/** Filter incoming emails to remove Spam
 * The script is set up to run automatically every hour.
 * Emails for the last 30 days that dont have a ContactUs Label are read. A maximum of 500 emails will be
 * processed in on read.
 * 
 * There are three sheets:
 *    +  'Regexes' - specifies actions and regular expressin matches as below.
 *    +  'Log'     - log of executions/ Note the execution logs associated with these scripts have more detail.
 *    +  'Config'  - not currently used
 * 
 * The emails are filtered according to the Regexes sheet.
 * Each row is processed:
 *   + If any subject starts with 'ContactUs:' This email has already been processed. Label with
 *     ContactUsRead and ignore.
 *   + If columns 1, ..., n match is found then the appropriate action in column 0 is applied
 *     no further processing takes place.
 * 
 * The column formats are 
 * Column 0  Action being one of
 *    + Read - Label as ContactUsRead and no futher action
 *    + Spam - Label as ContactUsSpam and no futher action
 *    + Fwd:<email address>
 *             Label with ContactUsFwd and forward to email address. No futher action
 * 
 * Colums 1, .., n are The match is run on each column and the column matches ANDed together.
 * If the result is True then the action in column 0 is taken as described above.
 * 
 * Each column has the format of [Optional not] <PartOfEmail>:<regular expression> where:
 *    + not implies negate match for that column
 *    + <PartOfEmail> can be email, subject or body or from
 *    + <regular expression> to which to match the part of the email. 
 *    + There is a special directive 'Default' where the action is applied if that row is reached.
 */

/**
 * Being Google workspace, script has 30 minutes. 25 minutes would be conservative
 * A free space is 6 minutes
 * But not expected to last more than 15 minutes.
 * 
 *  Will filter approximately 1.5 threads/sec
 *  Maximum number of threads per get is 500. So maximum time should be ~ 5.5m (333 secs)
 */
const maxTime = 15 * 60 * 1000;     // = 15 mins

/** This is a test entry funcion. Though at the moment it performs exactly the same as the 
 * entry function below that is used for regular triggers.
 */
function myFunction()
 {
  let config = ReadConfig();

  // If regexes then the result is an error string. Else everything is OK
  let regexes = getRegularExpressionFromSheet("RegExes");
  if (typeof regexes === 'string')
  {
    return;
  }

  let sQuery = "newer_than:30d  -label:ContactUsSpam -label:ContactUsFwd -label:ContactUsRead label:inbox"
  ApplyRegexesToEmails(config, regexes, sQuery);
}


/** TestRegexes
 * 
 * This entry is for the button on the regexes sheet.
 * 
 * Test the validity of the actions and regexes
 * 
 */
function TestRegexes()
 {
  let config = ReadConfig();

  let regexes = getRegularExpressionFromSheet("RegExes");

  // If regexes then the result is an error string. Else everything is OK
  if (typeof regexes === 'string')
  {
    Browser.msgBox(Utilities.formatString("Error in Regexes sheet.\n%s", regexes));
  }
  else
    Browser.msgBox("Everything is OK");
}


/**
 * Entry point for executio with exception handling
 * 
 */
function FilterEmailsWithExceptionHandling()
{
  let email_address = "ithelpdesk@worthingu3a.org.uk"
  let subject = "Error in U3A Contact email forwarding"
  //subject = "Test ignore"
  let _body = "The script in contact@worthingu3a.org.uk:FilterContactUsEmails has exceptioned. "

  try
  {
    FilterEmails()
  }
  catch (e)
  {
    let err_str = Utilities.formatString( "%s\n",e);
    console.log("Exception information:")
    console.log(err_str)
    console.log(e.stack)

    let body = Utilities.formatString( "%s:\n\n%s", _body, err_str)
    console.log("body=", body)
    let res = MailApp.sendEmail(email_address, subject, body, {});  
    
    throw "Filter emails has exceptioned"
  }
}


/** FilterEmails
 * 
 * Entry point for regular triggers.
 * 
 * Get Actions and expressions from regexes sheet
 * 
 * Apply actions/expressions to email that match specified query
 */
function FilterEmails()  
{
  // There is no config at the moment

  let config = ReadConfig();

  // Read regular expressions from spreadsheet
  // For Each row, AND together the columns B to E if True take action in column A
  // If regexes then the result is an error string. Else everything is OK
  let regexes = getRegularExpressionFromSheet("RegExes");
  if (typeof regexes === 'string')
  {
    return;
  }

  // Apply regular expressions
  let sQuery = "newer_than:30d -label:ContactUsSpam -label:ContactUsFwd -label:ContactUsRead "
  ApplyRegexesToEmails(config, regexes, sQuery);
}

/** ApplyRegexesToEmails
 * 
 * Apply regular expressions to emails
 * 
 * Outline description
 *  + Make sure labels are available.
 *  + Read email threads (only one message per thread is allowed)
 *  + Call ApplyRegexesToEmail to Apply Actions/expressions to emails
 *  + Add a row per execution to speadsheet log
 *  + Create a script execution log entry
 */
function ApplyRegexesToEmails(config, regexes, sQuery) 
{
  // Make sure labels exist
  let label_spam = GmailApp.getUserLabelByName("ContactUsSpam");
  if (label_spam == null)
  {
    GmailApp.createLabel("ContactUsSpam");
  }
  let label_fwd = GmailApp.getUserLabelByName("ContactUsFwd");
  if (label_fwd == null)
  {
    GmailApp.createLabel("ContactUsFwd");
  }
  let label_read = GmailApp.getUserLabelByName("ContactUsRead");
  if (label_read == null)
  {
    GmailApp.createLabel("ContactUsRead");
  }

  // Read threads that match specified query
  let threads = GmailApp.search(sQuery); //, 0, 500);//Get queries since given date

  // Apply Actions/expressions to emails
  let res = ApplyRegexesToEmail(config, regexes, threads, label_spam, label_fwd, label_read);

  SpreadsheetLog(res);

  console.log("SPAM\r\n"+ res.s_spam);
  console.log("FWD\r\n" + res.s_fwd);
  console.log("READ\r\n" + res.s_read);
  
  console.log(Utilities.formatString("%s%d Threads Processed of %d total",
                 res.timed_out ? "!!! Script Timed out - " : "",
                 res.threads_processed, res.threads_count));
}


/** ApplyRegexesToEmail
 * 
 * Apply Regexes to emails that have been read
 * 
 * Args:
 *    config - Config data. None at the moment.
 *    regexes - Array  1/row of arrays 
 *      [0]   Action
 *      [1]   Argument to action (i.e. text after ':')
 *      [2n]  Part of email to match and negate
 *      [2n+1] Regex
 *    label_... three labels to apply to emails when processed.
 *    
 * Outline view:
 *  For Each thread:
 *      For each message in a thread:
 *          Get individual parts of the message
 *          For each row in regexes array:
 *              Have we reached the default action. If so record action and argument
 *              For each regex in row of regexes:
 *                  Apply regex
 *                  If (match & negate) OR (not match and don't negate):
 *                      False so this row does not match. Flag no match and Exit For
 * 
 *              Test if all regexes in row returned true
 *                  If so this is the row that will define action. Record action and any argument.
 * 
 *      Apply action to email and label
 *      Test to make sure we are not out of time
 * 
 *  Return results
 */
function ApplyRegexesToEmail(config, regexes, threads, label_spam, label_fwd, label_read)
{
  let s_spam = '';
  let s_fwd = '';
  let s_read = '';
  let timed_out = false;
  const start = Date.now();
  let threads_processed = 0;

  for(let ix_thread=0; ix_thread < threads.length; ix_thread++)
  {
    let thread = threads[ix_thread];
    let labels = thread.getLabels();
    let doesThisThreadHaveTheLabel = false;
    
    // Actions are  0=> Read, 1=> Fwd, 2 = Spam
    // prefixes are 0 => default,  2=> body, 4=> subject 6=> from 8=> msg_from; odd => negate
    let messages = GmailApp.getMessagesForThread(thread);
    for (let j=0; j<messages.length;j++) 
    {
      let message_date_str = Utilities.formatDate(messages[j].getDate(), "GMT", "yyyy-MM-dd HH:mm", );
      let message=messages[j].getBody()
      let subject=messages[j].getSubject();
      let from=messages[j].getFrom();
      let to=messages[j].getTo();

      // Limit the length of the subject that is logged.
      let subject_log = (subject.length > 35) ? subject.substring(0, 32) + "..." : subject

      let row_result = 0;   // 0=> Read, 1=> Fwd, 2 = Spam
      let _match = message.match(/Email Address.......: (?<from>[a-zA-Z0-9_.@]*) /)
      let msg_from = _match == null ? null : _match[1];
      let row_result_arg1 = null;

      /**

      if (_match == null)
      {
        // This message was sent directly to contact@wor... not via ContactUs
        // assume it is just read
        // This may be fixed in the future to allow filtering on such messages
        s_read +=Utilities.formatString('From: %s Date: %s Subject %s\n', from, message_date_str, subject_log);

        label_read.addToThread(thread);
        continue;
      }
      */
      // For emails submitted via the form, the user entered 'from' is embedded into the body of the email

      // Rows form an OR of conditions to be Spam
      for (let rexes_row = 0; rexes_row < regexes.length; rexes_row++)
      {
        // let rexes = rex_strings[rexes_row];
        let rexes = regexes[rexes_row];

        // Have we found the default action? There will only be 2 columns 
        if (rexes[2] == 0)
        {
          //Default action
          row_result = rexes[0];
          row_result_arg1 = rexes[1];
          break;
        }
        // Columns form an AND of conditions to be Spam
        let row_col_result = true;
        for (let rexes_col = 2; rexes_col < rexes.length; rexes_col+=2)
        {
          // Split the prefix enumeration into part of email and negate result
          // prefixes are 2=> body, 4=> subject 6=> from; odd => negate
          let email_part = rexes[rexes_col] >> 1;
          let negate = (rexes[rexes_col] & 1) != 0;

          // Find part of email to apply regex to.
          let rex_string_to_match = (email_part == 1) ? message:
                                    (email_part == 2) ? subject :  
                                    (email_part == 3) ? from :  "";
                                    (email_part == 4) ? msg_from :  "";
          let _match = rex_string_to_match.match(rexes[rexes_col + 1])

          // If result is false row does not match. End this loop
          if (((_match == null) && !negate) || ((_match != null) && negate))
          {
            row_col_result = false;
            break;
          }
        }
        // Did this row match. If so record action and any argument
        if (row_col_result)
        {
          row_result = rexes[0];   // 0=> Read, 1=> Fwd, 2 = Spam
          row_result_arg1 = rexes[1];
          break;
        }
      }

      // Create log string. Label and apply any action
      let sEmail = Utilities.formatString('From: %s msg_from %s Date %s Subject: %s\n',
                                               from, msg_from, message_date_str, subject_log);
      // Label as Read no further action
      if (row_result == 0)
      {
        s_read += sEmail;
        label_read.addToThread(thread);
      }
      // Label as Fwd and forward with 'reply to' set to submitters address
      // make sure subject is not too long
      let trimmed_subject = subject.substring(0, 100) + "..."
      if (row_result == 1)
      {
        s_fwd += sEmail;
        messages[j].forward(row_result_arg1,
        {
          subject: "ContactUs: " + trimmed_subject,
          replyTo: msg_from
        });
        label_fwd.addToThread(thread);
      }
      // Label as Spam, no further action
      else if (row_result == 2)
      {
        s_spam += sEmail;
        label_spam.addToThread(thread);
      }
    } //for (let j=0; j<messages.length;j++)
    
    // Check we are not going to time out.
    elapsedTime = Date.now() - start;
    if (elapsedTime > maxTime)
    {
      timed_out = true;
      break;
    }

    threads_processed++;
  }    //for(let ix_thread=0; ix_thread < threads.length; ix_thread++)

  return {timed_out: timed_out,
          threads_count: threads.length, threads_processed: threads_processed,
          s_spam, s_fwd, s_read };
}

/** =========================================================================
 * 
 * getRegularExpressionFromSheet
 * 
 * read regular expressions from given sheet
 * 
 * returns array of arrays
 *    each
 * 
 */
function getRegularExpressionFromSheet(WorksheetName)
{
  let shRexEs = SpreadsheetApp.getActive().getSheetByName(WorksheetName);

  let permissible_actions = ["read", "fwd", "spam"];
  let permissible_prefixes = ["default", "body", "subject", "from", "msg_from"];

  let valid_actions_str = "Valid Actions are: 'Spam', 'Fwd:<EmailAddress>' or 'Read'"
  let valid_prefixes_str = "Valid prefixes are: [Optional not]  'body:', 'subject:', 'from:', 'msg_from:' or Default"

  //Check sheet exists
  if (shRexEs == null)
  {
    return null;
  }
  
  // Get data from sheet into an array
  var rng_rexes = shRexEs.getRange(2, 1, 100, 5);
  var rexes_data_rows = rng_rexes.getValues();

  let regexes = new Array(0);


  for (let row = 0; row < rexes_data_rows.length; row++)
  {
    let regex_row_data = rexes_data_rows[row];
    let regex = new Array(0); //regex_row_data.length * 2);

    // null first column in array terminates data
    if (regex_row_data[0].trim().length == 0)
    {
      break;
    }

    // Invalid regular expressions will cause an exception. Catch them so we can report the error
    // Also use exceptions for reporting invalid actions or parts of email
    try
    {  /** actions are spam, read (= processed) or Fwd:<email address>
        *
        * rex format is 
        * [0] action spam, read (= processed) or Fwd:<email address>
        * [1] text after : if any, null otherwise
        * [n] regex prefix, "read", "fwd", "spam"
        * [n + 1] regular expression
        */
      let action_arg = (regex_row_data[0].toLowerCase());

      let colon_ix = action_arg.indexOf(":");
      let action = (colon_ix >= 0) ? action_arg.substring(0, colon_ix) : action_arg;

      if (permissible_actions.indexOf(action) >= 0)
      {
        regex[0] = permissible_actions.indexOf(action);
        regex[1] = (colon_ix >= 0) ? action_arg.substring(colon_ix + 1): null;
      }
      else
      {
        throw(Utilities.formatString("Invalid Action, row %d: '%s'\\n\\n%s", row + 2, action, valid_actions_str));
      }

      // Now process [not] <partOfEmail>:<regulkar expression>
      for (let col = 1; col < regex_row_data.length; col++)
      {
        // Default is a catchall regular expression
        if (regex_row_data[col].toLowerCase() == "default")
        {
          //What to do if we fall off the end of the rows and no row_results = true
            regex[2 * col] = 0;               // 0 => Default
            break;
        }

        // Null column terminates columns
        if (regex_row_data[col].trim().length == 0)
        {
          break;
        }

        // Split [not] <PartOfEmail> and <regular expression>
        let colon_ix = regex_row_data[col].indexOf(":");
        if (colon_ix < 0)
        {
          throw(Utilities.formatString("No colon (prefix separator) row %d : '%s'",
            row + 2, regex_row_data[col] ));
        }
        let prefix_type = 0;
        let regex_prefix = regex_row_data[col].substring(0, colon_ix).toLowerCase();

        if (regex_prefix.substring(0, 4) == "not ")
        {
          prefix_type = 1;                          // negate is odd
          regex_prefix = regex_prefix.substring(4);
        }

        // Check part of emails is valid
        if (permissible_prefixes.includes(regex_prefix))
        {
          // Prefixes are default=0, body=2, subject = 4, from=6, msg_from=8
          prefix_type += permissible_prefixes.indexOf(regex_prefix) * 2;
        }
        else
        {
          throw(Utilities.formatString("Invalid Prefix row %d: '%s'\\n\\n%s",
           row + 2, regex_prefix, valid_prefixes_str ));
        }

        regex.length = 2 * col + 1;
        regex[2 * col] = prefix_type;

        // Creating a new regular expression object parses reg ex. Need to catch errors.
        try
        {
          regex[2 * col + 1] = new RegExp(regex_row_data[col].substring(colon_ix + 1), 'i');
        }
        catch(e)
        {
           throw(Utilities.formatString("Invalid regular expression row %d:\\n\\n '%s'\\n\\n %s",
            row + 2, regex_row_data[col].substring(colon_ix + 1), e));
        }
      }
      regexes.length = row + 1;
      regexes[row] = regex;
    }

    // Report any errors
    catch(e)
    {
      let err_str = Utilities.formatString( "%s\n",e);
      console.log(err_str)
      return err_str;      
    }
  }

  console.log(Utilities.formatString("%d Regular expressions read from spreadsheet", regexes.length));
  return regexes;
}

/** =============================================================================
 * ReadConfig
 * 
 * Read configuration data from Config sheet.
 * There is currently none.
 */
function ReadConfig()
{
  let shConfig = SpreadsheetApp.getActive().getSheetByName("Config");
  let rng_config = shConfig.getRange(2, 1, 20, 2);
  let config_data_rows = rng_config.getValues();
  let config_dict = {};
  for (let rows_config = 0; rows_config < config_data_rows.length; rows_config++)
  {
    let row_data = config_data_rows[rows_config]
    if (row_data[0].trim().length == 0)
      break;
    config_dict[row_data[0].trim()] = row_data[1].trim();
  }

  return config_dict;
}

/**  SpreadsheetLog
 * 
 * Add log row to Log Sheet in spreadsheet
*/
function SpreadsheetLog(res)
{
  let shLog = SpreadsheetApp.getActive().getSheetByName("Log");
  let log_str =   Utilities.formatString("%s%d Threads Processed of %d total, emails: %d fwd, %d spam, %d read",
                 res.timed_out ? "!!! Script Timed out - " : "",
                 res.threads_processed, res.threads_count,
                 res.s_fwd.split('\n').length - 1, res.s_spam.split('\n').length - 1,
                 res.s_read.split('\n').length - 1);

  //shLog.appendRow([Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm"), log_str]);

  shLog.insertRows(1);

  let rng = shLog.getRange(1,1,1, 2) ;
  let data = rng.getValues();
  data[0][0] = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm");
  data[0][1] = log_str;
  rng.setValues(data);

  //Limit number of rows to 1000 - 1023
  if ( shLog.getMaxRows() >= 1024)
  {
    shLog.deleteRows(1001, shLog.getMaxRows() - 1000);
  }

}


