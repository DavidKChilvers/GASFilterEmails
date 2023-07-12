/** Potentially useful code
  let sQuery = `after:${getEpoch(new Date("2022-01-01 00:02:03"))} before:${getEpoch(new Date("2023-02-01 00:02:03"))}`;

  //let query = Utilities.formatString("after:%d-%02d-%02d", dtQuery.getFullYear(), dtQuery.getMonth()+1, dtQuery.getDate() );              
  //let query = 'after:' + d.getDate()  + "-" + (d.getMonth()+1) + "-" + d.getFullYear() 
  // + " " + d.getHours() + ":" + d.getMinutes();

  //let d = new Date();
  //d.setDate(d.getDate() - 10)
*/

/** ============================================================================
 * getEpoch
 * 
 * Convert date object to a seconds since epoch date, i.e January 1, 1970
 * 
 * Numirc seconds Can be used as a date
 * 
 * epoch date is number of ms elapsed since i.e January 1, 1970
function getEpoch(dtDate)
{
  const secondsSinceEpoch = (date) => Math.floor(date.getTime() / 1000);
  const dt = new Date();
  
  //dt.setHours(12, 15, 0, 0);
  return(secondsSinceEpoch(dtDate));
};
 */




/**
function testruntime(times_run = 0)
{
  if (times_run >= 0 )
    console.log(`ran ${times_run} times`);
  Utilities.sleep(30000);
  testruntime(times_run + 1);
}
*/

/**
 * Archive emails by batches preventing controlling limiting the execution time and  
 * creating a trigger if there are still threads pending to be archived.
 
function batchArchiveEmail(){
  const start = Date.now();
  /** 
   * Own execution time limit for the search and archiving operations to prevent an 
   * uncatchable error. As the execution time check is done in do..while condition there  
   * should be enough time to one search and archive operation and to create a trigger 
   * to start a new execution. 
   * / 
  const maxTime = 25 * 60 * 1000; // Instead of 25 use 3 for Google free accounts
  const batchSize = 100;
  let threads, elapsedTime;
  /** Search and archive threads, then repeat until the search returns 0 threads or the 
   * maxTime is reached
   * / 
  do {
    threads = GmailApp.search('label:inbox before:2021/1/1');
    // console.log(`${threads[0] ? threads[0].getFirstMessageSubject() : 'No more threads'}`)
    for (let j = 0; j < threads.length; j += batchSize) {
      GmailApp.moveThreadsToArchive(threads.slice(j, j + batchSize));
    };
    /**
     * Used to prevent to have too many calls in a short time, might not be 
     * necessary with a large enough batchSize
     * /
    Utilities.sleep(`2000`); 
    elapsedTime = Date.now() - start;
  } while (threads.length > 0 &&  elapsedTime < maxTime);
  if(threads.length > 0){
    /** Delete the last trigger * /
    deleteTriggers();

    /** Create a one-time new trigger * /
    ScriptApp
    .newTrigger('batchArchiveEmail')
    .timeBased()
    .after(60 * 1000)
    .create();
    console.log(`trigger created`)
  } else {
    /** Delete the last trigger * /
    deleteTriggers();
    console.log(`No more threads to archive`);
  }
}
*/

/**
 * Creates the first trigger to call batchArchiveEmail. 
function init(){
   ScriptApp
    .newTrigger('batchArchiveEmail')
    .timeBased()
    .after(60 * 1000)
    .create();
    console.log(`trigger created`)
}
*/

/**
 * As there is a limit on the number of triggers that a user might have for each 
 * project, delete the triggers.
function deleteTriggers()
{
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
  ScriptApp.deleteTrigger(triggers[i]);
  }
}
 */

/**
function save_attachment_in_folder()
{

  let image_types_re = /jpg$/ig;
  var folder = DriveApp.getFolderById('1sLQS5mFV5gPUXC-VElA18cSwOlcCsqXB');
  var userId = "somptingvillagemorris@gmail.com";
  let query = 'has:attachment after:2021-10-20';
  query = "has:attachment";
  query = 'has:attachment after:2021-10-20';
  //var res = Gmail.Users.Messages.list(userId, {q: query});//I assumed that this works
  let email_threads = GmailApp.search(query, 0, 10);

  for (let ix in email_threads)
  {
    // Log all the subject lines in the first thread of your inbox
    //let thread = GmailApp.getInboxThreads(0, 1)[0];
    let thread = email_threads[ix]
//  console.log( Utilities.formatString("%d %s",
//        thread.getMessageCount(), thread.getLastMessageDate()));
    let messages = thread.getMessages();
    for (let i = 0 ; i < messages.length; i++)
    {
      let msg = messages[i];
      //console.log(Utilities.formatString("%s %s %s", msg.getSubject(), msg.getDate(), msg.getFrom()));
      let attA=messages[i].getAttachments();
      attA.forEach(function(a)
      {
        let file_name = a.getName(); 
        if (file_name.match(image_types_re) != null)
        {
        console.log( Utilities.formatString("%s %s Att Name %s",
          msg.getDate(),  msg.getFrom(), a.getName()));
        var ts=Utilities.formatDate(new Date(),Session.getScriptTimeZone(), "yyMMddHHmmss");
        folder.createFile(a.copyBlob()).setName(msg.getSubject() + ' ' + ts + a.getName());
        }
      });
    }    
  }
}
*/