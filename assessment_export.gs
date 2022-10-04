/*
  This script extracts data from emails for the getdathelp account. It pulls information from the thread labels, getdatahelp form results (in the first email message), and other data from the messages.

  The ID fields for each row are the thread ID (will not change if messages are added), and the messages IDs (a comma-separated list of all the message IDs). The messages ID will change as messages are added to the threads.

  For labels, if there are sublabels, such as for Staff, DSC, Concepts, and tools, the sublabels are the values under the label column. So for Staff, "Staff/Shannon" and "Staff/Joddy" would get saved under the Staff column with multiple staff labels being combined. If there are not sublabels, such as "Data Science Academy" and "Instruction requests", the column is set to "yes" to indicate that the label is present for the thread.
  
  It will check for thread rows that have already been added, and update them without overwriting any manually entered data. The email body field has a limit of 50000 characters, so the messages get truncated after that.

  To run the export you will need to be logged into the full account for getdatahelp, and open apps script within the assessment export sheet.

  TODO:
    1. Check for whether a thread has already been added, and update it with the new messages.
    2. Decide which columns to remove/hide
    3. Decide what to do for body cells > 50000 characters. Could either create a column for the first N messages in a thread, or leave as is and truncate.
    4. Once manual entry begins, need to verify that the manually entered data is not overwritten. The code only operates on the columns that are automatically pulled, but this still needs to be tested to be sure.
    5. Potentially add in parameters to the search to ignore non-consultation emails (IT tickets, security messages, etc.).
*/

function saveEmails() {
  // The string for the search. We may want to exclude messages that are not consultations (IT tickets, security messages, etc.).
  var COMPLETE = 'label: Status/Complete';

  // Pager variables to look at 500 messages at a time.
  var start = 0;
  var max = 500;
  
  // Searching the getdatahelp Gmail account for threads marked as complete
  var threads = GmailApp.search(COMPLETE, start, max);
  if (threads!=null){
    console.log(threads.length + " threads found 🎉");
  } else {
    console.warn("No emails found within search criteria 😢");
    return;
  }

  // Add the column headings to the sheet
  // TODO: Check to see if the headings are already there before appending.
  appendData(1, [["threadID", "MessageIDs", "Subject", "Content", "First message date", "Last message date", "Email Address", "Name",
    "Department/Major", "University Status", "Question/Request", "Message Status", "Topic", "Software", "Staff Responsible", "DSC Responsible", "DSA Request", "Instruction request",
    "Location", "Misc", "NC State College", "NC State Status", "Referral", "Consultation?", "Medium", "Reference", "transaction?",
    "Duration (minutes)", "Public science/scholarship?", "Notes (optional)", "Quote"]]);
  
  var totalThreads = 0;

  // The rows of new data are saved here
  var rows = [];

  // Get the IDs that are saved into the sheet
  var ids_range = cleanRange(SpreadsheetApp.getActiveSpreadsheet().getRange("A:B").getValues());
  existing_ids = new Map(ids_range.map(
    row => {
      return [row[0], row[1]];
    }
  ));

  // Print out the threadIDs and messageIDs that are already in the sheet
  /* existing_ids.forEach((value,key,map) => {
    console.log(`${key} = ${value}`);
  }); */

  // Outer loop to iterate over all threads
  for (var i in threads) {
    var thread=threads[i];

    // Ignore empty rows
    if(thread == null || thread == undefined) {
      continue;
    } else {
      // This is a real thread, log it
      // console.log("thread: " + thread.getId());
    }

    // Thread ID, saved as the row key
    var threadID = thread.getId();

    // Get the date of last message
    var last_date = thread.getLastMessageDate();
    
    // Get the labels from the thread
    var labels_raw = thread.getLabels();
    var labelnames = [];
    var labels = [];

    // These are the labels we are extracting, saved into a Map associative array, keyed by the top-level label
    var labels_final = new Map([
      ["Concepts", []],
      ["Data Science Academy", []],
      ["DSC", []],
      ["Instruction request", []],
      ["Location", []],
      ["Misc", []],
      ["Staff", []],
      ["Status", []],
      ["Tools", []]
    ]);

    // Extract the labels we are interested in
    for (var i in labels_raw) {
      labelname = labels_raw[i].getName();

      labelnames[i] = labelname;

      labels.push(splitLabels(labelnames[i]));
    }
    
    // Add the final label values into the Map
    for(var i in labels) {
     labels_final.get(labels[i][0]).push(labels[i][1]);
    }

      // All messages from the current thread
      var msgs = thread.getMessages();
      var messageID = [];

      var subject;
      var content = [];
      
      // Map holding the form values entered into the getdatahelp form on the website. They get saved into the first message of the thread.
      var patron_details = new Map();

      // Loop over each message to extract data
      for (var j in msgs) {
        var msg = msgs[j];
        var msg_id = msg.getId();
        // Add this message to the messageIDs list in the threadID row
        messageID.push(msg_id);    
        
        // Process the patron-entered details that are in the first message
        if(j == 0) {
          // Use the subject from the first message
          subject = msg.getSubject();
          // Process values in form email to get the columns
          body = msg.getBody();
          lines = body.split(/\r?\n/);

          // Loop through each line to get the field data
          for(var k in lines) {
            line = lines[k];

            // Fill in the details with empty strings if the fields do not exist in the first message (i.e. it did not come from the website form).
            if(!lines[0].includes("Name: ")) {
              patron_details.set("Name", "");
              patron_details.set("Email", "");
              patron_details.set("University Status", "");
              patron_details.set("Department/Major", "");
              patron_details.set("Request", "");

              // Leave the loop since this is not a form message
              break;
            }

            /* Format:
            "Name: J******* S*****
            Email: j*******@ncsu.edu
            University Status: Graduate Student
            Department/Major: Educational Leadership
            Request: full version of ATLAS.ti"
            */

            if(k <= 4) {
              // Get Name, Email, University Status, Department/Major, and the first line of Request
              detail = line.split(": ");
              // console.log("detail: " + detail);
              patron_details.set(detail[0], detail[1]);
            } else {
              // Get the rest of the Request lines (since it can be multi-line)
              patron_details.set("Request", patron_details.get("Request") + line);
            }
        }

          // Get the date from the first message
          var first_date = msg.getDate();  
        }
          // Add the message body content into the thread content column
          content.push(msg.getBody());
      }

      // Combine the message IDs into a comma-separated list
      messageIDs = messageID.join(",");

      // Columns:
      /*   appendData(1, [["threadID", "MessageIDs", "Subject", "Content", "First message date", "Last message date", "Email Address", "Name",
    "Department/Major", "University Status", "Question/Request", "Message Status", "Topic", "Software", "Staff Responsible", "DSC Responsible", "DSA Request", "Instruction request",
    "Location", "Misc", "NC State College", "NC State Status", "Referral", "Consultation?", "Medium", "Reference", "transaction?",
    "Duration (minutes)", "Public science/scholarship?", "Notes (optional)", "Quote"]]);
      */
      // TODO: Decide which columns to remove
    var dataLine = [
      threadID, messageIDs, subject, content.join(",").slice(0, 50000), first_date, last_date, // There is a 50000 character limit to cells, so the content is truncated.
      patron_details.get("Email"), patron_details.get("Name"), patron_details.get("Department/Major"),
      patron_details.get("University Status"), patron_details.get("Request"),
      labels_final.get("Status").join(","), labels_final.get("Concepts").join(","), labels_final.get("Tools").join(","),
      labels_final.get("Staff").join(","), labels_final.get("DSC").join(","), labels_final.get("Data Science Academy").join(","),
      labels_final.get("Instruction request").join(","), labels_final.get("Location").join(","), labels_final.get("Misc").join(",") ];

    // The data for the thread row is added into the threads data structure.
    rows.push(dataLine);
  }

  // The number of threads added
  totalThreads = totalThreads + rows.length;

  // Go to the next page of emails, 500 at a time
  if (threads.length == max){
      console.log("Reading next page...");
  } else {
      console.log("Last page read 🏁");
  }

  // Update which set of 500 emails we are on
  start = start + max;

  // Perform another search for the next 500 emails
  threads = GmailApp.search(COMPLETE, start, max);

  /* TODO Add in test to see if:
    1. The thread is already in the spreadsheet
      - if 1, check that if there are new messages
    2. There are any new messages in the thread
      - if not 2, skip that threadID row
      - if 2, update that threadID row

    Care must be taken not to mess up the manual entry rows
  */

  appendData(2, rows);

  console.info(totalThreads + " threads added to sheet 🎉");
}

/* Add data into the spreadsheet. Will need to be rewritten to add lines one at a time, depending on whether the thread exists/needs to be updated.
 Vars:
    line - The row in the spreadsheet to start on
    array2d - the spreadsheet data to add
  Return:
    none
*/
function appendData(line, array2d) {
  // Load the assessment export spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  // Add the data at the first line (includes the headings). This will need to be rewritten to allow for updating one thread at a time.
  sheet.getRange(line, 1, array2d.length, array2d[0].length).setValues(array2d);
}

/* Process a label, returning an array with the label name/value pair. Checks whether a label has sublabels (there is a '/'), and if not, "yes" is entered as the value to show that the label is present.
  Vars:
    labelname - the label string, in the format "Label/Sublabel" or "Label"
  Returns:
    labelpair - a 2 element array with element 0) the label name and 1) the label value
*/
function splitLabels(labelname) {
  var labelpair = [];
  // Check for sublabels. If there is not a sublabel, assign "yes" as the value to indicate presence of the label.
  if(!labelname.includes('/')) {
    labelpair = [labelname, "yes"];
  } else {
    // This label has a sublabel. Save the sublabel as the value.
    labelpair = labelname.split('/');
  }
  return labelpair;
}

/* Get rid of the empty space and headings on a range from the sheet.
  Vars:
    range - The Google Sheet Range from the api call
  Returns:
    range - The Google Sheet Range without headings and empty rows at the end
*/
function cleanRange(range) {
  // Clear the headings
  range.shift();

    // Clear the empty values from the end
    for(var i = range.length; i > 0; i--) {
      popped = range.pop();
      if(popped[0] !== "") {
        break;
      }
    }

    // Return a Range without heading or empty rows at the end
    return range;
}