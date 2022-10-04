
function saveEmails() {
  var COMPLETE = 'label: Status/Complete';

  var start = 0;
  var max = 500;
  
  var threads = GmailApp.search(COMPLETE, start, max);
  if (threads!=null){
    console.log(threads.length + " threads found 🎉");
  } else {
    console.warn("No emails found within search criteria 😢");
    return;
  }

  // TODO: Check to see if the headings are already there before appending.
  appendData(1, [["threadID", "MessageIDs", "Subject", "Content", "First message date", "Last message date", "Email Address", "Name",
    "Department/Major", "University Status", "Question/Request", "Message Status", "Topic", "Software", "Staff Responsible", "DSC Responsible", "DSA Request", "Instruction request",
    "Location", "Misc", "NC State College", "NC State Status", "Referral", "Consultation?", "Medium", "Reference", "transaction?",
    "Duration (minutes)", "Public science/scholarship?", "Notes (optional)", "Quote"]]);
  
  var totalEmails = 0;
  var rows = [];

  // existing_thread_ids = SpreadsheetApp.getActiveSpreadsheet().getRange("A:A").getValues();
  // existing_message_ids = SpreadsheetApp.getActiveSpreadsheet().getRange("B:B").getValues();

  var ids_range = cleanRange(SpreadsheetApp.getActiveSpreadsheet().getRange("A:B").getValues());

  existing_ids = new Map(ids_range.map(
    row => {
      return [row[0], row[1]];
    }
  ));

  /* existing_ids.forEach((value,key,map) => {
    console.log(`${key} = ${value}`);
  }); */

  // return;

  for (var i in threads) {
    var thread=threads[i];

    // Ignore empty rows
    if(thread == null || thread == undefined) {
      continue;
    } else {
      // This is a real thread, log it
      // console.log("thread: " + thread.getId());
    }


    var threadID = thread.getId();

    


    var last_date = thread.getLastMessageDate();
      
    var labels_raw = thread.getLabels();
    var labelnames = [];
    var labels = [];

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

    for (var i in labels_raw) {
      labelname = labels_raw[i].getName();

      labelnames[i] = labelname;

      labels.push(splitLabels(labelnames[i]));
    }
    
    for(var i in labels) {
     labels_final.get(labels[i][0]).push(labels[i][1]);
    }

      var msgs = thread.getMessages();
      var messageID = [];

      var subject;
      var content = [];
      
      var patron_details = new Map();

      for (var j in msgs) {
        var msg = msgs[j];
        var msg_id = msg.getId();
        messageID.push(msg_id);    
        
        if(j == 0) {
          // Use the subject from the first message
          subject = msg.getSubject();
          // Process values in form email to get the columns
          body = msg.getBody();
          lines = body.split(/\r?\n/);
          // console.log("lines: " + lines);

          // console.log("first line: " + lines[0]);
          for(var k in lines) {
            line = lines[k];

            if(!lines[0].includes("Name: ")) {
              patron_details.set("Name", "");
              patron_details.set("Email", "");
              patron_details.set("University Status", "");
              patron_details.set("Department/Major", "");
              patron_details.set("Request", "");

              break;
            }

            if(k <= 4) {
              detail = line.split(": ");
              // console.log("detail: " + detail);
              patron_details.set(detail[0], detail[1]);
            } else {
              patron_details.set("Request", patron_details.get("Request") + line);
            }
        }

          /* Format:
          "Name: Jennifer Swartz
          Email: jwswartz@ncsu.edu
          University Status: Graduate Student
          Department/Major: Educational Leadership
          Request: full version of ATLAS.ti"
          */
          var first_date = msg.getDate();  
        }
          content.push(msg.getBody());
      }

      messageIDs = messageID.join(",");

      // Columns:
      /*   appendData(1, [["threadID", "MessageIDs", "Subject", "Content", "First message date", "Last message date", "Email Address", "Name",
    "Department/Major", "University Status", "Question/Request", "Message Status", "Topic", "Software", "Staff Responsible", "DSC Responsible", "DSA Request", "Instruction request",
    "Location", "Misc", "NC State College", "NC State Status", "Referral", "Consultation?", "Medium", "Reference", "transaction?",
    "Duration (minutes)", "Public science/scholarship?", "Notes (optional)", "Quote"]]);
      */
      // TODO: Add in the new columns
    var dataLine = [
      threadID, messageIDs, subject, content.join(",").slice(0, 50000), first_date, last_date, // There is a 50000 character limit to cells, so the content is truncated.
      patron_details.get("Email"), patron_details.get("Name"), patron_details.get("Department/Major"),
      patron_details.get("University Status"), patron_details.get("Request"),
      labels_final.get("Status").join(","), labels_final.get("Concepts").join(","), labels_final.get("Tools").join(","),
      labels_final.get("Staff").join(","), labels_final.get("DSC").join(","), labels_final.get("Data Science Academy").join(","),
      labels_final.get("Instruction request").join(","), labels_final.get("Location").join(","), labels_final.get("Misc").join(",") ];



    rows.push(dataLine);
  }

  totalEmails = totalEmails + rows.length;

  if (threads.length == max){
      console.log("Reading next page...");
  } else {
      console.log("Last page read 🏁");
  }
  start = start + max; 
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

console.info(totalEmails+" emails added to sheet 🎉");
}

function appendData(line, array2d) {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(line, 1, array2d.length, array2d[0].length).setValues(array2d);
}

function splitLabels(labelname) {
  var labelpair = [];
  if(!labelname.includes('/')) {
    labelpair = [labelname, "yes"];
  } else {
    labelpair = labelname.split('/');
  }
  return labelpair;
}

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

    return range;
}