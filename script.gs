// Configuration Variables

var sheetID = "dummyID";
var timeZone = "America/New_York";

// Debugging Purposes
var sheetIDDebugLogs = "dummyID";
var developerEmail = "xxxx@gmail.com";

//Sheet Configs
var input = SpreadsheetApp.openById(sheetID).getSheets()[0].getRange("A2:B").getValues().filter(String);
var debugSheet = SpreadsheetApp.openById(sheetIDDebugLogs).getSheets()[0]

//input variables
var input = input.filter(function(row) {
	return !row.every(function(cell) {
		return cell === '';
	});
});
var inputJSON = arrayToJsonObject(input);

const emailToFilter = inputJSON["Email To Filter"];
const emailToNotify = inputJSON["Email To Notify"];
const labelName = inputJSON["Label Name"];
const folderId = inputJSON["GDrive FolderID"];
const runByDate = inputJSON["Run By Date"];

//Setting Time and formatting
if (runByDate) {
	console.log("run by date is :- " + runByDate)

	var today = new Date(runByDate);
  var tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  if (tomorrow.getMonth() !== today.getMonth()) {
    tomorrow.setDate(1); // Set the day to 1 to handle month rollover
    tomorrow.setMonth(today.getMonth() + 1); // Move to the next month

    // Check if the next month is in the next year
    if (tomorrow.getMonth() === 0) {
      tomorrow.setYear(today.getFullYear() + 1); // Move to the next year
    }
  }

  console.log("tommorow is"+tomorrow);

	SpreadsheetApp.openById(sheetID).getSheets()[0].getRange("B6").setValue("");
} else {
	var today = new Date();

  var tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  if (tomorrow.getMonth() !== today.getMonth()) {
    tomorrow.setDate(1); // Set the day to 1 to handle month rollover
    tomorrow.setMonth(today.getMonth() + 1); // Move to the next month

    // Check if the next month is in the next year
    if (tomorrow.getMonth() === 0) {
      tomorrow.setYear(today.getFullYear() + 1); // Move to the next year
    }
  }

  console.log("tommorow is"+tomorrow);
}

console.log(today);

formattedTime = Utilities.formatDate(today, timeZone, "yyyy-MM-dd");
formattedTommorrow = Utilities.formatDate(tomorrow, timeZone, "yyyy-MM-dd");

console.log(formattedTime);

outputCreateSheet = createGoogleSheetInFolder(folderId, formattedTime);

var sheetToUse = SpreadsheetApp.openById(outputCreateSheet[0]).getSheets()[0];
var reportSheetUrl = outputCreateSheet[1];

//Archive Manage Code

var archiveLabel = GmailApp.getUserLabelByName(labelName);

if (!archiveLabel) {
	archiveLabel = GmailApp.createLabel(labelName);
}

// Run this function to retrieve and store the emails
function rUNTHIS() {
	try {
		getEmailsAndStoreInSheet();
		sendEmailWithTable();
	} catch (error) {

		debugSheet.appendRow([error.message,formattedTime]);

		MailApp.sendEmail({
			to: developerEmail,
			subject: "Critical : Met Error Automation Script (General Run this error) : Clickup Sheet Summary",
			htmlBody: error.message,
		});


	}
}

function getEmailsAndStoreInSheet() {

	var errorFlag = 0;

	var threads = GmailApp.search("after:" + formattedTime +" before:"+ formattedTommorrow +" from:" + emailToFilter);

  console.log("Searching Emails for ...");
  console.log("after:" + formattedTime +" before:"+ formattedTommorrow +" from:" + emailToFilter);

	sheetToUse.appendRow(["Formatted Date", "Timestamp", "Task Name", "Owner", "Notes", "Link"]);
	// Iterate through threads and messages


	for (var i = 0; i < threads.length; i++) {
		var messages = threads[i].getMessages();

		for (var j = 0; j < messages.length; j++) {

			try {

				var message = messages[j];
				var date = message.getDate();
        var formattedEmailDate = Utilities.formatDate(date, timeZone, "yyyy-MM-dd");
				var subject = message.getSubject();
				var plainbody = message.getPlainBody();
				var body = message.getBody();
				if (body.indexOf("https://clickup.com/") !== -1) {

					var link = body.match(/https:\/\/app\.clickup\.com\/\S+/)[0].replace('"', "").replace("&amp;", "&");

          if(body.match(/">by (.*?)<\/p>/))
          {
            var whoDid = body.match(/">by (.*?)<\/p>/)[1];
            var content = body.match(/">by(.*?)<\/p>[\s\S]*?">Replies to this email will be added as comments<\/p>/)[0].replace("Replies to this email will be added as comments", "").replace(/">by(.*?)<\/p>/,'');
          }
          else
          {
            var whoDid="NA"
            var content = body.match(/">(.*?)<\/p>[\s\S]*?">Replies to this email will be added as comments<\/p>/)[0].replace("Replies to this email will be added as comments", "").replace('">','');
          }
					
					//sheetToUse.appendRow([body]);
					content = removeUnwantedText(htmlToPlainText(content));

          if(formattedTime==formattedEmailDate && returnFilteredEmail(body,link))
          {
            sheetToUse.appendRow([formattedTime, date, subject, whoDid, htmlToPlainText(content), link]);
          }

					

					threads[i].addLabel(archiveLabel);
					threads[i].moveToArchive();
				}

			} catch (error) {
				console.log("Met With an Error - Reported to developer with logs")
				debugSheet.appendRow([body,date, error.message]);
				Logger.log("An error occurred: " + error.message);

				errorFlag = 1;

			}
		}

	}

	if (errorFlag == 1) {
		MailApp.sendEmail({
			to: developerEmail,
			subject: "Warning : Single/Multiple Errors in getEmailsAndStoreInSheet : Clickup Sheet Summary",
			htmlBody: "Check the debug sheet for details",
		});
	}

	customSheetFormat(sheetToUse);
  sheetToUse.getRange("A2:A").setNumberFormat('@');
}

function returnFilteredEmail(emailBody,link)
{
  if(link.includes("=assignee_add") && emailBody.includes("Assigned to You"))
  {
    return false
  }
  else if(link.includes("=comment") && emailBody.includes("New comment"))
  {
    return false
  }
  else if(link.includes("=reaction") && emailBody.includes("liked your comment"))
  {
    return false
  }
  else
  {
    return true
  }
  
}

function htmlToPlainText(html) {
	// Remove HTML tags using a regular expression
	var plainText = html.replace(/<[^>]+>/g, ' ');

	// Replace common HTML entities with their plain text equivalents
	plainText = plainText.replace(/</g, '<');
	plainText = plainText.replace(/>/g, '>');
	plainText = plainText.replace(/&/g, '&');
	plainText = plainText.replace(/ /g, ' ');

	return plainText;
}

function sendEmailWithTable() {

	rowsmatched = MatchRows(sheetToUse, formattedTime);

	var datesRow = sheetToUse.getRange('B2:F').getValues();


	var table = '<table border="1" style="border-collapse: collapse;">'; // Add border attributes and collapse styling

	if (rowsmatched.length > 0) {
		table += '<tr style="background-color: #f0f8ff;"><td style="border: 1px solid black; padding: 5px;"><center><b>SR:</b></center></td><td style="border: 1px solid black; padding: 5px;"><center><b>Time Stamp</b></center></td><td style="border: 1px solid black; padding: 5px;"><center><b>Task Name</b></center></td><td style="border: 1px solid black; padding: 5px;"><center><b>Item Owner</b></center></td><td style="border: 1px solid black; padding: 5px;"><center><b>Note</b></center></td><td style="border: 1px solid black; padding: 5px;"><center><b>Link</b></center></td></tr>';
	}

	for (var i = 0; i < rowsmatched.length; i++) {
		table_row = i + 1
		table += '<tr><td style="border: 1px solid black; padding: 5px;"><center>' + table_row + '</center></td>';
		for (var j = 0; j < datesRow[i].length; j++) {
			table += '<td style="border: 1px solid black; padding: 5px;">' + datesRow[i][j] + '</td>'; // Add border and padding styling
		}
		table += '</tr>';
	}
	table += '</table>';

	var subjectToSend = 'Daily Summary Report :- Click Up Notifications';
	var bodyToSend = "Dear User, <br> <br> Here's a summary of the ClickUp notifications for you on <b>" + formattedTime + '</b>:<br><br>There were ' + rowsmatched.length + ' Email Notifications <br><br>' + table + '\n\n\n <p><a href="' + reportSheetUrl + '"><b>Click Here</b></a> for Google Sheet Reports : ' + "</p>\n\n<p style='color:#aaaaaa; font-size:10px;'>Note: These emails are labeled and moved to the archive. You can find the emails under the label :- " + labelName + "</p>";

	//console.log(bodyToSend);

	MailApp.sendEmail({
		to: emailToNotify,
		subject: subjectToSend,
		htmlBody: bodyToSend,
	});
}

function createGoogleSheetInFolder(folderId, sheetname) {

	output = []
	var folder = DriveApp.getFolderById(folderId);

	var sheet = SpreadsheetApp.create(sheetname);

	DriveApp.getFileById(sheet.getId()).moveTo(folder);

	output.push(sheet.getId());
	output.push(sheet.getUrl());

	return output;
}

function MatchRows(sheetToUse, tomatch) {
	var rowmatched = [];
	var datesRow = sheetToUse.getRange('A2:A').getValues();
	//console.log(datesRow);
	for (var j = 0; j <= datesRow.filter(String).length; j++) {
		if (datesRow[j] == tomatch) {
			rowmatched.push(j);
		}
	}
	console.log(rowmatched);
	return (rowmatched);

}

function arrayToJsonObject(data) {
	var jsonObject = {};
	for (var i = 0; i < data.length; i++) {
		var key = data[i][0];
		var value = data[i][1];
		jsonObject[key] = value;
	}
	return jsonObject
}

function customSheetFormat(sheet) {
	var range = sheet.getDataRange();

	range.setWrap(true);
	range.setBorder(true, true, true, true, true, true);

	var firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn());
	firstRow.setFontWeight("bold");
	firstRow.setHorizontalAlignment("center");

	Logger.log("Sheet formatting complete.");
}

function removeUnwantedText(text) {
  //console.log(text);
  var cleanedText = text.replace(/\n{2,}/g, '\n');
  var cleanedText = cleanedText.replace(/\s{5,}/g, ' ');
  var cleanedText = cleanedText.replace(/[\r\s]+/g, ' ');
  
	cleanedText = cleanedText.replace(/&nbsp;/g, ' ')
                           .replace(/&gt;/g, '>')
                           .replace(/&#x2F;/g, '/')
                           .replace(/&#x27;/g, "'")
                           .replace(/&amp;/g, '&')
                           .replace(/&#x3D;/g, '=')  // Replace equals (=)
                           .replace(/&#x3A;/g, ':')  // Replace colons (:)
                           .replace(/&#x2E;/g, '.')  // Replace periods (.)
                           .replace(/&#x2C;/g, ',')  // Replace commas (,)
                           .replace(/&#x20;/g, ' ');  // Replace spaces ( );

  console.log(cleanedText);
	return cleanedText;
}

function getLineNumber() {
	try {
		throw new Error();
	} catch (e) {
		return e.lineNumber;
	}
}
