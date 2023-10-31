// Configuration Variables

var sheetID = "1vAX2P-kQ3P1YI3z4aENxSdvaEqtRsr57BN_LSuu_4Ws";
var timeZone = "America/New_York";

//Sheet Configs
var input = SpreadsheetApp.openById(sheetID).getSheets()[0].getRange("A2:B").getValues().filter(String);

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

//Setting Time and formatting
var today = new Date();
console.log(today);
formattedTime = Utilities.formatDate(today, timeZone, "yyyy-MM-dd");
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
	getEmailsAndStoreInSheet();
	sendEmailWithTable();
}

function getEmailsAndStoreInSheet() {

	var threads = GmailApp.search("after:" + formattedTime + " from:" + emailToFilter);
	//var threads = GmailApp.search("from:" + emailToFilter);
	sheetToUse.appendRow(["Formatted Date", "Timestamp", "Task Name", "Owner", "Notes", "Link"]);
	// Iterate through threads and messages
	for (var i = 0; i < threads.length; i++) {
		var messages = threads[i].getMessages();

		for (var j = 0; j < messages.length; j++) {
			var message = messages[j];
			var date = message.getDate();
			var subject = message.getSubject();
			var plainbody = message.getPlainBody();
			var body = message.getBody();
			if (body.indexOf("https://clickup.com/") !== -1) {

				var link = body.match(/https:\/\/app\.clickup\.com\/\S+/)[0].replace('"', "").replace("&amp;", "&");
				var whoDid = body.match(/">by (.*?)<\/p>/)[1];
				var content = body.match(/">@(.*?)<\/tbody>/)[1];

				sheetToUse.appendRow([formattedTime, date, subject, whoDid, htmlToPlainText(content), link]);
				sheetToUse.getRange("A2:A").setNumberFormat('@');

				threads[i].addLabel(archiveLabel);
				threads[i].moveToArchive();
			}
		}

	}
	customSheetFormat(sheetToUse);
}

function htmlToPlainText(html) {
	// Remove HTML tags using a regular expression
	var plainText = html.replace(/<[^>]+>/g, '');

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
	// Find the target folder by name
	var folder = DriveApp.getFolderById(folderId);

	// Create a new Google Sheet
	var sheet = SpreadsheetApp.create(sheetname);

	// Move the newly created Google Sheet to the target folder
	DriveApp.getFileById(sheet.getId()).moveTo(folder);

	output.push(sheet.getId());
	output.push(sheet.getUrl());

	// Return the Google Sheet ID
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

	// Apply multiple formatting in a single batch update
	range.setWrap(true);
	range.setBorder(true, true, true, true, true, true);

	var firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn());
	firstRow.setFontWeight("bold");
	firstRow.setHorizontalAlignment("center");

	Logger.log("Sheet formatting complete.");
}
