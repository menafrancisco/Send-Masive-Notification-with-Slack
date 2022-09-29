//https://github.com/menafrancisco/Send-Masive-Notification-with-Slack.git

var POST_MESSAGE_ENDPOINT = "https://slack.com/api/chat.postMessage";

// This code adds a menu item to the Google Sheet that you can use to send your message

function onOpen(e) {
   SpreadsheetApp.getUi()
       .createMenu('Messager Menu')
       .addItem('Send to Recipients', 'postLoop')
       .addToUi();
 }

// This code gets the list of user ID's from your Google Sheet and sends your Slack message to them one-by-one.

function postLoop () {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA");
  var rangeValues = sheet.getDataRange().getValues();
  var initcell = rangeValues[1][1]-1;
  var token = rangeValues[2][1];
  let body;

  const lastColumn = sheet.getLastColumn();
  const [header, ...data] = sheet.getRange(initcell, 1, sheet.getLastRow() - 2, lastColumn)
    .getValues();
  const template= sheet.getRange(1, 2).getValue();
  var channel=""; 
  var country = "";

  sheet.getRange(1, sheet.getLastColumn())
        .setValue("PROCESSING...").setFontColor('#e69138');
      SpreadsheetApp.flush();

  cleanstatus(sheet, data, initcell);

  //Loop thru data, replace body placeholders
  data.forEach((row, r) => {
    
    body = template;
    channel=""; 
    country = "";
   
    //Replace body placeholders
    header.forEach((hCol, h) => {

      if (hCol.includes("{{") && hCol.includes("}}") ) {

        let rg = new RegExp(hCol, "g");
        //console.log(row[h]);
        body = body.replace(rg, row[h]);
      }

      if (hCol === "{{Slack ID}}" ) {
        channel = row[h];
      };

      if (hCol === "{{Holiday Country}}") {
        country = row[h];
      };      
    });
    console.log(row[0])
    var message = body;

    if(row[0] =='yes'){
      var answer = postToSlack(channel, message, token);
      console.log(answer.getResponseCode());
      console.log(answer.getContentText());

      //Update status on EMAILS sheet
      let rw = r + initcell+1;
      sheet.getRange(rw, sheet.getLastColumn())
        .setValue("MSG SENT").setFontColor('#6aa84f');
      SpreadsheetApp.flush();
      if (r % 5) {
        Utilities.sleep(50);
      }
    }else{
      //Update status on EMAILS sheet
      let rw = r + initcell+1;
      sheet.getRange(rw, sheet.getLastColumn())
        .setValue("NO ACTION").setFontColor('#6aa84f');
      SpreadsheetApp.flush();
      if (r % 5) {
        Utilities.sleep(50);
      }
    }   
  });

  sheet.getRange(1, sheet.getLastColumn())
        .setValue("COMPLETED").setFontColor('#6aa84f');
      SpreadsheetApp.flush();

}

function cleanstatus(sh, data, initcell) {
  data.forEach((row, r) => {
    //Update status 
    let rw = r + initcell+1;    
    sh.getRange(rw, sh.getLastColumn())
        .setValue("");
      SpreadsheetApp.flush();
      if (r % 5) {
        Utilities.sleep(50);
      }
  });
}

// This is the code that sends your message to Slack, it is called by the above function postLoop()

function postToSlack(channel, message, token) {
  var payload = {
    'channel' : channel,
    'as_user' : true,
	"blocks": [
		{
			"type": "section",
			"text": {
				"type": "mrkdwn",
				"text": message
			},
		}
	]
};
return UrlFetchApp.fetch(
  POST_MESSAGE_ENDPOINT,
  {
    method             : 'post',
    contentType        : 'application/json',
    headers            : {
      Authorization : 'Bearer ' + token
    },
    payload            : JSON.stringify(payload),
    muteHttpExceptions : true,
})
}
