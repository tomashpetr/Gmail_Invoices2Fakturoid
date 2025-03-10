// Initiate the spreadsheet objects for data logging

let sheetLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FakturoidLog');
let sheetProcessedEmls = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EmlsRecords');

// Initiate the config values

function getConfigValues() {

  var configJSON = {
    "clientId": 'xxxx',     // Replace with your own user-level client_id, client_secret from fakturoid.cz
    "clientSecret": 'yyy',     // Replace with your own user-level client_id, client_secret from fakturoid.cz
    "accountSlug": 'mycompany', // Define Fakturoid account slug (mycompany if domain is app.fakturoid.cz/mycompany/)
    "urlMatchPattern": /(https?:\/\/[^<>\s"()]+invoice%2Fpdf[^<>\s"()]*)/, //define pattern to search in email body for the URL (depend)
    "useragentcustom": "Your_app_name (your_email@gmail.com)", // define fakturoid HTTP user agent name - required by fakturoid API:
    "GmailQuery": 'from:receipts@bolt.eu subject:"your bolt drive" newer_than:10d' //for Bolt invoices (adjust as needed)
  };

  return configJSON;

}


/**
 * 1) Retrieve access token using Client Credentials Flow (v3).
 */

function getFakturoidAccessTokenClientCredentials(configVals) {

  var tokenUrl = 'https://app.fakturoid.cz/api/v3/oauth/token';

  // Basic Auth header = 'Basic ' + base64("client_id:client_secret")
  var authHeader = 'Basic ' + Utilities.base64Encode(configVals['clientId'] + ':' + configVals['clientSecret']);

  var optionsGetToken = {
    method: 'post',
    payload: JSON.stringify(
      {
        'grant_type': 'client_credentials'
      }
    ),
    muteHttpExceptions: true,
    headers: {
      'Authorization': authHeader,
      'User-Agent': configVals['useragentcustom'], // Required by Fakturoid
      'Accept': 'application/json',
      'Content-Type': 'application/json',
    },
    muteHttpExceptions: true
  };

  var httprequest = UrlFetchApp.getRequest(tokenUrl, optionsGetToken)
  Logger.log('Fakturoid HTTP request: ' + JSON.stringify(httprequest));

  Logger.log('Fakturoid token call | URL : ' + tokenUrl);
  Logger.log('Fakturoid token call | custom agent: ' + configVals['useragentcustom']);
  Logger.log('Fakturoid token call | Call API URL start...');


  var response = UrlFetchApp.fetch(tokenUrl, optionsGetToken);
  var code = response.getResponseCode();
  var body = response.getContentText();

  Logger.log('Fakturoid token call | response code: ' + code);
  Logger.log('Fakturoid token call | response body: ' + body);
  Logger.log('Fakturoid token call | Response Headers: ' + JSON.stringify(response.getAllHeaders()));

  if (code >= 200 && code < 300) {
    var json = JSON.parse(body);
    return json.access_token;
  } else {
    throw new Error('Fakturoid token call |Failed to get token. Code: ' + code + ', Body: ' + body);
  }
}


/**
 * 2) Main function to:
 *    - Find Bolt invoices in Gmail
 *    - Extract PDF link
 *    - Upload to Fakturoid as an expense
 */
function processBoltInvoices() {
  // 1) Get the Bearer token for Fakturoid

  //initialization of local variables - taken over from config file (global var defined there)

var configVals = getConfigValues();

  var accountSlugInst = configVals['accountSlug'];
  var urlMatchPatternInst = configVals['urlMatchPattern'];
 // var clientIdInst = configVals['clientId']; not used
 // var clientSecretInst = configVals['clientSecret']; not used

  // call fakturoid function to get access token
  var accessToken = getFakturoidAccessTokenClientCredentials(configVals);

  // 2) Define Fakturoid account slug (mycompany if domain is app.fakturoid.cz/mycompany/)
  // will use accountSlug variable passed via the function call (takenover from config file - global variable);

  /*  avoid using labels for marking it in gmail
  // define or fetch the label object for processed messages
  var processedLabelName = 'bolt_inv_processed1';
  var processedLabel = GmailApp.getUserLabelByName(processedLabelName);
  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(processedLabelName);
  }
  */

  // 3) Gmail search query 
  var searchQuery = configVals['GmailQuery'];
  var threads = GmailApp.search(searchQuery);
  Logger.log('Script control | Gmail - found ' + threads.length + ' Gmail thread(s) for: ' + searchQuery);

  const existingIds = sheetProcessedEmls.getDataRange().getValues().map(row => row[0]); // Column A

  // 4) Loop through matching emails
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var msg = messages[j];

      /* // Check if this message already has our processed label
      var hasLabel = false;
      var labels = msg.getLabels();
      for (var lb = 0; lb < labels.length; lb++) {
        if (labels[lb].getName() === processedLabelName) {
          hasLabel = true;
          break;
        }
      }
      */
      
      
      // Skip if already recorded in Gsheet list "EmlsRecords" => the message is considered as already processed before
      if (existingIds.includes(msg.getId())) {
        Logger.log('Script control | thread ' + i + ' - message ' + j + ' - Skipping (already processed) msg: ' + msg.getId());
        continue;
      }
      

      // Extract PDF link from email body
      var body = msg.getBody();
      var invoiceLink = extractInvoiceLink(body, urlMatchPatternInst);
      if (!invoiceLink) {
        // if no URL link is found in the email
        Logger.log('Script control | thread ' + i + ' - message ' + j + ' -  No invoice link found in msg: ' + msg.getId());

        sheetLog.appendRow([
          new Date(),
          msg.getId(),
          'No invoice link found in msg',
          'Response code - Skipped fakturoid API',
          'Fakturoid response ID - Skipped fakturoid API',
          'Script control | thread ' + i + ' - message ' + j + ' -  No invoice link found in msg: ' + msg.getId()
        ]);

        // Mark message as processed (will be skipped next time - no link found)
        sheetProcessedEmls.appendRow([msg.getId(),new Date()]);
        continue;
      }

      try {
        // Download PDF
        Logger.log('Script control | thread ' + i + ' - message ' + j + ' -  Invoice link found: ' + invoiceLink);
        var responsePdf = UrlFetchApp.fetch(invoiceLink);
        var pdfBlob = responsePdf.getBlob().setName('Bolt-Invoice-' + msg.getId() + '.pdf');
        Logger.log('Script control | thread ' + i + ' - message ' + j + ' - PDF downloaded (size: ' + pdfBlob.getBytes().length + ' bytes).');


        // Log the result in the sheetLog.
        sheetLog.appendRow([
          new Date(),
          msg.getId(),
          'Invoice link: ' + invoiceLink,
          'Invoice (PDF) link HTTP response: ' + responsePdf.getResponseCode(),
          'Fakturoid response ID - no call to fakturoid  yet',
          'Script control | thread ' + i + ' - message ' + j + ' - PDF downloaded (size: ' + pdfBlob.getBytes().length + ' bytes).'
        ]);


        // 5) Create expense with attachment in Fakturoid
        sendPDFToFakturoidInbox(configVals, accessToken, pdfBlob, msg, invoiceLink, i, j);


        // Mark message as processed
        //msg.addLabel(processedLabel);
        sheetProcessedEmls.appendRow([msg.getId(),new Date()]);

      } catch (err) {
        // Log any errors - to console
        Logger.log('Error handling PDF for msg ' + msg.getId() + ': ' + err);
        //msg.star(); // or omit if you want to retry next time
        // Log any errors - to gsheet
        sheetLog.appendRow([new Date(), msg.getId(), invoiceLink, 'ERROR', err.message, '']);
      }
    }
  }
}


/**
 * Helper: Actually POST the expense + PDF to Fakturoid using Bearer token.
 */
function sendPDFToFakturoidInbox(configVals, accessToken, pdfBlob, msg, invoiceLink, i, j) {
  var fakturoidUrl = 'https://app.fakturoid.cz/api/v3/accounts/' + configVals['accountSlug'] + '/inbox_files.json';

  // 1) Convert the PDF blob to a Base64 string
  var base64Content = Utilities.base64Encode(pdfBlob.getBytes());

  // 2) Build the data URI: "data:application/pdf;base64,<base64>"
  //    Fakturoid uses this format to recognize the file.
  var dataUri = 'data:application/pdf;base64,' + base64Content;
  //Logger.log('Prepared PDF payload - long text ' + dataUri + "==");

  // 3) Prepare the JSON body
  var requestBody = {
    'attachment': dataUri,
    'filename': pdfBlob.getName() || 'bolt-invoice ' + msg.getId + '.pdf',
    // 'type': 'application/pdf' // optional if the extension is .pdf,
    //"send_to_ocr": false
  };

  var requestBodyOLD = {
    'page': 1,
  };


  // 4) Make the POST request with JSON
  var options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + accessToken,
      'User-Agent': configVals['useragentcustom'], // Required by Fakturoid
      'Content-Type': 'application/json'
    }
  };


  Logger.log('Fakturoid inbox call | URL : ' + fakturoidUrl);
  Logger.log('Fakturoid inbox call | custom agent: ' + configVals['useragentcustom']);
  Logger.log('Fakturoid inbox call | Call API URL start...');

  var response = UrlFetchApp.fetch(fakturoidUrl, options);
  var code = response.getResponseCode();
  var body = response.getContentText();

  Logger.log('Fakturoid inbox call | Create expense for msg: ' + msg.getId() + ' => code: ' + code);
  Logger.log('Fakturoid inbox call | Response: ' + body);

  if (code >= 200 && code < 300) {
    Logger.log('Fakturoid inbox call | Success! Inbox entry created for Bolt invoice.');
    // Log the result in the sheetLog.
    sheetLog.appendRow([
      new Date(),
      msg.getId(),
      invoiceLink,
      'Fakturoid API call result code: ' + code,
      'Fakturoid API call result text: ' + body,
      'Script control | thread ' + i + ' - message ' + j + ' - Fakturoid inbox call | Success! Inbox entry created for Bolt invoice.'
    ]);

  } else {
    // Log the result in the sheetLog.
    sheetLog.appendRow([
      new Date(),
      msg.getId(),
      invoiceLink,
      'Fakturoid API call result code: ' + code,
      'Fakturoid API call result text: ' + body,
      'Script control | thread ' + i + ' - message ' + j + ' - Fakturoid inbox call | Expense creation failed.'
    ]);

    throw new Error('Fakturoid inbox call | Inbox entry creation failed. Code: ' + code + ', Body: ' + body);
  }
}


/**
 * Helper: Extract PDF link from email body. pattern taken from config
 */
function extractInvoiceLink(emailBody, urlMatchPattern) {
  // Example pattern if the link includes "invoice%2Fpdf"
  // and we want to avoid trailing ">" or whitespace.
  var match = emailBody.match(urlMatchPattern);
  return match ? match[1] : null;
}
