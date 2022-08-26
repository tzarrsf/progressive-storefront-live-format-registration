//Set this to true to send emails to the replyTo address not the one in the sheet
const debugMode = false;
const productName = "B2C Commerce";
const eventName = "PWA Kit & Managed Runtime Beta Workshop";
const sheetName = "Form Responses 1";

//TODO: Make this one Jonathan: jonathan.tucker@salesforce.com
const replyTo = "jonathan.tucker@salesforce.com";
const fyiList = "jonathan.tucker@salesforce.com"

/*
Assumed header text values in use:

- Partner company name
- Full Name
- Email Address
- Requested training date
- Country
- Have you completed the defined prerequisites and/or validated that you have the required skills (see top of form)? These are required to attend the training. 

Mapping all columns by index makes debugging Apps Script far easier along with getting older data when it might be needed later for responses to duplicates, etc.
*/

const partnerCompanyNameHeader = "Partner company name";
const fullNameHeader = "Full Name";
const emailAddressHeader = "Email Address";
const requestedTrainingDateHeader = "Requested training date";
const countryHeader = "Country";
const prerequisitesHeader = "Have you completed the defined prerequisites and/or validated that you have the required skills (see top of form)? These are required to attend the training.";
const partnerCompanyNameColumnIndex = getHeaderIndex(sheetName, partnerCompanyNameHeader);
const fullNameColumnIndex = getHeaderIndex(sheetName, fullNameHeader);
const emailAddressColumnIndex = getHeaderIndex(sheetName, emailAddressHeader);
const requestedTrainingDateColumnIndex = getHeaderIndex(sheetName, requestedTrainingDateHeader);
const countryColumnIndex = getHeaderIndex(sheetName, countryHeader);
const prerequisitesColumnIndex = getHeaderIndex(sheetName, prerequisitesHeader);

function onFormSubmit(e) {
  //Get the row of submitted data by names and other means where needed
  let submission = { 
    emailAddress: e.values[emailAddressColumnIndex]
    , partnerCompanyName: e.values[partnerCompanyNameColumnIndex]    
    , fullName: e.values[fullNameColumnIndex]
    , country: e.values[countryColumnIndex]
    , requestedTrainingDate: e.values[requestedTrainingDateColumnIndex]
    , prerequisitesMet: e.values[prerequisitesColumnIndex]
  };
  
  //You can avoid mistakes or spamming when first assigning this code to your form by using the debugMode variable at the top
  if(debugMode)
  {
    submission.emailAddress = replyTo;  
  }
  
  //Check if we have a record on file for the email address and same session time
  let alreadyRegisteredResult = isDuplicateRegistration(submission.emailAddress, submission.requestedTrainingDate);
  
  if(alreadyRegisteredResult.alreadyRegistered === true)
  {
    //Send notification of duplicate - we already have them for the email and time slot (requestedTrainingDate)
    MailApp.sendEmail({
      to: submission.emailAddress
      , replyTo: replyTo
      , subject: "Duplicate request for " + productName + " " + eventName
      , htmlBody: "Hi " + submission.fullName + ",<br /><br />It looks like you submitted a duplicate request for this event:<br /><br />"
      + makeRegistrationTable(submission) + "<br />"
      + "Your previous request for " + submission.requestedTrainingDate + " remains intact. You do not need to register again.<br /><br />"
      + "Thanks for your interest in the " + productName + " " + eventName + "."
    });
    Logger.log("Sent Duplicate registration request for " + productName + ": " + eventName + " to '" + submission.emailAddress + "'");

    //Send a Fwd email to internal staff
    MailApp.sendEmail({
      to: fyiList
      , replyTo: replyTo
      , subject: "FYI - Duplicate request for " + productName + " " + eventName
      , htmlBody: "Hi " + submission.fullName + ",<br /><br />It looks like you submitted a duplicate request for this event:<br /><br />"
      + makeRegistrationTable(submission) + "<br />"
      + "Your previous request for " + submission.requestedTrainingDate + " remains intact. You do not need to register again.<br /><br />"
      + "Thanks for your interest in the " + productName + " " + eventName + "."
    });
    Logger.log("Sent Fwd of Duplicate registration request for " + productName + ": " + eventName + " to '" + fyiList + "'");
  }
  else
  {
    //NOTE: If your form allows editing you will get an exception for not having a recipient here so don't allow editing in your form settings :)
    //Send Ack email to the person registering - this is a new registration
      MailApp.sendEmail({
      to: submission.emailAddress
      , replyTo: replyTo
      , subject: "Initial request for " + productName + " " + eventName
      , htmlBody: "Hi " + submission.fullName + ",<br /><br />This email is confirmation that we have received your request to attend this event:<br /><br />"
      + makeRegistrationTable(submission) + "<br />Thanks for your interest in the " + productName + " " + eventName + "."
    });
    Logger.log("Sent Registration acknowledgement for " + productName + ": " + eventName + " to '" + submission.emailAddress + "'");
    
    //Send an FYI email to internal staff - this is a new registration
      MailApp.sendEmail({
      to: fyiList
      , replyTo: replyTo
      , subject: "FYI - Initial request for " + productName + " " + eventName
      , htmlBody: "Hi " + submission.fullName + ",<br /><br />This email is confirmation that we have received your request to attend this event:<br /><br />"
      + makeRegistrationTable(submission) + "<br />Thanks for your interest in the " + productName + " " + eventName + "."
    });
    Logger.log("Sent Fwd Registration acknowledgement for " + productName + ": " + eventName + " to '" + fyiList + "'");
  }
}

function getHeaderIndex(sheetName, headerText)
{
  let result = -1;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if(sheet == null)
  {
    console.log('Sheet not located in getIndexByHeader using: ' + sheetName);
    return result;
  }

  let data = sheet.getDataRange().getValues();

  for (let i=0; i < data[0].length; i++)
  {
    if(data[0][i].toString() == headerText)
    {
      result = i;
      break;
    }
  }
  return result;
}

//Make a fungible old-school HTML table of name-value pairs we can embed in comms
function makeRegistrationTable(data)
{
  let registrationTable = "";
  registrationTable += "<b>Event</b>: " + productName + " " + eventName  + "<br />";
  registrationTable += "<b>Requested Date</b>: " + data.requestedTrainingDate + "<br />";
  registrationTable += "<b>Full Name</b>: " + data.fullName + "<br />";
  registrationTable += "<b>Country</b>: " + data.country + "<br />";
  registrationTable += "<b>Partner Company Name</b>: " + data.partnerCompanyName + "<br />";
  registrationTable += "<b>Prerequisites met</b>: " + data.prerequisitesMet + "<br />";
  return registrationTable;
}

//Detect a duplicate based on the email address and training slot (requestedTrainingDate)
function isDuplicateRegistration(email, requestedTrainingDate)
{
  Logger.log("isDuplicateRegistration(" + email + "," + requestedTrainingDate + ") invoked.");
  //The current submission does not get stopped so we need to just track it and kill the real dup looking at any matches added beyond a length of 1
  let hits = [];
  let sheet = SpreadsheetApp.getActiveSheet();
  let data  = sheet.getDataRange().getValues();
  let i = 0;
        
  for (i = 0; i < data.length; i++)
  {
    if (data[i][emailAddressColumnIndex].toString().toLowerCase() === email.toLowerCase()
    && data[i][requestedTrainingDateColumnIndex].toString().toLowerCase() === requestedTrainingDate.toLowerCase()){
      Logger.log("Pushing: " + JSON.stringify(data[i]));  
      hits.push(data[i]);
    }
  }
  
  //Return some slightly enriched results we can use making only 1 hit on the scan
  if(hits.length > 1)
  {
    Logger.log("Located duplicate email: <" + email + "> with training date : " + requestedTrainingDate.toLowerCase());
    //Remove the newest addition using the tracking array (like it never happened)
    sheet.deleteRow(i);
    //Get the data for the original (old) submission into something we can use
    let originalRow = hits[hits.length - 2];
    
    let priorRegistration = {
      emailAddress: originalRow.values[emailAddressColumnIndex]
      , partnerCompanyName: originalRow.values[partnerCompanyNameColumnIndex]    
      , fullName: originalRow.values[fullNameColumnIndex]
      , country: originalRow.values[countryColumnIndex]
      , requestedTrainingDate: originalRow.values[requestedTrainingDateColumnIndex]
      , prerequisitesMet: originalRow.values[prerequisitesColumnIndex]
    };
    
    return JSON.parse("{\"alreadyRegistered\": true, \"priorData\": " + JSON.stringify(priorRegistration) +"}")
  }
  else
  {
    return JSON.parse("{\"alreadyRegistered\": false, \"priorData\": null}");
  }
}
