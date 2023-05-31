/**
When a specified form is submitted,
triggers a series of functions that result in a slide deck created
with values from the submitted form response.
*/

// ID of slide deck template
const SLIDES_TEMPLATE_FILE_ID = '<REPLACE-THIS-WITH-THE-ID-OF-YOUR-SLIDE-DECK-TEMPLATE-FILE>'
// ID of folder to save to
const LIGHTNING_TALKS_SLIDES_FOLDER_ID = '<REPLACE-THIS-WITH-THE-ID-OF-YOUR-FOLDER>'
// URL of spreadsheet associated with form responses
const LIGHTNING_TALKS_SHEET = SpreadsheetApp.openByUrl('<REPLACE-THIS-WITH-THE-URL-OF-YOUR-FORM-RESPONSE-SLIDESHEET>');

function formatFormResponse(formResponse) {
  // Create key-value pairs from form submission event values
  
  console.log("Generating form response key-value pairs from form response with key 'name': " + formResponse[2]);
  
  var formResponseObject = {};
  formResponseObject['timestamp'] = formResponse[0];
  formResponseObject['name'] = formResponse[1];
  formResponseObject['topic'] = formResponse[2];
  formResponseObject['title'] = formResponse[3];
  formResponseObject['tag_line'] = formResponse[4];
  formResponseObject['email'] = formResponse[5];

  console.log("Form response key-value pairs created")  
  
  return formResponseObject
}

function createNewSlideDeck(formResponseObject) {
  // Create a new slide deck by copying the template slide deck

  console.log("Creating a new slide deck by copying the template slide deck with form response with key 'name': " + formResponseObject['name'])
  
  // Get the template file
  var file = DriveApp.getFileById(SLIDES_TEMPLATE_FILE_ID);

  // Get the folder we want to save our new slide deck in and create a copy of the template there
  var folder = DriveApp.getFolderById(LIGHTNING_TALKS_SLIDES_FOLDER_ID);
  var filename = formResponseObject['name'] + '-lightning-talk'
  var newSlideDeck = file.makeCopy(filename, folder);

  console.log("New slide deck created with filename: " + filename)

  return newSlideDeck
}

function populateNewSlideDeck(formResponseObject) {
  // Populate the new slide deck with the form responses

  console.log("Populating the new slide deck using the form response with key 'name': " + formResponseObject['name']) 

  // Create the copy
  var newSlideDeck = createNewSlideDeck(formResponseObject) 
  // Open the new slide deck
  var openedSlideDeck = SlidesApp.openById(newSlideDeck.getId())
  // Get the first slide in the deck
  var titleSlide = openedSlideDeck.getSlides()[0];
  
  // Replace all instances of the template variables with the corresponding form response
  titleSlide.replaceAllText('{{title}}', formResponseObject['title']); 
  titleSlide.replaceAllText('{{tag_line}}', formResponseObject['tag_line']);  
  titleSlide.replaceAllText('{{name}}', formResponseObject['name']);  

  // Save and close the slide deck to persist our changes
  openedSlideDeck.saveAndClose();

  console.log("Slide deck saved and closed")

  // Add the form respondent's email as editor for the slide deck
  DriveApp.getFileById(newSlideDeck.getId()).addEditor(formResponseObject['email']);

  console.log("Editor added to slide deck")

  return newSlideDeck
}

function sendEmailWithSlideDeck(formResponseObject, newSlideDeck) {
  // Send email to slide deck respondent so they know their slide deck has been created

    console.log("Sending email to: " + formResponseObject['name'])
  
  // URL for the slide deck
  var slideDeckUrl = newSlideDeck.getUrl();
  // Link to the informational document about lightning talks
  var lightningTalkDoc = '<REPLACE-THIS-WITH-THE-URL-TO-GOOGLE-DOC>'
  
  // Send an email to the form respondent with the following values
  var emailAddress = formResponseObject['email'];
  var emailSubject = "Your lightning talk slide deck";
  var emailMessageBody = `Your lightning talk slide deck has been created here: ${slideDeckUrl}. \n\nFor more information on lightning talks, click here: ${lightningTalkDoc} \n\nPlease reach out if you have any problems or questions!`;
  GmailApp.sendEmail(emailAddress, emailSubject, emailMessageBody);

  console.log("Email sent")
}

function formWorkflow(e) {
  // Final workflow for the process:

  console.log("Starting workflow for form response with name: " + e.values[1]);

  // 1. Create key-value pairs from form response event
  var formResponseObject = formatFormResponse(e.values);

  // 2. Create a new slide deck and populate it with the key-value pairs
  newSlideDeck = populateNewSlideDeck(formResponseObject);

  // 3. Send the email
  sendEmailWithSlideDeck(formResponseObject, newSlideDeck);

  console.log("Workflow completed");
}

function main() {
  // Build the trigger that executes formWorkflow when a form is submitted to the linked spreadsheet

  console.log("Building trigger ...")

  ScriptApp.newTrigger('formWorkflow')
    .forSpreadsheet(LIGHTNING_TALKS_SHEET)
    .onFormSubmit()
    .create();
}