/* Get and declare the user's previously saved headings information */
var userProperties = PropertiesService.getUserProperties();
var firstName = userProperties.getProperty('first-name');
var lastName = userProperties.getProperty('last-name');
var teacherName = userProperties.getProperty('teacher-name');
var courseName = userProperties.getProperty('course-name');

/* Sets up a navigable menu when the user opens the document */
function onOpen() {
  DocumentApp.getUi() 
    .createMenu("Essay")
    .addItem("Format", "format")
    .addItem("Set Info", "openSidebar")
    .addItem("About", "openAbout")
    .addToUi();
}

/* Opens a sidebar where the user can fill in their information */
function openSidebar() {
  var html = HtmlService.createTemplateFromFile("sidebar")
    .evaluate()
    .setTitle("Essay Formatter")
  DocumentApp.getUi()
    .showSidebar(html);
}

/* Opens a sidebar with instructions on how to use the Essay Formatter */
function openAbout() {
  var html = HtmlService.createTemplateFromFile("about")
    .evaluate()
    .setTitle("About")
  DocumentApp.getUi()
    .showSidebar(html);
}

/* Used to insert html from one file into another html file */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/* Saves the entered information to the user of the add-on */
function saveInfo(firstName, lastName, teacherName, courseName) {
  userProperties.setProperties({
    "first-name": firstName,
    "last-name": lastName,
    "teacher-name": teacherName,
    "course-name": courseName
  });
}

/* Updates the document according to MLA format, using the user's saved information
 * Alerts the user if there is missing information
 */ 
function format() {
  if(firstName=="null" || lastName=="null" || teacherName=="null" || courseName=="null") {
    alertUser();
  }
  else {
    var doc = DocumentApp.getActiveDocument();
    var headerStyle = {};
    headerStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
    headerStyle[DocumentApp.Attribute.FONT_FAMILY] = "Times New Roman";
    headerStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
    var bodyStyle = {};
    bodyStyle[DocumentApp.Attribute.LINE_SPACING] = 2.0;
    bodyStyle[DocumentApp.Attribute.FONT_FAMILY] = "Times New Roman";
    bodyStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
    doc.getHeader()
      .insertParagraph(0, lastName+' 1')
      .setAttributes(headerStyle);
    doc.getBody()
      .insertParagraph(0, firstName+' '+lastName+'\n'+teacherName+'\n'+courseName+'\n'+getDateHeading_())
      .setAttributes(bodyStyle);
  }
}

/* Prompts the user to complete setting up the headings, and opens the Set Info sidebar if user agrees. */
function alertUser() {
  var ui = DocumentApp.getUi();
  var response = ui.alert("Missing Information", "Please set up your headings before continuing.", ui.ButtonSet.OK_CANCEL);
  if(response == ui.Button.OK) {
    openSidebar();
  }
}

/* Returns the current date in MLA format.
 * @return (String) The date in MLA format
 */
function getDateHeading_() {
  var d = new Date();
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  return d.getDate()+' '+months[d.getMonth()]+' '+d.getFullYear();
}
