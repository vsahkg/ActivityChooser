// -- Course Choices - By achiu@ais.edu.hk --
// This App displays logged-in student's credits (from spreadsheet) and displays them, 
// along with choices (from spreadsheet) for next yr's course selections.
//

// -- GLOBALS --
//var myListDocID = '1nOOS8pF3bZo4raMg88-l2O1KyaMmKCouK0ZCeM7a6Mw'; //used if separated from sheets
 var ListDoc = SpreadsheetApp.getActiveSpreadsheet();
 var SurveyDoc = ListDoc;


//Spreadsheet sheet names for Get Lists...
var myListSheetName = 'Choices';
var myStudentDataSheetName = 'StudentList'; //Sheet for authorising students
var myApproverListSheetName = 'Approver';

var myCreditsSheetName = 'Transcript';
var myGradReqsSheetName = 'GradReqs';

//Spreadsheet for saving student choices and Teacher Recommendations. 
//var mySurveyCollector = myListDocID; //'19_3IqX_uocu8mPKH3_GLIRKyp4pj7yBoWUsFRwXOlfY'; //Live collector in 2012-13 for 2013-14 choices
//var mySurveyCollector = '1LsiB1BFZgc-RjmQ0zFUw4cNvAkyz7nSnKqLmPT70tCY'; //Trial DEV Spreadsheet


var mySurveySheetName = 'LiveResults'; //Results sheet 
//var mySurveySheetName = 'Results-V2'; //Current results sheet with 2 Adv List Choices (obsolete - moved over to "results")
var mySurveyTripCounts = 'LiveCounts';
var mySurveyCourseCounts = 'CourseCounts';
var myRecommendationSheetName = 'Recommendation';

//Get User

var thisUser = Session.getActiveUser().getEmail(); //Logged In User


//Globals for displaying course options in each block
var coursePerLine = 3; //obsolete!
var courseFontSize = '9px';

var DisableDateLocking = false; //turns on or off the date lock out.
var showSecondChoice = true; //turns on or off for displaying a second choice.

//For HTMLService app (from template)
//var userSheetName = 'Members';
var challengeSheetName = 'Activity Selections';
var appTitle = 'Choose Earth Day Activity';
var entityTitle = 'Selections'; // used in titles throughout app

  var approverEmailCol = 0;
  var studentEmailCol = 0;
  var parentEmailCol = 12;
  var advisorEmailCol = 20;
  var endDateCol = 18; //Survey end date col on student sheet
  var startDateCol = 17; //survey start date col on student sheet
  var canPostCol = 16;

//
// -----
//  doGet - main function for web app
// -----
function doGet(){
  var myDoc = 'index';  
  return HtmlService.createTemplateFromFile(myDoc).evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
  
// ----
// function loadGInfo() - Gets all info from spreadsheets and passes it back to the client
// -----
function loadGInfo() {

  //var ListDoc = SpreadsheetApp.openById(myListDocID);
  //var SurveyDoc = SpreadsheetApp.openById(mySurveyCollector);
  var ListDoc = SpreadsheetApp.getActiveSpreadsheet();
  var SurveyDoc = ListDoc;
  
  //Check are we adventure week approver, HRM teacher or student.
  var userType = {isStudent: false, isAdvisor: false, isApprover: false, isParent: false};
  var error = {status: false, class: "bg-danger", msg: "Error"};
  
  var approverList = ListDoc.getSheetByName(myApproverListSheetName).getDataRange().getValues();
  var studentList = ListDoc.getSheetByName(myStudentDataSheetName).getDataRange().getValues();
  var choicesList = SurveyDoc.getSheetByName(mySurveySheetName).getDataRange().getValues();



  var myApprover = getRowsMatching(approverList, approverEmailCol, thisUser);
  var myAdvisor = getRowsIncluding(studentList, advisorEmailCol, thisUser);
  var myStudentInfo = getRowsMatching(studentList, studentEmailCol, thisUser);
  var myChildInfo = getRowsMatching(studentList, parentEmailCol, thisUser);
  var studentInfo = [];
  var choiceData = [];
  var canPost = false; 
  
  userType.isApprover = (myApprover.length > 0);//We are an approver
  userType.isAdvisor = (myAdvisor.length > 0); //We are an advisor
  userType.isStudent = (myStudentInfo.length > 0); //We are a student
  userType.isParent = (myChildInfo.length > 0); //We are a parent
  
  if (userType.isApprover) { //We are an approver

    studentInfo = textifyDates(studentList.slice(1)); //prepare student info
    choiceData = textifyDates(choicesList);
    //Logger.log(studentInfo);
    
  } else if (userType.isAdvisor) { //we are an advisor
    studentInfo = textifyDates(myAdvisor);
    choiceData = textifyDates(choicesList);
    //canPost = checkCanPost(myStudentInfo); 
  } else if (userType.isStudent) { //we are a student
      Logger.log(myStudentInfo);
    canPost = checkCanPost(myStudentInfo); 
    studentInfo = textifyDates(myStudentInfo);
    //Logger.log(canPost);
    //Logger.log(studentInfo);
    choiceData = textifyDates(getRowsMatching(choicesList,0,thisUser));
    
    //if(!myStudentInfo[0][17]) error = {status: true, class: "bg-danger", msg: "Choices not available for you."}
  } else if (userType.isParent) { //we are a student
    studentInfo = textifyDates(myChildInfo);
    for (var child = 0; child < studentInfo.length; child++){
      var myChildChoice = textifyDates(getRowsMatching(choicesList,0,studentInfo[child][0]));
      Logger.log(myChildChoice);
      if(myChildChoice.length > 0){
        for (var cc = 0; cc < myChildChoice.length; cc++){
          choiceData.push(myChildChoice[cc]);
        }
      }
    }
  } else {
    error.msg = "You don't have permissions to view anything here.";
    error.status = true;
  }
  
  //Get a list of possible choices for this user
  var possibleChoices = [];
  var possibleChoiceList = textifyDates(ListDoc.getSheetByName(myListSheetName).getDataRange().getValues().splice(1));
  /* for (var pc = 0; pc < possibleChoiceList.length; pc++){
    possibleChoices.push(possibleChoiceList[pc][0]);
  } */
  var tripCounts = getTripCounts();
  
  var refreshTime = Utilities.formatDate(new Date(), "GMT+08:00", "dd-MMM-yyyy hh:mm:ss")
  Logger.log(choiceData);
  return {studentInfo: studentInfo, choiceData: choiceData, userType: userType, possibleChoiceList: possibleChoiceList, canPost: canPost, error: error, tripCounts: tripCounts, refreshed: refreshTime, secondChoice: showSecondChoice, user: thisUser};
}

function getTripCounts(){
  //var SurveyDoc = SpreadsheetApp.openById(mySurveyCollector);
  var tripCounts = SurveyDoc.getSheetByName(mySurveyTripCounts).getDataRange().getValues();
  
  return tripCounts;
  
}

function checkCanPost(studentInfo){
  try{
    var timeOpen = timeCheck(new Date(studentInfo[0][startDateCol]), new Date(studentInfo[0][endDateCol]));
    var canPost = (studentInfo[0][canPostCol] && timeOpen); //Check "Survey Permitted" column for this student, and if timeOpen is TRUE (line above)
    return canPost;
  } catch(e){
    return false;
  }
  
}
//-----
// function getStudentInfo(myID) returns the student listing info for thisUser
//-----
function getStudentInfo(ListDoc, myID){

  var myStudentListsheet = ListDoc.getSheetByName(myStudentDataSheetName);
  var studentInfo = getRowsMatching(myStudentListsheet.getDataRange().getValues(),1,myID);
  Logger.log(studentInfo);
  return studentInfo;
}

//-----
// function getStudentInfo(myID) returns all student listing info if thisUser is in Principals list
//-----
function getPrincipalInfo(myID){

  var studentInfo = [];
  //var myPrincipalList = SpreadsheetApp.openById(myListDocID).getSheetByName(myApproverListSheetName).getDataRange().getValues();
  var myPrincipalList = ListDoc.getSheetByName(myApproverListSheetName).getDataRange().getValues();
  for (var i =0; i < myPrincipalList.length; i++){
    //if (myID === myPrincipalList[i][0]) studentInfo = SpreadsheetApp.openById(myListDocID).getSheetByName(myStudentDataSheetName).getDataRange().getValues().splice(1);  
    if (myID === myPrincipalList[i][0]) studentInfo = ListDoc.getSheetByName(myStudentDataSheetName).getDataRange().getValues().splice(1);  
  }
  return studentInfo;
}


//--------
// studentPostData(data) - expects choice data from the interface - posts choices to sheet
//--------
function studentPostData(data, sindex) {
   // var userType = data.userType; // Get user type from data
    var emailToUse = thisUser; // Default to the logged-in user email
  /*  if (userType.isParent) {
        emailToUse = data.parentEmail; // Use parent's email if user is a parent
    }
  */
  //var ListDoc = SpreadsheetApp.openById(myListDocID);
  //var SurveyDoc = SpreadsheetApp.openById(mySurveyCollector);
  
  var msg = "";
  var studentList = ListDoc.getSheetByName(myStudentDataSheetName).getDataRange().getValues(); 
  var myStudentInfo = getRowsMatching(studentList, studentEmailCol, thisUser);

  if(myStudentInfo.length > 0){//check we can find this student user
    //Check student can post
    var canPost = checkCanPost(myStudentInfo);
    var myChoices = getRowsMatching(SurveyDoc.getSheetByName(mySurveySheetName).getDataRange().getValues(),1,thisUser);//Get Current Choices
    var hrmApproved = false;
    var principalApproved = false;  
    if(myChoices.length > 0){
      hrmApproved = myChoices[0][4];
      principalApproved = myChoices[0][9];
    }
    
    if (canPost && !(hrmApproved) && !(principalApproved)){ //If student can post (survey is open) and principal / HRM have not approved
      var myC = new Array();  
      myC.push([emailToUse, 
                  thisUser,
                  new Date(),
                  data.hrm,
                  data.advisorChecked,
                  data.choice1,
                  data.choice2,
                  data.choice3,
                  data.approvalChange,
                  data.approvalDate,
                  data.principalChecked,
                  data.paid,
                  data.house]);
   
      return postData(myC, SurveyDoc, true, sindex, canPost); // Post data with the correct email
    } else {
      msg = "No permissions to post - refresh to check status.";
    }
  } else {
    msg = "An error occured - student not found."
  }
  return {error: {status: true, msg: msg, class: "bg-danger", log: textifyDates(myStudentInfo)}};
}

//--------
// ParentPostData(data) - expects choice data from the interface - posts choices to sheet
//--------
function parentPostData(data, sindex) {
  var studentEmailToUse = data.id; // Use the student's ID from their choice data object.
  //var ListDoc = SpreadsheetApp.openById(myListDocID);
  //var SurveyDoc = SpreadsheetApp.openById(mySurveyCollector);
  var msg = "";
  var studentList = ListDoc.getSheetByName(myStudentDataSheetName).getDataRange().getValues(); 
  var myStudentInfo = getRowsMatching(studentList, studentEmailCol, studentEmailToUse);
  

 if(myStudentInfo.length > 0){//check we can find this student user
   //Check student can post AND Check the current user is the parent of this student
   var canPost = checkCanPost(myStudentInfo) && (myStudentInfo[0][parentEmailCol] == thisUser);
   var myChoices = getRowsMatching(SurveyDoc.getSheetByName(mySurveySheetName).getDataRange().getValues(),1,studentEmailToUse);//Get Current Choices
   var hrmApproved = false;
   var principalApproved = false;  
   if(myChoices.length > 0){
     hrmApproved = myChoices[0][4];
     principalApproved = myChoices[0][9];
   }
   
   if (canPost && !(hrmApproved) && !(principalApproved)){ //If student can post (survey is open) and principal / HRM have not approved
     var myC = new Array();  
     myC.push([studentEmailToUse, 
                 thisUser,
                 new Date(),
                 data.hrm,
                 data.advisorChecked,
                 data.choice1,
                 data.choice2,
                 data.choice3,
                 data.approvalChange,
                 data.approvalDate,
                 data.principalChecked,
                 data.paid,
                 data.house]); 
  
     return postData(myC, SurveyDoc, true, sindex, canPost); // Post data with the correct email
   } else {
     msg = "No permissions to post - refresh to check status.";
   }
 } else {
   msg = "An error occured - student not found."
 }
 return {error: {status: true, msg: msg, class: "bg-danger", log: textifyDates(myStudentInfo)}};
}

//--------
// approverPostData(data) - expects choice data from the interface - posts choices to sheet
//--------
function approverPostData(data, sindex){
  
  var passTest = true;
  //Check we have permission to approve
  passTest = getPrincipalInfo(thisUser).length > -1; 
  //check choices data exists and has been filled 

  if(passTest){   
    var mySurveyFile = SurveyDoc;
      var myC = new Array();  
      myC.push([data.id, 
                  data.email,
                  data.timeDate,
                  data.hrm,
                  data.principalChecked,
                  data.choice1,
                  data.choice2,
                  data.choice3,
                  thisUser,
                  new Date(),
                  data.principalChecked,
                  data.paid,
                  data.house]);
      return postData(myC, mySurveyFile, false, sindex, false);
  } else {
      return {error: {status: true, msg: "No posting permissions", class: "bg-danger"}};
  }
}

//-------
//DO POST
//-------

function postData(myC, SurveyDoc, checkQuota, sindex, canPost) {

  var sheet = SurveyDoc.getSheetByName(mySurveySheetName);
  var countSheet = SurveyDoc.getSheetByName(mySurveyTripCounts);
  var statusMessage = "";
  var problemEnc = false;
  
  //Get Lock
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);

  //Get the survey data.
  var lastRow = sheet.getLastRow();
  var sheetData = sheet.getRange(1,1,lastRow,14).getValues();
  var thisRow = getExistingRow(sheetData,0,myC[0][0]);  
  var nowTime = new Date();
    
  //If record already exists, clear it first (this is to ensure that we get accurate course spot counts without including this user's previous choice)
  if (thisRow < lastRow){   
    var clearRange = sheet.getRange(thisRow+1, 1, 1, 8).setValues([[myC[0][0], myC[0][1], myC[0][2], myC[0][3], myC[0][4], '', '', '']] );    
    //sheet.deleteRow(thisRow+1);
  }
  
  if(checkQuota){//If we are checking quota
    //Check not over quota. Get quota info...
    var myCounts = countSheet.getDataRange().getValues();
    // For each course entered, if spots taken greater or equal spots available, blank that course and add error message.
    for (var x=5; x <= 7; x++){
      var RowX = ArrayLib.indexOf(myCounts, 0, myC[0][x]);
      if (RowX > -1){ 
        if (myCounts[RowX][2] >= myCounts[RowX][1]) {
          statusMessage += myC[0][x] + ' already full! Please select another choice! '; 
          problemEnc = true;
          myC[0][x] = '';
        }
      }
    } 
  }

  var targetRange = sheet.getRange(thisRow+1, 1, 1, myC[0].length).setValues(myC);    
    
  SpreadsheetApp.flush();
  // clean up and release the lock
  lock.releaseLock();
  
  //var canPost = checkCanPost(thisUser);
  
  var choice = {};
  var myChoices = textifyDates(getRowsMatching(sheet.getDataRange().getValues(),0,myC[0][0]));
  if (myChoices.length > 0) {
    choice = { id: myChoices[0][0],
                       email: myChoices[0][1],
                       timeDate: myChoices[0][2],
                       hrm: myChoices[0][3],
                       advisorChecked: myChoices[0][4],
                       choice1: myChoices[0][5],
                       choice2: myChoices[0][6],
                       choice3: myChoices[0][7],
                       approvalChange: myChoices[0][8],
                       approvalDate: myChoices[0][9],
                       principalChecked: myChoices[0][10],
                       paid: myChoices[0][11],
                       canPost: canPost,
                       house: myChoices[0][12]
                        }
  } else {
    problemEnc = true;
    statusMessage = "Could not reload your data. Try reloading this page.";
  }

  return {error: {status: problemEnc, msg: statusMessage, class: "bg-danger", log: myChoices}, choice: choice, sindex: sindex};
}





//-----
// function getStudentInfo(myID) returns the student listing info for thisUser
//-----
function getStudentInfo(myID){

  //var myStudentListsheet = SpreadsheetApp.openById(myListDocID).getSheetByName(myStudentDataSheetName);  
  var myStudentListsheet = ListDoc.getSheetByName(myStudentDataSheetName);  
  var LastSsRow = myStudentListsheet.getLastRow()-1;
  var studentInfo = getRowsMatching(myStudentListsheet.getRange(2, 1, LastSsRow,myStudentListsheet.getLastColumn()).getValues(),1,myID);

  return studentInfo;
}


//-----
// getExistingRow - returns the row that contains the data matching the criteria, or returns the next row in the spreadsheet.
//
//-----
function getExistingRow(myList,checkCol,checkCriteria){
  var myRow = 0;
  while (myRow < myList.length){
    if (myList[myRow][checkCol] == checkCriteria) {
      return myRow;
    }
    myRow++;
  } 
  return myRow;
}


function myClickHandler(e) {
  var app = UiApp.getActiveApplication();

  var label = app.getElementById('statusLabel');
  label.setVisible(true);

  app.close();
  return app;
}

function usableColValues(coldata, lastrow) {
  for( var i = (lastrow - 1) ; i > 0; i--){

    if(coldata[i] != "") {
      return coldata.slice(0,i+1);
      };
  };
  return coldata;
}



//getRowsMatching takes a data list and searches the sortIndex for all values that match valueToFind, returning the rows that match this value

function getRowsMatching(DataList, sortIndex, valueToFind){
  
  var foundList = new Array();
  var myDataList = DataList.slice(0); //Added this line and changed the function parameter from myDataList to DataList - hoping to stop changes to original array
  myDataList.sort(function(a, b){ //Sort the items by sortIndex
    var x = a[sortIndex];
    var y = b[sortIndex];
    return (x < y ? -1 : (x > y ? 1 : 0));});
  
  var cdr = 0;
  var found = false; 

  while ( cdr < myDataList.length){
    if (myDataList[cdr][sortIndex] == valueToFind) {
      found=true;
      foundList.push(myDataList[cdr])
    }
    else if (found){
      return foundList;
    }
    cdr++;
  }

  return foundList;
  
}

//getRowsIncluding takes a data list and searches the sortIndex for values that include valueToFind, returning the rows that match this value. It may be less efficient than getRowsMatching because it goes through all rows.

function getRowsIncluding(DataList, sortIndex, valueToFind){
  
  var foundList = new Array();
  var myDataList = DataList.slice(0); //Added this line and changed the function parameter from myDataList to DataList - hoping to stop changes to original array
  
  for (var cdr = 0; cdr < myDataList.length; cdr++){
    if (myDataList[cdr][sortIndex].includes(valueToFind)) {
      found=true;
      foundList.push(myDataList[cdr])
    }
  }
  return foundList;
  
}


//-----
// function timeCheck (studentInfo)
//-----
function timeCheck(openTime, closeTime) {
  //Logger.log([openTime, closeTime, (new Date() >= openTime)]);
  return DisableDateLocking || ((new Date() >= openTime) && (new Date() <= closeTime));
}

// -----
// include - include files
// -----
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

// ------------------------------------ USEFUL FUNCTIONS--------------------------------------------------
//

//-----
// function textifyDates(myArr) - converts all dates into text format - assumes a 2D array as an input, returns the array.
//-----
function textifyDates(myArr){
  
  
  for(var r=0; r < myArr.length; r++){
    for(var c=0; c < myArr[r].length; c++){
      if (Object.prototype.toString.call(myArr[r][c]) === '[object Date]'){
        try {           
          //myArr[r] = myArr[r].toString();
          myArr[r][c] = Utilities.formatDate(myArr[r][c], "GMT+08:00", "dd-MMM-yyyy hh:mm:ss")
        } 
        catch(err) { myArr[r][c] = err};
      }
    }
  }
  return myArr;
}

//-----
// function textifyDates(myArr) - converts all dates into text format - assumes a 2D array as an input, returns the array.
//-----
function textifyDates1D(myArr){
  
  for(var r=0; r < myArr.length; r++){
      if (Object.prototype.toString.call(myArr[r]) === '[object Date]'){
        try {
          //myArr[r] = myArr[r].toString();
          myArr[r] = Utilities.formatDate(myArr[r], "GMT+08:00", "dd-MMM-yyyy")
          } 
        catch(err) { myArr[r] = err};
      }
    
  }
  return myArr;
}
