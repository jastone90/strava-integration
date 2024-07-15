// Changelog
// bdk230117: define global var with name of data input sheet

var CLIENT_ID = '';
var CLIENT_SECRET = '';

var MILAGE_DATA_ROW_START = 10;

var DATA_SHEET = '2024';  // bdk230117 var data sheet

//Athlete data 
const ALAN = {
  id: 123456,           // Strava unique identifier 
  col: "C",               // Date Column
  col_end: "I",           // Column before calculated points
  strava_button_loc: "E3" // Location of strava import button
};
const ANOTHERATHELETE = {
  id: 123,
  col: "L",
  col_end: "R",
  strava_button_loc: "E4"
};

// called in log_mileage.gs 
function stravaIntegration(ui) {
  ui.createMenu('Import')
    .addItem('Strava','importStravaActivity')
    .addToUi();
}

//Alternative way to kick off Strava import
function stravaButtonClickAS(){ stravaButtonClick(ALAN);}
function stravaButtonClickBK(){ stravaButtonClick(BRIAN);}
function stravaButtonClickMWPA(){ stravaButtonClick(MARTIN);}
function stravaButtonClickWHB(){ stravaButtonClick(WILL);}

//Wrapper for alerts, catch getUI exception when running on a trigger
function alert(msg){
  try{
    SpreadsheetApp.getUi().alert(msg);
  } catch(err) {
    // do nothing
  }
}

//Listen for checkbox being checked
function stravaButtonClick(athlete){
  //var stravaButton = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data").getRange(athlete.strava_button_loc);     // bdk230117 var data sheet
    var stravaButton = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET).getRange(athlete.strava_button_loc); // bdk230117 var data sheet

  
  if (stravaButton.getValue() === true) {
    //importStravaActivity
    importStravaActivity();
    //Set back to false after
    stravaButton.setValue('FALSE');
  }
}

function importStravaActivity () {
  //Attempt to fetch strava activity info
  var response = callStravaAPI();


  if (typeof response === 'undefined') {
        alert("Retry the script after accepting OAuth stuff");
        return;
  }
  if(response.length > 1) {
    alert("More than one activity found");
  } else if(response.length == 0) {
    alert("no workouts found today");
  } else {
    //One activity found, proceed to updating sheet
    var activity = getStravaActivity(response[0]);
    updateSpreadSheet(activity);
  }
  
}

// Get rid of noise and convert to freedom units
function getStravaActivity (res) {
  var activity = {
    athlete: res.athlete.id,
    type: res.type,
    distance: (res.distance / 1609.344).toFixed(2), // m -> mi
    elevation: (res.total_elevation_gain * 3.28084).toFixed(0) // m -> ft
  };
  return activity;
}

function updateSpreadSheet(activity) {
  var range;

  switch(activity.athlete) {
    case ALAN.id:
      // Find the cells that need to be updated
      range = getAtheleteRange(ALAN);  
      break;    
    case BRIAN.id:
      range = getAtheleteRange(BRIAN); 
      break;
    case MARTIN.id:
      range = getAtheleteRange(MARTIN); 
      break;
    case WILL.id:
      range = getAtheleteRange(WILL);
      break;
    default:
      alert("you are under arrest");
      return;
  }

  //Finally update cells with strava activity info
  updateRow(activity, range);
}

function getAtheleteRange(athlete){
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");   // bdk230117 var data sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET); // bdk230117 var data sheet
  
  //Find the first empty row in the atheletes Date column 
  var col = sheet.getRange(athlete.col +"10:"+athlete.col+"500");
  var row = getFirstEmptyRow(col);

  //Find the range of cells that should be updated for this athlete
  return sheet.getRange(athlete.col + row + ":"+ athlete.col_end + row);
}

function updateRow(act, range){
  var data = range.getValues();
  data[0][0]= getDateWithoutTime();

  switch(act.type){
    case "VirtualRide":
      data[0][1]=true;
      data[0][4]=act.distance;
      data[0][5]=act.elevation;
      break;
    case "Ride":
      data[0][4]=act.distance;
      data[0][5]=act.elevation;
      break;
    case "Run":
      data[0][2]=act.distance;
      data[0][3]=act.elevation;
      break;
    default:
      alert("Screenshot this message and send to our signal chat: " +act.type);
      return;
  }

  range.setValues(data);
}

function getDateWithoutTime(){
  var tempDate = new Date();
  return tempDate.getMonth() + 1 + "/" + tempDate.getDate() + "/" + tempDate.getFullYear();
}

function getFirstEmptyRow(col) {
  var values = col.getValues(); 
  
  var i = 0;
  while ( values[i] && values[i][0] != "" ) {
    i++;
  }
  return (i+MILAGE_DATA_ROW_START); //  add 10 since activity entry starts on row 10
}

function getEpochDate () {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");   // bdk230117 var data sheet
  //var date = new Date(sheet.getRange("Data!A1").getValue());                  // bdk230117 var data sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET); // bdk230117 var data sheet
  var date = new Date(sheet.getRange( DATA_SHEET + "!A1").getValue());
  // var offset = new 
  // console.log (date);
  // var epochDate =(date - Date(1970,1,1)); //*86400;
  // console.log ("epoch date: " + Date.parse(date))

  //testing
  // return 1641674878;

  return Date.parse(date)/1000;  //strava wtf take my milliseconds
}


function callStravaAPI () {
  // set up the service
  var service = getStravaService();
  
  if (service.hasAccess()) {
    console.log('App has access.');
    var endpoint = 'https://www.strava.com/api/v3/athlete/activities';
    var params = '?after=' + getEpochDate ();

    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true
    };
    
    var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
    return response;
  } else {
    console.log("App has no access yet.");
    // open this url to gain authorization from strava 
    var authorizationUrl = service.getAuthorizationUrl();
    // console.log(authorizationUrl);
    // SpreadsheetApp.getUi().alert("Watch this informational video detailing oauth setup for strava:\n\n https://youtube.com/watch?v=USCi-NmSGkk"); 
    // SpreadsheetApp.getUi().alert("Put the youtube presentation into fullscreen mode and dose yourself"); 
    alert("Open the following URL and grant permissions:\n\n"+ authorizationUrl); 

  }
}

// configure the service
function getStravaService() {
  //LInk to OAuth readme : https://github.com/googleworkspace/apps-script-oauth2

  return OAuth2.createService('Strava')
    .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
    .setTokenUrl('https://www.strava.com/oauth/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('activity:read_all');
}

// handle the callback
function authCallback(request) {
  var stravaService = getStravaService();
  var isAuthorized = stravaService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}

