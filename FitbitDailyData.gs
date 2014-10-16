/* FitbitDownload.gs
This script will access your Fitbit data via the Fitbit API and insert it into a Google spreadsheet.
The first row of the spreadsheet will be a header row containing data element names. Subsequent rows will contain data, one day per row.
Note that Fitbit uses metric units (weight, distance) so you may wish to convert them.

Original script by loghound@gmail.com
Original instructional video by Ernesto Ramirez at http://vimeo.com/26338767
Modifications by Mark Leavitt (PDX Quantified Self organizer) www.markleavitt.com
Here's to your (quantified) health!
*/

// Key of ScriptProperty for Firtbit consumer key.
var CONSUMER_KEY_PROPERTY_NAME = "fitbitConsumerKey";
// Key of ScriptProperty for Fitbit consumer secret.
var CONSUMER_SECRET_PROPERTY_NAME = "fitbitConsumerSecret";
// Default loggable resources (from Fitbit API docs).
var LOGGABLES = ["activities/log/steps", "activities/log/distance",
    "activities/log/activeScore", "activities/log/activityCalories",
    "activities/log/calories", "foods/log/caloriesIn",
    "activities/log/minutesSedentary",
    "activities/log/minutesLightlyActive",
    "activities/log/minutesFairlyActive",
    "activities/log/minutesVeryActive", "sleep/timeInBed",
    "sleep/minutesAsleep", "sleep/minutesAwake", "sleep/awakeningsCount",
    "body/weight", "body/bmi", "body/fat",];

// function authorize() makes a call to the Fitbit API to fetch the user profile    
function authorize() {
    var oAuthConfig = UrlFetchApp.addOAuthService("fitbit");
    oAuthConfig.setAccessTokenUrl("https://api.fitbit.com/oauth/access_token");
    oAuthConfig.setRequestTokenUrl("https://api.fitbit.com/oauth/request_token");
    oAuthConfig.setAuthorizationUrl("https://api.fitbit.com/oauth/authorize");
    oAuthConfig.setConsumerKey(getConsumerKey());
    oAuthConfig.setConsumerSecret(getConsumerSecret());
    var options = {
        "oAuthServiceName": "fitbit",
        "oAuthUseToken": "always",
    };
    // get the profile to force authentication
    Logger.log("Function authorize() is attempting a fetch...");
    try {
       var result = UrlFetchApp.fetch("https://api.fitbit.com/1/user/-/profile.json", options);
       var o = Utilities.jsonParse(result.getContentText());
       return o.user;
    }
    catch (exception) {
       Logger.log(exception);
       Browser.msgBox("Error attempting authorization");
       return null;
    }
}

// function setup accepts and stores the Consumer Key, Consumer Secret, firstDate, and list of Data Elements
function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var app = UiApp.createApplication().setTitle("Setup Fitbit Download");
    app.setStyleAttribute("padding", "10px");

    var consumerKeyLabel = app.createLabel("Fitbit OAuth Consumer Key:*");
    var consumerKey = app.createTextBox();
    consumerKey.setName("consumerKey");
    consumerKey.setWidth("100%");
    consumerKey.setText(getConsumerKey());
    var consumerSecretLabel = app.createLabel("Fitbit OAuth Consumer Secret:*");
    var consumerSecret = app.createTextBox();
    consumerSecret.setName("consumerSecret");
    consumerSecret.setWidth("100%");
    consumerSecret.setText(getConsumerSecret());
    var firstDate = app.createTextBox().setId("firstDate").setName("firstDate");
    firstDate.setName("firstDate");
    firstDate.setWidth("100%");
    firstDate.setText(getFirstDate());

    // add listbox to select data elements
    var loggables = app.createListBox(true).setId("loggables").setName(
      "loggables");
    loggables.setVisibleItemCount(4);
    // add all possible elements (in array LOGGABLES)
    var logIndex = 0;
    for (var resource in LOGGABLES) {
        loggables.addItem(LOGGABLES[resource]);
        // check if this resource is in the getLoggables list
        if (getLoggables().indexOf(LOGGABLES[resource]) > -1) {
          // if so, pre-select it
          loggables.setItemSelected(logIndex, true);
        }
        logIndex++;
    }
    // create the save handler and button
    var saveHandler = app.createServerClickHandler("saveSetup");
    var saveButton = app.createButton("Save Setup", saveHandler);

    // put the controls in a grid
    var listPanel = app.createGrid(6, 3);
    listPanel.setWidget(1, 0, consumerKeyLabel);
    listPanel.setWidget(1, 1, consumerKey);
    listPanel.setWidget(2, 0, consumerSecretLabel);
    listPanel.setWidget(2, 1, consumerSecret);
    listPanel.setWidget(3, 0, app.createLabel(" * (obtain these at dev.fitbit.com)"));
    listPanel.setWidget(4, 0, app.createLabel("Start Date for download (yyyy-mm-dd)"));
    listPanel.setWidget(4, 1, firstDate);
    listPanel.setWidget(5, 0, app.createLabel("Data Elements to download:"));
    listPanel.setWidget(5, 1, loggables);
    
    // Ensure that all controls in the grid are handled
    saveHandler.addCallbackElement(listPanel);
    // Build a FlowPanel, adding the grid and the save button
    var dialogPanel = app.createFlowPanel();
    dialogPanel.add(listPanel);
    dialogPanel.add(saveButton);
    app.add(dialogPanel);
    doc.show(app);
}

// function sync() is called to download all desired data from Fitbit API to the spreadsheet                 
function sync() {
    // if the user has never performed setup, do it now
    if (!isConfigured()) {
        setup();
        return;
    }

    var user = authorize();
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    doc.setFrozenRows(1);
    var options = {
        "oAuthServiceName": "fitbit",
        "oAuthUseToken": "always",
        "method": "GET"
    };
    // prepare and format today's date, and a list of desired data elements
    var dateString = formatToday();
    var activities = getLoggables();
    // for each data element, fetch a list beginning from the firstDate, ending with today
    for (var activity in activities) {
        var currentActivity = activities[activity];
        try {
            var result = UrlFetchApp.fetch("https://api.fitbit.com/1/user/-/"
          + currentActivity + "/date/" + getFirstDate() + "/"
          + dateString + ".json", options);
        } catch (exception) {
            Logger.log(exception);
            Browser.msgBox("Error downloading " + currentActivity);
        }
        var o = Utilities.jsonParse(result.getContentText());

        // set title
        var titleCell = doc.getRange("a1");
        titleCell.setValue("date");
        var cell = doc.getRange('a2');

        // fill the spreadsheet with the data
        var index = 0;
        for (var i in o) {
            // set title for this column
            var title = i.substring(i.lastIndexOf('-') + 1);
            titleCell.offset(0, 1 + activity * 1.0).setValue(title);

            var row = o[i];
            for (var j in row) {
                var val = row[j];
                cell.offset(index, 0).setValue(val["dateTime"]);
                // set the date index
                cell.offset(index, 1 + activity * 1.0).setValue(val["value"]);
                // set the value index index
                index++;
            }
        }
    }
}

function isConfigured() {
    return getConsumerKey() != "" && getConsumerSecret() != "";
}

function setConsumerKey(key) {
    ScriptProperties.setProperty(CONSUMER_KEY_PROPERTY_NAME, key);
}

function getConsumerKey() {
    var key = ScriptProperties.getProperty(CONSUMER_KEY_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
}

function setLoggables(loggable) {
    ScriptProperties.setProperty("loggables", loggable);
}

function getLoggables() {
    var loggable = ScriptProperties.getProperty("loggables");
    if (loggable == null) {
        loggable = LOGGABLES;
    } else {
        loggable = loggable.split(',');
    }
    return loggable;
}

function setFirstDate(firstDate) {
    ScriptProperties.setProperty("firstDate", firstDate);
}

function getFirstDate() {
    var firstDate = ScriptProperties.getProperty("firstDate");
    if (firstDate == null) {
        firstDate = "2012-01-01";
    }
    return firstDate;
}

function formatToday() {
    var todayDate = new Date;
    return todayDate.getFullYear()
    + '-'
    + ("00" + (todayDate.getMonth() + 1)).slice(-2)
    + '-'
    + ("00" + todayDate.getDate()).slice(-2);
}

function setConsumerSecret(secret) {
    ScriptProperties.setProperty(CONSUMER_SECRET_PROPERTY_NAME, secret);
}

function getConsumerSecret() {
    var secret = ScriptProperties.getProperty(CONSUMER_SECRET_PROPERTY_NAME);
    if (secret == null) {
        secret = "";
    }
    return secret;
}

// function saveSetup saves the setup params from the UI
function saveSetup(e) {
    setConsumerKey(e.parameter.consumerKey);
    setConsumerSecret(e.parameter.consumerSecret);
    setLoggables(e.parameter.loggables);
    setFirstDate(e.parameter.firstDate);
    var app = UiApp.getActiveApplication();
    app.close();
    return app;
}

// function onOpen is called when the spreadsheet is opened; adds the Fitbit menu
function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{
        name: "Sync",
        functionName: "sync"
    }, {
        name: "Setup",
        functionName: "setup"
    }, {
        name: "Authorize",
        functionName: "authorize"
    }];
    ss.addMenu("Fitbit", menuEntries);
}

// function onInstall is called when the script is installed (obsolete?)
function onInstall() {
    onOpen();
}
