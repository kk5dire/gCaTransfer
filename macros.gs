function testn1() { // Locate Data from GApps to be used
  var calendar = CalendarApp.getAllCalendars
  var spreadsheet = SpreadsheetApp.getActive();
  console.log(CalendarApp.getEventsForDay)
  console.log(CalendarApp.getEvents)
  spreadsheet.getCurrentCell().setValue(CalendarApp.getEventsForDay);
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.getCurrentCell().setValue(CalendarApp.getEvents);
  spreadsheet.getCurrentCell().offset(1, 0).activate();
};

function valset() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().setValue('val1');
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.getCurrentCell().setValue('val2');
  spreadsheet.getCurrentCell().offset(1, 0).activate();
};

function onOpen() {    // Read this chunk on open, is used to create UI objects for easy access
   var ui = SpreadsheetApp.getUi();
  ui.createMenu('iCal Tools')
      .addItem('Create Events', 'crevnt')
      .addToUi();
  // Intro alert
 SpreadsheetApp.getUi().alert("ℹ️ Information: Press 'iCal Tools'/'Create Events' to update the sheet, sheet is updated automatically hourly. If you are a viewer this does not apply")
}
function crevnt() { // Alert that the calendar is updating
  SpreadsheetApp.getUi().alert('🛠️ Working: Sheet is updating from calendar manually, If you are an editor, check the logs for errors');
  listPrimary()
}
function listPrimary() {
  
  // Delete all cached data
  deleteall() // Lets check all events from the primary calendar 
  // I would note that this tool was hand crafted for a specific reason and if you wish to use this for yourself you will indeed need to modify some seconds of it
  // for example, Calendar IDs, and number of functions called to merge calendars into primary
  //As this is the first fun I will break it down for those who want to understand it 
  
  
  spreadsheet = SpreadsheetApp.getActive();
  // Get the selected Spreadsheet  (aka primary)
  var scriptPrp = PropertiesService.getScriptProperties()
  // Get the properties of said sheet
  var calid = scriptPrp.getProperty('newcalid')
  
  var calendarId = 'primary'
  Logger.log(calendarId + ": Calendar ID")
  //Bind a calendar and calendar ID to it
  var now = new Date();
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  //Set the params to fetch payload
  if (events.items && events.items.length > 0) { // lets check this list is not empty
    spreadsheet = SpreadsheetApp.getActive()
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      if (event.start.date) {
        // All-day event.
        var start = new Date(event.start.date);
        Logger.log('%s (%s)', event.summary, start.toLocaleDateString());
        Logger.log(listCalendars()) // Dump the info and make sure the results were not a false positive
if (event.summary == null) Logger.log(event.summary + event.source + ": returned null")
        if (event.summary == null) Logger.log(event.summary + event.source + ": returned null")
        var valwow = event.summary + ' ' + event.start
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      } else { // Add the data as an offset of the last into the sheet
        var start = new Date(event.start.dateTime);
        Logger.log('%s (%s)', event.summary, start.toLocaleString());
        var valwow = ``
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      }
    }
    listPrimary2()
    // Utilities.sleep(30 * 1000);
  } else {
    // SpreadsheetApp.getUi().alert("Primary iCal: No events found");
    Logger.log('No events found.'); // No events found, lets back out of the calendar and return an error in gapps logs
  }
}
function listPrimary2() { // check events from the secondary calendar and update the payload
  spreadsheet = SpreadsheetApp.getActive();
  var scriptPrp = PropertiesService.getScriptProperties()
  var calid = scriptPrp.getProperty('newcalid')
  var calendarId = 'aj6a@import.calendar.google.com'
  Logger.log(calendarId + ": Calendar ID")
  var now = new Date();
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  if (events.items && events.items.length > 0) {
    spreadsheet = SpreadsheetApp.getActive()
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      if (event.start.date) {
        // All-day event.
        var start = new Date(event.start.date);
        Logger.log('%s (%s)', event.summary, start.toLocaleDateString());
        Logger.log(listCalendars())
if (event.summary == null) Logger.log(event.summary + event.source + ": returned null")
        var valwow = event.summary + ' ' + event.start
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      } else {
        var start = new Date(event.start.dateTime);
        Logger.log('%s (%s)', event.summary, start.toLocaleString());
        var valwow = ''
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      }
    }
    listPrimary3()
    // Utilities.sleep(30 * 1000);
  } else {
    // SpreadsheetApp.getUi().alert("1 iCal: No events found");
    Logger.log('No events found.');
  }
}
function listPrimary3() {
  spreadsheet = SpreadsheetApp.getActive();
  var scriptPrp = PropertiesService.getScriptProperties()
  var calid = scriptPrp.getProperty('newcalid')   // Check active classroom calendars and update the payload
  var calendarId = 'c_classroom@group.calendar.google.com'
  Logger.log(calendarId + ": Calendar ID")
  var now = new Date();
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  if (events.items && events.items.length > 0) {
    spreadsheet = SpreadsheetApp.getActive()
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      if (event.start.date) {
        // All-day event.
        var start = new Date(event.start.date);
        Logger.log('%s (%s)', event.summary, start.toLocaleDateString());
        Logger.log(listCalendars())
if (event.summary == null) Logger.log(event.summary + event.source + ": returned null")
        var valwow = event.summary + ' ' + event.start
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      } else {
        var start = new Date(event.start.dateTime);
        Logger.log('%s (%s)', event.summary, start.toLocaleString());
        var valwow = ''
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      }
    }
    listPrimary4()
    // Utilities.sleep(30 * 1000);
  } else {
    // SpreadsheetApp.getUi().alert("2 iCal: No events found");
    Logger.log('No events found.');
  }
}
function listPrimary4() {
  spreadsheet = SpreadsheetApp.getActive();
  var scriptPrp = PropertiesService.getScriptProperties()
  var calid = scriptPrp.getProperty('newcalid')
  var calendarId = 'c_@group.calendar.google.com'
  Logger.log(calendarId + ": Calendar ID")
  var now = new Date();
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  if (events.items && events.items.length > 0) {
    spreadsheet = SpreadsheetApp.getActive()
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      if (event.start.date) {
        // All-day event.
        var start = new Date(event.start.date);
        Logger.log('%s (%s)', event.summary, start.toLocaleDateString());
        Logger.log(listCalendars())
if (event.summary == null) Logger.log(event.summary + event.source + ": returned null")
        var valwow = event.summary + ' ' + event.start
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      } else {
        var start = new Date(event.start.dateTime);
        Logger.log('%s (%s)', event.summary, start.toLocaleString());
        var valwow = ''
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      }
    }
    listPrimary5()
    // Utilities.sleep(30 * 1000);
  } else {
    // SpreadsheetApp.getUi().alert("3 iCal: No events found");
    Logger.log('No events found.');
  }
}
function listPrimary5() {
  spreadsheet = SpreadsheetApp.getActive();
  var scriptPrp = PropertiesService.getScriptProperties()
  var calid = scriptPrp.getProperty('newcalid')
  var calendarId = 'c@group.calendar.google.com'
  Logger.log(calendarId + ": Calendar ID")
  var now = new Date();
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  if (events.items && events.items.length > 0) {
    spreadsheet = SpreadsheetApp.getActive()
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      if (event.start.date) {
        // All-day event.
        var start = new Date(event.start.date);
        Logger.log('%s (%s)', event.summary, start.toLocaleDateString());
        Logger.log(listCalendars())
if (event.summary == null) Logger.log(event.summary + event.source + ": returned null")
        var valwow = event.summary + ' ' + event.start
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      } else {
        var start = new Date(event.start.dateTime);
        Logger.log('%s (%s)', event.summary, start.toLocaleString());
        var valwow = ''
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      }
    }
    listPrimary6()
    // Utilities.sleep(30 * 1000);
  } else {
    // SpreadsheetApp.getUi().alert("4 iCal: No events found");
    Logger.log('No events found.');
  }
}
function listPrimary6() {
  spreadsheet = SpreadsheetApp.getActive();
  var scriptPrp = PropertiesService.getScriptProperties()
  var calid = scriptPrp.getProperty('newcalid')
  var calendarId = 'c_c@calendar.google.com'
  Logger.log(calendarId + ": Calendar ID")
  var now = new Date();
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  if (events.items && events.items.length > 0) {
    spreadsheet = SpreadsheetApp.getActive()
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      if (event.start.date) {
        // All-day event.
        var start = new Date(event.start.date);
        Logger.log('%s (%s)', event.summary, start.toLocaleDateString());
        Logger.log(listCalendars())
if (event.summary == null) Logger.log(event.summary + event.source + ": returned null")
        var valwow = event.summary + ' ' + event.start
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      } else {
        var start = new Date(event.start.dateTime);
        Logger.log('%s (%s)', event.summary, start.toLocaleString());
        var valwow = ''
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      }
    }
    listPrimary7()
    // Utilities.sleep(30 * 1000);
  } else {
    // SpreadsheetApp.getUi().alert("5 iCal: No events found");
    Logger.log('No events found.');
  }
}
function listPrimary7() {
  spreadsheet = SpreadsheetApp.getActive();
  var scriptPrp = PropertiesService.getScriptProperties()
  var calid = scriptPrp.getProperty('newcalid')
  var calendarId = 'c_@group.google.com'
  Logger.log(calendarId + ": Calendar ID")
  var now = new Date();
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  if (events.items && events.items.length > 0) {
    spreadsheet = SpreadsheetApp.getActive()
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      if (event.start.date) {
        // All-day event.
        var start = new Date(event.start.date);
        Logger.log('%s (%s)', event.summary, start.toLocaleDateString());
        Logger.log(listCalendars())
if (event.summary == null) Logger.log(event.summary + event.source + ": returned null")
        var valwow = event.summary + ' ' + event.start
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      } else {
        var start = new Date(event.start.dateTime);
        Logger.log('%s (%s)', event.summary, start.toLocaleString());
        var valwow = ''
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      }
    }
    listPrimary8()
    // Utilities.sleep(30 * 1000);
  } else {
    // SpreadsheetApp.getUi().alert("6 iCal: No events found");
    Logger.log('No events found.');
  }
}
function listPrimary8() {
  spreadsheet = SpreadsheetApp.getActive();
  var scriptPrp = PropertiesService.getScriptProperties()
  var calid = scriptPrp.getProperty('newcalid')
  var calendarId = 'c_ocb@group.calendar.google.com'
  Logger.log(calendarId + ": Calendar ID")
  var now = new Date();
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  if (events.items && events.items.length > 0) {
    spreadsheet = SpreadsheetApp.getActive()
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      if (event.start.date) {
        // All-day event.
        var start = new Date(event.start.date);
        Logger.log('%s (%s)', event.summary, start.toLocaleDateString());
        Logger.log(listCalendars())
if (event.summary == null) Logger.log(event.summary + event.source + ": returned null")
        var valwow = event.summary + ' ' + event.start
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      } else {
        var start = new Date(event.start.dateTime);
        Logger.log('%s (%s)', event.summary, start.toLocaleString());
        var valwow = ''
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      }
    }
    listPrimary9()
    // Utilities.sleep(30 * 1000);
  } else {
    // SpreadsheetApp.getUi().alert("7 iCal: No events found");
    Logger.log('No events found.');
  }
}
function listPrimary9() {  // A simple calendar function to get holidays in the united states coming up and add them to the list
  spreadsheet = SpreadsheetApp.getActive();
  var scriptPrp = PropertiesService.getScriptProperties()
  var calid = scriptPrp.getProperty('newcalid')
  var calendarId = 'en.usa#holiday@group.v.calendar.google.com'
  Logger.log(calendarId + ": Calendar ID")
  var now = new Date();
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  if (events.items && events.items.length > 0) {
    spreadsheet = SpreadsheetApp.getActive()
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      if (event.start.date) {
        // All-day event.
        var start = new Date(event.start.date);
        Logger.log('%s (%s)', event.summary, start.toLocaleDateString());
        Logger.log(listCalendars())
if (event.summary == null) Logger.log(event.summary + event.source + ": returned null")
        var valwow = event.summary + ' ' + event.start
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      } else {
        var start = new Date(event.start.dateTime);
        Logger.log('%s (%s)', event.summary, start.toLocaleString());
        var valwow = ''
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      }
    }
    // Utilities.sleep(30 * 1000);
  } else {
    // SpreadsheetApp.getUi().alert("8 iCal: No events found");
    Logger.log('No events found.');
  }
}
function lcal() {
  var calendars;
  var pageToken;
  do {
    calendars = Calendar.CalendarList.list({
      maxResults: 100,
      pageToken: pageToken
    });
    if (calendars.items && calendars.items.length > 0) {
      for (var i = 0; i < calendars.items.length; i++) {
        var calendar = calendars.items[i];
        Logger.log(calendar.id);
        var scriptPrp = PropertiesService.getScriptProperties()
        scriptPrp.setProperty('newcalid', calendar.id)
      }
    } else { // No Calendars found, alert the user
      // SpreadsheetApp.getUi().alert("iCal Error: No Calendars found");
      Logger.log('No calendars found.');
    }
    pageToken = calendars.nextPageToken;
  } while (pageToken);
}

function deleteall() { // Clear calendar cache to prevent data conflicting
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  var currentCell = spreadsheet.getCurrentCell();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};
function sheetchange() { // Calendar was updated alert the user
 SpreadsheetApp.getUi().alert("⚠️Warning: Sheet was modified, This may not reflect accurate data")
}
