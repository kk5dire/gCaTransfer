function testn1() {
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

function onOpen() {
   var ui = SpreadsheetApp.getUi();
  ui.createMenu('iCal Tools')
      .addItem('Create Events', 'crevnt')
      .addToUi();
 SpreadsheetApp.getUi().alert("ℹ️ Information: Press 'iCal Tools'/'Create Events' to update the sheet, sheet is updated automatically hourly. If you are a viewer this does not apply")
}
function crevnt() {
  SpreadsheetApp.getUi().alert('🛠️ Working: Sheet is updating from calendar manually, If you are an editor, check the logs for errors');
  listPrimary()
}
function listPrimary() {
  deleteall()
  spreadsheet = SpreadsheetApp.getActive();
  var scriptPrp = PropertiesService.getScriptProperties()
  var calid = scriptPrp.getProperty('newcalid')
  var calendarId = 'primary'
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
        if (event.summary == null) Logger.log(event.summary + event.source + ": returned null")
        var valwow = event.summary + ' ' + event.start
        spreadsheet.getCurrentCell().setValue(valwow);
        spreadsheet.getCurrentCell().offset(1, 0).activate();
      } else {
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
    Logger.log('No events found.');
  }
}
function listPrimary2() {
  spreadsheet = SpreadsheetApp.getActive();
  var scriptPrp = PropertiesService.getScriptProperties()
  var calid = scriptPrp.getProperty('newcalid')
  var calendarId = 'aj6ib22496f56gr2hcqses2fphtp3paa@import.calendar.google.com'
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
  var calid = scriptPrp.getProperty('newcalid')
  var calendarId = 'c_classroom368aba6f@group.calendar.google.com'
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
  var calendarId = 'c_classroomfc919c1c@group.calendar.google.com'
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
  var calendarId = 'c_classroom2fe7040c@group.calendar.google.com'
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
  var calendarId = 'c_classroom7634b5ac@group.calendar.google.com'
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
  var calendarId = 'c_classroomc0226b5d@group.calendar.google.com'
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
  var calendarId = 'c_ocbf2apginsaqv017stvdc1pbc@group.calendar.google.com'
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
function listPrimary9() {
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
    } else {
      // SpreadsheetApp.getUi().alert("iCal Error: No Calendars found");
      Logger.log('No calendars found.');
    }
    pageToken = calendars.nextPageToken;
  } while (pageToken);
}

function deleteall() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  var currentCell = spreadsheet.getCurrentCell();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};
function sheetchange() {
 SpreadsheetApp.getUi().alert("⚠️Warning: Sheet was modified, This may not reflect accurate data")
}