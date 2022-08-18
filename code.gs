function start() {
  var calendar = CalendarApp.getAllCalendars
  var spreadsheet = SpreadsheetApp.getActive
  console.log(CalendarApp.getEventsForDay)
  console.log(CalendarApp.getEvents)
}

/**
 * Lists the calendars shown in the user's calendar list.
 */
function listCalendars() {
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
        Logger.log('%s (ID: %s)', calendar.summary, calendar.id);
      }
    } else {
      Logger.log('No calendars found.');
    }
    pageToken = calendars.nextPageToken;
  } while (pageToken);
}
function listNext10Events() {
  var calendarId = 'primary';
  var now = new Date();
  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  if (events.items && events.items.length > 0) {
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      if (event.start.date) {
        // All-day event.
        var start = new Date(event.start.date);
        Logger.log('%s (%s)', event.summary, start.toLocaleDateString());
      } else {
        var start = new Date(event.start.dateTime);
        Logger.log('%s (%s)', event.summary, start.toLocaleString());
      }
    }
    // Utilities.sleep(30 * 1000);
  } else {
    Logger.log('No events found.');
  }
}
/**
 * Retrieve and log events from the given calendar that have been modified
 * since the last sync. If the sync token is missing or invalid, log all
 * events from up to a month ago (a full sync).
 *
 * @param {string} calendarId The ID of the calender to retrieve events from.
 * @param {boolean} fullSync If true, throw out any existing sync token and
 *        perform a full sync; if false, use the existing sync token if possible.
 */
function logSyncedEvents(calendarId, fullSync) {
  var properties = PropertiesService.getUserProperties();
  var options = {
    maxResults: 100
  };
  var syncToken = properties.getProperty('syncToken');
  if (syncToken && !fullSync) {
    options.syncToken = syncToken;
  } else {
    // Sync events up to thirty days in the past.
    options.timeMin = getRelativeDate(-30, 0).toISOString();
  }

  // Retrieve events one page at a time.
  var events;
  var pageToken;
  do {
    try {
      options.pageToken = pageToken;
      events = Calendar.Events.list(calendarId, options);
    } catch (e) {
      // Check to see if the sync token was invalidated by the server;
      // if so, perform a full sync instead.
      if (e.message === 'Sync token is no longer valid, a full sync is required.') {
        properties.deleteProperty('syncToken');
        logSyncedEvents(calendarId, true);
        return;
      } else {
        throw new Error(e.message);
      }
    }

    if (events.items && events.items.length > 0) {
      for (var i = 0; i < events.items.length; i++) {
        var event = events.items[i];
        if (event.status === 'cancelled') {
          console.log('Event id %s was cancelled.', event.id);
        } else if (event.start.date) {
          // All-day event.
          var start = new Date(event.start.date);
          console.log('%s (%s)', event.summary, start.toLocaleDateString());
        } else {
          // Events that don't last all day; they have defined start times.
          var start = new Date(event.start.dateTime);
          console.log('%s (%s)', event.summary, start.toLocaleString());
        }
      }
    } else {
      console.log('No events found.');
    }

    pageToken = events.nextPageToken;
  } while (pageToken);

  properties.setProperty('syncToken', events.nextSyncToken);
}
