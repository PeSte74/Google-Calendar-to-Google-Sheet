// References:
// https://github.com/Davepar/gcalendarsync
// https://developers.google.com/apps-script/reference/calendar/

// Set this value to match your calendar!!!
// Calendar ID can be found in the "Calendar Address" section of the Calendar Settings.
var calendarId = '';
var companyDomains = ['geotab.com','geotabinc.com','fleetcarma.com','intendia.com'];
var dataSheet = 'Data'

// Set the beginning and end dates that should be synced. beginDate can be set to Date() to use
// today. The numbers are year, month, date, where month is 0 for Jan through 11 for Dec.
//var beginDate = new Date(2020, 2, 7);  // Default to Jan 1, 1970
var endDate = new Date();  // Default to today
var beginDate = new Date();
beginDate.setDate(beginDate.getDate() - 30);

// Date format to use in the spreadsheet.
var dateFormat = 'dd/MM/yyyy HH:mm';

var titleRowMap = {
  'title': 'Title',
  'description': 'Description',
  'location': 'Location',
  'starttime': 'Start Time',
  'endtime': 'End Time',
  'duration': 'Duration',
  'recurring': 'Recurring',
  'type': 'Type',
  'status': 'Status',
  'creators': 'Creator',
  'guests' : 'Guests',
  'guestsDetails': 'Guests Details',
  'color': 'Color',
  'id': 'Id'
};

var titleRowKeys = ['title', 'description', 'location', 'starttime', 'endtime', 'duration', 'recurring', 'type','status','creators','guestsDetails', 'guests', 'color', 'id'];
var requiredFields = ['id', 'title', 'starttime', 'endtime'];

// Creates a mapping array between spreadsheet column and event field name
function createIdxMap(row) {
  var idxMap = [];
  for (var idx = 0; idx < row.length; idx++) {
    var fieldFromHdr = row[idx];
    for (var titleKey in titleRowMap) {
      if (titleRowMap[titleKey] == fieldFromHdr) {
        idxMap.push(titleKey);
        break;
      }
    }
    if (idxMap.length <= idx) {
      // Header field not in map, so add null
      idxMap.push(null);
    }
  }
  return idxMap;
}

// Set up formats and hide ID column for empty spreadsheet
function setUpSheet(sheet, fieldKeys) {
  sheet.getRange(1, fieldKeys.indexOf('starttime') + 1, 999).setNumberFormat(dateFormat);
  sheet.getRange(1, fieldKeys.indexOf('endtime') + 1, 999).setNumberFormat(dateFormat);
  sheet.getRange(1, fieldKeys.indexOf('duration') + 1, 999).setNumberFormat('0.0');
  sheet.hideColumns(fieldKeys.indexOf('id') + 1);
  sheet.hideColumns(fieldKeys.indexOf('description') + 1);
}

// Display error alert
function errorAlert(msg, evt, ridx) {
  var ui = SpreadsheetApp.getUi();
  if (evt) {
    ui.alert('Skipping row: ' + msg + ' in event "' + evt.title + '", row ' + (ridx + 1));
  } else {
    ui.alert(msg);
  }
}

// Determine whether required fields are missing
function areRequiredFieldsMissing(idxMap) {
  return requiredFields.some(function(val) {
    return idxMap.indexOf(val) < 0;
  });
}

//Determine if the calendar event is One to One, Internal or External 
function getMeetingType(creators, guests) {
  var meetingType = 'External';

  if(creators.toString().toLowerCase().indexOf(companyDomains[0])>-1 ) {
    if(guests) {
      let regexp = new RegExp(companyDomains.join("|"), "gi");
      if(guests.match(regexp).length == guests.split(',').length  ) {
        meetingType = 'Internal';
        if(guests.split(',').length == 1) {meetingType = 'One to One';}
      }     
    } else { meetingType = 'Own Time';}  
  } 
  return meetingType; 
}

function calculateDurationHours(startTime, endTime) {
  console.info(endTime)
  return duration = (endTime- startTime)/1000/60/60;
  
}

// Converts a calendar event to a psuedo-sheet event.
function convertCalEvent(calEvent) {
  convertedEvent = {
    'id': calEvent.getId(),
    'title': calEvent.getTitle(),
    'description': calEvent.getDescription(),
    'location': calEvent.getLocation(),
    'recurring': calEvent.isRecurringEvent(),
    'creators': calEvent.getOriginalCalendarId(), //getCreators(),
    'status': calEvent.getMyStatus(),
    //'guests': calEvent.getGuestList().map(function(x) {return x.getEmail();}).filter(function(x) {return x.indexOf('resource.calendar.google.com')===-1;}).join(','),
    'color': calEvent.getColor()
  };
  
  let guestsDetails = calEvent.getGuestList().map(function(x) {return x.getEmail();}).filter(function(x) {return x.indexOf('resource.calendar.google.com')===-1;});
  convertedEvent.guests = guestsDetails.length;
  convertedEvent.guestsDetails = guestsDetails.join(',');
  convertedEvent.type = getMeetingType(convertedEvent.creators, convertedEvent.guestsDetails);
  
  if(convertedEvent.creators.startsWith("geotabinc.com_")) {convertedEvent.creators = calEvent.getCreators()}
  
  if (calEvent.isAllDayEvent()) {
    convertedEvent.starttime = calEvent.getAllDayStartDate();
    let endtime = calEvent.getAllDayEndDate();
    if (endtime - convertedEvent.starttime === 24 * 3600 * 1000) {
      convertedEvent.endtime = '';
    } else {
      convertedEvent.endtime = endtime;
      if (endtime.getHours() === 0 && endtime.getMinutes() == 0) {
        convertedEvent.endtime.setSeconds(endtime.getSeconds() - 1);
      }
    }
  } else {
    convertedEvent.starttime = calEvent.getStartTime();
    convertedEvent.endtime = calEvent.getEndTime();
    if(convertedEvent.status != 'NO') {
      convertedEvent.duration = calculateDurationHours(convertedEvent.starttime.getTime(), convertedEvent.endtime.getTime());
      }
    }
  return convertedEvent;
}

// Converts calendar event into spreadsheet data row
function calEventToSheet(calEvent, idxMap, dataRow) {
  convertedEvent = convertCalEvent(calEvent);

  for (var idx = 0; idx < idxMap.length; idx++) {
    if (idxMap[idx] !== null) {
      dataRow[idx] = convertedEvent[idxMap[idx]];
    }
  }
}

function syncFromCalendar() {
  console.info('Starting sync from calendar');
  console.info(beginDate);
  console.info(endDate);         
  // Get calendar and events
  var calendar = CalendarApp.getCalendarById(calendarId);
  var calEvents = calendar.getEvents(beginDate, endDate);
  
  // Get spreadsheet and data
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(dataSheet);
  var range = sheet.getDataRange();
  var data = range.getValues();
  var eventFound = new Array(data.length);
  
  // Check if spreadsheet is empty and add a title row
  var titleRow = [];
  for (var idx = 0; idx < titleRowKeys.length; idx++) {
    titleRow.push(titleRowMap[titleRowKeys[idx]]);
  }
  if (data.length < 1) {
    data.push(titleRow);
    range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    setUpSheet(sheet, titleRowKeys);
  }

  if (data.length == 1 && data[0].length == 1 && data[0][0] === '') {
    data[0] = titleRow;
    range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    setUpSheet(sheet, titleRowKeys);
  }
  
  // Map spreadsheet headers to indices
  var idxMap = createIdxMap(data[0]);
  var idIdx = idxMap.indexOf('id');

  // Verify header has all required fields
  if (areRequiredFieldsMissing(idxMap)) {
    var reqFieldNames = requiredFields.map(function(x) {return titleRowMap[x];}).join(', ');
    errorAlert('Spreadsheet must have ' + reqFieldNames + ' columns');
    return;
  }
  
  
  // Array of IDs in the spreadsheet
  var sheetEventIds = data.slice(1).map(function(row) {return row[idIdx];});
  
  // Loop through calendar events
  for (var cidx = 0; cidx < calEvents.length; cidx++) {
    var calEvent = calEvents[cidx];
    var calEventId = calEvent.getId();

    var ridx = sheetEventIds.indexOf(calEventId) + 1;
    if (ridx < 1) {
      // Event not found, create it
      ridx = data.length;
      var newRow = [];
      var rowSize = idxMap.length;
      while (rowSize--) newRow.push('');
      data.push(newRow);
    } else {
      eventFound[ridx] = true;
    }
    // Update event in spreadsheet data
    calEventToSheet(calEvent, idxMap, data[ridx]);
  }
  
  // Remove any data rows not found in the calendar
  var rowsDeleted = 0;
  for (var idx = eventFound.length - 1; idx > 0; idx--) {
    //event doesn't exists and has an event id
    if (!eventFound[idx] && sheetEventIds[idx - 1]) {
      data.splice(idx, 1);
      rowsDeleted++;
    }
  }

  // Save spreadsheet changes
  range = sheet.getRange(1, 1, data.length, data[0].length);
  range.setValues(data);
  if (rowsDeleted > 0) {
    sheet.deleteRows(data.length + 1, rowsDeleted);
  }

  
  
}
