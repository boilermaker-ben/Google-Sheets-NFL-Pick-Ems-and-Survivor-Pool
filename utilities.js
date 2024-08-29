// UTILITIES
//------------------------------------------------------------------------
// RESET Function to reset and create menu for runFirst
function resetSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Return to spreadsheet for prompts');
  let prompt = ui.alert('Reset spreadsheet and delete all data?', ui.ButtonSet.YES_NO);
  if (prompt == 'YES') {
    
    let promptTwo = ui.alert('Are you sure? This would be very difficult to recover from.',ui.ButtonSet.YES_NO);
    if (promptTwo == 'YES') {
      let ranges = ss.getNamedRanges();
      for (let a = 0; a < ranges.length; a++){
        ranges[a].remove();
      }
      let sheets = ss.getSheets();
      let baseSheet = ss.insertSheet();
      for (let a = 0; a < sheets.length; a++){
        ss.deleteSheet(sheets[a]);
      }
      let protections = ss.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (let a = 0; a < protections.length; a++){
        protections[a].remove();
      }
      protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (let a = 0; a < protections.length; a++){
        protections[a].remove();
      }      
      baseSheet.setName('Sheet1');

      // Deletes initialization, time zone, and any other response-associated properites
      let properties = PropertiesService.getScriptProperties();
      properties.deleteAllProperties();

      deleteTriggers();

      initializeMenu();

    } else {
      ss.toast('Canceled reset');
    }
  } else {
    ss.toast('Canceled reset');
  }
  
}

// FETCH SPREADSHEET - Checks that the 'ss' variable passed into a script is not null, undefined, or a non-spreadsheet
function fetchSpreadsheet(ss) {
  try {
    if (ss && typeof ss.getSheets === 'function' && typeof ss.getId === 'function') {
      return ss;
    } else {
      throw new Error('Invalid Spreadsheet object');
    }
  } catch (err) {
    if (ss !== null && ss !== undefined) {
      Logger.log('ALERT: The function \'' + (new Error()).stack.split('\n')[2].trim().split(' ')[1] + '\' passed ' + typeof ss + ' \'' + ss + '\' to the \'fetchSpreadsheet\' function.');
      Logger.log(err.stack);
    }
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  return ss;
}

// FETCH UI - Checks that the 'ui' variable passed into a script is not null, undefined, or a non-UI
function fetchUi(ui) {
  try{
    if (typeof ui.showModalDialog !== 'function') {
      throw new Error('Non-UI passed');
    }
  }
  catch (err) {
    if (ui !== null && ui !== undefined) {
      Logger.log('ALERT: The function \'' + (new Error()).stack.split('\n')[2].trim().split(' ')[1] + '\' passed ' + typeof ui + ' \'' + ui + '\' to the \'fetchUi\' function.')
    }
    ui = SpreadsheetApp.getUi();
  }
  return ui;
}

// SERVICE Function to remove all ONOPEN triggers on project
function deleteOnOpenTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let a = 0; a < triggers.length; a++) {
    if (triggers[a].getEventType() === ScriptApp.EventType.ON_OPEN) {
      ScriptApp.deleteTrigger(triggers[a]);
    }
  }
}

// SERVICE Function to remove all triggers on project
function deleteTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let a = 0; a < triggers.length; a++) {
    ScriptApp.deleteTrigger(triggers[a]);
  }
}

// ADJUST ROWS - Cleans up rows of a sheet by providing the total rows that currently exist with data
function adjustRows(sheet,rows,verbose){
  let maxRows = sheet.getMaxRows(); 
  if (rows == undefined || rows == null) {
    rows = sheet.getLastRow();
  }
  if (rows > 0 && rows > maxRows) {
    sheet.insertRowsAfter(maxRows,(rows-maxRows));
    if(verbose) return Logger.log('Added ' + (rows-maxRows) + ' rows');
  } else if (rows < maxRows && rows != 0){
    sheet.deleteRows((rows+1), (maxRows-rows));
    if(verbose) return Logger.log('Removed ' + (maxRows - rows) + ' rows');
  } else {
    if(verbose) return Logger.log('Rows not adjusted');
  }
}

// ADJUST COLUMNS - Cleans up columns of a sheet by providing the total columns that currently exist with data
function adjustColumns(sheet,columns,verbose){
  let maxColumns = sheet.getMaxColumns(); 
  if (columns == undefined || columns == null) {
    columns = sheet.getLastColumn();
  }
  if (columns > 0 && columns > maxColumns) {
    sheet.insertColumnsAfter(maxColumns,(columns-maxColumns));
    if(verbose) return Logger.log('Added ' + (columns-maxColumns) + ' columns');
  }  else if (columns < maxColumns && columns != 0){
    sheet.deleteColumns((columns+1), (maxColumns-columns));
    if(verbose) return Logger.log('Removed ' + (maxColumns - columns) + ' column(s)');
  } else {
    if(verbose) return Logger.log('Columns not adjusted');
  }
}

// NEXT WEEK - Detects what is the next highest integer weekly sheet for prompting creation of new forms (or sheets)
function nextWeek() {
  let weekObj = fetchWeekCompare();
  if (weekObj.api <= 0 && weekObj.prop == null) {
    Logger.log('No responses and currently preseason active');
    return 1;
  } else if (weekObj.prop > 0) {
    Logger.log('Responses recorded for week ' + weekObj.prop + ', advancing to week ' + (weekObj.prop + 1));
    return (weekObj.prop + 1);
  } else {
    return weekObj.api;
  }
}

// MAX WEEK - Detects what is the likely maximum week based on API, property, and CONFIG sheet inputs
function maxWeek() {
  let weekObj = fetchWeekCompare();
  let max = 1;
  Object.keys(weekObj).forEach(value => {
    max = max < weekObj[value] ? weekObj[value] : max;
  });
  return max;
}

// FETCH HIGHEST WEEKLY SHEET - PULLS THE HIGHEST WEEK SHEET THAT EXISTS
function fetchWeeklySheet(ss) {
  if (ss = null) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  let week = 0;
  let sheets = ss.getSheets();
  let regex = new RegExp('^'+weeklySheetPrefix+'[0-9]{1,2}'); // sheetName = weeklySheetPrefix [global var] + week [1-18 integer]
  for (let a = 0; a < sheets.length; a++) {
    if (regex.test(sheets[a].getSheetName())) {
      let number = parseInt(sheets[a].getName().replace(weeklySheetPrefix,''));
      if (week < number) {
        week = number;
      }
    }
  }
  if (week > 0) {
    return week;
  } else {
    ss.toast('Issue detecting previous forms or which week is next, assuming week 1');
    return 1;
  }
}

// FETCH WEEK COMPARE - PULL INFO FROM MULTIPLE SOURCES AND COMPARE; USED WITHIN 'nextWeek' FUNCTION
function fetchWeekCompare(ss) {
  // Default to week 1
  let ssWeek, apiWeek, propWeek = 1;

  // Leverage existing script of 'fetchWeek' for API week
  apiWeek = fetchWeek(true);
  weeks = fetchWeeks();

  // Get most recently recorded set of Form responses logged in Script Properties
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    let picksProp = scriptProperties.getProperty('picks_1');

    // Loop to find the most recent properties value logged
    while (picksProp != null) {
      propWeek++;
      picksProp = scriptProperties.getProperty('picks_'+ propWeek);
    }
    propWeek--;
  }
  catch (err) {
    Logger.log('Error fetching variable for week ' + apiWeek + ' property of stored responses, assuming it has not occured.')
  }

  // Get recorded value of which week within the CONFIG sheet
  try {
    // Get spreadsheet if not provided as input
    if (ss == null) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    ssWeek = ss.getRangeByName('WEEK').getValue();
  }
  catch (err) {
    Logger.log('Issue with fetching a week from the spreadsheet, returning API response for week\r\n\r\n' + err.stack);
    ssWeek = null;
  }
  return {
          'ss': ssWeek,
          'api':apiWeek,
          'prop':propWeek
        };
}

// GENERATES HEX GRADIENT - Provide a start and end and a count of values and this function generates a HEX gradient. Midpoint value is optional.
function hexGradient(start, end, count, midpoint) { // start and end in either 3 or 6 digit hex values, count is total values in array to return
  if (count < 2 || count.isNaN) {
    Logger.log('ERROR: Please provide a \'count\' value of 2 or greater');
    return null;
  } else {
    count = Math.ceil(count);
    if (midpoint == null || midpoint == undefined) {
      // strip the leading # if it's there
      start = start.replace(/^\s*#|\s*$/g, '');
      end = end.replace(/^\s*#|\s*$/g, '');

      // convert 3 char codes --> 6, e.g. `E0F` --> `EE00FF`
      if(start.length == 3){
        start = start.replace(/(.)/g, '$1$1');
      }

      if(end.length == 3){
        end = end.replace(/(.)/g, '$1$1');
      }

      let arr = ['#'+start];
      let tmpRed, tmpGreen, tmpBlue;

      // get colors
      let startRed = parseInt(start.substr(0, 2), 16),
          startGreen = parseInt(start.substr(2, 2), 16),
          startBlue = parseInt(start.substr(4, 2), 16);
      let endRed = parseInt(end.substr(0, 2), 16),
          endGreen = parseInt(end.substr(2, 2), 16),
          endBlue = parseInt(end.substr(4, 2), 16);
      let stepRed = (endRed-startRed)/(count-1),
          stepGreen = (endGreen-startGreen)/(count-1),
          stepBlue = (endBlue-startBlue)/(count-1);
      

      for (let a = 1; a < count-1; a++) {
        // calculate the step differential for each color
        tmpRed = ((stepRed * a) + startRed).toString(16).split('.')[0];
        tmpGreen = ((stepGreen * a) + startGreen).toString(16).split('.')[0];
        tmpBlue = ((stepBlue * a) + startBlue).toString(16).split('.')[0];
        // ensure 2 digits by color
        if( tmpRed.length == 1 ) tmpRed = '0' + tmpRed;
        if( tmpGreen.length == 1 ) tmpGreen = '0' + tmpGreen;
        if( tmpBlue.length == 1 ) tmpBlue = '0' + tmpBlue;
        arr.push(('#' + tmpRed + tmpGreen + tmpBlue).toUpperCase());
      }
      arr.push('#'+end);
      return arr;
    } else {
      count = Math.ceil(count);
      if (count % 2 == 0) {
        count++;
        // Logger.log('Even number provided with midpoint, increasing count to ' + count);
      }
      let half = Math.ceil(count/2);
      let arr = hexGradient(start,midpoint,half);
      arr.pop();
      let arr2 = hexGradient(midpoint,end,half);
      arr = arr.concat(arr2);
      return arr;
    }
  }
}

// ENSURE ARRAY IS RECTANGULAR - a function to ensure that an array has blank values if it fails to have a full set of columns per row
function makeArrayRectangular(arr) {
  const maxLength = Math.max(...arr.map(row => row.length));
  for (let a = 0; a < arr.length; a++) {
    // While the row's length is less than the maximum length, push a placeholder value
    while (arr[a].length < maxLength) {
      arr[a].push('');
    }
  }
  return arr;
}


// GET TIMEZONE
function timezoneSet() {
  // Get the value for the script property timezone
  const scriptProperties = PropertiesService.getScriptProperties();
  const tz = scriptProperties.getProperty('tz');
  if (tz != null) {
    return true
  } else {
    Logger.log('No timezone confirmation has been done yet');
    return false
  }
}

// SET PROPRTY - sets a script property based on an inputted name (string) and a value (string/array/object) (essentially this ia global variable)
function setProperty(property,value){
  const scriptProperties = PropertiesService.getScriptProperties();
  if (typeof value === 'object' && !Array.isArray(value) && value !== null) {
    scriptProperties.setProperty(property,JSON.stringify(value));
  } else {
    scriptProperties.setProperty(property,value);
  }
}

// OPEN URL - Quick script to open a new tab with the newly created form, in this case
function openUrl(url,week){
  if (!url || typeof url !== 'string') {
    throw new Error("Invalid URL provided.");
  }
  if (week == null) {
    week = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('WEEK').getValue();
  }
  if (week == undefined) {
    week = fetchWeek();
  }

  // Create the HTML content with the Montserrat font
  let htmlContent = `
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700&display=swap" rel="stylesheet">
    <div style="font-family: 'Montserrat', sans-serif; text-align: center; padding: 20px;">
      <p style="font-size: 22px;"><a href="${url}" target="_blank" style="font-weight: bold;">Click for Week ` + week + ` Form</a></p>
    </div>
  `;

  let htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(350)
    .setHeight(180);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}
// FAILURE TO OPEN ON SOME MACHINES
//   let html = HtmlService.createHtmlOutput('<script>window.open("' + url + '", "_blank");google.script.host.close();</script>')
//     .setHeight(15)
//     .setWidth(100);
//   SpreadsheetApp.getUi().showModalDialog(html, 'Opening Form...');
// }

// VIEW USER PROPERTIES - Shows all set variables within Google user properties
// This is a back-end and unused script, these variables aren't isolated to the sheet/script but used by the form/sheet connection when triggering onSubmit calls
function viewUserProperties() {
  let userProperties = PropertiesService.getUserProperties().getProperties();
  Logger.log('User Properties:');
  for (let key in userProperties) {
    Logger.log(key + ': ' + userProperties[key]);
  }
}
