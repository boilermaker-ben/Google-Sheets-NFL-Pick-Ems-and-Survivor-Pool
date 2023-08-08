// PRELIM SETUP - Creation of all needed initial sheets, prompt to import NFL
function runFirst() {
  var year = fetchYear();
  var week = fetchWeek();
  var weeks = fetchTotalWeeks();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create all default sheets if they don't exist
  if ( ss.getSheetByName('SUMMARY') == null ) {
    ss.insertSheet('SUMMARY');
    Logger.log('Summary sheet created');
  }
  if ( ss.getSheetByName('WINNERS') == null ) {
    ss.insertSheet('WINNERS');
    Logger.log('Winners sheet created');  
  }
  if ( ss.getSheetByName('SURVIVOR') == null ) {
    ss.insertSheet('SURVIVOR');
    Logger.log('Survivor sheet created');
  }
  if ( ss.getSheetByName('MNF') == null ) {
    ss.insertSheet('MNF');
    Logger.log('Monday Night Football sheet created');
  }
  if ( ss.getSheetByName('OVERALL') == null ) {
    ss.insertSheet('OVERALL');
    Logger.log('Overall sheet created');
  }
  if ( ss.getSheetByName('OVERALL_RANK') == null ) {
    ss.insertSheet('OVERALL_RANK');
    Logger.log('OverallRank sheet created');
  }
  if ( ss.getSheetByName('OVERALL_PCT') == null ) {
    ss.insertSheet('OVERALL_PCT');
    Logger.log('OverallPct sheet created');
  }
  if ( ss.getSheetByName('FORM') == null ) {
    ss.insertSheet('FORM');
  }
  Logger.log('Form sheet created');
  if ( ss.getSheetByName('MEMBERS') == null ) {
    ss.insertSheet('MEMBERS');
  }
  Logger.log('Members sheet created');
  var sheet = ss.getSheetByName('Sheet1');
  if ( sheet != null ) {
    ss.deleteSheet(sheet);
  }
  // Pull in NFL Schedule data and create sheet
  fetchNFL();
  Logger.log('Fetched NFL Schedule');
  
  // Run through all sheet information population
  try {
    // DATA All matchup and NFL data that's been imported by 'fetchNFL' script
    data = ss.getRangeByName('NFL_' + year).getValues(); //Grab again if wasn't populated before      
        
    var members = memberList();
    var totalMembers = members.length;
    Logger.log('Total Members: ' + totalMembers);

    // Creates Member sheet (calling function)
    var member = memberSheet(members);
    Logger.log('Deployed Members sheet');

    // Creates Weekly Sheets for the Current Week (calling function)
    var weekSheet = weeklySheet(year,week,members,false);
    Logger.log('Deployed Weekly sheet');

    // Creates Overall Record Sheet (calling function)
    var overall = overallSheet(year,weeks,members);
    Logger.log('Deployed Overall sheet');

    // Creates Overall Rank Record Sheet (calling function)
    var overallRnk = overallRankSheet(year,weeks,members);
    Logger.log('Deployed Overall Rank sheet');

    // Creates Overall Percent Record Sheet (calling function)
    var overallPct = overallPctSheet(year,weeks,members);
    Logger.log('Deployed Overall Percent sheet');

    // Creates Survivor Sheet (calling function)
    var survivor = survivorSheet(year,weeks,members);
    Logger.log('Deployed Survivor sheet');

    // Creates Summary Record Sheet (calling function)
    var summary = summarySheet(year,members);
    Logger.log('Deployed Summary sheet');

    // Creates Winners Sheet (calling function)
    var winners = winnersSheet(year,weeks,members);
    Logger.log('Deployed Winners sheet');

    // Creates MNF Sheet (calling function)
    var mnf = mnfSheet(year,weeks,members);
    Logger.log('Deployed MNF Sheet');
    
    // Creates previous weekly sheets if needed
    var ui = SpreadsheetApp.getUi();
    if (week > 2) {
      var prompt = ui.alert('There are previous weeks that can be configured that you can populate the previous data, create those now?', ui.ButtonSet.YES_NO)
    } else if (week == 2) {
      var prompt = ui.alert('There is a previous week that can be configured that you can populate the previous data, create that now?', ui.ButtonSet.YES_NO)
    } 
    if (prompt == 'YES') {
      weeklySheetCreation();
      Logger.log('Created previous week(s)');
    }

    Logger.log('Initial setup complete, proceeding to update menu');
    var lockMembers = ui.alert('Keep members unlocked and allow new members to be added to the pool through the Google Form?', ui.ButtonSet.YES_NO);
    if (lockMembers == 'YES') {
      createMenuUnlockedWithTriggerFirst();
    } else {
      createMenuLockedWithTriggerFirst();
    }
    Logger.log('Created final menu.');
    
    // Hide unnecessary sheets -- comment out with '//' to automatically hide these sheets
    
    ss.getSheetByName('NFL_' + year).hideSheet();
    //ss.getSheetByName('FORM').hideSheet();
    //overall.hideSheet();
    //overallRnk.hideSheet();
    //overallPct.hideSheet();
    //survivor.hideSheet();
    //summary.hideSheet();
    //winners.hideSheet();

  var createForm = ui.alert('Create first Pick \'Ems submission Google Form now?', ui.ButtonSet.YES_NO);
  if (createForm == 'YES'){
    formFiller(true);
    Logger.log('Created initial form');
  } else {
    ui.alert('To create the initial form and create each new weekly form, use the \'Pick\'Ems\' menu function \'Update Form\'.', ui.ButtonSet.OK);
  }

  Logger.log('You\'re all set, have fun!');

  }
  catch (err) {
    Logger.log('runFirstStack ' + err.stack);
  }    
}

//------------------------------------------------------------------------
// CREATE MENU - this is the ideal setup once the sheet has been configured and the data is all imported
function createMenuUnlocked(trigger) {
  var menu = SpreadsheetApp.getUi().createMenu('Pick\'Ems')
  menu.addItem('Update Form','formFiller')
      .addItem('Open Form','openForm')
      .addItem('Check NFL Scores','fetchNFLScores')
      .addSeparator()
      .addItem('Check Responses','formCheckAlertCall')
      .addItem('Import Picks','dataTransfer')
      .addItem('Import Thursday Picks','dataTransferTNF')
      .addSeparator()
      .addItem('Add Member','memberAdd')
      .addItem('Lock Members','createMenuLockedWithTrigger')
      .addSeparator()
      .addItem('Update NFL Schedule', 'fetchNFL')
      .addItem('Rebuild Calculations', 'allFormulasUpdate')
      .addToUi();
  if (trigger == true) {
    deleteTriggers()
    var id = SpreadsheetApp.getActiveSpreadsheet().getId();
    ScriptApp.newTrigger('createMenuUnlocked')
      .forSpreadsheet(id)
      .onOpen()
      .create();
  }
}

// CREATE MENU For general use with locked MEMBERS sheet
function createMenuLocked(trigger) {
  var menu = SpreadsheetApp.getUi().createMenu('Pick\'Ems')
  menu.addItem('Update Form','formFiller')
      .addItem('Open Form','openForm')
      .addItem('Check NFL Scores','fetchNFLScores')
      .addSeparator()
      .addItem('Check Responses','formCheckAlertCall')
      .addItem('Import Picks','dataTransfer')
      .addItem('Import Thursday Picks','dataTransferTNF')
      .addSeparator()
      .addItem('Reopen Members','createMenuUnlockedWithTrigger')
      .addSeparator()
      .addItem('Update NFL Schedule', 'fetchNFL')
      .addItem('Rebuild Calculations', 'allFormulasUpdate')      
      .addToUi();
  if (trigger == true) {
    deleteTriggers()
    var id = SpreadsheetApp.getActiveSpreadsheet().getId();
    ScriptApp.newTrigger('createMenuLocked')
      .forSpreadsheet(id)
      .onOpen()
      .create();
  }
}

// CREATE MENU UNLOCKED MEMBERSHIP with Trigger Input
function createMenuUnlockedWithTrigger(init) {
  createMenuUnlocked(true);
  membersSheetUnlock();
  Logger.log('Menu updated to an open membership, MEMBERS unlocked');
  if (init != true) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('New entrants will be allowed through the Google Form and through the \'Pick\'Ems\' menu function: \'Add Member\'. Run \'Lock Members\' to prevent new additions in the Google Form and menu.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
// CREATE MENU UNLOCKED MEMBERSHIP with Trigger Input on first pass (skips prompt)
function createMenuUnlockedWithTriggerFirst() {
  createMenuUnlockedWithTrigger(true);
}

// CREATE MENU LOCKED MEMBERSHIP with Trigger Input
function createMenuLockedWithTrigger(init) {
  createMenuLocked(true);
  membersSheetLock();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MEMBERS').hideSheet();
  Logger.log('Menu updated to a locked membership, MEMBERS locked');
  if (init != true) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('New entrants will not be allowed through the Google Form nor through the menu unless \'Reopen Members\' script is run. Run \'Reopen Members\' to allow new additions in the Google Form and menu', SpreadsheetApp.getUi().ButtonSet.OK);
    var prompt = ui.alert('Recreate form to take into account this change?', ui.ButtonSet.YES_NO);
    if (prompt == 'YES'){
      formFiller(true)
    }
  }
}
// CREATE MENU LOCKED MEMBERSHIP with Trigger Input on first pass (skips prompt)
function createMenuLockedWithTriggerFirst() {
  createMenuLockedWithTrigger(true);
}

//------------------------------------------------------------------------
// MEMBERS List for editing in future years
function memberList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var members = [];
  try {
    members = ss.getRangeByName('MEMBERS').getValues();
    if (members[0] == '') {
      throw new Error();  
    }
    return members;
  } 
  catch (err) {
    Logger.log('No member list found, prompting for creation... [Go to spreadsheet]');
    var ui = SpreadsheetApp.getUi();
    var prompt = ui.prompt('Entery preliminary list of members as a comma separated list.', ui.ButtonSet.OK_CANCEL)
    if ( prompt.getSelectedButton() == 'OK') {
      var arr = [];
      membersStr = prompt.getResponseText();
      members = membersStr.split(',');
      for (var a = 0; a < members.length; a++) {
        arr.push([toTitleCase(members[a].trim())]);
      }
      members = arr;
    } else {
      ss.toast('Canceled members list creation, try again and enter at least one name.');
    }
    return members;
  }
}

// MEMBERS Addition for adding new members later in the season
function memberAdd(name) {
  var members = [];
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prompt, a;
  var proceed = 1;
  try {
    var membersSheet = ss.getSheetByName('MEMBERS');
    var range = ss.getRangeByName('MEMBERS');
    var members = range.getValues();
  }
  catch (err) {
    proceed = 0;
    Logger.log('memberAdd 1 ' + err.message);
    var prompt = ui.alert('No members sheet created yet, do that now?', ui.ButtonSet.YES_NO)
    if (prompt == "YES") {
      proceed = 1;
      //memberSheet(memberList());
      var membersSheet = ss.getSheetByName('MEMBERS');
      var range = ss.getRangeByName('MEMBERS');
      var members = range.getValues();
    }
  }
  if (proceed == 1 && name == null) {   
    prompt = ui.prompt('Please enter one member to add:', ui.ButtonSet.OK_CANCEL)
    name = prompt.getResponseText();
  } 
  var response;
  try {
    response = prompt.getSelectedButton() != 'CANCEL'
  }
  catch (err) {
    response = 'OK';
  }
  if (name != null && proceed == 1 && response != 'CANCEL') {
    name = toTitleCase(name);
    members.push([name]);
    membersSheet.insertRows(1,1);
    var range = membersSheet.getRange(1,1,membersSheet.getMaxRows(),1);
    range.setValues(members);
    ss.setNamedRange('MEMBERS',range);

    var year = fetchYear();
    var first, week;
    try {
      week = ss.getRangeByName('WEEK').getValue();
      first = false;
      if (week == null) {
        var week = fetchWeek();
        first = true;
      }
    }
    catch (err){
      Logger.log('No Week Set Yet, checking API info');
      week = fetchWeek();
      first = true;
    }
    var weeks = fetchTotalWeeks();
    //-------------------
    // Update WEEKLY SHEETS
    Logger.log('Working on week ' + week);
    nflData = ss.getRangeByName('NFL_'+year).getValues();
    weeklySheet(year,week,members,true);
    ss.toast('Recreated weekly sheet for week ' + week)

    // Create all default sheets if they don't exist
    Logger.log('Fetched NFL Schedule');
    if ( ss.getSheetByName('SUMMARY') == null ) {
      ss.insertSheet('SUMMARY');
      Logger.log('Summary sheet created');
    }
    if ( ss.getSheetByName('WINNERS') == null ) {
      ss.insertSheet('WINNERS');
      Logger.log('Winners sheet created');  
    }
    if ( ss.getSheetByName('SURVIVOR') == null ) {
      ss.insertSheet('SURVIVOR');
      Logger.log('Survivor sheet created');
    }
    if ( ss.getSheetByName('MNF') == null ) {
      ss.insertSheet('MNF');
      Logger.log('Monday Night Football sheet created');
    }    
    if ( ss.getSheetByName('OVERALL') == null ) {
      ss.insertSheet('OVERALL');
      Logger.log('Overall sheet created');
    }
    if ( ss.getSheetByName('OVERALL_RANK') == null ) {
      ss.insertSheet('OVERALL_RANK');
      Logger.log('OverallRank sheet created');
    }
    if ( ss.getSheetByName('OVERALL_PCT') == null ) {
      ss.insertSheet('OVERALL_PCT');
      Logger.log('OverallPct sheet created');
    }

    members = memberList();
    Logger.log(members);
    // Creates Overall Record Sheet (calling function)
    var overall = overallSheet(year,weeks,members);
    Logger.log('Recreated Overall sheet');

    // Creates Overall Rank Record Sheet (calling function)
    var overallRnk = overallRankSheet(year,weeks,members);
    Logger.log('Recreated Overall Rank sheet');

    // Creates Overall Percent Record Sheet (calling function)
    var overallPct = overallPctSheet(year,weeks,members);
    Logger.log('Recreated Overall Percent sheet');

    // Creates Survivor Sheet (calling function)
    var survivor = survivorSheet(year,weeks,members);
    Logger.log('Recreated Survivor sheet');

    // Creates Summary Record Sheet (calling function)
    var summary = summarySheet(year,members);
    Logger.log('Recreated Summary sheet');

    // Creates Winners Sheet (calling function)
    var winners = winnersSheet(year,weeks,members);
    Logger.log('Recreated Winners sheet');
    
    // Creates MNF Sheet (calling function)
    var mnf = mnfSheet(year,weeks,members);
    Logger.log('Recreated MNF Sheet');
    

  } else {
    ss.toast('Member add canceled due to no MEMBERS sheet.');
  }
}

//------------------------------------------------------------------------
// FETCH CURRENT YEAR
function fetchYear() {
  var obj;
  obj = JSON.parse(UrlFetchApp.fetch('http://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard').getContentText());
  var year = obj['season']['year'];
  return year;
}

//------------------------------------------------------------------------
// FETCH CURRENT WEEK
function fetchWeek() {
  var obj;
  obj = JSON.parse(UrlFetchApp.fetch('http://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard').getContentText());
  var week = 1;
  if(obj['events'][0]['season']['slug'] != 'preseason'){
    week = obj['week']['number'];
  }
  Logger.log('Current week: ' + week);
  return week;
}

//------------------------------------------------------------------------
// FETCH TOTAL WEEKS
function fetchTotalWeeks() {
  var content = UrlFetchApp.fetch('http://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard').getContentText();
  var obj = JSON.parse(content);
  var calendar = obj['leagues'][0]['calendar'];
  for (var a = 0; a < calendar.length; a++) {
    if (calendar[a]['value'] == 2) {
      var weeks = calendar[a]['entries'].length;
      break;
    }
  }
  return weeks;
}

//------------------------------------------------------------------------
// ESPN TEAMS - Fetches the ESPN-available API data on NFL teams
function fetchTeamsESPN() {
  var year = fetchYear(); // First array value is year
  var obj = JSON.parse(UrlFetchApp.fetch('http://fantasy.espn.com/apis/v3/games/ffl/seasons/' + year + '?view=proTeamSchedules').getContentText());
  var objTeams = obj['settings']['proTeams'];
  return objTeams;
}

//------------------------------------------------------------------------
// NFL TEAM INFO - script to fetch all NFL data for teams
function fetchNFL() {
  // Calls the linked spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    
  // Declaration of script variables
  var maxRows;
  var maxCols;
  var year = fetchYear();
  var arr = [];
  var nfl = [];
  var abbr;
  var name;
  var location;
  var objTeams = fetchTeamsESPN();
  var teamsLen;
  var espnId = [];
  var espnAbbr = [];
  var espnName = [];
  var espnLocation = [];
  var teamsLen = objTeams.length;
  
  for (var i = 0 ; i < teamsLen ; i++ ) {
    arr = [];
    if(objTeams[i]['id'] != 0 ) {
      abbr = objTeams[i]['abbrev'].toUpperCase();
      name = objTeams[i]['name'];
      location = objTeams[i]['location'];
      espnId.push(objTeams[i]['id']);
      espnAbbr.push(abbr);
      espnName.push(name);
      espnLocation.push(location);
      arr = [objTeams[i]['id'],abbr,location,name,objTeams[i]['byeWeek']];
      nfl.push(arr);
    }
  }

  var sheet;
  var range;
  var ids = [];
  var abbrs = [];
  for ( var i = 0 ; i < espnId.length ; i++ ) {
    ids.push(espnId[i].toFixed(0));
    abbrs.push(espnAbbr[i]);
  }
  // Declaration of variables
  var arr = [];
  var schedule = [];
  var home = [];
  var dates = [];
  var allDates = [];
  var hours = [];
  var allHours = [];
  var minutes = [];
  var allMinutes = [];
  var location = [];
  var byeIndex;
  var id;
  var data;
  var j;
  var k;
  var date;
  var hour;
  var minute;
  var weeks = 1; 
  for(var key in objTeams[0]['proGamesByScoringPeriod']){
    weeks++;
  }
  
  for ( i = 0 ; i < teamsLen ; i++ ) {
    
    arr = [];
    home = [];
    dates = [];
    hours = [];
    minutes = [];
    byeIndex = objTeams[i]['byeWeek'].toFixed(0);
    if ( byeIndex != 0 ) {
      id = objTeams[i]['id'].toFixed(0);
      arr.push(abbrs[ids.indexOf(id)]);
      home.push(abbrs[ids.indexOf(id)]);
      dates.push(abbrs[ids.indexOf(id)]);
      hours.push(abbrs[ids.indexOf(id)]);
      minutes.push(abbrs[ids.indexOf(id)]);
      for (var j = 1 ; j <= weeks ; j++ ) {
        if ( j == byeIndex ) {
          arr.push('BYE');
          home.push('BYE');
          dates.push('BYE');
          hours.push('BYE');
          minutes.push('BYE');
        } else {
          if ( objTeams[i]['proGamesByScoringPeriod'][j][0]['homeProTeamId'].toFixed(0) === id ) {
            arr.push(abbrs[ids.indexOf(objTeams[i]['proGamesByScoringPeriod'][j][0]['awayProTeamId'].toFixed(0))]);
            home.push(1);
            date = new Date(objTeams[i]['proGamesByScoringPeriod'][j][0]['date'])
            dates.push(date);
            hour = date.getHours()
            hours.push(hour);
            minute = date.getMinutes();
            minutes.push(minute);
          } else {
            arr.push(abbrs[ids.indexOf(objTeams[i]['proGamesByScoringPeriod'][j][0]['homeProTeamId'].toFixed(0))]);
            home.push(0);
            date = new Date(objTeams[i]['proGamesByScoringPeriod'][j][0]['date'])
            dates.push(date);
            hour = date.getHours()
            hours.push(hour);
            minute = date.getMinutes();
            minutes.push(minute);
          }
        }
      }
      schedule.push(arr);
      location.push(home);
      allDates.push(dates);
      allHours.push(hours);
      allMinutes.push(minutes)
    }
  }
  
  // This section creates a nice table to be used for lookups and queries about NFL season
  var week;
  var awayTeam;
  var awayTeamName;
  var awayTeamLocation;
  var homeTeam;
  var homeTeamName;
  var homeTeamLocation;
  var formData = [];
  var mnf;
  var day;
  var dayName;
  arr = [];
  
  for ( j = 0; j < (teamsLen - 1); j++ ) {
    for ( k = 1; k <= 18; k++ ) {
      if (location[j][k] == 1) {
        week = k;
        awayTeam = schedule[j][k];
        awayTeamName = espnName[espnAbbr.indexOf(awayTeam)];
        awayTeamLocation = espnLocation[espnAbbr.indexOf(awayTeam)];
        homeTeam = schedule[j][0];
        homeTeamName = espnName[espnAbbr.indexOf(homeTeam)];
        homeTeamLocation = espnLocation[espnAbbr.indexOf(homeTeam)];
        date = allDates[j][k];
        hour = allHours[j][k];
        minute = allMinutes[j][k];
        day = date.getDay();
        mnf = 0;
        if ( day == 1 ) {
          mnf = 1;
          dayName = 'Monday';
        } else if ( day == 0 ) {
          dayName = 'Sunday';
        } else if ( day == 4 ) {
          day = -3;
          dayName = 'Thursday';
        } else if ( day == 5 ) {
          day = -2;
          dayName = 'Friday';
        } else if ( day == 6 ) {
          day = -1;
          dayName = 'Saturday';
        }
        arr = [week, date, day, hour, minute, dayName, awayTeam, homeTeam, awayTeamLocation, awayTeamName, homeTeamLocation, homeTeamName];
        formData.push(arr);
      }
    }
  }
  var headers = ['week','date','day','hour','minute','dayName','awayTeam','homeTeam','awayTeamLocation','awayTeamName','homeTeamLocation','homeTeamName'];
  var sheetName = 'NFL_' + year;
  var rows = formData.length + 1;
  var columns = formData[0].length;
  
  sheet = ss.getActiveSheet();
  if ( sheet.getSheetName() == 'Sheet1' && ss.getSheetByName(sheetName) == null) {
    sheet.setName(sheetName);
  }
  sheet = ss.getSheetByName(sheetName);  
  if (sheet == null) {
    ss.insertSheet(sheetName,0);
    sheet = ss.getSheetByName(sheetName);
  }
  
  maxRows = sheet.getMaxRows();
  if (maxRows < rows){
    sheet.insertRows(maxRows,rows - maxRows - 1);
  } else if (maxRows > rows){
    sheet.deleteRows(rows,maxRows - rows);
  }
  maxCols = sheet.getMaxColumns();
  for ( j = maxCols; j < columns; j++){
    sheet.insertColumnAfter(j);
  } 
  if (maxCols > columns){
    sheet.deleteColumns(columns,maxCols - columns);
  }
  sheet.setColumnWidths(1,columns,30);
  sheet.setColumnWidth(2,60);
  sheet.setColumnWidth(6,60);
  sheet.setColumnWidths(9,4,80);
  sheet.clear();
  range = sheet.getRange(1,1,1,columns)
  range.setValues([headers]);
  ss.setNamedRange(sheetName+'_HEADERS',range);
  range = sheet.getRange(1,1,rows,columns)
  range.setFontSize(8);
  range.setVerticalAlignment('middle');  
  range = sheet.getRange(2,1,formData.length,columns)
  range.setValues(formData);
  ss.setNamedRange(sheetName,range);
  range.setHorizontalAlignment('left');  
  range.sort([{column: 1, ascending: true},{column: 2, ascending: true},{column: 4, ascending: true},
              {column:  5, ascending: true},{column: 6, ascending: true},{column: 8, ascending: true}]); 
  sheet.getRange(1,3).setNote('-3: Thursday, -2: Friday, -1: Saturday, 0: Sunday, 1: Monday, 2: Tuesday');
  sheet.protect().setDescription(sheetName);
  try {
    sheet.hideSheet();
  }
  catch (err){
  }
  ss.toast('Imported all NFL schedule data');
}

//------------------------------------------------------------------------
// NFL ACTIVE WEEK SCORES - script to check and pull down any completed matches and record them to the sheet
function fetchNFLScores(){
  var url = 'http://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard'
  
  var obj = JSON.parse(UrlFetchApp.fetch(url));
  var ui = SpreadsheetApp.getUi();
  var week;
  if(obj['events'][0]['season']['slug'] == 'preseason'){
    week = obj['week']['number'];
    alert = ui.alert('Regular season not yet started. \r\n Currently preseason week ' + week + '.', ui.ButtonSet.OK);
    alert = 'CANCEL';
  } else {
    week = obj['week']['number'];
    var year = obj['season']['year'];
    var week = obj['week']['number'];
    var games = obj['events'];
    var count = games.length;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var writeRange = ss.getRangeByName('NFL_'+year+'_'+week);
    var writeCell;
    var sheet = writeRange.getSheet();  
    var tiebreakerCell = sheet.getRange(writeRange.getRow(),writeRange.getLastColumn()+1);
    var tiebreakerValue = tiebreakerCell.getValue();
    var winners = [];
    var tiebreakers = [];
    var undecided = 0;
    var existing = 0;
    var completed = false;
    var remaining = 0;

    for (let a = 0; a < count; a++){
      if (games[a]['status']['type']['completed']==true) {
        competitors = games[a]['competitions'][0]['competitors']
        if (competitors[0]['winner'] == true) {
          winners.push(competitors[0]['team']['abbreviation']);
          tiebreakers.push(parseInt(competitors[0]['score'])+parseInt(competitors[1]['score']));
        } else if (competitors[1]['winner'] == true) {
          winners.push(competitors[1]['team']['abbreviation']);
          tiebreakers.push(parseInt(competitors[0]['score'])+parseInt(competitors[1]['score']));
        } else {
          winners.push(competitors[0]['team']['abbreviation'] + '=' + competitors[1]['team']['abbreviation']);
          tiebreakers.push(parseInt(competitors[0]['score'])+parseInt(competitors[1]['score']));          
        }
      } else {
        undecided++;
      }
    }
    var matchups = sheet.getRange(writeRange.getRow()-1,writeRange.getColumn(),1,writeRange.getNumColumns()).getValues().flat();
    if (games.length == winners.length){
      completed = true;
    }
    for (a = 0; a < matchups.length; a++){
      if (sheet.getRange(writeRange.getRow(),writeRange.getColumn()+a).getValue() != '' && (sheet.getRange(writeRange.getRow(),writeRange.getColumn()+a).getValue() == 'TIE' || winners.indexOf(matchups[a].split('@')[0]) >= 0 || winners.indexOf(matchups[a].split('@')[1]) >= 0 )){
        existing++;
      }
    }
    remaining = winners.length - existing;
    var alert = 'CANCEL';
    if (completed && remaining == 1) {
      alert = ui.alert('WEEK ' + week + ' COMPLETE: \r\n Record the final unmarked match and tiebreaker?', ui.ButtonSet.OK_CANCEL);
    } else if (completed && remaining > 0){
      alert = ui.alert('WEEK ' + week + ' COMPLETE: \r\n Record all unmarked matches and tiebreaker?', ui.ButtonSet.OK_CANCEL);
    } else if (remaining > 1 && undecided > 1) {
      alert = ui.alert('WEEK ' + week + ' INCOMPLETE: \r\n Record ' + remaining + ' unmarked, completed matches? \r\n (There are ' + undecided + ' undecided matches remaining)', ui.ButtonSet.OK_CANCEL);
    } else if (remaining > 1 && undecided == 1) {
      alert = ui.alert('WEEK ' + week + ' INCOMPLETE: \r\n Record ' + remaining + ' unmarked, completed matches? \r\n (There is one undecided match)', ui.ButtonSet.OK_CANCEL);  
    } else if (remaining == 1) {
      alert = ui.alert('WEEK ' + week + ' INCOMPLETE: \r\n Record the one unmarked, completed match? \r\n (There are ' + undecided + ' matches remaining)', ui.ButtonSet.OK_CANCEL);
    } else if (remaining == 1 && undecided > 1) {
      alert = ui.alert('WEEK ' + week + ' INCOMPLETE: \r\n All completed games recorded. \r\n (There are ' + undecided + ' undecided matches remaining)', ui.ButtonSet.OK_CANCEL);
    } else if (remaining == 0 && undecided == 1) {
      alert = ui.alert('WEEK ' + week + ' INCOMPLETE: \r\n All completed games recorded. \r\n (There is one undecided match)', ui.ButtonSet.OK_CANCEL);  
    } else if (remaining == 0 && tiebreakerValue == '' && undecided == 0) {
      alert = ui.alert('WEEK ' + week + ' COMPLETE: \r\n Record tiebreaker?', ui.ButtonSet.OK_CANCEL);
    } else if (remaining == 0 && undecided == 0) {
      alert = ui.alert('WEEK ' + week + ' COMPLETE: \r\n All matches and tiebreaker recorded.', ui.ButtonSet.OK);
      alert = 'CANCEL';
    } else {
      alert = ui.alert('WEEK ' + week + ' NOT YET STARTED: \r\n No matches completed yet.', ui.ButtonSet.OK);
      alert = 'CANCEL';
    }
    if (alert == 'OK') {
      for (a = 0; a < matchups.length; a++){
        let teamA = matchups[a].split('@')[0];
        let teamB = matchups[a].split('@')[1];
        writeCell = sheet.getRange(writeRange.getRow(),writeRange.getColumn()+a);
        if ((winners.indexOf(teamA) >= 0 && writeCell.getValue() == teamA) || (winners.indexOf(teamB) >= 0 && writeCell.getValue() == teamB)){
          matchups[a] = '';
        }
        if (winners.indexOf(teamA) >= 0 && writeCell.getValue() == ''){
          writeCell.setValue(teamA);
          matchups[a] = '';
        }
        if (winners.indexOf(teamB) >= 0 && writeCell.getValue() == ''){
          writeCell.setValue(teamB);
          matchups[a] = '';
        }
        if (completed && a == (matchups.length-1) && (winners.indexOf(teamA) >= 0 || winners.indexOf(teamB) >= 0)){
          tiebreakerCell.setValue(tiebreakers[a]);
        }
      }
      var rule, rules;
      var args = [];
      for (let b = 0; b < winners.length; b++) {
        args = [];
        if (winners[b].split('=')[1] != undefined){
          for (a = 0; a < matchups.length; a++){
            if (matchups[a] != '') {
              let teamA = matchups[a].split('@')[0];
              let teamB = matchups[a].split('@')[1];
              matchups[a] = '';
              writeCell = sheet.getRange(writeRange.getRow(),writeRange.getColumn()+a);
              rule = writeCell.getDataValidation();
              args = [teamA,teamB,'TIE'];
              rules = SpreadsheetApp.newDataValidation().requireValueInList(args, true).build();
              writeCell.setDataValidation(rules)
              writeCell.setValue('TIE');
            }
          }
          if (completed && b == (winners.length-1)) {
            tiebreakerCell.setValue(tiebreakers[b]);
          }          
        }
      }
    }
    if (completed){
      var prompt = ui.alert('WEEK ' + week + ' COMPLETE: \r\n Advance survivor pool?', ui.ButtonSet.YES_NO);
      if ( prompt == 'YES' ){
        ss.getRangeByName('WEEK').setValue(week+1);
      } else {
        ss.toast('Completed scores import');
      }
    } else {
      ss.toast('Completed scores import');
    }
  }
}

//------------------------------------------------------------------------
// FETCHES MNF Boolean for if there's a MNF game this week, provide week (int) and receive mnf output (true/false)
function fetchMNF (week) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mnf = false;
  var nflData = ss.getRangeByName('NFL_2021').getValues();
  for (var b = 0; b < nflData.length; b++) {
    if (nflData[b][0] == week) {
      if (nflData[b][2] == 1){
        mnf = true;
      }
    }
  }
  return mnf;
}

//------------------------------------------------------------------------
// MEMBERS Sheet Creation / Adjustment 
function memberSheet(members) {
  
  if (members == null) {
    members = memberList();
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var totalMembers = members.length;
  
  var sheetName = 'MEMBERS';
  var sheet = ss.getSheetByName(sheetName)
  if (sheet == null) {
    ss.insertSheet(sheetName,0);
    sheet = ss.getSheetByName(sheetName);
  }
  
  var rows = Math.max(members.length,1);
  var maxRows = sheet.getMaxRows();
  if ( rows < maxRows ) {
    sheet.deleteRows(rows,maxRows-rows);
  }
  var maxCols = sheet.getMaxColumns();
  if ( maxCols > 1 ) {
    sheet.deleteColumns(1,maxCols-1);
  }
  var range = sheet.getRange(1,1,rows,1);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('left');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  ss.setNamedRange(sheetName,range);
  if (members.length > 0) {
    sheet.getRange(1,1,totalMembers,1).setValues(members);
  }
  sheet.setColumnWidth(1,120);
  sheet.hideSheet();
  return sheet;
}

// MEMBERS Sheet Check if protected returns true or false
function membersSheetProtected() {
  try {
    var locked = false;
    var protections = SpreadsheetApp.getActiveSpreadsheet().getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var a = 0; a < protections.length; a++) {
      if (protections[a].getDescription() == "MEMBERS PROTECTION") {
        locked = true;
      }
    }
  }
  catch (err) {
    Logger.log('error ' + err.message)
    return locked;
  }
  Logger.log('Membership lock is ' + locked);
  return locked;
}

// MEMBERS Sheet Locking (protection)
function membersSheetLock() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MEMBERS');
  sheet.protect().setDescription('MEMBERS PROTECTION');
  Logger.log('locked MEMBERS');
}

// MEMBERS Sheet Unlocking (remove protection);
function membersSheetUnlock() {
  try {
    var protections = SpreadsheetApp.getActiveSpreadsheet().getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var a = 0; a < protections.length; a++) {
      if (protections[a].getDescription() == "MEMBERS PROTECTION") {
        protections[a].remove();
        Logger.log('unlocked MEMBERS');
      }
    }
  }
  catch (err) {
    Logger.log('error ' + err.message)
  }  
}

//------------------------------------------------------------------------
// WEEKLY Sheet Creation - creates all previous weeks
function weeklySheetCreation(){
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var year = fetchYear();
    var week = fetchWeek();
    var members = memberList();
    var sheetName;
    var sheetsCreated = [];
    var sheetsCreatedString = '';
    for ( var a = week; a >= 1; a-- ) {
      if ( a < 10 ) {
        sheetName = year + '_0' + a;
      } else {
        sheetName = year + '_' + a;
      }
      sheet = ss.getSheetByName(sheetName);  
      weeklySheet(year,a,members,false);
      sheetsCreated.push(a);
    }
    if ( sheetsCreated.length > 1 ) {
      for ( a = 0; a < sheetsCreated.length; a++ ) {
        if ( a < ( sheetsCreated.length - 1 ) ) {
          sheetsCreatedString = sheetsCreatedString.concat(sheetsCreated[a] + ', ');
        } else {
          sheetsCreatedString = sheetsCreatedString.concat(sheetsCreated[a]);
        }
      }
      ss.toast('Created sheets for weeks ' + sheetsCreated);
    } else if ( sheetsCreated.length == 1 ) { 
      ss.toast('Created a sheet for week ' + sheetsCreated);
    } else {
      ss.toast('No new sheets needed');
    }
    Logger.log(sheetsCreated);
    var ui = SpreadsheetApp.getUi();
    ui.alert('Previous weekly sheets created.', ui.ButtonSet.OK);
  }
  catch (err) {
    Logger.log('weeklySheetCreation ' + err.message);
    var ui = SpreadsheetApp.getUi();
    ui.alert('Didn\'t finish, restarting script...', ui.ButtonSet.OK);
    weeklySheetCreation();
  }
}

// WEEKLY Sheet Function - creates a sheet with provided year and week
function weeklySheet(year,week,members,dataRestore) {
  
  if (members == null){
    members = memberList();
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet;
  var sheetName;
  var data = ss.getRangeByName('NFL_' + year).getValues(); //Grab again if wasn't populated before      
  
  var mnf = false;
  var mnfStart;
  var mnfEnd; 
  var diffCount = 5; // Number of results to display for most similar weekly picks
  
  if ( week < 10 ) {
    sheetName = year + '_0' + week;
  } else {
    sheetName = year + '_' + week;
  }

  var totalMembers = members.length;
  var rows = members.length + 3; // Accounting for the top two rows above member rows
  var columns;

  var fresh = false;
  sheet = ss.getSheetByName(sheetName);  
  if (sheet == null) {
    ss.insertSheet(sheetName,ss.getNumSheets()+1);
    sheet = ss.getSheetByName(sheetName);
    fresh = true;
  }

  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();

  if (dataRestore == true && fresh == false){
    var headers = sheet.getRange('A1:1').getValues().flat();
    headers.unshift('COL INDEX ADJUST')
    var tiebreakerCol = headers.indexOf('TIEBREAKER');
    var commentCol = headers.indexOf('COMMENT');
    if (tiebreakerCol  >= 0) {
      var previousDataRange = sheet.getRange(2,5,maxRows-2,tiebreakerCol-4);
      var previousData = previousDataRange.getValues();
      ss.toast('Gathered previous data for week ' + week + ', recreating sheet now');
    }
    if (commentCol  >= 0) {
      var previousCommentRange = sheet.getRange(3,commentCol,maxRows-3,1);
      var previousComment = previousCommentRange.getValues();

    }
  }
  sheet.clear();

  // Removing extra rows, reducing to only member count and the additional 2
  if (maxRows < rows){
    sheet.insertRows(maxRows,rows - maxRows);
    Logger.log('added ' + (rows - maxRows) + ' rows');
  } else if (maxRows > rows){
    sheet.deleteRows(rows,maxRows - rows);
    Logger.log('deleted ' + (maxRows - rows) + ' rows');
  }    
  
  // Insert Members
  sheet.getRange(3,1,totalMembers,1).setValues(members);
  var bottomHeaders = ['PREFERRED','AWAY','HOME'];
  sheet.getRange(rows,1,1,3).setValues([bottomHeaders]);
  
  // Setting header values
  sheet.getRange(1,1).setValue('WEEK ' + week);
  sheet.setColumnWidth(1,120);
  
  sheet.getRange(1,2,2,1).setValues([['TOTAL'],['CORRECT']]);
  sheet.setColumnWidth(2,90);
  
  sheet.getRange(1,3,2,1).setValues([['WEEKLY'],['RANK']]);
  sheet.setColumnWidth(3,90);
  
  sheet.getRange(1,4,2,1).setValues([['PERCENT'],['CORRECT']]);
  sheet.setColumnWidth(4,90);

  // Setting headers for the week's matchups with format of 'AWAY' + '@' + 'HOME', then creating a data validation cell below each
  var rule,matches = 0;
  var column = 5;
  for ( var j = 0; j < data.length; j++ ) {
    if ( data[j][0] == week ) {
      matches++;
      if ( data[j][2] == 1 ) {
        mnf = true;
      }
      sheet.getRange(1,column).setValue(data[j][6] + '@' + data[j][7]);
      if ( data[j][2] == 1 ) {
        if ( mnfStart == undefined ) {
          mnfStart = column;
        }
        mnfEnd = column;
      }
      rule = SpreadsheetApp.newDataValidation().requireValueInList([data[j][6],data[j][7]], true).build();
      sheet.getRange(2,column).setDataValidation(rule);
      sheet.setColumnWidth(column,75);
      column++;
    }
  }
  var finalMatchColumn = (column - 1)
  sheet.getRange(2,1).setValue((finalMatchColumn-4) + ' NFL GAMES');
  sheet.getRange(1,column).setValue('TIEBREAKER');
  validRule = SpreadsheetApp.newDataValidation().requireNumberBetween(0,150)
  .setHelpText('Must be a number')
  .build();
  sheet.getRange(2,column).setDataValidation(validRule);
  sheet.setColumnWidth(column,100);
  column++;
  sheet.getRange(1,column).setValue('DIFFERENCE');
  sheet.setColumnWidth(column,100);
  column++;
  sheet.getRange(1,column).setValue('WIN');
  sheet.setColumnWidth(column,50);
  column++;
  sheet.getRange(1,column).setValue('MNF');
  sheet.setColumnWidth(column,50);
  column++;
  sheet.getRange(1,column).setValue('COMMENT'); // Added to allow submissions to have amusing comments, if desired
  sheet.setColumnWidth(column,125);
  column = column + diffCount;
  
  // Headers completed, now adjusting number of columns once headers are populated
  maxCols = sheet.getMaxColumns();
  
  sheet.getRange(1,column - diffCount + 1,2,1).setValues([['MOST SIMILAR'],['\[\# DIFFERENT\]']]); // Added to allow submissions to have amusing comments, if desired
  if (column > maxCols) {
    sheet.insertColumnsAfter(maxCols,column-maxCols);
  }
  sheet.setColumnWidths((column - diffCount),diffCount,140);
  
  if (maxCols > column){
    sheet.deleteColumns(column,maxCols - column + 1);
  }  
  maxCols = sheet.getMaxColumns();
  
  // Declare NFL Winners range for the week
  ss.setNamedRange('NFL_'+year+'_'+week,sheet.getRange(2,5,1,finalMatchColumn-4));
  
  for ( j = 3; j < rows; j++ ) {
    if ( j % 2 == 0 ) {
      sheet.getRange('R'+j+'C1:R'+j+'C'+(maxCols-diffCount)).setBackground('#f0f0f0');
    }
    // Formula to determine how many correct on the week
    sheet.getRange(j,2).setFormulaR1C1('=iferror(if(and(counta(R2C[3]:R2C['+(finalMatchColumn-2)+'])>0,counta(R[0]C[3]:R[0]C['+(finalMatchColumn-2)+'])>0),mmult(arrayformula(if(R2C[3]:R2C['+(finalMatchColumn-2)+']=R[0]C[3]:R[0]C['+(finalMatchColumn-2)+'],1,0)),transpose(arrayformula(if(not(isblank(R[0]C[3]:R[0]C['+(finalMatchColumn-2)+'])),1,0)))),))');
    // Formula to determine weekly rank
    sheet.getRange(j,3).setFormulaR1C1('=iferror(if(and(counta(R2C[2]:R2C['+(finalMatchColumn-3)+'])>0,not(isblank(R[0]C[-1]))),rank(R[0]C[-1],R3C2:R'+(totalMembers+2)+'C2,false),))');
    // Formula to determine weekly correct percent
    sheet.getRange(j,4).setFormulaR1C1('=iferror(if(and(counta(R2C[1]:R2C['+(finalMatchColumn-4)+'])>0,not(isblank(R[0]C[-2]))),R'+j+'C[-2]/counta(R2C[1]:R2C['+(finalMatchColumn-4)+']),),)');
    if ( mnf == true ) {
      // Formula to determine difference of tiebreaker from final MNF score
      sheet.getRange(j,finalMatchColumn+2).setFormulaR1C1('=iferror(if(or(isblank(R[0]C[-1]),isblank(R2C'+(finalMatchColumn+1)+')),,abs(R[0]C[-1]-R2C'+(finalMatchColumn+1)+')))');
    }
    // Formula to denote winner with a '1'
    sheet.getRange(j,finalMatchColumn+3).setFormulaR1C1('=iferror(if(sum(arrayformula(if(isblank(R2C5:R2C'+(finalMatchColumn+1)+'),1,0)))>0,,match(R[0]C1,filter(filter(R3C1:R'+(totalMembers+2)+'C1,R3C2:R'+(totalMembers+2)+'C2=max(R3C2:R'+(totalMembers+2)+'C2)),filter(R3C[-1]:R'+(totalMembers+2)+'C[-1],R3C2:R'+(totalMembers+2)+'C2=max(R3C2:R'+(totalMembers+2)+'C2))=min(filter(R3C[-1]:R'+(totalMembers+2)+'C[-1],R3C2:R'+(totalMembers+2)+'C2=max(R3C2:R'+(totalMembers+2)+'C2)))),0)^0))');
    // Formula to determine MNF win status sum (can be more than 1 for rare weeks)
    if ( mnf == true ) {
      sheet.getRange(j,finalMatchColumn+4).setFormulaR1C1('=iferror(if(mmult(arrayformula(if(R2C'+mnfStart+':R2C'+mnfEnd+'=R[0]C'+mnfStart+':R[0]C'+mnfEnd+',1,0)),transpose(arrayformula(if(not(isblank(R[0]C'+mnfStart+':R[0]C'+mnfEnd+')),1,0))))=0,0,mmult(arrayformula(if(R2C'+mnfStart+':R2C'+mnfEnd+'=R[0]C'+mnfStart+':R[0]C'+mnfEnd+',1,0)),transpose(arrayformula(if(not(isblank(R[0]C'+mnfStart+':R[0]C'+mnfEnd+')),1,0))))))');
    }
    // Formula to generate array of similar pickers on the week
    sheet.getRange(j,finalMatchColumn+6).setFormulaR1C1('=iferror(if(isblank(R[0]C5),,transpose(arrayformula({query({R3C1:R'+(totalMembers+2)+'C1,arrayformula(mmult(if(R3C5:R'+(totalMembers+2)+'C'+(finalMatchColumn)+'=R[0]C5:R[0]C'+(finalMatchColumn)+',1,0),transpose(arrayformula(column(R[0]C5:R[0]C'+(finalMatchColumn)+')\^0))))},\"select Col1 where Col1 \<\> \'\"\&R[0]C1\&\"\' order by Col2 desc, Col1 asc limit '+diffCount+
      '\")\&\" [\"\&arrayformula('+(finalMatchColumn-2)+'-query({R3C1:R'+(totalMembers+2)+'C1,arrayformula(mmult(if(R3C5:R'+(totalMembers+2)+'C'+(finalMatchColumn)+'=R[0]C5:R[0]C'+(finalMatchColumn)+',1,0),transpose(arrayformula(column(R[0]C5:R[0]C'+(finalMatchColumn)+')\^0))))},\"select Col2 where Col1 <> \'\"\&R[0]C1\&\"\' order by Col2 desc, Col1 asc limit '+diffCount+'\"))-2\&\"]\"}))))');
  }
  sheet.getRange(rows,1,1,maxCols).setBackground('#dbdbdb');
  sheet.getRange(rows,2).setBackground('#fffee3');
  sheet.getRange(rows,3).setBackground('#e3fffe'); 
  for ( j = 5; j <= finalMatchColumn; j++ ) {
    if (totalMembers >= 3) { // adjusts an if statement conditional for varying amounts of members
      cellsPopulatedCheck = 'or(not(isblank(R3C[0])),not(isblank(R4C[0])),not(isblank(R5C[0])))';
    } else if (totalMembers == 2){
      cellsPopulatedCheck = 'or(not(isblank(R3C[0])),not(isblank(R4C[0])))';
    } else if (totalMembers == 1) {
      cellsPopulatedCheck = 'not(isblank(R3C[0]))';
    }
    sheet.getRange(rows,j).setFormulaR1C1('=iferror(if(counta(R3C[0]:R[-1]C[0])>0,if(countif(R3C[0]:R'+(totalMembers+2)+'C[0],regexextract(R1C[0],"[A-Z]{2,3}"))=counta(R3C[0]:R'+(totalMembers+2)+'C[0])/2,"SPLIT",if(countif(R3C[0]:R'+(totalMembers+2)+'C[0],regexextract(R1C[0],"[A-Z]{2,3}"))<counta(R3C[0]:R'+(totalMembers+2)+'C[0])/2,regexextract(right(R1C[0],3),"[A-Z]{2,3}")&"|"&round(100*countif(R3C[0]:R'+(totalMembers+2)+'C[0],regexextract(right(R1C[0],3),"[A-Z]{2,3}"))/counta(R3C[0]:R'+(totalMembers+2)+'C[0]),1)&"%",regexextract(R1C[0],"[A-Z]{2,3}")&"|"&round(100*countif(R3C[0]:R'+(totalMembers+2)+'C[0],regexextract(R1C[0],"[A-Z]{2,3}"))/counta(R3C[0]:R'+(totalMembers+2)+'C[0]),1)&"%")),))');
  }
  sheet.getRange(rows,j+5).setFormulaR1C1('=iferror(if(isblank(R[0]C5),,transpose(query({arrayformula(R3C1:R'+(totalMembers+2)+'C1&\" [\"&(counta(R1C5:R1C'+finalMatchColumn+')-mmult(arrayformula(if(R3C5:R'+(totalMembers+2)+'C'+finalMatchColumn+'=arrayformula(regexextract(R'+(totalMembers+3)+'C5:R'+(totalMembers+3)+'C'+finalMatchColumn+',\"[A-Z]+\")),1,0)),transpose(arrayformula(if(arrayformula(len(R1C5:R1C'+finalMatchColumn+'))>1,1,1)))))&\"]\"),mmult(arrayformula(if(R3C5:R'+(totalMembers+2)+'C'+finalMatchColumn+'=arrayformula(regexextract(R'+(totalMembers+3)+'C5:R'+(totalMembers+3)+'C'+finalMatchColumn+',\"[A-Z]+\")),1,0)),transpose(arrayformula(if(arrayformula(len(R1C5:R1C'+finalMatchColumn+'))>1,1,1))))},\"select Col1 order by Col2 desc, Col1 desc limit '+diffCount+'\"))))');
  
 // AWAY TEAM BIAS FORMULA 
  sheet.getRange(rows,2,1,1).setFormulaR1C1('=iferror(if(counta(R3C5:R'+(totalMembers+2)+'C'+finalMatchColumn+')>10,"AWAY|"&round(100*(sum(arrayformula(if(regexextract(R1C5:R1C'+finalMatchColumn+',"^[A-Z]{2,3}")=R1C5:R'+(totalMembers+2)+'C'+finalMatchColumn+',1,0)))/counta(R3C5:R'+(totalMembers+2)+'C'+finalMatchColumn+')),1)&"%","AWAY"),"AWAY")');
  // HOME TEAM BIAS FORMULA
  sheet.getRange(rows,3,1,1).setFormulaR1C1('=iferror(if(counta(R3C5:R'+(totalMembers+2)+'C'+finalMatchColumn+')>10,"HOME|"&round(100*(sum(arrayformula(if(regexextract(R1C5:R1C'+finalMatchColumn+',"[A-Z]{2,3}$")=R1C5:R'+(totalMembers+2)+'C'+finalMatchColumn+',1,0)))/counta(R3C5:R'+(totalMembers+2)+'C'+finalMatchColumn+')),1)&"%","HOME"),"HOME")');
  sheet.getRange(rows,4,1,1).setFormulaR1C1('=iferror(if(counta(R2C[1]:R2C['+(finalMatchColumn-4)+'])>2,average(R2C[0]:R'+(totalMembers+2)+'C[0]),))');


  // Setting conditional formatting rules
  sheet.clearConditionalFormatRules();    
  var range = sheet.getRange('R3C5:R'+(rows-1)+'C'+finalMatchColumn)
  // CORRECT PICK COLOR RULE
  var formatRuleCorrectEven = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R2C[0]\",false)=indirect(\"R[0]C[0]\",false),not(isblank(indirect(\"R2C[0]\",false))),iseven(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#c9ffdf')
    .setRanges([range])
    .build();
  var formatRuleCorrectOdd = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R2C[0]\",false)=indirect(\"R[0]C[0]\",false),not(isblank(indirect(\"R2C[0]\",false))),isodd(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#bbedd0')
    .setRanges([range])
    .build();
  // INCORRECT PICK COLOR RULE
  var formatRuleIncorrectEven = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(indirect(\"R2C[0]\",false)=indirect(\"R[0]C[0]\",false)),not(isblank(indirect(\"R2C[0]\",false))),iseven(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#ffc4ca')
    .setStrikethrough(true)
    .setRanges([range])
    .build();
  var formatRuleIncorrectOdd = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(indirect(\"R2C[0]\",false)=indirect(\"R[0]C[0]\",false)),not(isblank(indirect(\"R2C[0]\",false))),isodd(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#f2bdc2')
    .setStrikethrough(true)
    .setRanges([range])
    .build();
  // HOME PICK COLOR RULE
  var formatRuleHomeEven = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),split(indirect(\"R1C[0]\",false),\"\@\"),0)=2,iseven(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#e3fffe')
    .setRanges([range])
    .build();
  var formatRuleHomeOdd = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),split(indirect(\"R1C[0]\",false),\"\@\"),0)=2,isodd(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#d0f5f3')
    .setRanges([range])
    .build();
  // AWAY PICK COLOR RULE
  var formatRuleAwayEven = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),split(indirect(\"R1C[0]\",false),\"\@\"),0)=1,iseven(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#fffee3')
    .setRanges([range])
    .build();
  var formatRuleAwayOdd = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),split(indirect(\"R1C[0]\",false),\"\@\"),0)=1,isodd(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#faf9e1')
    .setRanges([range])
    .build();  
  
  // NAMES COLUMN NAMED RANGE
  range = sheet.getRange('R3C1:R'+(rows-1)+'C1');
  ss.setNamedRange('NAMES_'+year+'_'+week,range);

  // TOTALS GRADIENT RULE
  range = sheet.getRange('R3C2:R'+(rows-1)+'C2');
  ss.setNamedRange('TOT_'+year+'_'+week,range);
  var formatRuleTotals = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint("#75F0A1")
    .setGradientMinpoint("#FFFFFF")
    //.setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, (finalMatchColumn-2) - 3) // Max value of all correct picks (adjusted by 3 to tighten color range)
    //.setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, (finalMatchColumn-2) / 2)  // Generates Median Value
    //.setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, 0 + 3) // Min value of all correct picks (adjusted by 3 to tighten color range)
    .setRanges([range])
    .build();
  // RANKS GRADIENT RULE
  range = sheet.getRange('R3C3:R'+(rows-1)+'C3');
  ss.setNamedRange('RANK_'+year+'_'+week,range);
  var formatRuleRanks = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, members.length)
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, members.length/2)
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([range])
    .build();
  // PERCENT GRADIENT RULE
  range = sheet.getRange('R3C4:R'+(rows)+'C4')
  range.setNumberFormat('##0.0%');
  var formatRulePercent = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, ".70")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, ".60")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, ".50")
    .setRanges([range])
    .build();
  ss.setNamedRange('PCT_'+year+'_'+week,sheet.getRange('R3C4:R'+(rows-1)+'C4'));    
  // WINNER COLUMN RULE
  range = sheet.getRange('R3C'+(finalMatchColumn+3)+':R'+(rows-1)+'C'+(finalMatchColumn+3));
  ss.setNamedRange('WIN_'+year+'_'+week,range);
  var formatRuleNotWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotEqualTo(1)
    .setBackground('#FFFFFF')
    .setFontColor('#FFFFFF')
    .setRanges([range])
    .build();     
  var formatRuleWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#75F0A1')
    .setFontColor('#75F0A1')
    .setRanges([range])
    .build();
  // WINNER NAME RULE
  range = sheet.getRange('R3C1:R'+rows+'C1');
  var formatRuleWinnerName = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=indirect(\"R[0]C'+(finalMatchColumn+3)+'\",false)=1')
    .setBackground('#75F0A1')
    .setRanges([range])
    .build();  
  // MNF GRADIENT RULE
  range = sheet.getRange('R3C'+(finalMatchColumn+4)+':R'+(rows-1)+'C'+(finalMatchColumn+4));
  if (mnf == true) {
    ss.setNamedRange('MNF_'+year+'_'+week,range);
  }
  var formatRuleMNFEmpty = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=isblank(indirect("R[0]C[0]"))')
    .setFontColor('#FFFFFF')
    .setBackground('#FFFFFF')
    .setRanges([range])
    .build();
  var formatRuleMNFZero = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setFontColor('#FFFFFF')
    .setBackground('#FFFFFF')
    .setRanges([range])
    .build();    
  var formatRuleMNF = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint("#C2FF7D") // Max value of all correct picks, min 1
    .setGradientMinpoint("#FFFFFF") // Min value of all correct picks  
    .setRanges([range])
    .build();
  range = sheet.getRange(3,column-diffCount+1,totalMembers+1,diffCount);
  var formatCommonPicker0 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))=0')
    .setBackground('#46f081')
    .setRanges([range])
    .build();
  var formatCommonPicker1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))=1')
    .setBackground('#75F0A1')
    .setRanges([range])
    .build();
  var formatCommonPicker2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))=2')
    .setBackground('#a4edbe')
    .setRanges([range])
    .build();
  var formatCommonPicker3 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))=3')
    .setBackground('#e4f0e8')
    .setRanges([range])
    .build();
  // DIFFERENCE TIEBREAKER COLUMN FORMATTING
  range = sheet.getRange(3,finalMatchColumn+2,totalMembers,1);
  var formatRuleDiff = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint("#FFFFFF")
    .setGradientMinpoint("#5EDCFF")
    .setRanges([range])
    .build();
  // PREFERENCE COLOR SCHEMES
  range = sheet.getRange(rows,4,1,finalMatchColumn-3);
  // Away Favored 90%
  var formatRuleAway90 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R1C[0]\",false),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=90)')
    .setBackground('#fffb7d')
    .setRanges([range])
    .build();
  // Home Favored 90%
  var formatRuleHome90 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R1C[0]\",false),3),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=90)')
    .setBackground('#7dfffb')
    .setRanges([range])
    .build();
  // Away favored 80%
  var formatRuleAway80 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R1C[0]\",false),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=80)')
    .setBackground('#fffc96')
    .setRanges([range])
    .build();
  // Home Favored 80%
  var formatRuleHome80 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R1C[0]\",false),3),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=80)')
    .setBackground('#96fffc')
    .setRanges([range])
    .build();
  // Away Favored 70%
  var formatRuleAway70 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R1C[0]\",false),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=70)')
    .setBackground('#fffcb0')
    .setRanges([range])
    .build();
  // Home Favored 70%
  var formatRuleHome70 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R1C[0]\",false),3),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=70)')
    .setBackground('#b0fffc')
    .setRanges([range])
    .build();
  // Away Favored 60%
  var formatRuleAway60 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R1C[0]\",false),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=60)')
    .setBackground('#fffdc9')
    .setRanges([range])
    .build();
  // Home Favored 60%
  var formatRuleHome60 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R1C[0]\",false),3),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=60)')
    .setBackground('#c9fffd')
    .setRanges([range])
    .build();
  // Away Favored
  var formatRuleAway50 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R1C[0]\",false),\"[A-Z]{2,3}\")')
    .setBackground('#fffee3')
    .setRanges([range])
    .build();
  // Home Favored
  var formatRuleHome50 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R1C[0]\",false),3),\"[A-Z]{2,3}\")')
    .setBackground('#e3fffe')
    .setRanges([range])
    .build();
  var formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleCorrectEven);
  formatRules.push(formatRuleCorrectOdd);
  formatRules.push(formatRuleIncorrectEven);
  formatRules.push(formatRuleIncorrectOdd);
  formatRules.push(formatRuleHomeEven);
  formatRules.push(formatRuleHomeOdd);
  formatRules.push(formatRuleAwayEven);
  formatRules.push(formatRuleAwayOdd);
  formatRules.push(formatRuleTotals);
  formatRules.push(formatRuleRanks);
  formatRules.push(formatRulePercent);
  formatRules.push(formatRuleNotWinner);
  formatRules.push(formatRuleWinner);
  formatRules.push(formatRuleWinnerName);
  formatRules.push(formatRuleMNFEmpty);
  formatRules.push(formatRuleMNFZero);
  formatRules.push(formatRuleMNF);
  formatRules.push(formatRuleDiff);
  formatRules.push(formatRuleHome90);
  formatRules.push(formatRuleAway90);
  formatRules.push(formatRuleHome80);
  formatRules.push(formatRuleAway80);
  formatRules.push(formatRuleHome70);
  formatRules.push(formatRuleAway70);
  formatRules.push(formatRuleHome60);
  formatRules.push(formatRuleAway60);
  formatRules.push(formatRuleHome50);
  formatRules.push(formatRuleAway50);    
  formatRules.push(formatCommonPicker0);
  formatRules.push(formatCommonPicker1);
  formatRules.push(formatCommonPicker2);
  formatRules.push(formatCommonPicker3);
  sheet.setConditionalFormatRules(formatRules);
  
  // Setting size, alignment, frozen columns
  columns = sheet.getMaxColumns();
  range = sheet.getRange(1,1,rows,columns)
  range.setFontSize(10);
  range.setVerticalAlignment('middle');
  range.setHorizontalAlignment('center');
  range.setFontFamily("Montserrat");
  sheet.getRange(3,column-diffCount,totalMembers+1,diffCount+1).setHorizontalAlignment('left');
  range = sheet.getRange(1,1,rows,1);
  range.setHorizontalAlignment('left');
  sheet.setFrozenColumns(4);
  sheet.setFrozenRows(2);
  range = sheet.getRange(1,1,2,columns);
  range.setBackground('black');
  range.setFontColor('white');

  if (dataRestore == true && fresh == false) {
    if (tiebreakerCol  >= 0) {
      previousDataRange.setValues(previousData);
      ss.toast('Previous values restored for week ' + week);
    } else {
      Logger.log('ERROR: Previous data not transferred! Undo immediately');
      ss.toast('ERROR: Previous data not transferred! Undo immediately');
    }
    if (commentCol  >= 0) {
      previousCommentRange.setValues(previousComment);
    }
  }

  return sheet;
}

//------------------------------------------------------------------------
// OVERALL Sheet Creation / Adjustment
function overallSheet(year,weeks,members) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'OVERALL';
  var sheet = ss.getSheetByName(sheetName)
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.clear();
  var totalMembers = members.length;
  
  var rows = totalMembers + 2;
  var maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  if ( weeks + 2 < maxCols ) {
    sheet.deleteColumns(weeks + 2,maxCols-(weeks + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('CORRECT');
  sheet.getRange(1,2).setValue('TOTAL');
  sheet.getRange(rows,1).setValue('AVERAGES');

  var mask;
  for ( var i = 0; i < weeks; i++ ) {
    sheet.getRange(1,i+3).setValue(i+1);
    sheet.setColumnWidth(i+3,30);
    if (i+1 < 10 ) { 
      mask = '0' + (i+1);
    } else {
      mask = (i+1);
    }
    sheet.getRange(2,i+3).setFormula('=iferror(arrayformula(TOT_'+year+'_'+mask+'))')
  }
  
  var range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(members); 
  sheet.getRange(1,1,totalMembers+2,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);
  
  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  sheet.getRange(totalMembers+2,1,1,weeks+2).setBackground('#e6e6e6');
  
  sheet.getRange(2,2,totalMembers+1,weeks+1).setNumberFormat('#.#');

  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1); 

  // SET OVERALL NAMES Range
  var rangeOverallTotNames = sheet.getRange('R2C1:R'+rows+'C1');
  ss.setNamedRange('TOT_OVERALL_'+year+'_NAMES',rangeOverallTotNames); 
  sheet.clearConditionalFormatRules(); 
  // OVERALL TOTAL GRADIENT RULE
  var rangeOverallTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('TOT_OVERALL_'+year,rangeOverallTot);
  var valuesOverallTot = rangeOverallTot.getValues().flat();
  var formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("TOT_OVERALL_'+year+'"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("TOT_OVERALL_'+year+'"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("TOT_OVERALL_'+year+'"))') // Min value of all correct picks
    .setRanges([rangeOverallTot])
    .build();
  // OVERALL SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks+2));
  ss.setNamedRange('TOT_WEEKLY_'+year,range);
  var formatRuleOverallHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R[0]C[0]\",false)>0,indirect(\"R[0]C[0]\",false)=max(indirect(\"R2C[0]:R'+maxRows+'C[0]\",false)))')
    .setBackground('#75F0A1')
    .setBold(true)
    .setRanges([range])
    .build();
  var formatRuleOverall = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, "15")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "10")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "5")
    .setRanges([range])
    .build();
  var formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallHigh);
  formatRules.push(formatRuleOverall);
  formatRules.push(formatRuleOverallTot);
  sheet.setConditionalFormatRules(formatRules);
  
  overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',true);
  overallMainFormulas(sheet,totalMembers,weeks,year,'TOT',true);
  
  return sheet;  
}

//------------------------------------------------------------------------
// OVERALL RANK Sheet Creation / Adjustment
function overallRankSheet(year,weeks,members) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'OVERALL_RANK';
  var sheet = ss.getSheetByName(sheetName)
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.clear();
  if (members == null) {
    members = memberList();
  }

  var totalMembers = members.length;
  
  var rows = totalMembers + 1;
  var maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  if ( weeks + 2 < maxCols ) {
    sheet.deleteColumns(weeks + 2,maxCols-(weeks + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('RANKS');
  sheet.getRange(1,2).setValue('AVERAGE');

  var mask;
  for ( var i = 0; i < weeks; i++ ) {
    sheet.getRange(1,i+3).setValue(i+1);
    sheet.setColumnWidth(i+3,30);
    if (i+1 < 10 ) { 
      mask = '0' + (i+1);
    } else {
      mask = (i+1);
    }
    sheet.getRange(2,i+3).setFormula('=iferror(arrayformula(RANK_'+year+'_'+mask+'))')
  }
  
  var range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(members); 
  sheet.getRange(1,1,totalMembers+1,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);
  
  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);

  // SET OVERALL RANK NAMES Range
  var rangeOverallTotRnkNames = sheet.getRange('R2C1:R'+rows+'C1');
  ss.setNamedRange('TOT_OVERALL_RANK_'+year+'_NAMES',rangeOverallTotRnkNames);  
  sheet.clearConditionalFormatRules(); 
  // RANKS TOTAL GRADIENT RULE
  var rangeOverallRankTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('TOT_OVERALL_RANK_'+year,rangeOverallRankTot);
  var formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([rangeOverallRankTot])
    .build();
  // RANKS SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks+2));
  ss.setNamedRange('TOT_WEEKLY_RANK_'+year,range);
  var formatRuleOverallWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#00E1FF')
    .setBold(true)
    .setRanges([range])
    .build();
  var formatRuleOverall = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([range])
    .build();
  var formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallWinner);
  formatRules.push(formatRuleOverall);
  formatRules.push(formatRuleOverallTot);
  sheet.setConditionalFormatRules(formatRules);
  
  overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',false);
  overallMainFormulas(sheet,totalMembers,weeks,year,'RANK',false);
  
  
  return sheet;  
}

//------------------------------------------------------------------------
// OVERALL PERCENT Sheet Creation / Adjustment
function overallPctSheet(year,weeks,members) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'OVERALL_PCT';
  var sheet = ss.getSheetByName(sheetName)
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.clear();
  
  if (members == null) {
    members = memberList();
  }
  var totalMembers = members.length;
  
  var rows = totalMembers + 2; // 2 additional rows
  var maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  if ( weeks + 2 < maxCols ) {
    sheet.deleteColumns(weeks + 2,maxCols-(weeks + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('PERCENTS');
  sheet.getRange(1,2).setValue('AVERAGE');
  sheet.getRange(totalMembers + 2,1).setValue('AVERAGES');

  var mask;
  for ( var a = 0; a < weeks; a++ ) {
    sheet.getRange(1,a+3).setValue(a+1);
    sheet.setColumnWidth(a+3,48);
  }
  
  var range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(members); 
  sheet.getRange(1,1,totalMembers+2,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);
  
  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  sheet.getRange(totalMembers+2,1,1,weeks+2).setBackground('#e6e6e6'); 

  sheet.getRange(2,2,totalMembers+1,1).setNumberFormat("##.#%");  
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);

  // SET OVERALL PCT NAMES Range
  var rangeOverallTotPctNames = sheet.getRange('R2C1:R'+(rows-1)+'C1');
  ss.setNamedRange('TOT_OVERALL_PCT_'+year+'_NAMES',rangeOverallTotPctNames);
  sheet.clearConditionalFormatRules(); 
  // OVERALL PCT TOTAL GRADIENT RULE
  var rangeOverallTotPct = sheet.getRange('R2C2:R'+(rows-1)+'C2');
  ss.setNamedRange('TOT_OVERALL_PCT_'+year,rangeOverallTotPct);
  rangeOverallTotPct = sheet.getRange('R2C2:R'+rows+'C2');
  var valuesOverallTot = rangeOverallTotPct.getValues().flat();
  var formatRuleOverallPctTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("TOT_OVERALL_PCT_'+year+'"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("TOT_OVERALL_PCT_'+year+'"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("TOT_OVERALL_PCT_'+year+'"))') // Min value of all correct picks  
    .setRanges([rangeOverallTotPct])
    .build();  
  // OVERALL PCT SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+(rows-1)+'C'+(weeks+2));
  ss.setNamedRange('TOT_WEEKLY_PCT_'+year,range);
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks+2)); 
  var formatRuleOverallPctHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R[0]C[0]\",false)>0,indirect(\"R[0]C[0]\",false)=max(indirect(\"R2C[0]:R'+maxRows+'C[0]\",false)))')
    .setBackground('#75F0A1')
    .setBold(true)
    .setRanges([range])
    .build();
  var formatRuleOverallPct = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "0.5")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setRanges([range])
    .build();
  var formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallPctHigh);
  formatRules.push(formatRuleOverallPct);
  formatRules.push(formatRuleOverallPctTot);
  sheet.setConditionalFormatRules(formatRules);

  overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',true);
  overallMainFormulas(sheet,totalMembers,weeks,year,'PCT',true);

  return sheet;  
}

//------------------------------------------------------------------------
// MNF Sheet Creation / Adjustment
function mnfSheet(year,weeks,members) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'MNF';
  var sheet = ss.getSheetByName(sheetName)
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.clear();

  if (members == null) {
    members = memberList();
  }
  var totalMembers = members.length;
  
  var rows = totalMembers + 1;
  var maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  if ( weeks + 2 < maxCols ) {
    sheet.deleteColumns(weeks + 2,maxCols-(weeks + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('CORRECT');
  sheet.getRange(1,2).setValue('TOTAL');

  var mask;
  for ( var a = 0; a < weeks; a++ ) {
    sheet.getRange(1,a+3).setValue(a+1);
    sheet.setColumnWidth(a+3,30);
  }
  
  var range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(members); 
  sheet.getRange(1,1,totalMembers+1,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);
  
  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1); 

  sheet.clearConditionalFormatRules(); 
  
  // CORRECT MNF COLOR RULE
  var formatRuleCorrectEven = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R2C[0]\",false)=indirect(\"R[0]C[0]\",false),not(isblank(indirect(\"R2C[0]\",false))),iseven(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#c9ffdf')
    .setRanges([range])
    .build();
  // INCORRECT MNF COLOR RULE
  var formatRuleIncorrectEven = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(indirect(\"R2C[0]\",false)=indirect(\"R[0]C[0]\",false)),not(isblank(indirect(\"R2C[0]\",false))),iseven(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#ffc4ca')
    .setStrikethrough(true)
    .setRanges([range])
    .build();

  // SET OVERALL NAMES Range
  var rangeMnfNames = sheet.getRange('R2C1:R'+rows+'C1');
  ss.setNamedRange('MNF_'+year+'_NAMES',rangeMnfNames); 
  // OVERALL TOTAL GRADIENT RULE
  var rangeMnfTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('MNF_'+year,rangeMnfTot);
  var valuesMnfTot = rangeMnfTot.getValues().flat();
  var formatRuleMnfTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#C9FFDF", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("MNF_'+year+'"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("MNF_'+year+'"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("MNF_'+year+'"))') // Min value of all correct picks
    .setRanges([rangeMnfTot])
    .build();
  // OVERALL SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks+2));
  ss.setNamedRange('MNF_WEEKLY_'+year,range);
  var formatRuleCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#C9FFDF')
    .setFontColor('#C9FFDF')
    .setBold(true)
    .setRanges([range])
    .build();
  var formatRuleIncorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#FFC4CA')
    .setFontColor('#FFC4CA')
    .setBold(true)
    .setRanges([range])
    .build();    
  var formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleCorrect);
  formatRules.push(formatRuleIncorrect);
  formatRules.push(formatRuleMnfTot);
  sheet.setConditionalFormatRules(formatRules);

  overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',false)
  overallMainFormulas(sheet,totalMembers,weeks,year,'MNF',false);

  return sheet;  
}

// OVERALL / OVERALL RANK / OVERALL PCT / MNF Combination formula for sum/average per player row
function overallPrimaryFormulas(sheet,totalMembers,maxCols,action,avgRow) {
  for ( var a = 1; a < totalMembers; a++ ) {
    if (action == 'average') {
      sheet.getRange(2,2,a+1,1).setFormulaR1C1('=iferror(if(counta(R[0]C3:R[0]C'+maxCols+')=0,,average(R[0]C3:R[0]C'+maxCols+')))');
    } else if (action == 'sum') {
      sheet.getRange(2,2,a+1,1).setFormulaR1C1('=iferror(if(counta(R[0]C3:R[0]C'+maxCols+')=0,,sum(R[0]C3:R[0]C'+maxCols+')))');
    }
    if (sheet.getSheetName() == 'OVERALL_PCT') {
      sheet.getRange(2,2,a+1,1).setNumberFormat("##.#%");
    } else if (action == 'sum') {
      sheet.getRange(2,2,a+1,1).setNumberFormat("##");
    } else {
      sheet.getRange(2,2,a+1,1).setNumberFormat("#0.0");
    }
  }
  if (avgRow == true){
    var rows = sheet.getMaxRows()
    sheet.getRange(rows,2).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))');
  }  
}

// OVERALL / OVERALL RANK / OVERALL PCT / MNF Combination formula for each column (week)
function overallMainFormulas(sheet,totalMembers,weeks,year,str,avgRow) {
  var b;
  for ( var a = 1; a <= weeks; a++ ) {
    b = 1;
    for (b ; b <= totalMembers; b++) {
      sheet.getRange(b+1,a+2).setFormula('=iferror(arrayformula(vlookup(R[0]C1,{NAMES_'+year+'_'+a+','+str+'_'+year+'_'+a+'},2,false)))');
      if (sheet.getSheetName() == 'OVERALL_PCT') {
        sheet.getRange(b+1,a+2).setNumberFormat("##.#%");
      } else {
        sheet.getRange(b+1,a+2).setNumberFormat("#0");
      }
    }
  }
  if (avgRow == true){
    for (a = 0; a < weeks; a++){
      var rows = sheet.getMaxRows()
      sheet.getRange(rows,a+3).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))');
    }
  }
}

// WEEKLY WINNERS Combination formula update
function winnersFormulas(sheet,weeks,year) {
  for ( var a = 1; a <= weeks; a++ ) {
    winRange = 'WIN_' + year + '_' + a;
    nameRange = 'NAMES_' + year + '_' + a;
    sheet.getRange(a+1,2).setFormulaR1C1('=iferror(join(", ",sort(filter('+nameRange+','+winRange+'=1),1,true)))');
  }
}

// REFRESH FORMULAS FOR OVERALL / OVERALL RANK / OVERALL PCT / MNF
function allFormulasUpdate(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var members = memberList();
  var weeks = fetchTotalWeeks();
  var year = fetchYear();

  var sheet = ss.getSheetByName('OVERALL');
  var maxCols = sheet.getMaxColumns();
  var totalMembers = members.length;
  overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',true);
  overallMainFormulas(sheet,totalMembers,weeks,year,'TOT',true);

  sheet = ss.getSheetByName('OVERALL_RANK');
  maxCols = sheet.getMaxColumns();
  overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',false);
  overallMainFormulas(sheet,totalMembers,weeks,year,'RANK',false);

  sheet = ss.getSheetByName('OVERALL_PCT');
  maxCols = sheet.getMaxColumns();
  overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',true);
  overallMainFormulas(sheet,totalMembers,weeks,year,'PCT',true);

  sheet = ss.getSheetByName('MNF');
  maxCols = sheet.getMaxColumns();
  overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',false)
  overallMainFormulas(sheet,totalMembers,weeks,year,'MNF',false);

  sheet = ss.getSheetByName('WINNERS');
  winnersFormulas(sheet,weeks,year);
}

//------------------------------------------------------------------------
// SURVIVOR Sheet Creation / Adjustment
function survivorSheet(year,weeks,members,dataRestore) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'SURVIVOR';
  var sheet = ss.getSheetByName(sheetName)
  var fresh = false;
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
    fresh = true;
  }

  var members = memberList();
  var totalMembers = members.length;
  var rows = members.length + 3; // Accounting for the top two rows above member rows
  var columns;

  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();

  if (dataRestore == true && fresh == false){
    var previousDataRange = sheet.getRange(2,3,maxRows-2,maxCols-2);
    var previousData = previousDataRange.getValues();
    ss.toast('Gathered previous data for SURVIVOR sheet, recreating sheet now');
  }
  sheet.clear();

  if (members == null) {
    members = memberList();
  }
  var totalMembers = members.length;
    
  var rows = totalMembers + 2;
  var maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();  
  var maxCols = sheet.getMaxColumns();
  if ( (weeks + 2) < maxCols ) {
    sheet.deleteColumns(weeks+2,maxCols - (weeks+2));
  }
  
  sheet.getRange(1,1).setValue('PLAYER');
  var eliminatedCol = 2;
  sheet.getRange(1,eliminatedCol).setValue('ELIMINATED');
  sheet.setColumnWidth(eliminatedCol,100);
  
  for ( var a = 0; a < weeks; a++ ) {
    sheet.getRange(1,a+3).setValue(a+1);
    sheet.setColumnWidth(a+3,30);
  }
  
  var formula;
  var c;
  for ( var b = 2; b <= totalMembers; b++ ) {
    formula = '=iferror(match(true,arrayformula\(\{';
    for ( c = 0; c < weeks; c++ ) {
      formula = formula.concat('if(or(R1C'+(c+3)+'>=WEEK,isblank(R[0]C'+(c+3)+')),false,iserror(match(R[0]C'+(c+3)+',indirect(\"NFL_'+year+'_\"\&R1C'+(c+3)+'),0)))');
      if ( c < (weeks-1) ) {
        formula = formula.concat(',');
      }
    }
    formula = formula.concat('\}\),0),)');  
    sheet.getRange(2,eliminatedCol,b,1).setFormulaR1C1(formula);
  }
  for ( b = 1; b < weeks; b++ ) {
    formula = '=iferror(if(sum(arrayformula(if(isblank(R2C[0]:R[-1]C[0]),0,1)))>0,counta(R2C1:R[-1]C1)-countif(R2C2:R[-1]C2,\"\<=\"\&R1C[0]),))'
    sheet.getRange(totalMembers+2,2+b).setFormulaR1C1(formula);
  }
  
  formula = '=iferror(rows(R2C[0]:R[-1]C[0])-counta(R2C[0]:R[-1]C[0]))';
  sheet.getRange(totalMembers+2,eliminatedCol).setFormulaR1C1(formula);
  
  var range = sheet.getRange(1,1,rows,weeks+2);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(members);
  sheet.getRange(totalMembers+2,1).setValue('REMAINING');
  sheet.getRange(1,1,totalMembers+2,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  
  range = sheet.getRange(1,1,1,weeks+2);
  range.setBackground('black');
  range.setFontColor('white');
  range = sheet.getRange(totalMembers+2,1,1,weeks+2);
  range.setBackground('#e6e6e6');
  
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);
  
  // Setting conditional formatting rules
  sheet.clearConditionalFormatRules();    
  range = sheet.getRange('R2C3:R'+(totalMembers+1)+'C'+(weeks+1));
  // BLANK COLOR RULE
  var formatRuleBlank = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground('#FFFFFF')
    .setRanges([range])
    .build();
  // CORRECT PICK COLOR RULE
  var formatRuleCorrectElim = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C'+eliminatedCol+'\",false))),not(iserror(match(indirect(\"R[0]C[0]\",false),indirect(\"NFL_'+year+'_\"\&indirect(\"R1C[0]\",false)),0))),column(indirect(\"R[0]C[0]"\,false))>(indirect(\"R[0]C2\",false)+2))')
    .setBackground('#ffeded')
    .setRanges([range])
    .build();
  // CORRECT PICK COLOR RULE
  var formatRuleCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=not(iserror(match(indirect(\"R[0]C[0]\",false),indirect(\"NFL_'+year+'_\"\&indirect(\"R1C[0]\",false)),0)))')
    .setBackground('#c9ffdf')
    .setRanges([range])
    .build();
  // INCORRECT PICK COLOR RULE
  var formatRuleIncorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=if(or(indirect(\"R1C[0]\",false)>=indirect(\"WEEK\"),isblank(indirect(\"R[0]C[0]\",false))),false,iserror(match(indirect(\"R[0]C[0]\",false),indirect(\"NFL_'+year+'_\"\&indirect(\"R1C[0]\",false)),0)))')
    .setBackground('#f2bdc2')
    .setStrikethrough(true)
    .setRanges([range])
    .build();  
  // ELIMINATED COLOR RULE
  range = sheet.getRange('R2C2:R'+(totalMembers+1)+'C2');
  var formatRuleEliminatedColorScale = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#f5d5d8')
    .setGradientMinpoint('#f07883')
    .setRanges([range])
    .build();
  // ELIMINATED COLOR RULE
  range = sheet.getRange('R2C1:R'+(totalMembers+1)+'C'+(weeks+2));
  var formatRuleEliminated = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=not(isblank(indirect(\"R[0]C'+eliminatedCol+'\",false)))')
    .setBackground('#f2bdc2')
    .setRanges([range])
    .build();
  // CORRECT PICK COLOR RULE
  var formatRuleMaybeCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),indirect(\"R1C[0]\",false)=indirect(\"WEEK\"))')
    .setBackground('#fffec9')
    .setRanges([range])
    .build();
  var formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleCorrectElim);
  formatRules.push(formatRuleCorrect);
  formatRules.push(formatRuleIncorrect);
  formatRules.push(formatRuleEliminatedColorScale);
  formatRules.push(formatRuleEliminated);
  formatRules.push(formatRuleMaybeCorrect);
  formatRules.push(formatRuleBlank);
  sheet.setConditionalFormatRules(formatRules);

  range = sheet.getRange('R2C'+(eliminatedCol-1)+':R'+(totalMembers+1)+'C'+(eliminatedCol-1));
  ss.setNamedRange('ELIMINATED_'+year+'_NAMES',range);  
  range = sheet.getRange('R2C'+eliminatedCol+':R'+(totalMembers+1)+'C'+eliminatedCol);
  ss.setNamedRange('ELIMINATED_'+year,range);

  if (dataRestore == true && fresh == false) {
    previousDataRange.setValues(previousData);
    ss.toast('Previous values restored for SURVIVOR sheet');
  }

  return sheet;
}

//------------------------------------------------------------------------
// WINNERS Sheet Creation / Adjustment
function winnersSheet(year,weeks,members) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'WINNERS';
  var sheet = ss.getSheetByName(sheetName)
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.clear();
  
  if (members == null) {
    members = memberList();
  }
  var totalMembers = members.length;
  
  var rows = weeks + 4;
  var maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  if ( 3 < maxCols ) {
    sheet.deleteColumns(3,maxCols-3);
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue(year);;
  sheet.getRange(1,2).setValue('WINNER');
  sheet.getRange(1,3).setValue('PAID');

  var range = sheet.getRange(1,1,rows,maxCols);
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,2,rows-1,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,80);
  sheet.setColumnWidth(2,150);
  sheet.setColumnWidth(3,40);

  range = sheet.getRange(2,3,weeks+3,1)
  range.insertCheckboxes();
  range.setHorizontalAlignment('center');
  range = sheet.getRange(1,1,rows,2);
  range.setHorizontalAlignment('left');

  for (var a = 0; a <= weeks; a++) {
    sheet.getRange(a+2,1,1,1).setValue(a+1);
  }
  sheet.getRange(a+1,1,1,1).setValue('SURVIVOR');
  sheet.getRange(a+2,1,1,1).setValue('MNF');
  sheet.getRange(a+3,1,1,1).setValue('OVERALL');

  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  
  sheet.setFrozenRows(1); 

  range = sheet.getRange('R2C2:R'+(weeks+1)+'C2');
  ss.setNamedRange('WEEKLY_WINNERS',range);

  sheet.clearConditionalFormatRules(); 
  // OVERALL SHEET GRADIENT RULE
  var fivePlusWins = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('=countif($2:B$'+(weeks+1)+',B2)>=5')
  .setBackground('#2CFF75')
  .setRanges([range])
  .build();
  var fourPlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks+1)+',B2)=4')
    .setBackground('#72FFA3')
    .setRanges([range])
    .build();
  var threePlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks+1)+',B2)=3')
    .setBackground('#BBFFD3')
    .setRanges([range])
    .build();
  var twoPlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks+1)+',B2)=2')
    .setBackground('#D3FFE2')
    .setRanges([range])
    .build();
  var formatRules = sheet.getConditionalFormatRules();
  formatRules.push(fivePlusWins);
  formatRules.push(fourPlusWins);
  formatRules.push(threePlusWins);
  formatRules.push(twoPlusWins);
  sheet.setConditionalFormatRules(formatRules);
  
  var winRange;
  var nameRange;
  for ( var b = 1; b <= weeks; b++ ) {
    winRange = 'WIN_' + year + '_' + (b);
    nameRange = 'NAMES_' + year + '_' + (b);
    sheet.getRange(b+1,2,1,1).setFormulaR1C1('=iferror(join(", ",sort(filter('+nameRange+','+winRange+'=1),1,true)))');
  }
  
  return sheet;  
}

//------------------------------------------------------------------------
// SUMMARY Sheet Creation / Adjustment
function summarySheet(year,members) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'SUMMARY';
  var sheet = ss.getSheetByName(sheetName)
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.clear();
  
  if (members == null) {
    members = memberList();
  }

  var headers = ['PLAYER','TOTAL CORRECT','TOTAL RANK','MNF CORRECT','MNF RANK','WEEKLY AVG %','WEEKLY AVG RANK','WEEKLY WINS','SURVIVOR (WEEK OUT)','NOTES'];
  var headersWidth = [120,80,80,80,80,80,80,80,90,160];
  var totalCol = headers.indexOf('TOTAL CORRECT') + 1;
  var mnfCol = headers.indexOf('MNF CORRECT') + 1;
  var weeklyCorrectAvgCol = headers.indexOf('WEEKLY AVG %') + 1;
  var weeklyRankAvgCol = headers.indexOf('WEEKLY AVG RANK') + 1;
  var weeklyWinsCol = headers.indexOf('WEEKLY WINS') + 1;
  var survivorCol = headers.indexOf('SURVIVOR (WEEK OUT)') + 1;
  var len = headers.length;
  var totalMembers = members.length;
  
  var rows = totalMembers + 1;
  var maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  if ( len < maxCols ) {
    sheet.deleteColumns(len,maxCols-len);
  }
  maxCols = sheet.getMaxColumns();
  
  sheet.getRange(1,1,1,len).setValues([headers]);
  for ( var a = 0; a < len; a++ ) {
    sheet.setColumnWidth(a+1,headersWidth[a]);
  }
  sheet.setRowHeight(1,40);
  var range = sheet.getRange(1,1,1,maxCols);
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  range = sheet.getRange(1,1,rows,len);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(members); 
  sheet.getRange(1,1,totalMembers+1,1).setHorizontalAlignment('left');
  
  range = sheet.getRange(1,1,1,len);
  range.setBackground('black');
  range.setFontColor('white');
  
  sheet.setFrozenColumns(1);
  sheet.setFrozenRows(1);
  
  sheet.clearConditionalFormatRules(); 
  // SUMMARY TOTAL GRADIENT RULE
  var rangeSummaryTot = sheet.getRange('R2C'+totalCol+':R'+rows+'C'+totalCol);
  var formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#75F0A1')
    .setGradientMinpoint('#FFFFFF')
    .setRanges([rangeSummaryTot])
    .build();
  // MNF TOTAL GRADIENT RULES
  var rangeMNFTot = sheet.getRange('R2C'+mnfCol+':R'+rows+'C'+mnfCol);
  //ss.setNamedRange('TOT_MNF_'+year,range);
  var formatRuleMNFTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#75F0A1')
    .setGradientMinpoint('#FFFFFF')
    .setRanges([rangeMNFTot])
    .build();
  // RANK GRADIENT RULES
  var rangeOverallRank = sheet.getRange('R2C'+(totalCol+1)+':R'+rows+'C'+(totalCol+1));
  var rangeMNFRank = sheet.getRange('R2C'+(mnfCol+1)+':R'+rows+'C'+(mnfCol+1));
  ss.setNamedRange('TOT_OVERALL_RANK_'+year,rangeOverallRank);
  ss.setNamedRange('TOT_MNF_RANK_'+year,rangeMNFRank);
  var formatRuleRanks = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([rangeOverallRank,rangeMNFRank])
    .build();
  // WEEKLY WINS GRADIENT/SINGLE COLOR RULES
  range = sheet.getRange('R2C'+weeklyWinsCol+':R'+rows+'C'+weeklyWinsCol);
  ss.setNamedRange('WEEKLY_WINS_'+year,range);
  var formatRuleWeeklyWins = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#ffee00')
    .setGradientMinpoint('#FFFFFF')
    .setRanges([range])
    .build();
  var formatRuleWeeklyWinsEmpty = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#FFFFFF')
    .setFontColor('#FFFFFF')
    .setRanges([range])
    .build();
  // WEEKLY CORRECT % AVG
  range = sheet.getRange('R2C'+weeklyCorrectAvgCol+':R'+rows+'C'+weeklyCorrectAvgCol);
  range.setNumberFormat('##.#%');
  var formatRuleCorrectAvg = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, ".70")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, ".60")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, ".50")
    .setRanges([range])
    .build();
  // WEEKLY RANK AVG
  range = sheet.getRange('R2C'+weeklyRankAvgCol+':R'+rows+'C'+weeklyRankAvgCol);
  range.setNumberFormat('#.#');
  var formatRuleCorrectRank = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, "5")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "10")
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "15")
    .setRanges([range])
    .build();
  // SURVIVOR "IN"
  range = sheet.getRange('R2C'+survivorCol+':R'+(totalMembers+1)+'C'+survivorCol);
  var formatRuleCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('IN')
    .setBackground('#C9FFDF')
    .setRanges([range])
    .build();
  // SURVIVOR "OUT"
  var formatRuleIncorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('OUT')
    .setBackground('#F2BDC2')
    .setRanges([range])
    .build();
  var formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleCorrect);
  formatRules.push(formatRuleIncorrect);
  formatRules.push(formatRuleOverallTot);
  formatRules.push(formatRuleMNFTot);
  formatRules.push(formatRuleRanks);
  formatRules.push(formatRuleWeeklyWinsEmpty);
  formatRules.push(formatRuleWeeklyWins);
  formatRules.push(formatRuleCorrectAvg);
  formatRules.push(formatRuleCorrectRank);
  sheet.setConditionalFormatRules(formatRules);
  
  // Creates all formulas for SUMMARY Sheet
  summarySheetFormulas(totalMembers,year);

  return sheet;  
}

function summaryformtest(){
  var totalMembers = memberList().length;
  var year = fetchYear();
  summarySheetFormulas(totalMembers,year);
}
// UPDATES SUMMARY SHEET FORMULAS
function summarySheetFormulas(totalMembers,year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SUMMARY');
  var headers = sheet.getRange('1:1').getValues().flat();
  var arr = ['PLAYER','TOTAL CORRECT','TOTAL RANK','MNF CORRECT','MNF RANK','WEEKLY AVG %','WEEKLY AVG RANK','WEEKLY WINS','SURVIVOR (WEEK OUT)','NOTES'];
  headers.unshift('COL INDEX ADJUST');
  var totalCorrect = false;
  var totalRnk = false;
  var mnfCorrect = false;
  var mnfRnk = false;
  var weeklyAvgPct = false;
  var weeklyAvgRnk = false;
  var weeklyWins = false;
  var survivor = false;
  for ( var a = 0; a < totalMembers; a++ ) {
    // TOTAL games correct column formula
    if (headers.indexOf(arr[1]) >= 0) {
      sheet.getRange(a+2,headers.indexOf(arr[1])).setFormulaR1C1('=iferror(vlookup(R[0]C1,{TOT_OVERALL_'+year+'_NAMES,TOT_OVERALL_'+year+'},2,false))');
    } else {
      totalCorrect = true;
    }
    // TOTAL RANK column formula
    if (headers.indexOf(arr[2]) >= 0) {
      sheet.getRange(a+2,headers.indexOf(arr[2])).setFormulaR1C1('=iferror(rank(R[0]C[-1],R2C[-1]:R'+ (totalMembers+1) + 'C[-1]))');
    } else {
      totalRnk = true;
    }
    // MNF games correct column formula
    if (headers.indexOf(arr[3]) >= 0) {
      sheet.getRange(a+2,headers.indexOf(arr[3])).setFormulaR1C1('=iferror(vlookup(R[0]C1,{MNF_'+year+'_NAMES,MNF_'+year+'},2,false))');
    } else {
      mnfCorrect = true;
    }
    // MNF RANK column formula
    if (headers.indexOf(arr[4]) >= 0) {
      sheet.getRange(a+2,headers.indexOf(arr[4])).setFormulaR1C1('=iferror(rank(R[0]C[-1],R2C[-1]:R'+ (totalMembers+1) + 'C[-1]))');
    } else {
      mnfRnk = true;
    }
    // WEEKLY AVG % column formula
    if (headers.indexOf(arr[5]) >= 0) {
      sheet.getRange(a+2,headers.indexOf(arr[5])).setFormulaR1C1('=iferror(vlookup(R[0]C1,{TOT_OVERALL_PCT_'+year+'_NAMES,TOT_OVERALL_PCT_'+year+'},2,false))');
    } else {
      weeklyAvgPct = true;
    }
    // WEEKLY AVG RANK column formula
    if (headers.indexOf(arr[6]) >= 0) { 
      sheet.getRange(a+2,headers.indexOf(arr[6])).setFormulaR1C1('=iferror(vlookup(R[0]C1,{TOT_OVERALL_RANK_'+year+'_NAMES,TOT_OVERALL_RANK_'+year+'},2,false))');
    } else {
      weekyAvgRnk = true;
    }
    // WEEKLY WINS column formula
    if (headers.indexOf(arr[7]) >= 0) {
      sheet.getRange(a+2,headers.indexOf(arr[7])).setFormulaR1C1('=iferror(countif(WEEKLY_WINNERS,R[0]C1))');
    } else {
      weeklyWins = true;
    }
    // SURVIVOR status column formula
    if (headers.indexOf(arr[8]) >= 0) {
      sheet.getRange(a+2,headers.indexOf(arr[8])).setFormulaR1C1('=iferror(arrayformula(if(isblank(vlookup(R[0]C1,{ELIMINATED_'+year+'_NAMES,ELIMINATED_'+year+'},2,false)),"IN","OUT ("&vlookup(R[0]C1,{ELIMINATED_'+year+'_NAMES,ELIMINATED_'+year+'},2,false)&")")))');
    } else {
      survivor = true;
    }
    var errArr = [totalCorrect,totalRnk,mnfCorrect,mnfRnk,weeklyAvgPct,weeklyAvgRnk,weeklyWins,survivor];
  }
  for (a = 0; a < errArr.length; a++) {
      if (errArr[a] == true){
        Logger.log('Error setting formula for ' + arr[a] + ' column')
      }
  }
  ss.setNamedRange('TOT_OVERALL_RANK_'+year,sheet.getRange('R2C'+headers.indexOf(arr[2])+':R'+(totalMembers+1)+'C'+headers.indexOf(arr[2])));
  ss.setNamedRange('TOT_MNF_RANK_'+year,sheet.getRange('R2C'+headers.indexOf(arr[4])+':R'+(totalMembers+1)+'C'+headers.indexOf(arr[4])));
  ss.setNamedRange('WEEKLY_WINS_'+year,sheet.getRange('R2C'+headers.indexOf(arr[7])+':R'+(totalMembers+1)+'C'+headers.indexOf(arr[7])));
  Logger.log('Updated formulas and ranges for summary sheet');
}

//------------------------------------------------------------------------
// CREATE FORMS - Tool to create initial form, and repopulate with matchups as needed
function formFiller(auto) {

  if (auto == null){
    auto = false;
  }
  // Fetch update to the NFL data to ensure most recent schedule
  fetchNFL(); 
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName('FORM');
  if ( formSheet == null ) {
    ss.insertSheet('FORM');
    formSheet = ss.getSheetByName('FORM');
  }
  
  var year = fetchYear();
  var first, week;
  try {
    week = ss.getRangeByName('WEEK').getValue();
    first = false;
    if (week == null) {
      var week = fetchWeek();
      first = true;
    }
  }
  catch (err){
    Logger.log('No Week Set Yet, checking API info');
    week = fetchWeek();
    ss.setNamedRange('WEEK',formSheet.getRange(1,2));
    ss.getRangeByName('WEEK').setValue(week);
    first = true;
  }
  
  var ui = SpreadsheetApp.getUi();
  try {
    if ( week < 11 ) {
      previousName = 'TOT_' + year + '_0' + (week-1);
    } else {
      previousName = 'TOT_' + year + '_' + (week-1);
    }
    var previousCorrect = 0;
    var previous = ss.getRangeByName(previousName).getValues().flat();
    for (var a = 0; a < previous.length; a++){
      previousCorrect = previousCorrect + previous[a];
    }
    if (previousCorrect <= 0) {
      week = week - 1;
    }
    if (week > 1) {
      if ( ss.getRangeByName('WEEK').getValue() != null ) {
        var formImported = ui.alert('Have you imported all data from previous form?', ui.ButtonSet.YES_NO);
        if ( formImported == 'NO' ) {
          var formImportNow = ui.alert('Run data transfer script right now?', ui.ButtonSet.YES_NO);
          if ( formImportNow == 'YES' ) {
            dataTransfer();
          }
        }
      }
    }
  }
  catch (err) {
    Logger.log('No previous form; moving on');
    first = true;
  }
  var formReset;
  var changeWeek;
  var newWeek;
  if ((week == 1 || first == true) && auto != true) {
    formReset = ui.alert('Initiate form for week ' + week + '?', ui.ButtonSet.YES_NO);
  } else if (auto == false) {
    formReset = ui.alert('Confirm erasing responses and recreate form for week ' + week + '?', ui.ButtonSet.YES_NO);
  } else {
    formReset = 'YES';
  }
  if ( formReset == 'NO' && auto != true ) {
    changeWeek = ui.alert('Create form for another week than ' + week + '?', ui.ButtonSet.YES_NO);
    if ( changeWeek == 'YES' ) {
      newWeek = ui.prompt('Specify new week:', ui.ButtonSet.OK);
      week = newWeek.getResponseText();
      formReset = ui.alert('Confirm erasing responses and recreate form for week ' + week + '?', ui.ButtonSet.YES_NO);
    }
  }
  if ( formReset == 'YES' ) {
    var mask = '_';
    if ( week < 10 ) {
      mask = '_0' 
    }
    
    ss.setNamedRange('WEEK',formSheet.getRange(1,2));
    ss.getRangeByName('WEEK').setValue(week);
    var sheet = ss.getSheetByName(year + mask + week);

    // Erases old version and re-creates/creates it, skips this if automatically recreating for membership locking/unlocking
    if (auto != true) {
      if ( sheet != null) {
        ss.deleteSheet(sheet);
      }
      weeklySheet(year,week,null,false)
    }
    ss.toast('Recreated sheet for week ' + week);
    var sheetId = ss.getId();
    var sheet;
    var urlSheetEdit = ss.getUrl();
    var urlSheetPub = urlSheetEdit.slice(0,-5);
    var form;
    var formId;
  
    var urlFormPub;
    var urlFormEdit;
    
    var arr = [['WEEK',week],['SHEET URL',urlSheetPub],['SHEET URL EDIT',urlSheetEdit],['SHEET ID',sheetId],['FORM URL',urlFormPub],['FORM URL EDIT',urlFormEdit],['FORM ID',formId]];
    
    var rows = arr.length;
    var cols = 2;
    var maxRows = formSheet.getMaxRows();
    var maxCols = formSheet.getMaxColumns();
    if (maxCols > cols) { formSheet.deleteColumns(cols,maxCols - cols) }
    if (maxRows > rows) { formSheet.deleteRows(rows,maxRows - rows) }
    formSheet.setColumnWidth(1,120);
    formSheet.setColumnWidth(2,700);
    var range = formSheet.getRange(rows,2);
    formSheet.getRange(1,2,rows,1).clearNote();
    formSheet.getRange(2,2).setNote('Use this to share to the group -- but make sure to make the spreadsheet shared for View Only with a link!');
    formSheet.getRange(4,2).setNote('Use this ID for the Google Form script ID needed');
    formSheet.getRange(5,2).setNote('Use this to share to the group for filling out the form');
    var first = false;
    if ( range.getValue() == '' ) {
      first = true;
      form = FormApp.create('Week ' + week + ' NFL Pick\’Ems');
      formId = form.getId();
      urlFormPub = form.getPublishedUrl();
      urlFormEdit = form.getEditUrl();
      var arr = [['WEEK',week],['SHEET URL',urlSheetPub],['SHEET URL EDIT',urlSheetEdit],['SHEET ID',sheetId],['FORM URL',urlFormPub],['FORM URL EDIT',urlFormEdit],['FORM ID',formId]];
      rows = arr.length;
      formSheet.getRange(1,1,rows,cols).setValues(arr);
    } else {
      formId = range.getValue();
      form = FormApp.openById(formId);
      urlFormPub = form.getPublishedUrl();
      urlFormEdit = form.getEditUrl();
      ss.setNamedRange('FORM_ID',range);
    }
    ss.setNamedRange('FORM_ID',range);
    ss.setNamedRange('WEEK',formSheet.getRange(1,2));  
    range = formSheet.getRange(1,1,rows,cols);
    range.setHorizontalAlignment('left');
    range.setVerticalAlignment('center');
    range.setFontFamily("Montserrat");
    range.setFontSize(10);
    
    form.deleteAllResponses();
    
    var year = fetchYear();
    
    var data;    
    try { data = ss.getRangeByName('NFL_' + year).getValues() }
    catch (err) {
      var ui = SpreadsheetApp.getUi();
      var fetchNFLPrompt = ui.alert('It looks like NFL data hasn\'t been brought in, import now?', ui.ButtonSet.YES_NO);
      
      if ( fetchNFLPrompt == 'YES' ) {
        fetchNFL();
      } else {
        ui.alert('Please run again and import NFL first or click \'YES\' next time', ui.ButtonSet.OK);
      }
    }
    
    try { 
      // Import all NFL data to create form
      data = ss.getRangeByName('NFL_' + year).getValues();

      var members = ss.getRangeByName('MEMBERS').getValues().flat();
      var item, day, time, minutes;
      var teams = [];
      
      // Update form title, ensure description and confirmation are set
      form.setTitle('Week ' + week + ' NFL Pick\’Ems')
        .setDescription('Good Luck!')
        .setConfirmationMessage('Thanks for responding!')
        .setAllowResponseEdits(false)
        .setAcceptingResponses(true);
      // Update the form's response destination.
      //form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      
      // Clear all questions from previous weeks
      var items = form.getItems().length;
      for (var i = items; i > 0; i--) {
        form.deleteItem(i-1);
      }
      
      // Create name prompt
      // Open a form by ID
      if (membersSheetProtected() == false) {
        // Creates an item for text entry for the first week of the season
        var membersText = members[0];
        if (members.length >= 3) {
          membersText = membersText + ', ';
          for (var a = 1; a < members.length; a++) {
          membersText = membersText + members[a];
            if (a < members.length - 2) {
              membersText = membersText + (', ');
            } else if (a < members.length - 1){
              membersText = membersText + (', and ');
            }
          }
        } else if (members.length == 2){
          membersText = membersText + ' and ' + members[1];
        }
        item = form.addTextItem();
        item.setTitle('Please type your first and last name')
        item.setHelpText('The following members exist: \r\n' + membersText + '\r\n Type your name as it appears above, if you\'re not already on the list, enter your name here.')
          .setRequired(true);
      } else {
        // Add drop-down list of previous entries
        item = form.addListItem();
        item.setTitle('Select your name')
          .setChoiceValues(members)
          .setRequired(true);
      }
      for ( i = 0; i < data.length; i++ ) {
        if ( data[i][0] == week ) {
          teams.push(data[i][6]);
          teams.push(data[i][7]);
          item = form.addMultipleChoiceItem();
          if ( data[i][2] == 1 ) {
            day = 'Monday Night Football';
          } else {
            day = data[i][5];
          }
          if (data[i][4] < 10) {
            minutes = '0' + data[i][4];
          } else {
            minutes = data[i][4];
          }
          if ( data[i][3] == 12 ) {
            time = data[i][3] + ':' + minutes + ' PM'; //case for 1pm start or later (24 hour time converted to standard 12 hour format)
          } else if ( data[i][3] > 12 ) {
            time = (data[i][3] - 12) + ':' + minutes + ' PM'; //case for 1pm start or later (24 hour time converted to standard 12 hour format)
          } else {
            time = data[i][3]  + ':' + minutes + ' AM'; // early (pre-noon) game start time with two digits for minutes
          }
          item.setTitle(day + ' at ' + time + ': ' + data[i][8] + ' ' + data[i][9] + ' at ' + data[i][10] + ' ' + data[i][11])
          .setChoices([
            item.createChoice(data[i][6]),
            item.createChoice(data[i][7])
          ])
          .showOtherOption(false)
          .setRequired(true);
        }
      }
      
      teams.sort();
      teams.unshift('NA');
      
      var numberValidation = FormApp.createTextValidation()
        .setHelpText('Input must be a whole number between 0 and 150')
        .requireWholeNumber()
        .requireNumberBetween(0,150)
        .build();      
      
      // Tiebreaker question
      item = form.addTextItem();
      item.setTitle('Tiebreaker (Total Points)')
        .setRequired(true)
        .setValidation(numberValidation);
      
      // Survivor question
      item = form.addListItem();
      item.setTitle('Survivor pick (\"NA\" if out):')
        .setChoiceValues(teams)
        .setRequired(true);
      
      // Passing Thoughts
      item = form.addTextItem();
      item.setTitle('Passing thoughts...');
       
      var ui = SpreadsheetApp.getUi();
      
    // Update all formulas to account for new weekly sheets that may have been created
      allFormulasUpdate();
      var tab;
      if (first == true) {
        tab = ui.alert('Google Form created for week ' + week + '! \r\n\r\nDon\'t forget to set a football picture for the Form and maybe adjust the color scheme to your preference.\r\n\r\nWould you like to open the editable Google Form in a new tab to make updates to the look?', ui.ButtonSet.YES_NO);
        if ( tab == 'NO') {
          var pub = ui.alert('Google Form for week ' + week + ' Shareable Link:\r\n' + urlFormPub + '\r\n\r\nWould you like to open the weekly Google Form in a new tab?', ui.ButtonSet.YES_NO);
          if ( pub == 'YES' ) {
            openUrl(urlFormPub);
          }
        } else {
          ui.alert('Google Form for week ' + week + ' Shareable Link:\r\n' + urlFormPub, ui.ButtonSet.OK);
          openUrl(urlFormEdit);
        }
      } else {
        tab = ui.alert('Google Form updated for week ' + week + '! \r\n\r\nShareable Link:\r\n' + urlFormPub + '\r\n\r\nWould you like to open the weekly Google Form in a new tab?', ui.ButtonSet.YES_NO);
        if ( tab == 'YES' ) {
          openUrl(urlFormPub);
        }
      }
    }
    catch (err) {
      Logger.log('Aborted due to error ' + err.message);
    } 
  } else {
    ss.toast('Canceled form creation');
  }
}

// OPEN URL - Quick script to open a new tab with the newly created form, in this case
function openUrl(url){
  var js = "<script>window.open('" + url + "', '_blank');google.script.host.close();</script>;"
  var html = HtmlService.createHtmlOutput(js)
    .setHeight(10)
    .setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening...');
}

//------------------------------------------------------------------------
// OPEN FORM - Function to open the Google Form quickly from the menu
function openForm() {
  var formId = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('FORM_ID').getValue();
  if (formId == null || formId == ''){
    var ui = SpreadsheetApp.getUi();
    var alert = ui.alert('No Form created yet, would you like to create one now?', ui.ButtonSet.YES_NO);
    if (alert == 'YES') {
      formFiller(auto);
    } else {
      ui.alert('Try again after you\'ve created the initial Google Form.', ui.ButtonSet.OK);
    }
  } else {
    var form = FormApp.openById(formId);
    var urlFormPub = form.getPublishedUrl();
    openUrl(urlFormPub);
  }
}

//------------------------------------------------------------------------
// CHECK SUBMISSIONS - Tool to check who's submitted the weekly form so far
function formCheck(request) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  try {
    var formId = ss.getRangeByName('FORM_ID').getValue();
    var form = FormApp.openById(formId);
    var formResponses = form.getResponses(); 
    var response;
    var itemResponses;
    var membersArr = memberList();
    var members = membersArr.flat();
    var name;
    var names = [];
    
    for (var a = 0; a < formResponses.length; a++ ) {
      response = formResponses[a];
      itemResponses = response.getItemResponses();    
      name = itemResponses[0].getResponse();
      names.push(name);
      if (members.indexOf(name) >= 0) {
        members.splice(members.indexOf(name),1);
      }
    }
    if (request == null || request == undefined || request == "missing") {
      return members;
    } else if (request == "received") {
      return names;
    } else if (request == "new") {
      for (var b = 0; b < membersArr.length; b++) {
        for (var c = 0; c < names.length; c++) {
          if (membersArr[b] == names[c]){
            names.splice(c,1);
          }          
        }
      }
      return names;
    }
  }
  catch (err) {
      Logger.log('formCheck ' + err.message);
      var noForm = ui.alert('No Google Form created yet, run \"Update Form\" from the \"Pick\’Ems\" menu', ui.ButtonSet.OK);
  } 
}

//------------------------------------------------------------------------
// ALERT FOR SUBMISSION CHECK
function formCheckAlert(missing) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var formSheet = ss.getSheetByName('FORM');
  try {
    var formId = ss.getRangeByName('FORM_ID').getValue();
    if ( formSheet == null || formId == null ) {
      var noForm = ui.alert('No Google Form created yet, run \"Update Form\" from the \"Pick\’Ems\" menu', ui.ButtonSet.OK);
    } else {
      var form = FormApp.openById(formId);
      var week = ss.getRangeByName('WEEK').getValue();
      var a;
      var membersNew = formCheck("new");
      var textNew = '';
      for ( a = 0; a < membersNew.length; a++) {
        if (a < membersNew.length - 1) {
          textNew = textNew.concat(membersNew[a] + '\r\n');
        } else {
          textNew = textNew.concat(membersNew[a]);
        }
      }
      if ( membersNew.length > 0 ) {
        for (var b = 0; b < membersNew.length; b++) {
          memberAdd(membersNew[b]);
        }
        if (membersNew.length > 1) {
          var membersNewAlert = ui.alert('Added new member(s): \r\n' + textNew, ui.ButtonSet.OK);
        } else {
          var membersNewAlert = ui.alert('Added one new member: ' + textNew, ui.ButtonSet.OK);
        }
      }
       
      var text = '';
      for ( a = 0; a < missing.length; a++) {
        if (a < missing.length - 1) {
          text = text.concat(missing[a] + '\r\n');
        } else {
          text = text.concat(missing[a]);
        }
      }
      var totalMembers = memberList().length;
      if (missing.length >= totalMembers) {
        var respondents = ui.alert('No responses recorded yet for this week.', ui.ButtonSet.OK);
      } else if (missing.length == 0) {
        var respondents = ui.alert('All responses logged for week ' + week + ', import data now?' + text, ui.ButtonSet.YES_NO);
        if ( respondents == 'YES' ) {
          dataTransfer(1);
        }
      } else if (missing.length == 1) {
        var respondents = ui.alert(text + ' is the only one who hasn\'t responded', ui.ButtonSet.OK);
      } else if (missing.length == 2) {
        var respondents = ui.alert(missing[0] + ' and ' + missing[1] + ' are the only two who haven\'t responded', ui.ButtonSet.OK);
      } else if (missing.length == 3) {
        var respondents = ui.alert(missing[0] + ', ' + missing[1] + ', and ' + missing[2] + ' are the only three who haven\'t responded', ui.ButtonSet.OK);
      } else if (missing.length >= 4) {
        var respondents = ui.alert('These ' + missing.length + ' players haven\'t responded for week ' + week + ': \r\n' + text, ui.ButtonSet.OK);
      }
    }
  }
  catch (err) {
    Logger.log('formCheckAlert ' + err.message);
  }
}

//------------------------------------------------------------------------
// REQUEST FORM CHECK ALERT
function formCheckAlertCall() {
  var members = formCheck("missing");
  if ( members != null ) {
    formCheckAlert(members);
  }
}

//------------------------------------------------------------------------
// DATA IMPORTING - Function to import responses from the surveys
function dataTransfer(redirect) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = fetchYear();
  var week = ss.getRangeByName('WEEK').getValue();
  var members = memberList().flat();
  var missing = formCheck("missing");
  var received = formCheck("received");
  var textMissing = '';
  var textReceived = '';
  for (var a = 0; a < missing.length; a++) {
    if (a < missing.length - 1) {
      textMissing = textMissing.concat(missing[a] + '\r\n');
    } else {
      textMissing = textMissing.concat(missing[a]);
    }
  }
  for (a = 0; a < received.length; a++) {
    if (a < received.length - 1) {
      textReceived = textReceived.concat(received[a] + '\r\n');
    } else {
      textReceived = textReceived.concat(received[a]);
    }
  }
  var remaining = members.length - (members.length - missing.length);
  
  var ui = SpreadsheetApp.getUi();
  if ( redirect == null ) {
    var prompt;
    if (received == 0) {
      prompt = ui.alert('No responses received yet', ui.ButtonSet.OK);
    } else if (received > 0 || week == 1) {
      prompt = ui.alert('Received responses from the following: \r\n' + textReceived + '\r\n Import picks now?', ui.ButtonSet.YES_NO);
    } else if (missing.length == 0) {
      prompt = ui.alert('All responses logged for week ' + week + '. Import picks now?', ui.ButtonSet.YES_NO);
    } else {
      if (missing.length == 1) {
        var respondents = ui.alert(textMissing + ' is the only one who hasn\'t responded', ui.ButtonSet.OK);
      } else if (missing.length == 2) {
        var respondents = ui.alert(missing[0] + ' and ' + missing[1] + ' are the only two who haven\'t responded', ui.ButtonSet.OK);
      } else if (missing.length == 3) {
        var respondents = ui.alert(missing[0] + ', ' + missing[1] + ', and ' + missing[2] + ' are the only three who haven\'t responded', ui.ButtonSet.OK);
      } else if (missing.length >= 4) {
        var respondents = ui.alert('These ' + missing.length + ' players haven\'t responded for week ' + week + ': \r\n' + textMissing, ui.ButtonSet.OK);
      }
      prompt = ui.alert('Would you like to still import picks despite missing ' + remaining + '?', ui.ButtonSet.YES_NO);
    }
  } else {
    prompt = 'YES';
  }
  if (prompt == 'YES') {
  
    var formId = ss.getRangeByName('FORM_ID').getValue();
    var form = FormApp.openById(formId);
    var formResponses = form.getResponses();    
    var itemResponse;    
    var response;
    var itemResponses;
    var week;
    var survivor;
    var membersArr = ss.getRangeByName('MEMBERS').getValues(); 
    var members = membersArr.flat();
    var sheet;
    var sheetName;
    var row;
    var name;
    var i;
    var j;
    
    for ( i = 0; i < formResponses.length; i++ ) {
      response = formResponses[i];
      itemResponses = response.getItemResponses();

      week = parseInt(week);
      if ( week < 10 ) {
        sheetName = year + '_0' + ( week );
      } else {
        sheetName = year + '_' + ( week );
      }
      sheet = ss.getSheetByName(sheetName);  
      if (sheet == null) {
        weeklySheet(year,week,membersArr,false);
        sheet = ss.getSheetByName(sheetName);
        Logger.log('New weekly sheet created for week ' + week + ', \"' + sheetName + "\.")
      }
      
      name = toTitleCase(itemResponses[0].getResponse());
      if (members.indexOf(name) < 0) {
          memberAdd(name);
          members.push(name);
      }
      row = members.indexOf(name) + 3;
      sheet.getRange(row,1).setValue(name);
      
      for ( j = 1; j < (itemResponses.length - 2); j++) {
        itemResponse = itemResponses[j];
        sheet.getRange(row,j+4).setValue(itemResponse.getResponse());
      }
      itemResponse = itemResponses[j+1];
      // Adding response to MISC column
      sheet.getRange(row,j+7).setValue(itemResponse.getResponse());
      itemResponse = itemResponses[j];
      survivor = itemResponse.getResponse();
      if ( survivor == 'NA' ) {
        survivor = '';
      }
      sheet = ss.getSheetByName('SURVIVOR');
      if (sheet == null) {
        ss.insertSheet('SURVIVOR');
        sheet = ss.getSheetByName('SURVIVOR');
      }
      sheet.getRange(row-1,week+2).setValue(survivor);

    }
    ss.toast('Imported/updated responses for week ' + week );
    
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Canceled');
  }
}

//------------------------------------------------------------------------
// DATA IMPORTING - Function to import responses from the surveys
function dataTransferTNF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = fetchYear();
  var week = ss.getRangeByName('WEEK').getValue();
  var members = memberList().flat();
  var missing = formCheck("missing");
  var received = formCheck("received");
  var textMissing = '';
  var textReceived = '';
  for ( var a = 0; a < missing.length; a++) {
    if (a < missing.length - 1) {
      textMissing = textMissing.concat(missing[a] + '\r\n');
    } else {
      textMissing = textMissing.concat(missing[a]);
    }
  }
  for (a = 0; a < received.length; a++) {
    if (a < received.length - 1) {
      textReceived = textReceived.concat(received[a] + '\r\n');
    } else {
      textReceived = textReceived.concat(received[a]);
    }
  }
  var thursGames = 0;
  var nfl = ss.getRangeByName('NFL_' + year).getValues();
  for ( a = 0; a < nfl.length; a++ ) {
    if ( nfl[a][0] == week && nfl[a][2] == -3 ) {
      Logger.log(nfl[a][2]);
      thursGames++;
    }
  }

  var remaining = members.length - (members.length - missing.length);
  var thursPrompt;
  var thursBool = false;
  var ui = SpreadsheetApp.getUi();
  if ( thursGames == 1 ) {
    thursPrompt = ui.alert('There is one Thursday games this week');
    thursBool = true;
  } else if ( thursGames > 1 ) {
    thursPrompt = ui.alert('There are ' + thursGames + ' Thursday games this week');
    thursBool = true;
  } else {
    thursPrompt = ui.alert('There are no Thursday games this week.');
  }
  var prompt;
  if (thursBool == true) {
    if (week == 1 && remaining < members.length) {
      prompt = ui.alert('Responses logged for week ' + week + ' by the following: \r\n' + textReceived + '\r\n Import picks now?', ui.ButtonSet.YES_NO);
    } else if (missing.length == 0) {
      prompt = ui.alert('All responses logged for week ' + week + '. Import ALL picks now?', ui.ButtonSet.YES_NO);
      if (prompt == 'YES') {
        prompt = 'ALL';
      } else {
        prompt = ui.alert('Import the Thursday picks now?', ui.ButtonSet.YES_NO);
      }
    } else {
      if (remaining == members.length) {
        var respondents = ui.alert('No responses recorded yet for this week.', ui.ButtonSet.OK);
        prompt == 'NO'
      } else if (missing.length == 1) {
        var respondents = ui.alert(textMissing + ' is the only one who hasn\'t responded', ui.ButtonSet.OK);
        prompt = ui.alert('Would you like to still import the Thursday picks despite missing 1 response?', ui.ButtonSet.YES_NO);
      } else if (missing.length == 2) {
        var respondents = ui.alert(missing[0] + ' and ' + missing[1] + ' are the only two who haven\'t responded', ui.ButtonSet.OK);
        prompt = ui.alert('Would you like to still import the Thursday picks despite missing 2 responses?', ui.ButtonSet.YES_NO);
      } else if (missing.length == 3) {
        var respondents = ui.alert(missing[0] + ', ' + missing[1] + ', and ' + missing[2] + ' are the only three who haven\'t responded', ui.ButtonSet.OK);
        prompt = ui.alert('Would you like to still import the Thursday picks despite missing 3 responses?', ui.ButtonSet.YES_NO);
      } else if (missing.length >= 4) {
        var respondents = ui.alert('These ' + missing.length + ' players haven\'t responded for week ' + week + ': \r\n' + textMissing, ui.ButtonSet.OK);
        prompt = ui.alert('Would you like to still import the Thursday picks despite missing ' + remaining + ' responses?', ui.ButtonSet.YES_NO);
      }
    }
    if ( prompt == 'ALL' ) { // Bring in all information since all submissions are completed
      dataTransfer(1);
    } else if ( prompt == 'YES' ) { // Bring in only submissions provided for Thursday night games thus far
      var formId = ss.getRangeByName('FORM_ID').getValue();
      var form = FormApp.openById(formId);
      var formResponses = form.getResponses();    
      var itemResponse;    
      var response;
      var itemResponses;
      var week;
      var members = ss.getRangeByName('MEMBERS').getValues().flat();
      var sheet;
      var sheetName;
      var row;
      var name;
      var i;
      var j;
      
      for ( i = 0; i < formResponses.length; i++ ) {
        response = formResponses[i];
        itemResponses = response.getItemResponses();
  
        week = parseInt(week);
        if (week < 10) {
          sheetName = year + '_0' + week;
        } else {
          sheetName = year + '_' + week;
        }
        sheet = ss.getSheetByName(sheetName);  
        if (sheet == null) {
          ss.insertSheet(sheetName);
          sheet = ss.getSheetByName(sheetName);
        }
        
        name = itemResponses[0].getResponse();
        if (members.indexOf(name) < 0) {
          memberAdd(name);
          members.push(name);
        }
        row = members.indexOf(name) + 3;
        sheet.getRange(row,1).setValue(name);
        
        for ( j = 1; j < 2 + (thursGames - 1); j++) {
          itemResponse = itemResponses[j];
          sheet.getRange(row,j+4).setValue(itemResponse.getResponse());
        }
        
      }
      ss.toast('Imported/updated responses for week ' + week );
      
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      ss.toast('Canceled');
    }
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Canceled');
  }
}

//------------------------------------------------------------------------
// SERVICE Function to change a string to title
function toTitleCase(str) {
  return str.replace(
    /\w\S*/g,
    function(txt) {
      return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    }
  );
}

// SERVICE Function to remove all triggers on project
function deleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function resetSpreadsheet() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.alert('Reset spreadsheet and delete all data?', ui.ButtonSet.YES_NO);
  if (prompt == 'YES') {
    var promptTwo = ui.alert('Are you sure? This would be very difficult to recover from.',ui.ButtonSet.YES_NO)
    if (promptTwo == 'YES') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var ranges = ss.getNamedRanges();
      for (var a = 0; a < ranges.length; a++){
        ranges[a].remove();
      }
      var sheets = ss.getSheets();
      var baseSheet = ss.insertSheet();
      for (a = 0; a < sheets.length; a++){
        ss.deleteSheet(sheets[a]);
      }
      var protections = ss.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (a = 0; a < protections.length; a++){
        protections[a].remove();
      }
      protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (a = 0; a < protections.length; a++){
        protections[a].remove();
      }      
      baseSheet.setName('Sheet1');
      var menu = SpreadsheetApp.getUi().createMenu('Setup')
      menu.addItem('Run First','runFirst')
      .addToUi();
    } else {
      ss.toast('Canceled reset');
    }
  } else {
    ss.toast('Canceled reset');
  }
  
}

// 2022 - Created by Ben Powers
// ben.powers.creative@gmail.com
