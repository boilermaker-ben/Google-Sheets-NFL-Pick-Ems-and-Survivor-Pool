
// SEASON INFORMATION FUNCTIONS
//------------------------------------------------------------------------
// FETCH CURRENT YEAR
function fetchYear() {
  let year;
  const scriptProperties = PropertiesService.getScriptProperties();
    try {
      year = scriptProperties.getProperty('year');
      if (year == null) {
        try {
          const obj = JSON.parse(UrlFetchApp.fetch(scoreboard).getContentText());
          scriptProperties.setProperty('year',obj.season.year.toString());
          year = scriptProperties.getProperty('year');
        }
        catch (err) {
          Logger.log('Unable to fetch year from ESPN API currently, try again later');
          year = fallbackYear;
        }
      }
    }
    catch (err) {
      Logger.log('Error fetching year script variable or using API, using fallback year [global variable]');
      year = fallbackYear;
    }
  return (parseInt(year).toFixed(0));
}

// FETCH CURRENT WEEK
function fetchWeek(negative) {
  let weeks, week, advance = 0;
  try {
    
    const obj = JSON.parse(UrlFetchApp.fetch(scoreboard).getContentText());
    
    let season = obj.season.type;
    
    obj.leagues[0].calendar.forEach(entry => {
      if (entry.value == season) {
        weeks = entry.entries.length;
      }
    });

    obj.events.forEach(event => {
      if (event.status.type.state != 'pre') {
        advance = 1; // At least one game has started and therefore the script will prompt for the next week
      }
    });

    let name;
    switch (season) {
      case 1:
        name = 'Preseason';
        week = obj.week.number - (weeks + 1);
        break;
      case 2: 
        name = 'Regular season';
        week = obj.week.number + advance;
        break;
      case 3:
        name = 'Postseason';
        week = obj.week.number + obj.leagues[0].calendar[1].entries.length + advance;
        break;
    }
    Logger.log(name + ' is currently active with ' + weeks + ' weeks in total'); 
    if (negative) {
      return week;
    } else {
      week = week <= 0 ? 1 : week;
      return week;
    }
  }
  catch (err) {
    Logger.log('ESPN API has an issue right now' + err.stack);
    return null;
  }
}

// FETCH TOTAL WEEKS - doesn't remove the week 22 pro bowl
function fetchWeeks() {
  try {
    let weeks = 0;
    const content = UrlFetchApp.fetch(scoreboard).getContentText();
    const obj = JSON.parse(content);
    const calendar = obj.leagues[0].calendar;
    for (let a = 0; a < calendar.length; a++) {
      if (calendar[a].value == 2 || calendar[a].value == 3) {
        weeks += calendar[a].entries.length;
      }
    }
    return weeks;
  }
  catch (err) {
    Logger.log('ESPN API has an issue right now');
    return 23;
  }
}

// ESPN FUNCTIONS
//------------------------------------------------------------------------
// ESPN TEAMS - Fetches the ESPN-available API data on NFL teams
function fetchTeamsESPN(year) {
  if (year == undefined) {
    year = fetchYear();
  }
  let obj = {};
  try {
    let string = schedulePrefix + year + scheduleSuffix;
    obj = JSON.parse(UrlFetchApp.fetch(string).getContentText());
    let objTeams = obj.settings.proTeams;
    return objTeams;
  }
  catch (err) {
    Logger.log('ESPN API has an issue right now');
  }  
}


// NFL TEAM INFO - script to fetch all NFL data for teams - auto for setting up trigger allows for boolean entry in column near the end
function fetchSchedule(ss,year,currentWeek,auto,overwrite) {
  // Calls the linked spreadsheet
  const timeFetched = new Date();
  ss = fetchSpreadsheet(ss);
  let all = false;
  if (currentWeek == undefined || currentWeek == null) {
    currentWeek = fetchWeek(null,true);
    all = true;
    ss.toast(`Fetching complete schedule data for the ${LEAGUE}`,`📅 FETCHING SCHEDULE`);
  } else {
    ss.toast(`Fetching only data for week ${currentWeek}, if available.`,`📅 FETCHING WEEK ${currentWeek}`);
  }
  // Declaration of script variables
  if (year == undefined || year == null) {
    year = fetchYear();
  }
  const objTeams = fetchTeamsESPN(year);
  const teamsLen = objTeams.length;
  let headers = ['week','date','day','hour','minute','dayName','awayTeam','homeTeam','awayTeamLocation','awayTeamName','homeTeamLocation','homeTeamName','type','divisional','division','overUnder','spread','spreadAutoFetched','timeFetched'];
  let sheetName = LEAGUE;
  let sheet, range, abbr, name, arr = [], nfl = [],espnId = [], espnAbbr = [], espnName = [], espnLocation = [], location = [], ids = [], abbrs = []; 
  
  for (let a = 0 ; a < teamsLen ; a++ ) {
    arr = [];
    if(objTeams[a].id != 0 ) {
      abbr = objTeams[a].abbrev.toUpperCase();
      name = objTeams[a].name;
      location = objTeams[a].location;
      espnId.push(objTeams[a].id);
      espnAbbr.push(abbr);
      espnName.push(name);
      espnLocation.push(location);
    }
  }

  for (let a = 0 ; a < espnId.length ; a++ ) {
    ids.push(espnId[a].toFixed(0));
    abbrs.push(espnAbbr[a]);
  }

  // Declaration of variables
  let schedule = [], home = [], dates = [], allDates = [], hours = [], allHours = [], minutes = [], allMinutes = [], byeIndex, id, date, hour, minute, weeks = Object.keys(objTeams[0].proGamesByScoringPeriod).length;
  if ( objTeams[0].byeWeek > 0 ) {
    weeks++;
  }

  location = [];
  
  for (let a = 0 ; a < teamsLen ; a++ ) {
    arr = [];
    home = [];
    dates = [];
    hours = [];
    minutes = [];
    byeIndex = objTeams[a].byeWeek.toFixed(0);
    if ( byeIndex != 0 ) {
      id = objTeams[a].id.toFixed(0);
      arr.push(abbrs[ids.indexOf(id)]);
      home.push(abbrs[ids.indexOf(id)]);
      dates.push(abbrs[ids.indexOf(id)]);
      hours.push(abbrs[ids.indexOf(id)]);
      minutes.push(abbrs[ids.indexOf(id)]);
      for (let b = 1 ; b <= weeks ; b++ ) {
        if ( b == byeIndex ) {
          arr.push('BYE');
          home.push('BYE');
          dates.push('BYE');
          hours.push('BYE');
          minutes.push('BYE');
        } else {
          if ( objTeams[a].proGamesByScoringPeriod[b][0].homeProTeamId.toFixed(0) === id ) {
            arr.push(abbrs[ids.indexOf(objTeams[a].proGamesByScoringPeriod[b][0].awayProTeamId.toFixed(0))]);
            home.push(1);
            date = new Date(objTeams[a].proGamesByScoringPeriod[b][0].date);
            dates.push(date);
            hour = date.getHours();
            hours.push(hour);
            minute = date.getMinutes();
            minutes.push(minute);
          } else {
            arr.push(abbrs[ids.indexOf(objTeams[a].proGamesByScoringPeriod[b][0].homeProTeamId.toFixed(0))]);
            home.push(0);
            date = new Date(objTeams[a].proGamesByScoringPeriod[b][0].date);
            dates.push(date);
            hour = date.getHours();
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
      allMinutes.push(minutes);
    }
  }
  
  // This section creates a nice table to be used for lookups and queries about NFL season
  let week, awayTeam, awayTeamName, awayTeamLocation, homeTeam, homeTeamName, homeTeamLocation, day, dayName, divisional, division, scheduleData = [];
  
  // Create an array of matchups per week where index of 0 is equivalent to week 1 and so forth
  let matchupsPerWeek = Array(WEEKS).fill(0);
  arr = [];
  let weekArr = [];
  for (let b = 0; b < (teamsLen - 1); b++ ) {
    for ( let c = 1; c <= weeks; c++ ) {
      if (location[b][c] == 1) {
        week = c;
        awayTeam = schedule[b][c];
        awayTeamName = espnName[espnAbbr.indexOf(awayTeam)];
        awayTeamLocation = espnLocation[espnAbbr.indexOf(awayTeam)];
        homeTeam = schedule[b][0];
        homeTeamName = espnName[espnAbbr.indexOf(homeTeam)];
        homeTeamLocation = espnLocation[espnAbbr.indexOf(homeTeam)];
        date = allDates[b][c];
        hour = allHours[b][c];
        minute = allMinutes[b][c];
        day = date.getDay();
        // Uses globalVariables.gs variable to determine day name and assign offset index
        dayName = DAY[day].name;
        day = DAY[day].index;
        divisional = LEAGUE_DATA[homeTeam].division_opponents.indexOf(awayTeam) > -1 ? 1 : 0;
        division = divisional == 1 ? LEAGUE_DATA[homeTeam].division : '';

        arr = [
          week,
          date,
          day,
          hour,
          minute,
          dayName,
          awayTeam,
          homeTeam,
          awayTeamLocation,
          awayTeamName,
          homeTeamLocation,
          homeTeamName,
          WEEKNAME.hasOwnProperty(c) ? WEEKNAME[c].name : 'Regular Season', // type
          divisional,
          division,
          '', // Placeholder for overUnder
          '', // Placeholder for spread
          '', // Placeholder for spreadAutoFetched
          timeFetched
        ];
        matchupsPerWeek[week-1] = matchupsPerWeek[week-1] + 1;
        scheduleData.push(arr);
      }
    }
  }

  scheduleData = scheduleData.sort((a,b) => a[1] - b[1]);
    
  for (let a = 0; a < scheduleData.length; a++) {
    weekArr.push(scheduleData[a][0]);
  }
  // Add the playoff schedule to that array of matchups per week
  Object.keys(WEEKNAME).forEach(weekNum => {
    matchupsPerWeek[weekNum-1] = WEEKNAME[weekNum].matchups;
    for (let a = 0; a < matchupsPerWeek[weekNum-1]; a++) {
      weekArr.push(parseInt(weekNum));
    }
  });

  // Create indexing array of when weeks begin and end
  let rowIndex = 2;
  let startingRow = Array(WEEKS).fill(0);
  for (let a = 1; a < startingRow.length; a++) {
    let start = 0;
    for (let b = 0; b < a; b++) {
      start = start + matchupsPerWeek[b];
    }
    startingRow[a] = 2 + start;
  }


  // Sheet formatting & Range Setting =========================
  sheet = ss.getActiveSheet();
  if ( sheet.getSheetName() == 'Sheet1' && ss.getSheetByName(sheetName) == null) {
    sheet.clear();
    sheet.setName(sheetName);
  }
  sheet = ss.getSheetByName(sheetName);  
  if (sheet == null) {
    ss.insertSheet(sheetName,0);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.setTabColor(scheduleTabColor);

  adjustColumns(sheet,headers.length);
  
  sheet.setColumnWidths(1,headers.length,30);
  sheet.setColumnWidth(headers.indexOf('date')+1,60);
  sheet.setColumnWidth(headers.indexOf('dayName')+1,60);
  sheet.setColumnWidths(headers.indexOf('awayTeamLocation')+1,4,80); // All Locations & Team Names
  sheet.setColumnWidth(headers.indexOf('type')+1,110);
  sheet.setColumnWidth(headers.indexOf('division')+1,60);
  sheet.setColumnWidth(headers.indexOf('spread')+1,60);
  sheet.setColumnWidth(headers.indexOf('timeFetched')+1,110);
  range = sheet.getRange(1,1,1,headers.length);
  
  range.setValues([headers]);
  ss.setNamedRange(sheetName+'_HEADERS',range);

  range = sheet.getRange(1,1,weekArr.length+1,headers.length);
  range.setFontSize(8);
  range.setVerticalAlignment('middle');  
 
  ss.setNamedRange(sheetName,range);
  let rangeData = sheet.getRange(2,1,weekArr.length,headers.length);

  rangeData.setHorizontalAlignment('left');
  sheet.getRange(1,3).setNote('-4: Wednesday, -3: Thursday, -2: Friday, -1: Saturday, 0: Sunday, 1: Monday, 2: Tuesday');
  
  // Fetches sorted data
  // Sets named ranges for weekly home and away teams to compare for survivor status
  awayTeam = headers.indexOf('awayTeam')+1;
  homeTeam = headers.indexOf('homeTeam')+1;
  ss.setNamedRange(`${LEAGUE}_MATCHUPS_HEADERS`,sheet.getRange(1,1,1,headers.length));
  for (let a = 0; a < WEEKS; a++) {
    if (matchupsPerWeek[a] > 0) {
      try {
        let start = weekArr.indexOf(a+1)+2;
        let len = matchupsPerWeek[a];
        ss.setNamedRange(`${LEAGUE}_AWAY_${a+1}`,sheet.getRange(start,awayTeam,len,1));
        ss.setNamedRange(`${LEAGUE}_HOME_${a+1}`,sheet.getRange(start,homeTeam,len,1));
        ss.setNamedRange(`${LEAGUE}_MATCHUPS_${a+1}`,sheet.getRange(start,1,len,headers.length));
      }
      catch (err) {
        Logger.log(`No data entered or available for week ${a} in the spreadsheet`);
        Logger.log(err.stack);
      }
    } else {
      Logger.log(`No matchups in week ${a}`);
    }
  }
  // Sheet formatting =========================


  // Set of loops to create blank entries for playoff schedule
  const blankRow = new Array(headers.length).fill('');
  for (let a = (REGULAR_SEASON+1); a <= WEEKS; a++) {
    if (WEEKNAME.hasOwnProperty(a)) {
      for (let b = 0; b < WEEKNAME[a].matchups; b++) {
        let newRow = [...blankRow];
        newRow[0] = a; // Replace first value with week number
        scheduleData.push(newRow);
      }
    }
  }

  // Get scoreboard data
  const obj = JSON.parse(UrlFetchApp.fetch(SCOREBOARD));
  let scoreboardData = [];
  for (let event = 0; event < obj.events.length; event++) {
    date = new Date(obj.events[event].date);
    hour = date.getHours();
    minute = date.getMinutes();
    day = date.getDay();
    const away = obj.events[event].competitions[0].competitors.filter(x => x.homeAway === 'away')[0].team;
    const home = obj.events[event].competitions[0].competitors.filter(x => x.homeAway === 'home')[0].team;
    divisional = LEAGUE_DATA[home.abbreviation].division_opponents.indexOf(away.abbreviation) > -1 ? 1 : 0;
    division = divisional == 1 ? LEAGUE_DATA[home.abbreviation].division : '';
    let arr = [
      currentWeek,
      date,
      DAY[day].index,
      hour,
      minute,
      DAY[day].name,
      away.abbreviation,
      home.abbreviation,
      away.location,
      away.name,
      home.location,
      home.name,
      WEEKNAME.hasOwnProperty(currentWeek) ? WEEKNAME[currentWeek].name : 'Regular Season', // type
      divisional,
      division,
      (obj.events[event].competitions[0]).hasOwnProperty('odds') ? obj.events[event].competitions[0].odds[0].overUnder : '',
      (obj.events[event].competitions[0]).hasOwnProperty('odds') ? obj.events[event].competitions[0].odds[0].details : '',
      auto ? 1 : 0,
      timeFetched
    ];
    scoreboardData.push(arr);
  }
  for (let a = 0; a < scheduleData.length; a++) {
    if (scheduleData[a][0] == currentWeek) {
      scheduleData.splice(a,1,scoreboardData[0]);
      scoreboardData.shift();
    }
  }
  scheduleData.splice(scheduleData.indexOf(currentWeek),scoreboardData.length,...scoreboardData);

  let rows = scheduleData.length + 1;
  let columns = scheduleData[0].length;
  
  // utilities.gs functions to remove/add rows that are blank
  adjustRows(sheet,rows);
  adjustColumns(sheet,columns);
  
  let existingData = rangeData.getValues();
  const regexOverUnder = new RegExp(/^[0-9\.]+$/);
  const regexSpread = new RegExp(/^[A-Z]{2,3}\ \-[0-9\.]+$/);
  let existing = {};
  for (let a = 0; a < existingData.length; a++) {
    // Log data for each week (over/under and spread) as well as the schedule data for postseason weeks to recall later if needed
    if ((regexOverUnder.test(existingData[a][headers.indexOf('overUnder')]) || regexSpread.test(existingData[a][headers.indexOf('spread')])) || existingData[a][0] > REGULAR_SEASON) {
      let matchup = `${existingData[a][headers.indexOf('awayTeam')]}@${[existingData[a][headers.indexOf('homeTeam')]]}`;
      let rowData = existingData[a];
      existing[existingData[a][0]] = existing[existingData[a][0]] || {};
      existing[existingData[a][0]][matchup] = {};
      existing[existingData[a][0]][matchup].row = rowData;
      existing[existingData[a][0]][matchup].placed = false;
      if (existingData[a][headers.indexOf('overUnder')]) {
        existing[existingData[a][0]][matchup].overUnder = existingData[a][headers.indexOf('overUnder')];
      }
      if (existingData[a][headers.indexOf('spread')]) {
        existing[existingData[a][0]][matchup].spread = existingData[a][headers.indexOf('spread')];
      }
    }
  }

  // Checking for postseason empty slots within recently pulled data
  let missingMatchups = {};
  if (currentWeek > REGULAR_SEASON) {
    for (let a = 0; a < scheduleData.length; a++) {
      let scheduleDataWeek = scheduleData[a][0];
      if (scheduleDataWeek > REGULAR_SEASON) {
        if (scheduleData[a][headers.indexOf('awayTeam')] == '' || scheduleData[a][headers.indexOf('homeTeam') == '']) {
          missingMatchups[scheduleDataWeek] = missingMatchups[scheduleDataWeek] || {};
          missingMatchups[scheduleDataWeek].rows = missingMatchups[scheduleDataWeek].rows || [];
          missingMatchups[scheduleDataWeek].rows.push(a);
          missingMatchups[scheduleDataWeek].count = missingMatchups[scheduleDataWeek].count + 1 || 1;
        }
      }
    }
  }

  Object.keys(missingMatchups).forEach(week => {
    if (missingMatchups[week].count == matchupsPerWeek[week-1]) {
      Object.keys(existing[week]).forEach(matchup => {
        if (!existing[week][matchup].placed) {
          scheduleData[missingMatchups[week].rows[0]] = existing[week][matchup].row;
          existing[week][matchup].placed = true;
          missingMatchups[week].rows.splice(0,1);
        } else {
          Logger.log(`Already placed week ${week} matchup of ${matchup}.`);
        }
      });
    } else {
      let emptyRows = [];
      let knownMatchups = [];
      for (let a = 0; a < scheduleData.length; a++) {
        if (scheduleData[a][0] == week) {
          if (scheduleData[a][headers.indexOf('awayTeam')] != '' && scheduleData[a][headers.indexOf('homeTeam')] != '') {
            knownMatchups.push(scheduleData[a]);
          } else {
            emptyRows.push(a);
          }
        }
      }
      for (let a = 0; a < knownMatchups.length; a++) {
        if (existing[knownMatchups[a][0]].hasOwnProperty(`${knownMatchups[a][headers.indexOf('awayTeam')]}@${knownMatchups[a][headers.indexOf('homeTeam')]}`)) {
          existing[knownMatchups[a][0]][`${knownMatchups[a][headers.indexOf('awayTeam')]}@${knownMatchups[a][headers.indexOf('homeTeam')]}`].placed = true;
        }
      }
      Object.keys(existing[week]).forEach(matchup => {
        if (!existing[week][matchup].placed) {
          scheduleData.splice(emptyRows[0],1,existing[week][matchup].row);
          emptyRows.shift();
          existing[week][matchup].placed = true;
        }
      });
    }
  });
  for (let a = 0; a < scheduleData.length; a++ ) {
    let scheduleDataWeek = scheduleData[a][0];
    if (existing.hasOwnProperty(scheduleDataWeek)) {     
      if (existing[scheduleDataWeek].hasOwnProperty('row')) {
        Logger.log(`Replacing ${scheduleData[a]} with object data: ${existing[scheduleDataWeek].row}`);
        scheduleData.splice(a,1,existing[scheduleDataWeek].row);
      }
    }
  }

  if (Object.keys(existing).length > 0) {
    let awayIndex = headers.indexOf('awayTeam');
    let homeIndex = headers.indexOf('homeTeam');
    let spreadIndex = headers.indexOf('spread');
    let overUnderIndex = headers.indexOf('overUnder');
    let spreadAutoIndex = headers.indexOf('spreadAutoFetched');
    let timeFetchedIndex = headers.indexOf('timeFetched');
    for (let a = 0; a < scheduleData.length; a++) {
      let dataWeek = scheduleData[a][0];
      let matchup = `${scheduleData[a][awayIndex]}@${[scheduleData[a][homeIndex]]}`;
      if (dataWeek != currentWeek) {
        if (existing.hasOwnProperty(dataWeek)) {
          if (existing[dataWeek].hasOwnProperty(matchup)) {
            if (existing[dataWeek][matchup].hasOwnProperty('overUnder')) {
              scheduleData[a][overUnderIndex] = existing[dataWeek][matchup].overUnder;
            }
            if (existing[dataWeek][matchup].hasOwnProperty('spread')) {
              scheduleData[a][spreadIndex] = existing[dataWeek][matchup].spread;
            }
            scheduleData[a][spreadAutoIndex] = existing[dataWeek][matchup].auto;
            scheduleData[a][timeFetchedIndex] = existing[dataWeek][matchup].timeFetched;
          }
        }
      }
    }
    if (!overwrite && existing.hasOwnProperty(currentWeek)) {
      let ui = fetchUi();
      let replaceAlert = ui.alert(`Found previous over/under and spread data for week ${currentWeek} in the existing NFL data. Would you like to overwright with new values?`, ui.ButtonSet.YES_NO_CANCEL);
      if (replaceAlert !== ui.Button.YES) {
        for (let a = 0; a < scheduleData.length; a++) {
          let dataWeek = scheduleData[a][0];
          let matchup = `${scheduleData[a][awayIndex]}@${[scheduleData[a][homeIndex]]}`;
          if (dataWeek === currentWeek) {
            if (existing.hasOwnProperty(dataWeek)) {
              if (existing[dataWeek].hasOwnProperty(matchup)) {
                scheduleData[a][overUnderIndex] = existing[dataWeek][matchup].overUnder;
                scheduleData[a][spreadIndex] = existing[dataWeek][matchup].spread;
                scheduleData[a][spreadAutoIndex] = auto ? 1 : 0;
                scheduleData[a][timeFetchedIndex] = timeFetched;
              }
            }
          }
        }
      }
    }
  }
  rangeData.setValues(scheduleData);

  sheet.protect().setDescription(sheetName);
  try {
    sheet.hideSheet();
  }
  catch (err){
    // Logger.log('fetchSchedule hiding: Couldn\'t hide sheet as no other sheets exist');
  }
  ss.toast(`Imported all ${LEAGUE} schedule data`);
}

// NFL GAMES - output by week input and in array format: [date,day,hour,minute,dayName,awayTeam,homeTeam,awayTeamLocation,awayTeamName,homeTeamLocation,homeTeamName]
function fetchGames(week) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (week == null) {
    week = fetchWeek();
  }
  try {
    const nfl = ss.getRangeByName(league).getValues().shift();
    let games = [];
    for (let a = 0; a < nfl.length; a++) {
      if (nfl[a][0] == week) {
        games.push(nfl[a].slice(1));
      }
    }
    return games;
  }
  catch (err) {
    let text = 'Attempted to fetch NFL matches for week ' + week + ' but no NFL data exists, fetching now...';
    Logger.log(text);
    ss.toast(text);
    fetchSchedule();
    return fetchGames(week);
  }
}

// NFL ACTIVE WEEK SCORES - script to check and pull down any completed matches and record them to the weekly sheet
function recordWeeklyScores(){
  
  const outcomes = fetchWeeklyScores();
  if (outcomes[0] > 0) {
    const week = outcomes[0];
    const games = outcomes[1];
    const completed = outcomes[2];
    const remaining = outcomes[3];
    const data = outcomes[4];

    const done = (games == completed);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    const pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
    const survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
    const tiebreakerInclude = ss.getRangeByName('TIEBREAKER_PRESENT').getValue();
    let outcomesRecorded = [];
    let range;
    let alert = 'CANCEL';
    if (done) {
      let text = 'WEEK ' + week + ' COMPLETE\r\n\r\nMark all game outcomes';
      if (pickemsInclude) {
        text = text + ' and tiebreaker?';
      } else {
        text = text + '?';
      }
      alert = ui.alert(text, ui.ButtonSet.OK_CANCEL);
    } else if (remaining == 1) {
      alert = ui.alert('WEEK ' + week + ' INCOMPLETE\r\n\r\nRecord completed game outcomes?\r\n\r\n(There is one undecided game)\r\n\r\n', ui.ButtonSet.OK_CANCEL);
    } else if (remaining > 0 && remaining != games){
      alert = ui.alert('WEEK ' + week + ' INCOMPLETE\r\n\r\nRecord completed game outcomes?\r\n\r\n(There are ' + remaining + ' undecided games remaining)\r\n\r\n', ui.ButtonSet.OK_CANCEL);
    } else if (remaining == games) {
      ui.alert('WEEK ' + week + ' NOT YET STARTED\r\n\r\nNo game outcomes to record.\r\n\r\n', ui.ButtonSet.OK);
    }
    if (alert == 'OK') {
      if (pickemsInclude) {
        let sheet,matchupRange,matchups,cols,outcomeRange,outcomesRecorded,writeRange;
        try {
          sheet = ss.getSheetByName(weeklySheetPrefix+week);
          matchupRange = ss.getRangeByName(league + '_'+week);
          matchups = matchupRange.getValues().flat();
          outcomeRange = ss.getRangeByName(league + '_PICKEM_OUTCOMES_'+week);
          outcomesRecorded = outcomeRange.getValues().flat();
          if (tiebreakerInclude) {
            cols = matchups.length+1; // Adds one more column for tiebreaker value
          } else {
            cols = matchups.length;
          }
          writeRange = sheet.getRange(outcomeRange.getRow(),outcomeRange.getColumn(),1,cols);
        }
        catch (err) {
          Logger.log(err.stack);
          ss.toast('Issue with fetching weekly sheet or named ranges on weekly sheet, recreating now.');
          weeklySheet(week,memberList(ss),false);
        }
        let regex = new RegExp('[A-Z]{2,3}','g');
        let arr = [];
        for (let a = 0; a < matchups.length; a++){
          let game = matchups[a].match(regex);
          let away = game[0];
          let home = game[1];
          let outcome;        
          try {
            outcome = [];
            for (let b = 0; b < data.length; b++) {
              if (data[b][0] == away  && data[b][1] == home) {
                outcome = data[b];
              }
            }
            if (outcome.length <= 0) {
              throw new Error ('No game data for game at index ' + (a+1) + ' with teams given as ' + away + ' and ' + home);
            }
            //outcome = data.filter(game => game[0] == away && game[1] == home)[0];
            if (outcome[2] == away || outcome[2] == home) {
              if (regex.test(outcome[2])) {
                arr.push(outcome[2]);
              } else {
                arr.push(outcomesRecorded[a]);
              }
            } else if (outcome[2] == 'TIE') {
              let writeCell = sheet.getRange(outcomeRange.getRow(),outcomeRange.getColumn()+a);
              let rules = SpreadsheetApp.newDataValidation().requireValueInList([away,home,'TIE'], true).build();
              writeCell.setDataValidation(rules);
            } else {
              arr.push(outcomesRecorded[a]);
            }
          }
          catch (err) {
            Logger.log('No game data for ' + away + '@' + home);
            arr.push(outcomesRecorded[a]);
          }
          if (tiebreakerInclude) {
            try {
              if (a == (matchups.length - 1)) {
                if (outcome.length <= 0) {
                  throw new Error('No tiebreaker yet');
                }
                arr.push(outcome[3]); // Appends tiebreaker to end of array
              }
            }
            catch (err) {
              Logger.log('No tiebreaker yet');
              let tiebreakerCell = ss.getRangeByName(league + '_TIEBREAKER_'+week);
              let tiebreaker = sheet.getRange(tiebreakerCell.getRow()-1,tiebreakerCell.getColumn()).getValue();
              arr.push(tiebreaker);
            }
          }
        }
        writeRange.setValues([arr]);
      } else if (survivorInclude) {
        let away = ss.getRangeByName(league + '_AWAY_'+week).getValues().flat();
        let home = ss.getRangeByName(league + '_HOME_'+week).getValues().flat();
        range = ss.getRangeByName(league + '_OUTCOMES_'+week);
        outcomesRecorded = range.getValues().flat();
        let arr = [];
        for (let a = 0; a < away.length; a++) {
          arr.push([null]);
          for (let b = 0; b < data.length; b++) {
            if (data[b][0] == away[a] && data[b][1] == home[a]) {
              if (data[b][2] != null  && (outcomesRecorded[a] == null || outcomesRecorded[a] == '')) {
                arr[a] = [data[b][2]];  
              } else {
                arr[a] = [outcomesRecorded[a]];
              }
            }
          }        
        }
        range.setValues(arr);
      }
    }
    if (done) {  
      if (survivorInclude) {
        let prompt = ui.alert('WEEK ' + week + ' COMPLETE\r\n\r\nAdvance survivor pool?', ui.ButtonSet.YES_NO); 
        if ( prompt == 'YES' ) {
          ss.getRangeByName('WEEK').setValue(week+1);
        } else {
          ss.toast('Complete: '+ completed + ' game outcomes recorded');
        }
      } else {
        ss.toast('Complete: '+ completed + ' game outcomes recorded');
      }
    } else if ( alert != 'CANCEL') {
      ss.toast('Complete: '+ completed + ' game outcomes recorded');
    } else {
    ss.toast('Canceled import.');
    }
  }
}

// NFL OUTCOMES - Records the winner and combined tiebreaker for each matchup on the NFL sheet
function fetchWeeklyScores(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let obj = {};
  try{
    obj = JSON.parse(UrlFetchApp.fetch(scoreboard));
  }
  catch (err) {
    Logger.log(err.stack);
    ui.alert('ESPN API isn\'t responding currently, try again in a moment.',ui.ButtonSet.OK);
    throw new Error('ESPN API issue, try later');
  }
  
  
  if (Object.keys(obj).length > 0) {
    let games = obj.events;
    let week = obj.week.number;
    let year = obj.season['year'];

    // Checks if preseason, if not, pulls in score data
    if(obj.events[0].season.slug == 'preseason'){
      ui.alert('Regular season not yet started.\r\n\r\n Currently preseason is still underway.', ui.ButtonSet.OK);
      return [0,null,null,null,null];
    } else {

      let teams = [];

      // Get value for TNF being included
      let tnfInclude = true;
      try{
        tnfInclude = ss.getRangeByName('TNF_PRESENT').getValue();
      }
      catch (err) {
        Logger.log('Your version doesn\'t have the TNF feature configured, add a named range "TNF_PRESENT" "somewhere on a blank CONFIG sheet cell (hidden by default) with a value TRUE or FALSE to include');
      }

      // Get existing matchup data for comparison to scores (only for TNF exclusion)
      let data = [];
      if (!tnfInclude) {
        try {
          data = ss.getRangeByName(league).getValues().shift();
        }
        catch (err) {
          ss.toast('No NFL data, importing now');
          fetchSchedule(year);
          data = ss.getRangeByName(league).getValues().shift();
        }
        for (let a = 0; a < data.length; a++) {
          if (data[a][0] == week && (tnfInclude || (!tnfInclude && data[a][2] >= 0))) {
            teams.push(data[a][6]);
            teams.push(data[a][7]);
          }
        }
      }
      // Loop through games provided and creates an array for placing
      let all = [];
      let count = 0;
      let away, awayScore,home, homeScore,tiebreaker,winner,competitors;
      for (let a = 0; a < games.length; a++){
        let outcomes = [];
        awayScore = '';
        homeScore = '';
        tiebreaker = '';
        winner = '';
        competitors = games[a].competitions[0].competitors;
        away = (competitors[1].homeAway == 'away' ? competitors[1].team.abbreviation : competitors[0].team.abbreviation);
        home = (competitors[0].homeAway == 'home' ? competitors[0].team.abbreviation : competitors[1].team.abbreviation);
        if (games[a].status.type.completed) {
          if (tnfInclude || (!tnfInclude && (teams.indexOf(away) >= 0 || teams.indexOf(home) >= 0))) {
            count++;
            awayScore = parseInt(competitors[1].homeAway == 'away' ? competitors[1].score : competitors[0].score);
            homeScore = parseInt(competitors[0].homeAway == 'home' ? competitors[0].score : competitors[1].score);
            tiebreaker = awayScore + homeScore;
            winner = (competitors[0].winner ? competitors[0].team.abbreviation : (competitors[1].winner ? competitors[1].team.abbreviation : 'TIE'));
            outcomes.push(away,home,winner,tiebreaker);
            all.push(outcomes);
          }
        }      
      }
      // Sets info variables for passing back to any calling functions
      let remaining = games.length - count;
      let completed = games.length - remaining;

      // Outputs total matches, how many completed, and how many remaining, and all matchups with outcomes decided;
      return [week,games.length,completed,remaining,all];
    }
  } else {
    Logger.log('ESPN API returned no games');
    ui.alert('ESPN API didn\'t return any game information. Try again later and make sure you\'re checking while the season is active',ui.ButtonSet.OK);
  }
}

// LEAGUE LOGOS - Saves URLs to logos to a Script Property variable named "logos"
function fetchLogos(){
  let obj = {};
  let logos = {};
  try{
    obj = JSON.parse(UrlFetchApp.fetch(scoreboard));
  }
  catch (err) {
    Logger.log(err.stack);
    ui.alert('ESPN API isn\'t responding currently, try again in a moment.',ui.ButtonSet.OK);
    throw new Error('ESPN API issue, try later');
  }
  
  if (Object.keys(obj).length > 0) {
    let games = obj.events;
    // Loop through games provided and creates an array for placing
    for (let a = 0; a < games.length; a++){
      let competitors = games[a].competitions[0].competitors;
      let teamOne = competitors[0].team.abbreviation;
      let teamTwo = competitors[1].team.abbreviation;
      let teamOneLogo = competitors[0].team.logo;
      let teamTwoLogo = competitors[1].team.logo;
      logos[teamOne] = teamOneLogo;
      logos[teamTwo] = teamTwoLogo;
    }
    Logger.log(logos);
    const scriptProperties = PropertiesService.getScriptProperties();
    try {
      let logoProp = scriptProperties.getProperty('logos');
      let tempObj = JSON.parse(logoProp);
      if (Object.keys(tempObj).length < nflTeams) {
        scriptProperties.setProperty('logos',JSON.stringify(logos));
      }
    }
    catch (err) {
      Logger.log('Error fetching logo object, creating one now');
      scriptProperties.setProperty('logos',JSON.stringify(logos));
    }
  }
  return logos;
}
