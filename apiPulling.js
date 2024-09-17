
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
    Logger.log(name + ' is currently active with ' + weeks + ' weeks in total, current week is: ' + week); 
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

// FETCH TOTAL WEEKS
function fetchWeeks() {
  try {
    let weeks;
    const content = UrlFetchApp.fetch(scoreboard).getContentText();
    const obj = JSON.parse(content);
    const calendar = obj.leagues[0].calendar;
    for (let a = 0; a < calendar.length; a++) {
      if (calendar[a].value == 2) {
        weeks = calendar[a].entries.length;
        break;
      }
    }
    return weeks;
  }
  catch (err) {
    Logger.log('ESPN API has an issue right now');
    return 18;
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

// NFL TEAM INFO - script to fetch all NFL data for teams
function fetchSchedule(year) {
  // Calls the linked spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Declaration of script variables
  let abbr, name, maxRows, maxCols;
  const objTeams = fetchTeamsESPN(year);
  const teamsLen = objTeams.length;
  let arr = [];
  let nfl = [];
  let espnId = [];
  let espnAbbr = [];
  let espnName = [];
  let espnLocation = [];
  let location = [];  
  
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
      arr = [objTeams[a].id,abbr,location,name,objTeams[a].byeWeek];
      nfl.push(arr);
    }
  }
  
  let sheet, range;
  let ids = [];
  let abbrs = [];
  for (let a = 0 ; a < espnId.length ; a++ ) {
    ids.push(espnId[a].toFixed(0));
    abbrs.push(espnAbbr[a]);
  }
  // Declaration of variables
  let schedule = [];
  let home = [];
  let dates = [];
  let allDates = [];
  let hours = [];
  let allHours = [];
  let minutes = [];
  let allMinutes = [];
  let byeIndex, id;
  let date, hour, minute;
  let weeks = Object.keys(objTeams[0].proGamesByScoringPeriod).length;
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
  let week, awayTeam, awayTeamName, awayTeamLocation, homeTeam, homeTeamName, homeTeamLocation, mnf, day, dayName;
  let formData = [];
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
        weekArr.push(week);
      }
    }
  }
  let headers = ['week','date','day','hour','minute','dayName','awayTeam','homeTeam','awayTeamLocation','awayTeamName','homeTeamLocation','homeTeamName'];
  let sheetName = league;
  let rows = formData.length + 1;
  let columns = formData[0].length;
  
  sheet = ss.getActiveSheet();
  if ( sheet.getSheetName() == 'Sheet1' && ss.getSheetByName(sheetName) == null) {
    sheet.setName(sheetName);
  }
  sheet = ss.getSheetByName(sheetName);  
  if (sheet == null) {
    ss.insertSheet(sheetName,0);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.setTabColor(scheduleTabColor);
  
  maxRows = sheet.getMaxRows();
  if (maxRows < rows){
    sheet.insertRows(maxRows,rows - maxRows - 1);
  } else if (maxRows > rows){
    sheet.deleteRows(rows,maxRows - rows);
  }
  maxCols = sheet.getMaxColumns();
  if (maxCols < columns) {
    sheet.insertColumnsAfter(maxCols,columns - maxCols);
  } else if (maxCols > columns){
    sheet.deleteColumns(columns,maxCols - columns);
  }
  sheet.setColumnWidths(1,columns,30);
  sheet.setColumnWidth(2,60);
  sheet.setColumnWidth(6,60);
  sheet.setColumnWidths(9,4,80);
  sheet.clear();
  range = sheet.getRange(1,1,1,columns);
  range.setValues([headers]);
  ss.setNamedRange(sheetName+'_HEADERS',range);
 
  range = sheet.getRange(1,1,rows,columns);
  range.setFontSize(8);
  range.setVerticalAlignment('middle');  
  range = sheet.getRange(2,1,formData.length,columns);
  range.setValues(formData);

  ss.setNamedRange(sheetName,range);
  range.setHorizontalAlignment('left');
  range.sort([{column: 1, ascending: true},{column: 2, ascending: true},{column: 4, ascending: true},
              {column:  5, ascending: true},{column: 6, ascending: true},{column: 8, ascending: true}]); 
  sheet.getRange(1,3).setNote('-3: Thursday, -2: Friday, -1: Saturday, 0: Sunday, 1: Monday, 2: Tuesday');
  
  // Fetches sorted data
  formData = range.getValues();
  weekArr = sheet.getRange(2,1,rows-1,1).getValues().flat();
  // Sets named ranges for weekly home and away teams to compare for survivor status
  awayTeam = headers.indexOf('awayTeam')+1;
  homeTeam = headers.indexOf('homeTeam')+1;
  for (let a = 1; a <= weeks; a++) {
    let start = weekArr.indexOf(a)+2;
    let end = weekArr.indexOf(a+1)+2;
    if (a == weeks) {
      end = rows+1;
    }
    let len = end - start;
    ss.setNamedRange(league + '_AWAY_'+a,sheet.getRange(start,awayTeam,len,1));
    ss.setNamedRange(league + '_HOME_'+a,sheet.getRange(start,homeTeam,len,1));
  }
  sheet.protect().setDescription(sheetName);
  try {
    sheet.hideSheet();
  }
  catch (err){
    // Logger.log('fetchSchedule hiding: Couldn\'t hide sheet as no other sheets exist');
  }
  ss.toast('Imported all NFL schedule data');
}

// NFL GAMES - output by week input and in array format: [date,day,hour,minute,dayName,awayTeam,homeTeam,awayTeamLocation,awayTeamName,homeTeamLocation,homeTeamName]
function fetchGames(week) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (week == null) {
    week = fetchWeek();
  }
  try {
    const nfl = ss.getRangeByName(league).getValues();
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
          data = ss.getRangeByName(league).getValues();
        }
        catch (err) {
          ss.toast('No NFL data, importing now');
          fetchSchedule(year);
          data = ss.getRangeByName(league).getValues();
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
