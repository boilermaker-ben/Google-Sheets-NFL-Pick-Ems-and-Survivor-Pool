// BONUS TOOLS
//------------------------------------------------------------------------
// BONUS STATE - writes bonus state, reveals or hides bonus row of current week, and adds named range if missing
function bonusState(bonus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (bonus == null) {
    try {
      bonus = ss.getRangeByName('BONUS_PRESENT').getValue();
    }
    catch (err) {
      bonus = false;
      ss.toast('No bonus state provided to bonusState function and no existing value in sheet, setting to \'false\'');
      Logger.log('No bonus state provided to bonusState function and no existing value in sheet, setting to \'false\'');
    }
  }
  try{
    let bonusCell = ss.getRangeByName('BONUS_PRESENT');
    bonusCell.setValue(bonus);
  }
  catch (err) {
    Logger.log('Creating a \'BONUS_PRESENT\' value on the CONFIG page');
    let labelRange = ss.getSheetByName('CONFIG').getRange(ss.getRangeByName('PICKEMS_PRESENT').getRow(),ss.getRangeByName('PICKEMS_PRESENT').getColumn()+1);
    let valueRange = ss.getSheetByName('CONFIG').getRange(labelRange.getRow(),labelRange.getColumn()+1);
    labelRange.setValue('BONUS GAMES')
      .setFontWeight('bold');
    valueRange.setValue(bonus);
    ss.setNamedRange('BONUS_PRESENT',valueRange);
  }
  const week = fetchWeek();
  try {
    let range = ss.getRangeByName(league + '_BONUS_' + week);
    if (bonus) {
      range.getSheet().showRows(range.getRow());
      ss.toast('Bonus row for week ' + week + ' is now visible');
      Logger.log('Bonus row for week ' + week + ' is now visible');
    } else {
      let range = ss.getRangeByName(league + '_BONUS_' + week);
      range.getSheet().hideRows(range.getRow());
      ss.toast('Bonus row for week ' + week + ' is now hidden');
      Logger.log('Bonus row for week ' + week + ' is now hidden');
    }
  }
  catch (err) {
    if (bonus) {
      ss.toast('No bonus row exists for week ' + week + ' future weeks will have a bonus row present. Update the \'WEEK\' value on the \'CONFIG\' sheet if you intended to reveal the bonus row on another week');
      Logger.log('The week '+ week + ' sheet doesn\'t have a bonus row to reveal, future weeks will have a bonus row present');
    } else {
      ss.toast('No bonus row exists for week ' + week + ' future weeks will have a bonus row hidden. Update the \'WEEK\' value on the \'CONFIG\' sheet if you intended to hide the bonus row on another week');
      Logger.log('The week '+ week + ' sheet doesn\'t have a bonus row to hide, future weeks will have a bonus row hidden');
    }
  }
  createMenu(null,true);
}

// BONUS STATE TRUE - calls bonus state function to write value as "TRUE"
function bonusUnhide() {
  bonusState(true);
}

// BONUS STATE FALSE - calls bonus state function to write value as "FALSE"
function bonusHide() {
  bonusState(false);
}

// DOUBLE MNF STATE - writes double MNF state, changes bonus row value if present, and adds named range if missing
function bonusDoubleMNF(double) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let valueRange, mnfRange, text, arr = [];
  if (double == null) {
    try {
      double = ss.getRangeByName('MNF_DOUBLE').getValue();
    }
    catch (err) {
      double = false;
      ss.toast('No double MNF state provided to bonusDoubleMNF function and no existing value in sheet, setting to \'false\'');
      Logger.log('No double MNF state provided to bonusDoubleMNF function and no existing value in sheet, setting to \'false\'');
    }
  }
  try{
    let doubleCell = ss.getRangeByName('MNF_DOUBLE');
    doubleCell.setValue(double);
  }
  catch (err) {
    Logger.log('Creating a \'MNF_DOUBLE\' value on the CONFIG page');
    let labelRange = ss.getSheetByName('CONFIG').getRange(ss.getRangeByName('MNF_PRESENT').getRow(),ss.getRangeByName('MNF_PRESENT').getColumn()+1);
    valueRange = ss.getSheetByName('CONFIG').getRange(labelRange.getRow(),labelRange.getColumn()+1);
    labelRange.setValue('MNF DOUBLE')
      .setFontWeight('bold');
    valueRange.setValue(double);
    ss.setNamedRange('MNF_DOUBLE',valueRange);
  }
  let week = maxWeek();
  try {
    mnfRange = ss.getRangeByName(league + '_MNF_' + week);
    for (let a = 0; a < mnfRange.getNumColumns(); a++) {
      if (double) { 
        arr.push(2);
      } else {
        arr.push(1);
      }
    }
  }
  catch (err) {
    text = 'MNF ERROR\r\n\r\nNo MNF games exist for week ' + week + ' or there was an error finding the MNF named range for the week. Future weeks will include MNF games marked as ';
    if (double) {
      text = text.concat('double.');
    } else {
      text = text.concat('as a normal game.');
    }
    ui.alert(text,ui.ButtonSet.OK);
    Logger.log(text);
  }
  try {
    let bonusRange = ss.getRangeByName(league + '_BONUS_' + week);
    let doubleMNFRange = bonusRange.getSheet().getRange(bonusRange.getRow(),mnfRange.getColumn(),1,mnfRange.getNumColumns());
    let notifyText = 'MNF DOUBLE\r\n\r\nWould you like to mark this week\'s ';
    if (arr.length > 1) {
      notifyText = notifyText.concat(arr.length + ' MNF games as ');
    } else {
      notifyText = notifyText.concat(' MNF game as ');
    }
    if (double) {
      notifyText = notifyText.concat('double for week ' + week + ' and future weeks?');
    } else {
      notifyText = notifyText.concat('a normal game for week ' + week + ' and also count future MNF games as normal games?');
    }
    let notify = ui.alert(notifyText,ui.ButtonSet.YES_NO);
    if (notify == ui.Button.YES) {
      doubleMNFRange.setValues([arr]);
      text = 'The ';
      if (arr.length > 1) {
        text = text.concat(arr.length + ' MNF games for week ' + week + ' were ');
      } else {
        text = text.concat('MNF game for week ' + week + ' was ');
      }
      text = text.concat('marked to be weighted as ');
      if (double) { 
        text = text.concat('double.');
      } else {
        text = text.concat('a normal game.');
      }
      ss.toast(text);
      Logger.log(text);
    } else {
      double = !double;
      valueRange.setValue(double);
      Logger.log('Canceled MNF double operation and reset the value to \'' + double + '\'');
    }
  }
  catch (err) {
    Logger.log(err.stack);
    let text = 'No bonus row exists for week ' + week + ' future weeks will have a bonus row present and will mark MNF as ';
    if (double) {
      text = text.concat('double.');
    } else {
      text = text.concat('a standard game.');
    }
    ss.toast(text);
    Logger.log(text);
  }
  createMenu(null,true);
}

// DOUBLE MNF STATE ENABLE - uses bonusDoubleMNF with true value for menu
function bonusDoubleMNFEnable() {
  bonusDoubleMNF(true);
}

// DOUBLE MNF STATE ENABLE - uses bonusDoubleMNF with false value for menu
function bonusDoubleMNFDisable() {
  bonusDoubleMNF(false);
}

// GAME OF THE WEEK SHEET FUNCTION - selects one random game for 2x multiplier to be applied
function bonusRandomGameSet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let tnf = true, bonusRange, mnfDouble = false, text;
  const week = maxWeek();
  let sheet, sheetExisted = true;
  try { 
    ss.getRangeByName('BONUS_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('No \'BONUS_PRESENT\' named range');
    ui.alert('BONUS PRESENT NOT SET\r\n\r\nNo bonus present range established for inclusion/exclusion of bonus game weighting, please run the enable/disable bonus function and try this function again after that has been set', ui.ButtonSet.OK);
    throw new Error('Canceled due to no bonus feature');
  }
  try { 
    tnf = ss.getRangeByName('TNF_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('No \'TNF_PRESENT\' named range, assuming true');
    ui.alert('THURSDAY NIGHT FOOTBALL EXCLUSION NOT SET\r\n\r\nNo Thursday present range established for inclusion/exclusion of Thursday\'s games', ui.ButtonSet.OK);
  }
  try { 
    mnfDouble = ss.getRangeByName('MNF_DOUBLE').getValue();
  }
  catch (err) {
    Logger.log('No \'MNF_DOUBLE\' named range, assuming \'false\' for double MNF and proceeding.');
  }
  try {
    sheet = ss.getSheetByName(weeklySheetPrefix + week);
  } catch (err) {
    Logger.log('No sheet for week ' + week);
    let prompt = ui.alert('NO SHEET\r\n\r\nThe week ' + week + ' sheet does not exist. Create a sheet now for week ' + week + '?\r\n\r\n(Selecting \'Cancel\' will exit and no game will be selected)', ui.ButtonSet.OK_CANCEL);
    if (prompt = ui.Button.OK) {
      sheet = weeklySheet(ss,week,memberList(ss),false);
      sheetExisted = false;
    } else {
      throw new Error('Exited when new sheet creation was declined');
    }
  }
  try {
    bonusRange = ss.getRangeByName(league + '_BONUS_' + week);
  }
  catch (err) {
    Logger.log('No \'BONUS\' named range for week ' + week);
    ui.alert('NO BONUS\r\n\r\nThe week ' + week + ' sheet lacks the bonus game feature. Would you like to recreate the week ' + week + ' sheet now?\r\n\r\n(Selecting \'Cancel\' will exit and no game will be selected)', ui.ButtonSet.OK_CANCEL);
    if (prompt == ui.Button.OK) {
      sheet = weeklySheet(ss,week,memberList(ss),false);
      sheetExisted = false;
    } else {
      throw new Error('Exited when new sheet creation was declined');
    }
  }
  bonusRange = ss.getRangeByName(league + '_BONUS_' + week);

  let mnf = false, mnfRange, bonusValues = bonusRange.getValues().flat();

  if (mnfDouble) {
    try {
      mnfRange = ss.getRangeByName(league + '_MNF_' + week);
      bonusValues.splice(bonusValues.length-mnfRange.getNumColumns(),mnfRange.getNumColumns());
      bonusRange = sheet.getRange(bonusRange.getRow(),bonusRange.getColumn(),1,bonusRange.getNumColumns()-mnfRange.getNumColumns());
      if (mnfRange.getValues().length > 0) {
        mnf = true;
      }
    }
    catch (err) {
      Logger.log('No MNF range for week ' + week + '. Including all games in randomization.');
    }
  }  

  if (sheetExisted) {
    for (let a = 0; a < bonusValues.length; a++) {
      if (bonusValues[a] > 1) {
        text = 'BONUS GAME ALREADY MARKED\r\n\r\nYou already have one or more games marked for 2x or greater weighting.\r\n\r\nMark all ';
        if (mnfDouble && mnf) {
          text = text.concat('non-MNF games\' weighting to 1 and try again');
        } else {
          text = text.concat('games\' weighting to 1 and try again');
        }
        ui.alert(text,ui.ButtonSet.OK);
        throw new Error('Other games marked as bonus prior to running random Game of the Week function');
      }
    }
  }

  text = 'GAME OF THE WEEK\r\n\r\nWould you like to randomly select one game this week to count as double?';
  if (mnfDouble) {
    text = text.concat('\r\n\r\nAny MNF games will be excluded since you have the MNF Double feature enabled');
  }
  let gameOfTheWeek;
  let randomPrompt = ui.alert(text,ui.ButtonSet.YES_NO);
  if (randomPrompt == ui.Button.YES) {
    gameOfTheWeek = bonusRandomGame(week,tnf,mnfDouble);
    let matchupNames = ss.getRangeByName(league + '_' + week).getValues().flat();
    let regex = new RegExp(/[A-Z]{2,3}/,'g');
    let matchupRegex = [];
    matchupNames.forEach(a => matchupRegex.push(a.match(regex)[0]+ '@' + a.match(regex)[1]));
    bonusValues[matchupRegex.indexOf(gameOfTheWeek)] = 2;
    bonusRange.setValues([bonusValues]);
  }

  let formId = ss.getRangeByName('FORM_WEEK_' + week).getValue();
  try {
    let form = FormApp.openById(formId);
    let prompt = ui.alert('FORM EXISTS\r\n\r\nYou\'ve already created a form for week ' + week + ', would you like to designate the Game of the Week on the Form?',ui.ButtonSet.YES_NO);
    if (prompt == ui.Button.YES) {
      let form = FormApp.openById(formId);
      let questions = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
      for (let a = 0; a < questions.length; a++) {
        try{
          let choices = questions[a].asMultipleChoiceItem().getChoices();
          let matchup = choices[0].getValue() + '@' + choices[1].getValue();
          if (matchup == gameOfTheWeek) {
            questions[a].setTitle('GAME OF THE WEEK (Double Points)\n' + questions[a].getTitle());
            break;
          }
        }
        catch (err) {
          Logger.log('Issue with getting choices for question with title ' + questions[a].getTitle() + ' or setting the title.');
        }
      }
    }
  }
  catch (err) {
    Logger.log('No form exists for week ' + week + ' or there was an error getting the questions for the form.'); 
  }
}

// GAME OF THE WEEK SELECTION - selects one random game for 2x multiplier to be applied
function bonusRandomGame(week,tnf,mnfDouble) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (week == null) {
    week = maxWeek();
  }

  let games = fetchGames(week);
  
  let abbrevs = [];
  let teams = [];
  for (let a = games.length - 1; a >= 0; a--) {
    if ((games[a][1] == 1 && mnfDouble) || (games[a][1] == -3 && !tnf)) {
      games.splice(a,1);
    } else {
      abbrevs.push(games[a][5] + '@' + games[a][6])
      teams.push(games[a][7] + ' ' + games[a][8] + ' at ' + games[a][9] + ' ' + games[a][10])
    }
  }

  let gameOfTheWeekIndex = getRandomInt(0,abbrevs.length-1);

  text = 'For week ' + week + ', your Game of the Week has been randomly selected as:\r\n\r\n';
  try {
    let gameOfTheWeek = abbrevs[gameOfTheWeekIndex];
    text = text.concat(teams[gameOfTheWeekIndex] + '\r\n\r\nWould you like to mark it as such?');
    let verify = ui.alert(text,ui.ButtonSet.OK_CANCEL);
    if (verify == ui.Button.OK) {
      return gameOfTheWeek;
    } else {
      ss.toast('Canceled Game of the Week selection');
    }
  }
  catch (err) {
    ss.toast('Error fetching matches or selecting Game of the Week\r\n\r\nError:\r\n' + err.message);
    Logger.log('Error fetching matches or selecting Game of the Week\r\n\r\nError:\r\n' + err.message);
  }
}

// RANDOM - random integer function for selecting Game of the Week
function getRandomInt(min, max) {
      min = Math.ceil(min);
      max = Math.floor(max);
      return Math.floor(Math.random() * (max - min + 1)) + min;
}
