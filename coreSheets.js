// SHEET CREATION
//------------------------------------------------------------------------
// SHEET FOR LOGGING ALL OUTCOMES - creates a set of columns (one per week) on a sheet with a dedicated data validation rule per game to select from if not using import features
function outcomesSheet(ss) {
  ss = fetchSpreadsheet(ss);
  const weeks = fetchWeeks();
  const sheetName = league+'_OUTCOMES';
  let sheet = ss.getSheetByName(sheetName);
  if ( sheet == null ) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clearFormats();
  sheet.setTabColor(dayColorsFilled[dayColorsFilled.length-1]);

  let data;
  try {
    data = ss.getRangeByName(league).getValues();
  }
  catch (err) {
    ss.toast('No ' + league + ' data, importing now');
    fetchSchedule();
    data = ss.getRangeByName(league).getValues();
  }

  let tnfInclude = true;
  try{
    tnfInclude = ss.getRangeByName('TNF_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('Your version doesn\'t have the TNF feature configured, add a named range "TNF_PRESENT" "somewhere on a blank CONFIG sheet cell (hidden by default) with a value TRUE or FALSE to include');
  }

  let headers = [];
  let headersWidth = [];
  let headerRow = 2;

  for (let a = 1; a <= weeks; a++) {
    headers.push(a);
    headersWidth.push(60);
    ss.setNamedRange(league + '_OUTCOMES_' + a,sheet.getRange(headerRow+1,a,maxGames,1)); // maxGames is a global variable
  }

  // Adjust the rows and columns of the sheet, and set maxCols/maxRows variables
  let maxCols = sheet.getMaxColumns();
  if (maxCols < headers.length) {
    sheet.insertColumnsAfter(maxCols,headers.length-maxCols);
  } else if (maxCols > headers.length) {
    sheet.deleteColumns(headers.length + 1, maxCols - headers.length);
  }
  maxCols = sheet.getMaxColumns();

  let rowTarget = (headerRow + maxGames); // maxGames is a global variable
  let maxRows = sheet.getMaxRows();
  if (maxRows < rowTarget) {
    sheet.insertRowsAfter(maxRows,rowTarget - maxRows);
  } else if (maxRows > rowTarget) {
    sheet.deleteRows(rowTarget + 1, maxRows - rowTarget);
  }
  maxRows = sheet.getMaxRows();

  sheet.getRange(1,1,maxRows,maxCols).clearDataValidations();
  sheet.getRange(1,1,maxRows,maxCols).clearNote();

  // Formatting sheet
  let range = sheet.getRange(1,1);
  sheet.setRowHeight(1,70);
  range.setValue(sheetName.replace(/\_/g,' '));
  range.setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontColor('#ffffff')
    .setBackground('#666666')
    .setFontSize(18)
    .setFontFamily("Montserrat")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  range.setNote('Home team will be BOLD when selected. Colors represent days of the week the game is played on (yellow is Thursday moving to blue for Monday)');
  sheet.getRange(1,1,1,maxCols).mergeAcross(); // Merges top row horizontally
  range = sheet.getRange(headerRow,1,1,headers.length);
  sheet.setRowHeight(1,35);
  range.setValues([headers]);
  range.setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontColor('#ffffff')
    .setBackground('#000000')
    .setFontSize(12)
    .setFontFamily("Montserrat")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  range = sheet.getRange(headerRow+1,1,maxRows-headerRow,maxCols);
  range.setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontColor('#000000')
    .setBackground('#ffffff')
    .setFontSize(9)
    .setFontFamily("Montserrat")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  for (let a = 0; a < headersWidth.length; a++) {
    sheet.setColumnWidth(a+1,headersWidth[a]);
  }

  let matchups = sheet.getRange(headerRow+1,1,maxRows-headerRow,maxCols);
  matchups.setBackground('#dddddd');

  let game = 1;
  let week;
  let formats = [];
  for (let row = 0; row < data.length; row++) {
    if (tnfInclude || (!tnfInclude && data[row][2] >= 0)) {
      if (data[row][0] != week) { // Checks if new row has a new week value
        game = 1;
      } else {
        game++;
      }
      week = data[row][0]; // Sets week variable to the week stated in the data row
      let writeCell = sheet.getRange(game+headerRow,week);
      let rules = SpreadsheetApp.newDataValidation().requireValueInList([data[row][6],data[row][7],'TIE'], true).build();
      writeCell.setDataValidation(rules);
      let awayWin = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(data[row][6])
        .setBold(false)
        .setRanges([writeCell]);
      let homeWin = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(data[row][7])
        .setBold(true)
        .setRanges([writeCell]);
      // Color Coding Days
      let dayIndex = data[row][2] + 3; // Numeric day used for gradient application (-3 is Thursday, 1 is Monday);
      writeCell.setBackground(dayColors[dayIndex]);
      awayWin.setBackground(dayColorsFilled[dayIndex]);
      homeWin.setBackground(dayColorsFilled[dayIndex]);
      awayWin.build();
      homeWin.build();
      formats.push(awayWin);
      formats.push(homeWin);
    }
  }
  let ties = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('TIE')
    .setBold(false)
    .setBackground('#aaaaaa')
    .setRanges([matchups])
    .build();
  formats.push(ties);
  sheet.setConditionalFormatRules(formats);
  Logger.log('Completed setting up ' + league + ' Winners sheet');
}

// UPDATE OUTCOMES - Updates the data validation, color scheme, and matchups for a specific week on the winners sheet
function outcomesSheetUpdate(ss,week,equations) {
  const startRow = 3;

  ss = fetchSpreadsheet(ss);
  if (week == null) {
    week = fetchWeek();
  }
  let sheet = ss.getSheetByName(league + '_OUTCOMES');
  if (sheet == null) {
    sheet = outcomesSheet(ss);
  }

  let data = ss.getRangeByName(league).getValues();
  if (data == null) {
    fetchSchedule();
  }
  let tnfInclude = true;
  try{
    tnfInclude = ss.getRangeByName('TNF_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('Your version doesn\'t have the TNF feature configured, add a named range "TNF_PRESENT" somewhere on a blank CONFIG sheet cell (hidden by default) with a value TRUE or FALSE to include');
  }

  let days = [], games = [];
  for (let a = 0; a < data.length; a++) {
    if (data[a][0] == week && ((tnfInclude && data[a][2] == -3) || data[a][2] != -3)) {
      days.push(data[a][2]+3); // Numeric day used for gradient application (-3 is Thursday, 1 is Monday);
      games.push([data[a][6],data[a][7]]);
    }
  }
  if (equations != true) {

    // Clears data validation and notes
    let matchups = ss.getRangeByName(league + '_OUTCOMES_' + week);
    matchups.clearDataValidations();
    matchups.clearNote();

    let existingRules = sheet.getConditionalFormatRules();
    let rulesToKeep = [];
    let newRules = [];
    for (let a = 0; a < existingRules.length; a++) {
      let ranges = existingRules[a].getRanges();
      for (let b = 0; b < ranges.length; b++) {
        if (ranges[b].getColumn() != matchups.getColumn()) {
          rulesToKeep.push(existingRules[a]);
        }
      }
    }

    let start = startRow;
    let end = start+1;
    for (let a = 0; a < days.length; a++) {
      sheet.getRange(a+startRow,week).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([games[a][0],games[a][1],'TIE'], true).build());
      // Color Coding Days
      if (days[a] != days[a+1]) {
        sheet.getRange(start,week,end-start,1).setBackground(dayColors[days[a]]);
        let homeWin = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=iferror(match(indirect(\"R[0]C[0]\",false),indirect(\"'+league+'_HOME_'+week+'\"),0)>=0,false)')
          .setBackground(dayColorsFilled[days[a]])
          .setBold(true)
          .setRanges([sheet.getRange(start,week,end-start,1)])
          .build();
        newRules.push(homeWin);
        let awayWin = SpreadsheetApp.newConditionalFormatRule()
          .whenCellNotEmpty()
          .setBackground(dayColorsFilled[days[a]])
          .setRanges([sheet.getRange(start,week,end-start,1)])
          .build();
        newRules.push(awayWin);
        start = end;
      }
      end++;
    }
    let ties = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('TIE')
      .setBold(false)
      .setBackground('#BBBBBB')
      .setRanges([sheet.getRange(startRow,week,days.length,1)])
      .build();
    newRules.push(ties);

    let allRules = rulesToKeep.concat(newRules);
    //clear all rules first and then add again
    sheet.clearConditionalFormatRules();
    sheet.setConditionalFormatRules(allRules);
  }

  let weeklySheetName = (weeklySheetPrefix + week);
  let sourceSheet = ss.getSheetByName(weeklySheetName);
  if (sourceSheet != null && ss.getRangeByName('PICKEMS_PRESENT').getValue()) {
    const targetSheet = ss.getSheetByName(league + '_OUTCOMES');
    const sourceRange = ss.getRangeByName(league + '_PICKEM_OUTCOMES_'+week);
    const targetRange = ss.getRangeByName(league + '_OUTCOMES_'+week);
    let row = sourceRange.getRow();
    let data = targetRange.getValues().flat();
    let regex = new RegExp(/^[A-Z]{2,3}/);
    for (let a = 1; a <= sourceRange.getNumColumns(); a++) {
      if (!regex.test(data[a-1])) {
        targetSheet.getRange(targetRange.getRow()+(a-1),targetRange.getColumn()).setFormula(
          '=\''+weeklySheetName+'\'!'+sourceSheet.getRange(row,sourceRange.getColumn()+(a-1)).getA1Notation()
        );
      } else {
        Logger.log('Found matching value of ' + data[a-1] + ' on outcomes sheet in row ' + (a + 2) + '; avoiding re-writing formula for this cell');
      }
    }
  } else if (sourceSheet == null) {
    Logger.log('No sheet created yet for week ' + week);
    const ui = SpreadsheetApp.getUi();
    let prompt = ui.alert('No sheet created yet for week ' + week + '.\r\n\r\nWould you like to create a weekly sheet now?',ui.ButtonSet.OK_CANCEL);
    if (prompt == ui.Button.OK) {
      weeklySheetCreate(week,false);
    }
  } else {
    Logger.log('Pick \'Ems not present for running the formula portion of Outcomes Update script');
  }
}

// CONFIG SHEET - Sheet with all recorded customizations as well as logging the URLs for the weekly forms created by the script
function configSheet(ss,name,year,week,weeks,pickemsInclude,mnfInclude,tnfInclude,tiebreaker,commentInclude,bonus,survivorInclude,survivorStart) {
  ss = fetchSpreadsheet(ss);
  let sheetName = 'CONFIG';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName,0);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.setTabColor(configTabColor);

  try {
    if (pickemsInclude == null) {
      pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
    }
    if (mnfInclude == null) {
      mnfInclude = ss.getRangeByName('MNF_PRESENT').getValue();
    }
    if (tnfInclude == null) {
      tnfInclude = ss.getRangeByName('TNF_PRESENT').getValue();
    }
    if (tiebreaker == null) {
      tiebreaker = ss.getRangeByName('TIEBREAKER_PRESENT').getValue();
    }
    if (bonus == null) {
      bonus = ss.getRangeByName('BONUS_PRESENT').getValue();
    }
    if (survivorInclude == null) {
      survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
    }
    if (survivorInclude) {
      if (week == 1) {
        survivorStart = 1;
      } else {
        survivorStart = ss.getRangeByName('SURVIVOR_START').getValue();
      }
    }
    if (commentInclude == null) {
      commentInclude = ss.getRangeByName('COMMENTS_PRESENT').getValue();
    }
  }
  catch (err) {
    ss.toast('Error with getting information from CONFIG sheet or from runFirst script input, you may need to recreate everything or look at your version history');
    Logger.log('sheetSheet Error: ' + err.stack);
  }

  // Establish generic name if null provided
  if (name == null) {
    if (pickemsInclude) {
      name = league + ' Pick \'Ems';
    } else {
      name = league + ' Survivor Pool';
    }
  }
  let array = [['NAME',name],['ACTIVE\ WEEK',week],['TOTAL\ WEEKS',weeks],['YEAR',year],['PICK\ \'EMS',pickemsInclude],['MNF',mnfInclude],['TNF',tnfInclude],['TIEBREAKERS',tiebreaker],['BONUS GAMES',bonus],['MNF DOUBLE',false],['COMMENTS',commentInclude],['SURVIVOR',survivorInclude],['SURVIVOR\ DONE',''],['SURVIVOR\ START',survivorStart]];
  let endData = array.length;
  let arrayNamedRanges = ['NAME','WEEK','WEEKS','YEAR','PICKEMS_PRESENT','MNF_PRESENT','TNF_PRESENT','TIEBREAKER_PRESENT','BONUS_PRESENT','MNF_DOUBLE','COMMENTS_PRESENT','SURVIVOR_PRESENT','SURVIVOR_DONE','SURVIVOR_START'];
  const trueFalseCount = 9; // Number of named ranges with true and false values for conditional formatting
  const dataValidationCount = 8; // Number of named ranges with data validation rules

  // Fix total rows and columns
  if(sheet.getMaxRows() > (endData + weeks + 2)) {
    sheet.deleteRows((endData + weeks + 2) + 1,sheet.getMaxRows() - (endData + weeks + 2));
  } else if (sheet.getMaxRows() < ((endData+1) + weeks)) {
    sheet.insertRowsAfter(sheet.getMaxRows(),((endData+1)  + weeks + 1) - sheet.getMaxRows());
  }
  if(sheet.getMaxColumns() > 4) {
    sheet.deleteColumns(5,sheet.getMaxColumns()-4);
  } else if (sheet.getMaxColumns() < 4 ) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(),4-sheet.getMaxColumns());
  }

  sheet.getRange(1,1,endData,2).setValues(array);
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.getRange(1,2,1,4).mergeAcross();
  sheet.getRange(1,2).setValue(name);
  sheet.getRange(endData+1,1,2,1).setValues([['TITLE'],['SHEET']]);
  let weeksArr = [];
  for(let a = 1; a <= weeks; a++) {
    sheet.getRange(a+(endData+2),1).setValue(a);
    ss.setNamedRange('FORM_WEEK_'+a,sheet.getRange(a+(endData+2),2));
    weeksArr.push(a);
  }

  // Setting values and named ranges of Config sheet
  sheet.getRange(endData+1,2).setValue('ID');
  sheet.getRange(endData+2,2).setValue(ss.getId());
  sheet.getRange(endData+1,3).setValue('SHAREABLE');
  sheet.getRange(endData+2,3).setValue(ss.getUrl().slice(0,-5));
  sheet.getRange(endData+1,4).setValue('EDITABLE');
  sheet.getRange(endData+2,4).setValue(ss.getUrl());
  // Sets all named ranges of those values in array from above
  for (let a = 0; a < arrayNamedRanges.length; a++) {
    ss.setNamedRange(arrayNamedRanges[a],sheet.getRange(arrayNamedRanges.indexOf(arrayNamedRanges[a])+1,2));
  }

  // Puts formula in survivor done cell (likely needs to be replaced to trigger recalculation later)
  survivorDoneFormula(ss);

  // Rules for dropdowns on Config sheet
  let rule = SpreadsheetApp.newDataValidation().requireValueInList(weeksArr, true).build();
  sheet.getRange(2,2).setDataValidation(rule);

  rule = SpreadsheetApp.newDataValidation().requireValueInList([true,false], true).build();
  let range = sheet.getRange(5,2,dataValidationCount,1);
  range.setDataValidation(rule);
  rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(weeksArr)
    .build();
  sheet.getRange(endData,2).setDataValidation(rule);

  // TRUE COLOR FORMAT
  range = sheet.getRange(5,2,trueFalseCount,1);
  let formatTrue = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("true")
    .setBackground('#c9ffdf')
    .setRanges([range])
    .build();
  // FALSE COLOR FORMAT
  let formatFalse = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("false")
    .setBackground('#ffc4ca')
    .setRanges([range])
    .build();

  let formats = [formatTrue,formatFalse];
  sheet.setConditionalFormatRules(formats);

  // Formats sheet
  sheet.setColumnWidths(1,1,150);
  sheet.setColumnWidths(2,1,60);
  sheet.setColumnWidths(3,2,120);
  sheet.getRange(1,2,sheet.getMaxRows(),1).clearNote();
  sheet.getRange(2,2).setNote('Week that the form script will reference');
  sheet.getRange(endData-2,2).setNote('This is a calculated value, don\'t change unless you know what you\'re doing');
  sheet.getRange(endData-2,2).setNote('Prompts in the form creation should result in this being changed automatically, only change if you know what you\'re doing');
  sheet.getRange((endData+2),3).setNote('Use this to share to the group -- but make sure to make the spreadsheet shared for View Only with a link!');
  range = sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns());
  range.setHorizontalAlignment('left');
  range.setVerticalAlignment('center');
  let style = SpreadsheetApp.newTextStyle()
    .setFontFamily("Montserrat")
    .setFontSize(10)
    .build();
  range.setTextStyle(style);
  style = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .build();
  sheet.getRange(1,1,sheet.getMaxRows(),1).setTextStyle(style);
  sheet.getRange(endData+1,1,1,sheet.getMaxColumns()).setTextStyle(style);
  style = SpreadsheetApp.newTextStyle()
    .setFontFamily("Montserrat")
    .setFontSize(14)
    .setBold(true)
    .build();
  sheet.getRange(1,1,1,2).setTextStyle(style);
  sheet.hideSheet();
  return sheet;
}

// MEMBERS Sheet Creation / Adjustment
function memberSheet(ss,members) {
  ss = fetchSpreadsheet(ss);
  if (members == null) {
    members = memberList(ss);
  }
  let totalMembers = members.length;

  let sheetName = 'MEMBERS';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName,0);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.setTabColor(configTabColor);

  let rows = Math.max(members.length,1);
  let maxRows = sheet.getMaxRows();
  if ( rows < maxRows ) {
    sheet.deleteRows(rows,maxRows-rows);
  }
  let maxCols = sheet.getMaxColumns();
  if ( maxCols > 1 ) {
    sheet.deleteColumns(1,maxCols-1);
  }
  let range = sheet.getRange(1,1,rows,1);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('left');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  ss.setNamedRange(sheetName,range);
  let arr = [];
  for (let a = 0; a < members.length; a++) {
    arr.push([members[a]]);
  }
  if (members.length > 0) {
    sheet.getRange(1,1,totalMembers,1).setValues(arr);
  }
  memberList(ss);
  sheet.setColumnWidth(1,120);
  sheet.hideSheet();
  return sheet;
}

// MEMBERS Sheet Check if protected returns true or false
function membersSheetProtected() {
  let locked = false;
  try {
    let protections = SpreadsheetApp.getActiveSpreadsheet().getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (let a = 0; a < protections.length; a++) {
      if (protections[a].getDescription() == "MEMBERS PROTECTION") {
        locked = true;
      }
    }
  }
  catch (err) {
    Logger.log('membersSheetProtected error: ' + err.message + ' \r\n' + err.stack);
    return locked;
  }
  Logger.log('Membership lock is ' + locked);
  return locked;
}

// MEMBERS Sheet Locking (protection)
function membersSheetLock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('MEMBERS');
  sheet.protect().setDescription('MEMBERS PROTECTION');
  Logger.log('locked MEMBERS');
}

// MEMBERS Sheet Unlocking (remove protection);
function membersSheetUnlock() {
  try {
    let protections = SpreadsheetApp.getActiveSpreadsheet().getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (let a = 0; a < protections.length; a++) {
      if (protections[a].getDescription() == "MEMBERS PROTECTION") {
        protections[a].remove();
        Logger.log('unlocked MEMBERS');
      }
    }
  }
  catch (err) {
    Logger.log('membersSheetUnlock error: ' + err.message + ' \r\n' + err.stack);
  }
}

// TOTAL Sheet Creation / Adjustment
function totSheet(ss,weeks,members) {
  ss = fetchSpreadsheet(ss);
  let sheetName = 'TOTAL';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.clear();
  sheet.setTabColor(generalTabColor);

  if (weeks == null) {
    weeks = fetchWeeks();
  }
  let totalMembers = members.length;

  let rows = totalMembers + 2;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }

  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( weeks + 2 < maxCols ) {
    sheet.deleteColumns(weeks + 2,maxCols-(weeks + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('CORRECT');
  sheet.getRange(1,2).setValue('TOTAL');
  sheet.getRange(2,1).setValue('AVERAGES');

  for ( let a = 0; a < weeks; a++ ) {
    sheet.getRange(1,a+3).setValue(a+1);
    sheet.setColumnWidth(a+3,30);
    sheet.getRange(2,a+3).setFormula('=iferror(arrayformula(countif(filter('+league+'_PICKS_'+(a+1)+',NAMES_'+(a+1)+'=$A2)='+league+'_PICKEM_OUTCOMES_'+(a+1)+',true)),)');
  }

  let range = sheet.getRange(1,1,rows,maxCols);
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
  let rangeOverallTotNames = sheet.getRange('R2C1:R'+rows+'C1');
  ss.setNamedRange('TOT_OVERALL_NAMES',rangeOverallTotNames);
  sheet.clearConditionalFormatRules();
  // OVERALL TOTAL GRADIENT RULE
  let rangeOverallTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('TOT_OVERALL',rangeOverallTot);
  let formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("TOT_OVERALL"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("TOT_OVERALL"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("TOT_OVERALL"))') // Min value of all correct picks
    .setRanges([rangeOverallTot])
    .build();
  // OVERALL SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks+2));
  ss.setNamedRange('TOT_WEEKLY',range);
  let formatRuleOverallHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R[0]C[0]\",false)>0,indirect(\"R[0]C[0]\",false)=max(indirect(\"R2C[0]:R'+maxRows+'C[0]\",false)))')
    .setBackground('#75F0A1')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOverall = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, "15")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "10")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "5")
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallHigh);
  formatRules.push(formatRuleOverall);
  formatRules.push(formatRuleOverallTot);
  sheet.setConditionalFormatRules(formatRules);

  overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',true);
  overallMainFormulas(sheet,totalMembers,weeks,'TOT',true);

  return sheet;
}

// RNK Sheet Creation / Adjustment
function rnkSheet(ss,weeks,members) {
  ss = fetchSpreadsheet(ss);
  let sheetName = 'RNK';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.clear();
  sheet.setTabColor(generalTabColor);

  if (weeks == null) {
    weeks = fetchWeeks();
  }
  if (members == null) {
    members = memberList(ss);
  }

  let totalMembers = members.length;

  let rows = totalMembers + 1;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( weeks + 2 < maxCols ) {
    sheet.deleteColumns(weeks + 2,maxCols-(weeks + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('RANKS');
  sheet.getRange(1,2).setValue('AVERAGE');

  let range = sheet.getRange(1,1,rows,maxCols);
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
  let rangeOverallTotRnkNames = sheet.getRange('R2C1:R'+rows+'C1');
  ss.setNamedRange('TOT_OVERALL_RNK_NAMES',rangeOverallTotRnkNames);
  sheet.clearConditionalFormatRules();
  // RANKS TOTAL GRADIENT RULE
  let rangeOverallRankTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('TOT_OVERALL_RANK',rangeOverallRankTot);
  let formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([rangeOverallRankTot])
    .build();
  // RANKS SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks+2));
  ss.setNamedRange('TOT_WEEKLY_RANK',range);
  let formatRuleOverallWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#00E1FF')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOverall = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallWinner);
  formatRules.push(formatRuleOverall);
  formatRules.push(formatRuleOverallTot);
  sheet.setConditionalFormatRules(formatRules);

  overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',false);
  overallMainFormulas(sheet,totalMembers,weeks,'RANK',false);

  return sheet;
}

// PCT Sheet Creation / Adjustment
function pctSheet(ss,weeks,members) {
  ss = fetchSpreadsheet(ss);

  let sheetName = 'PCT';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.clear();
  sheet.setTabColor(generalTabColor);

  if (weeks == null) {
    weeks = fetchWeeks();
  }
  if (members == null) {
    members = memberList(ss);
  }
  let totalMembers = members.length;

  let rows = totalMembers + 2; // 2 additional rows
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( weeks + 2 < maxCols ) {
    sheet.deleteColumns(weeks + 2,maxCols-(weeks + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('PERCENTS');
  sheet.getRange(1,2).setValue('AVERAGE');
  sheet.getRange(totalMembers + 2,1).setValue('AVERAGES');

  for ( let a = 0; a < weeks; a++ ) {
    sheet.getRange(1,a+3).setValue(a+1);
    sheet.setColumnWidth(a+3,48);
  }

  let range = sheet.getRange(1,1,rows,maxCols);
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
  let rangeOverallTotPctNames = sheet.getRange('R2C1:R'+(rows-1)+'C1');
  ss.setNamedRange('TOT_OVERALL_PCT_NAMES',rangeOverallTotPctNames);
  sheet.clearConditionalFormatRules();
  // PCT TOTAL GRADIENT RULE
  let rangeOverallTotPct = sheet.getRange('R2C2:R'+(rows-1)+'C2');
  ss.setNamedRange('TOT_OVERALL_PCT',rangeOverallTotPct);
  rangeOverallTotPct = sheet.getRange('R2C2:R'+rows+'C2');
  let formatRuleOverallPctTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("TOT_OVERALL_PCT"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("TOT_OVERALL_PCT"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("TOT_OVERALL_PCT"))') // Min value of all correct picks
    .setRanges([rangeOverallTotPct])
    .build();
  // PCT SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+(rows-1)+'C'+(weeks+2));
  ss.setNamedRange('TOT_WEEKLY_PCT',range);
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks+2));
  let formatRuleOverallPctHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R[0]C[0]\",false)>0,indirect(\"R[0]C[0]\",false)=max(indirect(\"R2C[0]:R'+maxRows+'C[0]\",false)))')
    .setBackground('#75F0A1')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOverallPct = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "0.5")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallPctHigh);
  formatRules.push(formatRuleOverallPct);
  formatRules.push(formatRuleOverallPctTot);
  sheet.setConditionalFormatRules(formatRules);

  overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',true);
  overallMainFormulas(sheet,totalMembers,weeks,'PCT',true);

  return sheet;
}

// MNF Sheet Creation / Adjustment
function mnfSheet(ss,weeks,members) {
  ss = fetchSpreadsheet(ss);

  let sheetName = 'MNF';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.clear();
  sheet.setTabColor(generalTabColor);

  if (weeks == null) {
    weeks = fetchWeeks();
  }
  if (members == null) {
    members = memberList(ss);
  }
  let totalMembers = members.length;

  Logger.log('Checking for Monday games, if any');
  let data = ss.getRangeByName(league).getValues();
  let text = '0';
  let result = text.repeat(weeks);
  let mondayNightGames = Array.from(result);
  for (let a = 0; a < data.length; a++) {
    if ( data[a][2] == 1 && data[a][3] >= 17) {
      mondayNightGames[(data[a][0]-1)]++;
    }
  }
  let rows = totalMembers + 2; // +1 for header row and +1 for footer/stat row
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( weeks + 2 < maxCols ) {
    sheet.deleteColumns(weeks + 2,maxCols-(weeks + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('CORRECT');
  sheet.getRange(1,2).setValue('TOTAL');

  let range = sheet.getRange(1,1,rows,maxCols);
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

  let headers = [];
  for ( let a = 0; a < weeks; a++ ) {
    if (mondayNightGames[a] == 2) {
      range = sheet.getRange(1,a+3);
      range.setNote('Two MNF Games')
        .setFontWeight('bold')
        .setBackground('#666666');
    } else if (mondayNightGames[a] == 3) {
      range = sheet.getRange(1,a+3);
      range.setNote('Three MNF Games')
        .setFontWeight('bold')
        .setBackground('#AAAAAA');
    }
    sheet.setColumnWidth(a+3,30);
    headers.push(a+1);
  }
  sheet.getRange(1,3,1,weeks).setValues([headers]);

  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);

  sheet.clearConditionalFormatRules();

  // SET MNF NAMES Range
  let rangeMnfNames = sheet.getRange('R2C1:R'+rows+'C1');
  ss.setNamedRange('MNF_NAMES',rangeMnfNames);
  // MNF TOTAL GRADIENT RULE
  let rangeMnfTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('MNF',rangeMnfTot);
  let formatRuleMnfTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#C9FFDF", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("MNF"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("MNF"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("MNF"))') // Min value of all correct picks
    .setRanges([rangeMnfTot])
    .build();
  // MNF SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks+2));
  ss.setNamedRange('MNF_WEEKLY',range);
  let formatRuleTwoCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(2)
    .setBackground('#9CFFC4')
    .setFontColor('#9CFFC4')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOneCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#C9FFDF')
    .setFontColor('#C9FFDF')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleIncorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=or(and(not(isblank(indirect(\"R[0]C[0]\",false))),indirect(\"R[0]C[0]\",false)=0),and(isblank(indirect(\"R[0]C[0]\",false)),indirect(\"WEEK\")>=indirect(\"R1C[0]\",false)))')
    .setBackground('#FFC4CA')
    .setFontColor('#FFC4CA')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleTwoCorrect);
  formatRules.push(formatRuleOneCorrect);
  formatRules.push(formatRuleIncorrect);
  formatRules.push(formatRuleMnfTot);
  sheet.setConditionalFormatRules(formatRules);

  overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',false);
  overallMainFormulas(sheet,totalMembers,weeks,'MNF',false);

  return sheet;
}

// SURVIVOR Sheet Creation / Adjustment
function survivorSheet(ss,weeks,members,dataRestore) {
  ss = fetchSpreadsheet(ss);

  const sheetName = 'SURVIVOR';
  let sheet = ss.getSheetByName(sheetName);
  let fresh = false;
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
    fresh = true;
  }

  sheet.setTabColor(winnersTabColor);

  if (members == null) {
    members = memberList(ss);
  }
  const totalMembers = members.length;

  let maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();

  let previousDataRange, previousData;
  if (dataRestore && !fresh){
    previousDataRange = sheet.getRange(2,3,maxRows-2,weeks);
    previousData = previousDataRange.getValues();
    ss.toast('Gathered previous data for SURVIVOR sheet, recreating sheet now');
  }
  sheet.clear();

  let rows = totalMembers + 2;
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let cols = weeks + 2;
  if (cols < maxCols) {
    sheet.deleteColumns(cols + 1,maxCols-cols);
  } else if (cols > maxCols) {
    sheet.insertColumnsAfter(maxCols,cols-maxCols);
  }
  maxCols = sheet.getMaxColumns();

  sheet.getRange(1,1).setValue('PLAYER');
  let eliminatedCol = 2;
  sheet.getRange(1,eliminatedCol).setValue('ELIMINATED');
  sheet.setColumnWidth(eliminatedCol,100);

  for (let a = 0; a < weeks; a++ ) {
    sheet.getRange(1,a+3).setValue(a+1);
    sheet.setColumnWidth(a+3,30);
  }

  let formula;
  for (let b = 2; b <= totalMembers; b++ ) {
    formula = '=iferror(vlookup(indirect(\"R[0]C1\",false),SURVIVOR_EVAL,2,false))';
    sheet.getRange(2,eliminatedCol,b,1).setFormulaR1C1(formula);
  }
  for (let b = 1; b < weeks; b++ ) {
    formula = '=if(indirect(\"R1C[0]\",false)<SURVIVOR_START,,iferror(if(sum(arrayformula(if(isblank(R2C[0]:R[-1]C[0]),0,1)))>0,counta(R2C1:R[-1]C1)-countif(R2C2:R[-1]C2,\"\<=\"\&R1C[0]),)))';
    sheet.getRange(totalMembers+2,eliminatedCol+b).setFormulaR1C1(formula);
  }

  formula = '=iferror(rows(R2C[0]:R[-1]C[0])-counta(R2C[0]:R[-1]C[0]))';
  sheet.getRange(totalMembers+2,eliminatedCol).setFormulaR1C1(formula);

  let range = sheet.getRange(1,1,rows,weeks+2);
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
  range = sheet.getRange('R2C3:R'+(totalMembers+1)+'C'+(weeks+2));
  // BLANK COLOR RULE
  let formatRuleBlank = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground('#FFFFFF')
    .setRanges([range])
    .build();
  // ELIMINATED COLOR RULE
  let formatRuleCorrectElim = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C'+eliminatedCol+'\",false))),(indirect(\"R1C[0]\",false)-(indirect(\"SURVIVOR_START\")-1))>indirect(\"R[0]C'+eliminatedCol+'\",false))')
    .setBackground('#ffeded')
    .setFontColor('#ffeded')
    .setRanges([range])
    .build();
  // CORRECT PICK COLOR RULE PREVIOUS
  let formatRuleCorrectPrevious = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R1C[0]\",false)<indirect(\"SURVIVOR_START\"),vlookup(indirect(\"R[0]C1\",false),indirect(\"SURVIVOR_EVAL\"),match(indirect(\"R1C[0]\",false),indirect(\"SURVIVOR_EVAL_WEEKS\"),0)+'+eliminatedCol+',false)=0,not(isblank(vlookup(indirect(\"R[0]C1\",false),indirect(\"SURVIVOR_EVAL\"),match(indirect(\"R1C[0]\",false),indirect(\"SURVIVOR_EVAL_WEEKS\"),0)+'+eliminatedCol+',false))),not(isblank(indirect(\"R[0]C[0]\",false))))')
    .setBackground('#f2fff7')
    .setFontColor('#bcd1c4')
    .setRanges([range])
    .build();
  // CORRECT PICK COLOR RULE
  let formatRuleCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(vlookup(indirect(\"R[0]C1\",false),indirect(\"SURVIVOR_EVAL\"),match(indirect(\"R1C[0]\",false),indirect(\"SURVIVOR_EVAL_WEEKS\"),0)+'+eliminatedCol+',false)=0,not(isblank(vlookup(indirect(\"R[0]C1\",false),indirect(\"SURVIVOR_EVAL\"),match(indirect(\"R1C[0]\",false),indirect(\"SURVIVOR_EVAL_WEEKS\"),0)+'+eliminatedCol+',false))),not(isblank(indirect(\"R[0]C[0]\",false))))')
    .setBackground('#c9ffdf')
    .setRanges([range])
    .build();

  // INCORRECT PICK COLOR RULE PREVIOUS
  let formatRuleIncorrectPrevious = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R1C[0]\",false)<indirect(\"SURVIVOR_START\"),vlookup(indirect(\"R[0]C1\",false),indirect(\"SURVIVOR_EVAL\"),match(indirect(\"R1C[0]\",false),indirect(\"SURVIVOR_EVAL_WEEKS\"),0)+'+eliminatedCol+',false)=1)')
    .setBackground('#f7dfe1')
    .setFontColor('#ccb6b7')
    .setStrikethrough(true)
    .setRanges([range])
    .build();
  // INCORRECT PICK COLOR RULE
  let formatRuleIncorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=vlookup(indirect(\"R[0]C1\",false),indirect(\"SURVIVOR_EVAL\"),match(indirect(\"R1C[0]\",false),indirect(\"SURVIVOR_EVAL_WEEKS\"),0)+'+eliminatedCol+',false)=1')
    .setBackground('#f2bdc2')
    .setStrikethrough(true)
    .setRanges([range])
    .build();
  // ELIMINATED COLOR RULE
  range = sheet.getRange('R2C2:R'+(totalMembers+1)+'C2');
  let formatRuleEliminatedColorScale = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#f5d5d8')
    .setGradientMinpoint('#f07883')
    .setRanges([range])
    .build();
  // NOT ELIMINATED NAME RULE
  range = sheet.getRange('R2C1:R'+(totalMembers+1)+'C1');
  let formatRuleNotEliminatedName = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=isblank(indirect(\"R[0]C'+eliminatedCol+'\",false))')
    .setBackground('#ffffff')
    .setFontColor('#000000')
    .setBold(true)
    .setRanges([range])
    .build();
  // ELIMINATED COLOR RULE
  range = sheet.getRange('R2C1:R'+(totalMembers+1)+'C'+(weeks+2));
  let formatRuleEliminated = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=not(isblank(indirect(\"R[0]C'+eliminatedCol+'\",false)))')
    .setBackground('#f2bdc2')
    .setRanges([range])
    .build();
  // CORRECT PICK COLOR RULE
  let formatRuleMaybeCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),indirect(\"R1C[0]\",false)=indirect(\"WEEK\"))')
    .setBackground('#fffec9')
    .setRanges([range])
    .build();
  // HEADER RULE FOR PREVIOUS WEEKS
  range = sheet.getRange(1,eliminatedCol+1,1,weeks);
  let formatHeadersPrevious = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=indirect(\"R1C[0]\",false)<indirect(\"SURVIVOR_START\")')
    .setBackground('#999999')
    .setFontColor('#bbbbbb')
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleCorrectElim);
  formatRules.push(formatRuleCorrectPrevious);
  formatRules.push(formatRuleIncorrectPrevious);
  formatRules.push(formatRuleCorrect);
  formatRules.push(formatRuleIncorrect);
  formatRules.push(formatRuleEliminatedColorScale);
  formatRules.push(formatRuleNotEliminatedName);
  formatRules.push(formatRuleEliminated);
  formatRules.push(formatRuleMaybeCorrect);
  formatRules.push(formatRuleBlank);
  formatRules.push(formatHeadersPrevious);
  sheet.setConditionalFormatRules(formatRules);

  range = sheet.getRange('R2C'+(eliminatedCol-1)+':R'+(totalMembers+1)+'C'+(eliminatedCol-1));
  ss.setNamedRange('SURVIVOR_NAMES',range);
  range = sheet.getRange('R2C'+eliminatedCol+':R'+(totalMembers+1)+'C'+eliminatedCol);
  ss.setNamedRange('SURVIVOR_ELIMINATED',range);
  range = sheet.getRange('R2C'+(eliminatedCol+1)+':R'+(totalMembers+1)+'C'+(weeks+2));
  ss.setNamedRange('SURVIVOR_PICKS',range);


  if (dataRestore && !fresh) {
    previousDataRange.setValues(previousData);
    ss.toast('Previous values restored for SURVIVOR sheet if they were present');
  }

  return sheet;
}

// SURVIVOR DONE FORMULA Updates the formula for the survivor pool completion status
function survivorDoneFormula(ss) {
  // Replace the formula in the Survivor Done cell to re-evaluate
  ss = fetchSpreadsheet(ss);
  ss.getRangeByName('SURVIVOR_DONE').setValue('=iferror(if(SURVIVOR_EVAL_REMAINING<=1,true,false))');
}

// SURVIVOR EVAL Sheet Creation / Adjustment
function survivorEvalSheet(ss,weeks,members,survivorStart) {
  ss = fetchSpreadsheet(ss);

  let sheetName = 'SURVIVOR_EVAL';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.setTabColor(configTabColor);

  if (members == null) {
    members = memberList(ss);
  }
  let totalMembers = members.length;

  if (survivorStart == null) {
    survivorStart = ss.getRangeByName('SURVIVOR_START').getValue();
  }
  if (survivorStart == null || survivorStart == '') {
    survivorStart = 1;
  }
  if (weeks == null || weeks == '') {
    weeks = fetchWeeks();
  }

  let maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();

  sheet.clear();

  let rows = totalMembers + 2;
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  if ( (weeks + 2) < maxCols ) {
    sheet.deleteColumns(weeks+2,maxCols - (weeks+2));
  }
  maxCols = sheet.getMaxColumns();

  sheet.getRange(1,1).setValue('PLAYER');
  let eliminatedCol = 2;
  sheet.getRange(1,eliminatedCol).setValue('ELIMINATED');
  sheet.setColumnWidth(eliminatedCol,100);

  for (let a = 0; a < weeks; a++ ) {
    sheet.getRange(1,a+3).setValue(a+1);
    sheet.setColumnWidth(a+3,30);
  }

  let formula;
  for (let b = 2; b <= totalMembers; b++ ) {
    formula = '=iferror(match(1,indirect(\"R[0]C\"\&(indirect(\"SURVIVOR_START\")+2)\&\"\:R[0]C20\",false),0),)';
    // Alternative formula if actual week eliminated is desired
    // formula = '=iferror(match(1,indirect(\"R[0]C\"\&(indirect(\"SURVIVOR_START\")+2)\&\"\:R[0]C20\",false),0)+(indirect(\"SURVIVOR_START\")-1),)'
    sheet.getRange(2,eliminatedCol,b,1).setFormulaR1C1(formula);
  }
  for (let b = 1; b < weeks; b++ ) {
    formula = '=if(indirect(\"R1C[0]\",false)<indirect(\"SURVIVOR_START\"),,iferror(if(sum(arrayformula(if(isblank(R2C[0]:R[-1]C[0]),0,1)))>0,counta(R2C1:R[-1]C1)-countif(R2C2:R[-1]C2,\"\<=\"\&R1C[0]),)))';
    sheet.getRange(totalMembers+2,eliminatedCol+b).setFormulaR1C1(formula);
  }
  formula = '=iferror(iferror(if(match(indirect(\"SURVIVOR!R[0]C[0]\",false),indirect(\"'+league+'_OUTCOMES_\"\&indirect(\"R1C[0]\",false),false),0)>0,0,1),iferror(if(match(iferror(vlookup(indirect(\"SURVIVOR!R[0]C[0]\",false),{indirect(\"'+league+'_AWAY_\"\&indirect(\"R1C[0]\",false)),indirect(\"'+league+'_HOME_\"\&indirect(\"R1C[0]\",false))},2,false),vlookup(indirect(\"SURVIVOR!R[0]C[0]\",false),{indirect(\"'+league+'_HOME_\"\&indirect(\"R1C[0]\",false)),indirect(\"'+league+'_AWAY_\"\&indirect(\"R1C[0]\",false))},2,false)),indirect(\"'+league+'_OUTCOMES_\"\&indirect(\"R1C[0]\",false),false),0)>0,1,0),if(and(isblank(indirect(\"SURVIVOR!R[0]C[0]\",false)),indirect(\"R1C[0]\",false)<WEEK),1,if(and(isblank(indirect(\"SURVIVOR!R[0]C[0]\",false)),indirect(\"R1C[0]\",false)<WEEK,indirect(\"R1C[0]\",false)<>indirect(\"SURVIVOR_START\")),1,)))))';
  sheet.getRange(2,eliminatedCol+1,totalMembers,weeks).setFormulaR1C1(formula);

  formula = '=iferror(rows(R2C[0]:R[-1]C[0])-counta(R2C[0]:R[-1]C[0]))';
  sheet.getRange(totalMembers+2,eliminatedCol).setFormulaR1C1(formula);

  let range = sheet.getRange(1,1,rows,weeks+2);
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
  range = sheet.getRange('R2C3:R'+(totalMembers+1)+'C'+(weeks+2));
  // BLANK COLOR RULE
  let formatRuleBlank = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground('#FFFFFF')
    .setRanges([range])
    .build();
  // CORRECT PICK COLOR RULE PREVIOUS
  let formatRuleCorrectPrevious = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R1C[0]\",false)<indirect(\"SURVIVOR_START\"),indirect(\"R[0]C[0]\",false)=0,and(not(isblank(indirect(\"R[0]C[0]\",false)))))')
    .setBackground('#f0fcf5')
    .setFontColor('#9dcfb1')
    .setRanges([range])
    .build();
  // CORRECT PICK COLOR RULE
  let formatRuleCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R[0]C[0]\",false)=0,and(not(isblank(indirect(\"R[0]C[0]\",false)))))')
    .setBackground('#c9ffdf')
    .setFontColor('#6bffa7')
    .setRanges([range])
    .build();
  // INCORRECT PICK COLOR RULE PREVIOUS
  let formatRuleIncorrectPrevious = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R1C[0]\",false)<indirect(\"SURVIVOR_START\"),or(indirect(\"R[0]C[0]\",false)=1,and(isblank(indirect(\"R[0]C[0]\",false)),indirect(\"R1C[0]\",false)<indirect(\"WEEK\"))))')
    .setBackground('#fcf2f3')
    .setFontColor('#dbb2b6')
    .setRanges([range])
    .build();
  // INCORRECT PICK COLOR RULE
  let formatRuleIncorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=or(indirect(\"R[0]C[0]\",false)=1,and(isblank(indirect(\"R[0]C[0]\",false)),indirect(\"R1C[0]\",false)<indirect(\"WEEK\")))')
    .setBackground('#f2bdc2')
    .setFontColor('#f57884')
    .setRanges([range])
    .build();
  // MAYBE CORRECT PICK COLOR RULE
  let formatRuleMaybeCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(isblank(indirect(\"R[0]C[0]\",false)),indirect(\"R1C[0]\",false)=indirect(\"WEEK\"))')
    .setBackground('#fffec9')
    .setFontColor('#fffec9')
    .setRanges([range])
    .build();
  // ELIMINATED COLOR RULE
  range = sheet.getRange('R2C2:R'+(totalMembers+1)+'C2');
  let formatRuleEliminatedColorScale = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#f5d5d8')
    .setGradientMinpoint('#f07883')
    .setRanges([range])
    .build();
  // ELIMINATED NAME COLOR RULE
  range = sheet.getRange('R2C1:R'+(totalMembers+1)+'C1');
  let formatRuleEliminatedName = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=not(isblank(indirect(\"R[0]C'+eliminatedCol+'\",false)))')
    .setBackground('#f2bdc2')
    .setFontColor('#222222')
    .setRanges([range])
    .build();
  let formatRuleNotEliminatedName = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=isblank(indirect(\"R[0]C'+eliminatedCol+'\",false))')
    .setBackground('#ffffff')
    .setFontColor('#000000')
    .setBold(true)
    .setRanges([range])
    .build();
  // HIDE VISIBILITY OF UNEVALUATED NUMBERS
  range = sheet.getRange('R2C3:R'+(totalMembers+1)+'C'+(weeks+2));
  let formatRuleWhite = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=not(isblank(indirect(\"R[0]C[0]\",false)))')
    .setBackground('#ffffff')
    .setFontColor('#ffffff')
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleMaybeCorrect);
  formatRules.push(formatRuleCorrectPrevious);
  formatRules.push(formatRuleIncorrectPrevious);
  formatRules.push(formatRuleCorrect);
  formatRules.push(formatRuleIncorrect);
  formatRules.push(formatRuleBlank);
  formatRules.push(formatRuleEliminatedColorScale);
  formatRules.push(formatRuleNotEliminatedName);
  formatRules.push(formatRuleEliminatedName);
  formatRules.push(formatRuleWhite);
  sheet.setConditionalFormatRules(formatRules);

  range = sheet.getRange(2,(eliminatedCol-1),totalMembers,1);
  ss.setNamedRange('SURVIVOR_EVAL_NAMES',range);
  range = sheet.getRange(2,eliminatedCol,totalMembers,1);
  ss.setNamedRange('SURVIVOR_EVAL_ELIMINATED',range);
  range = sheet.getRange(totalMembers+2,eliminatedCol);
  ss.setNamedRange('SURVIVOR_EVAL_REMAINING',range);
  range = sheet.getRange(1,(eliminatedCol+1),1,weeks);
  ss.setNamedRange('SURVIVOR_EVAL_WEEKS',range);
  range = sheet.getRange(2,1,totalMembers,weeks+2);
  ss.setNamedRange('SURVIVOR_EVAL',range);

  survivorDoneFormula(ss);

  sheet.hideSheet();

  return sheet;
}

// WINNERS Sheet Creation / Adjustment
function winnersSheet(ss,year,weeks,members) {
  ss = fetchSpreadsheet(ss);

  let sheetName = 'WINNERS';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.clear();
  sheet.setTabColor(winnersTabColor);

  let checkboxRange = sheet.getRange(2,3,weeks+3,1);
  let checkboxes = checkboxRange.getValues();

  if (members == null) {
    members = memberList(ss);
  }

  let rows = weeks + 4;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( 3 < maxCols ) {
    sheet.deleteColumns(3,maxCols-3);
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue(year);
  sheet.getRange(1,2).setValue('WINNER');
  sheet.getRange(1,3).setValue('PAID');

  let range = sheet.getRange(1,1,rows,maxCols);
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,2,rows-1,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,80);
  sheet.setColumnWidth(2,150);
  sheet.setColumnWidth(3,40);

  range = sheet.getRange(2,3,weeks+3,1);
  range.insertCheckboxes();
  range.setHorizontalAlignment('center');
  range = sheet.getRange(1,1,rows,2);
  range.setHorizontalAlignment('left');
  let a = 0;
  for (a; a <= weeks; a++) {
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
  let fivePlusWins = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('=countif($2:B$'+(weeks+1)+',B2)>=5')
  .setBackground('#2CFF75')
  .setRanges([range])
  .build();
  let fourPlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks+1)+',B2)=4')
    .setBackground('#72FFA3')
    .setRanges([range])
    .build();
  let threePlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks+1)+',B2)=3')
    .setBackground('#BBFFD3')
    .setRanges([range])
    .build();
  let twoPlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks+1)+',B2)=2')
    .setBackground('#D3FFE2')
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(fivePlusWins);
  formatRules.push(fourPlusWins);
  formatRules.push(threePlusWins);
  formatRules.push(twoPlusWins);
  sheet.setConditionalFormatRules(formatRules);

  // Rewrites the checkboxes if they previously had any checked.
  let col = checkboxRange.getColumn();
  for (let a = 0; (a < checkboxes.length || a < (weeks + 3)); a++) {
    if (checkboxes[a][0]) {
      sheet.getRange(a+1,col).check();
    }
  }
  let winRange;
  let nameRange;

  for ( let b = 1; b <= weeks; b++ ) {
    winRange = 'WIN_' + (b);
    nameRange = 'NAMES_' + (b);
    sheet.getRange(b+1,2,1,1).setFormulaR1C1('=iferror(join(", ",sort(filter('+nameRange+','+winRange+'=1),1,true)))');
  }

  return sheet;

}

// SUMMARY Sheet Creation / Adjustment
function summarySheet(ss,members,pickemsInclude,mnfInclude,survivorInclude) {
  ss = fetchSpreadsheet(ss);

  if (pickemsInclude == null) {
    pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  }

  if (pickemsInclude) {
    if (mnfInclude == null) {
      mnfInclude = ss.getRangeByName('MNF_PRESENT').getValue();
    }
  }

  if (survivorInclude == null) {
    survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
  }
  let restoreNotes = false;
  let notesRange, notes, sheetName = 'SUMMARY';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  } else {
    restoreNotes = true;
    notesRange = sheet.getRange(2,sheet.getRange(1,1,sheet.getMaxRows()-1,sheet.getMaxColumns()).getValues().flat().indexOf('NOTES')+1,sheet.getMaxRows()-1,1);
    notes = notesRange.getValues();
  }
  sheet.clear();
  sheet.setTabColor(winnersTabColor);

  if (members == null) {
    members = memberList(ss);
  }

  let headers = ['PLAYER'];
  let headersWidth = [120];
  let mnfCol;
  if (pickemsInclude) {
    headers = headers.concat(['TOTAL CORRECT','TOTAL RANK','AVG % CORRECT','AVG % CORRECT RANK','WEEKLY WINS']);
    headersWidth = headersWidth.concat([90,90,90,90,90]);
    if (mnfInclude) {
      headers = headers.concat(['MNF CORRECT','MNF RANK']);
      headersWidth = headersWidth.concat([90,90]);
      mnfCol = headers.indexOf('MNF CORRECT') + 1;
    }
  }

  let survivorCol;
  if (survivorInclude) {
    headers.push('SURVIVOR (WEEK OUT)');
    headersWidth.push(90);
    survivorCol = headers.indexOf('SURVIVOR (WEEK OUT)')+1;
  }
  headers.push('NOTES');
  headersWidth.push(160);

  let totalCol = headers.indexOf('TOTAL CORRECT') + 1;
  let weeklyPercentCol = headers.indexOf('AVG % CORRECT') + 1;
  let weeklyRankAvgCol = headers.indexOf('AVG % CORRECT RANK') + 1;
  let weeklyWinsCol = headers.indexOf('WEEKLY WINS') + 1;
  let notesCol = headers.indexOf('NOTES') + 1;

  let len = headers.length;
  let totalMembers = members.length;

  let rows = totalMembers + 1;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( len < maxCols ) {
    sheet.deleteColumns(len,maxCols-len);
  } else if ( len > maxCols ) {
    sheet.insertColumnsAfter(maxCols, len - maxCols);
  }
  maxCols = sheet.getMaxColumns();

  sheet.getRange(1,1,1,len).setValues([headers]);
  if(restoreNotes) {
    sheet.getRange(2,notesCol,notes.length,1).setValues(notes);
  }

  for ( let a = 0; a < len; a++ ) {
    sheet.setColumnWidth(a+1,headersWidth[a]);
  }
  sheet.setRowHeight(1,40);
  let range = sheet.getRange(1,1,1,maxCols);
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
  let formatRules = sheet.getConditionalFormatRules();
  if (pickemsInclude) {
    // SUMMARY TOTAL GRADIENT RULE
    let rangeSummaryTot = sheet.getRange('R2C'+totalCol+':R'+rows+'C'+totalCol);
    let formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpoint('#75F0A1')
      .setGradientMinpoint('#FFFFFF')
      .setRanges([rangeSummaryTot])
      .build();
    formatRules.push(formatRuleOverallTot);
    // MNF TOTAL GRADIENT RULES
    let rangeMNFTot, rangeMNFRank, formatRuleMNFRank;
    if (mnfInclude) {
      rangeMNFTot = sheet.getRange('R2C'+mnfCol+':R'+rows+'C'+mnfCol);
      //ss.setNamedRange('TOT_MNF',range);
      let formatRuleMNFTot = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint('#75F0A1')
        .setGradientMinpoint('#FFFFFF')
        .setRanges([rangeMNFTot])
        .build();
      formatRules.push(formatRuleMNFTot);
      // RANK MNF GRADIENT RULE
      rangeMNFRank = sheet.getRange('R2C'+(mnfCol+1)+':R'+rows+'C'+(mnfCol+1));
      ss.setNamedRange('TOT_MNF_RANK',rangeMNFRank);
      formatRuleMNFRank = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
        .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
        .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
        .setRanges([rangeMNFRank])
        .build();
      formatRules.push(formatRuleMNFRank);
    }
    // RANK OVERALL RULE
    let rangeOverallRank = sheet.getRange('R2C'+(totalCol+1)+':R'+rows+'C'+(totalCol+1));
    ss.setNamedRange('TOT_OVERALL_RANK',rangeOverallRank);
    let formatRuleRank = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
      .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
      .setRanges([rangeOverallRank])
      .build();
    formatRules.push(formatRuleRank);
    // WEEKLY WINS GRADIENT/SINGLE COLOR RULES
    range = sheet.getRange('R2C'+weeklyWinsCol+':R'+rows+'C'+weeklyWinsCol);
    ss.setNamedRange('WEEKLY_WINS',range);
    let formatRuleWeeklyWinsEmpty = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#FFFFFF')
      .setFontColor('#FFFFFF')
      .setRanges([range])
      .build();
    formatRules.push(formatRuleWeeklyWinsEmpty);
    let formatRuleWeeklyWins = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpoint('#ffee00')
      .setGradientMinpoint('#FFFFFF')
      .setRanges([range])
      .build();
    formatRules.push(formatRuleWeeklyWins);
    // OVERALL AND WEEKLY CORRECT % AVG
    range = sheet.getRange('R2C'+weeklyPercentCol+':R'+rows+'C'+weeklyPercentCol);
    range.setNumberFormat('##.#%');
    let formatRuleCorrectAvg = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, ".70")
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, ".60")
      .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, ".50")
      .setRanges([range])
      .build();
    formatRules.push(formatRuleCorrectAvg);
    // WEEKLY RANK AVG
    range = sheet.getRange('R2C'+weeklyRankAvgCol+':R'+rows+'C'+weeklyRankAvgCol);
    range.setNumberFormat('#.#');
    let formatRuleCorrectRank = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, "5")
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "10")
      .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "15")
      .setRanges([range])
      .build();
    formatRules.push(formatRuleCorrectRank);
  }
  if (survivorInclude) {
  // SURVIVOR "IN"
    range = sheet.getRange('R2C'+survivorCol+':R'+(totalMembers+1)+'C'+survivorCol);
    let formatRuleIn = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('IN')
      .setBackground('#C9FFDF')
      .setRanges([range])
      .build();
    // SURVIVOR "OUT"
    let formatRuleOut = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('OUT')
      .setBackground('#F2BDC2')
      .setRanges([range])
      .build();
    formatRules.push(formatRuleIn);
    formatRules.push(formatRuleOut);
  }
  sheet.setConditionalFormatRules(formatRules);
  // Creates all formulas for SUMMARY Sheet
  summarySheetFormulas(totalMembers);

  return sheet;
}

// UPDATES SUMMARY SHEET FORMULAS
function summarySheetFormulas(totalMembers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SUMMARY');
  let headers = sheet.getRange('1:1').getValues().flat();
  let arr = ['PLAYER','TOTAL CORRECT','TOTAL RANK','MNF CORRECT','MNF RANK','AVG % CORRECT','AVG % CORRECT RANK','WEEKLY WINS','SURVIVOR (WEEK OUT)','NOTES'];
  headers.unshift('COL INDEX ADJUST');
  for (let a = 0; a < arr.length; a++) {
    for (let b = 0; b < totalMembers; b++) {
      if (headers[a] == 'TOTAL CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{TOT_OVERALL_NAMES,TOT_OVERALL},2,false))');
      } else if (headers[a] == 'TOTAL RANK' || headers[a] == 'AVG % CORRECT RANK' || headers[a] == 'MNF RANK') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(rank(R[0]C[-1],R2C[-1]:R'+ (totalMembers+1) + 'C[-1]))');
        ss.setNamedRange('TOT_OVERALL_RANK',sheet.getRange(2,headers.indexOf('TOTAL RANK'),totalMembers,1));
      } else if (headers[a] == 'MNF CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{MNF_NAMES,MNF},2,false))');
        ss.setNamedRange('TOT_MNF_RANK',sheet.getRange(2,headers.indexOf('MNF RANK'),totalMembers,1));
      } else if (headers[a] == 'AVG % CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{TOT_OVERALL_PCT_NAMES,TOT_OVERALL_PCT},2,false))');
      } else if (headers[a] == 'WEEKLY WINS') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(countif(WEEKLY_WINNERS,R[0]C1))');
        ss.setNamedRange('WEEKLY_WINS',sheet.getRange(2,headers.indexOf('WEEKLY WINS'),totalMembers,1));
      } else if (headers[a] == 'SURVIVOR (WEEK OUT)') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(arrayformula(if(isblank(vlookup(R[0]C1,{SURVIVOR_EVAL_NAMES,SURVIVOR_EVAL_ELIMINATED},2,false)),"IN","OUT ("\&vlookup(R[0]C1,{SURVIVOR_EVAL_NAMES,SURVIVOR_EVAL_ELIMINATED},2,false)\&")")))');
      }
    }
  }
  Logger.log('Updated formulas and ranges for summary sheet');
}

// TOT / RANK / PCT / MNF Combination formula for sum/average per player row
function overallPrimaryFormulas(sheet,totalMembers,maxCols,action,avgRow) {
  for ( let a = 1; a < totalMembers; a++ ) {
    if (action == 'average') {
      sheet.getRange(2,2,a+1,1).setFormulaR1C1('=iferror(if(counta(R[0]C3:R[0]C'+maxCols+')=0,,average(R[0]C3:R[0]C'+maxCols+')))');
    } else if (action == 'sum') {
      sheet.getRange(2,2,a+1,1).setFormulaR1C1('=iferror(if(counta(R[0]C3:R[0]C'+maxCols+')=0,,sum(R[0]C3:R[0]C'+maxCols+')))');
    }
    if (sheet.getSheetName() == 'PCT') {
      sheet.getRange(2,2,a+1,1).setNumberFormat("##.#%");
    } else if (action == 'sum') {
      sheet.getRange(2,2,a+1,1).setNumberFormat("##");
    } else {
      sheet.getRange(2,2,a+1,1).setNumberFormat("#0.0");
    }
  }
  if (avgRow && sheet.getSheetName() == 'PCT'){
    sheet.getRange(sheet.getMaxRows(),2).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))')
      .setNumberFormat('##.#%');
  } else {
    sheet.getRange(sheet.getMaxRows(),2).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))')
      .setNumberFormat("#0.0");
  }
}

// TOT / RNK / PCT / MNF Combination formula for each column (week)
function overallMainFormulas(sheet,totalMembers,weeks,str,avgRow) {
  let b;
  for (let a = 1; a <= weeks; a++ ) {
    b = 1;
    for (b ; b <= totalMembers; b++) {
      if (str == 'TOT') {
        sheet.getRange(b+1,a+2).setFormula('=iferror(if(or(iserror(vlookup($A'+(b+1)+',NAMES_'+a+',1,false)),counta(filter('+league+'_PICKS_'+a+',NAMES_'+a+'=$A'+(b+1)+'))=0),,arrayformula(countifs(filter('+league+'_PICKS_'+a+',NAMES_'+a+'=$A'+(b+1)+')='+league+'_PICKEM_OUTCOMES_'+a+',true,filter('+league+'_PICKS_'+a+',NAMES_'+a+'=$A'+(b+1)+'),\"<>\"))),)');
      } else {
        sheet.getRange(b+1,a+2).setFormula('=iferror(arrayformula(vlookup(R[0]C1,{NAMES_'+a+','+str+'_'+a+'},2,false)))');
      }
      if (sheet.getSheetName() == 'PCT') {
        sheet.getRange(b+1,a+2).setNumberFormat("##.#%");
      } else {
        sheet.getRange(b+1,a+2).setNumberFormat("#0");
      }
    }
  }
  if (avgRow){
    for (let a = 0; a < weeks; a++){
      let rows = sheet.getMaxRows();
      sheet.getRange(rows,a+3).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))');
    }
  }
}

// WEEKLY WINNERS Combination formula update
function winnersFormulas(sheet,weeks) {
  for (let a = 1; a <= weeks; a++ ) {
    let winRange = 'WIN_' + a;
    let nameRange = 'NAMES_' + a;
    sheet.getRange(a+1,2).setFormulaR1C1('=iferror(join(", ",sort(filter('+nameRange+','+winRange+'=1),1,true)))');
  }
}

// REFRESH FORMULAS FOR TOT / RNK / PCT / MNF
function allFormulasUpdate(ss){
  ss = fetchSpreadsheet(ss);
  const pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  const mnfInclude = ss.getRangeByName('MNF_PRESENT').getValue();
  const members = memberList(ss);
  const weeks = fetchWeeks();
  let sheet, totalMembers, maxCols;

  if ( pickemsInclude ) {
    sheet = ss.getSheetByName('TOT');
    maxCols = sheet.getMaxColumns();
    totalMembers = members.length;
    overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',true);
    overallMainFormulas(sheet,totalMembers,weeks,'TOT',true);

    sheet = ss.getSheetByName('RNK');
    maxCols = sheet.getMaxColumns();
    overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',false);
    overallMainFormulas(sheet,totalMembers,weeks,'RNK',false);

    sheet = ss.getSheetByName('PCT');
    maxCols = sheet.getMaxColumns();
    overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',true);
    overallMainFormulas(sheet,totalMembers,weeks,'PCT',true);

    if (mnfInclude) {
      sheet = ss.getSheetByName('MNF');
      maxCols = sheet.getMaxColumns();
      overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',false);
      overallMainFormulas(sheet,totalMembers,weeks,'MNF',false);
    }

    sheet = ss.getSheetByName('WINNERS');
    winnersFormulas(sheet,weeks);
  }
}

