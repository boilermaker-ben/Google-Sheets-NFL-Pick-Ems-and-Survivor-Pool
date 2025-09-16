// WEEKLY Sheet Creation Function - creates a sheet with provided week, members [array], and if data should be restored
function weeklySheet(ss,week,members,dataRestore) {
  ss = fetchSpreadsheet(ss);
  if (week == undefined) {
    week = fetchWeek();
  } else {
    ss.getRangeByName('WEEK').setValue(week);
  }
  if (members == undefined){
    members = memberList(ss);
  }
  let totalMembers = members.length;

  if (totalMembers <= 0) {
    let ui = SpreadsheetApp.getUi();
    ui.alert('MEMBER ISSUE\r\n\r\nThere was an issue fetching the members to create the weekly sheet, ensure that there are member names in the named range \"MEMBERS\" and try again', ui.ButtonSet.OK);
    throw new Error('Error fetching members to create weekly sheet');
  }

  let mnfInclude = ss.getRangeByName('MNF_PRESENT').getValue();
  let tnfInclude = ss.getRangeByName('TNF_PRESENT').getValue();
  let tiebreaker = ss.getRangeByName('TIEBREAKER_PRESENT').getValue();
  let bonus = ss.getRangeByName('BONUS_PRESENT').getValue();
  let mnfDouble = ss.getRangeByName('MNF_DOUBLE').getValue();
  let commentInclude = ss.getRangeByName('COMMENTS_PRESENT').getValue();

  let sheet, sheetName = weeklySheetPrefix + week;
  let data = ss.getRangeByName(league).getValues().shift();
  
  let diffCount = (totalMembers - 1) >= 5 ? 5 : (totalMembers - 1); // Number of results to display for most similar weekly picks (defaults to 5, or 1 fewer than the total member count, whichever is larger)

  const matchRow = 1; // Row for all matchups
  const dayRow = matchRow + 1; // Row for denoting day of the week
  const entryRowStart = dayRow + 1; // Row of first user input on weekly sheet
  const entryRowEnd = (entryRowStart - 1) + totalMembers; // Includes any header rows (entryRowStart-1) and adds two additional for final row of home/away splits and then bonus values
  const summaryRow = entryRowEnd + 1; // Row for group averages (away/home) and other calculated values
  const outcomeRow = summaryRow + 1; // Row for matchup outcomes
  const bonusRow = outcomeRow + 1; // Row for adding bonus drop-downs
  const rows = bonusRow; // Declare row variable, unnecessary, but easier to work with  
  const pointsCol = 2;

  let columns, fresh = false;
  
  // Checks for sheet presence and creates if necessary
  sheet = ss.getSheetByName(sheetName);  
  if (sheet == null) {
    dataRestore = false;
    ss.insertSheet(sheetName,ss.getNumSheets()+1);
    sheet = ss.getSheetByName(sheetName);
    fresh = true;
  }

  // Adds tab colors
  weeklySheetTabColors(ss,sheet); 

  let maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  
  // DATA GATHERING IF DATA RESTORE ACTIVE
  let regex, commentCol, tiebreakerCol = -1, firstInput, finalInput, previousRows, previousNames, previousData, previousOutcomes, previousComment, previousTiebreaker, previousTiebreakers, previousBonus = null, previousNamesRange, previousDataRange, previousOutcomesRange, previousCommentRange, previousTiebreakersRange, matchupStartCol, matchupEndCol;
  if (dataRestore && !fresh) {
    
    // Get first column of values from sheet and check for data input range rows
    let firstCol = sheet.getRange('A1:A').getValues().flat();
    firstCol.unshift('ROW INDEX ADJUST');
    try {
      previousNamesRange = ss.getRangeByName('NAMES_'+week);
      previousNames = previousNamesRange.getValues().flat();
      firstInput = previousNamesRange.getRow();
      finalInput = previousNamesRange.getLastRow();
    }
    catch (err) {
      Logger.log('No previous names named range found, attempting to find data rows by column values');
      regex = new RegExp(/MATCHES$/);
      if (regex.test(firstCol)) {
        for (let a = 0; a < firstCol.length; a++) {
          if (regex.test(firstCol[a])) {
            firstInput = a+1; // Denotes the row below finding "MATCHES" to mark as first user input row
          }
        }
      } else {
        firstInput = entryRowStart;
      }
      finalInput = firstCol.indexOf('PREFERRED');
      if (finalInput == -1) {
        finalInput = maxRows - 1;
      }
    }
    previousRows = finalInput - firstInput;
    
    let previousHeaders = sheet.getRange('A1:1').getValues().flat();
    previousHeaders.unshift('COL INDEX ADJUST');

    // Get matchup range and values, use column lookup if fails
    let confirmation = 'Gathered any available previous data for week ' + week + ', recreating sheet now';
    let noMatchups = false;
    try {
      previousDataRange = ss.getRangeByName(league + '_PICKS_' + week);
      previousData = previousDataRange.getValues();
      matchupStartCol = previousDataRange.getColumn();
      matchupEndCol = previousDataRange.getColumn() + previousData[0].length;
      ss.toast(confirmation);
    }
    catch (err) {
      Logger.log('No previous matchup named range found, attempting to find by header index');
      regex = new RegExp(/[A-Z]{2,3}@[A-Z]{2,3}/);
      for (let a = 0; a < previousHeaders.length; a++) {
        if (regex.test(previousHeaders[a].replace(/\s/g,''))) {
          if (matchupStartCol == undefined) {
            matchupStartCol = a;
          }
          matchupEndCol = a;
        }
      }
      if (matchupStartCol != null) {
        previousDataRange = sheet.getRange(firstInput,matchupStartCol,previousRows,matchupEndCol-matchupStartCol);
        previousData = previousDataRange.getValues();
        ss.toast(confirmation);
      } else {
        noMatchups = true;
      }
    }

    // Check if data exists, then set dataRestore to false if no data present
    regex = new RegExp(/^[A-Z]{2,3}/);
    if (!regex.test(previousData) || noMatchups) {
      dataRestore = false;
      ss.toast('Intended to restore data, but no data found. If there was any information present. please undo immediately if you want to retain information on sheet ' + sheetName);
    }
    
    if (dataRestore) {
      // Recover any marked outcomes if present
      try {
        previousOutcomesRange = ss.getRangeByName(league + '_PICKEM_OUTCOMES_'+week);
        previousOutcomes = previousOutcomesRange.getValues().flat();
      }
      catch (err) {
        Logger.log('No previous matchup outcomes named range found, referencing general location');
        previousOutcomesRange = sheet.getRange(previousDataRange.getRow()+1,matchupStartCol,1,matchupEndCol-matchupStartCol);
        previousOutcomes = previousOutcomesRange.getValues().flat();
      } 
      
      if (tiebreaker) {
        // Get tiebreaker range and values if present, use column lookup if fails
        try {
          previousTiebreakersRange = ss.getRangeByName(league + '_TIEBREAKER_' + week);
          previousTiebreakers = previousTiebreakersRange.getValues();
        }
        catch (err) {
          Logger.log('No previous tiebreaker named range found, attempting to find by header index');
          tiebreakerCol = previousHeaders.indexOf('TIEBREAKER');
          if (tiebreakerCol  >= 0) {
            try {
              previousTiebreakersRange = sheet.getRange(firstInput,tiebreakerCol,previousRows,1);
              previousTiebreakers = previousTiebreakersRange.getValues();
            }
            catch (err) {
              Logger.log('No previous tiebreaker data found to retain');
            }
          }
        }
      
        // Get previous tiebreaker final game value
        try {
          previousTiebreaker = sheet.getRange(previousTiebreakersRange.getLastRow() + 2, previousTiebreakersRange.getColumn()).getValue();
        }
        catch (err) {
          Logger.log('Unable to gather previous tiebreaker final game score value');
        }
      }

      // Get comment range and values if present, use column lookup if fails
      try {
        previousCommentRange = ss.getRangeByName('COMMENTS_' + week);
        previousComment = previousCommentRange.getValues();
      }
      catch (err) {
        Logger.log('No previous comment named range found, attempting to find by header index');
        commentCol = previousHeaders.indexOf('COMMENT');
        if (commentCol  >= 0 && commentInclude) {
          try {
            previousCommentRange = sheet.getRange(firstInput,commentCol,previousRows,1);
            previousComment = previousCommentRange.getValues();
          }
          catch (err) {
            Logger.log('No previous comment data found to retain');
          }
        }
      }      
      
      // Get bonus values if present, use row lookup if fails
      try {
        previousBonus = ss.getRangeByName(league + '_BONUS_' + week).getValues().flat();
      }
      catch (err) {
        Logger.log('No previous bonus named range found, attempting to find by column index');
        let previousBonusRow = firstCol.indexOf('BONUS');
        if (bonusRow  >= 0) {
          try {
            previousBonus = sheet.getRange(previousBonusRow,matchupStartCol,1,matchupEndCol-matchupStartCol).getValues().flat();
          }
          catch (err) {
            Logger.log('No previous bonus data found to retain');
          }
        }
      }
    } else {
      Logger.log('Skipping finding any comment, tiebreaker, or bonus data due to no matchup information found');
    }
  } else {
    Logger.log('Skipping any data restoration features.');
  }
  // DATA GATHER END

  // CLEAR AND PREP SHEET
  sheet.clear();
  sheet.clearNotes();
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).clearDataValidations();
  adjustRows(sheet,rows);
  
  // Insert Members
  sheet.getRange(entryRowStart,1,totalMembers,1).setValues(members);

  // Setting header values
  let headers = ['WEEK ' + week,'POINTS','WEEKLY\nRANK','PERCENT\nCORRECT'];
  let bottomHeaders = ['PREFERRED','AWAY','HOME'];
  sheet.getRange(summaryRow,1,1,3).setValues([bottomHeaders]);
  sheet.getRange(outcomeRow,1).setValue('OUTCOME');
  sheet.getRange(bonusRow,1).setValue('BONUS');
  let widths = [130,75,75,75];

  // Setting headers for the week's matchups with format of 'AWAY' + '@' + 'HOME', then creating a data validation cell below each
  let firstMatchCol = headers.length + 1;
  let mnfCol, mnfStartCol, mnfEndCol, tnfStartCol, tnfEndCol, winCol, days = [], dayRowColors = [], bonuses = [], formatRules = [];
  let mnf = false, tnf = false; // Preliminary establishing if there are Monday or Thursday games (false by default, fixed to true when looped through matchups)
  let rule, matches = 0;
  let exportMatches = [];
  for ( let a = 0; a < data.length; a++ ) {
    if ( data[a][0] == week && (tnfInclude || (!tnfInclude && data[a][2] >= 0))) {
      matches++;
      let day = data[a][2];
      let evening = data[a][3] >= 17 ? true : false;
      let away = data[a][6];
      let home = data[a][7];
      let matchup = away + '\n@' + home;
      exportMatches.push([day,away,home]);
      if ( previousBonus != null && (previousBonus[matches-1] >= 1 && previousBonus[matches-1] <= 3)) {
        bonuses.push(previousBonus[matches-1]);
      } else {
        if (bonus && day == 1 && mnfDouble) {
          bonuses.push(2);
        } else {
          bonuses.push(1);
        }
      }
      if ( day == 1 && evening ) {
        mnf = true;
        if ( mnfStartCol == undefined ) {
          mnfStartCol = headers.length + 1;
        }
        mnfEndCol = headers.length + 1;
      } else if ( day == -3 ) {
        tnf = true;
        if ( tnfStartCol == undefined ) {
          tnfStartCol = headers.length + 1;
        }
        tnfEndCol = headers.length + 1;
      }
      let dayIndex = day + 3;
      let writeCell = sheet.getRange(dayRow,firstMatchCol+(matches-1));
      let rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false)))')
        .setBackground(dayColorsFilled[dayIndex])
        .setBold(true)
        .setRanges([writeCell]);
      rule.build();
      formatRules.push(rule);
      dayRowColors.push(dayColors[dayIndex]);
      days.push(data[a][5]);
      headers.push(matchup);
      widths.push(50);
      rule = SpreadsheetApp.newDataValidation().requireValueInList([data[a][6],data[a][7]], true).build();
      sheet.getRange(outcomeRow,headers.length).setDataValidation(rule);
    }
  }

  const finalMatchCol = headers.length;

  if (tiebreaker) {
    headers.push('TIE\nBREAKER'); // Omitted if tiebreakers are removed
    widths.push(75);
    tiebreakerCol = headers.length;
    headers.push('DIFF');
    widths.push(50);
  }

  headers.push('WIN');
  widths.push(50);
  winCol = headers.indexOf('WIN')+1;

  if (mnfInclude && mnf) {
    headers.push('MNF'); // Added if user wants a MNF competition included
    widths.push(50);
    mnfCol = headers.indexOf('MNF')+1;
  }

  if (commentInclude) {
    headers.push('COMMENT'); // Added to allow submissions to have amusing comments, if desired
    widths.push(150);
    commentCol = headers.indexOf('COMMENT')+1;
  }

  let diffCol = headers.length+1;
  let finalCol = diffCol + (diffCount-1);

  // Headers completed, now adjusting number of columns once headers are populated
  adjustColumns(sheet,finalCol);
  maxCols = sheet.getMaxColumns();

  sheet.getRange(matchRow,1,1,headers.length).setValues([headers]);
  sheet.getRange(dayRow,firstMatchCol,1,matches).setValues([days]);
  let bonusRange = sheet.getRange(bonusRow,firstMatchCol,1,bonuses.length);
  bonusRange.setValues([bonuses]);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(['1','2','3'],true).build();
  bonusRange.setDataValidation(rule);

  for (let a = 0; a < widths.length; a++) {
    sheet.setColumnWidth(a+1,widths[a]);
  }
  
  sheet.getRange(matchRow,diffCol).setValue('SIMILAR SELECTIONS'); // Added to allow submissions to have amusing comments, if desired
  sheet.getRange(dayRow,diffCol).setValue('Displayed as the number of picks different and the name of the member')
    .setFontSize(8);

  // Create named ranges
  ss.setNamedRange(league + '_' + week,sheet.getRange(matchRow,firstMatchCol,1,matches));
  ss.setNamedRange(league + '_PICKEM_OUTCOMES_' + week,sheet.getRange(outcomeRow,firstMatchCol,1,matches));
  ss.setNamedRange(league + '_BONUS_' + week,sheet.getRange(bonusRow,firstMatchCol,1,matches));
  ss.setNamedRange(league + '_PICKS_' + week,sheet.getRange(entryRowStart,firstMatchCol,totalMembers,matches));

  if (tnfInclude && tnf) {
    ss.setNamedRange(league + '_THURS_PICKS_' + week,sheet.getRange(entryRowStart,tnfStartCol,totalMembers,tnfEndCol-tnfStartCol+1));
  }
  if (mnfInclude && mnf) {
    ss.setNamedRange(league + '_MNF_' + week,sheet.getRange(entryRowStart,mnfStartCol,totalMembers,mnfEndCol-(mnfStartCol-1)));
  }
  if (tiebreaker) {
    ss.setNamedRange(league + '_TIEBREAKER_' + week,sheet.getRange(entryRowStart,tiebreakerCol,totalMembers,1));
    let validRule = SpreadsheetApp.newDataValidation().requireNumberBetween(0,120)
      .setHelpText('Must be a number')
      .build();
    sheet.getRange(outcomeRow,tiebreakerCol).setDataValidation(validRule);
  }
  if (commentInclude) {
    ss.setNamedRange('COMMENTS_' + week,sheet.getRange(entryRowStart,commentCol,totalMembers,1));
  }

  for (let row = entryRowStart; row <= entryRowEnd; row++ ) {
    // Formula to determine how many correct on the week
    sheet.getRange(row,1,1,maxCols).setBorder(null,null,true,null,false,false,'#AAAAAA',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    sheet.getRange(row,pointsCol).setFormulaR1C1('=iferror(if(and(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+')>0,counta(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+')>0),sum(arrayformula(if(not(isblank(R'+row+'C'+firstMatchCol+':R'+row+'C'+finalMatchCol+')),if(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+'=R'+row+'C'+firstMatchCol+':R'+row+'C'+finalMatchCol+',1,0),0)*R'+bonusRow+'C'+firstMatchCol+':R'+bonusRow+'C'+finalMatchCol+')),))');

    // sheet.getRange(row,2).setFormulaR1C1('=iferror(if(and(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C['+finalMatchCol+'])>0,counta(R[0]C[3]:R[0]C['+finalMatchCol+'])>0),mmult(arrayformula(if(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+'=R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+',1,0)),transpose(arrayformula(if(not(isblank(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+')),1,0)))),))');
    
    // Formula to determine weekly rank
    sheet.getRange(row,pointsCol+1).setFormulaR1C1('=iferror(if(and(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+')>0,not(isblank(R[0]C'+pointsCol+'))),rank(R[0]C'+pointsCol+',R'+entryRowStart+'C2:R'+entryRowEnd+'C2,false),))');

    // Formula to determine weekly correct percent
    sheet.getRange(row,pointsCol+2).setFormulaR1C1('=iferror(if(and(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+')>0,not(isblank(R[0]C'+pointsCol+'))),sum(filter(arrayformula(if(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+'=R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+',1,0)),not(isblank(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+'))))/counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+'),),)');
    
    // Formula to determine difference of tiebreaker from final MNF score
    if (tiebreaker) {
      sheet.getRange(row,tiebreakerCol+1).setFormulaR1C1('=iferror(if(or(isblank(R[0]C[-1]),isblank(R'+outcomeRow+'C'+tiebreakerCol+')),,abs(R[0]C[-1]-R'+outcomeRow+'C'+tiebreakerCol+')))');
      // Formula to denote winner with a '1' if a clear winner exists
      sheet.getRange(row,winCol).setFormulaR1C1('=iferror(if(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+')=value(regexextract(R'+dayRow+'C1,\"[0-9]+\")),arrayformula(if(countif(array_constrain({R[0]C'+pointsCol+',R[0]C'+(tiebreakerCol+1)+'}=filter(filter({R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+',R'+entryRowStart+'C'+(tiebreakerCol+1)+':R'+entryRowEnd+'C'+(tiebreakerCol+1)+'},R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+'=max(R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+')),filter(R'+entryRowStart+'C'+(tiebreakerCol+1)+':R'+entryRowEnd+'C'+(tiebreakerCol+1)+',R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+'=max(R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+'))=min(filter(R'+entryRowStart+'C'+(tiebreakerCol+1)+':R'+entryRowEnd+'C'+(tiebreakerCol+1)+',R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+'=max(R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+')))),1,2),true)=2,1,0))),)');
    } else {
      // Formula to denote winner with a '1', with a tiebreaker allowed
      sheet.getRange(row,winCol).setFormulaR1C1('=iferror(if(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+')=value(regexextract(R'+dayRow+'C1,\"[0-9]+\")),if(rank(R'+row+'C'+pointsCol+',R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+',false)=1,1,0)),)');
    }

    // Formula to determine MNF win status sum (can be more than 1 for rare weeks)
    if (mnfInclude && mnf) {
      sheet.getRange(row,mnfCol).setFormulaR1C1('=iferror(if(and(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+')>0,counta(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+')>0),if(mmult(arrayformula(if(R'+outcomeRow+'C'+mnfStartCol+':R'+outcomeRow+'C'+mnfEndCol+'=R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+',1,0)),transpose(arrayformula(if(not(isblank(R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+')),1,0))))=0,0,mmult(arrayformula(if(R'+outcomeRow+'C'+mnfStartCol+':R'+outcomeRow+'C'+mnfEndCol+'=R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+',1,0)),transpose(arrayformula(if(not(isblank(R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+')),1,0))))),),)');
    }

    // Formula to generate array of similar pickers on the week
    sheet.getRange(row,diffCol).setFormulaR1C1('=iferror(if(isblank(R[0]C'+(firstMatchCol+2)+'),,transpose(arrayformula({arrayformula('+matches+'-query({R'+entryRowStart+'C1:R'+entryRowEnd+'C1,arrayformula(mmult(if(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+(finalMatchCol)+'=R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+',1,0),transpose(arrayformula(column(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+')\^0))))},\"select Col2 where Col1 <> \'\"\&R[0]C1\&\"\' order by Col2 desc, Col1 asc limit '+diffCount+'\"))&\": \"&query({R'+entryRowStart+'C1:R'+entryRowEnd+'C1,arrayformula(mmult(if(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+'=R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+',1,0),transpose(arrayformula(column(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+')\^0))))},\"select Col1 where Col1 \<\> \'\"\&R[0]C1\&\"\' order by Col2 desc, Col1 asc limit '+diffCount+
      '\")}))))');
  }

  // Sets the formula for home / away split for each matchup column
  for (let col = firstMatchCol; col <= finalMatchCol; col++ ) {
    sheet.getRange(summaryRow,col).setFormulaR1C1('=iferror(if(counta(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])>0,if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],regexextract(R'+matchRow+'C[0],"[A-Z]{2,3}"))=counta(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])/2,\"SPLIT\"&char(10)&\"50%\",if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],regexextract(R'+matchRow+'C[0],\"[A-Z]{2,3}\"))<counta(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])/2,regexextract(right(R'+matchRow+'C[0],3),\"[A-Z]{2,3}\")&char(10)&round(100*countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],regexextract(right(R'+matchRow+'C[0],3),\"[A-Z]{2,3}\"))/counta(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),1)&\"%\",regexextract(R'+matchRow+'C[0],\"[A-Z]{2,3}\")&char(10)&round(100*countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],regexextract(R'+matchRow+'C[0],\"[A-Z]{2,3}\"))/counta(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),1)&\"%\")),))');
  }
  
  if (tiebreaker) {
    sheet.getRange(matchRow,winCol).setFormulaR1C1('=if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)>1,\"TIE\",\"WIN\")');
    sheet.getRange(summaryRow,winCol).setFormulaR1C1('=iferror(if(not(isblank(R'+summaryRow+'C'+tiebreakerCol+')),if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)>1,countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)&\"\-WAY\"&char(10)&\"TIE\",),),)');
    sheet.getRange(summaryRow,tiebreakerCol).setFormulaR1C1('=iferror(if(sum(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])>0,\"AVG\"&char(10)&round(average(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),1),),)');
    sheet.getRange(summaryRow,tiebreakerCol+1).setFormulaR1C1('=iferror(if(sum(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])>0,\"AVG\"&char(10)&round(average(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),1),),)');
  } else {
    sheet.getRange(summaryRow,winCol).setFormulaR1C1('=iferror(if(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+')=value(regexextract(R'+dayRow+'C1,\"[0-9]+\")),if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)>1,countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)&\"\-WAY\"&char(10)&\"TIE\",\"DONE\"),),)');
    sheet.getRange(matchRow,winCol).setFormulaR1C1('=iferror(if(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+')=value(regexextract(R'+dayRow+'C1,\"[0-9]+\")),if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)=0,\"TIE\",\"WIN\"),\"WIN\"),)');
  }

  if (mnfInclude && mnf) {
    sheet.getRange(summaryRow,mnfCol).setFormulaR1C1('=iferror(if(counta(R'+outcomeRow+'C'+mnfStartCol+':R'+outcomeRow+'C'+mnfEndCol+')=columns(R'+outcomeRow+'C'+mnfStartCol+':R'+outcomeRow+'C'+mnfEndCol+'),\"MNF\"\&char(10)&(round(sum(mmult(arrayformula(if(R'+entryRowStart+'C'+mnfStartCol+':R'+entryRowEnd+'C'+mnfEndCol+'=R'+outcomeRow+'C'+mnfStartCol+':R'+outcomeRow+'C'+mnfEndCol+',1,0)),transpose(arrayformula(if(not(isblank(R'+outcomeRow+'C'+mnfStartCol+':R'+outcomeRow+'C'+mnfEndCol+')),1,0)))))/counta(R'+entryRowStart+'C'+mnfStartCol+':R'+entryRowEnd+'C'+mnfEndCol+'),3)*100)\&\"\%\",),)');
  }

  sheet.getRange(matchRow,pointsCol).setFormulaR1C1('=iferror(if(countif(R'+bonusRow+'C'+firstMatchCol+':R'+bonusRow+'C'+finalMatchCol+',\">1\")>0,\"TOTAL\"&char(10)&\"POINTS\",\"TOTAL\"&char(10)&\"CORRECT\"),)');

  sheet.getRange(summaryRow,pointsCol).setFormulaR1C1('=iferror(if(sum(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])>0,\"AVG\"\&char(10)&(round(average(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),1)),),)');

  sheet.getRange(summaryRow,diffCol).setFormulaR1C1('=iferror(if(isblank(R[0]C'+firstMatchCol+'),,transpose(query({arrayformula((counta(R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol+')-mmult(arrayformula(if(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+'=arrayformula(regexextract(R'+(totalMembers+3)+'C'+firstMatchCol+':R'+(totalMembers+3)+'C'+finalMatchCol+',\"[A-Z]+\")),1,0)),transpose(arrayformula(if(arrayformula(len(R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol+'))>1,1,1)))))&\": \"\&'+'R'+entryRowStart+'C1:R'+entryRowEnd+'C1),mmult(arrayformula(if(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+'=arrayformula(regexextract(R'+(totalMembers+3)+'C'+firstMatchCol+':R'+(totalMembers+3)+'C'+finalMatchCol+',\"[A-Z]+\")),1,0)),transpose(arrayformula(if(arrayformula(len(R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol+'))>1,1,1))))},\"select Col1 order by Col2 desc, Col1 desc limit '+diffCount+'\"))))');

  // AWAY TEAM BIAS FORMULA 
  sheet.getRange(summaryRow,2,1,1).setFormulaR1C1('=iferror(if(counta(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+')>10,"AWAY"&char(10)&round(100*(sum(arrayformula(if(regexextract(R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol+',"^[A-Z]{2,3}")=R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+',1,0)))/counta(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+')),1)&"%","AWAY"),"AWAY")');
  // HOME TEAM BIAS FORMULA
  sheet.getRange(summaryRow,3,1,1).setFormulaR1C1('=iferror(if(counta(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+')>10,"HOME"&char(10)&round(100*(sum(arrayformula(if(regexextract(R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol+',"[A-Z]{2,3}$")=R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+',1,0)))/counta(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+')),1)&"%","HOME"),"HOME")');
  sheet.getRange(summaryRow,4,1,1).setFormulaR1C1('=iferror(if(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+')>2,average(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),))');

  // Setting conditional formatting rules
  let bonusCount = 3;
  let parity = ['iseven','isodd'];
  let formatObj = [{'name':'correct_pick_even','color_start':'#c9ffdf','color_end':'#69ffa6','formula':'=and(indirect(\"R'+outcomeRow+'C[0]\",false)=indirect(\"R[0]C[0]\",false),not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false))),'+parity[0]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'correct_pick_odd','color_start':'#a0fdba','color_end':'#73ff9b','formula':'=and(indirect(\"R'+outcomeRow+'C[0]\",false)=indirect(\"R[0]C[0]\",false),not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false))),'+parity[1]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'incorrect_pick_even','color_start':'#FFF7F9','color_end':'#FCD4DC','formula':'=and(not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false))),'+parity[0]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'incorrect_pick_odd','color_start':'#FFF2F4','color_end':'#FFC3CC','formula':'=and(not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false))),'+parity[1]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'home_pick_even','color_start':'#e3fffe','color_end':'#9ef2ee','formula':'=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),arrayformula(trim(split(indirect(\"R'+matchRow+'C[0]\",false),\"\@\"))),0)=2,'+parity[0]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'home_pick_odd','color_start':'#d0f5f3','color_end':'#80f1ea', 'formula':'=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),arrayformula(trim(split(indirect(\"R'+matchRow+'C[0]\",false),\"\@\"))),0)=2,'+parity[1]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'away_pick_even','color_start':'#fffee3','color_end':'#fdf9a2','formula':'=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),arrayformula(trim(split(indirect(\"R'+matchRow+'C[0]\",false),\"\@\"))),0)=1,'+parity[0]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'away_pick_odd','color_start':'#faf9e1','color_end':'#fbf77f','formula':'=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),arrayformula(trim(split(indirect(\"R'+matchRow+'C[0]\",false),\"\@\"))),0)=1,'+parity[1]+'(row(indirect(\"R[0]C1\",false))))'}];

  sheet.clearConditionalFormatRules();    
  let range = sheet.getRange('R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol);
  Object.keys(formatObj).forEach(a => {
    let gradient = hexGradient(formatObj[a]['color_start'],formatObj[a]['color_end'],bonusCount);
    for (let b = gradient.length-1; b >= 0; b--) {
      let formula = formatObj[a]['formula'];
      if (b > 0) {
        // Appends the number bonus amount to the conditional formatting to pair with the gradient value assigned
        formula = formula.substring(0,formula.length-1).concat(',indirect(\"R'+bonusRow+'C[0]\",false)='+(b+1)+')');
      }
      let rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground(gradient[b])
        .setRanges([range]);
      if (formatObj[a]['name'].includes('incorrect')) {
        rule.setFontColor('#999999'); // Dark gray text for the incorrect picks
      }
      rule.build();
      
      formatRules.push(rule);
    }
  });

  // NAMES COLUMN NAMED RANGE
  range = sheet.getRange('R'+entryRowStart+'C1:R'+entryRowEnd+'C1');
  ss.setNamedRange('NAMES_'+week,range);

  // TOTALS GRADIENT RULE
  range = sheet.getRange('R'+entryRowStart+'C2:R'+entryRowEnd+'C2');
  ss.setNamedRange('TOT_'+week,range);
  let formatRuleTotals = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint("#75F0A1")
    .setGradientMinpoint("#FFFFFF")
    //.setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, (finalMatchCol-2) - 3) // Max value of all correct picks (adjusted by 3 to tighten color range)
    //.setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, (finalMatchCol-2) / 2)  // Generates Median Value
    //.setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, 0 + 3) // Min value of all correct picks (adjusted by 3 to tighten color range)
    .setRanges([range])
    .build();
  formatRules.push(formatRuleTotals);
  // RANKS GRADIENT RULE
  range = sheet.getRange('R'+entryRowStart+'C3:R'+entryRowEnd+'C3');
  ss.setNamedRange('RNK_'+week,range);
  let formatRuleRanks = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, members.length)
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, members.length/2)
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([range])
    .build();
  formatRules.push(formatRuleRanks);
  // PERCENT GRADIENT RULE
  range = sheet.getRange('R'+entryRowStart+'C4:R'+(rows)+'C4');
  range.setNumberFormat('##0.0%');
  let formatRulePercent = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, ".70")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, ".60")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, ".50")
    .setRanges([range])
    .build();
  formatRules.push(formatRulePercent);
  ss.setNamedRange('PCT_'+week,sheet.getRange('R'+entryRowStart+'C4:R'+entryRowEnd+'C4'));    
  // POINTS GRADIENT RULE
  range = sheet.getRange('R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol);
  let formatRulePoints = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect(\"R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]\",false))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect(\"R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]\",false))')
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect(\"R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]\",false))')
    .setRanges([range])
    .build();
  formatRules.push(formatRulePoints);


  // WINNER COLUMN RULE
  range = sheet.getRange('R'+entryRowStart+'C'+winCol+':R'+entryRowEnd+'C'+winCol);
  ss.setNamedRange('WIN_'+week,range);
  let formatRuleNotWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotEqualTo(1)
    .setBackground('#FFFFFF')
    .setFontColor('#FFFFFF')
    .setRanges([range])
    .build();     
  formatRules.push(formatRuleNotWinner);
  let formatRuleWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#75F0A1')
    .setFontColor('#75F0A1')
    .setRanges([range])
    .build();
  formatRules.push(formatRuleWinner);  
  // WINNER NAME RULE
  range = sheet.getRange('R'+entryRowStart+'C1:R'+entryRowEnd+'C1');
  let formatRuleWinnerName = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=indirect(\"R[0]C'+winCol+'\",false)=1')
    .setBackground('#75F0A1')
    .setRanges([range])
    .build();
  formatRules.push(formatRuleWinnerName);

  // MNF GRADIENT RULE
  let formatRuleMNFEmpty, formatRuleMNF;
  if (mnfInclude && mnf) {
    range = sheet.getRange('R'+entryRowStart+'C'+mnfCol+':R'+entryRowEnd+'C'+mnfCol);
    ss.setNamedRange('MNF_'+week,range);
    formatRuleMNFEmpty = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=or(isblank(indirect("R[0]C[0]",false)),indirect("R[0]C[0]",false)=0)')
      .setFontColor('#FFFFFF')
      .setBackground('#FFFFFF')
      .setRanges([range])
      .build();
    formatRules.push(formatRuleMNFEmpty);      
    if (mnfStartCol != mnfEndCol) { // Rules for when there are multiple MNF games
      formatRuleMNF = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint("#FFF624") // Max value of all correct picks, min 1
        .setGradientMinpoint("#FFFFFF") // Min value of all correct picks  
        .setRanges([range])
        .build();
    } else { // Rules for single MNF game 
      formatRuleMNF = SpreadsheetApp.newConditionalFormatRule()
        .setBackground("#FFF624")
        .setFontColor("#FFF624")
        .whenNumberEqualTo(1)
        .setRanges([range])
        .build();
    }
    formatRules.push(formatRuleMNF);
  }
 
  // DIFFERENCE TIEBREAKER COLUMN FORMATTING
  if (tiebreaker) {
    let offsets = [1,3,5,10,15,20,20];
    let offsetColors = hexGradient('#33FF7A','#FFFFFF',offsets.length);
    for (let a = 0; a < offsets.length; a++) {
      let rule;
      if (a < (offsets.length - 1)) {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false))),abs(indirect(\"R[0]C[0]\",false)-indirect(\"R'+outcomeRow+'C[0]:R'+outcomeRow+'C[0]\",false))<='+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(entryRowStart,tiebreakerCol,totalMembers,1)])
          .build();
      } else {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false))),abs(indirect(\"R[0]C[0]\",false)-indirect(\"R'+outcomeRow+'C[0]:R'+outcomeRow+'C[0]\",false))>'+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(entryRowStart,tiebreakerCol,totalMembers,1)])
          .build();        
      }
      formatRules.push(rule);
      rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false))),abs(value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))-indirect(\"R'+outcomeRow+'C[0]:R'+outcomeRow+'C[0]\",false))<='+offsets[a]+',)')
        .setBackground(offsetColors[a])
        .setRanges([sheet.getRange(summaryRow,tiebreakerCol)])
        .build();
      formatRules.push(rule);
    }
    offsetColors = hexGradient('#FFFFFF','#666666',offsets.length);
    for (let a = 0; a < offsets.length; a++) {
      let rule;
      let ruleOffsets;
      if (a < (offsets.length - 1)) {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[-1]\",false))),indirect(\"R[0]C[0]\",false)<='+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(entryRowStart,tiebreakerCol+1,totalMembers,1)])
          .build();
        ruleOffsets = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[-1]\",false))),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))<='+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(summaryRow,tiebreakerCol+1)])
          .build();
      } else {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[-1]\",false))),indirect(\"R[0]C[0]\",false)>'+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(entryRowStart,tiebreakerCol+1,totalMembers,1)])
          .build();
        ruleOffsets = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[-1]\",false))),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))>'+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(summaryRow,tiebreakerCol+1)])
          .build();              
      }
      formatRules.push(rule);
      formatRules.push(ruleOffsets);
    }
    // ADD ADDITIONAL COLOR VARIATION BASED ON TIEBREAKER VALUE PRESENT HERE
    let formatRuleTiebreakerEmptyAndDone = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=and(isblank(indirect(\"R[0]C[0]\",false)),counta(indirect(\"R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+'\",false))>=columns(indirect(\"R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+'\",false)))')
      .setBackground("#FF3FC7")
      .setRanges([sheet.getRange(outcomeRow,tiebreakerCol)])
      .build();
    formatRules.push(formatRuleTiebreakerEmptyAndDone);
    let formatRuleTiebreakerEmpty = SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty()
      .setBackground("#CCCCCC")
      .setRanges([sheet.getRange(outcomeRow,tiebreakerCol)])
      .build();
    formatRules.push(formatRuleTiebreakerEmpty);
    range = sheet.getRange(entryRowStart,tiebreakerCol,totalMembers,1);
    let formatRuleDiff = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpoint("#B7B7B7")
      .setGradientMinpoint("#FFFFFF")
      .setRanges([range])
      .build();
    formatRules.push(formatRuleDiff);
  }

  // PREFERENCE COLOR SCHEMES
  let homeAwayPercents = [90,80,70,60,50];
  let awayColors = ['#FFFB7D','#FFFC96','#FFFCB0','#FFFDC9','#FFFEE3'];
  let homeColors = ['#7DFFFB','#96FFFC','#B0FFFC','#C9FFFD','#E3FFFE'];
  let awayFormula = '=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R'+matchRow+'C[0]\",false),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=%%)'; // Replaceable "%%" for inserting percent number
  let homeFormula = '=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R'+matchRow+'C[0]\",false),3),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=%%)'; // Replaceable "%%" for inserting percent number
  range = sheet.getRange(summaryRow,firstMatchCol,1,matches); // Summary row of matches
  for (let a = 0; a < homeAwayPercents.length; a++) {
    let formula = awayFormula.replace('%%',homeAwayPercents[a]);

    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(awayColors[a])
      .setRanges([range]);
    rule.build();
    formatRules.push(rule);

    formula = homeFormula.replace('%%',homeAwayPercents[a]);
    rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(homeColors[a])
      .setRanges([range]);
    rule.build();
    formatRules.push(rule);    
  }

  // MATCHUP WEIGHTING RULE
  let formatRuleWeightedThree, formatRuleWeightedTwo;
  formatRuleWeightedThree = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),or(and(indirect(\"R'+bonusRow+'C[0]\",false)=2,countif(indirect(\"R'+bonusRow+'C'+firstMatchCol+':R'+bonusRow+'C'+finalMatchCol+'\",false),3)=0),indirect(\"R'+bonusRow+'C[0]\",false)=3))')
    .setBackground('#9C9C97')
    .setRanges([sheet.getRange('R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol),sheet.getRange('R'+outcomeRow+'C'+firstMatchCol+':R'+bonusRow+'C'+finalMatchCol)])
    .build();
  formatRules.push(formatRuleWeightedThree);
  formatRuleWeightedTwo = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),indirect(\"R'+bonusRow+'C[0]\",false)=2)')
    .setBackground('#949376')
    .setRanges([sheet.getRange('R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol),sheet.getRange('R'+outcomeRow+'C'+firstMatchCol+':R'+bonusRow+'C'+finalMatchCol)])
    .build();
  formatRules.push(formatRuleWeightedTwo);
  
  // Format rules for difference columns to emphasize the most common picker
  let commonPickersGradient = hexGradient('#46f081','#e4f0e8',8);
  let commonPickersFormula = '=value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))=%'; // Replaceable "%" for common picker number
  range = sheet.getRange(entryRowStart,diffCol,totalMembers+1,diffCount);
  for (let a = 0; a < commonPickersGradient.length; a++) {
    let formula = commonPickersFormula.replace('%',a); // Replaces "%" with index of commonPickersGradient (0-X)
    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(commonPickersGradient[a])
      .setRanges([range])
      .build();
    formatRules.push(rule);
  }

  // Sets all formerly pushed rules to the sheet
  sheet.setConditionalFormatRules(formatRules);

  // Setting size, alignment, frozen columns
  columns = sheet.getMaxColumns();
  sheet.getRange(1,1,rows,columns)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setFontSize(10)
    .setFontFamily("Montserrat");

  sheet.getRange(entryRowStart,diffCol,totalMembers+1,diffCount).setHorizontalAlignment('left');
  if (commentInclude) {
    sheet.getRange(2,commentCol,totalMembers+1,1).setHorizontalAlignment('left');
  }

  sheet.getRange(1,1,summaryRow,1)
    .setHorizontalAlignment('left');
 
  sheet.setFrozenColumns(firstMatchCol-1);
  sheet.setFrozenRows(dayRow);
  sheet.getRange(1,1,1,columns)
    .setBackground('black')
    .setFontColor('white')
    .setFontWeight('bold');
  sheet.setRowHeights(1,rows,21);

  sheet.getRange(matchRow,1,1,sheet.getMaxColumns()).setVerticalAlignment('middle');
  sheet.setRowHeight(matchRow,50);
  sheet.getRange(matchRow,1).setFontSize(18)
    .setHorizontalAlignment('center');
  
  sheet.getRange(dayRow,firstMatchCol,1,matches).setFontSize(7);
  sheet.getRange(dayRow,1,1,maxCols).setBackground('#CCCCCC');
  sheet.getRange(dayRow,firstMatchCol,1,dayRowColors.length).setBackgrounds([dayRowColors]);
  sheet.getRange(dayRow,1,1,firstMatchCol-1).mergeAcross();
  sheet.getRange(dayRow,1).setValue(matches + ' ' + league + ' MATCHES')
    .setHorizontalAlignment('left');
  
  sheet.getRange(outcomeRow,1,1,firstMatchCol-1).mergeAcross();
  sheet.getRange(outcomeRow,1,1,firstMatchCol-1).setHorizontalAlignment('right');  
  sheet.getRange(outcomeRow,1,1,maxCols).setBackground('black')
    .setFontColor('white')
    .setFontWeight('bold');
  sheet.getRange(bonusRow,1,1,maxCols).setBackground('black')
    .setFontColor('white');
  sheet.getRange(bonusRow,1,1,firstMatchCol-1).mergeAcross()
    .setHorizontalAlignment('right');
  if (!bonus) {
    sheet.hideRows(bonusRow);
  }

  sheet.setRowHeight(summaryRow,40);
  sheet.getRange(summaryRow,1,1,sheet.getMaxColumns()).setVerticalAlignment('middle');
  sheet.getRange(summaryRow,1,1,maxCols-diffCount).setBackground('#CCCCCC');
  sheet.getRange(summaryRow,2).setBackground(awayColors[1]);
  sheet.getRange(summaryRow,3).setBackground(homeColors[1]);

  sheet.setColumnWidths(diffCol,diffCount,90);
  sheet.getRange(1,diffCol,2,diffCount)
    .setHorizontalAlignment('left')
    .mergeAcross();

  if (tiebreaker) {
    sheet.getRange(outcomeRow,tiebreakerCol).setNote('Enter the summed score of the outcome of the final game of the week in this cell to complete the week and designate a winner');
  }
  sheet.getRange(dayRow,finalMatchCol+1,1,finalCol-finalMatchCol-diffCount).mergeAcross();

  // DATA RESTORATION
  if (dataRestore && !fresh) {
    let allPreviousPicks = [], allTiebreakers = [], allComments = [];
    for (let a = entryRowStart; a <= entryRowEnd; a++) {
      let previousPicks = [];
      let index = previousNames.indexOf(sheet.getRange(a,1).getValue());
      if (index >= 0) {
        // exportMatches format: [[day,away,home],.....]
        for (let b = 0; b < exportMatches.length; b++) {
          let away = exportMatches[b][1];
          let home = exportMatches[b][2];
          if (previousData[index].indexOf(away) >= 0) {
            previousPicks.push(away);
          } else if (previousData[index].indexOf(home) >= 0) {
            previousPicks.push(home);
          } else {
            previousPicks.push('');
          }
        }
        // If member found attempt to log tiebreaker and comment values, if present
        if (tiebreaker && tiebreakerCol > 5) {
          try {
            allTiebreakers.push(previousTiebreakers[index]);
          }
          catch (err) {
            allTiebreakers.push(['']);
          }
        } else {
          allTiebreakers.push(['']);
        }
        if (commentInclude && commentCol > 6) {
          try {
            allComments.push(previousComment[index]);
          }
          catch (err) {
            allComments.push(['']);
          }
        } else {
          allComments.push(['']);
        }
      } else {
        for (let b = 0; b < exportMatches; b++) {
          previousPicks.push('');
        }
      }
      allPreviousPicks.push(previousPicks);
    }
    
    // Fill out array with blank entries if a new user is added in the process of creating the sheet
    allPreviousPicks = makeArrayRectangular(allPreviousPicks);

    sheet.getRange(entryRowStart,firstMatchCol,allPreviousPicks.length,matches).setValues(allPreviousPicks);

    if (tiebreaker) {
      try {
        sheet.getRange(entryRowStart,tiebreakerCol,allTiebreakers.length,1).setValues(allTiebreakers);
      }
      catch (err) {
        Logger.log('Error setting tiebreaker column values or tiebreaker feature disabled.');
      }
    }

    try {
      sheet.getRange(entryRowStart,commentCol,allComments.length,1).setValues(allComments);
    }
    catch (err) {
      Logger.log('Error setting comment column values or comment feature disabled.');
    }
    let outcomes = [];
    for (let a = 0; a < exportMatches.length; a++) {
      let away = exportMatches[a][1];
      let home = exportMatches[a][2];
      if (previousOutcomes.indexOf(away) >= 0) {
        outcomes.push(away);
      } else if (previousOutcomes.indexOf(home) >= 0) {
        outcomes.push(home);
      } else {
        outcomes.push('');
      }
    }
    sheet.getRange(outcomeRow,firstMatchCol,1,matches).setValues([outcomes]);
    
    if (tiebreaker) {
      try {
        sheet.getRange(outcomeRow,tiebreakerCol).setValue(previousTiebreaker);
      }
      catch (err) {
        Logger.log('Unable to replace previous week tiebreaker value.');
      }
    }
    
    ss.toast('Previous data restored for week ' + week + ' if they were present');
    Logger.log('Previous data restored for week ' + week + ' if they were present');
  }
  // END DATA RESTORATION

  // Updates OUTCOMES sheet to reference weekly sheet for values along "outcomeRow"  
  outcomesSheetUpdate(ss,week);

  ss.toast('Completed creation of pick \'ems week ' + week + ' sheet.');
  return sheet;
}

// WEEKLY Sheet Creation Function - creates a sheet with provided week
function weeklySheetCreate(ss,next,restore) {
  ss = fetchSpreadsheet(ss);
  const ui = SpreadsheetApp.getUi();
  const weeks = fetchWeeks();
  
  if (next == undefined || next == null) {
    let all = [], next, weekString, missing = [], sheets = ss.getSheets();
    let regex = new RegExp('^'+weeklySheetPrefix+'[0-9]{1,2}'); // sheetName = weeklySheetPrefix [global var] + week [1-18 integer]
    for (let a = 1; a <= weeks; a++) {
      all.push(a);
      missing.push(a);
    }
    for (let a = 0; a < sheets.length; a++) {
      if (regex.test(sheets[a].getSheetName())) {
        let week = parseInt(sheets[a].getName().replace(weeklySheetPrefix,''));
        all[week-1] = '';
        missing.splice(missing.indexOf(week),1);
      }
    }
    next = all.lastIndexOf('') + 2; // Offset by 2: 1 for array indexing, 1 for going past last week
    let prompt, fresh = false;
    let ask = false;
    if (missing.length == 1) {
      weekString = 'Week ' + missing[0] + ' is the only one missing.';
    } else if (missing.length == 2) {
      weekString = 'Weeks ' + missing[0] + ' and ' + missing[1] + ' are missing.';
    } else if (missing.length > 2) {
      weekString = 'Weeks ' + missing[0];
      for (let a = 1; a < missing.length; a++) {
        if (a == missing.length - 1) {
          weekString = weekString.concat(', and ' + missing[a]);
        } else {
          weekString = weekString.concat(', ' + missing[a]);
        }
      }
      weekString = weekString.concat(' are missing.');  
    }
    if (next > 0) {
      prompt = ui.alert('Would you like to create a sheet for week ' + next + '?\r\n\r\n(Select \'NO\' to enter a different week number)', ui.ButtonSet.YES_NO_CANCEL);
      if (prompt == ui.Button.NO) {
        ask = true;
      } else if (prompt == ui.Button.CANCEL) {
        Logger.log('Canceled during prompt for creating week ' + next + '.');
        throw new Error('Canceled during prompt for creating week ' + next + '.');
      } else {
        fresh = true;
      }
    } else {
      prompt = ui.alert('All sheets for this season exist, would you like to recreate one of the sheets?', ui.ButtonSet.YES_NO);
      if (prompt == ui.Button.NO) {
        Logger.log('Canceled during prompt for creating another week since all weeks exist');
        throw new Error('Canceled during prompt for creating another week since all weeks exist');
      }
    }
    if (restore == null) {
      restore = false;
    }
    let confirm, other = 0;
    regex = new RegExp(/^[0-9]{1,2}/);
    let invalid = 'That week was invalid, please try again:';
    if (ask) {
      prompt = ui.prompt('Which sheet would you like to create or recreate?\r\n\r\n' + weekString, ui.ButtonSet.OK_CANCEL);
      other = prompt.getResponseText();
      let promptText = invalid;
      while (prompt.getSelectedButton() == 'OK' || !regex.test(other) || (other < 1 || other > weeks)) {
        while (prompt.getSelectedButton() == 'OK' && (!regex.test(other) || (other < 1 || other > weeks))) {
          prompt = ui.prompt(promptText, ui.ButtonSet.OK_CANCEL);
          other = prompt.getResponseText();
          promptText = invalid;
        }
        if (missing.indexOf(other) == -1 && prompt.getSelectedButton() == 'OK') {
          confirm = ui.alert('There is already a sheet for week ' + other + ' would you like to recreate it?\r\n\r\nClick \'NO\' to enter a different week.', ui.ButtonSet.YES_NO_CANCEL);
          restore = true;
        } else if (prompt.getSelectedButton() == 'OK') {
          confirm = ui.alert('Create a weekly sheet for week ' + other + '?\r\n\r\nClick \'NO\' to enter a different week.', ui.ButtonSet.YES_NO_CANCEL);
          restore = false;
        }
        if (confirm == ui.Button.YES) {
          next = other;
          break;
        } else if (confirm == ui.Button.NO) {
          other = 0;
          promptText = 'Try entering a different week value to create or recreate\r\n\r\n' + weekString;
        } else {
          break;
        }
      }
      if (prompt.getSelectedButton() != 'OK') {
        Logger.log('Canceled during prompt for entering custom week value');
        throw new Error('Canceled during prompt for entering custom week value');
      }
    }
    if (confirm == ui.Button.YES && restore) {
      ss.toast('Recreating sheet for week ' + next + ', data will be retained if possible');
      weeklySheet(ss,next,memberList(ss),restore);
    } else if(confirm == ui.Button.YES || fresh) {
      ss.toast('Creating a new sheet for week ' + next);
      weeklySheet(ss,next,memberList(ss),restore);
    }
  } else {
    ss.toast('Creating a sheet for week ' + next);
    if (restore == null) {
      restore = false;
    }
    weeklySheet(ss,next,memberList(ss),restore);
  }
  try {
    ss.toast('Refreshing formulas...');
    allFormulasUpdate(ss)
    ss.toast('Formulas refreshed succesfully');
  }
  catch (err) {
    ss.toast('Error refreshing formulas\r\n' + err.stack);
    Logger.log('Error refreshing formulas\r\n' + err.stack);
  }
}

// WEEKLY SHEET COLORATION - Adds a color to the weekly tabs that exist and uses the "dayColorsFilled" array [global variable]
function weeklySheetTabColors(ss,sheet) {
  let week;
  ss = fetchSpreadsheet(ss);
  if (sheet == undefined) {
    week = ss.getRangeByName('WEEK').getValue();
    sheet = ss.getSheetByName(weeklySheetPrefix + week);
  }
  try {
    if (sheet == undefined) {
      throw new Error();
    }
    let colors = [...dayColorsFilled];
    colors.push(winnersTabColor); // Adds a bright yellow to the end of the array for the active week tab
    let week = parseInt(sheet.getName().replace(weeklySheetPrefix,''));
    sheet.setTabColor(colors[colors.length-1]);
    colors.pop();
    for (let a = (week - 1); a > 0; a--) {
      let sheet = ss.getSheetByName(weeklySheetPrefix + a);
      if (sheet != null) {
        sheet.setTabColor(colors[colors.length-1]);
      }
      if (colors.length > 1) {
        colors.pop();
      }
    }
    Logger.log('changed all colors of tabs to reflect week shift');
  }
  catch (err) {
    Logger.log('Error assigning colors to weekly sheet tabs: ' + err.stack);
  }
}
