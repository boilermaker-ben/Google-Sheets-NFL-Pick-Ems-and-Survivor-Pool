//------------------------------------------------------------------------
// MEMBERS LIST - Creation and prompting of list of members to be stored on "MEMBERS" sheet
function memberList(ss,initial) {
  ss = fetchSpreadsheet(ss);
  let members = [];
  try {
    members = ss.getRangeByName('MEMBERS').getValues();
    if (members[0] == '') {
      throw new Error();
    }
    if (initial) {
      throw new Error();
    } else {
      return members;
    }
  } 
  catch (err) {
    Logger.log('No member list found, prompting for creation...');
    let ui = SpreadsheetApp.getUi();
    let valid = false;
    while (!valid) {
      let text = 'MEMBERS\r\n\r\nEnter a comma-separated list of members, more may be added later if you keep the membership unlocked.\r\n\r\nExample: \"Billy Joel, Hootie, Bon Jovi, Phil Collins\"\r\n\r\n';
      if (initial && members.length > 0) {
        text = 'MEMBERS\r\n\r\nExisting member list found: ' + members + '\r\n\r\nTo overwrite, enter a comma-separated list of members, more may be added later if you keep the membership unlocked.\r\n\r\nExample: \"Billy Joel, Hootie, Bon Jovi, Phil Collins\"\r\n\r\n'
      }
      let prompt = ui.prompt(text, ui.ButtonSet.OK_CANCEL);
      if ( prompt.getSelectedButton() == 'OK' ) {
        let arr = [];
        members = prompt.getResponseText().split(',');
        for (let a = 0; a < members.length; a++) {
          arr.push(members[a].trim());
        }
        members = arr;
        let duplicates = [];
        for (let a = 0; a < members.length; a++) {
          if (members.indexOf(members[a]) != -1 && members.indexOf(members[a]) != a && duplicates.indexOf(members[a]) == -1) {
            duplicates.push(' ' + members[a]);
          }
        }
        if (duplicates.length > 0) {
          ui.alert('DUPLICATE\r\n\r\nYou\'ve entered one or more duplicate names, try again and ensure each name is entered once.\r\n\r\nDuplicate(s): ' + duplicates, ui.ButtonSet.OK);
        } else if (members.length < 2) {
          ui.alert('MEMBER MINIMUM\r\n\r\nPlease enter at least 2 names', ui.ButtonSet.OK);
        } else {
          let text = '';
          for (let a = 0; a < members.length; a++) {
            text = text + members[a] + '\r\n';
          }
          prompt = ui.alert('MEMBERS\r\n\r\nThis is the list you entered:\r\n\r\n' + text + '\r\n\Would you like to proceed?', ui.ButtonSet.YES_NO);
          if (prompt == 'YES') {
            valid = true;
          }
        }
      } else {
        prompt = ui.alert('ALERT!\r\n\r\nIt is critical to create a member list for using this spreadsheet and form generator. Do you really want to cancel?', ui.ButtonSet.YES_NO);
        if (prompt == 'YES') {
          valid = true;
        }
        ss.toast('Restarting script for member list gathering.');     
      }
    }
    return members;
  }
}

// MEMBERS Addition for adding new members later in the season
function memberAdd(name) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let prompt;
  let membersSheet = ss.getSheetByName('MEMBERS');
  let range = ss.getRangeByName('MEMBERS');
  let members = range.getValues();
  const pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  const survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
  let mnfInclude;
  if (pickemsInclude) {
    mnfInclude = ss.getRangeByName('MNF_PRESENT').getValue();
  }
  let cancel = true;
  if (name == null) {
    prompt = ui.prompt('ADD MEMBER(S)\r\n\r\nPlease enter one member or a comma-separated list of members to add:', ui.ButtonSet.OK_CANCEL);
    name = prompt.getResponseText();
    if (prompt.getSelectedButton() == 'OK' && prompt.getResponseText() != null) {
      cancel = false;
    } else {
      ss.toast('Enter at least one name and click \"OK\" next time. Re-run \"Add Member(s)\" function to try again.');
    }
  } else {
    cancel = false;
  }
  if (name != null && !cancel) {
    let names = name.split(',');
    let arr = [];
    for (let a = 0; a < names.length; a++) {
      arr.push(names[a].trim());
    }
    names = [];
    for (let a = 0; a < arr.length; a++) {
      // Ensure no duplicate name is added
      if (members.flat().indexOf(arr[a]) == -1) {        
        members.push([arr[a]]);
        membersSheet.insertRows(1,1);
        range = membersSheet.getRange(1,1,membersSheet.getMaxRows(),1);
        range.setValues(members);
        ss.setNamedRange('MEMBERS',range);
        names.push(arr[a]);
      } else {
        prompt = ui.alert('MEMBER EXISTS\r\n\r\nA member with name ' + arr[a] + ' already exists.', ui.ButtonSet.OK);
        ss.toast('Unable to add ' + arr[a] + ' due to duplication.\r\n\r\nRe-run the \"Add Member(s)\" function again.');
      }
    }
    if (names.length > 0) {
      const week = fetchWeek(); 
      const weeks = fetchWeeks();
      // Update WEEKLY SHEETS
      if ( pickemsInclude ) {
        Logger.log('Working on week ' + week);
        weeklySheet(ss,week,members,true);
        ss.toast('Recreated weekly sheet for week ' + week);

        // Creates Weekly Totals Record Sheet
        totSheet(ss,weeks,members);
        Logger.log('Recreated Weekly Totals sheet');
        ss.toast('Recreated Weekly Totals sheet');

        // Creates Weekly Rank Record Sheet
        rnkSheet(ss,weeks,members);
        Logger.log('Recreated Weekly Rank sheet');
        ss.toast('Recreated Weekly Rank sheet');

        // Creates Weekly Percent Record Sheet
        pctSheet(ss,weeks,members);
        Logger.log('Recreated Weekly Percent sheet');
        ss.toast('Recreated Weekly Percent sheet');

        if ( mnfInclude ) {
          // Creates MNF Sheet
          mnfSheet(ss,weeks,members);
          Logger.log('Recreated MNF Sheet');
          ss.toast('Recreated MNF Sheet');
        }
      }

      if ( survivorInclude ) {
        // Creates Survivor Sheet
        survivorSheet(ss,weeks,members,true);
        Logger.log('Recreated Survivor sheet');
        ss.toast('Recreated Survivor sheet');

        survivorEvalSheet(ss,weeks,members,null);
        Logger.log('Recreated Survivor Eval sheet');
        ss.toast('Recreated Survivor Eval sheet');
      }

      // Creates Summary Record Sheet
      summarySheet(ss,members,pickemsInclude,mnfInclude,survivorInclude);
      Logger.log('Recreated Summary sheet');
      ss.toast('Recreated Summary sheet');
      
      memberAddForm(names,week);

      ss.toast('Completed addition of new member(s):\r\n\r\n' + names);
    } else {
      ss.toast('No new members added.');
    }
  } else {
    ss.toast('No new members added.');
  }
}

// MEMBERS Addition for adding new members later in the season
function memberRemove(name) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let prompt;
  let membersSheet = ss.getSheetByName('MEMBERS');
  let range = ss.getRangeByName('MEMBERS');
  let members = range.getValues();
  const pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  const survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
  let mnfInclude;
  if (pickemsInclude) {
    mnfInclude = ss.getRangeByName('MNF_PRESENT').getValue();
  }
  let cancel = true;
  if (name == null) {
    prompt = ui.prompt('Please type the name of the member you wish to remove:', ui.ButtonSet.OK_CANCEL);
    name = prompt.getResponseText().trim();
    if (prompt.getSelectedButton() == 'OK' && prompt.getResponseText() != null) {
      cancel = false;
    } else {
      ss.toast('No response was registered, try running again and entering the name of the member you wish to remove.');
    }
  } else {
    cancel = false;
  }
  if (name != null && !cancel && members.flat().indexOf(name) >= 0) {
    prompt = ui.alert('MEMBER FOUND\r\n\r\nFound member named ' + name + ', are you sure you want to remove this member?', ui.ButtonSet.YES_NO);
    if (prompt == ui.Button.YES) {
      membersSheet.deleteRow(members.flat().indexOf(name)+1);
      members.splice(members.flat().indexOf(name,1),1);
      range = membersSheet.getRange(1,1,membersSheet.getMaxRows(),1);
      range.setValues(members);
      ss.setNamedRange('MEMBERS',range);
      let rangeArr = [], names = [];
      if (pickemsInclude) {
        rangeArr = ['TOT_OVERALL_NAMES','TOT_RANKS_NAMES','TOT_PERCENT_NAMES'];
        if (mnfInclude) {
          rangeArr.push('MNF_NAMES');
        }
        nameRemove(rangeArr,name);
      }

      if (survivorInclude) {
        rangeArr = ['SURVIVOR_NAMES','SURVIVOR_EVAL_NAMES'];
        nameRemove(rangeArr,name);
      }

      let sheet = ss.getSheetByName('SUMMARY');
      range = sheet.getRange(1,1,sheet.getMaxRows(),1);
      names = range.getValues().flat();
      sheet.deleteRow(names.indexOf(name)+1);
      Logger.log('Deleted member ' + name + ' from SUMMARY sheet.');

      ss.toast('Completed removal of member: ' + name);
    } else {
      ss.toast('Member ' + name + ' not removed.');
    }
  } else {
    ss.toast('No member to remove.');
  }
  function nameRemove(rangeArr,name) {
    for (let a = 0; a < rangeArr.length; a++) {
      range = ss.getRangeByName(rangeArr[a]);
      let names = range.getValues().flat();
      let row = names.indexOf(name) + range.getRow();
      if (names.indexOf(name) >= 0) {
        range.getSheet().deleteRow(row);
        Logger.log('Deleted member ' + name + ' from ' + range.getSheet().getSheetName() + ' sheet.');
      }
    }
  }
}

// MEMBERS Addition for adding new members later in the season
function memberAddForm(names,week){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
  const survivorStart = ss.getRangeByName('SURVIVOR_START').getValue();

  if (week == null) {
    week = fetchWeek();
  }
  if (typeof names == 'string') {
    names = [names];
  } else if (names == null) {
    names = ['New User'];
  }
  let nameQuestion, gotoPage, newUserPage, found = false;
  try {
    let formId = ss.getRangeByName('FORM_WEEK_'+week).getValue();
    if (formId) {
      let form = FormApp.openById(formId);
      const items = form.getItems();
      for (let a = 0; a < items.length; a++) {
        if (items[a].getType() == 'LIST' && items[a].getTitle() == 'Name') {
          nameQuestion = items[a];
          found = true;
        } else if (items[a].getType() == 'PAGE_BREAK'){
          let pageBreakItem = items[a].asPageBreakItem();
          let pageTitle = pageBreakItem.getTitle();
          if (pageTitle == 'Survivor Start') {
            gotoPage = pageBreakItem;
          } else if (pageTitle == 'New User') {
            newUserPage = pageBreakItem;
          }
        }
      }
      if (found && nameQuestion) {
        let newChoice, choices = nameQuestion.asListItem().getChoices();
        if (survivorInclude && survivorStart == week) {
          try {
            for (let a = 0; a < names.length; a++) {
              if (names[a] == 'New User') {
                newChoice = nameQuestion.asListItem().createChoice(names[a],newUserPage);
                Logger.log('New user \"' + names[a] + '\" is redirected to the \"' + newUserPage.getTitle() + '\" Form page');
              } else {
                newChoice = nameQuestion.asListItem().createChoice(names[a],gotoPage);
                Logger.log('New user \"' + names[a] + '\" is redirected to the \"' + gotoPage.getTitle() + '\" Form page');
              }
              choices.unshift(newChoice);
              
            }
            nameQuestion.asListItem().setChoices(choices);
          }
          catch (err) {
            ss.toast('Issue locating survivor start question, you may need to add member manually');
            Logger.log('memberAdd error: ' + err.stack);
          }
        } else {
          try {
            for (let a = 0; a < names.length; a++) {
              if (names[a] == 'New User') {
                newChoice = nameQuestion.asListItem().createChoice(names[a],newUserPage);
                choices.unshift(newChoice);
                Logger.log('New user \"' + names[a] + '\" is redirected to the \"' + newUserPage.getTitle() + '\" Form page');
              } else {
                newChoice = nameQuestion.asListItem().createChoice(names[a],FormApp.PageNavigationType.SUBMIT);
                choices.unshift(newChoice);
                Logger.log('New user \"' + names[a] + '\" is redirected to the submit Form page');
              }
            }
            nameQuestion.asListItem().setChoices(choices);
          }
          catch (err) {
            ss.toast('Issue locating submit form value, you may need to add member manually');
            Logger.log('memberAdd error: ' + err.stack);
          }
        }
      }
    } else {
      Logger.log('No form created yet for week ' + week + ', skipping addition of ' + names + ' to form.');
    }
  }
  catch (err) {
    Logger.log(err.stack);
    ss.toast('Unable to add ' + names + ' to the form.');
  }
}
