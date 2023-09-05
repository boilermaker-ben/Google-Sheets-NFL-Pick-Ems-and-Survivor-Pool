// Google Sheets NFL Pick 'Ems & Survivor
// League Creator & Management Platform Tool
// v2.1 - 2023
// Created by Ben Powers
// ben.powers.creative@gmail.com

// CONSTANTS
const nflTeams = 32;
const maxGames = nflTeams/2;

// PRELIM SETUP- Creation of all needed initial sheets, prompt to import NFL
function runFirst() {
  
  // Initial Variables
  const year = fetchYear();
  const week = fetchWeek();
  const weeks = fetchWeeks();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const tz = ss.getSpreadsheetTimeZone();
  
  // Default group name
  let name = 'NFL Pick \'Ems';

  // Cue to go to spreadsheet for UI prompts from logger
  Logger.log('Answer the prompts that appear... [Go to spreadsheet]');

  // Prompt to start creation of spreadsheet
  let start = ui.alert('WELCOME\r\n\r\nThanks for checking out this Pick \'Ems script and using it. \r\n\r\nThere are a couple user inputs to gather before getting you rolling\r\n\r\n\- Ben', ui.ButtonSet.OK);
  let cancel = false;
  const cancelText = 'Setup canceled by user. Try again later.';
  if ( start == ui.Button.OK) {
    
    // Confirm timezone setting before continuing
    let timeZonePrompt = ui.alert('TIMEZONE\r\n\r\nThe timezone you\'re currently using is ' + tz + '. Is this correct?', ui.ButtonSet.YES_NO);
    if ( timeZonePrompt != 'YES') {
      let timeZoneFixPrompt = ui.alert('FIX TIMEZONE\r\n\r\nFollow these steps to change your projects time zone:\r\n\r\n1\. Return to the script editor\r\n2\. Select the gear icon on the left menu\r\n3\. Use the drop-down to select the correct timezone\r\n4\. Restart the script by clicking \'Run\' again', ui.ButtonSet.OK);
      ss.toast('Canceled due to incorrect time zone');
      throw new Error('Canceled during time zone confirmation question');
    }
    
    // Prompts to allow the user to create a league/pool/group name [defaults to NFL Pick 'Ems]
    let namePrompt = ui.alert('CUSTOMIZE NAME\r\n\r\nThe default name of the forms created will be called \"NFL Pick \'Ems\" or \"NFL Survivor Pool\", depending on your selections later. Do you want to change the name?', ui.ButtonSet.YES_NO);
    if ( namePrompt == ui.Button.YES) {
      // Loop to ensure name is acceptable
      let acceptance = false;
      let exit = false;
      while (acceptance == false && exit == false) {
        let namePrompt = ui.prompt('GROUP NAME\r\n\r\nWhat would you like to call your group?', ui.ButtonSet.OK);
        if (namePrompt.getSelectedButton() == ui.Button.OK && namePrompt.getResponseText() != '') {
          name = namePrompt.getResponseText();
        } else if ( namePrompt != ui.Button.OK ) {
          exit = true;
          cancel = true;
        }
        if (namePrompt.getResponseText() == '') {
          let retry = ui.alert('You didn\'t enter anything, want to try again?', ui.ButtonSet.YES_NO);
          if (retry == ui.Button.NO) {
            exit = true;
            cancel = true;
          } else {
            exit = false;
            acceptance = false;
          }
        } else {
          let examplePrompt = ui.alert('This is what your first form would be titled:\r\n\r\n'+name + ' - Week ' + week + ' - ' + year + '\r\n\r\nIs that correct?', ui.ButtonSet.YES_NO);
          if ( examplePrompt != ui.Button.YES && examplePrompt != ui.Button.NO ) {
            exit = true;
            cancel = true;
          }
          if ( examplePrompt == ui.Button.YES) {
            acceptance = true;
          }
        }
      }
    } else if ( namePrompt != ui.Button.YES && namePrompt != ui.Button.NO ) {
      cancel = true;
    }
    if (cancel == true) {
      ui.alert(cancelText, ui.ButtonSet.OK);
      throw new Error('Canceled during group naming question'); 
    }

    // Prompts for the inclusion of a pick 'ems contest
    let pickemsInclude = true;
    let pickemsCheck = ui.alert('PICK \'EMS\r\n\r\nThis script is intended to be used for running a weekly straight up pick \'ems style pool, but can be exclusively used for a survivor pool if desired.\r\n\r\nDo you intend to run a weekly pick \'ems pool?', ui.ButtonSet.YES_NO);
    if ( pickemsCheck == ui.Button.NO) {
      pickemsInclude = false;
      if (name == 'NFL Pick \'Ems') {
        name = 'NFL Survivor Pool';
      }
    } else if ( pickemsCheck != ui.Button.YES && pickemsCheck != ui.Button.NO ) {
      ui.alert(cancelText, ui.ButtonSet.OK);
      throw new Error('Canceled during pick \'ems prompt question'); 
    }

    // Prompts for the inclusion of a MNF tally
    let mnfInclude = false;
    if (pickemsInclude == true) {
      let mnfCheck = ui.alert('MONDAY NIGHT FOOTBALL\r\n\r\nWould you like to include a running tally of correct picks on Monday night football games?', ui.ButtonSet.YES_NO);
      if ( mnfCheck == ui.Button.YES) {
        mnfInclude = true;
      } else if ( mnfCheck != ui.Button.YES && mnfCheck != ui.Button.NO ) {
        ui.alert(cancelText, ui.ButtonSet.OK);
        throw new Error('Canceled during MNF question'); 
      }
    }

    // Prompts for the inclusion of a comment box at end of form
    let commentInclude = false;
    if (pickemsInclude == true) {
      let commentCheck = ui.alert('COMMENTS\r\n\r\nWould you like to include a comment box for members to include a note in their weekly submissions?', ui.ButtonSet.YES_NO);
      if ( commentCheck == ui.Button.YES ) {
        commentInclude = true;
      } else if ( commentCheck != ui.Button.YES && commentCheck != ui.Button.NO ) {
        ui.alert(cancelText, ui.ButtonSet.OK);
        throw new Error(); 
      }
    }

    // Prompts for the inclusion of a survivor pool
    let survivorInclude = true;
    let survivorStart = week;
    if (pickemsInclude == true) {
      let survivorCheck = ui.alert('SURVIVOR\r\n\r\nWould you like to include a survivor pool?', ui.ButtonSet.YES_NO);
      if ( survivorCheck == ui.Button.NO) {
        survivorInclude = false;
      }  else if ( survivorCheck != ui.Button.YES && survivorCheck != ui.Button.NO) {
        ui.alert(cancelText, ui.ButtonSet.OK);
        throw new Error('Canceled during survivor question'); 
      } else {
        if (week != 1) {
          ui.alert('Your survivor pool will start this week, week ' + week + ', rather than the standard starting point of week 1.', ui.ButtonSet.OK);
        }
      }
    }

    // Prompts for locking the number of participants
    let lockMembers = true;
    let lockMembersCheck = ui.alert('OPEN MEMBERSHIP\r\n\r\nAllow new members to be added to the pool through the Google Form?\r\n\r\n\(This can be changed later if you\'re not sure\)', ui.ButtonSet.YES_NO);
    if ( lockMembersCheck == ui.Button.YES) {
      lockMembers = false;
    } else if ( lockMembersCheck != ui.Button.YES && lockMembersCheck != ui.Button.NO ) {
      ui.alert(cancelText, ui.ButtonSet.OK);
      throw new Error('Canceled during lock membership question'); 
    }
    
    // Prompt if past week 1 to create previous week tables
    let oldWeeks, createOldWeeks = false;
    if (pickemsInclude == true) {
      if (week > 2) {
        oldWeeks = ui.alert('PREVIOUS WEEKS\r\n\r\nThere are ' + (week - 1) + ' previous weeks that could be created to manually enter old data, would you like to have those created?', ui.ButtonSet.YES_NO);
      } else if (week == 2) {
        oldWeeks = ui.alert('PREVIOUS WEEKS\r\n\r\nThere is a previous week that could be created to manually enter old data, would you like to create week 1 as well?', ui.ButtonSet.YES_NO);
      } 
      if ( oldWeeks == ui.Button.YES) {
        createOldWeeks = true;
      }
    }

    let createFormConfirm = false;
    let createForm = ui.alert('CREATE FORM\r\n\r\nCreate first Google Form after completing setup? \r\n\r\n\(You can still do this later through the menu\)', ui.ButtonSet.YES_NO);
    if (createForm == ui.Button.YES){
      createFormConfirm = true;
    }

    // Prompts for a set of member names to use
    let members = memberList();

    // Final prompt to start the longer script
    let text = 'Alright, that\'s it, now the script will do its thing!\r\n\r\nTimezone: ' + tz + '\r\nName: \"' + name + '\"\r\nPick \'Ems Pool: ' + (pickemsInclude==true?'YES':'NO');
    if (pickemsInclude == true) {
      text = text + '\r\nMNF Pool: ' + (mnfInclude==true?'YES':'NO') + '\r\nComments: ' + (commentInclude==true?'YES':'NO');
    }
    text = text + '\r\nSurvivor Pool: ' + (survivorInclude==true?'YES':'NO') + '\r\nMembers: ' + (lockMembers==true?'LOCKED':'UNLOCKED') + (week>1?('\r\nCreate Previous Weeks: ' + (createOldWeeks==true?'YES':'NO')):'') + '\r\nCreate Initial Form: ' + (createFormConfirm==true?'YES':'NO') + '\r\nInitial Member Count: ' + members.length;
    let finish = ui.alert(text, ui.ButtonSet.OK_CANCEL);
    if (finish == ui.Button.OK) {    
      // Pull in NFL Schedule data and create sheet
      fetchNFL();
      Logger.log('Fetched NFL Schedule');      
      
      // Run through all sheet information population
      try {
        // Creates Form Sheet (calls function)
        let config = configSheet(name,year,week,weeks,pickemsInclude,mnfInclude,commentInclude,survivorInclude,survivorStart);
        Logger.log('Deployed Config sheet');        

        // Creates Member sheet (calling function)
        memberSheet(members);
        Logger.log('Deployed Members sheet');
        
        members = memberList();
        // Creates winner selection sheet (NFL Outcomes)
        nflOutcomes(year);
        Logger.log('Deployed NFL Outcomes sheet');
        if (pickemsInclude == true) {
          // Creates Overall Record Sheet (calling function)
          overallSheet(year,weeks,members);
          Logger.log('Deployed Overall sheet');

          // Creates Overall Rank Record Sheet (calling function)
          overallRankSheet(year,weeks,members);
          Logger.log('Deployed Overall Rank sheet');
          
          // Creates Overall Percent Record Sheet (calling function)
          overallPctSheet(year,weeks,members);
          Logger.log('Deployed Overall Percent sheet');
        
          // Creates Winners Sheet (calling function)
          winnersSheet(year,weeks,members);
          Logger.log('Deployed Winners sheet');
          
          // Creates MNF Sheet (calling function)
          if (mnfInclude == true) {
            mnfSheet(year,weeks,members);
            Logger.log('Deployed MNF Sheet');
          }
        } else {
          // Deletes sheets if no Pickem's present if they were created by accident
          try {ss.deleteSheet(ss.getSheetByName('OVERALL'));} catch (err){}
          try {ss.deleteSheet(ss.getSheetByName('OVERALL_RANK'));} catch (err) {}
          try {ss.deleteSheet(ss.getSheetByName('OVERALL_PCT'));} catch (err) {}
          try {ss.deleteSheet(ss.getSheetByName('WINNERS'));} catch (err) {}
          try {ss.deleteSheet(ss.getSheetByName('MNF'));} catch (err) {}
        }
        if (survivorInclude == true) {
          // Creates Survivor Sheet (calling function)
          let survivor = survivorSheet(year,weeks,members,false);
          Logger.log('Deployed Survivor sheet');
          if (pickemsInclude == false) {
            survivor.activate();
          }
        } else {
          try{
            ss.deleteSheet(ss.getSheetByName('SURVIVOR'));
            } 
            catch (err) {}
        }
        
        // Creates Summary Record Sheet (calling function)
        summarySheet(year,members,pickemsInclude,mnfInclude,survivorInclude);
        Logger.log('Deployed Summary sheet');

        if (pickemsInclude == true) {    
          // Creates Weekly Sheets for the Current Week (calling function)
          let weekly = weeklySheet(year,week,members,false);
          Logger.log('Deployed Weekly sheet for week ' + week);
          weekly.activate();
        }       

        if (lockMembers == false) {
          createMenuUnlockedWithTriggerFirst();
        } else {
          createMenuLockedWithTriggerFirst();
        }
        Logger.log('Created final menu.');

        ss.getSheetByName('NFL_' + year).hideSheet();

      if (createFormConfirm == true){
        formCreate(true,week,year,name);
        Logger.log('Created initial form');
      }

      if (createOldWeeks == true && pickemsInclude == true) {
        try {
          for (let a = (week - 1); a > 1; a--) {
            weeklySheet(year,a,members,false);
          }
        }
        catch (err) {
          Logger.log('Issue creating previous weeks\r\n\r\n' + err.stack);
          ss.toast('Issue creating previous weeks pick \'ems sheets');
        }
      }
      let sheet = ss.getSheetByName('Sheet1');
      if ( sheet != null ) {
        ss.deleteSheet(sheet);
      }
      Logger.log('Deleted \'Sheet 1\'');

      config.hideSheet();
      Logger.log('You\'re all set, have fun!');

      }
      catch (err) {
        Logger.log('runFirstStack ' + err.stack);
      }
    } else {
      let canceled = ui.alert('You\'ve canceled the creation of the sheet and form. Do you really want to exit the initial setup?',ui.ButtonSet.YES_NO);
      if (canceled == ui.Button.NO) {
        runFirst();
      }
    }
  }
}

//------------------------------------------------------------------------
// CREATE MENU - this is the ideal setup once the sheet has been configured and the data is all imported
function createMenuUnlocked(trigger) {
  if (SpreadsheetApp.getActiveSpreadsheet().getRangeByName('PICKEMS_PRESENT').getValue() != false) {
    let menu = SpreadsheetApp.getUi().createMenu('Picks');
    menu.addItem('Create Form','formCreateAuto')
        .addItem('Open Form','openForm')
        .addItem('Check NFL Scores','recordNFLWeeklyScores')
        .addSeparator()
        .addItem('Check Responses','formCheckAlert')
        .addItem('Import Picks','dataTransfer')
        .addItem('Import Thursday Picks','dataTransferTNF')
        .addSeparator()
        .addItem('Add Member(s)','memberAdd')
        .addItem('Lock Members','createMenuLockedWithTrigger')
        .addSeparator()
        .addItem('Update NFL Schedule', 'fetchNFL')
        .addToUi();
  } else {
    let menu = SpreadsheetApp.getUi().createMenu('Picks');
    menu.addItem('Create Form','formCreateAuto')
        .addItem('Open Form','openForm')
        .addItem('Check NFL Scores','recordNFLWeeklyScores')
        .addSeparator()
        .addItem('Check Responses','formCheckAlert')
        .addItem('Import Picks','dataTransfer')
        .addSeparator()
        .addItem('Add Member(s)','memberAdd')
        .addItem('Lock Members','createMenuLockedWithTrigger')
        .addSeparator()
        .addItem('Update NFL Schedule', 'fetchNFL')
        .addToUi();
  }
  if (trigger == true) {
      deleteTriggers();
      let id = SpreadsheetApp.getActiveSpreadsheet().getId();
      ScriptApp.newTrigger('createMenuUnlocked')
        .forSpreadsheet(id)
        .onOpen()
        .create();
  }
}
// CREATE MENU For general use with locked MEMBERS sheet
function createMenuLocked(trigger) {
  if (SpreadsheetApp.getActiveSpreadsheet().getRangeByName('PICKEMS_PRESENT').getValue() != false) {
    let menu = SpreadsheetApp.getUi().createMenu('Picks');
    menu.addItem('Create Form','formCreateAuto')
        .addItem('Open Form','openForm')
        .addItem('Check NFL Scores','recordNFLWeeklyScores')
        .addSeparator()
        .addItem('Check Responses','formCheckAlert')
        .addItem('Import Picks','dataTransfer')
        .addItem('Import Thursday Picks','dataTransferTNF')
        .addSeparator()
        .addItem('Reopen Members','createMenuUnlockedWithTrigger')
        .addSeparator()
        .addItem('Update NFL Schedule', 'fetchNFL')  
        .addToUi();
  } else {
    let menu = SpreadsheetApp.getUi().createMenu('Picks');
    menu.addItem('Create Form','formCreateAuto')
        .addItem('Open Form','openForm')
        .addItem('Check NFL Scores','recordNFLWeeklyScores')
        .addSeparator()
        .addItem('Check Responses','formCheckAlert')
        .addItem('Import Picks','dataTransfer')
        .addSeparator()
        .addItem('Reopen Members','createMenuUnlockedWithTrigger')
        .addSeparator()
        .addItem('Update NFL Schedule', 'fetchNFL')  
        .addToUi();    
  }
  if (trigger == true) {
    deleteTriggers();
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
  memberAddForm(); // default action with no arguments is to add 'New User' to this week's form
  Logger.log('Menu updated to an open membership, MEMBERS unlocked');
  if (init != true) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('New entrants will be allowed through the Form and through the \"Picks\" menu function: \"Add Member(s)\". Run \"Lock Members\" to prevent new additions in the Form and menu.', SpreadsheetApp.getUi().ButtonSet.OK);
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
  removeNewUserQuestion(); // Removes 'New User' from Form
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MEMBERS').hideSheet();
  Logger.log('Menu updated to a locked membership, MEMBERS locked');
  if (init != true) {
    let ui = SpreadsheetApp.getUi();
    ui.alert('New entrants will not be allowed through the Form nor through the menu unless \"Reopen Members\" script is run. Run \"Reopen Members\" to allow new additions in the Form and menu', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
// CREATE MENU LOCKED MEMBERSHIP with Trigger Input on first pass (skips prompt)
function createMenuLockedWithTriggerFirst() {
  createMenuLockedWithTrigger(true);
}

//------------------------------------------------------------------------
// MEMBERS List for editing in future years
function memberList() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let members = [];
  try {
    members = ss.getRangeByName('MEMBERS').getValues();
    if (members[0] == '') {
      throw new Error();  
    }
    return members;
  } 
  catch (err) {
    Logger.log('No member list found, prompting for creation... [Go to spreadsheet]');
    let ui = SpreadsheetApp.getUi();
    
    let valid = false;
    while (valid == false) {
      let prompt = ui.prompt('MEMBERS\r\n\r\nEnter a comma-separated list of members, more may be added later if you keep the membership unlocked.\r\n\r\nExample: \"Billy Joel, Hootie, Bon Jovi, Phil Collins\"\r\n\r\n', ui.ButtonSet.OK_CANCEL);
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
          ui.alert('You\'ve entered one or more duplicate names, try again and ensure each name is entered once.\r\n\r\nDuplicate(s): ' + duplicates, ui.ButtonSet.OK);
        } else if (members.length < 2) {
          ui.alert('Please enter at least 2 names', ui.ButtonSet.OK);
        } else {
          let text = '';
          for (let a = 0; a < members.length; a++) {
            text = text + members[a] + '\r\n';
          }
          prompt = ui.alert('This is the list you entered:\r\n\r\n' + text + '\r\n\Would you like to proceed?', ui.ButtonSet.YES_NO);
          if (prompt == 'YES') {
            valid = true;
          }
        }
      } else {
        prompt = ui.alert('It is critical to create a member list for using this spreadsheet and form generator. Do you really want to cancel?', ui.ButtonSet.YES_NO);
        if (prompt == 'YES') {
          valid = true;
        }
        ss.toast('Restarting script for member list gathering.');     
      }
    }
    return members;
  }
}

//------------------------------------------------------------------------
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
  if (pickemsInclude == true) {
    mnfInclude = ss.getRangeByName('MNF_PRESENT').getValue();
  }
  const year = fetchYear();      
  const week = fetchWeek(); 
  const weeks = fetchWeeks();
  let cancel = true;
  if (name == null) {
    prompt = ui.prompt('Please enter one member or a comma-separated list of members to add:', ui.ButtonSet.OK_CANCEL);
    name = prompt.getResponseText();
    if (prompt.getSelectedButton() == 'OK' && prompt.getResponseText() != null) {
      cancel = false;
    } else {
      ss.toast('Enter at least one name and click \"OK\" next time. Re-run \"Add Member(s)\" function to try again.');
    }
  } else {
    cancel = false;
  }
  if (name != null && cancel == false) {
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
        prompt = ui.alert('A member with name ' + arr[a] + ' already exists.', ui.ButtonSet.OK);
        ss.toast('Unable to add ' + arr[a] + ' due to duplication.\r\n\r\nRe-run the \"Add Member(s)\" function again.');
      }
    }
    if (names.length > 0) {
      const year = fetchYear();      
      const week = fetchWeek(); 
      const weeks = fetchWeeks();
      // Update WEEKLY SHEETS
      if ( pickemsInclude == true) {
        Logger.log('Working on week ' + week);
        weeklySheet(year,week,members,true);
        ss.toast('Recreated weekly sheet for week ' + week);

        // Creates Overall Record Sheet (calling function)
        overallSheet(year,weeks,members);
        Logger.log('Recreated Overall sheet');

        // Creates Overall Rank Record Sheet (calling function)
        overallRankSheet(year,weeks,members);
        Logger.log('Recreated Overall Rank sheet');

        // Creates Overall Percent Record Sheet (calling function)
        overallPctSheet(year,weeks,members);
        Logger.log('Recreated Overall Percent sheet');
        
        // Creates Winners Sheet (calling function)
        winnersSheet(year,weeks,members);
        Logger.log('Recreated Winners sheet');

        if ( mnfInclude == true ) {
          // Creates MNF Sheet (calling function)
          mnfSheet(year,weeks,members);
          Logger.log('Recreated MNF Sheet');
        }
      }

      if ( survivorInclude == true ) {
        // Creates Survivor Sheet (calling function)
        survivorSheet(year,weeks,members,true);
        Logger.log('Recreated Survivor sheet');
        survivorEvalSheet(year,weeks,members,null);
        Logger.log('Recreated Survivor Eval sheet');
      }

      // Creates Summary Record Sheet (calling function)
      summarySheet(year,members,pickemsInclude,mnfInclude,survivorInclude);
      Logger.log('Recreated Summary sheet');

      memberAddForm(names,week);

      ss.toast('Completed addition of new member(s):\r\n\r\n' + names);
    } else {
      ss.toast('No new members added.');
    }
  } else {
    ss.toast('No new members added.');
  }
}

//------------------------------------------------------------------------
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
        if (survivorInclude == true && survivorStart == week) {
          try {
            for (let a = 0; a < names.length; a++) {
              if (names[a] == 'New User') {
                newChoice = nameQuestion.asListItem().createChoice(names[a],newUserPage);
                Logger.log('New user \"' + names[a] + '\" is redirected to the \"' + newUserPage.getTitle() + '\" Form page');
              } else {
                newChoice = nameQuestion.asListItem().createChoice(names[a],gotoPage);
                Logger.log('New user \"' + names[a] + '\" is redirected to the \"' + gotoPage.getTitle() + '\" Form page');
              }
              choices.push(newChoice);
              
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
                choices.push(newChoice);
                Logger.log('New user \"' + names[a] + '\" is redirected to the \"' + newUserPage.getTitle() + '\" Form page');
              } else {
                newChoice = nameQuestion.asListItem().createChoice(names[a],FormApp.PageNavigationType.SUBMIT);
                choices.push(newChoice);
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

//------------------------------------------------------------------------
// FETCH CURRENT YEAR
function fetchYear() {
  try {
    let year = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('YEAR').getValue();
    if (year != null) {
      return year;
    } else {
      try {
        const obj = JSON.parse(UrlFetchApp.fetch('http://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard').getContentText());
        year = obj.season.year;
        return year;
      }
      catch (err) {
        Logger.log('ESPN API has an issue right now');
      }
    }
  }
  catch (err) {
    const obj = JSON.parse(UrlFetchApp.fetch('http://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard').getContentText());
    let year = obj.season.year;
    return year;
  }
}

//------------------------------------------------------------------------
// FETCH CURRENT WEEK
function fetchWeek() {
  try {
    let week = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('WEEK').getValue();
    if (week != null) {
      return week;
    } else {
      try {
        const obj = JSON.parse(UrlFetchApp.fetch('http://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard').getContentText());
        let week = 1;
        if(obj.events[0].season.slug != 'preseason'){
          week = obj.week.number;
        }
        return week;
      }
      catch (err) {
        Logger.log('ESPN API has an issue right now');
      }
    }
  }
  catch (err) {
    try {
      const obj = JSON.parse(UrlFetchApp.fetch('http://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard').getContentText());
      let week = 1;
      if(obj.events[0].season.slug != 'preseason'){
        week = obj.week.number;
      }
      return week;
    }
    catch (err) {
      Logger.log('ESPN API has an issue right now');
    }
  }  
}

//------------------------------------------------------------------------
// FETCH TOTAL WEEKS
function fetchWeeks() {
  try {
    let weeks;
    const content = UrlFetchApp.fetch('http://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard').getContentText();
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

//------------------------------------------------------------------------
// ESPN TEAMS - Fetches the ESPN-available API data on NFL teams
function fetchTeamsESPN() {
  let year = fetchYear(); // First array value is year
  let obj = JSON.parse(UrlFetchApp.fetch('http://fantasy.espn.com/apis/v3/games/ffl/seasons/' + year + '?view=proTeamSchedules').getContentText());
  let objTeams = obj.settings.proTeams;
  return objTeams;
}

//------------------------------------------------------------------------
// NFL TEAM INFO - script to fetch all NFL data for teams
function fetchNFL() {
  // Calls the linked spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Declaration of script variables
  let abbr, name, maxRows, maxCols;
  const year = fetchYear();
  const objTeams = fetchTeamsESPN();
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
  let sheetName = 'NFL_' + year;
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
    ss.setNamedRange('NFL_'+year+'_AWAY_'+a,sheet.getRange(start,awayTeam,len,1));
    ss.setNamedRange('NFL_'+year+'_HOME_'+a,sheet.getRange(start,homeTeam,len,1));
  }
  sheet.protect().setDescription(sheetName);
  try {
    sheet.hideSheet();
  }
  catch (err){
    // Logger.log('fetchNFL hiding: Couldn\'t hide sheet as no other sheets exist');
  }
  ss.toast('Imported all NFL schedule data');
}

//------------------------------------------------------------------------
// NFL ACTIVE WEEK SCORES - script to check and pull down any completed matches and record them to the weekly sheet
function recordNFLWeeklyScores(){
  
  const outcomes = fetchNFLOutcomes();
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
  const year = fetchYear();
  let range;
  let weekMask = week < 10 ? '0' + week : week;
  let alert = 'CANCEL';
  if (done) {
    let text = 'WEEK ' + week + ' COMPLETE\r\n\r\nMark all game outcomes';
    if (pickemsInclude == true) {
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
    if (pickemsInclude == true) {
      let sheet,matchupRange,matchups,outcomeRange,outcomesRecorded,writeRange;
      try {
        sheet = ss.getSheetByName(year+'_'+weekMask);
        matchupRange = ss.getRangeByName('NFL_'+year+'_'+week);
        matchups = matchupRange.getValues().flat();
        outcomeRange = ss.getRangeByName('NFL_'+year+'_PICKEM_OUTCOMES_'+week);
        outcomesRecorded = outcomeRange.getValues().flat();
        writeRange = sheet.getRange(outcomeRange.getRow(),outcomeRange.getColumn(),1,matchups.length+1);
      }
      catch (err) {
        Logger.log(err.stack);
        ss.toast('Issue with fetching weekly sheet or named ranges on weekly sheet, recreating now.');
        weeklySheet(year,week,memberList(),false);
      }
      let arr = [];
      for (let a = 0; a < matchups.length; a++){
        let away = matchups[a].split('@')[0];
        let home = matchups[a].split('@')[1];
        let outcome;
        let regex = new RegExp('[A-Z]{2,3}');
        try {
          outcome = data.filter(game => game[0] == away && game[1] == home)[0];
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
        try {
          if (a == (matchups.length - 1)) {
            arr.push(outcome[3]); // Appends tiebreaker to end of array
          }
        }
        catch (err) {
          Logger.log('No tiebreaker yet');
          let tiebreakerCell = ss.getRangeByName('NFL_'+year+'_TIEBREAKER_'+week);
          let tiebreaker = sheet.getRange(tiebreakerCell.getRow()-1,tiebreakerCell.getColumn()).getValue();
          arr.push(tiebreaker);
        }
      }
      writeRange.setValues([arr]);
    } else if (survivorInclude == true) {
      let away = ss.getRangeByName('NFL_'+year+'_AWAY_'+week).getValues().flat();
      let home = ss.getRangeByName('NFL_'+year+'_HOME_'+week).getValues().flat();
      range = ss.getRangeByName('NFL_'+year+'_OUTCOMES_'+week);
      outcomesRecorded = range.getValues().flat();
      let arr = [];
      for (let a = 0; a < maxGames; a++) {
        arr.push([null]);
        for (let b = 0; b < away.length; b++) {
          if (data[b][0] == away[a] && data[b][1] == home[a]) {
            if (data[b][2] != null && outcomesRecorded[a] == null) {
              arr[a] = [data[b][2]];  
            } else {
              arr[a] = outcomesRecorded[a];
            }
          }
        }        
      }
      range.setValues(arr);
    }
  }
  if (done) {  
    if (survivorInclude == true) {
      let prompt = ui.alert('WEEK ' + week + ' COMPLETE \r\n\r\nAdvance survivor pool?', ui.ButtonSet.YES_NO); 
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

//------------------------------------------------------------------------
// NFL OUTCOMES - Records the winner and combined tiebreaker for each matchup on the NFL_{year} sheet
function fetchNFLOutcomes(){
  const ui = SpreadsheetApp.getUi();
  const obj = JSON.parse(UrlFetchApp.fetch('http://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard'));
  let games = obj.events;

  let year = fetchYear();
  if (year == null) {
    year = obj.season.year;
  }
  let week = obj.week.number;

  // Checks if preseason, if not, pulls in score data
  if(obj.events[0].season.slug == 'preseason'){
    ui.alert('Regular season not yet started.\r\n\r\n Currently preseason is still underway.', ui.ButtonSet.OK);
  } else {
    // Loop through games provided and creates an array for placing
    let headers = ['week','home','homeScore','away','awayScore','winner','tiebreaker'];
    let outcomes = [];
    let all = [];
    let count = 0;
    let away, awayScore,home, homeScore,tiebreaker,winner,competitors;
    for (let a = 0; a < games.length; a++){
      awayScore = '';
      homeScore = '';
      tiebreaker = '';
      winner = '';
      competitors = games[a].competitions[0].competitors;
      away = (competitors[1].homeAway == 'away' ? competitors[1].team.abbreviation : competitors[0].team.abbreviation);
      home = (competitors[0].homeAway == 'home' ? competitors[0].team.abbreviation : competitors[1].team.abbreviation);
      if (games[a].status.type.completed == true) {
        count++;
        awayScore = parseInt(competitors[1].homeAway == 'away' ? competitors[1].score : competitors[0].score);
        homeScore = parseInt(competitors[0].homeAway == 'home' ? competitors[0].score : competitors[1].score);
        tiebreaker = awayScore + homeScore;
        winner = (competitors[0].winner == true ? competitors[0].team.abbreviation : (competitors[1].winner == true ? competitors[1].team.abbreviation : 'TIE'));
        outcomes.push(home,away,winner,tiebreaker);
      }
      all.push(outcomes);
    }
    // Sets info variables for passing back to any calling functions
    let remaining = games.length - count;
    let completed = games.length - remaining;
    
    // Loop to add any empty rows of data
    for (let a = all.length; a < 16; a++) {
      outcomes = [week];
      for (let b = 1; b < headers.length; b++) {
        outcomes.push('');
      }
      all.push(outcomes);
    }
    // Outputs total matches, how many completed, and how many remaining;
    return [week,games.length,completed,remaining,all];
  }
}

//------------------------------------------------------------------------
// SHEET FOR LOGGING ALL OUTCOMES - creates a set of columns (one per week) on a sheet with a dedicated data validation rule per game to select from if not using import features
function nflOutcomes(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if ( year == null ) { 
    year = fetchYear();
  }
  const weeks = fetchWeeks();
  const sheetName = 'NFL_OUTCOMES';
  let sheet = ss.getSheetByName(sheetName);
  if ( sheet == null ) { sheet = ss.insertSheet(sheetName); }
  sheet.clearFormats();
  
  let data;
  try {
    data = ss.getRangeByName('NFL_'+year).getValues();
  }
  catch (err) {
    ss.toast('No NFL data, importing now');
    fetchNFL();
    data = ss.getRangeByName('NFL_'+year).getValues();
  }
  
  let headers = [];
  let headersWidth = [];
  let headerRow = 2;
  let gameCount = 16;

  for (let a = 1; a <= weeks; a++) {
    headers.push(a);
    headersWidth.push(60);
    ss.setNamedRange('NFL_'+year+'_OUTCOMES_'+a,sheet.getRange(headerRow+1,a,gameCount,1));
  }

  // Adjust the rows and columns of the sheet, and set maxCols/maxRows variables
  let maxCols = sheet.getMaxColumns();
  if (maxCols < headers.length) {
    sheet.insertColumnsAfter(maxCols,headers.length-maxCols);
  } else if (maxCols > headers.length) {
    sheet.deleteColumns(headers.length + 1, maxCols - headers.length);
  }
  maxCols = sheet.getMaxColumns();
  
  let rowTarget = (headerRow + gameCount); //  gameCount = 16 (max NFL Matches per week) data rows plus variable for headers
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
  let emptyArray = ['#fffdcc','#e7fed1','#cffdda','#bbfbe7','#adf7f5'];
  let filledArray = ['#fffb95','#d4ffa6','#abffbf','#89fddb','#74f7f3'];
  for (let row = 0; row < data.length; row++) {
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
    writeCell.setBackground(emptyArray[dayIndex]);
    awayWin.setBackground(filledArray[dayIndex]);
    homeWin.setBackground(filledArray[dayIndex]);
    awayWin.build();
    homeWin.build();
    formats.push(awayWin);
    formats.push(homeWin);
  }
  let ties = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('TIE')
    .setBold(false)
    .setBackground('#aaaaaa')
    .setRanges([matchups])
    .build();
  formats.push(ties);
  sheet.setConditionalFormatRules(formats);
  Logger.log('Completed setting up NFL Winners sheet');
}

//------------------------------------------------------------------------
// UPDATE OUTCOMES - Updates the data validation, color scheme, and matchups for a specific week on the NFL Winners sheet
function nflOutcomesUpdate(year,week,games) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (year == null) {
    year = fetchYear();
  }
  let sheet = ss.getSheetByName('NFL_OUTCOMES');
  if (sheet == null) {
    sheet = nflOutcomes(year);
  }

  let maxCols = sheet.getMaxColumns();
  
  // Clears data validation and notes
  let matchups = ss.getRangeByName('NFL_'+year+'_OUTCOMES_'+week);
  let rows = matchups.getNumRows();
  matchups.clearDataValidations();
  matchups.clearNote();

  let existingRules = sheet.getConditionalFormatRules();
  let rulesToKeep = [];
  let newRules = [];
  for (let a = 0; a < existingRules.length; a++) {
    let ranges = existingRules[a].getRanges();
    for (let b = 0; b < ranges.length; b++) {
      if (ranges[b].getA1Notation() != matchups) {
        rulesToKeep.push(existingRules[a]);
      }
    }
  }

  let emptyArray = ['#fffdcc','#e7fed1','#cffdda','#bbfbe7','#adf7f5'];
  let filledArray = ['#fffb95','#d4ffa6','#abffbf','#89fddb','#74f7f3'];
  
  for (let row = 0; row < games.length; row++) {
    let away = games[row][1];
    let home = games[row][2];
    
    let writeCell = sheet.getRange(row + 3,week);
    let rules = SpreadsheetApp.newDataValidation().requireValueInList([away,home,'TIE'], true).build();
    writeCell.setDataValidation(rules);
    let awayWin = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(away)
    .setBold(false)
    .setRanges([writeCell]);
    let homeWin = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(home)
    .setBold(true)
    .setRanges([writeCell]);
    // Color Coding Days
    let dayIndex = games[row][0] + 3; // Numeric day used for gradient application (-3 is Thursday, 1 is Monday);
    writeCell.setBackground(emptyArray[dayIndex]);
    awayWin.setBackground(filledArray[dayIndex]);
    homeWin.setBackground(filledArray[dayIndex]);
    awayWin.build();
    homeWin.build();
    newRules.push(awayWin);
    newRules.push(homeWin);
  }
  let ties = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('TIE')
    .setBold(false)
    .setBackground('#aaaaaa')
    .setRanges([sheet.getRange(matchups.getRow(),1,maxCols,rows)])
    .build();
  newRules.push(ties);

  let allRules = rulesToKeep.concat(newRules);
  //clear all rules first and then add again
  sheet.clearConditionalFormatRules();
  sheet.setConditionalFormatRules(allRules);
  let weeklySheetName = (year + '_' + week);
  if (week < 10) {
    weeklySheetName = (year + '_0' + week);
  }
  let sourceSheet = ss.getSheetByName(weeklySheetName);
  if (sourceSheet != null && ss.getRangeByName('PICKEMS_PRESENT').getValue() == true) {
    const targetSheet = ss.getSheetByName('NFL_OUTCOMES');
    const sourceRange = ss.getRangeByName('NFL_'+year+'_PICKEM_OUTCOMES_'+week);
    const targetRange = ss.getRangeByName('NFL_'+year+'_OUTCOMES_'+week);
    const row = sourceRange.getRow();
    let data = targetRange.getValues().flat();
    let regex = new RegExp('[A-Z]{2,3}');
    for (let a = 1; a <= games.length; a++) {
      if (regex.test(data[a-1])) {
        Logger.log('Found existing data at ' + (a+1) + ' of value ' + data[a-1]);
      } else {
        let formula = '=\''+weeklySheetName+'\'!'+sourceSheet.getRange(row,sourceRange.getColumn()+(a-1)).getA1Notation();
        targetSheet.getRange(targetRange.getRow()+(a-1),targetRange.getColumn()).setFormula(formula);        
      }
    }
  }  
}
function testConfig(){
  configSheet(null,2023,1,18,true,true,true,true,1);
}

//------------------------------------------------------------------------
function configSheet(name,year,week,weeks,pickemsInclude,mnfInclude,commentInclude,survivorInclude,survivorStart) {
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheetName = 'CONFIG';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName,0);
    sheet = ss.getSheetByName(sheetName);
  }
  try {
    if (pickemsInclude == null) {
      pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
    }
    if (mnfInclude == null) {
      mnfInclude = ss.getRangeByName('MNF_PRESENT').getValue();
    }
    if (survivorInclude == null) {
      survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
    }
    if (survivorInclude == true) {
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
  let array = [['NAME',name],['ACTIVE\ WEEK',week],['TOTAL\ WEEKS',weeks],['YEAR',year],['PICK\ \'EMS',pickemsInclude],['MNF',mnfInclude],['COMMENTS',commentInclude],['SURVIVOR',survivorInclude],['SURVIVOR\ DONE','=iferror(if(indirect(\"SURVIVOR_EVAL_REMAINING\")<=1,true,false))'],['SURVIVOR\ START',survivorStart]];
  let endData = array.length;
  let arrayNamedRanges = ['NAME','WEEK','WEEKS','YEAR','PICKEMS_PRESENT','MNF_PRESENT','COMMENTS_PRESENT','SURVIVOR_PRESENT','SURVIVOR_DONE','SURVIVOR_START'];

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

  // Rules for dropdowns on Config sheet
  let rule = SpreadsheetApp.newDataValidation().requireValueInList(weeksArr, true).build();
  sheet.getRange(2,2).setDataValidation(rule);
  
  rule = SpreadsheetApp.newDataValidation().requireValueInList([true,false], true).build();
  let range = sheet.getRange(5,2,4,1);
  range.setDataValidation(rule);
  rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(weeksArr)
    .build();
  sheet.getRange((endData-1),2).setDataValidation(rule);
  
  // TRUE COLOR FORMAT
  range = sheet.getRange(5,2,5,1);
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
  sheet.setColumnWidths(1,1,140);
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

//------------------------------------------------------------------------
// MEMBERS Sheet Creation / Adjustment 
function memberSheet(members) {
  
  if (members == null) {
    members = memberList();
  }
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalMembers = members.length;
  
  let sheetName = 'MEMBERS';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName,0);
    sheet = ss.getSheetByName(sheetName);
  }
  
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
  memberList();
  sheet.setColumnWidth(1,120);
  sheet.hideSheet();
  return sheet;
}

//------------------------------------------------------------------------
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

//------------------------------------------------------------------------
// MEMBERS Sheet Locking (protection)
function membersSheetLock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('MEMBERS');
  sheet.protect().setDescription('MEMBERS PROTECTION');
  Logger.log('locked MEMBERS');
}

//------------------------------------------------------------------------
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

//------------------------------------------------------------------------
// WEEKLY Sheet Function - creates a sheet with provided year and week
function weeklySheet(year,week,members,dataRestore) {
  
  if (year == null) {
    year = fetchYear();
  }
  if (week == null) {
    week = fetchWeek();
  }
  if (members == null){
    members = memberList();
  }
  let totalMembers = members.length;

  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let mnfInclude = ss.getRangeByName('MNF_PRESENT').getValue();
  let commentInclude = ss.getRangeByName('COMMENTS_PRESENT').getValue();

  let sheet, sheetName;
  let data = ss.getRangeByName('NFL_' + year).getValues(); //Grab again if wasn't populated before      
  
  let mnf = false;
  let tnf = false;
  let diffCount = (totalMembers - 1) >= 5 ? 5 : (totalMembers - 1); // Number of results to display for most similar weekly picks (defaults to 5, or 1 fewer than the total member count, whichever is larger)
  
  if ( week < 10 ) {
    sheetName = year + '_0' + week;
  } else {
    sheetName = year + '_' + week;
  }
  
  let rows = totalMembers + 3; // Accounting for the top two rows above member rows
  let columns;
  let fresh = false;
  sheet = ss.getSheetByName(sheetName);  
  if (sheet == null) {
    dataRestore = false;
    ss.insertSheet(sheetName,ss.getNumSheets()+1);
    sheet = ss.getSheetByName(sheetName);
    fresh = true;
  }

  let maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  
  // In case there were too few columns initially, adding extra
  if (maxCols < 24) {
    sheet.insertColumnsAfter(maxCols,24 - maxCols);
  }
  let tiebreakerCol, commentCol;
  let previousDataRange, previousData;
  let previousCommentRange, previousComment;
  if (dataRestore == true && fresh == false){
    let headers = sheet.getRange('A1:1').getValues().flat();
    headers.unshift('COL INDEX ADJUST');
    tiebreakerCol = headers.indexOf('TIEBREAKER');
    commentCol = headers.indexOf('COMMENT');
    if (tiebreakerCol  >= 0) {
      previousDataRange = sheet.getRange(2,5,maxRows-2,tiebreakerCol-4);
      previousData = previousDataRange.getValues();
      ss.toast('Gathered previous data for week ' + week + ', recreating sheet now');
    }    
    if (commentCol  >= 0 && commentInclude == true) {
      try {
        previousCommentRange = sheet.getRange(3,commentCol,maxRows-3,1);
        previousComment = previousCommentRange.getValues();
      }
      catch (err) {
        Logger.log('No previous comment data found to retain');
      }
    }
  }
  if (dataRestore == true && (tiebreakerCol == null || tiebreakerCol == -1)) {
    dataRestore = false;
    ss.toast('Intended to restore data, but no tiebreaker column found and therefore no data is believed to exist. Undo immediately if you want to retain information on sheet ' + sheetName);
  }
  sheet.clear();

  // Removing extra rows, reducing to only member count and the additional 2
  if (maxRows < rows){
    sheet.insertRows(maxRows,rows - maxRows);
  } else if (maxRows > rows){
    sheet.deleteRows(rows,maxRows - rows);
  }    
  
  // Insert Members
  sheet.getRange(3,1,totalMembers,1).setValues(members);
    
  // Setting header values
  let headersOne = ['WEEK '+week,'TOTAL','WEEKLY','PERCENT'];
  let headersTwo = ['CORRECT','RANK','CORRECT'];
  let bottomHeaders = ['PREFERRED','AWAY','HOME'];
  sheet.getRange(rows,1,1,3).setValues([bottomHeaders]);
  let widths = [120,90,90,90];

  // Setting headers for the week's matchups with format of 'AWAY' + '@' + 'HOME', then creating a data validation cell below each
  let firstMatchCol = (headersOne.indexOf('PERCENT'))+3;
  let mnfCol, mnfStartCol, mnfEndCol, tnfStartCol, tnfEndCol;
  let rule,matches = 0;
  let exportMatches = [];
  for ( let a = 0; a < data.length; a++ ) {
    if ( data[a][0] == week ) {
      matches++;
      let day = data[a][2];
      let away = data[a][6];
      let home = data[a][7];
      let matchup = away + '@' + home;
      exportMatches.push([day,away,home]);
      if ( day == 1 && mnfInclude == true) {
        mnf = true;
        if ( mnfStartCol == undefined ) {
          mnfStartCol = headersOne.length;
        }
        mnfEndCol = headersOne.length;
      } else if ( day == -3 ) {
        tnf = true;
        if ( tnfStartCol == undefined ) {
          tnfStartCol = headersOne.length + 1;
        }
        tnfEndCol = headersOne.length + 1;
      }
      headersOne.push(matchup);
      widths.push(75);
      rule = SpreadsheetApp.newDataValidation().requireValueInList([data[a][6],data[a][7]], true).build();
      sheet.getRange(2,headersOne.length).setDataValidation(rule);
    }
  }
  let finalMatchCol = headersOne.length;
  headersTwo.unshift(matches + ' NFL GAMES');

  headersOne = headersOne.concat(['TIEBREAKER','DIFFERENCE','WIN']);
  widths = widths.concat([100,100,50]); // Adding widths for above values
  tiebreakerCol =  headersOne.indexOf('TIEBREAKER')+1;

  if (mnfInclude == true && mnf == true) {
    headersOne.push('MNF');
    widths.push(50);
    mnfCol = headersOne.indexOf('MNF')+1;
  }
  if (commentInclude == true) {
    headersOne.push('COMMENT'); // Added to allow submissions to have amusing comments, if desired
    widths.push(125);
    commentCol = headersOne.indexOf('COMMENT')+1;
  }

  let diffCol = headersOne.length+1;
  let finalCol = diffCol + (diffCount-1);

  // Headers completed, now adjusting number of columns once headers are populated
  if (maxCols > finalCol){
    sheet.deleteColumns(finalCol+1,maxCols - finalCol);
  } else if (finalCol > maxCols) {
    sheet.insertColumnsAfter(maxCols,finalCol-maxCols);
  }
  maxCols = sheet.getMaxColumns();

  sheet.getRange(1,1,1,headersOne.length).setValues([headersOne]);
  sheet.getRange(2,1,1,headersTwo.length).setValues([headersTwo]);
  for (let a = 0; a < widths.length; a++) {
    sheet.setColumnWidth(a+1,widths[a]);
  }
  
  sheet.getRange(1,diffCol,2,1).setValues([['MOST SIMILAR'],['\[\# DIFFERENT\]']]); // Added to allow submissions to have amusing comments, if desired
  sheet.setColumnWidths(diffCol,diffCount,140);
  maxCols = sheet.getMaxColumns();

  let validRule = SpreadsheetApp.newDataValidation().requireNumberBetween(0,120)
    .setHelpText('Must be a number')
    .build();
  sheet.getRange(2,tiebreakerCol).setDataValidation(validRule);

  // Declare NFL Mathcups and Winners range for the week
  ss.setNamedRange('NFL_'+year+'_'+week,sheet.getRange(1,5,1,matches));
  ss.setNamedRange('NFL_'+year+'_PICKEM_OUTCOMES_'+week,sheet.getRange(2,5,1,matches));
  ss.setNamedRange('NFL_'+year+'_PICKS_'+week,sheet.getRange(3,5,totalMembers,matches));
  if (tnf == true) {
    ss.setNamedRange('NFL_'+year+'_THURS_PICKS_'+week,sheet.getRange(3,tnfStartCol,totalMembers,tnfEndCol-tnfStartCol+1));
  }
  if (mnfInclude == true) {
    ss.setNamedRange('NFL_'+year+'_MNF_'+week,sheet.getRange(3,mnfStartCol,totalMembers,mnfEndCol-(mnfStartCol-1)));
  }
  ss.setNamedRange('NFL_'+year+'_TIEBREAKER_'+week,sheet.getRange(3,tiebreakerCol,totalMembers,1));
  if (commentInclude == true) {
    ss.setNamedRange('COMMENTS_'+year+'_'+week,sheet.getRange(3,commentCol,totalMembers,1));
  }  

  for (let row = 3; row < rows; row++ ) {
    // Formula to determine how many correct on the week
    sheet.getRange(row,2).setFormulaR1C1('=iferror(if(and(counta(R2C[3]:R2C['+(finalMatchCol-2)+'])>0,counta(R[0]C[3]:R[0]C['+(finalMatchCol-2)+'])>0),mmult(arrayformula(if(R2C[3]:R2C['+(finalMatchCol-2)+']=R[0]C[3]:R[0]C['+(finalMatchCol-2)+'],1,0)),transpose(arrayformula(if(not(isblank(R[0]C[3]:R[0]C['+(finalMatchCol-2)+'])),1,0)))),))');
    // Formula to determine weekly rank
    sheet.getRange(row,3).setFormulaR1C1('=iferror(if(and(counta(R2C[2]:R2C['+(finalMatchCol-3)+'])>0,not(isblank(R[0]C[-1]))),rank(R[0]C[-1],R3C2:R'+(totalMembers+2)+'C2,false),))');
    // Formula to determine weekly correct percent
    sheet.getRange(row,4).setFormulaR1C1('=iferror(if(and(counta(R2C[1]:R2C['+(finalMatchCol-4)+'])>0,not(isblank(R[0]C[-2]))),R'+row+'C[-2]/counta(R2C[1]:R2C['+(finalMatchCol-4)+']),),)');
    // Formula to determine difference of tiebreaker from final MNF score
    sheet.getRange(row,finalMatchCol+2).setFormulaR1C1('=iferror(if(or(isblank(R[0]C[-1]),isblank(R2C'+(finalMatchCol+1)+')),,abs(R[0]C[-1]-R2C'+(finalMatchCol+1)+')))');
    // Formula to denote winner with a '1'
    sheet.getRange(row,finalMatchCol+3).setFormulaR1C1('=iferror(if(sum(arrayformula(if(isblank(R2C5:R2C'+(finalMatchCol+1)+'),1,0)))>0,,match(R[0]C1,filter(filter(R3C1:R'+(totalMembers+2)+'C1,R3C2:R'+(totalMembers+2)+'C2=max(R3C2:R'+(totalMembers+2)+'C2)),filter(R3C[-1]:R'+(totalMembers+2)+'C[-1],R3C2:R'+(totalMembers+2)+'C2=max(R3C2:R'+(totalMembers+2)+'C2))=min(filter(R3C[-1]:R'+(totalMembers+2)+'C[-1],R3C2:R'+(totalMembers+2)+'C2=max(R3C2:R'+(totalMembers+2)+'C2)))),0)^0))');
    // Formula to determine MNF win status sum (can be more than 1 for rare weeks)
    if ( mnfInclude == true && mnf == true ) {
      sheet.getRange(row,mnfCol).setFormulaR1C1('=iferror(if(and(counta(R2C'+firstMatchCol+':R2C'+finalMatchCol+')>0,counta(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+')>0),if(mmult(arrayformula(if(R2C'+mnfStartCol+':R2C'+mnfEndCol+'=R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+',1,0)),transpose(arrayformula(if(not(isblank(R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+')),1,0))))=0,0,mmult(arrayformula(if(R2C'+mnfStartCol+':R2C'+mnfEndCol+'=R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+',1,0)),transpose(arrayformula(if(not(isblank(R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+')),1,0))))),),)');
    }
    // Formula to generate array of similar pickers on the week
    sheet.getRange(row,diffCol).setFormulaR1C1('=iferror(if(isblank(R[0]C5),,transpose(arrayformula({query({R3C1:R'+(totalMembers+2)+'C1,arrayformula(mmult(if(R3C5:R'+(totalMembers+2)+'C'+(finalMatchCol)+'=R[0]C5:R[0]C'+(finalMatchCol)+',1,0),transpose(arrayformula(column(R[0]C5:R[0]C'+(finalMatchCol)+')\^0))))},\"select Col1 where Col1 \<\> \'\"\&R[0]C1\&\"\' order by Col2 desc, Col1 asc limit '+diffCount+
      '\")\&\" [\"\&arrayformula('+(finalMatchCol-2)+'-query({R3C1:R'+(totalMembers+2)+'C1,arrayformula(mmult(if(R3C5:R'+(totalMembers+2)+'C'+(finalMatchCol)+'=R[0]C5:R[0]C'+(finalMatchCol)+',1,0),transpose(arrayformula(column(R[0]C5:R[0]C'+(finalMatchCol)+')\^0))))},\"select Col2 where Col1 <> \'\"\&R[0]C1\&\"\' order by Col2 desc, Col1 asc limit '+diffCount+'\"))-2\&\"]\"}))))');
  }

  sheet.getRange(rows,1,1,maxCols).setBackground('#dbdbdb');
  sheet.getRange(rows,2).setBackground('#fffee3');
  sheet.getRange(rows,3).setBackground('#e3fffe'); 
  let cellsPopulatedCheck;
  for (let row = 5; row <= finalMatchCol; row++ ) {
    if (totalMembers >= 3) { // adjusts an if statement conditional for varying amounts of members
      cellsPopulatedCheck = 'or(not(isblank(R3C[0])),not(isblank(R4C[0])),not(isblank(R5C[0])))';
    } else if (totalMembers == 2){
      cellsPopulatedCheck = 'or(not(isblank(R3C[0])),not(isblank(R4C[0])))';
    } else if (totalMembers == 1) {
      cellsPopulatedCheck = 'not(isblank(R3C[0]))';
    }
    sheet.getRange(rows,row).setFormulaR1C1('=iferror(if(counta(R3C[0]:R[-1]C[0])>0,if(countif(R3C[0]:R'+(totalMembers+2)+'C[0],regexextract(R1C[0],"[A-Z]{2,3}"))=counta(R3C[0]:R'+(totalMembers+2)+'C[0])/2,"SPLIT",if(countif(R3C[0]:R'+(totalMembers+2)+'C[0],regexextract(R1C[0],"[A-Z]{2,3}"))<counta(R3C[0]:R'+(totalMembers+2)+'C[0])/2,regexextract(right(R1C[0],3),"[A-Z]{2,3}")&"|"&round(100*countif(R3C[0]:R'+(totalMembers+2)+'C[0],regexextract(right(R1C[0],3),"[A-Z]{2,3}"))/counta(R3C[0]:R'+(totalMembers+2)+'C[0]),1)&"%",regexextract(R1C[0],"[A-Z]{2,3}")&"|"&round(100*countif(R3C[0]:R'+(totalMembers+2)+'C[0],regexextract(R1C[0],"[A-Z]{2,3}"))/counta(R3C[0]:R'+(totalMembers+2)+'C[0]),1)&"%")),))');
  }
  sheet.getRange(rows,diffCol).setFormulaR1C1('=iferror(if(isblank(R[0]C5),,transpose(query({arrayformula(R3C1:R'+(totalMembers+2)+'C1&\" [\"&(counta(R1C5:R1C'+finalMatchCol+')-mmult(arrayformula(if(R3C5:R'+(totalMembers+2)+'C'+finalMatchCol+'=arrayformula(regexextract(R'+(totalMembers+3)+'C5:R'+(totalMembers+3)+'C'+finalMatchCol+',\"[A-Z]+\")),1,0)),transpose(arrayformula(if(arrayformula(len(R1C5:R1C'+finalMatchCol+'))>1,1,1)))))&\"]\"),mmult(arrayformula(if(R3C5:R'+(totalMembers+2)+'C'+finalMatchCol+'=arrayformula(regexextract(R'+(totalMembers+3)+'C5:R'+(totalMembers+3)+'C'+finalMatchCol+',\"[A-Z]+\")),1,0)),transpose(arrayformula(if(arrayformula(len(R1C5:R1C'+finalMatchCol+'))>1,1,1))))},\"select Col1 order by Col2 desc, Col1 desc limit '+diffCount+'\"))))');
  
 // AWAY TEAM BIAS FORMULA 
  sheet.getRange(rows,2,1,1).setFormulaR1C1('=iferror(if(counta(R3C5:R'+(totalMembers+2)+'C'+finalMatchCol+')>10,"AWAY|"&round(100*(sum(arrayformula(if(regexextract(R1C5:R1C'+finalMatchCol+',"^[A-Z]{2,3}")=R1C5:R'+(totalMembers+2)+'C'+finalMatchCol+',1,0)))/counta(R3C5:R'+(totalMembers+2)+'C'+finalMatchCol+')),1)&"%","AWAY"),"AWAY")');
  // HOME TEAM BIAS FORMULA
  sheet.getRange(rows,3,1,1).setFormulaR1C1('=iferror(if(counta(R3C5:R'+(totalMembers+2)+'C'+finalMatchCol+')>10,"HOME|"&round(100*(sum(arrayformula(if(regexextract(R1C5:R1C'+finalMatchCol+',"[A-Z]{2,3}$")=R1C5:R'+(totalMembers+2)+'C'+finalMatchCol+',1,0)))/counta(R3C5:R'+(totalMembers+2)+'C'+finalMatchCol+')),1)&"%","HOME"),"HOME")');
  sheet.getRange(rows,4,1,1).setFormulaR1C1('=iferror(if(counta(R2C[1]:R2C['+(finalMatchCol-4)+'])>2,average(R2C[0]:R'+(totalMembers+2)+'C[0]),))');

  // ALTERNATING COLORS / BANDING
  let range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  range.getBandings().forEach(banding => banding.remove());

  sheet.getRange(3,1,totalMembers,finalCol).applyRowBanding(SpreadsheetApp.BandingTheme)
    .setHeaderRowColor(null)
    .setSecondColumnColor('#F0F0F0');

  // Setting conditional formatting rules
  sheet.clearConditionalFormatRules();    
  range = sheet.getRange('R3C5:R'+(rows-1)+'C'+finalMatchCol);
  // CORRECT PICK COLOR RULE
  let formatRuleCorrectEven = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R2C[0]\",false)=indirect(\"R[0]C[0]\",false),not(isblank(indirect(\"R2C[0]\",false))),iseven(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#c9ffdf')
    .setRanges([range])
    .build();
  let formatRuleCorrectOdd = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R2C[0]\",false)=indirect(\"R[0]C[0]\",false),not(isblank(indirect(\"R2C[0]\",false))),isodd(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#bbedd0')
    .setRanges([range])
    .build();
  // INCORRECT PICK COLOR RULE
  let formatRuleIncorrectEven = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(indirect(\"R2C[0]\",false)=indirect(\"R[0]C[0]\",false)),not(isblank(indirect(\"R2C[0]\",false))),iseven(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#ffc4ca')
    .setStrikethrough(true)
    .setRanges([range])
    .build();
  let formatRuleIncorrectOdd = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(indirect(\"R2C[0]\",false)=indirect(\"R[0]C[0]\",false)),not(isblank(indirect(\"R2C[0]\",false))),isodd(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#f2bdc2')
    .setStrikethrough(true)
    .setRanges([range])
    .build();
  // HOME PICK COLOR RULE
  let formatRuleHomeEven = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),split(indirect(\"R1C[0]\",false),\"\@\"),0)=2,iseven(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#e3fffe')
    .setRanges([range])
    .build();
  let formatRuleHomeOdd = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),split(indirect(\"R1C[0]\",false),\"\@\"),0)=2,isodd(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#d0f5f3')
    .setRanges([range])
    .build();
  // AWAY PICK COLOR RULE
  let formatRuleAwayEven = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),split(indirect(\"R1C[0]\",false),\"\@\"),0)=1,iseven(row(indirect(\"R[0]C1\",false))))')
    .setBackground('#fffee3')
    .setRanges([range])
    .build();
  let formatRuleAwayOdd = SpreadsheetApp.newConditionalFormatRule()
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
  let formatRuleTotals = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint("#75F0A1")
    .setGradientMinpoint("#FFFFFF")
    //.setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, (finalMatchCol-2) - 3) // Max value of all correct picks (adjusted by 3 to tighten color range)
    //.setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, (finalMatchCol-2) / 2)  // Generates Median Value
    //.setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, 0 + 3) // Min value of all correct picks (adjusted by 3 to tighten color range)
    .setRanges([range])
    .build();
  // RANKS GRADIENT RULE
  range = sheet.getRange('R3C3:R'+(rows-1)+'C3');
  ss.setNamedRange('RANK_'+year+'_'+week,range);
  let formatRuleRanks = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, members.length)
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, members.length/2)
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([range])
    .build();
  // PERCENT GRADIENT RULE
  range = sheet.getRange('R3C4:R'+(rows)+'C4');
  range.setNumberFormat('##0.0%');
  let formatRulePercent = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, ".70")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, ".60")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, ".50")
    .setRanges([range])
    .build();
  ss.setNamedRange('PCT_'+year+'_'+week,sheet.getRange('R3C4:R'+(rows-1)+'C4'));    
  
  // WINNER COLUMN RULE
  range = sheet.getRange('R3C'+(finalMatchCol+3)+':R'+(rows-1)+'C'+(finalMatchCol+3));
  ss.setNamedRange('WIN_'+year+'_'+week,range);
  let formatRuleNotWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotEqualTo(1)
    .setBackground('#FFFFFF')
    .setFontColor('#FFFFFF')
    .setRanges([range])
    .build();     
  let formatRuleWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#75F0A1')
    .setFontColor('#75F0A1')
    .setRanges([range])
    .build();
  
  // WINNER NAME RULE
  range = sheet.getRange('R3C1:R'+rows+'C1');
  let formatRuleWinnerName = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=indirect(\"R[0]C'+(finalMatchCol+3)+'\",false)=1')
    .setBackground('#75F0A1')
    .setRanges([range])
    .build();  
  

  // MNF GRADIENT RULE
  let formatRuleMNFEmpty, formatRuleMNF;
  if (mnfInclude == true && mnf == true) {
    range = sheet.getRange('R3C'+(finalMatchCol+4)+':R'+(rows-1)+'C'+(finalMatchCol+4));
    ss.setNamedRange('MNF_'+year+'_'+week,range);
    formatRuleMNFEmpty = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=isblank(indirect("R[0]C[0]",false))')
      .setFontColor('#FFFFFF')
      .setBackground('#FFFFFF')
      .setRanges([range])
      .build();
    formatRuleMNF = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpoint("#C2FF7D") // Max value of all correct picks, min 1
      .setGradientMinpoint("#FFFFFF") // Min value of all correct picks  
      .setRanges([range])
      .build();
  }
  
  range = sheet.getRange(3,diffCol,totalMembers+1,diffCount);
  let formatCommonPicker0 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))=0')
    .setBackground('#46f081')
    .setRanges([range])
    .build();
  let formatCommonPicker1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))=1')
    .setBackground('#75F0A1')
    .setRanges([range])
    .build();
  let formatCommonPicker2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))=2')
    .setBackground('#a4edbe')
    .setRanges([range])
    .build();
  let formatCommonPicker3 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))=3')
    .setBackground('#e4f0e8')
    .setRanges([range])
    .build();
  
  // DIFFERENCE TIEBREAKER COLUMN FORMATTING
  range = sheet.getRange(3,finalMatchCol+2,totalMembers,1);
  let formatRuleDiff = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint("#FFFFFF")
    .setGradientMinpoint("#5EDCFF")
    .setRanges([range])
    .build();
  
  // PREFERENCE COLOR SCHEMES
  range = sheet.getRange(rows,4,1,finalMatchCol-3);
  // Away Favored 90%
  let formatRuleAway90 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R1C[0]\",false),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=90)')
    .setBackground('#fffb7d')
    .setRanges([range])
    .build();
  // Home Favored 90%
  let formatRuleHome90 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R1C[0]\",false),3),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=90)')
    .setBackground('#7dfffb')
    .setRanges([range])
    .build();
  // Away favored 80%
  let formatRuleAway80 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R1C[0]\",false),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=80)')
    .setBackground('#fffc96')
    .setRanges([range])
    .build();
  // Home Favored 80%
  let formatRuleHome80 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R1C[0]\",false),3),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=80)')
    .setBackground('#96fffc')
    .setRanges([range])
    .build();
  // Away Favored 70%
  let formatRuleAway70 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R1C[0]\",false),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=70)')
    .setBackground('#fffcb0')
    .setRanges([range])
    .build();
  // Home Favored 70%
  let formatRuleHome70 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R1C[0]\",false),3),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=70)')
    .setBackground('#b0fffc')
    .setRanges([range])
    .build();
  // Away Favored 60%
  let formatRuleAway60 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R1C[0]\",false),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=60)')
    .setBackground('#fffdc9')
    .setRanges([range])
    .build();
  // Home Favored 60%
  let formatRuleHome60 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R1C[0]\",false),3),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=60)')
    .setBackground('#c9fffd')
    .setRanges([range])
    .build();
  // Away Favored
  let formatRuleAway50 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R1C[0]\",false),\"[A-Z]{2,3}\")')
    .setBackground('#fffee3')
    .setRanges([range])
    .build();
  // Home Favored
  let formatRuleHome50 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R1C[0]\",false),3),\"[A-Z]{2,3}\")')
    .setBackground('#e3fffe')
    .setRanges([range])
    .build();

  let formatRules = sheet.getConditionalFormatRules();
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
  if (mnfInclude == true && mnf == true) {
    formatRules.push(formatRuleMNFEmpty);
    formatRules.push(formatRuleMNF);
  }
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
  range = sheet.getRange(1,1,rows,columns);
  range.setFontSize(10);
  range.setVerticalAlignment('middle');
  range.setHorizontalAlignment('center');
  range.setFontFamily("Montserrat");
  sheet.getRange(3,diffCol,totalMembers+1,diffCount).setHorizontalAlignment('left');
  if (commentInclude == true) {
    sheet.getRange(3,commentCol,totalMembers+1,1).setHorizontalAlignment('left');
  }
  range = sheet.getRange(1,1,rows,1);
  range.setHorizontalAlignment('left');
  sheet.setFrozenColumns(4);
  sheet.setFrozenRows(2);
  range = sheet.getRange(1,1,2,columns);
  range.setBackground('black');
  range.setFontColor('white');

  try {
    if (dataRestore == true && fresh == false) {
      if (tiebreakerCol  >= 0) {
        previousDataRange.setValues(previousData);
        ss.toast('Previous values restored for week ' + week + ' if they were present');
      } else {
        Logger.log('ERROR: Previous data not transferred! Undo immediately');
        ss.toast('ERROR: Previous data not transferred! Undo immediately');
      }
      if (commentCol  >= 0 && commentInclude == true) {
        try {
          previousCommentRange.setValues(previousComment);
        }
        catch (err) {
        }
      }
    }
  }
  catch (err) {
    Logger.log('ERROR: Previous data not transferred or didn\'t exist! Undo immediately');
    ss.toast('ERROR: Previous data not transferred or didn\'t exist! Undo immediately');
  }
  nflOutcomesUpdate(year,week,exportMatches);
  return sheet;
}

//------------------------------------------------------------------------
// OVERALL Sheet Creation / Adjustment
function overallSheet(year,weeks,members) {
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = 'OVERALL';

  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }

  if (weeks == null) {
    weeks = fetchWeeks();
  }
  sheet.clear();
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
  sheet.getRange(rows,1).setValue('AVERAGES');

  let mask;
  for ( let a = 0; a < weeks; a++ ) {
    sheet.getRange(1,a+3).setValue(a+1);
    sheet.setColumnWidth(a+3,30);
    if (a+1 < 10 ) { 
      mask = '0' + (a+1);
    } else {
      mask = (a+1);
    }
    sheet.getRange(2,a+3).setFormula('=iferror(arrayformula(TOT_'+year+'_'+mask+'))');
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
  ss.setNamedRange('TOT_OVERALL_'+year+'_NAMES',rangeOverallTotNames); 
  sheet.clearConditionalFormatRules(); 
  // OVERALL TOTAL GRADIENT RULE
  let rangeOverallTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('TOT_OVERALL_'+year,rangeOverallTot);
  let formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("TOT_OVERALL_'+year+'"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("TOT_OVERALL_'+year+'"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("TOT_OVERALL_'+year+'"))') // Min value of all correct picks
    .setRanges([rangeOverallTot])
    .build();
  // OVERALL SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks+2));
  ss.setNamedRange('TOT_WEEKLY_'+year,range);
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
  overallMainFormulas(sheet,totalMembers,weeks,year,'TOT',true);
  
  return sheet;  
}

//------------------------------------------------------------------------
// OVERALL RANK Sheet Creation / Adjustment
function overallRankSheet(year,weeks,members) {
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = 'OVERALL_RANK';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  if (weeks == null) {
    weeks = fetchWeeks();
  }
  sheet.clear();
  if (members == null) {
    members = memberList();
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

  let mask;
  for ( let i = 0; i < weeks; i++ ) {
    sheet.getRange(1,i+3).setValue(i+1);
    sheet.setColumnWidth(i+3,30);
    if (i+1 < 10 ) { 
      mask = '0' + (i+1);
    } else {
      mask = (i+1);
    }
    sheet.getRange(2,i+3).setFormula('=iferror(arrayformula(RANK_'+year+'_'+mask+'))');
  }
  
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
  ss.setNamedRange('TOT_OVERALL_RANK_'+year+'_NAMES',rangeOverallTotRnkNames);  
  sheet.clearConditionalFormatRules(); 
  // RANKS TOTAL GRADIENT RULE
  let rangeOverallRankTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('TOT_OVERALL_RANK_'+year,rangeOverallRankTot);
  let formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([rangeOverallRankTot])
    .build();
  // RANKS SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks+2));
  ss.setNamedRange('TOT_WEEKLY_RANK_'+year,range);
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
  overallMainFormulas(sheet,totalMembers,weeks,year,'RANK',false);
  
  return sheet;  
}

//------------------------------------------------------------------------
// OVERALL PERCENT Sheet Creation / Adjustment
function overallPctSheet(year,weeks,members) {
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = 'OVERALL_PCT';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  if (weeks == null) {
    weeks = fetchWeeks();
  }

  sheet.clear();
  
  if (members == null) {
    members = memberList();
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
  ss.setNamedRange('TOT_OVERALL_PCT_'+year+'_NAMES',rangeOverallTotPctNames);
  sheet.clearConditionalFormatRules(); 
  // OVERALL PCT TOTAL GRADIENT RULE
  let rangeOverallTotPct = sheet.getRange('R2C2:R'+(rows-1)+'C2');
  ss.setNamedRange('TOT_OVERALL_PCT_'+year,rangeOverallTotPct);
  rangeOverallTotPct = sheet.getRange('R2C2:R'+rows+'C2');
  let formatRuleOverallPctTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("TOT_OVERALL_PCT_'+year+'"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("TOT_OVERALL_PCT_'+year+'"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("TOT_OVERALL_PCT_'+year+'"))') // Min value of all correct picks  
    .setRanges([rangeOverallTotPct])
    .build();  
  // OVERALL PCT SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+(rows-1)+'C'+(weeks+2));
  ss.setNamedRange('TOT_WEEKLY_PCT_'+year,range);
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
  overallMainFormulas(sheet,totalMembers,weeks,year,'PCT',true);

  return sheet;  
}

//------------------------------------------------------------------------
// MNF Sheet Creation / Adjustment
function mnfSheet(year,weeks,members) {
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = 'MNF';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  if (weeks == null) {
    weeks = fetchWeeks();
  }
  sheet.clear();

  if (members == null) {
    members = memberList();
  }
  let totalMembers = members.length;
  
  Logger.log('Checking for Monday games, if any');
  let data = ss.getRangeByName('NFL_'+year).getValues();
  let text = '0';
  let result = text.repeat(weeks);
  let mondayGames = Array.from(result);
  for (let a = 0; a < data.length; a++) {
    if ( data[a][2] == 1 ) {
      mondayGames[(data[a][0]-1)]++;
    }
  }
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
    if (mondayGames[a] == 2) {
      range = sheet.getRange(1,a+3);
      range.setNote('Two MNF Games')
        .setFontWeight('bold')
        .setBackground('#666666');
    } else if (mondayGames[a] == 3) {
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
  ss.setNamedRange('MNF_'+year+'_NAMES',rangeMnfNames); 
  // MNF TOTAL GRADIENT RULE
  let rangeMnfTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('MNF_'+year,rangeMnfTot);
  let formatRuleMnfTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#C9FFDF", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("MNF_'+year+'"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("MNF_'+year+'"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("MNF_'+year+'"))') // Min value of all correct picks
    .setRanges([rangeMnfTot])
    .build();
  // MNF SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks+2));
  ss.setNamedRange('MNF_WEEKLY_'+year,range);
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
  overallMainFormulas(sheet,totalMembers,weeks,year,'MNF',false);

  return sheet;  
}

//------------------------------------------------------------------------
// OVERALL / OVERALL RANK / OVERALL PCT / MNF Combination formula for sum/average per player row
function overallPrimaryFormulas(sheet,totalMembers,maxCols,action,avgRow) {
  for ( let a = 1; a < totalMembers; a++ ) {
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
    sheet.getRange(sheet.getMaxRows(),2).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))');
  } 
}

//------------------------------------------------------------------------
// OVERALL / OVERALL RANK / OVERALL PCT / MNF Combination formula for each column (week)
function overallMainFormulas(sheet,totalMembers,weeks,year,str,avgRow) {
  let b;
  for (let a = 1; a <= weeks; a++ ) {
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
    for (let a = 0; a < weeks; a++){
      let rows = sheet.getMaxRows();
      sheet.getRange(rows,a+3).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))');
    }
  }
}

//------------------------------------------------------------------------
// WEEKLY WINNERS Combination formula update
function winnersFormulas(sheet,weeks,year) {
  for (let a = 1; a <= weeks; a++ ) {
    let winRange = 'WIN_' + year + '_' + a;
    let nameRange = 'NAMES_' + year + '_' + a;
    sheet.getRange(a+1,2).setFormulaR1C1('=iferror(join(", ",sort(filter('+nameRange+','+winRange+'=1),1,true)))');
  }
}

//------------------------------------------------------------------------
// REFRESH FORMULAS FOR OVERALL / OVERALL RANK / OVERALL PCT / MNF
function allFormulasUpdate(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  const members = memberList();
  const weeks = fetchWeeks();
  const year = fetchYear();
  let sheet, totalMembers, maxCols;

  if ( pickemsInclude == true ) {
    sheet = ss.getSheetByName('OVERALL');
    maxCols = sheet.getMaxColumns();
    totalMembers = members.length;
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
    overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',false);
    overallMainFormulas(sheet,totalMembers,weeks,year,'MNF',false);

    sheet = ss.getSheetByName('WINNERS');
    winnersFormulas(sheet,weeks,year);
  }
}

//------------------------------------------------------------------------
// SURVIVOR Sheet Creation / Adjustment
function survivorSheet(year,weeks,members,dataRestore) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'SURVIVOR';
  let sheet = ss.getSheetByName(sheetName);
  let fresh = false;
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
    fresh = true;
  }

  if (members == null) {
    members = memberList();
  }
  const totalMembers = members.length;

  let maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();

  let previousDataRange, previousData;
  if (dataRestore == true && fresh == false){
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
    formula = '=iferror(vlookup(indirect(\"R[0]C1\",false),indirect(\"SURVIVOR_EVAL\"),2,false))';
    sheet.getRange(2,eliminatedCol,b,1).setFormulaR1C1(formula);
  }
  for (let b = 1; b < weeks; b++ ) {
    formula = '=if(indirect(\"R1C[0]\",false)<indirect(\"SURVIVOR_START\"),,iferror(if(sum(arrayformula(if(isblank(R2C[0]:R[-1]C[0]),0,1)))>0,counta(R2C1:R[-1]C1)-countif(R2C2:R[-1]C2,\"\<=\"\&R1C[0]),)))';
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


  if (dataRestore == true && fresh == false) {
    previousDataRange.setValues(previousData);
    ss.toast('Previous values restored for SURVIVOR sheet if they were present');
  }

  return sheet;
}

//------------------------------------------------------------------------
// SURVIVOR Sheet Creation / Adjustment
function survivorEvalSheet(year,weeks,members,survivorStart) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = 'SURVIVOR_EVAL';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  if (year == null) {
    year = fetchYear();
  }
  if (members == null) {
    members = memberList();
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
  formula = '=iferror(iferror(if(match(indirect(\"SURVIVOR!R[0]C[0]\",false),indirect(\"NFL_'+year+'_OUTCOMES_\"\&indirect(\"R1C[0]\",false),false),0)>0,0,1),iferror(if(match(iferror(vlookup(indirect(\"SURVIVOR!R[0]C[0]\",false),{indirect(\"NFL_'+year+'_AWAY_\"\&indirect(\"R1C[0]\",false)),indirect(\"NFL_'+year+'_HOME_\"\&indirect(\"R1C[0]\",false))},2,false),vlookup(indirect(\"SURVIVOR!R[0]C[0]\",false),{indirect(\"NFL_'+year+'_HOME_\"\&indirect(\"R1C[0]\",false)),indirect(\"NFL_'+year+'_AWAY_\"\&indirect(\"R1C[0]\",false))},2,false)),indirect(\"NFL_'+year+'_OUTCOMES_\"\&indirect(\"R1C[0]\",false),false),0)>0,1,0),if(and(isblank(indirect(\"SURVIVOR!R[0]C[0]\",false)),indirect(\"R1C[0]\",false)<indirect(\"WEEK\")),1,if(and(isblank(indirect(\"SURVIVOR!R[0]C[0]\",false)),indirect(\"R1C[0]\",false)<indirect(\"WEEK\"),indirect(\"R1C[0]\",false)<>indirect(\"SURVIVOR_START\")),1,)))))';
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
  
  sheet.hideSheet();
  
  return sheet;
}

//------------------------------------------------------------------------
// WINNERS Sheet Creation / Adjustment
function winnersSheet(year,weeks,members) {
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = 'WINNERS';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  let checkboxRange = sheet.getRange(2,3,weeks+3,1);
  let checkboxes = checkboxRange.getValues();
  sheet.clear();
  
  if (members == null) {
    members = memberList();
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
    if (checkboxes[a][0] == true) {
      sheet.getRange(a+1,col).check();
    }
  }
  let winRange;
  let nameRange;

  for ( let b = 1; b <= weeks; b++ ) {
    winRange = 'WIN_' + year + '_' + (b);
    nameRange = 'NAMES_' + year + '_' + (b);
    sheet.getRange(b+1,2,1,1).setFormulaR1C1('=iferror(join(", ",sort(filter('+nameRange+','+winRange+'=1),1,true)))');
  }

  return sheet;

}

//------------------------------------------------------------------------
// SUMMARY Sheet Creation / Adjustment
function summarySheet(year,members,pickemsInclude,mnfInclude,survivorInclude) {
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (pickemsInclude == null) {
    pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  } 

  if (pickemsInclude == true) {
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
  
  if (members == null) {
    members = memberList();
  }

  let headers = ['PLAYER'];
  let headersWidth = [120];
  let mnfCol;
  if (pickemsInclude == true) {
    headers = headers.concat(['TOTAL CORRECT','TOTAL RANK','AVG % CORRECT','AVG % CORRECT RANK','WEEKLY WINS']);
    headersWidth = headersWidth.concat([90,90,90,90,90]);
    if (mnfInclude == true) {
      headers = headers.concat(['MNF CORRECT','MNF RANK']);
      headersWidth = headersWidth.concat([90,90]);
      mnfCol = headers.indexOf('MNF CORRECT') + 1;
    }
  }

  let survivorCol;
  if (survivorInclude == true) {
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
  if (pickemsInclude == true) {
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
    if (mnfInclude == true) {
      rangeMNFTot = sheet.getRange('R2C'+mnfCol+':R'+rows+'C'+mnfCol);
      //ss.setNamedRange('TOT_MNF_'+year,range);
      let formatRuleMNFTot = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint('#75F0A1')
        .setGradientMinpoint('#FFFFFF')
        .setRanges([rangeMNFTot])
        .build();
      formatRules.push(formatRuleMNFTot);    
      // RANK MNF GRADIENT RULE
      rangeMNFRank = sheet.getRange('R2C'+(mnfCol+1)+':R'+rows+'C'+(mnfCol+1));
      ss.setNamedRange('TOT_MNF_RANK_'+year,rangeMNFRank);
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
    ss.setNamedRange('TOT_OVERALL_RANK_'+year,rangeOverallRank);
    let formatRuleRank = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
      .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
      .setRanges([rangeOverallRank])
      .build();
    formatRules.push(formatRuleRank);
    // WEEKLY WINS GRADIENT/SINGLE COLOR RULES
    range = sheet.getRange('R2C'+weeklyWinsCol+':R'+rows+'C'+weeklyWinsCol);
    ss.setNamedRange('WEEKLY_WINS_'+year,range); 
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
  if (survivorInclude == true) {
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
  summarySheetFormulas(totalMembers,year);

  return sheet;  
}

//------------------------------------------------------------------------
// UPDATES SUMMARY SHEET FORMULAS
function summarySheetFormulas(totalMembers,year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SUMMARY');
  let headers = sheet.getRange('1:1').getValues().flat();
  let arr = ['PLAYER','TOTAL CORRECT','TOTAL RANK','MNF CORRECT','MNF RANK','AVG % CORRECT','AVG % CORRECT RANK','WEEKLY WINS','SURVIVOR (WEEK OUT)','NOTES'];
  headers.unshift('COL INDEX ADJUST');
  for (let a = 0; a < arr.length; a++) {
    for (let b = 0; b < totalMembers; b++) {
      if (headers[a] == 'TOTAL CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{TOT_OVERALL_'+year+'_NAMES,TOT_OVERALL_'+year+'},2,false))');
      } else if (headers[a] == 'TOTAL RANK' || headers[a] == 'AVG % CORRECT RANK' || headers[a] == 'MNF RANK') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(rank(R[0]C[-1],R2C[-1]:R'+ (totalMembers+1) + 'C[-1]))');
        ss.setNamedRange('TOT_OVERALL_RANK_'+year,sheet.getRange(2,headers.indexOf('TOTAL RANK'),totalMembers,1));
      } else if (headers[a] == 'MNF CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{MNF_'+year+'_NAMES,MNF_'+year+'},2,false))');
        ss.setNamedRange('TOT_MNF_RANK_'+year,sheet.getRange(2,headers.indexOf('MNF RANK'),totalMembers,1));
      } else if (headers[a] == 'AVG % CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{TOT_OVERALL_PCT_'+year+'_NAMES,TOT_OVERALL_PCT_'+year+'},2,false))');
      } else if (headers[a] == 'WEEKLY WINS') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(countif(WEEKLY_WINNERS,R[0]C1))');
        ss.setNamedRange('WEEKLY_WINS_'+year,sheet.getRange(2,headers.indexOf('WEEKLY WINS'),totalMembers,1));
      } else if (headers[a] == 'SURVIVOR (WEEK OUT)') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(arrayformula(if(isblank(vlookup(R[0]C1,{SURVIVOR_EVAL_NAMES,SURVIVOR_EVAL_ELIMINATED},2,false)),"IN","OUT ("\&vlookup(R[0]C1,{SURVIVOR_EVAL_NAMES,SURVIVOR_EVAL_ELIMINATED},2,false)\&")")))');
      }
    }
  }
  Logger.log('Updated formulas and ranges for summary sheet');
}

//------------------------------------------------------------------------
// CREATE BLANK FORM OR FETCH EXISTING - Creates a form from a template or locates an existing form
function formFetch(name,year,week,reset) {
  // Template form for creating new forms
  let id = '12fWFNFDbH5evyoSP8FdUUi6B3ZlZuGt0IWei-7IYuq0';

  let ss = SpreadsheetApp.getActiveSpreadsheet();

  if (week == null) {
    week = ss.getRangeByName('WEEK').getValue();
  }
  if (week == null) {
    week = fetchWeek();
  }
  if (year == null) {
    year = ss.getRangeByName('YEAR').getValue();
  }
  if (year == null) {
    year = fetchYear();
  }
  
  let current = ss.getRangeByName('FORM_WEEK_'+week).getValue();
  let form;

  if (current == '' || current == null || reset == true) {
    // Preliminary checks of folder for storing form files
    let folder = null;
    let folderName = year + ' ' + name + ' Forms';
    let folders = DriveApp.getFoldersByName(folderName);
    let existingForm;
    while (folders.hasNext()) {
      let current = folders.next();
      if(folderName == current.getName()) {         
        folder = current;
      }
    }
    try { 
      folder.getName();
    } catch (err) {
      Logger.log("No " + name + " folder created for " + year + ", creating one now");
      folder = DriveApp.createFolder(year + " " + name + " Forms");
    }
    
    // Preliminary check for existing form for specified week
    let formName = 'NFL Pick \’Ems - Week ' + week + ' - ' + year;

    if (name != null && name != '') {
      formName = name + ' - Week ' + week + ' - ' + year;
    }
    let files = folder.getFilesByName(formName); 
    while (files.hasNext()) {
      let current = files.next();
      if(formName == current.getName()) {         
        form = current;
      }
    }
    try {
      Logger.log('Checking for form by using name check');
      form.getName();
    } catch (err) {
      Logger.log("No form created for week " + week +", creating one now with name \"" + formName + "\"");
      form = DriveApp.getFileById(id).makeCopy(formName,folder);
      existingForm = false;
    }
    // Get Form object instead of File object
    form = FormApp.openById(form.getId());
    let formId = form.getId();
    let urlFormEdit = form.shortenFormUrl(form.getEditUrl());
    let urlFormPub = form.shortenFormUrl(form.getPublishedUrl()); 
    let range = ss.getRangeByName('FORM_WEEK_'+week);
    range.setValue(formId);
    let sheet = ss.getSheetByName('CONFIG');
    sheet.getRange(range.getRow(),range.getColumn()+1,1,1).setValue(urlFormPub);
    sheet.getRange(range.getRow(),range.getColumn()+2,1,1).setValue(urlFormEdit);
    return [form,existingForm];
  } else {
    try {
      form = FormApp.openById(current);
      return [form,true];
    }
    catch (err) {
      ss.toast('Error Opening Form - delete Form ID for week ' + week + ' and try again to create a new form');
    }
  }
}

//------------------------------------------------------------------------
// CREATE FORMS FOR CORRECT WEEK BY CHECKING RECORDED GAMES - Tool to create form and populate with matchups as needed, creates custom survivor selection drop-downs for each member
function formCreateAuto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let markedWeek = ss.getRangeByName('WEEK').getValue();
  let markedWeekForm = ss.getRangeByName('FORM_WEEK_'+markedWeek).getValue();
  let year = fetchYear();
  let week;
  ss.toast('Gathering information to suggest the week you intend to create...');
  if ((markedWeekForm == null || markedWeekForm == '') && markedWeek == 1) {
    week = 1; // If no form exists and week is noted as 1, then proceed
  } else {
    let data = ss.getRangeByName('NFL_'+year).getValues();
    let weeks = fetchWeeks();
    let outcomeCount, gameCount, matchesUnmarked = [];
    let regex = new RegExp(/[A-Z]{2,3}/);
    for (let week = 1; week <= weeks; week++) {
      gameCount = 0;
      outcomeCount = 0;
      let outcomes = ss.getRangeByName('NFL_'+year+'_OUTCOMES_'+week).getValues().flat();
      for (let a = 0; a < data.length; a++) {
      if (data[a][0] == week) {
        gameCount++;
        }
      }
      for (let a = 0; a < outcomes.length; a++) {
        try {
          if (regex.test(outcomes[a].trim())) {
            outcomeCount++;
          }
        }
        catch (err) {
          Logger.log('Issue with formCreateAuto trim function ' + err.stack);
          if (regex.test(outcomes[a])) {
            outcomeCount++;
          }
        }
      }
      matchesUnmarked.push(gameCount - outcomeCount);
    }
    week = matchesUnmarked.lastIndexOf(0) + 2; // Add 1 for index offset and add 1 for moving to the next week
  }
  ss.toast('Week ' + week + ' is the next week up of unmarked game scores, loading \"Form Create\" script.');
  formCreate(false,week,year,null);
}

//------------------------------------------------------------------------
// CREATE FORMS - Tool to create form and populate with matchups as needed, creates custom survivor selection drop-downs for each member
function formCreate(auto,week,year,name) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Estblish week and year if not provided
  if (week == null) {
    week = ss.getRangeByName('WEEK').getValue();
  }
  if (week == null || week == '') {
    week = fetchWeek();
  }
  if (year == null || year == '') {
    year = fetchYear();
  }
  if (auto == null) {
    auto = false;
  }

  // Establish variables if not passed into function
  const pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  const commentsInclude = ss.getRangeByName('COMMENTS_PRESENT').getValue();
  let survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
  let survivorStart = ss.getRangeByName('SURVIVOR_START').getValue();
  let survivorDone = ss.getRangeByName('SURVIVOR_DONE').getValue();
  
  // Begin creation of new form if either pickems or an active survivor pool is present
  if (pickemsInclude == true || (survivorInclude == true && survivorStart <= week)) {
  
    // Fetch update to the NFL data to ensure most recent schedule
    let data;  
    if ( auto != true && week != 1) {
      try { 
        data = ss.getRangeByName('NFL_' + year).getValues();
        let refreshNFLPrompt = ui.alert('Do you want to refresh the NFL schedule data?\r\n\r\n(Only necessary when NFL schedule changes occur)', ui.ButtonSet.YES_NO);
        if (refreshNFLPrompt == 'YES') {
          fetchNFL();
        }
      }
      catch (err) {
        let fetchNFLPrompt = ui.alert('It looks like NFL data hasn\'t been brought in, import now?', ui.ButtonSet.YES_NO);
        if ( fetchNFLPrompt == 'YES' ) {
          fetchNFL();
        } else {
          ui.alert('Please run again and import NFL first or click \'YES\' next time', ui.ButtonSet.OK);
        }
      }
    }
    
    // Import all NFL data to create form once confirming refreshing of data or leaving it as-is
    data = ss.getRangeByName('NFL_' + year).getValues();
    
    let members = memberList();
    let locked = membersSheetProtected();
    if (name == null) {
      name = ss.getRangeByName('NAME').getValue();
    }
    if (name == null || name == '') {
      let confirm = false;
      let nameCheck = ui.prompt('You don\'t appear to have a name for your group set, set one now if desired:',ui.ButtonSet.OK_CANCEL);
      while (name.length <= 1 && confirm == false) {
        if (nameCheck.getSelectedButton() == 'CANCEL') {
          confirm = true;
        } else {
          nameCheck = ui.prompt('Entry for group name was too short or blank, try again or cancel to use default name',ui.ButtonSet.OK_CANCEL);
          name = nameCheck.getResponseText();
          if (name != '' && name.length > 1) {
            confirm = true;
          }
        }
      }
    }
    if (name == null || name == '') {
      name = 'NFL Pick \'Ems';
      ss.getRangeByName('NAME').setValue(name);
    }

    let existingForm = ss.getRangeByName('FORM_WEEK_'+week).getValue();
    let deleteExisting = false;

    let formReset;
    let changeWeek;
    let newWeek;
    
    if (auto != true && (existingForm == null || existingForm == '')) {
      formReset = ui.alert('Initiate form for week ' + week + '?', ui.ButtonSet.YES_NO);
    } else if (auto != true && existingForm != null) {
      formReset = ui.alert('A form exists for week ' + week + ', do you want to delete the former form and create a new one?\r\n\r\n\ALERT: This will delete any previous form responses for this week.', ui.ButtonSet.YES_NO);
      if (formReset == 'YES') {
        deleteExisting = true;
      }
    } else {
      formReset = 'YES';
    }
    if ( formReset == ui.Button.NO && auto != true ) {
      changeWeek = ui.alert('Create form for another week than ' + week + '?', ui.ButtonSet.YES_NO);
      if ( changeWeek == 'YES' ) {
        newWeek = ui.prompt('Specify new week:', ui.ButtonSet.OK);
        week = newWeek.getResponseText();
        existingForm = ss.getRangeByName('FORM_WEEK_'+week).getValue();
        let fetched = formFetch(name,year,week);
        formId = fetched[0];
        if (fetched[1] == true) {
          formReset = ui.alert('A form exists for week ' + week + ', do you want to delete the former form and create a new one?\r\n\r\n\ALERT: This will delete any previous form responses for this week.', ui.ButtonSet.YES_NO);
          if (formReset == 'YES') {
            deleteExisting = true;
          }
        } else {
          formReset = 'YES';
        }
      }
    }
    if ( formReset == 'YES' ) {
      let survivorReset, survivorUnlock;
      if (survivorInclude == true && week != 1) {
        if (survivorDone == true) {
          survivorReset = ui.alert('Survivor contest has ended, would you like to restart the contest for week ' + week + '?', ui.ButtonSet.YES_NO);
          if (survivorReset == 'NO') {
            ss.getRangeByName('SURVIVOR_PRESENT').setValue(false);
          } else if (survivorReset == 'YES') {
            survivorUnlock = ui.alert('Membership can be re-opened for new additions, would you like to allow new members to join for this round?', ui.ButtonSet.YES_NO);
            if (survivorUnlock == 'YES') {
              createMenuUnlockedWithTrigger(true);
              ss.toast('Membership unlocked');
            } else {
              createMenuLockedWithTrigger(true);
              ss.toast('Membership locked');
            }
            ss.getRangeByName('SURVIVOR_START').setValue(week);
            survivorStart = week;
            survivorDone = false;
          }
        }
      }
      locked = membersSheetProtected();
      if (locked == false && survivorInclude == true && pickemsInclude == false && week > survivorStart) {
        createMenuLockedWithTrigger(true);
        ss.toast('Membership locked due to survivor already starting in week ' + survivorStart + '.');
      } else if (locked == false && week > survivorStart) {
        survivorUnlock = ui.alert('This week is past the start of the survivor pool, do you want to keep membership open to new members still?', ui.ButtonSet.YES_NO);
        if (survivorUnlock == 'NO') {
          createMenuLockedWithTrigger(true);
          ss.toast('Membership locked');
          locked = true;
        }
      }

      // Once script starts for creating form, set week to match
      ss.getRangeByName('WEEK').setValue(week);
      ss.toast('Beginning creation of form for week ' + week);


      // Attempt to clear former form if user opted to remove it
      if (deleteExisting == true) {
        let form = FormApp.openById(existingForm);
        try {
          form.deleteAllResponses();
        }
        catch (err) {
          Logger.log('Issue clearing previous responses');
        }
        try {
          let form = FormApp.openById(existingForm);
          let file = DriveApp.getFileById(form.getId());
          file.setTrashed(true);
          ss.getSheetByName('CONFIG').getRange(ss.getRangeByName('FORM_WEEK_'+week).getRow(),ss.getRangeByName('FORM_WEEK_'+week).getColumn(),1,3).setValue('');
        }
        catch (err) {
          Logger.log('Issue deleting previous form');
        }
      }

      let formFetchOutput = formFetch(name,year,week,true);
      form = formFetchOutput[0];
      let formId = form.getId();
      // urlFormEdit = form.shortenFormUrl(form.getEditUrl());
      form.deleteItem(form.getItems()[0]);
      let urlFormPub = form.shortenFormUrl(form.getPublishedUrl());
      let teams = [];
      
      // Name question
      let nameQuestion, day, time, minutes;
      // Update form title, ensure description and confirmation are set
      form.setTitle(name + ' - Week ' + week + ' - ' + year)
        .setDescription('Select who you believe will win each game.\r\n\r\nGood luck!')
        .setConfirmationMessage('Thanks for responding!')
        .setShowLinkToRespondAgain(false)
        .setAllowResponseEdits(false)
        .setAcceptingResponses(true);
      // Update the form's response destination.
      //form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());      
      // Add drop-down list of names entries
      nameQuestion = form.addListItem();
      nameQuestion.setTitle('Name')
        .setRequired(true);

      // Pick 'Ems questions
      let item;
      if(ss.getRangeByName('PICKEMS_PRESENT').getValue() == true) {
        try {
          let finalGame ='';
          let a = 0;
          for (a; a < data.length; a++ ) {
            if ( data[a][0] == week ) {
              teams.push(data[a][6]);
              teams.push(data[a][7]);
              item = form.addMultipleChoiceItem();
              if ( data[a][2] == 1 ) {
                day = 'Monday Night Football';
              } else {
                day = data[a][5];
              }
              if (data[a][4] < 10) {
                minutes = '0' + data[a][4];
              } else {
                minutes = data[a][4];
              }
              if ( data[a][3] == 12 ) {
                time = data[a][3] + ':' + minutes + ' PM'; //case for 1pm start or later (24 hour time converted to standard 12 hour format)
              } else if ( data[a][3] > 12 ) {
                time = (data[a][3] - 12) + ':' + minutes + ' PM'; //case for 1pm start or later (24 hour time converted to standard 12 hour format)
              } else {
                time = data[a][3]  + ':' + minutes + ' AM'; // early (pre-noon) game start time with two digits for minutes
              }
              item.setTitle(data[a][8] + ' ' + data[a][9] + ' at ' + data[a][10] + ' ' + data[a][11]);
              finalGame = data[a][8] + ' ' + data[a][9] + ' at ' + data[a][10] + ' ' + data[a][11]; // After loop completes, this will be the matchup used for the tiebreaker
              item.setHelpText(day + ' at ' + time)
                .setChoices([item.createChoice(data[a][6]),item.createChoice(data[a][7])])
                .showOtherOption(false)
                .setRequired(true);
            }
          }
          ss.toast('Created Form questions for all ' + (teams.length/2) + ' NFL games in week ' + week);
          teams.sort();
          
          let numberValidation = FormApp.createTextValidation()
            .setHelpText('Input must be a whole number between 0 and 100')
            .requireWholeNumber()
            .requireNumberBetween(0,120)
            .build();
          
          // Tiebreaker question
          item = form.addTextItem();
          item.setTitle('Tiebreaker')
            .setHelpText('Total Points combined between ' + finalGame)
            .setRequired(true)
            .setValidation(numberValidation);
          if(commentsInclude == true && pickemsInclude == true) {
            item = form.addTextItem();
            item.setTitle('Comments')
              .setHelpText('Passing thoughts...');
            Logger.log('Set comments question');
          }          
        }
        catch (err) {
          Logger.log('Aborted due to error with pick \'ems questions: ' + err.message + ' \r\n' + err.stack);
        }
      } else { 
        // This loops through the data for the weekly matchups and compiles all the participants from the weekend in case there is not a pick 'ems pool included.
        for (let a = 0; a < data.length; a++ ) {
          if ( data[a][0] == week ) {
            teams.push(data[a][6]);
            teams.push(data[a][7]);
          }
        }
      }
      teams.sort();

      // Creates a page for adding a new user after making picks based on entry from name dropdown on page 1
      let newUserPage, newNameQuestion;
      newUserPage = form.addPageBreakItem()
        .setTitle('New User');
      newUserPage.setGoToPage(FormApp.PageNavigationType.SUBMIT);
      Logger.log('Adding new user page');
      let nameValidation = FormApp.createTextValidation()
        .setHelpText('Enter text of 2 characters, up to 30.')
        .requireTextMatchesPattern("[A-Za-z]\\D{1,30}")
        .build();
      newNameQuestion = form.addTextItem();
      newNameQuestion.setTitle('Name')
        .setHelpText('Please enter your name as you would like it to appear in future forms and the overview spreadsheet')
        .setRequired(true)
        .setValidation(nameValidation);
      if(week == survivorStart && survivorInclude == true) {
        let survivorQuestion = form.addListItem();
        survivorQuestion.setTitle('Survivor')
          .setHelpText('Select which team you believe will win this week.')
          .setChoiceValues(teams)
          .setRequired(true);
      }

      if (pickemsInclude == true || (week == survivorStart && survivorInclude == true)) {
        // Final page for existing users who aren't in the survivor pool
        let finalPage = form.addPageBreakItem();
        finalPage.setGoToPage(FormApp.PageNavigationType.SUBMIT);
      }

      let entry;
      let nameOptions = [];
      // Survivor question
      if(survivorInclude == true && survivorStart <= week) {
        let survivorPages = [];
        let survivorPage;
        let survivorQuestions = [];
        
        try {
          let survivorWeekEliminated = ss.getRangeByName('SURVIVOR_ELIMINATED').getValues().flat();
          if (survivorWeekEliminated.indexOf('') != -1) {
            let survivorMembers = ss.getRangeByName('SURVIVOR_NAMES').getValues().flat();
            let included = [];
            if (week > survivorStart) {
              let survivorPicks;
              let range = ss.getRangeByName('SURVIVOR_PICKS');
              if (week != 1) {
                survivorPicks = range.getSheet().getRange(range.getRow(),range.getColumn()+(survivorStart-1),range.getNumRows(),range.getNumColumns()-(survivorStart-1)).getValues();
              } else {
                survivorPicks = ss.getRangeByName('SURVIVOR_PICKS').getValues(); // Gets all values picked by participants from Columns 3 through end of season
              }            
              for (let a = 0; a < survivorMembers.length; a++) {
                if (survivorWeekEliminated[a] == '') {
                  included[a] = true;
                  survivorPages[a] = form.addPageBreakItem();
                  survivorPages[a].setGoToPage(FormApp.PageNavigationType.SUBMIT);
                  survivorQuestions[a] = form.addListItem();
                  survivorQuestions[a].setTitle('Survivor')
                      .setHelpText('Select which team you believe will win this week. Teams you\'ve used in the past are not shown.')
                      .setRequired(true);
                } else {
                  included[a] = false;
                }
              }
              for (let a = 0; a < survivorMembers.length; a++) {
                let teamsCustom = [];
                if (included[a] == true) {
                  let survivorQuestion = survivorQuestions[a];
                  for (let b = 0; b < teams.length; b++) {
                    if (survivorPicks[a].indexOf(teams[b]) == -1) {
                      teamsCustom.push(teams[b]);
                    }
                  }
                  survivorQuestion.setTitle('Survivor')
                      .setHelpText('Select which team you believe will win this week. Teams you\'ve used in the past are not shown.')
                      .setChoiceValues(teamsCustom)
                      .setRequired(true);
                }
              }
            } else if (week == 1) {
                survivorPage = form.addPageBreakItem()
                  .setTitle('Survivor Start');
                let survivorQuestion = form.addListItem();
                survivorQuestion.setTitle('Survivor')
                  .setHelpText('Select which team you believe will win this week.')
                  .setChoiceValues(teams)
                  .setRequired(true);
            } else if (week == survivorStart) {
                survivorPage = form.addPageBreakItem()
                  .setTitle('Survivor Start');
                let survivorQuestion = form.addListItem();
                survivorQuestion.setTitle('Survivor')
                  .setHelpText('Survivor competition has been restarted beginning this week. Select which team you believe will win.')
                  .setChoiceValues(teams)
                  .setRequired(true);
            }     
            if (week > survivorStart) {
              for (let a = 0; a < survivorMembers.length; a++) {
                entry = null;
                if (included[a] == true) {
                  entry = nameQuestion.createChoice(survivorMembers[a],survivorPages[a]);
                } else if (pickemsInclude == true) {
                  entry = nameQuestion.createChoice(survivorMembers[a],FormApp.PageNavigationType.SUBMIT);
                }
                if (entry != null) {
                  nameOptions.push(entry);
                }
              }
            } else if (week == 1 || week == survivorStart) {
              for (let a = 0; a < survivorMembers.length; a++) {
                entry = nameQuestion.createChoice(survivorMembers[a],survivorPage);
                nameOptions.push(entry);
              }
            }
          }
        }
        catch (err) {
          Logger.log('Survivor Issue in formCreate: ' + err.stack);
        }
        ss.toast('Created survivor question(s)');
      } else if (pickemsInclude == true) {
        for (let a = 0; a < members.length; a++) {
          entry = null;
          if (commentsInclude == true && pickemsInclude == true) {
            entry = nameQuestion.createChoice(members[a],FormApp.PageNavigationType.CONTINUE);
          } else if (commentsInclude == false) {
            entry = nameQuestion.createChoice(members[a],FormApp.PageNavigationType.SUBMIT);
          }
          if (entry != null) {
            nameOptions.push(entry);
          }
        }
      }


      if (locked == false && (pickemsInclude == true || (survivorInclude == true && week == survivorStart))) {
        nameOptions.push(nameQuestion.createChoice('New User',newUserPage));
        nameQuestion.setHelpText('Select your name from the dropdown or select \'New User\' if you haven\'t joined yet.');
      } else if (pickemsInclude == false && survivorInclude == true && week != survivorStart) {
        nameQuestion.setHelpText('Select your name from the dropdown. If your name is not an option, then you were eliminated from the survivor pool.');
      } else {
        nameQuestion.setHelpText('Select your name from the dropdown.');
      }
      // Checks for nameOptions length and ensures there are valid names/navigation for pushing to the nameQuestion, though this is likely going to result in the inability to do the survivor pool correctly
      if (nameOptions.length == 0 || (nameOptions == 1 && survivorInclude == true)) {
        for (let a = 0; a < members.length; a++) {
          nameOptions.push(nameQuestion.createChoice(members[a],FormApp.PageNavigationType.SUBMIT));
        }
        ss.toast('Survivor member list error encountered. You may have the week advanced too far relative to game outcomes recorded or the survivor pool is complete.');
        Logger.log('No nameChoices provided through script geared towards survivorInclude (' + survivorInclude + '), created default list of all members to compensate, but this is likely inaccurate to what is desired');
      }
      nameQuestion.setChoices(nameOptions);

      // Update all formulas to account for new weekly sheets that may have been created
      allFormulasUpdate();

      // Final alert and prompt to open tab of form
      let pub = ui.alert('Form for week ' + week + ' shareable link:\r\n' + urlFormPub + '\r\n\r\nWould you like to open the weekly Form in a new tab?', ui.ButtonSet.YES_NO);
      if ( pub == 'YES' ) {
        openUrl(urlFormPub);
      }
    } else {
    ss.toast('Canceled form creation');
    }
  } else if (pickemsInclude == false && (survivorInclude == true && week < survivorStart)) {
    ui.alert('Your survivor pool start week is greater than this week (' + week + ') and you have no pick \'ems pool enabled. Change your start week for the survivor pool on the CONFIG sheet (normally hidden, but will be activate after you close this dialogue) or run the \"Create Form\" function again and it should prompt you to re-start survivor pool if not set', ui.ButtonSet.OK);
    ss.getSheetByName('CONFIG').activate();
  } else if (survivorInclude == false && pickemsInclude == false) {
    ui.alert('You have no pick \'ems competition included and survivor pool is either done or not enabled. Go to the CONFIG sheet (normally hidden, but will be activate after you close this dialogue) to change the presence of one or both and retry the \"Create Form\" function.', ui.ButtonSet.OK);
    ss.getSheetByName('CONFIG').activate();
  }
}

//------------------------------------------------------------------------
// REMOVE NEW USER OPTION - Removes the 'New User' option from the current week's Form
function removeNewUserQuestion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const week = fetchWeek();
  let nameQuestion, found = false;
  try {
      
    let form = FormApp.openById(ss.getRangeByName('FORM_WEEK_'+week).getValue());
    let items = form.getItems();
    for (let a = 0; a < items.length; a++) {
      if (items[a].getType() == 'LIST' && items[a].getTitle() == 'Name') {
        nameQuestion = items[a];
      }
    }

    let choices = nameQuestion.asListItem().getChoices();
    for (let a = 0; a < choices.length; a++) {
      if (choices[a].getValue() == 'New User') {
        choices.splice(a,1);
        found = true;
      }
    }
    if (found == true) {
      nameQuestion.asListItem().setChoices(choices);
      ss.toast('Removed the option of \"New User\" from the form.');
    } else {
      ss.toast('No \"New User\" option was present on the form.');
    }
  }
  catch (err) {
    ss.toast('Failed to remove the list item of \"New User\" from the form.');
  }
}

//------------------------------------------------------------------------
// OPEN URL - Quick script to open a new tab with the newly created form, in this case
function openUrl(url){
  var js = "<script>window.open('" + url + "', '_blank');google.script.host.close();</script>;";
  var html = HtmlService.createHtmlOutput(js)
    .setHeight(10)
    .setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening...');
}

//------------------------------------------------------------------------
// OPEN FORM - Function to open the Google Form quickly from the menu
function openForm(week) {
  if (week == null) {
    week = fetchWeek();
  }
  let formId = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('FORM_WEEK_'+week).getValue();
  if (formId == null || formId == ''){
    let ui = SpreadsheetApp.getUi();
    let alert = ui.alert('No Form created yet, would you like to create one now?', ui.ButtonSet.YES_NO);
    if (alert == 'YES') {
      formCreateAuto();
    } else {
      ui.alert('Try again after you\'ve created the initial Form.', ui.ButtonSet.OK);
    }
  } else {
    let form = FormApp.openById(formId);
    let urlFormPub = form.getPublishedUrl();
    openUrl(urlFormPub);
  }
}

//------------------------------------------------------------------------
// CHECK SUBMISSIONS - Tool to check who's submitted the weekly form so far
function formCheck(request,members,week) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    if (week == null) {
      week = fetchWeek();
    }
    let formId = ss.getRangeByName('FORM_WEEK_'+week).getValues()[0][0];
    let form = FormApp.openById(formId);
    let formResponses = form.getResponses(); 
    let itemResponses;
    if (members == null || members == undefined) {
      members = memberList();
    }
    let membersFlat = members.flat();
    let names = [];
    let duplicates = [];
    for (let a = 0; a < formResponses.length; a++ ) {
      let name = formResponses[a].getItemResponses()[0].getResponse();
      if (name == 'New User') {
        itemResponses = formResponses[a].getItemResponses();
        for (let b = 1; b < itemResponses.length; b++) {
          let itemResponse = itemResponses[b];
          if(form.getItemById(itemResponse.getItem().getId()).getTitle() == 'Name'){
            name = itemResponse.getResponse();
          }
        }
      }
      if (names.indexOf(name) >= 0) {
        duplicates.push(name);
      } else {
        names.push(name);
        if (membersFlat.indexOf(name) >= 0) {
          membersFlat.splice(membersFlat.indexOf(name),1);
        }        
      }
    }
    if (request == null || request == undefined || request == "missing") {
      // Logger.log(membersFlat);
      return membersFlat;
    } else if (request == "received") {
      // Logger.log(names);
      return names;
    } else if (request == "new") {
      for (var b = 0; b < members.length; b++) {
        for (var c = 0; c < names.length; c++) {
          if (members[b] == names[c]){
            names.splice(c,1);
          }          
        }
      }
      // Logger.log(names);
      return names;
    } else if (request == "duplicates") {
      // Logger.log(duplicates);
      return duplicates;
    } else if (request == "all") {
      let received = [];
      for (let b = 0; b < members.length; b++) {
        for (let c = 0; c < names.length; c++) {
          if (received.indexOf(names[c]) == -1) {
            received.push(names[c]);
          }
          if (members[b] == names[c]){
            names.splice(c,1);
          }
        }
      }
      return [received,membersFlat,names,duplicates];
    }
  }

  catch (err) {
      Logger.log('formCheck: ' + err.message + ' \r\n' + err.stack);
      let ui = SpreadsheetApp.getUi();
      ui.alert('No Form created yet, run \"Create Form\" from the \"Pick\’Ems\" menu', ui.ButtonSet.OK);
  } 
}

//------------------------------------------------------------------------
// ALERT FOR SUBMISSION CHECK
function formCheckAlert() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const configSheet = ss.getSheetByName('CONFIG');
  const pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  const survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();  
  let week = fetchWeek();
  let retry = false;
  try {
    let formId = ss.getRangeByName('FORM_WEEK_'+week).getValues()[0][0];
    if ( configSheet == null || formId == null ) {
      ui.alert('No Form created yet, run \"Create Form\" from the \"Pick\’Ems\" menu', ui.ButtonSet.OK);
    } else {
      let members = memberList();
      let totalMembers = members.length;
      let missing = formCheck("missing",members,week);
      let membersNew = formCheck("new",members,week);
      Logger.log('Current Total: ' + totalMembers);
      if (membersNew.length > 0) {
        Logger.log('New Members: ' + membersNew);
      }
      Logger.log('Missing: ' + missing);
      
      if (survivorInclude == true && pickemsInclude == false) {
        let survivorMembers = ss.getRangeByName('SURVIVOR_EVAL_NAMES').getValues().flat();
        let survivorMembersEliminated = ss.getRangeByName('SURVIVOR_EVAL_ELIMINATED').getValues().flat();
        for (let a = 0; a < survivorMembers.length; a++) {
          if (survivorMembersEliminated[a] > 0) {
            try {
              missing.splice(missing.indexOf(survivorMembers[a]),1);
            }
            catch (err) {
              Logger.log('Could not find/remove entry from user ' + survivorMembers[a] + ' from missing array');
            }
          }
        }
      }

      let submittedText = submittedTextOutput(week,members,missing);
      if (membersNew.length == 0) {
        let respondents;
        if (missing.length == 0) {
          respondents = ui.alert(submittedText, ui.ButtonSet.YES_NO);

        } else if (missing.length >= totalMembers || missing.length >= 1) {
          respondents = ui.alert(submittedText, ui.ButtonSet.OK);
          if (respondents == 'OK') {
            respondents = ui.alert('Despite missing ' + missing.length + ', would you still like to transfer the responses?', ui.ButtonSet.YES_NO);
          }
        }
        if ( respondents == 'YES' ) {
          dataTransfer(1);
        } else if ( respondents != 'NO') {
          ui.alert('Re-run \'Form Check\' function again to check submissions or import picks.', ui.ButtonSet.OK);
        }
      } else {
        let prompt;
        if (membersNew.length == 1) {
          prompt = ui.alert(membersNew[0] + ' filled out a form as a new member, would you like to update membership including this individual?',ui.ButtonSet.YES_NO);
        } else if (membersNew.length == 2) {
          prompt = ui.alert(membersNew[0] + ' and ' + membersNew[1] + ' both filled out forms as new members, would you like to update membership to include these inviduals?', ui.ButtonSet.YES_NO);
        } else {
          let listed = membersNew[0] + ', ' + membersNew[1];
          for (let a = 2; a < membersNew.length; a++) {
            if (a == membersNew.length - 1) {
              listed = listed + ', and ' + membersNew[a];
            } else {
              listed = listed + ', ' + membersNew[a];
            }
          }
          prompt = ui.alert(listed + ' filled out forms as new members, would you like to update membership to include these individuals?', ui.ButtonSet.YES_NO);
        }
        let skip = false;
        if (prompt == 'YES') {
          retry = true;
          skip = true;
        } else {
          // DELETES Responses from newly submitted names if undesired.
          let formId = ss.getRangeByName('FORM_WEEK_'+week).getValues()[0][0];  
          let form = FormApp.openById(formId);
          let formResponses = form.getResponses();
          let deleteIdArr = [];
          let deleteNameArr = [];
          for (let b = 0; b < membersNew.length; b++) {
            for (let c = 0; c < formResponses.length; c++){
              let itemResponses = formResponses[c].getItemResponses();
              let itemResponse = itemResponses[0].getResponse();
              if (itemResponse == membersNew[b]) {
                let scrub = ui.alert('Do you want to remove the form entry for ' + itemResponse.getResponse() + '?\r\n\r\nThis will delete this individual\'s form response and picks entirely', ui.ButtonSet.YES_NO);
                if (scrub == 'YES') {
                  deleteIdArr.push(formResponses[c].getId());
                  deleteNameArr.push(membersNew[b]);
                }
              } else {
                for (let d = 0; d < itemResponses.length; d++) {
                  if (itemResponses[d].getResponse() == membersNew[b]) {
                    itemResponse = itemResponses[d].getResponse();
                    let scrub = ui.alert('Do you want to remove the form entry for ' + itemResponse + '?\r\n\r\nThis will delete this individual\'s form response and picks entirely', ui.ButtonSet.YES_NO);
                    if (scrub == 'YES') {
                      deleteIdArr.push(formResponses[c].getId());
                      deleteNameArr.push(membersNew[b]);
                    }
                  }
                }  
              }
            }
          }
          // Deletes unwanted additions, then indicates which are being retained
          if (deleteIdArr.length > 0) {
            if (deleteIdArr.length > 1) {
              Logger.log('Deleting these submitted responses: ' + deleteNameArr);
            } else {
              Logger.log('Deleting this submitted response: ' + deleteNameArr[0]);
            }
            for (let a = 0; a < deleteIdArr.length; a++) {
              form.deleteResponse(deleteIdArr[a]);
              membersNew.splice(membersNew.indexOf(deleteNameArr[a]),1);
            }
            retry = true;
          } else {
            Logger.log('Retained all new submissions (' + membersNew + ')');
          }
        }
        let continueAdd = 'CANCEL'; // Temporary gating variable for adding members
        if (membersNew.length > 0 && prompt != 'NO' && skip == false) { // Prompt for confirmation if previously responded with a "NO"
          let text = 'Proceed with adding ';
          if (membersNew.length == 1) {
            text = text + ' ' + membersNew[0] + ' as a new member?';
          } else {
            text = text + ':\r\n';
            for (let a = 0; a < membersNew.length; a++) {
              text = text + membersNew[a];
              if (a+1 < membersNew.length) {
                text = text + '\r\n';
              }
            }
          }
          continueAdd = ui.alert(text, ui.ButtonSet.OK_CANCEL);
        } else if (membersNew.length > 0 && prompt != 'NO' && skip == true) { // Skip prompt if responded "YES" earlier
          continueAdd = 'OK';
        }
        if (continueAdd == 'OK') {
          Logger.log('New Member(s) being added: ' + membersNew);
          for (let a = 0; a < membersNew.length; a++) {
            memberAdd(membersNew[a]);
            retry = true;
          }
        } 
        if (retry == true) {
          let restart = ui.alert('Would you like to run the check again for form submissions now that membership is confirmed/updated?', ui.ButtonSet.YES_NO);
          if (restart == 'YES') {
            formCheckAlert();
          } else {
            members = memberList();
            totalMembers = members.length;
            missing = formCheck("missing",members,week);
            Logger.log('Total: ' + totalMembers);
            Logger.log('Missing: ' + missing);
            submittedTextOutput(week,members,missing);
            ui.alert(submittedText + '\r\n\r\nRe-run \'Form Check\' function again to check submissions or import picks.', ui.ButtonSet.OK);
          }
        } else {
          ui.alert(submittedText + '\r\n\r\nRe-run \'Form Check\' function again to check submissions or import picks.', ui.ButtonSet.OK);
        }
      }
    }
  }
  catch (err) {
    Logger.log('formCheckAlert: ' + err.message + ' \r\n' + err.stack);
  }
  function submittedTextOutput(week,members,missing){
    let submittedText = '';
    let totalMembers = members.length;
    let text = '';
    for (let a = 0; a < missing.length; a++) {
      if (a < missing.length - 1) {
        text = text.concat(missing[a] + '\r\n');
      } else {
        text = text.concat(missing[a]);
      }
    }
    if (missing.length >= totalMembers) {
      submittedText = 'No responses recorded yet for this week.';
    } else if (missing.length == 0) {
      submittedText = 'All responses logged for week ' + week + ', import data now?';
    } else if (missing.length == 1) {
      submittedText = text + ' is the only one who hasn\'t responded.';
    } else if (missing.length == 2) {
      submittedText = missing[0] + ' and ' + missing[1] + ' are the only two who haven\'t responded.';
    } else if (missing.length == 3) {
      submittedText = missing[0] + ', ' + missing[1] + ', and ' + missing[2] + ' are the only three who haven\'t responded.';
    } else if (missing.length >= 4) {
      submittedText = 'These ' + missing.length + ' players haven\'t responded for week ' + week + ': \r\n' + text;
    }
    return submittedText;
  }
}

//------------------------------------------------------------------------
// DATA IMPORTING - Function to import responses from the surveys
function dataTransfer(redirect,thursOnly) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi();
  const year = fetchYear();
  let week = fetchWeek();
  const membersArr = memberList();
  const members = membersArr.flat();
  const pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  const survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
  const commentInclude = ss.getRangeByName('COMMENTS_PRESENT').getValue();

  if (survivorInclude == true && pickemsInclude == false) {
    let survivorMembers = ss.getRangeByName('SURVIVOR_EVAL_NAMES').getValues().flat();
    let survivorMembersEliminated = ss.getRangeByName('SURVIVOR_EVAL_ELIMINATED').getValues().flat();
    for (let a = 0; a < survivorMembers.length; a++) {
      if (survivorMembersEliminated[a] > 0) {
        try {
          missing.splice(missing.indexOf(survivorMembers[a]),1);
        }
        catch (err) {
          Logger.log('Could not find/remove entry from user ' + survivorMembers[a] + ' from missing array');
        }
      }
    }
  }
  let continueImport = false;
  if (redirect == null) {
    let weekPrompt = ui.alert('Import picks from week ' + week + '?\r\n\r\nSelectiong \'NO\' will allow you to select a different week', ui.ButtonSet.YES_NO_CANCEL);
    if (weekPrompt == 'NO') {
      weekPrompt = ui.prompt('Type the number of the week you\'d like to import:',ui.ButtonSet.OK_CANCEL);
      let regex = new RegExp('[0-9]{1,2}');
      if (regex.test(weekPrompt.getResponseText())) {
        continueImport = true;
        week = weekPrompt.getResponseText();
      }
    } else if (weekPrompt == 'YES') {
      continueImport = true;
    }
  } else {
    continueImport = true;
  }
  let sheet, sheetName, thursRange, thursValues, thursMarked = false;
  if (continueImport) {
    ss.toast('Checking for received responses, missing members, duplicates, and new members');
    let formCheckAll = formCheck("all",members,week);
    
    let received = formCheckAll[0];
    let missing = formCheckAll[1];
    let membersNew = formCheckAll[2];
    let duplicates = formCheckAll[3];
    
    if (pickemsInclude == true) {
      if (week < 10) {
        sheetName = (year + '_0' + week);
      } else {
        sheetName = (year + '_' + week);
      }
      // Pulls the 'sheet' based on week by name
      sheet = ss.getSheetByName(sheetName);
      if (sheet == null) {
        weeklySheet(year,week,membersArr,false);
        sheet = ss.getSheetByName(sheetName);
        Logger.log('New weekly sheet created for week ' + week + ', \"' + sheetName + "\.");
      } else {
        try {
          // Checking for any populated Thursday games to retain those picks and overwrite after importing the rest of the picks.
          thursRange = ss.getRangeByName('NFL_'+year+'_THURS_PICKS_'+week);
          thursValues = thursRange.getValues();
          for (let row = 0; row < thursValues.length; row++) {
            for (let col = 0; col < thursValues[row].length; col++) {
              if (thursValues[row][col] !== '') {
                thursMarked = true; // Exit the loop if a non-blank cell is found
              }
            }
          }
        }
        catch (err) {
          ss.toast('No Thursday games range found');
        }
      }
      let thursOverwrite;
      if (thursMarked == true) {
        thursOverwrite = ui.alert('There are responses recorded for the Thursday games this week, do you want to allow the new submissions to be included for Thursday night football?\r\n\r\nNOTE: Selecting \'NO\' will mean new members\' picks will only be recorded for non-Thursday games.', ui.ButtonSet.YES_NO);
        if (thursOverwrite == 'YES') {
          thursMarked = false;
        }
      }
    }
    
    let textMissing = '';
    let textReceived = '';
    let transfer = false;
    if (thursOnly != true) {
      thursOnly = false;
    }
    if (membersNew.length > 0) {
      let membersNewList = membersNew[0];
      for (let a = 1; a < membersNew.length; a++) {
        membersNewList = membersNewList + '\r\n' + membersNew[a];
      }
      let prompt = ui.alert('You have the following new submissions that are not added to the membership:\r\n\r\n' + membersNewList + '\r\n\r\nWould you like to update membership now?', ui.ButtonSet.YES_NO_CANCEL);
      if (prompt == 'YES') {
        formCheckAlert();
      } if (prompt == 'NO') {
        transfer = true;
      }
    } else {
      transfer = true;
    }
    if (transfer == true) {
      for (let a = 0; a < missing.length; a++) {
        if (a < missing.length - 1) {
          textMissing = textMissing.concat(missing[a] + '\r\n');
        } else {
          textMissing = textMissing.concat(missing[a]);
        }
      }   
      for (let a = 0; a < received.length; a++) {
        if (a < received.length - 1) {
          textReceived = textReceived.concat(received[a] + '\r\n');
        } else {
          textReceived = textReceived.concat(received[a]);
        }
      }
      // Creates an object with format of {"NAME":"COUNT EXTRA"} of all duplicates
      let duplicateCounts = {};
      let remaining = members.length - (members.length - missing.length);
      for (let a = 0; a < duplicates.length; a++) {
        if(!duplicateCounts[duplicates[a]])
          duplicateCounts[duplicates[a]] = 0;
          ++duplicateCounts[duplicates[a]];
      }
      // Creates a string output for use in the prompts of "NAME (COUNT EXTRA)" separated by commas in the event of more than one duplicated name
      let duplicateText = '';
      let duplicatedArr = Object.entries(duplicateCounts);
      duplicatedArr.sort((a, b) => b[1] - a[1]);
      for (let a = 0; a < duplicatedArr.length; a++) {
        if (a > 0) {
          duplicateText = duplicateText + ', ';
        }
        duplicateText = duplicateText + (duplicatedArr[a][0] + ' (' + duplicatedArr[a][1] + ')');
      }

      let prompt;
      if ( redirect == null ) {
        
        if (received == 0) {
          prompt = ui.alert('No responses received yet', ui.ButtonSet.OK);
        } else if (missing == 0) {
          if (membersNew == 0 && duplicates.length == 0) {
            prompt = ui.alert('All member responses logged for week ' + week + '.\r\n\r\nImport picks now?', ui.ButtonSet.YES_NO);
          } else if (membersNew == 0 && duplicates.length > 0) {
            prompt = ui.alert('All member responses logged for week ' + week + '.\r\n\r\nThese members had duplicated responses (newest response will be imported):\r\n\r\n' + duplicateText + '\r\n\r\nImport picks now?', ui.ButtonSet.YES_NO);          
          } else if (membersNew > 0) {
            prompt = ui.alert('Received responses from the following: \r\n\r\n' + textReceived + '\r\n\r\nWith these duplicates (newest response will be imported):\r\n\r\n' + duplicateText + '\r\n\r\nImport picks now?', ui.ButtonSet.YES_NO);
          }
        } else {
          if (missing.length == 1) {
            ui.alert(textMissing + ' is the only one who hasn\'t responded.', ui.ButtonSet.OK);
          } else if (missing.length == 2) {
            ui.alert(missing[0] + ' and ' + missing[1] + ' are the only two who haven\'t responded.', ui.ButtonSet.OK);
          } else if (missing.length == 3) {
            ui.alert(missing[0] + ', ' + missing[1] + ', and ' + missing[2] + ' are the only three who haven\'t responded.', ui.ButtonSet.OK);
          } else if (missing.length >= 4) {
            ui.alert('These ' + missing.length + ' players haven\'t responded for week ' + week + ': \r\n' + textMissing, ui.ButtonSet.OK);
          }
          let promptText = 'Would you like to still import ';
          if (thursOnly == true) {
            promptText = promptText + ' Thursday picks despite missing ' + remaining + '?';
          } else {
            promptText = promptText + ' all picks despite missing ' + remaining + '?';
          }
          prompt = ui.alert(promptText, ui.ButtonSet.YES_NO);
        }
      } else {
        prompt = 'YES';
      }
      let responses = [];
      if (prompt == 'YES') {
        ss.toast('Fetching form responses now, this may take some time depending on the number of respondents.');
        let title, response;

        let formId = ss.getRangeByName('FORM_WEEK_'+week).getValues()[0][0];
        let form = FormApp.openById(formId);
        let formResponses = form.getResponses();
        //Determine Thursday games if pickems included
        let thursTeams = [];
        if (pickemsInclude == true && thursOnly == true) {
          ss.toast('Checking for what games happen on Thursday, if any');
          let data = ss.getRangeByName('NFL_'+year).getValues();
          for (let a = 0; a < data.length; a++) {
            if (data[a][0] == week && data[a][2] == -3) {
              thursTeams.push(data[a][6]);
              thursTeams.push(data[a][7]);
            }
          }
        }
        for (let b = 0; b < formResponses.length; b++) {
          let itemResponses = formResponses[b].getItemResponses();
          let itemResponse = itemResponses[0];
          responses[b] = {};
          if (pickemsInclude == true) {
            responses[b].games = [];
          }
          responses[b].timestamp = formResponses[b].getTimestamp();
          
          let user = '';
          for (let c = 0; c < itemResponses.length; c++) {
            itemResponse = itemResponses[c];
            response = itemResponse.getResponse();
            title = itemResponse.getItem().getTitle();
            if (title == 'Name' && response != 'New User') {
              responses[b].name = response;
              user = response;
            } else if (survivorInclude == true && title == 'Survivor') {
              responses[b].survivor = response;
            } else if (commentInclude == true && title == 'Comments') {
              responses[b].comment = response;
            } else if (pickemsInclude == true) {
              if ( title == 'Tiebreaker') {
                responses[b].tiebreaker = response;
              } else if (response.match(/[A-Z]{2,3}/g)) {
                if ( ( thursOnly == true && thursTeams.indexOf(response) >= 0) || thursOnly == false ) {
                  responses[b].games.push(response);
                }
              }
            }          
          }
          ss.toast('Fetched response for ' + user);
            //(itemResponse.getItem().getType() == 'MULTIPLE_CHOICE' ? (' and the item\'s choices are ' + form.getItemById(itemResponse.getItem().getId()).getChoices()) : (' and it is a text box')));
        }
        
        if (pickemsInclude == true) {
          // PICK 'EMS CONTENT
          let sheetMembers, matchups, picks, tiebreaker, mnf, comment;
          let blankMatchups = [];
          let allPicks = [];
          let tiebreakers = [];
          let comments = [];

          sheetMembers = ss.getRangeByName('NAMES_'+year+'_'+week).getValues().flat();
          matchups = ss.getRangeByName('NFL_'+year+'_'+week).getValues().flat();
          picks = ss.getRangeByName('NFL_'+year+'_PICKS_'+week);
          if (thursOnly == true) {
            for (let a = 0; a < thursTeams.length/2; a++) {
              blankMatchups.push(null);
            }
          } else {
            for (let a = 0; a < matchups.length; a++) {
              blankMatchups.push(null);
            }
          }
          tiebreaker = ss.getRangeByName('NFL_'+year+'_TIEBREAKER_'+week);
          try {
            mnf = ss.getRangeByName('NFL_'+year+'_MNF_'+week);
          }
          catch (err) {
            Logger.log('NO MNF INCLUDED');
          }
          try { 
            comment = ss.getRangeByName('COMMENTS_'+year+'_'+week);
          }
          catch (err) {
            Logger.log('NO COMMENTS PRESENT');
          }

          // Create arrays for placing pick 'ems choices
          for (let a = 0; a < sheetMembers.length; a++) {
            let single = {};
            for (let b = 0; b < responses.length; b++) {
              if ( responses[b].name == sheetMembers[a]) {
                if (Object.keys(single).length == 0) {
                  single = responses[b];
                  Logger.log('Received response from ' + sheetMembers[a] + ' adding it to variable');
                } else {
                  if ( single.timestamp < responses[b].timestamp) {
                    single = responses[b];
                    Logger.log('Got more than one response for ' + sheetMembers[a] + ', replacing with newer timestamp entry');
                  }
                }
              }
            }
            if (Object.keys(single).length != 0) {
              try {
                allPicks.push(single.games);
              } 
              catch (err) {
                Logger.log('No matchups picks for ' + sheetMembers[a]);
                allPicks.push(blankMatchups);
              }
              try {
                tiebreakers.push([single.tiebreaker]);
              }
              catch (err) {
                Logger.log('No tiebreaker pick for ' + sheetMembers[a]);
                tiebreakers.push([null]);
              }
              if (commentInclude == true) {
                try {
                  comments.push([single.comment]);
                }
                catch (err) {
                  Logger.log('No comment for ' + sheetMembers[a]);
                  comments.push([null]);
                }
              }
            } else {
              Logger.log('No response received from ' + sheetMembers[a]);
              allPicks.push(blankMatchups);
              tiebreakers.push([null]);
              comments.push([null]);
            }
          }

          // Record pick 'ems choices
          try {
            // Set pick 'ems choices
            let range = picks.getSheet().getRange(picks.getRow(),picks.getColumn(),sheetMembers.length,allPicks[0].length);
            range.setValues(allPicks);
          }
          catch (err) {
            Logger.log('Error placing weekly Pick \'Ems values: ' + err.stack);
          }
          if (thursOnly == false) {
            try {
              // Set tiebreakers
              tiebreaker.setValues(tiebreakers);
            }
            catch (err) {
              Logger.log('Error placing tiebreaker values: ' + err.stack);
            }
            if (commentInclude == true) {
              try {
                // Set comments
                comment.setValues(comments);
              }
              catch (err) {
                Logger.log('Error placing comment values: ' + err.stack);
              }
            }
          }
          if (thursMarked) {
            thursRange.setValues(thursValues);
            ss.toast('Successfully recorded week ' + week + ' pick \'ems selections and retained former Thursday picks');
          } else {
            ss.toast('Successfully recorded week ' + week + ' pick \'ems selections');        
          }
        }

        // SURVIVOR CONTENT
        let survivorMembers, survivorPicks;
        let survivors = [];
        // Create array for survivor selections
        if (survivorInclude == true) {
          survivorMembers = ss.getRangeByName('SURVIVOR_NAMES').getValues().flat();
          survivorPicks = ss.getRangeByName('SURVIVOR_PICKS');
          for (let a = 0; a < survivorMembers.length; a++) {
            let single = {};
            for (let b = 0; b < responses.length; b++) {
              if ( responses[b].name == survivorMembers[a]) {
                if (Object.keys(single).length == 0) {
                  single = responses[b];
                  Logger.log('Received response from ' + survivorMembers[a] + ' adding it to variable');
                } else {
                  if ( single.timestamp < responses[b].timestamp) {
                    single = responses[b];
                    Logger.log('Got more than one response for ' + survivorMembers[a] + ', replacing with newer timestamp entry');
                  }
                }
              }
            }
            try {
              if (single.name == survivorMembers[a]) {
                survivors.push([single.survivor]);
              } else {
                survivors.push(['']);
              }
            }
            catch (err)
            {
              Logger.log('No survivor response recorded for ' + survivorMembers[a]);
              survivors.push(['']);
            }
          }
          // Set values on Survivor sheet
          try {
            let range = survivorPicks.getSheet().getRange(survivorPicks.getRow(),survivorPicks.getColumn()+(week-1),survivorMembers.length,1);
            range.setValues(survivors);
            ss.toast('Successfully recorded week ' + week + ' survivor selections');
          }
          catch (err){
            Logger.log('Error placing Survivor picks: ' + err.stack);
            ss.toast('Error placing survivor selections. Make sure you haven\'t modified the Members or Survivor sheets.\r\n\r\n' + err.message);
          }        
        }
      } else {
        ss.toast('Canceled');
      }
    } else {
      ss.toast('Canceled');
    } 
  } else {
    ss.toast('Canceled');
  }
}

//------------------------------------------------------------------------
// DATA IMPORTING - Function to import responses from the surveys for Thursday only
function dataTransferTNF() {
  dataTransfer(null,true);
}

//------------------------------------------------------------------------
// SERVICE Function to remove all triggers on project
function deleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

//------------------------------------------------------------------------
// RESET Function to reset and create menu for runFirst
function resetSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let prompt = ui.alert('Reset spreadsheet and delete all data?', ui.ButtonSet.YES_NO);
  if (prompt == 'YES') {
    
    var promptTwo = ui.alert('Are you sure? This would be very difficult to recover from.',ui.ButtonSet.YES_NO);
    if (promptTwo == 'YES') {
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
      var menu = SpreadsheetApp.getUi().createMenu('Setup');
      menu.addItem('Run First','runFirst')
      .addToUi();
    } else {
      ss.toast('Canceled reset');
    }
  } else {
    ss.toast('Canceled reset');
  }
  
}
