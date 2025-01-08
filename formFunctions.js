// FORM TOOLS
//------------------------------------------------------------------------
// FORM EXISTING CHECKS - allows for two different systems for checking the existence of a form (spreadsheet ID provided or storage checking)
function formExistingCheck(week,year,source,name) {
  // Checking for form by ID provided in spreadsheet "CONFIG" page
  if (source == 'ss' || source == null) {
    let range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('FORM_WEEK_'+week);
    let current = range.getValue();
    if (current != '') {
      Logger.log('Checking for form by using ID provided in spreadsheet');
      if(formNoResponses(current,week,source)) {
        range.getSheet().getRange(range.getRow(),range.getColumn(),1,3).setValue(''); // Sets the row related to the existing spreadsheet to blanks for repopulating with new form info
        return true;
      } else {
        return false;
      }
    } else {
      return true;
    }  

  // Checking Google Drive for folder for season and any matching forms by name
  } else if (source == 'drive') {

    let folder = null;
    if (name == null || name == '') {
      name = league + ' Pick \’Ems';
    }
    let folderName = year + ' ' + name + ' Forms';
    let folders = DriveApp.getFoldersByName(folderName);
    while (folders.hasNext()) {
      let current = folders.next();
      if(folderName == current.getName()) {         
        folder = current;
      }
    }
    try { 
      folder.getName();
    } 
    catch (err) {
      Logger.log('No ' + name + ' folder created for ' + year + ', creating one now');
      folder = DriveApp.createFolder(year + " " + name + ' Forms');
    }

    let formName = name + ' - Week ' + week + ' - ' + year;
    let files = folder.getFilesByName(formName); 
    let matches = [];
    while (files.hasNext()) {
      let current = files.next();
      if(formName == current.getName()) {         
        matches.push(current);
      }
    }
    if (matches.length == 0) {
      return true;
    } else if (matches.length == 1) {
      let form = matches[0];
      try {
        Logger.log('Checking for form in Google Drive');
        form.getName();
        if(formNoResponses(form.getId(),week,'drive')) {
          let file = DriveApp.getFileById(form.getId());
          file.setTrashed(true);
          return true;
        } else {
          return false;
        }
      }
      catch (err) {
        Logger.log('No form created for week ' + week + '.');
        return true;
      }
    } else {
      Logger.log('Multiple instances of form with name \'' + formName + '\' detected.');
      let responses = 0;
      let pending = [];
      let count = matches.length;
      for (let a = 0; a < count; a++) {
        try {
          let form = matches[a];
          Logger.log('Checking for responses of form ' + (a+1) + ' of ' + count + ' for week ' + week + ' in Google Drive');
          form.getName();
          if(formNoResponses(form.getId(),week,'drive',(a+1),count)) {
            let file = DriveApp.getFileById(form.getId());
            file.setTrashed(true);
          } else {
            responses++;
            pending.push(matches[a]);
          }
        }
        catch (err) {
          Logger.log('No form created for week ' + week + '.');
          return true;
        }
      }
      if (responses > 0) {
        const ui = SpreadsheetApp.getUi();
        let alert = ui.alert('There were ' + matches.length + ' forms found in your drive folder; ' + responses + ' had submitted responses, would you like to delete these now?', ui.ButtonSet.OK_CANCEL);
        if (alert == 'OK') {
          for (let a = 0; a < pending.length; a++) {
            let file = DriveApp.getFileById(pending[a].getId());
            file.setTrashed(true);
          }
          return true;
        } else {
          return false;
        }
      } else {
        return true;
      }
    }
  }
}

// FETCHES FORM - processes checks for existing forms and then gives command to create new form
function formFetch(name,year,week,reset) {
  if (reset) {
    return [newForm(week,year,name),true];
  } else {
    filename = formExistingCheck(week,year,'drive',name);
    if (filename) {
      let fileId = formExistingCheck(week,year,'ss');
      if (fileId) {
        return [newForm(week,year,name),true];
      }
    }
    return [null,false];
  }
}

// FORM RESPONSES CHECK - ensures there are no responses to the formId provided, and if so, prompts for form deletion
// 'multiple' and 'count' are not required, but provide information when checking through multiple instances of a form with the same name
function formNoResponses(formId,week,source,multiple,count) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let form = FormApp.openById(formId);
  let extraText;
  if (source == 'ss') {
    extraText = ' documented in the spreadseheet';
  } else if (source == 'drive') {
    extraText = ' in your Google Drive folder';
  }
  let responses = form.getResponses().length;
  if (responses > 0) {
    const ui = SpreadsheetApp.getUi();
    let text = 'EXISTING FORM WITH A RESPONSE\r\n\r\nThere is an existing response for the form for week ' + week;
    if (responses > 1) {
      text = 'EXISTING FORM WITH RESPONSES\r\n\r\nThere are existing responses for the form for week ' + week;
    }
    if (source == 'ss' || source == 'drive') {
      text = text.concat(extraText);
    }
    text = text.concat(', are you sure you want to remove this form and create a new one?');
    
    if (multiple > 0) {
      return false;
    } else {
      let prompt = ui.alert(text, ui.ButtonSet.YES_NO);
      if (prompt == ui.Button.YES) {
        let file = DriveApp.getFileById(form.getId());
        file.setTrashed(true);
        return true;
      } else {
        return false;
      }
    }
  } else {
    if (multiple > 0) {
      if (count == null) {
        count = 'unkown quantity';
      }
      let text = 'Existing form ' + multiple + ' of ' + count + ' for week ' + week + ' in Google Drive, but no responses logged.';
      ss.toast(text);
      Logger.log(text);
      return true;
    }
    let text = 'Existing form for week ' + week + extraText + ', but no responses logged.';
    ss.toast(text);
    Logger.log(text);
    return true;
  }
}

// NEW FORM - creates a copy of the template form for use in new form creation
function newForm(week,year,name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  // Template form for creating new forms
  let id = '12fWFNFDbH5evyoSP8FdUUi6B3ZlZuGt0IWei-7IYuq0';
  
  // Check for folder and create one if doesn't exist
  let folder = null;
  if (name == null || name == '') {
    name = league + ' Pick \’Ems';
  }
  let folderName = year + ' ' + name + ' Forms';
  let folders = DriveApp.getFoldersByName(folderName);
  while (folders.hasNext()) {
    let current = folders.next();
    if(folderName == current.getName()) {         
      folder = current;
    }
  }
  try { 
    folder.getName();
  } 
  catch (err) {
    Logger.log('No ' + name + ' folder created for ' + year + ', creating one now');
    folder = DriveApp.createFolder(year + " " + name + ' Forms');
  }
  
  // Establish name of form
  
  let formName = name + ' - Week ' + week + ' - ' + year;
  let form = DriveApp.getFileById(id).makeCopy(formName,folder);
  
  // Get Form object instead of File object
  form = FormApp.openById(form.getId());
  let formId = form.getId();

  // Sets a user-wide property associating this form's ID to the parent spreadsheet ID for allowing the form to write to the spreadsheet's script properties effectively
  let userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(formId,ssId);

  let urlFormEdit = form.shortenFormUrl(form.getEditUrl());
  let urlFormPub = form.shortenFormUrl(form.getPublishedUrl()); 
  let range = ss.getRangeByName('FORM_WEEK_'+week);
  range.setValue(formId);
  range.getSheet().getRange(range.getRow(),range.getColumn()+1,1,1).setValue(urlFormPub);
  range.getSheet().getRange(range.getRow(),range.getColumn()+2,1,1).setValue(urlFormEdit);
  return form;    
}

// CREATE FORMS FOR CORRECT WEEK BY CHECKING EXISTING WEEKLY SHEETS - Tool to launch formCreate script without inputs, runs checks for existing and suggests which week to create
function formCreateAuto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let markedWeek = ss.getRangeByName('WEEK').getValue();
  let markedWeekForm = ss.getRangeByName('FORM_WEEK_'+markedWeek).getValue();
  let year = fetchYear();
  ss.toast('Gathering information to suggest the week you intend to create...');
  if ((markedWeekForm == null || markedWeekForm == '') && markedWeek == 1) {
    formCreate(ss,false,1,year,null); // If no form exists and week is noted as 1, then proceed
  } else {
    Logger.log(nextWeek());
    formCreate(ss,false,nextWeek(),year,null);
  }
}

// CREATE FORMS - Tool to create form and populate with matchups as needed, creates custom survivor selection drop-downs for each member
function formCreate(ss,first,week,year,name) { // KEEP YEAR
  ss = fetchSpreadsheet(ss);
  const ui = SpreadsheetApp.getUi();
  
  // Establish week and year if not provided
  if (week == null) {
    week = ss.getRangeByName('WEEK').getValue();
  }
  if (week == null || week == '') {
    week = fetchWeek();
  }
  if (year == null || year == '') { // KEEP YEAR
    year = fetchYear();
  }
  if (first == null) {
    first = false;
  }

  // Establish variables if not passed into function
  const pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  const commentsInclude = ss.getRangeByName('COMMENTS_PRESENT').getValue();
  let survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
  let survivorStart = ss.getRangeByName('SURVIVOR_START').getValue();
  let survivorDone = ss.getRangeByName('SURVIVOR_DONE').getValue();
  let tiebreaker = ss.getRangeByName('TIEBREAKER_PRESENT').getValue();
  let tnfInclude = ss.getRangeByName('TNF_PRESENT').getValue();
  let bonus = ss.getRangeByName('BONUS_PRESENT').getValue();
  let mnfDouble = ss.getRangeByName('MNF_DOUBLE').getValue();

  // Begin creation of new form if either pickems or an active survivor pool is present
  if (pickemsInclude || (survivorInclude && survivorStart <= week)) {
    // Fetch update to the season data to ensure most recent schedule
    if (!first & week != 1 & week <= 18) {
      try { 
        ss.getRangeByName(league);
        let refreshSchedulePrompt = ui.alert(league + ' REFRESH\r\n\r\nDo you want to refresh the ' + league + ' schedule data?\r\n\r\n(Only necessary when ' + league + ' schedule changes occur)', ui.ButtonSet.YES_NO);
        if (refreshSchedulePrompt == 'YES') {
          fetchSchedule(year);
        }
      }
      catch (err) {
        let fetchSchedulePrompt = ui.alert(league + ' SCHEDULE IMPORT\r\n\r\nIt looks like ' + league + ' data hasn\'t been brought in, import now?', ui.ButtonSet.YES_NO);
        if ( fetchSchedulePrompt == 'YES' ) {
          fetchSchedule(year);
        } else {
          ui.alert('RETRY\r\n\r\nPlease run again and import ' + league + ' first or click \'YES\' next time', ui.ButtonSet.OK);
        }
      }
    }
    
    let members = memberList(ss);
    let locked = membersSheetProtected();
    if (name == null) {
      name = ss.getRangeByName('NAME').getValue();
    }
    if (name == null || name == '') {
      let confirm = false;
      let nameCheck = ui.prompt('You don\'t appear to have a name for your group set, set one now if desired:',ui.ButtonSet.OK_CANCEL);
      while (name.length <= 1 && !confirm) {
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
      name = league + ' Pick \'Ems';
      ss.getRangeByName('NAME').setValue(name);
    }

    let form, alert, proceed = false;
    week = nextWeek() <= 0 ? 1 : nextWeek();
    ss.toast('Week ' + week + ' is the next week up of unmarked game scores, prompting for form creation.');
    if (first) {
      alert = ui.alert('SUCCESS\r\n\r\nYou\'ve created all necessary starting assets for the spreadsheet.\r\n\r\nFIRST FORM\r\n\r\nWould you like to create a form for week ' + week + '?\r\n\r\nSelecting \'NO\' will allow you to enter a number for a different week.\r\nYou can always create a form through the \'Picks\' menu later.', ui.ButtonSet.YES_NO_CANCEL);
      if (alert == 'NO') {
        first = false;
      }
    }
    if (!first) {
      if (alert != 'NO') { // Case where first was not previously active and user did not provide any input for a different week
        alert = ui.alert('CREATE FORM\r\n\r\nInitiate form for week ' + week + '?\r\n\r\nSelecting \'NO\' will allow you to enter a number for a different week.', ui.ButtonSet.YES_NO_CANCEL);
      }
      if (alert == 'NO') {
        let regex = new RegExp(/[0-9]{1,2}/);
        let prompt = ui.prompt('Type the number of the week you\'d like to create:', ui.ButtonSet.OK_CANCEL);
        week = prompt.getResponseText();
        while (prompt.getSelectedButton() != 'CANCEL' && prompt.getSelectedButton() != 'CLOSE' && (!regex.test(week) || week < 1 || week > 23)) {
          prompt = ui.prompt('You didn\'t provide a valid week number.\r\n\r\nType the number of the week you\'d like to create:', ui.ButtonSet.OK_CANCEL);
          week = prompt.getResponseText();
        }
        if (prompt.getSelectedButton() == 'OK') {
          let fetched = formFetch(name,year,week);
          form = fetched[0];
          proceed = fetched[1];
        } else {
          ss.toast('Canceled selecting custom week');
        }
      }
    }
    if (first || alert == 'YES') {
      let fetched = formFetch(name,year,week);
      form = fetched[0];
      proceed = fetched[1];
    }
    if (proceed) {
      
      Logger.log('Creating form for week ' + week + '...');
      ss.toast('Creating form for week ' + week);

      // Import all schedule data to create form once confirming refreshing of data or leaving it as-is
      const data = ss.getRangeByName(league).getValues();
    
      let survivorReset, survivorUnlock;
      if (survivorInclude && week != 1) {
        if (survivorDone) {
          survivorReset = ui.alert('SURVIVOR COMPLETE\r\n\r\nSurvivor contest has ended, would you like to restart the contest for week ' + week + '?', ui.ButtonSet.YES_NO);
          if (survivorReset == 'NO') {
            ss.getRangeByName('SURVIVOR_PRESENT').setValue(false);
          } else if (survivorReset == 'YES') {
            survivorUnlock = ui.alert('NEW MEMBERS\r\n\r\nMembership can be re-opened for new additions, would you like to allow new members to join for this round?', ui.ButtonSet.YES_NO);
            if (survivorUnlock == 'YES') {
              createMenuUnlocked();
              ss.toast('Membership unlocked');
            } else {
              createMenuLocked();
              ss.toast('Membership locked');
            }
            ss.getRangeByName('SURVIVOR_START').setValue(week);
            survivorStart = week;
            survivorDone = false;
          }
        }
      }
      locked = membersSheetProtected();
      if (!locked && survivorInclude&& !pickemsInclude && week > survivorStart) {
        createMenuLocked();
        ss.toast('Membership locked due to survivor already starting in week ' + survivorStart + '.');
      } else if (!locked && week > survivorStart) {
        survivorUnlock = ui.alert('This week is past the start of the survivor pool, do you want to keep membership open to new members still?', ui.ButtonSet.YES_NO);
        if (survivorUnlock == 'NO') {
          createMenuLocked();
          ss.toast('Membership locked');
          locked = true;
        }
      }

      // Once script starts for creating form, set week to match
      ss.getRangeByName('WEEK').setValue(week);
      ss.toast('Beginning creation of form for week ' + week);

      let formFetchOutput = formFetch(name,year,week,true);
      form = formFetchOutput[0];
      form.deleteItem(form.getItems()[0]);
      let urlFormPub = form.shortenFormUrl(form.getPublishedUrl());
      let teams = [];
      // Name question
      let nameQuestion, item, day, time, minutes;
      // Update form title, ensure description and confirmation are set
      form.setTitle(name + ' - Week ' + week + ' - ' + year)
        .setDescription('Select who you believe will win each game.\r\n\r\nGood luck!')
        .setConfirmationMessage('Thanks for responding!')
        .setShowLinkToRespondAgain(false)
        .setAllowResponseEdits(false)
        .setAcceptingResponses(true);

      // Add drop-down list of names entries
      nameQuestion = form.addListItem();
      nameQuestion.setTitle('Name')
        .setRequired(true);

      // Pick 'Ems questions
      if(ss.getRangeByName('PICKEMS_PRESENT').getValue()) {
        try {
          let finalGame ='';
          let a = 0;
          for (a; a < data.length; a++ ) {
            if ( data[a][0] == week && (tnfInclude || (!tnfInclude && data[a][2] >= 0))) {
              teams.push(data[a][6]);
              teams.push(data[a][7]);
              evening = data[a][3] >= 17 ? true : false;
              item = form.addMultipleChoiceItem();
              if ( data[a][2] == 1 && bonus && mnfDouble && evening) {
                day = 'DOUBLE POINTS Monday Night Football';
              } else if (data[a][2] == 1 && evening) {
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
              if (tiebreaker) {
                finalGame = data[a][8] + ' ' + data[a][9] + ' at ' + data[a][10] + ' ' + data[a][11]; // After loop completes, this will be the matchup used for the tiebreaker
              }
              item.setHelpText(day + ' at ' + time)
                .setChoices([item.createChoice(data[a][6]),item.createChoice(data[a][7])])
                .showOtherOption(false)
                .setRequired(true);
            }
          }
          ss.toast('Created Form questions for all ' + (teams.length/2) + ' ' + league + ' games in week ' + week);
          teams.sort();
          
          if (tiebreaker) { // Excludes tiebreaker question if tiebreaker is disabled
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
          }

          if(commentsInclude && pickemsInclude) { // Excludes comment question if comments are disabled
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
          if ( data[a][0] == week && (tnfInclude || (!tnfInclude && data[a][2] >= 0))) {
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
        .setHelpText('Enter a minimum of 2 characters, up to 30.')
        .requireTextMatchesPattern("[A-Za-z0-9]\\D{1,30}")
        .build();
      newNameQuestion = form.addTextItem();
      newNameQuestion.setTitle('Name')
        .setHelpText('Please enter your name as you would like it to appear in future forms and the overview spreadsheet')
        .setRequired(true)
        .setValidation(nameValidation);
      if(week == survivorStart && survivorInclude) {
        let survivorQuestion = form.addListItem();
        survivorQuestion.setTitle('Survivor')
          .setHelpText('Select which team you believe will win this week.')
          .setChoiceValues(teams)
          .setRequired(true);
      }

      if (pickemsInclude || (week == survivorStart && survivorInclude)) {
        // Final page for existing users who aren't in the survivor pool
        let finalPage = form.addPageBreakItem();
        finalPage.setGoToPage(FormApp.PageNavigationType.SUBMIT);
      }

      let entry, nameOptions = [];
      // Survivor question
      if(survivorInclude && survivorStart <= week) {
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
                if (included[a]) {
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
                if (included[a]) {
                  entry = nameQuestion.createChoice(survivorMembers[a],survivorPages[a]);
                } else if (pickemsInclude) {
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
      } else if (pickemsInclude) {
        for (let a = 0; a < members.length; a++) {
          entry = nameQuestion.createChoice(members[a],FormApp.PageNavigationType.SUBMIT);
          nameOptions.push(entry);
        }
      }

      if (!locked && (pickemsInclude || (survivorInclude && week == survivorStart))) {
        nameOptions.unshift(nameQuestion.createChoice('New User',newUserPage));
        nameQuestion.setHelpText('Select your name from the dropdown or select \'New User\' if you haven\'t joined yet.');
      } else if (!pickemsInclude && survivorInclude && week != survivorStart) {
        nameQuestion.setHelpText('Select your name from the dropdown. If your name is not an option, then you were eliminated from the survivor pool.');
      } else {
        nameQuestion.setHelpText('Select your name from the dropdown.');
      }
      // Checks for nameOptions length and ensures there are valid names/navigation for pushing to the nameQuestion, though this is likely going to result in the inability to do the survivor pool correctly
      if (nameOptions.length == 0 || (nameOptions.length == 1 && survivorInclude)) {
        for (let a = 0; a < members.length; a++) {
          nameOptions.push(nameQuestion.createChoice(members[a],FormApp.PageNavigationType.SUBMIT));
        }
        ss.toast('Survivor member list error encountered. You may have the week advanced too far relative to game outcomes recorded or the survivor pool is complete.');
        Logger.log('No nameChoices provided through script geared towards survivorInclude (' + survivorInclude + '), created default list of all members to compensate, but this is likely inaccurate to what is desired');
      }
      nameQuestion.setChoices(nameOptions);

      // Update all formulas to account for new weekly sheets that may have been created
      allFormulasUpdate(ss);

      // Create trigger to fetch all values of submitted responses and log those via the fetchResponses function
      formSubmitTrigger(week);

      // Final alert and prompt to open tab of form
      let pub = ui.alert('WEEK ' + week + ' FORM CREATED\r\n\r\nShareable link:\r\n' + urlFormPub + '\r\n\r\nWould you like to open the weekly Form now?', ui.ButtonSet.YES_NO);
      if ( pub == 'YES' ) {
        openUrl(urlFormPub,week);
      }
    } else {
    ss.toast('Canceled form creation');
    }
  } else if (!pickemsInclude && (survivorInclude && week < survivorStart)) {
    ui.alert('SURVIVOR ERROR\r\n\r\nYour survivor pool start week is greater than this week (' + week + ') and you have no pick \'ems pool enabled. Change your start week for the survivor pool on the CONFIG sheet (normally hidden, but will be activate after you close this dialogue) or run the \"Create a Form\" function again and it should prompt you to re-start survivor pool if not set', ui.ButtonSet.OK);
    ss.getSheetByName('CONFIG').activate();
  } else if (!survivorInclude && !pickemsInclude) {
    ui.alert('NO GAME\r\n\r\nYou have no pick \'ems competition included and survivor pool is either done or not enabled. Go to the CONFIG sheet (normally hidden, but will be activate after you close this dialogue) to change the presence of one or both and retry the \"Create a Form\" function.', ui.ButtonSet.OK);
    ss.getSheetByName('CONFIG').activate();
  }
}

// REMOVE NEW USER OPTION - Removes the 'New User' option from the current week's Form
function removeNewUserQuestion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let week = fetchWeek();
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
    if (found) {
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

// OPEN FORM - Function to open the Google Form quickly from the menu
function openForm(week) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (week == null) {
    week = ss.getRangeByName('WEEK').getValue();
  }
  if (week == undefined) {
    week = fetchWeek();
  }  
  let formId = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('FORM_WEEK_'+week).getValue();
  if (formId == null || formId == ''){
    let ui = SpreadsheetApp.getUi();
    let alert = ui.alert('NO FORM\r\n\r\nNo Form created yet, would you like to create one now?', ui.ButtonSet.YES_NO);
    if (alert == 'YES') {
      formCreateAuto();
    } else {
      ui.alert('NO FORM\r\n\r\nTry again after you\'ve created the initial Form.', ui.ButtonSet.OK);
    }
  } else {
    let form = FormApp.openById(formId);
    let urlFormPub = form.getPublishedUrl();
    openUrl(urlFormPub,week);
  }
}

// CHECK SUBMISSIONS - Tool to check who's submitted the weekly form so far
function formCheck(request,members,week) {
  Logger.log('Form Check Running');
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    if (week == null) {
      week = fetchWeek();
    }
    if (members == null || members == undefined) {
      members = memberList(ss);
    }

    let membersFlat = members.flat();
    let names = [], duplicates = [];

    let scriptProperties = PropertiesService.getScriptProperties();
    let propertyValue = scriptProperties.getProperty('picks_' + week);

    let object = {};
    if (propertyValue) {
      object = JSON.parse(propertyValue);
    }
    
    // Form call to compare object to form responses
    let formId = ss.getRangeByName('FORM_WEEK_'+week).getValues()[0][0];
    let form = FormApp.openById(formId);
    let formResponses = form.getResponses();
    if (formResponses.length == object.total) {
      // CASE WHERE OBJECT USED
      Logger.log('Data Source: script properties');
      Object.keys(object).forEach(entry => {
        if (Object.hasOwn(object[entry],'name')) {
          names.push(entry);
          if (membersFlat.indexOf(entry) >= 0) {
            membersFlat.splice(membersFlat.indexOf(entry),1);
          }
          // Check for duplicate (previous) entries
          if (Object.hasOwn(object[entry], 'previous')) {
            for (let a = 0; a < object[entry].previous.length; a++) {
              duplicates.push(entry);
            }
          }
        }
      });
    } else {
      // CASE WHERE FORM USED
      Logger.log('Data Source: form');
      let itemResponses;
      for (let a = 0; a < formResponses.length; a++ ) {
        let name = formResponses[a].getItemResponses()[0].getResponse();
        if (name == 'New User') {
          itemResponses = formResponses[a].getItemResponses();
          for (let b = 1; b < itemResponses.length; b++) {
            let itemResponse = itemResponses[b];
            if(form.getItemById(itemResponse.getItem().getId()).getTitle() == 'Name'){
              name = itemResponse.getResponse().trim();
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
    }
    
    if (request == null || request == undefined || request == "missing") {
      return membersFlat;
    } else if (request == "received") {
      return names;
    } else if (request == "new") {
      for (let b = 0; b < members.length; b++) {
        for (let c = 0; c < names.length; c++) {
          if (members[b] == names[c]){
            names.splice(c,1);
          }          
        }
      }
      return names;
    } else if (request == "duplicates") {
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
      ui.alert('ERROR\r\n\r\nIssue fetching the form or some other error related to response checking. More details in the execution logs.', ui.ButtonSet.OK);
  } 
}

// ALERT FOR SUBMISSION CHECK
function formCheckAlert(ss,ui,week,pickemsInclude,survivorInclude) {
  ss = fetchSpreadsheet(ss);
  if (ui == undefined) {
    ui = SpreadsheetApp.getUi();
  }
  if (week == undefined) {
    week = ss.getRangeByName('WEEK').getValue();
    if (week == null) {
      week = fetchWeek();
    }
  }
  if (pickemsInclude == undefined) {
    pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  }
  if (survivorInclude == undefined) {
    survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
  }
  const configSheet = ss.getSheetByName('CONFIG');
  let retry = false;
  let survivorMembers, survivorMembersEliminated;
  try {
    let formId = ss.getRangeByName('FORM_WEEK_'+week).getValues()[0][0];
    if ( configSheet == null || formId == null ) {
      ui.alert('NO FORM\r\n\r\nNo Form created yet, run \"Create a Form\" from the \"Pick\’Ems\" menu', ui.ButtonSet.OK);
    } else {
      let members = memberList(ss);
      let totalMembers = members.length;

      let formCheckAll = formCheck("all",members,week);
      let missing = formCheckAll[1];
      let membersNew = formCheckAll[2];

      Logger.log('Current Total: ' + totalMembers);
      if (membersNew.length > 0) {
        Logger.log('New Members: ' + membersNew);
      }
      if (missing.length > 0) {
        Logger.log('Missing: ' + missing);
      } else {
        Logger.log('All added members responses recorded');
      }
      
      // Removes eliminated members from the "missing" array if they're eliminated from survivor and no pick 'ems present
      if (survivorInclude && !pickemsInclude) {
        survivorMembers = ss.getRangeByName('SURVIVOR_EVAL_NAMES').getValues().flat();
        survivorMembersEliminated = ss.getRangeByName('SURVIVOR_EVAL_ELIMINATED').getValues().flat();
        for (let a = 0; a < survivorMembers.length; a++) {
          if (survivorMembersEliminated[a] > 0 && missing.indexOf(survivorMembers[a]) >= 0) {
            missing.splice(missing.indexOf(survivorMembers[a]),1);
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
          prompt = ui.alert('NEW MEMBER\r\n\r\n' + membersNew[0] + ' filled out a form as a new member, would you like to update membership including this individual?',ui.ButtonSet.YES_NO);
        } else if (membersNew.length == 2) {
          prompt = ui.alert('NEW MEMBERS\r\n\r\n' + membersNew[0] + ' and ' + membersNew[1] + ' both filled out forms as new members, would you like to update membership to include these inviduals?', ui.ButtonSet.YES_NO);
        } else {
          let listed = membersNew[0] + ', ' + membersNew[1];
          for (let a = 2; a < membersNew.length; a++) {
            if (a == membersNew.length - 1) {
              listed = listed + ', and ' + membersNew[a];
            } else {
              listed = listed + ', ' + membersNew[a];
            }
          }
          prompt = ui.alert('NEW MEMBERS\r\n\r\n' + listed + ' filled out forms as new members, would you like to update membership to include these individuals?', ui.ButtonSet.YES_NO);
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
                let scrub = ui.alert('REMOVE ENTRY\r\n\r\nDo you want to remove the form entry for ' + itemResponse.getResponse() + '?\r\n\r\nThis will delete this individual\'s form response and picks entirely', ui.ButtonSet.YES_NO);
                if (scrub == 'YES') {
                  deleteIdArr.push(formResponses[c].getId());
                  deleteNameArr.push(membersNew[b]);
                }
              } else {
                for (let d = 0; d < itemResponses.length; d++) {
                  if (itemResponses[d].getResponse() == membersNew[b]) {
                    itemResponse = itemResponses[d].getResponse();
                    let scrub = ui.alert('REMOVE ENTRY\r\n\r\nDo you want to remove the form entry for ' + itemResponse + '?\r\n\r\nThis will delete this individual\'s form response and picks entirely', ui.ButtonSet.YES_NO);
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
        if (membersNew.length > 0 && prompt != 'NO' && !skip) { // Prompt for confirmation if previously responded with a "NO"
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
        } else if (membersNew.length > 0 && prompt != 'NO' && skip) { // Skip prompt if responded "YES" earlier
          continueAdd = 'OK';
        }
        if (continueAdd == 'OK') {
          Logger.log('New Member(s) being added: ' + membersNew);
          memberAdd(membersNew.join(","));
          retry = true;
        } 
        if (retry) {
          members = memberList(ss);
          totalMembers = members.length;
          missing = formCheck("missing",members,week);
          
          // Removes eliminated members from the "missing" array if they're eliminated from survivor and no pick 'ems present
          if (survivorInclude && !pickemsInclude) {
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

          Logger.log('Total: ' + totalMembers);
          Logger.log('Missing: ' + missing);
          submittedText = submittedTextOutput(week,members,missing);
          return true;
        } else {
          ui.alert(submittedText + '\r\n\r\nRe-run \'Check Responses\' function again to check submissions or import picks.', ui.ButtonSet.OK);
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

// DATA IMPORTING - Function to import responses from the surveys
function dataTransfer(redirect,thursOnly) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi();
  let week = ss.getRangeByName('WEEK').getValue();
  if (week == null) {
    week = fetchWeek();
  }
  const membersArr = memberList(ss);
  const members = membersArr.flat();
  const pickemsInclude = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  const survivorInclude = ss.getRangeByName('SURVIVOR_PRESENT').getValue();
  const commentInclude = ss.getRangeByName('COMMENTS_PRESENT').getValue();
  let continueImport = false;
  let object = {};
  if (redirect == null) {
    let weekPrompt = ui.alert('IMPORT PICKS\r\n\r\nBring in picks from week ' + week + '?\r\n\r\nSelecting \'NO\' will allow you to select a different week', ui.ButtonSet.YES_NO_CANCEL);
    if (weekPrompt == 'NO') {
      weekPrompt = ui.prompt('ENTER WEEK\r\n\r\nType the number of the week you\'d like to import:',ui.ButtonSet.OK_CANCEL);
      let regex = new RegExp(/[0-9]{1,2}/);
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
    
    // Removes eliminated members from the "missing" array if they're eliminated from survivor and no pick 'ems present
    if (survivorInclude && !pickemsInclude) {
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

    if (pickemsInclude) {
      sheetName = weeklySheetPrefix + week
      sheet = ss.getSheetByName(sheetName);

      if (sheet == null) {
        weeklySheet(ss,week,membersArr,false);
        sheet = ss.getSheetByName(sheetName);
        Logger.log('New weekly sheet created for week ' + week + ' named: \"' + sheetName + "\"");
      } else {
        try {
          // Checking for any populated Thursday games to retain those picks and overwrite after importing the rest of the picks.
          thursRange = ss.getRangeByName(league + '_THURS_PICKS_' + week);
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
      if (thursMarked) {
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
        let formCheckAlertOutcome = formCheckAlert(ss,ui,week,pickemsInclude,survivorInclude);
        if (formCheckAlertOutcome) {
          transfer = true;
        }
      } if (prompt == 'NO') {
        transfer = true;
      }
    } else {
      transfer = true;
    }
    if (transfer) {
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
          if (thursOnly) {
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
        //Determine Thursday games if pickems included
        let thursTeams = [];
        
        if (pickemsInclude && thursOnly) {
          ss.toast('Checking for what games happen on Thursday, if any');
          let data = ss.getRangeByName(league).getValues();
          for (let a = 0; a < data.length; a++) {
            if (data[a][0] == week && data[a][2] == -3) {
              thursTeams.push(data[a][6]);
              thursTeams.push(data[a][7]);
            }
          }
        }
        
        ss.toast('Fetching form responses now...');
        // Calls function to gather script properties where submitted responses are stored, re-gathers if needed based on total received responses
        object = formResponsesPropertyFetch(week,received.length);
        
        if (pickemsInclude) {
          // PICK 'EMS CONTENT
          let sheetMembers, matchups, picks, tiebreaker, mnf, comment;
          let blankMatchups = [];
          let allPicks = [];
          let tiebreakers = [];
          let comments = [];

          sheetMembers = ss.getRangeByName('NAMES_' + week).getValues().flat();
          matchups = ss.getRangeByName(league + '_' + week).getValues().flat();
          picks = ss.getRangeByName(league + '_PICKS_' + week);
          if (thursOnly) {
            for (let a = 0; a < thursTeams.length/2; a++) {
              blankMatchups.push(null);
            }
          } else {
            for (let a = 0; a < matchups.length; a++) {
              blankMatchups.push(null);
            }
          }
          tiebreaker = ss.getRangeByName(league + '_TIEBREAKER_' + week);
          try {
            mnf = ss.getRangeByName(league + '_MNF_' + week);
          }
          catch (err) {
            Logger.log('NO MNF INCLUDED');
          }
          try { 
            comment = ss.getRangeByName('COMMENTS_' + week);
          }
          catch (err) {
            Logger.log('NO COMMENTS PRESENT');
          }

          // Create arrays for placing pick 'ems choices
          for (let a = 0; a < sheetMembers.length; a++) {
            let single = {};
            if (Object.hasOwn(object, sheetMembers[a])) {
              single = object[sheetMembers[a]];
              try {
                if(thursOnly) {
                  allPicks.push(single.games.slice(0,thursTeams.length/2));
                } else {
                  allPicks.push(single.games);
                }
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
              if (commentInclude) {
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
            let range = picks.getSheet().getRange(picks.getRow(),picks.getColumn(),sheetMembers.length,allPicks[0].length)
            range.setValues(allPicks);
          }
          catch (err) {
            Logger.log('Error placing weekly Pick \'Ems values: ' + err.stack);
          }
          if (!thursOnly) {
            if (tiebreaker) {
              try {
                // Set tiebreakers
                tiebreaker.setValues(tiebreakers);
              }
              catch (err) {
                Logger.log('Error placing tiebreaker values: ' + err.stack);
              }
            }
            if (commentInclude) {
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
        if (survivorInclude && !thursOnly) {
          survivorMembers = ss.getRangeByName('SURVIVOR_NAMES').getValues().flat();
          survivorPicks = ss.getRangeByName('SURVIVOR_PICKS');
          for (let a = 0; a < survivorMembers.length; a++) {
            let added = false;
            if (Object.hasOwn(object, survivorMembers[a])) {
              try {
                if (object[survivorMembers[a]].survivor != null) {
                  survivors.push([object[survivorMembers[a]].survivor]);
                  added = true;
                }
              }
              catch (err)
              {
                Logger.log('Error fetching survivor response for ' + survivorMembers[a]);
              }
            }
            if (!added) {
              Logger.log('No survivor response recorded for ' + survivorMembers[a]);
              survivors.push(['']);
            }
          }
          Logger.log(survivors);
          Logger.log('survivors length: ' + survivors.length);
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

        Logger.log(object);
        Logger.log(JSON.stringify(object));

      } else {
        ss.toast('Canceled 1');
      }
    } else {
      // ss.toast('Canceled');
    } 
  } else {
    ss.toast('Canceled 3');
  }
}

function testFormResponses() {
  formResponses();
}
// FORM RESPONSE GATHERING - gathers all responses and stores them to a script property for efficient fetching
function formResponses(e) {
  
  let formId, ss, week, userProperties = PropertiesService.getUserProperties();
  
  if (e) { // Case where function called by form via trigger
    try {
      formId = e.source.getId();
      try {
        ss = SpreadsheetApp.openById(userProperties.getProperty(formId));
        week = ss.getRangeByName('WEEK').getValue();
        if (week == undefined) {
          week = fetchWeek();
        }
      }
      catch (err) {
        Logger.log('Error fetching the spreadsheet ID from user properties, perhaps the form wasn\'t initialized correctly. Importing responses will simply take longer but will call this script to run from within the context of the sheet directly, rather than from the onSubmit trigger that is in place');
        ss = null;
      }
    }
    catch (err) {
      Logger.log('Form called via a function call and was unable to get ID that way');
      try {
        ss = SpreadsheetApp.getActiveSpreadsheet();
        week = fetchWeek();
        formId = ss.getRangeByName('FORM_WEEK_'+week).getValues()[0][0];
      }
      catch (err) {
        Logger.log('Error fetching form: ' + err.stack);
      }
    }
  } else { // Case where function was called from within the script
    ss = SpreadsheetApp.getActiveSpreadsheet();
    week = ss.getRangeByName('WEEK').getValue();
    if (week == undefined) {
      week = fetchWeek();
    }
    formId = ss.getRangeByName('FORM_WEEK_'+week).getValues()[0][0];
  }
  if (ss != null) {
    
    let title, response, responses = [], object = {"week":week,"new":0,"respondents":0,"total":0};

    const form = FormApp.openById(formId);
    let formResponses = form.getResponses();

    for (let b = 0; b < formResponses.length; b++) {
      let itemResponses = formResponses[b].getItemResponses();
      let itemResponse = itemResponses[0];
      responses[b] = {};
      responses[b].games = [];
      responses[b].timestamp = formResponses[b].getTimestamp();
      
      let user = '';
      for (let c = 0; c < itemResponses.length; c++) {
        itemResponse = itemResponses[c];
        response = itemResponse.getResponse();
        title = itemResponse.getItem().getTitle();
        switch (title) {
          case 'Name':
            if (response != 'New User') {
              responses[b].name = response.trim();
              user = response.trim();
            } else {
              // Increments "new" top-level object key count and adds a boolean indicator to the user if submitting a "New User" response
              object['new']++;
              responses[b]['new'] = true;
            }
            break;
          case 'Survivor':
            responses[b].survivor = response;
            break;
          case 'Comments':
            responses[b].comment = response;
            break;
          case 'Tiebreaker':
            responses[b].tiebreaker = response;
            break;
          default:
            if (response.match(/[A-Z]{2,3}/g)) {
              responses[b].games.push(response);
            } else {
              Logger.log('Found a title to a question of ' + title + ' with a response of ' + response);
            }
            break;
        }
      }
      Logger.log('Fetched response for ' + user);
      
      if (Object.hasOwn(object, user)) {
        let previousResponses = [];
        if (object[user].timestamp < responses[b].timestamp) {
          if (Object.hasOwn(object[user], 'previous')) {
            previousResponses = object[user].previous;
            delete object[user].previous;
            previousResponses.push(object[user]);
          } else {
            previousResponses[0] = object[user];
          }
          object[user] = responses[b];
          object[user].previous = previousResponses;
        } else {
          if (Object.hasOwn(object[user], 'previous')) {
            (object[user].previous).push(responses[b]);
          } else {
            object[user].previous = responses[b];
          }
        }
      } else {
        object[user] = responses[b];
        // Increment respondents value if it is a newly added username to the object
        object.respondents++;
      }
      object.total++;
    }
    Logger.log(JSON.stringify(object));
    setProperty('picks_'+week,object);
    return object;
  }
}

// SHEET PROPERTY FORM RESPONSE FETCH - checks the respondent count and compares to stored property to more efficiently bring in responses to the spreadsheet
function formResponsesPropertyFetch(week,count) {
  if (week == null) {
    week = fetchWeek();
  }
  if (count == null) {
    count = memberList(ss).length;
  }
  let mismatch = false, object = {};
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    let propertyValue = scriptProperties.getProperty('picks_' + week);
    if (propertyValue) {
      object = JSON.parse(propertyValue);
    } else {
      object = {};
    }
    if (Object.hasOwn(object,'respondents')) {
      if (object.respondents != count) {
        Logger.log('Mismatch of submitted responses and stored response property value, re-importing responses now');
        mismatch = true;
        throw new Error();
      }
    } else {
      Logger.log('No respondent count in object from script property');
      throw new Error();
    }
  }
  catch (err) {
    if (!mismatch) {
      Logger.log('Error fetching script property object with recorded values for week ' + week + ', importing and recording to object');
    }
    object = formResponses(week);
  }
  return object;
}

// DATA IMPORTING - Function to import responses from the surveys for Thursday only
function dataTransferTNF() {
  dataTransfer(null,true);
}

// FORM SUBMIT TRIGGER - Creates a form submit trigger for current week/form that is active
function formSubmitTrigger(week) {
  if (week == null) {
    week = fetchWeek();
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let allTriggers = ScriptApp.getProjectTriggers();
  for (let a = 0; a < allTriggers.length; a++){
    if('formResponses'==allTriggers[a].getHandlerFunction()){
      ScriptApp.deleteTrigger(allTriggers[a]);
      break;
    }
  }
  let form = FormApp.openById(ss.getRangeByName('FORM_WEEK_'+week).getValue());
  ScriptApp.newTrigger('formResponses')
    .forForm(form)
    .onFormSubmit()
    .create();
  Logger.log('Created form OnSubmit trigger for week ' + week);
  ss.toast('OnSubmit trigger created for week ' + week + ' form');
}
