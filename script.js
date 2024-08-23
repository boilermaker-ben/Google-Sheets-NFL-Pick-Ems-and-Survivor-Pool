/** GOOGLE SHEETS PICK 'EMS & SURVIVOR
 * League Creator & Management Platform Tool
 * v2.6
 * 08/23/2024
 * 
 * Created by Ben Powers
 * ben.powers.creative@gmail.com
 * 
 * ------------------------------------------------------------------------
 * DESCRIPTION:
 * A series of Google Apps Scripts that generate multiple sheets and a weekly Google Form
 * to be utilized for gathering pick 'ems and survivor selections.
 * 
 * ------------------------------------------------------------------------
 * INSTRUCTIONS:
 * Once copying this code into the Apps Script (Extensions > Apps Script) console,
 * run the "runFirst" function, then move to the sheet you created to answer prompts.
 * 
 * ------------------------------------------------------------------------
 * MENU OPTIONS WITH FUNCTION EXPLANATIONS:
 * 
 * CREATE A FORM
 * This function will allow you to create a new form for the week, 
 * there are safety checks to ensure you don’t erase previous entry information 
 * and it allows you to decline creating a form for the proposed week and enter your own
 * 
 * OPEN CURRENT FORM
 * This will open the week's form for you to share with your group if you'd rather not go to the CONFIG sheet
 * 
 * WEEK SHEET CREATION
 * Utility to create the current week's sheet manually. This will usually happen automatically when importing 
 * picks if it doesn't already exist
 * 
 * ----------------------------
 * 
 * CHECK RESPONSESE
 * Checks the responses in the Google Form without revealing picks 
 * so you can hound the worthless members who haven’t submitted picks yet;
 * prompts to import if all responses are submitted and checks for new users
 * 
 * IMPORT THURSDAY PICKS
 * In case you have lagging members who you allow to submit 
 * their picks late (and not count the Thursday game for them), this allows you to only import
 * the Thursday night game matchup picks from your faithful members (not available with survivor-only)
 * 
 * IMPORT PICKS
 * Direct function to import all pick’em information submitted, 
 * it does check responses first and confirm you’d like to submit as well as checking for new members first
 * 
 * ----------------------------
 * 
 * CHECK SCORES
 * Won’t work until the first week starts; this will fetch all completed matches and the tiebreaker
 * information from the MNF game, if available
 * 
 * UPDATE SCHEDULE
 * Re-imports the NFL regular season schedule
 * 
 * ----------------------------
 * 
 * BONUS
 * This submenu will let you reveal bonus multipliers. When enabled, it will show a multiplier
 * row across the bottom of weekly sheets that can add 2x and 3x weight to games.
 * MNF can be double-weighted by default and it can also randomly pick a double-weight game
 * for the week to be designated on the Form as the "GAME OF THE WEEK"
 * 
 * ----------------------------
 * 
 * ADD MEMBERS (Hidden if membership is locked)
 * Prompts to bring in a new member or multiple (comma-separated) members.
 * This will add them to the survivor activity, if present and in the first week of competition,
 * otherwise just adds them to a pick ‘ems pool
 * 
 * REOPEN MEMBERS / LOCK MEMBERS
 * Toggles between whether you can add members or not,
 * will add “New User” option in the Form or remove it and will
 * add or remove the “Add Member” function in the menu
 * 
 * ------------------------------------------------------------------------
 * 
 * HELP & SUPPORT
 * 
 * Opens an HTML pop-up that has a link to send me an email and this project hosted on GitHub
 * 
 * If you're feeling generous and would like to support my work,
 * here's a link to support my wife, five kiddos, and me:
 * https://www.buymeacoffee.com/benpowers
 * 
 * Thanks for checking out the script!
 * 
 * **/

//------------------------------------------------------------------------
// INITIALIZATION - Initializes a menu to authorize script and begin process of creating a pick 'ems and/or survivor sheet
function onOpen() {
  let scriptProperties = PropertiesService.getScriptProperties()
  const init = scriptProperties.getProperty('initialized');
  let tz = scriptProperties.getProperty('tz');
  let ui;
  if (tz == null) {
    ui = SpreadsheetApp.getUi();
    ui.alert('WELCOME\r\n\r\nThanks for checking out this Pick \'Ems and Survivor script. \r\n\r\nBefore you get started, you\'ll need to allow the scripts to run and also check that your time zone is set correctly.', ui.ButtonSet.OK);
    timezoneCheck(ui,scriptProperties);
    ui.alert('Next, run the \'Initialize Sheet\' script from the \'Picks\' menu along the top bar.', ui.ButtonSet.OK);
    initializeMenu(init);
  } else if (!init) {
    ui = SpreadsheetApp.getUi();
    ui.alert('WELCOME\r\n\r\nThanks for checking out this Pick \'Ems and Survivor script. \r\n\r\nBefore you get started, you\'ll need to allow the scripts to run if not already authorized and initialize the sheet.\r\n\r\nRun the \'Initialize Sheet\' script from the \'Picks\' menu along the top bar. .', ui.ButtonSet.OK);
    initializeMenu(init);
  } else if (checkOnOpenTriggers()) {
    createMenu(undefined,true);
  }
}

function initializeMenu(scriptProperties,init) {
  if (scriptProperties == undefined) {
    scriptProperties = PropertiesService.getScriptProperties();
  }
  if (init == undefined) {
    init = scriptProperties.getProperty('initialized');
  }
  let simpleMenu = false;
  if (!init) {
    simpleMenu = true;
  } else if (checkOnOpenTriggers()) {
      createMenu(undefined,true);
  } else {
    simpleMenu = true;
  }
  if (simpleMenu) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Adding menu for authorization and initialization');
    SpreadsheetApp.getUi().createMenu('Picks')
      .addItem('Initialize Sheet', 'runFirst')
      .addToUi();
  }
}

function checkOnOpenTriggers() {
  // Get all triggers for the current script
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getEventType() === ScriptApp.EventType.ON_OPEN) {
      return true;
    }
  }
  return false;
}

function timezoneCheck(ui,scriptProperties,lost) {
  if (scriptProperties == undefined) {
    scriptProperties = PropertiesService.getScriptProperties();
  }
  ui = fetchUi(ui);
  const tzProp = scriptProperties.getProperty('tz');
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  year = fetchYear(); // Gets year and sets a script property if not already set

  // Confirm timezone setting before continuing
  if (tzProp == null) {
    let text = 'TIMEZONE\r\n\r\nThe timezone you\'re currently using is ' + tz + '. Is this correct?';
    if (lost) {
      text = 'TIMEZONE VALUE NOT SET\r\n\r\nDespite initialization, it seems like the timezone variable property isn\'t set. The one your account is set to is ' + tz + '. Is this correct?';
    }
    let timeZonePrompt = ui.alert(text, ui.ButtonSet.YES_NO);
    if ( timeZonePrompt != 'YES') {
      ui.alert('FIX TIMEZONE\r\n\r\nFollow these steps to change your projects time zone:\r\n\r\n1\. Go to the \'Extensions\' > \'Apps Script\' menu\r\n2\. Select the gear icon on the left menu\r\n3\. Use the drop-down to select the correct timezone\r\n4\. Close the \'Apps Script\' editor and return to the sheet\r\n5\. Restart the script through the \'Picks\' menu', ui.ButtonSet.OK);
      return false;
    } else if ( timeZonePrompt == 'YES') {
      scriptProperties.setProperty('tz',tz);
      return true;
    }
  }
}

//------------------------------------------------------------------------
// PRELIM SETUP - User input and group structure option gathering
function runFirst(onOpen) {
  const ui = SpreadsheetApp.getUi();
  
  let tzSet = timezoneCheck(ui);

  if (tzSet) {
    // Cue to go to spreadsheet for UI prompts from logger
    let start;
    if (!onOpen) {
      Logger.log('Answer the prompts that appear... [Go to spreadsheet]');
      // Prompt to start creation of spreadsheet
      start = ui.alert('SETUP\r\n\r\nThanks for checking out this Pick \'Ems and Survivor script. \r\n\r\nThere are some user inputs to gather before getting you rolling.\r\nPlease read them carefully to avoid having to restart this one-time setup.\r\n\r\nHave a great season!\r\n\r\n\- Ben', ui.ButtonSet.OK);
    }
    if ( start == ui.Button.OK || onOpen) {
      showConfigDialog();
    } else {
      ui.alert('Run the \'Initialize Sheet\' script again to begin setup', ui.ButtonSet.OK);
    }
  }
}

//------------------------------------------------------------------------
// CONFIG POPUP FOR USER INPUT - Loads HTML "configPrompt.html" file
function showConfigDialog() {
  let html = HtmlService.createHtmlOutputFromFile('configPrompt.html')
      .setWidth(500)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'CONFIGURATION');
}

//------------------------------------------------------------------------
// SUPPORT POPUP FOR HELP - Loads HTML "supportPrompt.html" file
function showSupportDialog() {
  let html = HtmlService.createHtmlOutputFromFile('supportPrompt.html')
      .setWidth(600)
      .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

//------------------------------------------------------------------------
// INPUT ASSESSMENT - this function is called by the HTML "configPrompt.html" file upon submission to continue setup
function gatherInput(values) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let popup, obj = {};
  Logger.log(values);
  try {
    // Establish variables
    let pickemsInclude = false, survivorInclude = false, lockMembers = false, tnfInclude = true, mnfInclude = true, bonus = false, tiebreaker = true, commentInclude = true;

    for (let a = 0; a < values.length; a++) {
      // Based on HTML form input, update variables:
      switch (values[a]) {
        case 'pickemsInclude':
          pickemsInclude = true;
          break;
        case 'survivorInclude':
          survivorInclude = true;
          break;
        case 'membershipLocked':
          lockMembers = true;
          break;
        case 'tnfExclude':
          tnfInclude = false;
          break;
        case 'mnfExclude':
          mnfInclude = false;
          break;
        case 'bonusInclude':
          bonus = true;
          break;
        case 'tiebreakerExclude':
          tiebreaker = false;
          break;
        case 'commentsExclude':
          commentInclude = false;
          break;
      }
    }
    obj = {
      pickemsInclude,
      survivorInclude,
      lockMembers,
      tnfInclude,
      mnfInclude,
      bonus,
      tiebreaker,
      commentInclude
    }
    let text = 'Thanks for the initial input. Here are your selections:\r\n\r\n';
    
    if (pickemsInclude) {
      text = text + 'Pick \'Ems Pool: ' + (pickemsInclude ?'YES':'NO') + '\r\n - Thursday Games: ' + (tnfInclude ?'YES':'NO') + '\r\n - MNF Pool: ' + (mnfInclude ?'YES':'NO') + '\r\n - Tiebreakers: ' + (tiebreaker ?'YES':'NO') + '\r\n - Bonus Games: ' + (bonus ?'YES':'NO') + '\r\n - Comments: ' + (commentInclude ?'YES':'NO') + '\r\n\r\nSurvivor Pool: ' + (survivorInclude ?'YES':'NO');
    } else {
      text = text + '\r\nSurvivor Pool: ' + (survivorInclude ?'YES':'NO') + '\r\n - Thursday Games: ' + (tnfInclude ?'YES':'NO') + '\r\n\r\nPick \'Ems Pool: NO'; 
    }
    text = text + '\r\n\r\nMembers: ' + (lockMembers ?'LOCKED':'UNLOCKED');
    popup = ui.alert(text, ui.ButtonSet.OK_CANCEL);
  }
  catch (err) {
    Logger.log('Error with input, please try again\r\n' + err.stack);
    ss.toast('Error with input provided, please try again');
  }
  if (popup == ui.Button.OK) {
    ss.toast('Continuing setup');
    // Calls continuation operation for initial setup
    continueSetup(obj,ss,ui);
  } else {
    ss.toast('Restart setup to get up and running');
  }
}

//------------------------------------------------------------------------
// CONTINUATION OF SETUP - After a successful submission of the HTML prompt, this script picks up for some finishing questions and then runs the setup
function continueSetup(obj,ss,ui) {
  const year = fetchYear();
  let week = fetchWeek();
  week = week <= 0 ? 1 : week;
  const weeks = fetchWeeks();
  ss = fetchSpreadsheet(ss);
  if (ui == undefined) {
    ui = SpreadsheetApp.getUi();
  }

  let cancel = false;
  const cancelText = 'Setup canceled by user. Try again later.';

  // Rename the spreadsheet if "Copy of" is present at the beginning
  let docName = ss.getName();
  if (docName.startsWith("Copy of ")) {
    docName = docName.substring(9);
    ss.rename(docName);
    Logger.log('Document renamed to \''+docName+'\'');
  }

  // Default group name
  let name = league + ' Pick \'Ems';

  let namePromptTxt = 'CUSTOMIZE NAME\r\n\r\nThe name of the forms created will be called ';
  if (obj.pickemsInclude) {
    namePromptTxt = namePromptTxt.concat('\"' + name + '\" followed by the week and year (' + year + '). Do you want to change the group name?');
  } else {
    namePromptTxt = namePromptTxt.concat('\"' + league + ' Survivor Pool\" followed by the week and year (' + year + '). Do you want to change the group name?');
  }
  // Prompts to allow the user to create a league/pool/group name [defaults to Pick 'Ems]
  let namePrompt = ui.alert(namePromptTxt, ui.ButtonSet.YES_NO);
  if ( namePrompt == ui.Button.YES) {
    // Loop to ensure name is acceptable
    let acceptance = false;
    let exit = false;
    while (!acceptance && !exit) {
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
  if (cancel) {
    ui.alert(cancelText, ui.ButtonSet.OK);
    throw new Error('Canceled during group naming question'); 
  }

  // Prompts for the inclusion of a survivor pool
  let survivorStart = week;
  if (obj.survivorInclude && week > 1) {
    ui.alert('Your survivor pool will start this week, week ' + week + ', rather than the standard starting point of week 1.', ui.ButtonSet.OK);
  }
  
  // Prompt if past week 1 to create previous week sheets
  let oldWeeks, createOldWeeks = false;
  if (obj.pickemsInclude) {
    if (week > 2) {
      oldWeeks = ui.alert('PREVIOUS WEEKS\r\n\r\nThere are ' + (week - 1) + ' previous weekly sheets that could be created to manually enter old data, would you like to have those created?', ui.ButtonSet.YES_NO);
    } else if (week == 2) {
      oldWeeks = ui.alert('PREVIOUS WEEK\r\n\r\nThere is a previous weekly sheet that could be created to manually enter old data, would you like to create a week 1 sheet as well?', ui.ButtonSet.YES_NO);
    } 
    if ( oldWeeks == ui.Button.YES) {
      createOldWeeks = true;
    }
  }

  // Prompts for a set of member names to use
  let members = memberList(ss,true);
  let finish;
  if (members.length == null || !(members.length > 0)) {
    finish = false;
  } else {
    finish = ui.alert('Alright, that\'s it for inputs! Click "OK" and hold tight as the script finishes setup.', ui.ButtonSet.OK_CANCEL);
  }

  if (finish == ui.Button.OK) {

    // Pull in schedule data and create sheet
    fetchSchedule();
    Logger.log('Fetched ' + league + ' Schedule');      
    
    // Run through all sheet information population
    try {
      // Creates Form Sheet (calls function)
      let config = configSheet(ss,name,year,week,weeks,obj.pickemsInclude,obj.mnfInclude,obj.tnfInclude,obj.tiebreaker,obj.commentInclude,obj.bonus,obj.survivorInclude,survivorStart); // KEEP YEAR
      Logger.log('Deployed Config sheet');        

      // Creates Member sheet
      memberSheet(ss,members);
      Logger.log('Deployed Members sheet');
      ss.toast('Deployed Members sheet');
      
      members = memberList(ss);
      // Creates winner selection sheet
      outcomesSheet(ss);
      Logger.log('Deployed ' + league + ' Outcomes sheet');
      if (obj.pickemsInclude) {
        // Creates Weekly Totals Record Sheet
        totSheet(ss,weeks,members);
        Logger.log('Deployed Weekly Totals sheet');
        ss.toast('Deployed Weekly Totals sheet');

        // Creates Weekly Rank Record Sheet
        rnkSheet(ss,weeks,members);
        Logger.log('Deployed Weekly Rank sheet');
        ss.toast('Deployed Weekly Rank sheet');
        
        // Creates Weekly Percentages Record Sheet
        pctSheet(ss,weeks,members);
        Logger.log('Deployed Weekly Percentages sheet');
        ss.toast('Deployed Weekly Percentages sheet');
      
        // Creates Winners Sheet
        winnersSheet(ss,year,weeks,members); // KEEP YEAR
        Logger.log('Deployed Winners sheet');
        ss.toast('Deployed Winners sheet');
        
        // Creates MNF Sheet
        if (obj.mnfInclude) {
          mnfSheet(ss,weeks,members);
          Logger.log('Deployed MNF sheet');
          ss.toast('Deployed MNF sheet');
        }
      } else {
        // Deletes sheets if no Pickem's present if they were created by accident
        try {ss.deleteSheet(ss.getSheetByName('OVERALL'));} catch (err){}
        try {ss.deleteSheet(ss.getSheetByName('RNK'));} catch (err) {}
        try {ss.deleteSheet(ss.getSheetByName('PCT'));} catch (err) {}
        try {ss.deleteSheet(ss.getSheetByName('WINNERS'));} catch (err) {}
        try {ss.deleteSheet(ss.getSheetByName('MNF'));} catch (err) {}
      }
      if (obj.survivorInclude) {
        // Creates Survivor Sheet
        let survivor = survivorSheet(ss,weeks,members,false);
        Logger.log('Deployed Survivor sheet');
        ss.toast('Deployed Survivor sheet');

        if (!obj.pickemsInclude) {
          survivor.activate();
        }

        survivorEvalSheet(ss,weeks,members,true);
        Logger.log('Deployed Survivor Eval sheet');
        ss.toast('Deployed Survivor Eval sheet');
      } else {
        try{ss.deleteSheet(ss.getSheetByName('SURVIVOR'));} catch (err) {}
      }
      
      // Creates Summary Record Sheet
      summarySheet(ss,members,obj.pickemsInclude,obj.mnfInclude,obj.survivorInclude);
      Logger.log('Deployed Summary sheet');

      if (obj.pickemsInclude) {    
        // Creates Weekly Sheets for the Current Week
        let weekly = weeklySheet(ss,week,members,false);
        Logger.log('Deployed Weekly sheet for week ' + week);
        ss.toast('Deployed Weekly sheet for week ' + week);
        weekly.activate();
      }      

      createMenuFirst(obj.lockMembers);
      Logger.log('Created menu');
      ss.toast('Created menu');

      ss.getSheetByName(league).hideSheet();

      formCreate(ss,true,week,year,name); // KEEP YEAR
      Logger.log('Created initial form for week ' + week);
      ss.toast('Created initial form for week ' + week);

      if (createOldWeeks && obj.pickemsInclude) {
        try {
          for (let a = (week - 1); a > 1; a--) {
            weeklySheet(ss,a,members,false);
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
      
      // Set script property of completion
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperties({
        'initialized': true
      });
      
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


