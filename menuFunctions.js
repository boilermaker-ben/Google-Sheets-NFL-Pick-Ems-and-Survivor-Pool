
//------------------------------------------------------------------------
// CREATE MENU - this is the standard setup once the sheet has been configured and the data is all imported
function createMenu(lock,trigger) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (lock == undefined || lock == null) {
    lock = membersSheetProtected();
  }
  let pickems = false;
  try{
    pickems = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('Issue gathering PICKEMS_PRESENT cell, you may not have completed setup correctly.');
    pickems = true;
  }
  let tnfInclude = true;
  try{
    tnfInclude = ss.getRangeByName('TNF_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('Issue gathering TNF_PRESENT cell, you may not have completed setup correctly.');
  }
  let bonus = false;
  try{
    bonus = ss.getRangeByName('BONUS_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('Issue gathering BONUS_PRESENT cell, you may not have completed setup correctly.');
  }
  let mnfDouble = false;
  try{
    mnfDouble = ss.getRangeByName('MNF_DOUBLE').getValue();
  }
  catch (err) {
    Logger.log('Issue gathering MNF_DOUBLE cell, you may not have completed setup correctly.');
  }
  let menu = ui.createMenu('Picks');
    menu.addItem('Create a Form','formCreateAuto')
      .addItem('Open Current Form','openForm');
  if (pickems) {
    menu.addItem('Week Sheet Creation','weeklySheetCreate');
  }
  menu.addSeparator();
  if (tnfInclude) {
    menu.addItem('Check Responses','formCheckAlert')
      .addItem('Import Thursday Picks','dataTransferTNF')
      .addItem('Import Picks','dataTransfer');
  } else {
    menu.addItem('Check Responses','formCheckAlert')
      .addItem('Import Picks','dataTransfer');
  }
  menu.addSeparator()
    .addItem('Check ' + league + ' Scores','recordWeeklyScores')
    .addItem('Update ' + league + ' Schedule', 'fetchSchedule');
  menu.addSeparator();
  if (!bonus) {
    menu.addItem('Enable Bonus','bonusUnhide');
  } else if (mnfDouble) {
    menu.addSubMenu(ui.createMenu('Bonus')
      .addItem('Hide Game Bonus Value Row','bonusHide')
      .addItem('MNF Double Value Disable','bonusDoubleMNFDisable')
      .addItem('Random Game of the Week','bonusRandomGameSet'));
  } else {
    menu.addSubMenu(ui.createMenu('Bonus')
      .addItem('Hide Game Bonus Value Row','bonusHide')
      .addItem('MNF Double Value Enable','bonusDoubleMNFEnable')
      .addItem('Random Game of the Week','bonusRandomGameSet'));
  }
  menu.addSeparator();
  if (!lock) {
  menu.addItem('Add Member(s)','memberAdd')
    .addItem('Remove Member','memberRemove')
    .addItem('Lock Members','createMenuLocked');
  } else {
    menu.addItem('Reopen Members','createMenuUnlocked');
  }
  menu.addSeparator();
  menu.addItem('Refresh Formulas','allFormulasUpdate')
    .addItem('Help & Support','showSupportDialog')
    .addToUi();
  if (trigger) {
    deleteOnOpenTriggers();
    let id = ss.getId();
    ScriptApp.newTrigger('createMenu')
      .forSpreadsheet(id)
      .onOpen()
      .create();
  }
}

// CREATE MENU LOCKED
function createMenuLocked() {
  createMenu(true,true);
  membersSheetLock();
  removeNewUserQuestion(); // Removes 'New User' from Form
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MEMBERS').hideSheet();
  Logger.log('Menu updated to a locked membership, MEMBERS locked');
  let ui = SpreadsheetApp.getUi();
  ui.alert('MEMBERSHIP LOCKED\r\n\r\nNew entrants will not be allowed through the Form nor through the menu unless \"Reopen Members\" script is run.\r\n\r\nRun \"Reopen Members\" to allow new additions in the Form and menu', SpreadsheetApp.getUi().ButtonSet.OK); 
}

// CREATE MENU UNLOCKED MEMBERSHIP - with Trigger Input
function createMenuUnlocked() {
  createMenu(false,true);
  membersSheetUnlock();
  memberAddForm(); // default action with no arguments is to add 'New User' to this week's form
  Logger.log('Menu updated to an open membership, MEMBERS unlocked');
  let ui = SpreadsheetApp.getUi();
  ui.alert('MEMBERSHIP UNLOCKED\r\n\r\nNew entrants will be allowed through the Form and through the \"Picks\" menu function: \"Add Member(s)\".\r\n\r\nRun \"Lock Members\" to prevent new additions in the Form and menu.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// CREATE MENU UNLOCKED MEMBERSHIP with Trigger Input on first pass (skips prompt)
function createMenuFirst(lock) {
  createMenu(lock,true);
}
