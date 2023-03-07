# Google Sheets / Forms NFL Pick'Ems
## Run your own pick'ems pool with this handy script and your own Google account

Google Script to run within Google Sheets that creates a Google Form to create a season-long NFL Pick'Ems league run through a Google Sheet with a weekly updating Google Form to receive responses from members.

-------------------------

**Welcome!** The project below was developed over two seasons of NFL play to create a semi-robust way of managing an NFL pick’ems league. It creates a series of sheets in a spreadsheet for tracking all picks through the 18 regular season games of an NFL season. It also includes a “survivor” pool, a Monday Night Football most correct season-long winner, a weekly most correct winner, and a season-long most correct winner. The tool will also create a simple Google Form (questionnaire) that is used to collect responses from members that can be imported to the spreadsheet easily. Lastly, another custom script, among the many below, will pull in match results and tiebreaker scores. The final Monday Night Football game score total each week is used as the tiebreaker (some weeks we do have 2 MNF games). 

It’s up to the person running the league to import the picks for the week (ideally before Thursday night) and also to update the form for the coming week (usually done Tuesday or Wednesday morning to send to the members).

I was keen to help a friend create a more robust way to track a family and friends league two seasons ago and the effort resulted in this massive block of code that’s over 3000 lines (albeit many comments and probably some serious inefficiencies). Hopefully if you have any changes you’d like to make, you can update the scripts yourself if you know Javascript. I’m not a coder by training, I’m an industrial designer. I hope it doesn’t break for you.

Disclaimer: This set of functions relies on the use of the ESPN API for pulling NFL game data. Here are the ESPN terms of use of their API. I’m sharing these scripts with you with the intent that you are taking on the responsibilities of the terms of use for your own personal use and don’t condone or endorse your use of the code here for monetization of “apps” or any other content. The terms outline the need for an “Information Form” to be submitted by a parent or guardian if you are a minor. 

It’s amazing to have the power of creating these kinds of spreadsheets and forms through Google. This content is not intended to be published nor executed outside of the use by personal users. 

Lastly, there are some safeguards Google has in place to avoid allowing users to execute any malicious code from the Google Scripts console. Please feel free to review the code below, as it contains no functions to share information, transfer information, or send emails. Information only travels between your personal Google Sheet and your Google Form that are created in the process outlined below. All sharing of content must be done by you directly (via the links that are created), such as sharing the link to the Google Form with your members and sharing a “view only” version of the spreadsheet with them to allow the members to see their league’s standings.

-------------------------

## **Table of Contents**


### **1. Example Sheets** - Screenshots of the output from a league done in 2021

- WEEKLY Sheet
  
- SUMMARY Sheet
  
- OVERALL Sheet
  
- MNF Sheet
  
- SURVIVOR Sheet
  
### **2. Example Form** - Screenshot of form from week 18 in 2021

### **3. Setup Instructions** - create new document, create script, paste code, run initial setup

### **4. Usage** - how to use the tool

### **5. Custom Functions Overview** - description of all custom functions in the “Pick’Ems” menu

-------------------------

# **1. Example Sheets**

<h3 align="center">WEEKLY Sheet</h3>
<p align="center">
<img src="https://raw.githubusercontent.com/boilermaker-ben/Google-NFL-Pick-Ems/main/images/examples/GoogleSheets-Picks-Example-01.jpg" width="600" alt="WEEKLY Sheet">
</p>



<h3 align="center">SUMMARY Sheet</h3>
<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/raw/main/images/examples/GoogleSheets-Picks-Example-02.jpg" width="600" alt="SUMMARY Sheet">
</p>



<h3 align="center">MNF Sheet</h3>
<p align="center">
<img src="./images/examples/GoogleSheets-Picks-Example-03.jpg" width="600" alt="MNF Sheet">
</p>



<h3 align="center">OVERALL Sheet</h3>
<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/examples/GoogleSheets-Picks-Example-04.jpg" width="600" alt="OVERALL Sheet">
</p>



<h3 align="center">SURVIVOR Sheet</h3>
<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/examples/GoogleSheets-Picks-Example-05.jpg" width="600" alt="SURVIVOR Sheet">
</p>


-------------------------

# **2. Example Form**
Update your form to look like this, or whatever you prefer. The script will create all the weekly entries for each matchup of the week, a survivor pool prompt, a tiebreaker entry field, and a comments section. When membership is unlocked, the form will have a text entry field, rather than the dropdown, for “Name”.

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/examples/GoogleSheets-Picks-Example-06.jpg" width="500" alt="Example Form part 1">
</p>

<h3 align="center">[MANY MATCHES LATER]</h3>

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/examples/GoogleSheets-Picks-Example-07.jpg" width="500" alt="Example Form part 2">
</p>


-------------------------

# **3. Setup Instructions**
1. Go to Google Sheets and **create a blank spreadsheet,** and give it a name → [click here to automatically create a new spreadsheet](http://spreadsheet.new)

2. **Select “Extensions” > “Apps Script”**

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/instructions/GoogleSheets-Picks-Instructions-01.jpg" width="600" alt="Apps Script Menu">
</p>

3. In the tab that opens, **replace entire code** [in this repository] in place of existing text

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/instructions/GoogleSheets-Picks-Instructions-02.jpg" width="600" alt="Replace dummy code">
</p>

4. Click the **“Save project”** icon, wait for it to save; “untitled.gs” will be renamed to “Code.gs” (change if you want)

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/instructions/GoogleSheets-Picks-Instructions-03.jpg" width="770" alt="Save Project">
</p>

5. **Ensure that ‘runFirst’ is the selected function** from the function dropdown

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/instructions/GoogleSheets-Picks-Instructions-04.jpg" width="770" alt="Select runFirst">
</p>

6. **Click “Run”** button to start initial setup

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/instructions/GoogleSheets-Picks-Instructions-05.jpg" width="770" alt="Run script">
</p>

7. After 5-10 seconds, a “Authorization required” box will appear, **click “Review permissions”**

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/instructions/GoogleSheets-Picks-Instructions-06.jpg" width="600" alt="Review Permissions">
</p>

8. **Select your preferred Google account** for managing the spreadsheet and form

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/instructions/GoogleSheets-Picks-Instructions-07.jpg" width="400" alt="Select Google Account">
</p>

9. "App isn't verified" pops up, **click “Advanced” on bottom left**

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/instructions/GoogleSheets-Picks-Instructions-08.jpg" width="400" alt="Advanced verification">
</p>

10. **Click “Go to Untitled project (unsafe)”** on bottom left

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/instructions/GoogleSheets-Picks-Instructions-09.jpg" width="400" alt="Got to project (unsafe) prompt">
</p>

11. Review permissions, scroll down and **click “Allow”**

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/instructions/GoogleSheets-Picks-Instructions-10.jpg" width="400" alt="Allow script to run">
</p>

12. The initial script will run. It’s going to make a Google Form (image below of how I customized mine) and do a lot more. **Go back to the spreadsheet to answer prompts--and be patient.**

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/instructions/GoogleSheets-Picks-Instructions-11.jpg" width="600" alt="Form Template image">
</p>

13. You should now be ready to start running the pick’ems league. **More detailed usage below. Cheers!**

-------------------------

# **4. Usage**

1. The first prompt you’ll see when you return to the sheet should be the following:

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/usage/GoogleSheets-Picks-Usage-01.jpg" width="370" alt="First Prompt">
</p>

2. You’ll next be prompted to select if you’d like to keep the members list unlocked. At the beginning of the season, you may leave the members unlocked and the form will have a text entry question for Name. Upon locking membership, you will then have a dropdown selection for the Form, which is easier and less likely to result in errors or falsely creating new entries.

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/usage/GoogleSheets-Picks-Usage-02.jpg" width="300" alt="Lock members">
</p>

3. Next you’ll be asked if you’d like to create the first form (ideally this is prior to week 1 of the regular season, though I may have succeeded in making this a robust enough tool to start midway through). Recommended to create it now, though you can create it later.

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/usage/GoogleSheets-Picks-Usage-03.jpg" width="300" alt="Create first form">
</p>

4. Lastly, you should be shown a message like the following that gives a shareable link for the form (to give to your members)

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/usage/GoogleSheets-Picks-Usage-04.jpg" width="600" alt="Form URL">
</p>

NOTE: If you ever miss the edit form link or the shareable form link, you can always find them in the “FORM” tab that usually is hidden by default; sheets can be unhidden by selecting the specific sheet in the flyout menu on the lower left (indicated with an arrow).

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/usage/GoogleSheets-Picks-Usage-05.jpg" width="600" alt="FORM Tab">
</p>

NOTE: Use the “Share” button on the upper right of the spreadsheet to allow the Spreadsheet link to be viewable by your members (be sure to set the link role to “Viewer” -- which is the default).

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/usage/GoogleSheets-Picks-Usage-06.jpg" width="400" alt="Share with friends">
</p>

-------------------------

# **5. Custom Functions Overview**
Once the scripts have completed, you should have a new menu option on the top ribbon entitled “Pick’Ems”

<p align="center">
<img src="https://github.com/boilermaker-ben/Google-NFL-Pick-Ems/blob/main/images/GoogleSheets-Picks-Functions-01.jpg?raw=true" width="600" alt="Functions menu">
</p>

**Update Form:** This function will allow you to create a new form for the week, there are safety checks to ensure you don’t erase previous entry information and it allows you to override the current week (pulled from API) to select whatever week you’d like

**Check NFL Scores:** won’t work until the first week starts; this can bring down all completed matches and the tiebreaker information from the MNF game, if available.

**Check Responses:** checks the responses in the Google Form without revealing picks so you can hound the worthless members who haven’t submitted picks yet; prompts to import if all responses are submitted.

**Import Picks:** direct function to import all pick’em information submitted, it does check responses first and confirm you’d like to submit.

**Import Thursday Picks:** in case you have lagging members who you allow to submit their picks late (and not count the Thursday game for them), this allows you to only import the Thursday night game matchup picks from your faithful members.

**Add Member:** prompts to bring in a new member; data from previous weeks will be temporarily stored in an array and then re-populated once each sheet is updated to include the new member; this is a clunky system for doing multiple additions, but you should be able to repeat as many times as needed.

**Lock Members:** recommended to lock membership down as soon as you can; at the beginning of the season, you may leave the members unlocked and the form will have a TEXT entry question for "Name", but upon locking membership, you will then have a DROPDOWN selection for the Form, which is easier and less likely to result in errors or falsely creating new entries.

**Update NFL Schedule:** pulls any changes from NFL scheduling updates (likely not needed, but some games are flexed into primetime).

**Rebuild Calculations:** updates all calculated cells on the main sheets (not weekly sheets), likely unneeded for most cases


-------------------------

Hopeful improvements for future versions:

- Google User confirmation (auto-detection for submissions, may be above my head)
- Better data management with a hidden "database" of submissions
- Add column for participants in each weekly set of games
- Picking against the spread rather than straight up option
- More metrics (suggestions welcome!)


-------------------------


Thanks for checking out the project and for making it to the end!

I'll just leave this here if anyone is feeling generous or finds this useful [https://www.buymeacoffee.com/benpowers](https://www.buymeacoffee.com/benpowers).
