# Google Sheets / Forms NFL Pick'Ems and Survivor Pool 
## Creation and Management Tool for Running your Own Group

Google Script to run within Google Sheets that creates a Google Form to create a season-long NFL Pick'Ems league run through a Google Sheet with a weekly updating Google Form to receive responses from members.

-------------------------

**Welcome!** The project below was developed over three seasons of NFL play to create a semi-robust way of managing an NFL pick ’ems or survivor league. It creates a series of sheets in a spreadsheet for tracking all picks through the 18 regular season games of an NFL season. It also includes a Monday Night Football most correct season-long winner, a weekly most correct winner, and a season-long most correct winner. The tool will also create a simple Google Form (questionnaire) that is used to collect responses from members that can be imported to the spreadsheet easily. Lastly, another custom script, among the many below, will pull in match results and tiebreaker scores. The final Monday Night Football game score total each week is used as the tiebreaker for the pick ‘ems weekly competition (some weeks we do have 2 MNF games). 

It’s up to the person running the league to import the picks for the week (ideally before Thursday night) and also to update the form for the coming week (usually done Tuesday or Wednesday morning to send to the members).

I was keen to help a friend create a more robust way to track a family and friends league three seasons ago and the effort resulted in this massive and complex block of code that’s over 4,500 lines. I’m not a coder by training, I’m an industrial designer. I hope it doesn’t break for you--but let me know if it does! If you’re inclined and have enjoyed the script and care to support my wife, my kiddos, and me, you can [buy me a coffee](https://www.buymeacoffee.com/benpowers)--no pressure though, I’m just excited that you’re using this tool!


**Disclaimer:** This set of functions relies on the use of the ESPN API for pulling NFL game data. Here are the ESPN terms of use of their API. I’m sharing these scripts with you with the intent that you are taking on the responsibilities of the terms of use for your own personal use and don’t condone or endorse your use of the code here for monetization of “apps” or any other content. The terms outline the need for an “Information Form” to be submitted by a parent or guardian if you are a minor. This content is not intended to be published nor executed outside of the use by personal users. 

Lastly, there are some safeguards Google has in place to avoid allowing users to execute any malicious code from the Google Scripts console. Please feel free to review the code below, as it contains no functions to share information, transfer information, or send emails. Information only travels between your personal Google Sheet and your Google Form (copied from a template form) that are created in the process outlined below. All sharing of content must be done by you directly (via the links that are created), such as sharing the link to the Google Form with your members and sharing a “view only” version of the spreadsheet with them to allow the members to see their league’s standings.

-------------------------

**Notable Changes**
This newest update to version 2.0 includes the following new improved features/changes:
- Creation of a **Drive Folder for storing Google Forms** (as there will be multiple, see next bullet)
- **A new Form for each week is created**, rather than the same Form being cleared and re-created
  - This means that all previous picks and data will be backed up, so to speak, in a previous week Form
  - Be sure to get the shortened URL from the pop-up that occurs after Form creation (or find it in the CONFIG tab) to send to members
- **Survivor pool ONLY option** (I had multiple requests for this, as many of you don’t want to do a weekly pick ‘ems)
- **Survivor pool restart** feature
 - Last season, my group enjoyed restarting the survivor competition after about week 7 with how quickly it ended due to the nature of the NFL
 - If the entire pool is eliminated (should be auto-detected), the “Create Form” menu function script should prompt to restart the survivor pool
- **Member-specific Survivor** Form questions
  - After the first week, this will create a specific follow up question to name selection in the Google Form to only allow each user to select a team from their remaining Survivor teams. This helps avoid someone accidentally submitting a team they’ve already used.
  - This is likely the most complex addition to the script and does make form creation a little slower
- **New NFL Outcomes sheet** (see representation in Example Sheets section)
  - If running a survivor-only group, this tab is where the weekly winners will be entered
  - If running a group with pick ‘ems included, this is for reference only (as a formula for populating the cells will be present
  - If the whole column of games is populated for the given week, this sheet will be what the “Create Form” function will assess in order to suggest the next week
- Tons of efficiencies, error checks, and prompts added in hopes of making it more intuitive and easier to use

-------------------------

## **Table of Contents**


### **1. Example Sheets** - Screenshots of the output from a league done in 2021

- NFL_OUTCOMES Sheet
  
- WEEKLY Sheet
  
- SUMMARY Sheet
  
- OVERALL Sheet
  
- MNF Sheet
  
- SURVIVOR Sheet
  
### **2. Example Form** - Screenshot of form from week 18 in 2021

### **3. Setup Instructions** - create new document, create script, paste code, run initial setup

### **4. Usage** - how to use the tool

### **5. Custom Functions Overview** - description of all custom functions in the “Picks” menu

-------------------------

# **1. Example Sheets**

<h3 align="center">NFL OUTCOMES Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-pickems-outcomes-sheet.png" width="600" alt="WEEKLY Sheet">
</p>

<h3 align="center">WEEKLY Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/09/googlesheets-pickems-outcomes-sheet.png" width="600" alt="WEEKLY Sheet">
</p>



<h3 align="center">SUMMARY Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example02.png" width="600" alt="SUMMARY Sheet">
</p>



<h3 align="center">MNF Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example03.png" width="600" alt="MNF Sheet">
</p>



<h3 align="center">OVERALL Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example04.png" width="600" alt="OVERALL Sheet">
</p>



<h3 align="center">SURVIVOR Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example05.png" width="600" alt="SURVIVOR Sheet">
</p>


-------------------------

# **2. Example Form**
Update your form to look like this, or whatever you prefer. The script will create all the weekly entries for each matchup of the week, a survivor pool prompt, a tiebreaker entry field, and a comments section. When membership is unlocked, the form will have a text entry field, rather than the dropdown, for “Name”.

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example06.png" width="500" alt="Example Form part 1">
</p>

<h3 align="center">[MANY MATCHES LATER]</h3>

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example07.png" width="500" alt="Example Form part 2">
</p>


-------------------------

# **3. Setup Instructions**
1. Go to Google Sheets and **create a blank spreadsheet,** and give it a name → [click here to automatically create a new spreadsheet](http://spreadsheet.new)

2. **Select “Extensions” > “Apps Script”**

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions01.png" width="600" alt="Apps Script Menu">
</p>

3. In the tab that opens, **replace entire code** [in this repository] in place of existing text

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions02.png" width="600" alt="Replace dummy code">
</p>

4. Click the **“Save project”** icon, wait for it to save; “untitled.gs” will be renamed to “Code.gs” (change if you want)

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions03.png" width="770" alt="Save Project">
</p>

5. **Ensure that ‘runFirst’ is the selected function** from the function dropdown

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions04.png" width="770" alt="Select runFirst">
</p>

6. **Click “Run”** button to start initial setup

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions05.png" width="770" alt="Run script">
</p>

7. After 5-10 seconds, a “Authorization required” box will appear, **click “Review permissions”**

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions06.png" width="600" alt="Review Permissions">
</p>

8. **Select your preferred Google account** for managing the spreadsheet and form

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions07.png" width="400" alt="Select Google Account">
</p>

9. "App isn't verified" pops up, **click “Advanced” on bottom left**

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions08.png" width="400" alt="Advanced verification">
</p>

10. **Click “Go to Untitled project (unsafe)”** on bottom left

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions09.png" width="400" alt="Got to project (unsafe) prompt">
</p>

11. Review permissions, scroll down and **click “Allow”**

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions10.png" width="400" alt="Allow script to run">
</p>

12. The initial script will run. It’s going to make a Google Form (image below of how I customized mine) and do a lot more. **Go back to the spreadsheet to answer prompts--and be patient.**

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions11.png" width="600" alt="Form Template image">
</p>

13. You should now be ready to start running the pick’ems league. **More detailed usage below. Cheers!**

-------------------------

# **4. Usage**

1. You’ll be guided through setting up the group. These questions will be asked:
 - **Timezone confirmation** (will prompt to correct and cancel setup if it's not set correctly)
 - **Name of group** customization (defaults to “NFL Pick ‘Ems” or “NFL Survivor Pool”)
 - **Inclusion of pick ‘ems** (in the event you want to only run a survivor league)
 - **Inclusion of a Monday Night Football** tally competition
 - **Inclusion of a comments** box in the Form
 - **Inclusion of a survivor pool** (will automatically skip if you don’t include pick ‘ems pool)
 - **Locked membership** (if you do or don’t want to allow users to enter their own name in the Form)
 - Creation of previous blank pick ‘ems weeks (for manually entering if you start this later in the season)
 - Form creation (to make the first week’s form upon completion of setup)
 - **Initial member list** (you can enter names separated by commas: “Bobby, Billy, George”
2. Weekly usage:
 - **Share the Form** with your group
 - **Check responses** via the menu and **import picks** (ideally before the Thursday night game, if present).
 - Through the weekend, as games are completed you should be able to run the “Check NFL Scores” function and **import game outcomes** via that method
 - Survivor Only: Alternatively, enter the game outcomes manually on the NFL Outcomes sheet
 - Pick ‘Ems: Alternatively, enter the game outcomes manually across the top of the correct sheet with the name following this format: “YYYY_WW” (year and week pick ‘ems sheet). Note: Be sure to enter a tiebreaker (sum of the last MNF game score) value in the cell below the one labeled “Tiebreaker” or the weekly winner won’t be declared!
 - Upon completing the week (usually after the MNF game), you can **run the “Create Form” function again**, read through the prompts, and start the process over again for the next week
 - **Repeat**
3. CONFIG sheet example below
 - CONFIG is hidden by default and isn’t required to run the group
 - Hopefully the only reason you bring it up is to fetch the short URL for the Google form
 - You can manually disable survivor or pick ‘ems (or enable) through this sheet
 - Don’t edit the “SURVIVOR DONE” cell unless you’re having an issue (it’s formula-driven)
 - You can make the CONFIG sheet visible by going to the hamburger menu on the bottom and selecting it from the list (see red arrow in image below)
<h3 align="center">CONFIG Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/09/googlesheets-pickems-config-sheet.png" width="400" alt="WEEKLY Sheet">
</p>


-------------------------

# **5. Custom Functions Overview**
Once the scripts have completed, you should have a new menu option on the top ribbon entitled “Picks”

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/09/googlesheets-picks-menu.png" width="600" alt="Functions menu">
</p>

**Create Form:** This function will allow you to create a new form for the week, there are safety checks to ensure you don’t erase previous entry information and it allows you to decline creating a form for the proposed week and enter your own
**Check NFL Scores:** won’t work until the first week starts; this can bring down all completed matches and the tiebreaker information from the MNF game, if available.
**Check Responses:** checks the responses in the Google Form without revealing picks so you can hound the worthless members who haven’t submitted picks yet; prompts to import if all responses are submitted and checks for new users
**Import Picks:** direct function to import all pick’em information submitted, it does check responses first and confirm you’d like to submit as well as checking for new members first
**Import Thursday Picks:** in case you have lagging members who you allow to submit their picks late (and not count the Thursday game for them), this allows you to only import the Thursday night game matchup picks from your faithful members
**Add Member:** prompts to bring in a new member or multiple (comma-separated) members. This will add them to the survivor activity, if present and in the first week of competition, otherwise just adds them to a pick ‘ems pool
**Reopen Members / Lock Members:** toggles between whether you can add members or not, will add “New User” option in the Form or remove it and will add or remove the “Add Member” function in the menu
**Update NFL Schedule:** pulls any changes from NFL scheduling updates (likely not needed, but some games are flexed into primetime). You will also be prompted each week if you’d like to update the NFL schedule data, which doesn’t hurt, but adds a few seconds to the process

-------------------------

Hopeful improvements for future versions:

- Google User confirmation (auto-detection for submissions, may be above my head)
- Reorganize member names alphabetically as an option
- Clear former Forms and avoid re-pulling a previous form if the name of the group is unchanged
- Add column for participants in each weekly set of games
- Picking against the spread rather than straight up option  (tough due to timing of when people submit and the changing nature of lines)
- Multiple entries per user
- Option to have user removed upon submission from Form
- Option to "revive" in survivor pool (possibly with two correct picks)
- Column for payment/entry fee received per weekly sheet
- Adjust naming conventions for named ranges and weekly sheets
- Member rename function
- More metrics (suggestions welcome!)

-------------------------

Thanks for checking out the project and for making it to the end!

