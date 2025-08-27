# Google Sheets / Forms NFL Pick'Ems and Survivor Pool 
## Creation and Management Tool for Running your Own Group

Google Sheet document with multiple script files to generate Google Forms for season-long NFL Pick'Ems or Survivor league management

-------------------------

**NOTE: I am developing a NEW VERSION FOR 2025 and hope to get it going before the season starts--it should be worth the wait!**

If you can't wait for a few more days, then you can likely just use last year's version without any issue (it'll just be missing a lot of the cool new features I'm adding in)

**TLDR: Go [here](https://docs.google.com/spreadsheets/d/12ZV1sEuEw4J-giEuWhY_rly11Gk6jFfWQOgfeAPA_JU/edit?usp=sharing) and make a copy of the sheet. Follow prompts. Enjoy!**

-------------------------


**Welcome!** The project below was developed over three seasons of NFL play (maybe NCAAF eventually) to create a semi-robust way of managing an NFL pick ’ems or survivor league. Copying the file will enable you to customize and generate a series of sheets in your copy of the spreadsheet for tracking all picks through the 18 regular season games of an NFL season. It also includes a Monday Night Football most correct season-long winner, a weekly most correct winner, and a season-long most correct winner. The tool will also create a weekly Google Form (questionnaire) that is used to collect responses from members that can be imported to the spreadsheet easily. Match results and tiebreaker scores can also be pulled in via scripts in the 'Picks' menu. The final Monday Night Football game score total each week is used as the tiebreaker for the pick ‘ems weekly competition (some weeks we do have 2 MNF games). Tiebreakers, comments, exclusion of Thursday games, and more can be disabled/enabled via the setup.

It’s up to the person running the league to import the picks for the week (ideally before Thursday night) and also to update the form for the coming week (usually done Tuesday or Wednesday morning to send to the members).

I was keen to help a friend create a more robust way to track a family and friends league three seasons ago and the effort resulted in this massive and complex block of thousands of lines of code. I’m not a coder by training, I’m an industrial designer and product manager. I hope it doesn’t break for you--but let me know if it does! If you’re inclined and have enjoyed the script and care to support my wife, my five kiddos, and me, you can [buy me a coffee](https://www.buymeacoffee.com/benpowers)--no pressure though, I’m just excited that you’re using this tool!


**Disclaimer:** This set of functions relies on the use of the ESPN API for pulling NFL game data. You can find the ESPN terms of use [here](https://disneytermsofuse.com/). I’m sharing these scripts with you with the intent that you are taking on the responsibilities of the terms of use for your own personal use and don’t condone or endorse your use of the code here for monetization of “apps” or any other content. The terms outline the need for an “Information Form” to be submitted by a parent or guardian if you are a minor. This content is not intended to be published nor executed outside of the use by personal users. 

Lastly, there are some safeguards Google has in place to avoid allowing users to execute any malicious code from the Google Scripts console. Please feel free to review the code, as it contains no functions to share information, transfer information, or send emails. Information only travels between your personal Google Sheet and your Google Form (copied from a template form) that are created in the process outlined below. All sharing of content must be done by you directly (via the links that are created), such as sharing the link to the Google Form with your members and sharing a “view only” version of the spreadsheet with them to allow the members to see their league’s standings.

-------------------------

**Notable Changes**
This newest update to version 2.6 includes the following new improved features/changes:
- **More intuitive setup and initialization tool**
- Use of **Script Properties**
  - Storing timezone, year, and a few other fixed values
  - Stores Form responses by a triggered script that vastly reduces import time
- New **Bonus** features of a 2x or 3x multiplier for games
  - Set your MNF games to always count as double as an option
  - Randomly pick a "Game of the Week" via the menu for added randomization
  - Manually adjust game weighting using the bottom row within the sheet
- Updated **color-coding of sheet tabs**
- Fixed an issue with a color-banding attempt that caused issues loading on mobile
- Added a "Help & Support" HTML popup with a few links
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

-------------------------

## **Table of Contents**


### **1. Example Sheets** - Screenshots of the output from a league done in 2021 (Some of these have yet to be updated to visually represent the new format)

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

<h3 align="center">WEEKLY Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2024/08/2024_weekly_sheet.png" width="600" alt="WEEKLY Sheet">
</p>

<h3 align="center">NFL OUTCOMES Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/09/googlesheets-pickems-outcomes-sheet.png" width="600" alt="NFL OUTCOMES">
</p>

<h3 align="center">SUMMARY Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2024/08/2024_summary_sheet.png" width="600" alt="SUMMARY Sheet">
</p>

<h3 align="center">MNF Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example03.png" width="600" alt="MNF Sheet">
</p>

<h3 align="center">OVERALL Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example04.png" width="600" alt="OVERALL Sheet">
</p>

<h3 align="center">RANK Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2024/08/2024_rank_sheet.png" width="600" alt="RANK Sheet">
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
1. Go to my Google Sheet and **create a copy,** → [click here to open the spreadsheet](https://docs.google.com/spreadsheets/d/12ZV1sEuEw4J-giEuWhY_rly11Gk6jFfWQOgfeAPA_JU/edit?usp=sharing)

2. An onOpen trigger will prompt for a timezone and give you instructions.

3. Once you run the "Initialize Sheet" function, an “Authorization required” box will appear, **click “Review permissions”**

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions06.png" width="600" alt="Review Permissions">
</p>

4. **Select your preferred Google account** for managing the spreadsheet and form

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions07.png" width="400" alt="Select Google Account">
</p>

5. "App isn't verified" pops up, **click “Advanced” on bottom left**

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions08.png" width="400" alt="Advanced verification">
</p>

6. **Click “Go to Untitled project (unsafe)”** on bottom left

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions09.png" width="400" alt="Got to project (unsafe) prompt">
</p>

7. Review permissions, scroll down and **click “Allow”**

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions10.png" width="400" alt="Allow script to run">
</p>

8. You should be able to now **re-run the "Initialize Sheet"** function. Fill out the HTML popup questionnaire and let the script do its thing.

4. A week 1 form should be prompted for creation, but if not, **use the "Picks" menu to start off your first week.**

5. Most functions are self-explanatory, but please go to the **"Extensions" > "Apps Script" > "config.gs"** for some detail on each function


-------------------------

1. Weekly usage:
 - **Share the Form** with your group
 - **Check responses** via the menu and **import picks** (ideally before the Thursday night game, if present).
 - Through the weekend, as games are completed you should be able to run the “Check NFL Scores” function and **import game outcomes** via that method
 - Survivor Only: Alternatively, enter the game outcomes manually on the NFL Outcomes sheet
 - Pick ‘Ems: Alternatively, enter the game outcomes manually across the bottom of the correct weekly sheet. Note: If using a tiebreaker (sum of the last MNF game score), be sure to enter it in the cell to the right of the final match column or the weekly winner won’t be declared!
 - Upon completing the week (usually after the MNF game), you can **run the “Create Form” function again**, read through the prompts, and start the process over again for the next week
 - **Repeat**
2. CONFIG sheet example below (THIS MAY NOT REPRESENT ALL THE CURRENT OPTIONS)
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

Hopeful improvements for future versions:

- Google User confirmation (auto-detection for submissions, may be above my head)
- Reorganize member names alphabetically as an option
- Add column for participation toggling in each weekly set of games
- Picking against the spread rather than straight up option (tough due to timing of when people submit and the changing nature of lines)
- Multiple entries per user
- Option to have user removed upon submission from Form to avoid duplication
- Option to "revive" in survivor pool (possibly with two correct picks)
- Column for payment/entry fee received per weekly sheet
- Adjust naming conventions for named ranges and weekly sheets
- Member rename function
- Confidence pick 'ems capability
- Opting out of survivor competition in the Form
- NCAA Football capability
- More metrics (suggestions welcome!)

-------------------------

Thanks for checking out the project and for making it to the end!

