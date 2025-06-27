GTM Documentation Automation - Complete Setup Guide
## Check out the YouTube guide where I walk through the steps below: http://bit.ly/44BMR7O

## Prerequisites
- Google account with access to Google Sheets
- Google Tag Manager account with minimum (Read) access
- Basic familiarity with Google Tag Manager

## Step 1: Create a New Google Sheet
- Option A (Fastest):
Type sheets.new in your browser address bar and press Enter

- Option B (Traditional):
Go to sheets.google.com
Click the "+" button to create a new sheet

## Step 2: Open Google Apps Script Editor
- In your Google Sheet, click Extensions in the top menu
- Select Apps Script from the dropdown
- A new tab will open with the Apps Script editor

## Step 3: Replace the Default Code
- You'll see a file called Code.gs with some default text
- Select all the existing code (Ctrl+A or Cmd+A)
- Delete it completely
- Paste the GTM automation code (from the GitHub repository)
- Click the Save icon (💾) or press Ctrl+S (Cmd+S on Mac)
- When prompted, give your project a name like "GTM Documentation Tool"

## Step 4: Enable the Tag Manager API Service
- In the Apps Script editor, look for Services in the left sidebar
- Click the "+" (plus) icon next to Services
- This opens the "Add a service" dialog box

## Step 5: Add Tag Manager API
- In the service dialog, scroll down to find "Tag Manager API"
- Click on "Tag Manager API"
- Click the "Add" button in the bottom-right corner of the dialog
- The dialog will close and you'll see "TagManager" added under Services

## Step 6: Run and Authenticate the Script
- In the Apps Script editor, click Run in the top toolbar
- First-time authentication process:
       -> Click "Review permissions"
       -> Choose your Google account
       -> Click "Advanced" (if you see a warning)
       -> Click "Go to [Your Project Name] (unsafe)"
       -> Click "Allow" to grant permissions

- Check the execution log:
       -> Look at the bottom panel for "Execution log"
       -> You should see "Execution started" followed by "Execution completed"
       -> If you see errors, check the troubleshooting section below

## Step 7: Access the GTM Reports Menu
- Return to your Google Sheet (switch back to the first tab)
- Refresh the page (F5 or Ctrl+R)
- Look at the top menu - you'll now see "GTM Reports" next to the Help button
- Click "GTM Reports" to see the dropdown with two options:
        -> Generate GA4 Event Tag Report
        -> Configure GTM Settings

## Step 8: Configure Your GTM Settings
- Click "Configure GTM Settings" first
- A popup window will appear with instructions
- Find your GTM IDs:
       -> Open Google Tag Manager in another tab
       -> Navigate to your desired workspace
       -> Look at the URL:  https://tagmanager.google.com/#/container/accounts/12345/containers/67890/workspaces/11121314/
       -> Extract the three numbers:
            Account ID: 12345
            Container ID: 67890
            Workspace ID: 11121314
- Enter the IDs in the configuration popup
- Click "Save Configuration"

## Step 9: Generate Your Documentation
- Click "GTM Reports" in the menu again
- Select "Generate GA4 Event Tag Report"
- The script will run and create several sheets:
       -> GTM documentation: Main report with all tags and parameters
       -> Parameters: List of all parameter values and descriptions
       -> Events: List of all event names and descriptions
       
## Step 10: Bonus - Schedule Automatic Updates
- In Apps Script editor, click the clock icon (Triggers) in the left sidebar
- Click "+ Add Trigger"
- Configure:
       -> Function: listGa4EventTagsAndParameters
       -> Event source: Time-driven
       -> Type: Day timer (for daily updates) or Week timer (for weekly)
       -> Time: Choose your preferred time
- Click "Save"
