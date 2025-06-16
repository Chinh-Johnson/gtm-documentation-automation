/**
 * Adds a custom menu to the Google Sheet UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GTM Reports')
      .addItem('Generate GA4 Event Tag Report', 'listGa4EventTagsAndParameters')
      .addSeparator()
      .addItem('Configure GTM Settings', 'showConfigurationDialog')
      .addToUi();
}

/**
 * Shows a dialog for configuring GTM account, container, and workspace IDs.
 */
function showConfigurationDialog() {
  // Make sure the Config sheet exists
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
    
    // Create detailed instructions in a merged cell at the top
    configSheet.getRange('A1:B1').merge();
    const instructionsText = 'Open Google Tag Manager, select the workspace you want, and find the IDs in the URL. For example, if the URL is https://tagmanager.google.com/#/container/accounts/12345/containers/45678/workspaces/7891011/ then:';
    configSheet.getRange('A1').setValue(instructionsText);
    configSheet.getRange('A1:B1').setBackground('#ffffcc');
    configSheet.getRange('A1:B1').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    
    // Add a blue "account ID" hyperlink
    const richValue = SpreadsheetApp.newRichTextValue()
        .setText(instructionsText)
        .setLinkUrl(instructionsText.indexOf('12345'), instructionsText.indexOf('12345') + 5, 'https://tagmanager.google.com/')
        .setTextStyle(instructionsText.indexOf('12345'), instructionsText.indexOf('12345') + 5, SpreadsheetApp.newTextStyle().setForegroundColor('#0000ff').build())
        .build();
    configSheet.getRange('A1').setRichTextValue(richValue);
    
    // Create header row
    configSheet.appendRow(['Setting', 'Value']);
    
    // Add settings rows with concise descriptions
    configSheet.appendRow(['GTM Account ID', '']);
    configSheet.appendRow(['GTM Container ID', '']);
    configSheet.appendRow(['GTM Workspace ID', '']);
    
    // Format the sheet
    configSheet.setFrozenRows(2);  // Freeze both instruction and header rows
    configSheet.getRange('A2:C2').setFontWeight('bold');
    configSheet.setColumnWidth(1, 150);
    configSheet.setColumnWidth(2, 150);
    configSheet.setColumnWidth(3, 300);
    configSheet.getRange('C:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  } else {
    // For existing sheets, ensure the text wrapping is applied
    configSheet.getRange('C:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  
  // Get current values (adjust row numbers to account for the instruction row)
  const configData = configSheet.getRange('A3:B5').getValues();
  const accountId = configData[0][1] || '';
  const containerId = configData[1][1] || '';
  const workspaceId = configData[2][1] || '';
  
  // Create and show the dialog
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 20px;
            line-height: 1.6;
          }
          .form-group {
            margin-bottom: 15px;
          }
          label {
            display: block;
            font-weight: bold;
            margin-bottom: 5px;
          }
          input[type="text"] {
            width: 100%;
            padding: 8px;
            box-sizing: border-box;
            border: 1px solid #ccc;
            border-radius: 4px;
          }
          .text {
            background-color: #ffffcc;
            margin-bottom: 15px;
            padding: 10px;
            border-radius: 4px;
          }
          .button {
            background-color: #4285f4;
            border: none;
            color: white;
            padding: 10px 20px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 14px;
            margin: 4px 2px;
            cursor: pointer;
            border-radius: 4px;
          }
          .error {
            color: red;
            margin-top: 5px;
          }
          .url-example {
            word-break: break-all;
            display: inline-block;
            margin-top: 5px;
          }
          .bold-id {
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <h2>GTM Configuration</h2>
        <div class="text">
          <p>Open Google Tag Manager, select the workspace you want, and find the ID after "accountid" in the URL. For example, if the URL is</p>
          <span class="url-example">https://tagmanager.google.com/#/container/accounts/12345/containers/45678/workspaces/7891011/</span>
          <ul>
            <li>account ID is <span class="bold-id">12345</span></li>
            <li>container ID is <span class="bold-id">45678</span></li>
            <li>workspace ID is <span class="bold-id">7891011</span></li>
          </ul>
        </div>
        <div class="form-group">
          <label for="accountId">GTM Account ID:</label>
          <input type="text" id="accountId" value="${accountId}" placeholder="e.g., 12345">
        </div>
        <div class="form-group">
          <label for="containerId">GTM Container ID:</label>
          <input type="text" id="containerId" value="${containerId}" placeholder="e.g., 45678">
        </div>
        <div class="form-group">
          <label for="workspaceId">GTM Workspace ID:</label>
          <input type="text" id="workspaceId" value="${workspaceId}" placeholder="e.g., 7891011">
        </div>
        <div id="error" class="error"></div>
        <button class="button" onclick="saveConfig()">Save Configuration</button>
        
        <script>
          function saveConfig() {
            const accountId = document.getElementById('accountId').value.trim();
            const containerId = document.getElementById('containerId').value.trim();
            const workspaceId = document.getElementById('workspaceId').value.trim();
            
            // Basic validation
            if (!accountId || !containerId || !workspaceId) {
              document.getElementById('error').textContent = 'All fields are required.';
              return;
            }
            
            // Validate that inputs are numeric
            if (!/^\\d+$/.test(accountId) || !/^\\d+$/.test(containerId) || !/^\\d+$/.test(workspaceId)) {
              document.getElementById('error').textContent = 'All IDs must be numeric values.';
              return;
            }
            
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                document.getElementById('error').textContent = error.message || 'An error occurred while saving.';
              })
              .saveGtmConfig(accountId, containerId, workspaceId);
          }
        </script>
      </body>
    </html>
  `)
  .setWidth(450)
  .setHeight(450)
  .setTitle('GTM Configuration');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'GTM Configuration');
}


// Built-in GTM variable descriptions from Google documentation
const BUILTIN_VARIABLE_DESCRIPTIONS = {
  // Clicks
  'Click Element': {
    description: 'Accesses the gtm.element key in the dataLayer, which is set by Click triggers. This will be a reference to the DOM element where the click occurred.',
    type: 'object'
  },
  'Click Classes': {
    description: 'Accesses the gtm.elementClasses key in the dataLayer, which is set by Click triggers. This will be the string value of the classes attribute on the DOM element that was clicked.',
    type: 'string'
  },
  'Click ID': {
    description: 'Accesses the gtm.elementId key in the dataLayer, which is set by Click triggers. This will be the string value of the id attribute on the DOM element that was clicked.',
    type: 'string'
  },
  'Click Target': {
    description: 'Accesses the gtm.elementTarget key in the dataLayer, which is set by Click triggers.',
    type: 'string'
  },
  'Click URL': {
    description: 'Accesses the gtm.elementUrl key in the dataLayer, which is set by Click triggers.',
    type: 'string'
  },
  'Click Text': {
    description: 'Accesses the gtm.elementText key in the dataLayer, which is set by Click triggers.',
    type: 'string'
  },
  
  // Errors
  'Error Message': {
    description: 'Accesses the gtm.errorMessage key in the dataLayer, which is set by JavaScript Error triggers. This will be a string that contains the error message.',
    type: 'string'
  },
  'Error URL': {
    description: 'Accesses the gtm.errorUrl key in the dataLayer, which is set by JavaScript Error triggers. This will be a string that contains the URL where the error occurred.',
    type: 'string'
  },
  'Error Line': {
    description: 'Accesses the gtm.errorLine key in the dataLayer, which is set by JavaScript Error triggers. This will be a number of the line in the file where the error occurred.',
    type: 'number'
  },
  'Debug Mode': {
    description: 'Returns true if the container is currently in preview mode.',
    type: 'boolean'
  },
  
  // Forms
  'Form Classes': {
    description: 'Accesses the gtm.elementClasses key in the dataLayer, which is set by Form triggers. This will be the string value of the classes attribute on the form.',
    type: 'string'
  },
  'Form Element': {
    description: 'Accesses the gtm.element key in the dataLayer, which is set by Form triggers. This will be a reference to the form\'s DOM element.',
    type: 'object'
  },
  'Form ID': {
    description: 'Accesses the gtm.elementId key in the dataLayer, which is set by Form triggers. This will be the string value of the id attribute on the form.',
    type: 'string'
  },
  'Form Target': {
    description: 'Accesses the gtm.elementTarget key in the dataLayer, which is set by Form triggers.',
    type: 'string'
  },
  'Form Text': {
    description: 'Accesses the gtm.elementText key in the dataLayer, which is set by Form triggers.',
    type: 'string'
  },
  'Form URL': {
    description: 'Accesses the gtm.elementUrl key in the dataLayer, which is set by Form triggers.',
    type: 'string'
  },
  
  // History
  'History Source': {
    description: 'Accesses the gtm.historyChangeSource key in the dataLayer, which is set by History Change triggers.',
    type: 'string'
  },
  'New History Fragment': {
    description: 'Accesses the gtm.newUrlFragment key in the dataLayer, which is set by History Change triggers. Will be the string value of the fragment (aka hash) portion of the page\'s URL after the history event.',
    type: 'string'
  },
  'New History State': {
    description: 'Accesses the gtm.newHistoryState key in the dataLayer, which is set by History Change triggers. Will be the state object that the page pushed onto the history to cause the history event.',
    type: 'object'
  },
  'Old History Fragment': {
    description: 'Accesses the gtm.oldUrlFragment key in the dataLayer, which is set by History Change triggers. Will be the string value of the fragment (aka hash) portion of the page\'s URL before the history event.',
    type: 'string'
  },
  'Old History State': {
    description: 'Accesses the gtm.oldHistoryState key in the dataLayer, which is set by History Change triggers. Will be the state object that was active before the history event took place.',
    type: 'object'
  },
  
  // Pages
  'Page Hostname': {
    description: 'Provides the hostname portion of the current URL.',
    type: 'string'
  },
  'Page Path': {
    description: 'Provides the path portion of the current URL.',
    type: 'string'
  },
  'Page URL': {
    description: 'Provides the full URL of the current page.',
    type: 'string'
  },
  'Referrer': {
    description: 'Provides the full referrer URL for the current page.',
    type: 'string'
  },
  
  // Scroll
  'Scroll Depth Threshold': {
    description: 'Accesses the gtm.scrollThreshold key in the dataLayer, which is set by Scroll Depth triggers. This will be a numeric value that indicates the scroll depth that caused the trigger to fire.',
    type: 'number'
  },
  'Scroll Depth Units': {
    description: 'Accesses the gtm.scrollUnits key in the dataLayer, which is set by Scroll Depth triggers. This will be either \'pixels\' or \'percent\', that indicates the unit specified for the threshold.',
    type: 'string'
  },
  'Scroll Direction': {
    description: 'Accesses the gtm.scrollDirection key in the dataLayer, which is set by Scroll Depth triggers. This will be either \'vertical\' or \'horizontal\', that indicates the direction of the threshold.',
    type: 'string'
  },
  
  // Utilities
  'Container ID': {
    description: 'Provides the container\'s public ID. Example value: GTM-XKCD11',
    type: 'string'
  },
  'Container Version': {
    description: 'Provides the version number of the container, as a string.',
    type: 'string'
  },
  'Environment Name': {
    description: 'Returns the user-provided name of the current environment, if the container request was made from an environment "Share Preview" link or from an environment snippet.',
    type: 'string'
  },
  'Event': {
    description: 'Accesses the event key in the dataLayer, which is the name of the current dataLayer event (e.g. gtm.js, gtm.dom, gtm.load, or custom event names).',
    type: 'string'
  },
  'HTML ID': {
    description: 'Allows custom HTML tags to signal if they had succeeded or failed; used with tag sequencing.',
    type: 'string'
  },
  'Random Number': {
    description: 'Returns a random number value.',
    type: 'number'
  },
  
  // Videos
  'Video Current Time': {
    description: 'Accesses the gtm.videoCurrentTime key in the dataLayer, which is an integer that represents the time in seconds at which an event occurred in the video.',
    type: 'number'
  },
  'Video Duration': {
    description: 'Accesses the gtm.videoDuration key in the dataLayer, which is an integer that represents the total duration of the video in seconds.',
    type: 'number'
  },
  'Video Percent': {
    description: 'Accesses the gtm.VideoPercent key in the dataLayer, which is an integer (0-100) that represents the percent of video played at which an event occurred.',
    type: 'number'
  },
  'Video Provider': {
    description: 'Accesses the gtm.videoProvider key in the dataLayer, which is set by YouTube Video triggers. This will be the name of the video provider, i.e. \'YouTube\'.',
    type: 'string'
  },
  'Video Status': {
    description: 'Accesses the gtm.videoStatus key in the dataLayer, which is the state of the video when an event was detected, e.g. \'play\', \'pause\'.',
    type: 'string'
  },
  'Video Title': {
    description: 'Access the gtm.videoTitle key in the dataLayer, which is set by YouTube Video triggers. This will be the title of the video.',
    type: 'string'
  },
  'Video URL': {
    description: 'Access the gtm.videoUrl key in the dataLayer, which is set by YouTube Video triggers. This will be the URL of the video.',
    type: 'string'
  },
  'Video Visible': {
    description: 'Access the gtm.videoVisible key in the dataLayer, which is set by YouTube Video triggers. This will be set to true if the video is visible in the viewport.',
    type: 'boolean'
  },
  
  // Visibility
  'Percent Visible': {
    description: 'Accesses the gtm.visibleRatio key in the dataLayer, which is set by Element Visibility triggers. This will be a numeric value (0-100) that indicates how much of the selected element is visible.',
    type: 'number'
  },
  'On-Screen Duration': {
    description: 'Accesses the gtm.visibleTime key in the dataLayer, which is set by Element Visibility triggers. This will be a numeric value that indicates how many milliseconds the selected element has been visible.',
    type: 'number'
  }
};

// Helper function to format built-in variable descriptions
function formatBuiltInDescription(variableInfo) {
  return `parameter_description: <<Built-in: ${variableInfo.description}>>\nparameter_type: <<${variableInfo.type}>>`;
}


/**
 * Saves the GTM configuration to the Config sheet.
 */
function saveGtmConfig(accountId, containerId, workspaceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  // Update the values (adjust row numbers to account for the instruction row)
  configSheet.getRange('B3').setValue(accountId);
  configSheet.getRange('B4').setValue(containerId);
  configSheet.getRange('B5').setValue(workspaceId);
  
  return { success: true };
}

/**
 * Gets the GTM configuration from the Config sheet.
 * Returns default values if the Config sheet doesn't exist or values are not set.
 */
function getGtmConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  // Default values
  let config = {
    accountId: '',
    containerId: '',
    workspaceId: ''
  };
  
  if (configSheet) {
    // Apply text wrapping to ensure consistency
    configSheet.getRange('C:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    
    // Get values (adjust row numbers to account for the instruction row)
    const configData = configSheet.getRange('A3:B5').getValues();
    config.accountId = configData[0][1] || '';
    config.containerId = configData[1][1] || '';
    config.workspaceId = configData[2][1] || '';
  }
  
  return config;
}

/**
 * Converts a header string to lowercase with underscores
 * e.g., "Tag Name" -> "tag_name"
 */
function formatHeaderName(header) {
  // Convert to lowercase and replace spaces with underscores
  return header.toLowerCase().replace(/\s+/g, '_');
}

/**
 * Fetches GA4 Event Tags and their parameters from Google Tag Manager
 * and writes the data to the "GTM documentation" sheet.
 * Also updates the Parameters sheet with new parameter values
 * and the Events sheet with event names and descriptions.
 */
function listGa4EventTagsAndParameters() {
  // Get configuration from the Config sheet
  const config = getGtmConfig();
  
  // Check if required config values are set
  if (!config.accountId || !config.containerId || !config.workspaceId) {
    SpreadsheetApp.getUi().alert(
      'GTM Configuration Missing',
      'Please configure your GTM Account ID, Container ID, and Workspace ID using the "Configure GTM Settings" menu option.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // Open the configuration dialog
    showConfigurationDialog();
    return;
  }
  
  const accountId = config.accountId;
  const containerId = config.containerId;
  const workspaceId = config.workspaceId;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create the "GTM documentation" sheet
  let docSheet;
  try {
    docSheet = ss.getSheetByName("GTM documentation");
    if (!docSheet) {
      Logger.log("'GTM documentation' sheet not found. Creating one.");
      docSheet = ss.insertSheet("GTM documentation");
    }
  } catch (error) {
    Logger.log(`Error accessing 'GTM documentation' sheet: ${error.toString()}`);
    SpreadsheetApp.getUi().alert('Error', 'Could not access or create the "GTM documentation" sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Get the Parameters sheet for tracking parameter values
  let parametersSheet;
  try {
    parametersSheet = ss.getSheetByName("Parameters");
    if (!parametersSheet) {
      Logger.log("Parameters sheet not found. Creating one.");
      parametersSheet = ss.insertSheet("Parameters");
      parametersSheet.appendRow(["Parameter Value", "Description"]);
    }
  } catch (error) {
    Logger.log(`Error accessing Parameters sheet: ${error.toString()}`);
  }
  
  // Get the Events sheet for tracking event names and descriptions
  let eventsSheet;
  try {
    eventsSheet = ss.getSheetByName("Events");
    if (!eventsSheet) {
      Logger.log("Events sheet not found. Creating one.");
      eventsSheet = ss.insertSheet("Events");
      eventsSheet.appendRow(["Event Name", "Description"]);
    }
  } catch (error) {
    Logger.log(`Error accessing Events sheet: ${error.toString()}`);
  }
  
  // Load existing parameters from the Parameters sheet
  const existingParameters = new Map(); // Changed to Map to store both value and description
  if (parametersSheet) {
    const parameterData = parametersSheet.getRange(2, 1, Math.max(1, parametersSheet.getLastRow() - 1), 2).getValues();
    parameterData.forEach(row => {
      if (row[0]) existingParameters.set(row[0], row[1] || "");
    });
    Logger.log(`Loaded ${existingParameters.size} existing parameters from Parameters sheet.`);
  }
  
  // Load existing events from the Events sheet
  const existingEvents = new Map(); // Map to store event name and description
  if (eventsSheet) {
    const eventData = eventsSheet.getRange(2, 1, Math.max(1, eventsSheet.getLastRow() - 1), 2).getValues();
    eventData.forEach(row => {
      if (row[0]) existingEvents.set(row[0], row[1] || "");
    });
    Logger.log(`Loaded ${existingEvents.size} existing events from Events sheet.`);
  }
  
  // Set to track new parameters we find during processing
  const newParameters = new Map(); // Map to store parameter value and description
  
  // Set to track new events we find during processing
  const newEvents = new Map(); // Map to store event name and description
  
  // NEW: Add maps to track updates to existing parameters and events
  const parametersToUpdate = new Map(); // To track parameters that need description updates
  const eventsToUpdate = new Map(); // To track events that need description updates

  // Clear previous data in the GTM documentation sheet
  docSheet.clear();

  // Define original header names
  const originalHeaders = [
    'Tag Name',
    'Trigger',
    'Trigger Type',
    'Event Name',
    'Event Description',
    'Parameter Name',
    'Parameter Value',
    'Parameter Description'
  ];
  
  // Convert headers to lowercase with underscores
  const formattedHeaders = originalHeaders.map(header => formatHeaderName(header));
  
  // Set header row with formatted headers
  docSheet.appendRow(formattedHeaders);
  
  // Apply header formatting
  docSheet.getRange(1, 1, 1, formattedHeaders.length).setFontWeight('bold');
  docSheet.setFrozenRows(1);

  // Base path for API calls using the TagManager service
  const apiPath = `accounts/${accountId}/containers/${containerId}/workspaces/${workspaceId}`;

  let triggersMap = {}; // Map trigger ID to trigger object
  let variableNotesMap = {}; // Map variable name to its notes

  try {
    // 1. Fetch Triggers using the TagManager advanced service
    const triggersResponse = TagManager.Accounts.Containers.Workspaces.Triggers.list(apiPath);

    if (triggersResponse && triggersResponse.trigger) {
       triggersResponse.trigger.forEach(trigger => {
         triggersMap[trigger.triggerId] = trigger;
       });
       Logger.log(`Fetched ${Object.keys(triggersMap).length} triggers using TagManager service.`);
    } else {
       Logger.log('No triggers found in the workspace using TagManager service.');
    }
    
    // Add mapping for common built-in triggers with known IDs
    // These IDs can vary by GTM container, but some are consistent
    const builtInTriggers = {
      '2147479553': { name: 'All Pages', type: 'pageview' },
      '2147479554': { name: 'Window Loaded', type: 'windowLoaded' },
      '2147479572': { name: 'DOM Ready', type: 'domReady' },
      '2147479573': { name: 'History Change', type: 'historyChange' },
      // Add more built-in triggers as you identify them
    };
    
    // Add built-in triggers to our map
    Object.entries(builtInTriggers).forEach(([id, triggerInfo]) => {
      if (!triggersMap[id]) {
        triggersMap[id] = {
          triggerId: id,
          name: triggerInfo.name,
          type: triggerInfo.type
        };
      }
    });
    
  } catch (error) {
    Logger.log(`Error fetching triggers using TagManager service: ${error.toString()}`);
    docSheet.appendRow([`Error fetching triggers: ${error.toString()}`]);
    return; // Stop execution if triggers can't be fetched
  }
  
  // 2. Fetch Variables to get their notes
  try {
    const variablesResponse = TagManager.Accounts.Containers.Workspaces.Variables.list(apiPath);
    
    if (variablesResponse && variablesResponse.variable) {
      Logger.log(`Fetched ${variablesResponse.variable.length} variables from GTM.`);
      
      // Create a mapping of variable names to their notes
      variablesResponse.variable.forEach(variable => {
        // Add variable to the map - using both with and without {{ }} format
        const variableName = variable.name;
        const variableNotes = variable.notes || "";
        
        // Store the variable with its notes
        variableNotesMap[variableName] = variableNotes;
        
        // Also store with {{ }} format for direct matching
        variableNotesMap[`{{${variableName}}}`] = variableNotes;
      });
      
      Logger.log(`Created variable notes map with ${Object.keys(variableNotesMap).length} entries.`);
    } else {
      Logger.log('No variables found in the workspace.');
    }
  } catch (error) {
    Logger.log(`Error fetching variables: ${error.toString()}`);
    // Continue execution - variables are useful but not critical
  }

  try {
    // 3. Fetch Tags using the TagManager advanced service
    const tagsResponse = TagManager.Accounts.Containers.Workspaces.Tags.list(apiPath);

    if (tagsResponse && tagsResponse.tag) {
      const tags = tagsResponse.tag;

      if (!tags || tags.length === 0) {
        Logger.log('No tags found in the workspace.');
        docSheet.appendRow(['No tags found in the workspace.']);
        return;
      }

      Logger.log(`Found ${tags.length} tags.`);

      // Collect all data rows before writing to sheet
      const allDataRows = [];

      // Process each tag
      tags.forEach(tag => {
        Logger.log('Processing tag type: ' + tag.type + ' for tag: ' + tag.name);

        // Only process GA4 Event tags (type 'gaawe')
        if (tag.type === 'gaawe' && tag.parameter) {
          const tagName = tag.name;
          const tagNotes = tag.notes || ""; // Extract notes from the tag
          const firingTriggerIds = tag.firingTriggerId || [];
          
          // Extract event name directly
          let eventName = '';
          let parameterMap = new Map(); // Will hold parameter name -> value mappings

          // Find the event name parameter
          const eventNameParam = tag.parameter.find(param => 
            param.key === 'eventName' && param.type && param.type.toLowerCase() === 'template');
          
          if (eventNameParam) {
            eventName = eventNameParam.value;
            
            // MODIFIED: Check if this event exists and if the description has changed
            const description = tagNotes || "(Need to specify)";
            
            if (existingEvents.has(eventName)) {
              const existingDescription = existingEvents.get(eventName);
              // If GTM description is not empty and different from existing, mark for update
              if (description !== "(Need to specify)" && description !== existingDescription) {
                eventsToUpdate.set(eventName, description);
                Logger.log(`Event ${eventName} description will be updated from "${existingDescription}" to "${description}"`);
              }
            } else if (!newEvents.has(eventName)) {
              // Track as new event if not already tracked
              newEvents.set(eventName, description);
            }
          }

          // Find the eventSettingsTable parameter
          const settingsTableParam = tag.parameter.find(param => 
            param.key === 'eventSettingsTable' && param.type && param.type.toLowerCase() === 'list');
          
          // Process parameters if eventSettingsTable exists
          if (settingsTableParam && settingsTableParam.list) {
            settingsTableParam.list.forEach(item => {
              // Each item in the list is a parameter mapping
              if (item.type === 'map' && item.map) {
                let paramName = '';
                let paramValue = '';
                
                // Find parameter name and value in this map
                item.map.forEach(mapItem => {
                  if (mapItem.key === 'parameter' && mapItem.type && 
                      mapItem.type.toLowerCase() === 'template') {
                    paramName = mapItem.value;
                  }
                  if (mapItem.key === 'parameterValue' && mapItem.type && 
                      mapItem.type.toLowerCase() === 'template') {
                    paramValue = mapItem.value;
                    
                  // MODIFIED: Check if this is a variable parameter and handle updates
                  if (paramValue && paramValue.includes('{{') && paramValue.includes('}}')) {
                    // Get the notes for this parameter if available
                    let paramDescription = "(Need to specify)";
                    
                    // First try exact match with the variable notes map
                    if (variableNotesMap[paramValue]) {
                      paramDescription = variableNotesMap[paramValue];
                    } else {
                      // Try to extract variable name from the parameter value
                      const varNameMatch = paramValue.match(/\{\{([^\}]+)\}\}/);
                      if (varNameMatch && varNameMatch[1]) {
                        const variableName = varNameMatch[1].trim();

                        // Check if it's a built-in variable first
                      if (BUILTIN_VARIABLE_DESCRIPTIONS[variableName]) {
                        paramDescription = formatBuiltInDescription(BUILTIN_VARIABLE_DESCRIPTIONS[variableName]);
                      } else if (variableNotesMap[variableName]) {
                        // Otherwise check custom variable notes
                        paramDescription = variableNotesMap[variableName];
                      }
                      }
                    }

                      // Check if this parameter exists and if the description has changed
                      if (existingParameters.has(paramValue)) {
                        const existingDescription = existingParameters.get(paramValue);
                        // If GTM description is not empty and different from existing, mark for update
                        if (paramDescription !== "(Need to specify)" && paramDescription !== existingDescription) {
                          parametersToUpdate.set(paramValue, paramDescription);
                          Logger.log(`Parameter ${paramValue} description will be updated from "${existingDescription}" to "${paramDescription}"`);
                        }
                      } else if (!newParameters.has(paramValue)) {
                        // Track as new parameter if not already tracked
                        newParameters.set(paramValue, paramDescription);
                      }
                    }
                  }
                });
                
                // Add to our parameter map if we found both name and value
                if (paramName) {
                  parameterMap.set(paramName, paramValue || '');
                }
              }
            });
          }

          // Generate rows for this tag - one row per parameter
          if (parameterMap.size > 0) {
            // Convert Map to arrays for easier processing
            const paramNames = Array.from(parameterMap.keys());
            const paramValues = paramNames.map(name => parameterMap.get(name));
            
            // Combine all triggers into a single cell with newlines
            let combinedTriggerNames = '';
            let combinedTriggerTypes = '';
            
            if (firingTriggerIds.length === 0) {
              combinedTriggerNames = 'None';
              combinedTriggerTypes = 'N/A';
            } else {
              // Process all triggers
              firingTriggerIds.forEach((triggerId, index) => {
                const trigger = triggersMap[triggerId];
                const triggerName = trigger ? trigger.name : `Unknown Trigger ID: ${triggerId}`;
                const triggerType = trigger ? trigger.type : 'N/A';
                
                // Add newline between entries if not the first one
                if (index > 0) {
                  combinedTriggerNames += '\n';
                  combinedTriggerTypes += '\n';
                }
                
                combinedTriggerNames += triggerName;
                combinedTriggerTypes += triggerType;
              });
            }
            
            // MODIFIED: Get event description prioritizing updates
            let eventDescription = "(Need to specify)"; // Default
            
            if (eventsToUpdate.has(eventName)) {
              eventDescription = eventsToUpdate.get(eventName);
            } else if (existingEvents.has(eventName) && existingEvents.get(eventName)) {
              eventDescription = existingEvents.get(eventName);
            } else if (newEvents.has(eventName)) {
              eventDescription = newEvents.get(eventName);
            }
            
            // Process each parameter row
            for (let i = 0; i < paramNames.length; i++) {
              const paramName = paramNames[i];
              const paramValue = paramValues[i];
              
              // MODIFIED: Get parameter description prioritizing updates
              let paramDescription = "(Need to specify)"; // Default
              const paramKey = paramValue;
              
              if (parametersToUpdate.has(paramKey)) {
                paramDescription = parametersToUpdate.get(paramKey);
              } else if (existingParameters.has(paramKey) && existingParameters.get(paramKey)) {
                paramDescription = existingParameters.get(paramKey);
              } else if (newParameters.has(paramKey)) {
                paramDescription = newParameters.get(paramKey);
                } else if (paramValue && paramValue.includes('{{') && paramValue.includes('}}')) {
                  // Try to get description from variable notes
                  const varNameMatch = paramValue.match(/\{\{([^\}]+)\}\}/);
                  if (varNameMatch && varNameMatch[1]) {
                    const variableName = varNameMatch[1].trim();
                    
                    // Check if it's a built-in variable first
                    if (BUILTIN_VARIABLE_DESCRIPTIONS[variableName]) {
                      paramDescription = BUILTIN_VARIABLE_DESCRIPTIONS[variableName];
                    } else if (variableNotesMap[variableName]) {
                      // Otherwise check custom variable notes
                      paramDescription = variableNotesMap[variableName];
                    }
                  }
                }
              
              // Add row with all columns including descriptions
              allDataRows.push([
                tagName, 
                combinedTriggerNames, 
                combinedTriggerTypes, 
                eventName, 
                eventDescription,
                paramName, 
                paramValue,
                paramDescription
              ]);
            }
          } else {
            // No parameters found, just output basic tag info
            let combinedTriggerNames = '';
            let combinedTriggerTypes = '';
            
            if (firingTriggerIds.length === 0) {
              combinedTriggerNames = 'None';
              combinedTriggerTypes = 'N/A';
            } else {
              // Process all triggers
              firingTriggerIds.forEach((triggerId, index) => {
                const trigger = triggersMap[triggerId];
                const triggerName = trigger ? trigger.name : `Unknown Trigger ID: ${triggerId}`;
                const triggerType = trigger ? trigger.type : 'N/A';
                
                // Add newline between entries if not the first one
                if (index > 0) {
                  combinedTriggerNames += '\n';
                  combinedTriggerTypes += '\n';
                }
                
                combinedTriggerNames += triggerName;
                combinedTriggerTypes += triggerType;
              });
            }
            
            // MODIFIED: Get event description prioritizing updates
            let eventDescription = "(Need to specify)"; // Default
            
            if (eventsToUpdate.has(eventName)) {
              eventDescription = eventsToUpdate.get(eventName);
            } else if (existingEvents.has(eventName) && existingEvents.get(eventName)) {
              eventDescription = existingEvents.get(eventName);
            } else if (newEvents.has(eventName)) {
              eventDescription = newEvents.get(eventName);
            }
            
            allDataRows.push([
              tagName, 
              combinedTriggerNames, 
              combinedTriggerTypes, 
              eventName, 
              eventDescription,
              'N/A', 
              'N/A',
              ''
            ]);
          }
        }
      });

      // Write all data rows to the sheet
      if (allDataRows.length > 0) {
        docSheet.getRange(docSheet.getLastRow() + 1, 1, allDataRows.length, allDataRows[0].length)
             .setValues(allDataRows);
        Logger.log('Report generated successfully.');
        
        // Auto-size columns for better readability
        for (let i = 1; i <= formattedHeaders.length; i++) {
          docSheet.autoResizeColumn(i);
        }
      } else {
        Logger.log('No GA4 Event tags found with parameters.');
        docSheet.appendRow(['No GA4 Event tags found with parameters.']);
      }
      
      // Update the Parameters sheet with any new parameters found
      if (newParameters.size > 0 && parametersSheet) {
        Logger.log(`Found ${newParameters.size} new parameters to add to Parameters sheet.`);
        
        const newParamsToAdd = [];
        newParameters.forEach((description, paramValue) => {
          newParamsToAdd.push([paramValue, description]);
        });
        
        // Append new parameters to the Parameters sheet
        const lastRow = parametersSheet.getLastRow();
        parametersSheet.getRange(lastRow + 1, 1, newParamsToAdd.length, 2)
                       .setValues(newParamsToAdd);
        
        Logger.log(`Updated Parameters sheet with ${newParamsToAdd.length} new parameters.`);
      }
      
      // Update the Events sheet with any new events found
      if (newEvents.size > 0 && eventsSheet) {
        Logger.log(`Found ${newEvents.size} new events to add to Events sheet.`);
        
        const newEventsToAdd = [];
        newEvents.forEach((description, eventName) => {
          newEventsToAdd.push([eventName, description]);
        });
        
        // Append new events to the Events sheet
        const lastRow = eventsSheet.getLastRow();
        eventsSheet.getRange(lastRow + 1, 1, newEventsToAdd.length, 2)
                   .setValues(newEventsToAdd);
        
        Logger.log(`Updated Events sheet with ${newEventsToAdd.length} new events.`);
      }
      
      // NEW: Update existing parameters with new descriptions
      if (parametersToUpdate.size > 0 && parametersSheet) {
        Logger.log(`Found ${parametersToUpdate.size} parameters that need description updates.`);
        
        const parameterData = parametersSheet.getRange(2, 1, Math.max(1, parametersSheet.getLastRow() - 1), 2).getValues();
        
        // Loop through each parameter to update
        parametersToUpdate.forEach((newDescription, paramValue) => {
          // Find the row index for this parameter
          const rowIndex = parameterData.findIndex(row => row[0] === paramValue);
          
          if (rowIndex !== -1) {
            // Update the description in the Parameters sheet (add 2 because rowIndex is 0-based and we skip header)
            parametersSheet.getRange(rowIndex + 2, 2).setValue(newDescription);
            Logger.log(`Updated parameter "${paramValue}" description to "${newDescription}"`);
          }
        });
      }
      
      // NEW: Update existing events with new descriptions
      if (eventsToUpdate.size > 0 && eventsSheet) {
        Logger.log(`Found ${eventsToUpdate.size} events that need description updates.`);
        
        const eventData = eventsSheet.getRange(2, 1, Math.max(1, eventsSheet.getLastRow() - 1), 2).getValues();
        
        // Loop through each event to update
        eventsToUpdate.forEach((newDescription, eventName) => {
          // Find the row index for this event
          const rowIndex = eventData.findIndex(row => row[0] === eventName);
          
          if (rowIndex !== -1) {
            // Update the description in the Events sheet (add 2 because rowIndex is 0-based and we skip header)
            eventsSheet.getRange(rowIndex + 2, 2).setValue(newDescription);
            Logger.log(`Updated event "${eventName}" description to "${newDescription}"`);
          }
        });
      }
      
      // Show success message with update counts
      const updateMessage = `GTM documentation report has been generated successfully.
${newParameters.size} new parameters added.
${newEvents.size} new events added.
${parametersToUpdate.size} parameter descriptions updated.
${eventsToUpdate.size} event descriptions updated.`;
      
      SpreadsheetApp.getUi().alert('Success', updateMessage, SpreadsheetApp.getUi().ButtonSet.OK);
      
    } else {
      Logger.log(`Unexpected response when listing tags: ${JSON.stringify(tagsResponse)}`);
      docSheet.appendRow(['Unexpected response when listing tags']);
    }
  } catch (error) {
    Logger.log(`Error fetching tags using TagManager service: ${error.toString()}`);
    docSheet.appendRow([`Error fetching tags: ${error.toString()}`]);
    SpreadsheetApp.getUi().alert('Error', 'An error occurred while generating the report: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
