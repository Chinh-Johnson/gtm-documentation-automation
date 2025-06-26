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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
    
    // Prepare all initial data in arrays for batch operations
    const initialData = [
      ['Open Google Tag Manager, select the workspace you want, and find the IDs in the URL. For example, if the URL is https://tagmanager.google.com/#/container/accounts/12345/containers/45678/workspaces/7891011/ then:', '', ''],
      ['GTM Configuration', '', ''],
      ['Setting', 'Value', ''],
      ['GTM Account ID', '', ''],
      ['GTM Container ID', '', ''],
      ['GTM Workspace ID', '', ''],
      ['', '', ''],
      ['', '', ''],
      ['Report Guide', '', '']
    ];
    
    // Batch write all initial data
    configSheet.getRange(1, 1, initialData.length, 3).setValues(initialData);
    
    // Batch format operations
    configSheet.getRange('A1:C1').merge()
               .setBackground('#ffffcc')
               .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    
    // Add documentation table
    addReportDocumentationTableOptimized(configSheet);
    
    // Apply formatting in batch
    formatConfigSheetOptimized(configSheet);
    
  } else {
    // For existing sheets, check if documentation exists
    const data = configSheet.getDataRange().getValues();
    let hasDocumentation = false;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().includes('What to Expect When You Generate the Report')) {
        hasDocumentation = true;
        break;
      }
    }
    
    if (!hasDocumentation) {
      const currentLastRow = configSheet.getLastRow();
      configSheet.getRange(currentLastRow + 1, 1, 2, 3).setValues([
        ['', '', ''],
        ['Report Guide', '', '']
      ]);
      addReportDocumentationTableOptimized(configSheet);
    }
    
    // Reapply formatting
    formatConfigSheetOptimized(configSheet);
  }
  
  // Get current values in one batch read
  const configData = configSheet.getRange('A3:B6').getValues();
  const accountId = configData[1][1] || '';
  const containerId = configData[2][1] || '';
  const workspaceId = configData[3][1] || '';
  
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
          .instructions {
            background-color: #ffffcc;
            margin-bottom: 15px;
            padding: 10px;
            border-radius: 4px;
            font-size: 13px;
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
            display: block;
            margin: 5px 0;
            font-family: monospace;
            background: #f5f5f5;
            padding: 5px;
            border-radius: 3px;
          }
          .bold-id {
            font-weight: bold;
            color: #d73027;
          }
        </style>
      </head>
      <body>
        <h2>GTM Configuration</h2>
        <div class="instructions">
          <p>Open Google Tag Manager, select the workspace you want, and find the IDs in the URL. For example, if the URL is:</p>
          <div class="url-example">https://tagmanager.google.com/#/container/accounts/<span class="bold-id">12345</span>/containers/<span class="bold-id">45678</span>/workspaces/<span class="bold-id">7891011</span>/</div>
          <p>then: account ID is <span class="bold-id">12345</span>, container ID is <span class="bold-id">45678</span>, workspace ID is <span class="bold-id">7891011</span></p>
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
            
            if (!accountId || !containerId || !workspaceId) {
              document.getElementById('error').textContent = 'All fields are required.';
              return;
            }
            
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
  .setHeight(500)
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
 * Gets the GTM configuration from the Config sheet.
 * Optimized to use batch reads
 */
function getGtmConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  let config = {
    accountId: '',
    containerId: '',
    workspaceId: ''
  };
  
  if (configSheet) {
    // Apply text wrapping in one operation
    configSheet.getRange('C:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    
    // Get all values in one batch read
    const configData = configSheet.getRange('B4:B6').getValues();
    config.accountId = configData[0][0] || '';
    config.containerId = configData[1][0] || '';
    config.workspaceId = configData[2][0] || '';
  }
  
  return config;
}

/**
 * Saves GTM config using batch operations
 */
function saveGtmConfig(accountId, containerId, workspaceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  // Use batch write instead of individual setValue calls
  configSheet.getRange('B4:B6').setValues([
    [accountId],
    [containerId],
    [workspaceId]
  ]);
  
  return { success: true };
}

/**
 * Converts a header string to lowercase with underscores
 */
function formatHeaderName(header) {
  return header.toLowerCase().replace(/\s+/g, '_');
}

/**
 * Main function - optimized with batch operations
 */
function listGa4EventTagsAndParameters() {
  const config = getGtmConfig();
  
  if (!config.accountId || !config.containerId || !config.workspaceId) {
    SpreadsheetApp.getUi().alert(
      'GTM Configuration Missing',
      'Please configure your GTM Account ID, Container ID, and Workspace ID using the "Configure GTM Settings" menu option.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    showConfigurationDialog();
    return;
  }
  
  const accountId = config.accountId;
  const containerId = config.containerId;
  const workspaceId = config.workspaceId;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create sheets
  let docSheet = ss.getSheetByName("GTM documentation") || ss.insertSheet("GTM documentation");
  let parametersSheet = ss.getSheetByName("Parameters") || ss.insertSheet("Parameters");
  let eventsSheet = ss.getSheetByName("Events") || ss.insertSheet("Events");
  
  // Clear and set headers for all sheets in batch operations
  docSheet.clear();
  parametersSheet.clear();
  eventsSheet.clear();
  
  // Define headers
  const originalHeaders = [
    'Tag Name', 'Tag Status', 'Trigger', 'Trigger Type', 
    'Event Name', 'Event Description', 'Parameter Name', 
    'Parameter Value', 'Parameter Description'
  ];
  const formattedHeaders = originalHeaders.map(header => formatHeaderName(header));
  
  // Set all headers at once
  docSheet.getRange(1, 1, 1, formattedHeaders.length).setValues([formattedHeaders]);
  parametersSheet.getRange(1, 1, 1, 2).setValues([["Parameter Value", "Description"]]);
  eventsSheet.getRange(1, 1, 1, 2).setValues([["Event Name", "Description"]]);
  
  // Apply header formatting
  docSheet.getRange(1, 1, 1, formattedHeaders.length).setFontWeight('bold');
  docSheet.setFrozenRows(1);

  const apiPath = `accounts/${accountId}/containers/${containerId}/workspaces/${workspaceId}`;
  
  let triggersMap = {};
  let variableNotesMap = {};
  let allParameters = new Map();
  let allEvents = new Map();

  try {
    // Fetch Triggers
    const triggersResponse = TagManager.Accounts.Containers.Workspaces.Triggers.list(apiPath);
    if (triggersResponse && triggersResponse.trigger) {
      triggersResponse.trigger.forEach(trigger => {
        triggersMap[trigger.triggerId] = trigger;
      });
    }
    
    // Add built-in triggers
    const builtInTriggers = {
      '2147479553': { name: 'All Pages', type: 'pageview' },
      '2147479554': { name: 'Window Loaded', type: 'windowLoaded' },
      '2147479572': { name: 'DOM Ready', type: 'domReady' },
      '2147479573': { name: 'History Change', type: 'historyChange' },
    };
    
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
    Logger.log(`Error fetching triggers: ${error.toString()}`);
    docSheet.getRange(2, 1).setValue(`Error fetching triggers: ${error.toString()}`);
    return;
  }
  
  // Fetch Variables
  try {
    const variablesResponse = TagManager.Accounts.Containers.Workspaces.Variables.list(apiPath);
    if (variablesResponse && variablesResponse.variable) {
      variablesResponse.variable.forEach(variable => {
        const variableName = variable.name;
        const variableNotes = variable.notes || "";
        variableNotesMap[variableName] = variableNotes;
        variableNotesMap[`{{${variableName}}}`] = variableNotes;
      });
    }
  } catch (error) {
    Logger.log(`Error fetching variables: ${error.toString()}`);
  }

  try {
    // Fetch Tags
    const tagsResponse = TagManager.Accounts.Containers.Workspaces.Tags.list(apiPath);
    
    if (tagsResponse && tagsResponse.tag) {
      const tags = tagsResponse.tag;
      
      if (!tags || tags.length === 0) {
        docSheet.getRange(2, 1).setValue('No tags found in the workspace.');
        return;
      }
      
      // Collect ALL data rows before writing to sheet
      const allDataRows = [];
      
      // Process each tag
      tags.forEach(tag => {
        if (tag.type === 'gaawe' && tag.parameter) {
          const tagName = tag.name;
          const tagStatus = tag.paused ? 'Paused' : 'Live';
          const tagNotes = tag.notes || "";
          const firingTriggerIds = tag.firingTriggerId || [];
          
          let eventName = '';
          let parameterMap = new Map();
          
          // Find event name
          const eventNameParam = tag.parameter.find(param => 
            param.key === 'eventName' && param.type && param.type.toLowerCase() === 'template');
          
          if (eventNameParam) {
            eventName = eventNameParam.value;
            const eventDescription = tagNotes || "(Need to specify)";
            allEvents.set(eventName, eventDescription);
          }
          
          // Find parameters
          const settingsTableParam = tag.parameter.find(param => 
            param.key === 'eventSettingsTable' && param.type && param.type.toLowerCase() === 'list');
          
          if (settingsTableParam && settingsTableParam.list) {
            settingsTableParam.list.forEach(item => {
              if (item.type === 'map' && item.map) {
                let paramName = '';
                let paramValue = '';
                
                item.map.forEach(mapItem => {
                  if (mapItem.key === 'parameter' && mapItem.type && 
                      mapItem.type.toLowerCase() === 'template') {
                    paramName = mapItem.value;
                  }
                  if (mapItem.key === 'parameterValue' && mapItem.type && 
                      mapItem.type.toLowerCase() === 'template') {
                    paramValue = mapItem.value;
                    
                    let paramDescription = "(Need to specify)";
                    
                    if (paramValue && paramValue.includes('{{') && paramValue.includes('}}')) {
                      if (variableNotesMap[paramValue] && variableNotesMap[paramValue].trim()) {
                        paramDescription = variableNotesMap[paramValue];
                      } else {
                        const varNameMatch = paramValue.match(/\{\{([^\}]+)\}\}/);
                        if (varNameMatch && varNameMatch[1]) {
                          const variableName = varNameMatch[1].trim();
                          if (BUILTIN_VARIABLE_DESCRIPTIONS[variableName]) {
                            paramDescription = formatBuiltInDescription(BUILTIN_VARIABLE_DESCRIPTIONS[variableName]);
                          } else if (variableNotesMap[variableName] && variableNotesMap[variableName].trim()) {
                            paramDescription = variableNotesMap[variableName];
                          }
                        }
                      }
                    } else if (paramValue && paramValue.trim()) {
                      paramDescription = `Static: ${paramValue}`;
                    }
                    
                    allParameters.set(paramValue, paramDescription);
                  }
                });
                
                if (paramName) {
                  parameterMap.set(paramName, paramValue || '');
                }
              }
            });
          }
          
          // Generate rows
          if (parameterMap.size > 0) {
            const paramNames = Array.from(parameterMap.keys());
            const paramValues = paramNames.map(name => parameterMap.get(name));
            
            let combinedTriggerNames = '';
            let combinedTriggerTypes = '';
            
            if (firingTriggerIds.length === 0) {
              combinedTriggerNames = 'None';
              combinedTriggerTypes = 'N/A';
            } else {
              firingTriggerIds.forEach((triggerId, index) => {
                const trigger = triggersMap[triggerId];
                const triggerName = trigger ? trigger.name : `Unknown Trigger ID: ${triggerId}`;
                const triggerType = trigger ? trigger.type : 'N/A';
                
                if (index > 0) {
                  combinedTriggerNames += '\n';
                  combinedTriggerTypes += '\n';
                }
                
                combinedTriggerNames += triggerName;
                combinedTriggerTypes += triggerType;
              });
            }
            
            const eventDescription = allEvents.get(eventName) || "(Need to specify)";
            
            for (let i = 0; i < paramNames.length; i++) {
              const paramName = paramNames[i];
              const paramValue = paramValues[i];
              
              let paramDescription = allParameters.get(paramValue) || "(Need to specify)";
              
              if (paramDescription === "(Need to specify)" && paramValue) {
                if (paramValue.includes('{{') && paramValue.includes('}}')) {
                  const varNameMatch = paramValue.match(/\{\{([^\}]+)\}\}/);
                  if (varNameMatch && varNameMatch[1]) {
                    const variableName = varNameMatch[1].trim();
                    
                    if (BUILTIN_VARIABLE_DESCRIPTIONS[variableName]) {
                      paramDescription = formatBuiltInDescription(BUILTIN_VARIABLE_DESCRIPTIONS[variableName]);
                    } else if (variableNotesMap[variableName] && variableNotesMap[variableName].trim()) {
                      paramDescription = variableNotesMap[variableName];
                    }
                  }
                } else if (paramValue.trim()) {
                  paramDescription = `Static: ${paramValue}`;
                }
              }
              
              allDataRows.push([
                tagName, tagStatus, combinedTriggerNames, combinedTriggerTypes, 
                eventName, eventDescription, paramName, paramValue, paramDescription
              ]);
            }
          } else {
            // No parameters case
            let combinedTriggerNames = '';
            let combinedTriggerTypes = '';
            
            if (firingTriggerIds.length === 0) {
              combinedTriggerNames = 'None';
              combinedTriggerTypes = 'N/A';
            } else {
              firingTriggerIds.forEach((triggerId, index) => {
                const trigger = triggersMap[triggerId];
                const triggerName = trigger ? trigger.name : `Unknown Trigger ID: ${triggerId}`;
                const triggerType = trigger ? trigger.type : 'N/A';
                
                if (index > 0) {
                  combinedTriggerNames += '\n';
                  combinedTriggerTypes += '\n';
                }
                
                combinedTriggerNames += triggerName;
                combinedTriggerTypes += triggerType;
              });
            }
            
            const eventDescription = allEvents.get(eventName) || "(Need to specify)";
            
            allDataRows.push([
              tagName, tagStatus, combinedTriggerNames, combinedTriggerTypes, 
              eventName, eventDescription, 'N/A', 'N/A', ''
            ]);
          }
        }
      });
      
      // Write ALL data rows in one batch operation
      if (allDataRows.length > 0) {
        docSheet.getRange(2, 1, allDataRows.length, allDataRows[0].length).setValues(allDataRows);
        
        // Auto-resize all columns at once
        for (let i = 1; i <= formattedHeaders.length; i++) {
          docSheet.autoResizeColumn(i);
        }
      } else {
        docSheet.getRange(2, 1).setValue('No GA4 Event tags found with parameters.');
      }
      
      // Update Parameters sheet in batch
      if (allParameters.size > 0) {
        const parametersToAdd = [];
        allParameters.forEach((description, paramValue) => {
          parametersToAdd.push([paramValue, description]);
        });
        
        if (parametersToAdd.length > 0) {
          parametersSheet.getRange(2, 1, parametersToAdd.length, 2).setValues(parametersToAdd);
        }
      }
      
      // Update Events sheet in batch
      if (allEvents.size > 0) {
        const eventsToAdd = [];
        allEvents.forEach((description, eventName) => {
          eventsToAdd.push([eventName, description]);
        });
        
        if (eventsToAdd.length > 0) {
          eventsSheet.getRange(2, 1, eventsToAdd.length, 2).setValues(eventsToAdd);
        }
      }
      
      SpreadsheetApp.getUi().alert('Success', 'GTM documentation report has been generated successfully with the latest data from GTM.', SpreadsheetApp.getUi().ButtonSet.OK);
      
    } else {
      docSheet.getRange(2, 1).setValue('Unexpected response when listing tags');
    }
  } catch (error) {
    Logger.log(`Error fetching tags: ${error.toString()}`);
    docSheet.getRange(2, 1).setValue(`Error fetching tags: ${error.toString()}`);
    SpreadsheetApp.getUi().alert('Error', 'An error occurred while generating the report: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Optimized version of addReportDocumentationTable using batch operations
 */
function addReportDocumentationTableOptimized(configSheet) {
  const currentLastRow = configSheet.getLastRow();
  
  // Prepare all documentation data
  const allDocData = [
    ['', '', ''],
    ['What to Expect When You Generate the Report', '', ''],
    ['', '', ''],
    ['The "Generate GA4 Event Tag Report" button will create 3 sheets with the following information:', '', ''],
    ['', '', '']
  ];
  
  // Sheet 1 documentation
  const sheet1Data = [
    ['Sheet 1: GTM Documentation (Main Report)', '', ''],
    ['Column Name', 'What It Contains', 'Example'],
    ['tag_name', 'The name of your GA4 event tag in GTM', 'Purchase Complete'],
    ['tag_status', 'Whether the tag is active or paused', 'Live or Paused'],
    ['trigger', 'What triggers this tag to fire', 'Purchase Page View'],
    ['trigger_type', 'The type of trigger used', 'pageview, click, custom'],
    ['event_name', 'The GA4 event name being sent', 'purchase'],
    ['event_description', 'This describes the event being tracked, as specified in the Notes field of the tag within Google Tag Manager (GTM)', 'Tracks completed purchases'],
    ['parameter_name', 'Name of the event parameter', 'transaction_id'],
    ['parameter_value', 'The GTM variable used for this parameter', '{{Transaction ID}}'],
    ['parameter_description', 'If the parameter is a custom variable, its description comes from the Notes field in GTM. For static values, it shows "Static: ". Built-in variables use Googles default description', 'e.g For built-in variable Click Text: parameter_description: <<Built-in: Accesses the gtm.elementText key in the dataLayer, which is set by Click triggers.>> parameter_type: <<string>>'],
    ['', '', '']
  ];
  
  // Sheet 2 documentation
  const sheet2Data = [
    ['Sheet 2: Parameters', '', ''],
    ['Column Name', 'What It Contains', 'Example'],
    ['Parameter Value', 'The GTM variable name used', '{{Transaction ID}}'],
    ['Description', 'What this parameter represents', 'Unique identifier for the transaction'],
    ['', '', '']
  ];
  
  // Sheet 3 documentation
  const sheet3Data = [
    ['Sheet 3: Events', '', ''],
    ['Column Name', 'What It Contains', 'Example'],
    ['Event Name', 'The GA4 event name', 'purchase'],
    ['Description', 'What this event tracks', 'Tracks completed purchases'],
    ['', '', '']
  ];
  
  // Tip row
  const tipData = [
    ['💡 Tip: The report pulls fresh data from GTM each time you run it, so your documentation stays up-to-date!', '', '']
  ];
  
  // Combine all data
  const allData = [
    ...allDocData,
    ...sheet1Data,
    ...sheet2Data,
    ...sheet3Data,
    ...tipData
  ];
  
  // Write all data in one batch operation
  configSheet.getRange(currentLastRow + 1, 1, allData.length, 3).setValues(allData);
  
  // Apply formatting in batch
  const startRow = currentLastRow + 1;
  
  // Format main header
  configSheet.getRange(startRow + 1, 1, 1, 3).merge()
             .setFontSize(14)
             .setFontWeight('bold')
             .setBackground('#cfe2f3')
             .setHorizontalAlignment('center');
  
  // Format description
  configSheet.getRange(startRow + 3, 1, 1, 3).merge()
             .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Format sheet headers and table headers
  const sheetHeaderRows = [
    startRow + 5,  // Sheet 1 header
    startRow + 17, // Sheet 2 header  
    startRow + 22  // Sheet 3 header
  ];
  
  sheetHeaderRows.forEach(row => {
    configSheet.getRange(row, 1, 1, 3).setFontWeight('bold')
                                      .setBackground('#e1ecf7');
    configSheet.getRange(row + 1, 1, 1, 3).setFontWeight('bold')
                                           .setBackground('#cfe2f3');
  });
  
  // Format tip
  configSheet.getRange(startRow + allData.length - 1, 1, 1, 3).merge()
             .setBackground('#fff2cc')
             .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
             .setFontStyle('italic');
  
  // Apply text wrapping to all description columns
  configSheet.getRange(startRow, 2, allData.length, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

/**
 * Optimized version of formatConfigSheet
 */
function formatConfigSheetOptimized(configSheet) {
  try {
    // Set column widths individually
    configSheet.setColumnWidth(1, 180);
    configSheet.setColumnWidth(2, 180);
    configSheet.setColumnWidth(3, 350);
    
    // Get all data to find specific rows
    const data = configSheet.getDataRange().getValues();
    let docHeaderRow = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().includes('Report Guide')) {
        docHeaderRow = i + 1;
        break;
      }
    }
    
    // Apply all formatting in batch operations
    const formattingRanges = [
      // Instructions
      { range: 'A1', bg: '#ffffcc', wrap: true, vAlign: 'top' },
      // GTM Configuration header
      { range: 'A2:C2', merge: true, bg: '#b6d7ff', bold: true, fontSize: 12, hAlign: 'center' },
      // Column headers
      { range: 'A3:B3', bold: true, bg: '#d9e7ff' },
      // Input area
      { range: 'B4:B6', bg: '#ffd966', border: true }
    ];
    
    // Apply formatting for each range
    formattingRanges.forEach(format => {
      const range = configSheet.getRange(format.range);
      
      if (format.merge) range.merge();
      if (format.bg) range.setBackground(format.bg);
      if (format.bold) range.setFontWeight('bold');
      if (format.fontSize) range.setFontSize(format.fontSize);
      if (format.hAlign) range.setHorizontalAlignment(format.hAlign);
      if (format.vAlign) range.setVerticalAlignment(format.vAlign);
      if (format.wrap) range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      if (format.border) {
        range.setBorder(true, true, true, true, false, false, '#e69138', SpreadsheetApp.BorderStyle.SOLID);
      }
    });
    
    // Format Documentation header if found
    if (docHeaderRow > 0) {
      configSheet.getRange(`A${docHeaderRow}:C${docHeaderRow}`)
                 .merge()
                 .setBackground('#b6d7ff')
                 .setFontWeight('bold')
                 .setFontSize(12)
                 .setHorizontalAlignment('center');
    }
    
    // Apply wrap to column C
    configSheet.getRange('C:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    
  } catch (e) {
    console.log('Formatting error:', e.toString());
  }
}

/**
 * Helper function to find a row containing specific text
 */
function findRowWithText(sheet, searchText) {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().includes(searchText)) {
      return i + 1;
    }
  }
  return -1;
}
