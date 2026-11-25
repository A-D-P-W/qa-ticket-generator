// YouTrackAPI.js - YouTrack REST calls: create an issue, link issues through /api/commands, read an issue's State.

const REQUIRED_FIELD_DEFAULTS = { "Platforms": "PC" };

function QAGEN_createYouTrackIssue(issueObject, schemaMap) {
  if (!schemaMap) throw new Error('A valid schemaMap was not provided. Aborting.');
  
  if (issueObject.description) {
    issueObject.description = issueObject.description.replaceAll('\\n', '\n');
  }

  let customFields = [
    { name: "Priority", "$type": "SingleEnumIssueCustomField", value: { name: issueObject.priority } },
    { name: "Type", "$type": "SingleEnumIssueCustomField", value: { name: issueObject.type || "Task" } }
  ];

  for (const fieldName in REQUIRED_FIELD_DEFAULTS) {
    if (issueObject[fieldName]) {
        customFields.push({ name: fieldName, "$type": "MultiEnumIssueCustomField", "value": [{ "name": issueObject[fieldName] }] });
        continue;
    }
    if (customFields.some(cf => cf.name === fieldName)) continue;
    
    const fieldRules = schemaMap.get(fieldName);
    if (fieldRules && fieldRules.values && fieldRules.values.length > 0) {
      let valueToSet = REQUIRED_FIELD_DEFAULTS[fieldName];
      if (!fieldRules.values.includes(valueToSet)) {
        valueToSet = fieldRules.values[0];
      }
      Logger.log(`------> INFO: Applying workflow-required default: ${fieldName} = ${valueToSet}`);
      customFields.push({ name: fieldName, "$type": "MultiEnumIssueCustomField", "value": [{ "name": valueToSet }] });
    }
  }

  const payload = { 
    "project": { "id": YOUTRACK_PROJECT_ID }, 
    "summary": issueObject.summary, 
    "description": issueObject.description, 
    "usesMarkdown": true, // RESTORED: This enables full markdown rendering
    "customFields": customFields 
  };

  const fieldsQuery = 'fields=id,idReadable';
  const endpoint = `${YOUTRACK_BASE_URL}/api/issues?${fieldsQuery}`;
  const headers = QAGEN_getYouTrackHeaders_();
  const options = { method: 'post', headers: headers, payload: JSON.stringify(payload), muteHttpExceptions: true };
  
  try {
    Logger.log(`--------> ACTION: Sending ticket creation request to YouTrack for job "${issueObject.summary}".`);
    const response = UrlFetchApp.fetch(endpoint, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200 || responseCode === 201) {
      const newIssue = JSON.parse(responseBody);
      Logger.log(`--------> OUTCOME: ✅ Successfully created YouTrack ticket ${newIssue.idReadable} (Permanent ID: ${newIssue.id}).`);
      return newIssue;
    } else {
      Logger.log(`--------> ❌ FAILED: YouTrack server rejected the request.`);
      Logger.log(`--------> Status Code: ${responseCode}`);
      Logger.log(`--------> Server Response: ${responseBody}`);
      return null;
    }
  } catch (e) {
    _logApiError('QAGEN_createYouTrackIssue', endpoint, options, e);
    return null;
  }
}

function QAGEN_updateYouTrackIssueDescription_(issueIdReadable, newDescription) {
    const endpoint = `${YOUTRACK_BASE_URL}/api/issues/${issueIdReadable}?fields=id`;
    const headers = QAGEN_getYouTrackHeaders_();
    const payload = { "description": newDescription };
    const options = {
        method: 'post',
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(endpoint, options);
    if (response.getResponseCode() !== 200) {
        console.error(`❌ Failed to update description for ${issueIdReadable}. Response: ${response.getContentText()}`);
    } else {
        console.log(`✅ Successfully updated description for ${issueIdReadable}.`);
    }
}

function QAGEN_linkYouTrackIssues_(sourceReadableId, targetReadableId, linkType) {
    const endpoint = `${YOUTRACK_BASE_URL}/api/commands`;
    const headers = QAGEN_getYouTrackHeaders_();
    // This command now correctly modifies the CHILD issue, telling it about the PARENT.
    const commandQuery = `add ${linkType} ${targetReadableId}`;
    const payload = {
        "query": commandQuery,
        "issues": [{ "idReadable": sourceReadableId }]
    };
    const options = { method: 'post', headers: headers, payload: JSON.stringify(payload), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(endpoint, options);
    if (response.getResponseCode() !== 200) {
        console.error(`❌ Failed to link ${targetReadableId} to ${sourceReadableId}. Command: "${commandQuery}". Response: ${response.getContentText()}`);
    } else {
        console.log(`✅ Successfully linked ${targetReadableId} to ${sourceReadableId} as "${linkType}".`);
    }
}

function QAGEN_getYouTrackHeaders_() {
    const authToken = PropertiesService.getScriptProperties().getProperty('YOUTRACK_TOKEN');
    if (!authToken) throw new Error('YouTrack token not found.');
    return { 'Authorization': 'Bearer ' + authToken, 'Accept': 'application/json', 'Content-Type': 'application/json' };
}

function QAGEN_setYouTrackToken() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Set YouTrack Token', 'Please enter your permanent token:', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK && response.getResponseText().trim() !== '') {
        PropertiesService.getScriptProperties().setProperty('YOUTRACK_TOKEN', response.getResponseText().trim());
        ui.alert('✅ YouTrack API token has been securely stored.');
    }
}

function QAGEN_getYouTrackIssueDetails_(issueId) {
    const fieldsQuery = 'fields=customFields(name,value(name))';
    const endpoint = `${YOUTRACK_BASE_URL}/api/issues/${issueId}?${fieldsQuery}`;
    const headers = QAGEN_getYouTrackHeaders_();
    const options = { method: 'get', headers: headers, muteHttpExceptions: true }; // Method is 'get'
    
    try {
        const response = UrlFetchApp.fetch(endpoint, options);
        if (response.getResponseCode() === 200) {
            const issueDetails = JSON.parse(response.getContentText());
            const stateField = issueDetails.customFields.find(field => field.name === 'State');
            return stateField ? stateField.value.name : 'State Not Found';
        } else {
            Logger.log(`❌ Failed to get details for issue ${issueId}. Response: ${response.getContentText()}`);
            return null;
        }
    } catch (e) {
        // UPDATED: Call the new, centralized error logger
        _logApiError('QAGEN_getYouTrackIssueDetails_', endpoint, options, e);
        return null;
    }
}
/**
 * A centralized logger for handling and reporting API errors.
 * It provides a detailed, multi-line report for easier debugging.
 * @param {string} functionName The name of the function where the error occurred.
 * @param {string} endpoint The API endpoint that was called.
 * @param {object} options The options object sent with the request.
 * @param {Error} error The caught error object.
 */
function _logApiError(functionName, endpoint, options, error) {
    Logger.log(`!!! CRITICAL API ERROR in function: ${functionName} !!!`);
    Logger.log('-------------------------------------------------------------');
    Logger.log(`  - Endpoint: ${options.method.toUpperCase()} ${endpoint}`);
    
    // Safely log the payload if it exists
    if (options.payload) {
        try {
            const payloadObject = JSON.parse(options.payload);
            Logger.log(`  - Payload Sent:\n${JSON.stringify(payloadObject, null, 2)}`);
        } catch (e) {
            Logger.log(`  - Payload Sent (un-parsable):\n${options.payload}`);
        }
    }
    
    Logger.log(`  - Error Received: ${error.message}`);
    Logger.log('-------------------------------------------------------------');
}