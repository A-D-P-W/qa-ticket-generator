// TestCaseGenerator.js - turns one job row plus its blueprint(s) into ticket objects.

const MAX_CHECKLIST_ITEMS = 25;

/**
 * Builds the complete parameter map for a job. Diagnostics and generation both
 * call this one function, so what diagnostics validates is what generation fills in.
 * Precedence, lowest to highest:
 * 1. _Global_Parameters sheet (Key, Value).
 * 2. Every non-empty column of the job row (e.g. job['Task_Title_Override']).
 * 3. The job's 'Parameters' cell: "key:value" pairs separated by ';' (or ',' if no ';').
 * @param {object} job The job object from _generator_config.
 * @returns {Map<string, any>} The complete parameter map.
 */
function _buildParameterMap(job) {
  const paramMap = new Map();
  _getGlobalParameters().forEach((value, key) => paramMap.set(key, value));
  for (const key in job) {
    if (Object.prototype.hasOwnProperty.call(job, key) && job[key] !== '') {
      paramMap.set(key, job[key]);
    }
  }
  const paramString = job['Parameters'];
  if (paramString && typeof paramString === 'string' && paramString.trim() !== '') {
    const delimiter = paramString.includes(';') ? ';' : ',';
    paramString.split(delimiter).forEach(pair => {
      const parts = pair.split(/:(.*)/s);
      if (parts.length >= 2) {
        const key = parts[0].trim();
        const value = parts[1].trim();
        if (key) paramMap.set(key, value);
      }
    });
  }
  return paramMap;
}

/**
 * Replaces ${key} placeholders in text with values from the job's parameter map.
 * A key with no value becomes [UNDEFINED_PARAMETER:key], so a gap is visible in the ticket.
 * @param {string} text The text containing placeholders like ${key}.
 * @param {object} job The job object from _generator_config.
 * @returns {string} The text with placeholders replaced.
 */
function _applyParameters(text, job) {
  if (!text || typeof text !== 'string' || !job) return text;
  const paramMap = _buildParameterMap(job);
  return text.replace(/\$\{(.*?)\}/g, (match, key) => {
    return paramMap.has(key) ? paramMap.get(key) : `[UNDEFINED_PARAMETER:${key}]`;
  });
}


function QAGEN_generateTestCasesFromLibrary(job) {
  const masterDb = SpreadsheetApp.openById(MASTER_DB_ID);
  const librarySheet = masterDb.getSheetByName('_Master_Test_Library');
  if (!librarySheet) throw new Error('The "_Master_Test_Library" sheet was not found.');
  
  const allLibraryData = librarySheet.getDataRange().getValues();
  const headers = allLibraryData.shift();
  const headerMap = headers.reduce((acc, h, i) => (acc[h] = i, acc), {});
  
  // --- UPDATED LOGIC ---
  // 'Exploratory Suite' is now treated as a 'Single Task' because it has the same structure.
  if (job['Test_Type'] === 'Single Task' || job['Test_Type'] === 'Exploratory Suite') {
      return _generateSingleTaskTicket(allLibraryData, headerMap, job);
  }
  if (job['Test_Type'] === 'Checklist (Manual)') {
      return _generateChecklistManualTickets(allLibraryData, headerMap, job);
  }
  if (job['Test_Type'] === 'Checklist (from Sheet)') {
      return _generateChecklistFromSheetTickets(librarySheet, headerMap, job);
  }
  
  Logger.log(`--> WARNING: Unknown Test_Type "${job['Test_Type']}" for JobID '${job['Job_ID']}'. No ticket generated.`);
  return [];
}

function _generateSingleTaskTicket(libraryData, headerMap, job) {
  const row = libraryData.find(r => r[headerMap['Test_Type']] === job['Test_Type'] && r[headerMap['Category']] === job['Category'] && r[headerMap['Sub_Category']] === job['Sub_Category']);
  if (!row) {
      Logger.log(`------> ❌ WARNING: Could not find a matching blueprint for job: ${job['Category']} - ${job['Sub_Category']}`);
      return [];
  }

  // --- CORRECTED LOGIC ---
  // Simply get the raw text from the blueprint columns.
  const description = row[headerMap['Description']] || '';
  const steps = row[headerMap['Steps']] || '';
  const expectedResult = row[headerMap['Expected_Result']] || '';

  // Concatenate them in the correct order. The blueprint now controls all formatting.
  const descTemplate = `${description}\n\n---\n\n${steps}\n\n---\n\n${expectedResult}`;

  const finalSummary = _applyParameters(job['Task_Title_Override'] || `[QA Task] - ${job['Category']}: ${job['Sub_Category']}`, job);
  const finalDescription = _applyParameters(descTemplate, job);

  const testCase = { 
    isSuite: false, 
    summary: finalSummary, 
    description: finalDescription,
    priority: job['Priority'], 
    type: 'Task' 
  };
  Logger.log(`------> ✅ Generated 1 "${job['Test_Type']}" test case.`);
  return [testCase];
}


function _generateChecklistManualTickets(libraryData, headerMap, job) {
  const items = libraryData.filter(r => 
    r[headerMap['Test_Type']] === job['Test_Type'] && 
    r[headerMap['Category']] === job['Category'] && 
    r[headerMap['Sub_Category']] === job['Sub_Category']
  );

  if (items.length === 0) {
      Logger.log(`------> ❌ WARNING: Could not find any 'Checklist (Manual)' items for job: ${job['Category']} - ${job['Sub_Category']}`);
      return [];
  }
  
  Logger.log(`------> Found ${items.length} checklist item(s) for suite "${job['Sub_Category']}".`);
  const numTickets = Math.ceil(items.length / MAX_CHECKLIST_ITEMS);

  if (numTickets === 1) {
    Logger.log('--------> Single-child suite detected. Generating a single, combined ticket.');
    const checklist = items.map(item => `- [ ] ${item[headerMap['Description']]}`).join('\n');
    const descriptionTemplate = `**Objective:** Complete this checklist.\n\n---\n\n### Checklist\n${checklist}`;
    const singleTicket = {
      isSuite: false,
      summary: _applyParameters(job['Task_Title_Override'] || `[QA Checklist] - ${job['Category']}: ${job['Sub_Category']}`, job),
      description: _applyParameters(descriptionTemplate, job),
      priority: job['Priority'], type: 'Task'
    };
    return [singleTicket];
  }

  const parent = {
    isSuite: true,
    summary: _applyParameters(job['Task_Title_Override'] || `[QA Checklist Meta] - ${job['Category']}: ${job['Sub_Category']}`, job),
    description: _applyParameters(`Parent meta-task for the "${job['Sub_Category']}" test suite.`, job),
    priority: job['Priority'], type: 'Task', subTasks: []
  };

  for (let i = 0; i < numTickets; i++) {
    const chunk = items.slice(i * MAX_CHECKLIST_ITEMS, (i + 1) * MAX_CHECKLIST_ITEMS);
    const checklist = chunk.map(item => `- [ ] ${item[headerMap['Description']]}`).join('\n');
    const part = ` (Part ${i + 1} of ${numTickets})`;
    const subTaskDescriptionTemplate = `**Objective:** Complete this checklist.\n\n---\n\n### Checklist\n${checklist}`;

    parent.subTasks.push({
      summary: _applyParameters((job['Task_Title_Override'] || `[QA Checklist] - ${job['Category']}: ${job['Sub_Category']}`) + part, job),
      description: _applyParameters(subTaskDescriptionTemplate, job),
      priority: job['Priority'], type: 'Task'
    });
  }
  Logger.log(`------> ✅ Generated 1 Meta task with ${parent.subTasks.length} sub-task(s).`);
  return [parent];
}


function _generateChecklistFromSheetTickets(librarySheet, headerMap, job) {
  Logger.log(`------> Starting 'Checklist (from Sheet)' generation for: ${job['Category']} - ${job['Sub_Category']}`);
  const libraryData = librarySheet.getDataRange().getValues();

  const ruleRow = libraryData.find(r => 
    r[headerMap['Test_Type']] === job['Test_Type'] && 
    r[headerMap['Category']] === job['Category'] && 
    r[headerMap['Sub_Category']] === job['Sub_Category']
  );

  if (!ruleRow) {
    Logger.log(`--------> ❌ ERROR: Could not find a 'Checklist (from Sheet)' rule for job: ${job['Category']} - ${job['Sub_Category']}`);
    return [];
  }

  const sourceSheetName = job['Data_Source_Sheet'];
  const sourceColumnName = job['Data_Source_Column'];
  if (!sourceSheetName || !sourceColumnName) {
    Logger.log(`--------> ❌ ERROR: JobID '${job['Job_ID']}' is missing Data_Source_Sheet or Data_Source_Column info in _generator_config.`);
    return [];
  }

  const masterDb = SpreadsheetApp.openById(MASTER_DB_ID);
  const sourceSheet = masterDb.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    Logger.log(`--------> ❌ ERROR: Data_Source_Sheet "${sourceSheetName}" was not found.`);
    return [];
  }

  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData.shift();
  const sourceColIndex = sourceHeaders.indexOf(sourceColumnName);
  if (sourceColIndex === -1) {
    Logger.log(`--------> ❌ ERROR: Data_Source_Column "${sourceColumnName}" not found in "${sourceSheetName}".`);
    return [];
  }

  const items = sourceData.map(row => row[sourceColIndex]).filter(item => item && item.toString().trim() !== '');
  if (items.length === 0) {
    Logger.log(`--------> ⚠️ WARNING: Found 0 items in "${sourceSheetName}" under column "${sourceColumnName}".`);
    return [];
  }
  
  Logger.log(`--------> Found ${items.length} dynamic item(s) to test from sheet "${sourceSheetName}".`);
  const genericInstructions = `\n\n---\n\n### General Instructions\n**Steps:**\n${ruleRow[headerMap['Steps']]}\n\n**Expected Result for all items:**\n${ruleRow[headerMap['Expected_Result']]}`;
  const numTickets = Math.ceil(items.length / MAX_CHECKLIST_ITEMS);

  if (numTickets === 1) {
    const checklist = items.map(item => `- [ ] ${item}`).join('\n');
    const descriptionTemplate = `**Objective:** Complete this checklist.${genericInstructions}\n\n---\n\n### Checklist\n${checklist}`;
    const singleTicket = {
      isSuite: false,
      summary: _applyParameters(job['Task_Title_Override'] || `[QA Checklist] - ${job['Category']}: ${job['Sub_Category']}`, job),
      description: _applyParameters(descriptionTemplate, job),
      priority: job['Priority'], type: 'Task'
    };
    return [singleTicket];
  }

  const parent = {
    isSuite: true,
    summary: _applyParameters(job['Task_Title_Override'] || `[QA Checklist Meta] - ${job['Category']}: ${job['Sub_Category']}`, job),
    description: _applyParameters(`Parent meta-task for the "${job['Sub_Category']}" test suite.\n\n**High-Level Objective:**\n${ruleRow[headerMap['Description']]}`, job),
    priority: job['Priority'], type: 'Task', subTasks: []
  };

  for (let i = 0; i < numTickets; i++) {
    const chunk = items.slice(i * MAX_CHECKLIST_ITEMS, (i + 1) * MAX_CHECKLIST_ITEMS);
    const checklist = chunk.map(item => `- [ ] ${item}`).join('\n');
    const part = ` (Part ${i + 1} of ${numTickets})`;
    const subTaskDescriptionTemplate = `**Objective:** Complete this checklist.${genericInstructions}\n\n---\n\n### Checklist (Part ${i + 1})\n${checklist}`;

    parent.subTasks.push({
      summary: _applyParameters((job['Task_Title_Override'] || `[QA Checklist] - ${job['Category']}: ${job['Sub_Category']}`) + part, job),
      description: _applyParameters(subTaskDescriptionTemplate, job),
      priority: job['Priority'], type: 'Task'
    });
  }
  return [parent];
}

/**
 * Reads the _Global_Parameters sheet and returns a map of key-value pairs.
 * Caches the result in memory for the duration of the script execution to avoid redundant reads.
 */
let _globalParamsCache = null;
function _getGlobalParameters() {
    if (_globalParamsCache) {
        return _globalParamsCache;
    }
    try {
        const masterDb = SpreadsheetApp.openById(MASTER_DB_ID);
        const sheet = masterDb.getSheetByName('_Global_Parameters');
        if (!sheet) {
            Logger.log("--> WARNING: '_Global_Parameters' sheet not found. No global parameters will be applied.");
            _globalParamsCache = new Map();
            return _globalParamsCache;
        }
        const values = sheet.getDataRange().getValues();
        values.shift(); // Remove header row
        _globalParamsCache = values.reduce((map, row) => {
            if (row[0]) map.set(row[0], row[1]);
            return map;
        }, new Map());
        return _globalParamsCache;
    } catch (e) {
        Logger.log(`--> ERROR reading global parameters: ${e.message}`);
        _globalParamsCache = new Map();
        return _globalParamsCache;
    }
}
