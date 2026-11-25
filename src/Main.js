// Main.js - entry points: diagnostics, manual generation, campaigns, scheduled triggers,
// status sync and data validation. _processJobs is the shared engine behind the live runs.

/**
 * Read-only pre-flight check. Reads the sheets, never calls YouTrack, never writes a cell.
 * For every job marked TRUE or TEST it checks that a blueprint matches the job's
 * (Test_Type, Category, Sub_Category) key and that every ${placeholder} in that blueprint
 * has a value. For TEST jobs it also logs a preview of the ticket it would create.
 */
function QAGEN_runDiagnostics() {
    Logger.log('--- STARTING SYSTEM DIAGNOSTICS ---');
    let issuesFound = 0;
    const errorMessages = [];

    const ss = SpreadsheetApp.openById(MASTER_DB_ID);
    const configSheet = ss.getSheetByName('_generator_config');
    const librarySheet = ss.getSheetByName('_Master_Test_Library');
    const allConfigData = configSheet.getDataRange().getValues();
    const configHeaders = allConfigData.shift();
    const allLibraryData = librarySheet.getDataRange().getValues();
    const libraryHeaders = allLibraryData.shift();
    const libraryHeaderMap = libraryHeaders.reduce((acc, h, i) => (acc[h] = i, acc), {});

    const jobsToProcess = allConfigData.map((row, index) => {
        const job = {};
        configHeaders.forEach((header, i) => job[header] = row[i]);
        job.rowIndex = index + 2;
        return job;
    }).filter(job => {
        const generateValue = String(job['Generate?']).toUpperCase();
        return generateValue === 'TRUE' || generateValue === 'TEST';
    });

    if (jobsToProcess.length === 0) {
        Logger.log('No active jobs found to diagnose.');
        Logger.log('--- DIAGNOSTICS COMPLETE: Found 0 issue(s). ---');
        return;
    }
    Logger.log(`Found ${jobsToProcess.length} active job(s) to scan for issues...`);

    jobsToProcess.forEach(job => {
        let jobHasIssues = false;
        Logger.log(`--> Evaluating Row ${job.rowIndex} (JobID '${job.Job_ID}'):`);
        const logError = (message) => {
            const fullErrorMessage = `(Row ${job.rowIndex} - JobID '${job.Job_ID}'): ${message}`;
            Logger.log(`------> ❌ FAILED: ${message}`);
            errorMessages.push(`❌ ${fullErrorMessage}`);
            issuesFound++;
            jobHasIssues = true;
        };

        const ruleKey = `Type="${job.Test_Type}", Category="${job.Category}", Sub_Category="${job.Sub_Category}"`;
        const ruleRow = allLibraryData.find(r => r[libraryHeaderMap['Test_Type']] === job.Test_Type && r[libraryHeaderMap['Category']] === job.Category && r[libraryHeaderMap['Sub_Category']] === job.Sub_Category);

        if (!ruleRow) {
            logError(`A matching rule was not found for the combination: ${ruleKey}.`);
        } else {
            Logger.log(`------> ✅ PASSED: Found matching blueprint in library.`);
            const blueprintText = (ruleRow[libraryHeaderMap['Description']] || '') + (ruleRow[libraryHeaderMap['Steps']] || '') + (ruleRow[libraryHeaderMap['Expected_Result']] || '');
            const requiredParams = new Set((blueprintText.match(/\$\{(.*?)\}/g) || []).map(p => p.slice(2, -1)));

            if (requiredParams.size > 0) {
                // --- UNIFIED LOGIC ---
                // Use the same authoritative map builder as the execution engine.
                const fullParamMap = _buildParameterMap(job);
                let missingParams = [];
                requiredParams.forEach(reqParam => {
                    if (!fullParamMap.has(reqParam)) {
                        missingParams.push(reqParam);
                    }
                });

                if (missingParams.length > 0) {
                    logError(`Missing required parameter(s). The blueprint requires: [${missingParams.join(', ')}].`);
                }
            }
        }

        if (!jobHasIssues) {
            Logger.log(`--> ✅ PASSED: Job is configured correctly.`);
            if (String(job['Generate?']).toUpperCase() === 'TEST') {
                Logger.log('----> ACTION: "TEST" mode detected. Generating Mock Ticket Preview...');
                const generatedTests = QAGEN_generateTestCasesFromLibrary(job);
                if (generatedTests.length > 0) {
                    const preview = generatedTests[0];
                    Logger.log(`--------> MOCK PREVIEW for Job '${job.Job_ID}'`);
                    Logger.log(`-----------> Summary: ${preview.summary}`);
                    Logger.log(`-----------> Description:\n${preview.description}`);
                    if (preview.isSuite) {
                        Logger.log(`-----------> Contains ${preview.subTasks.length} sub-task(s).`);
                    }
                } else {
                    Logger.log(`--------> WARNING: Mock generation produced no output for Job '${job.Job_ID}'.`);
                }
            }
        }
    });

    Logger.log(`\n--- DIAGNOSTICS COMPLETE: Found ${issuesFound} issue(s). ---`);
    if (errorMessages.length > 0) {
        Logger.log('--- ERROR SUMMARY ---');
        errorMessages.forEach(msg => Logger.log(msg));
        Logger.log('--- END OF SUMMARY ---');
    }
}

function QAGEN_runCampaign() {
  Logger.log('--- UI ACTION: "Run Campaign..." Triggered ---');
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  const configSheet = ss.getSheetByName('_generator_config');
  if (!configSheet) {
    SpreadsheetApp.getUi().alert('Error: "_generator_config" sheet not found.');
    Logger.log('!!! CRITICAL ERROR: "_generator_config" sheet not found.');
    return;
  }

  const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
  const campaignColIndex = headers.indexOf('Campaign');

  if (campaignColIndex === -1) {
    SpreadsheetApp.getUi().alert('Error: "Campaign" column not found in _generator_config.');
    Logger.log('!!! CRITICAL ERROR: "Campaign" column not found.');
    return;
  }

  const campaignData = configSheet.getRange(2, campaignColIndex + 1, configSheet.getLastRow() - 1, 1).getValues();
  
  const allCampaigns = campaignData.flat()
    .filter(c => c && typeof c === 'string' && c.trim() !== '')
    .flatMap(c => c.split(',').map(name => name.trim()))
    .filter(name => name);

  const uniqueCampaigns = [...new Set(allCampaigns)].sort();
  Logger.log(`Found ${uniqueCampaigns.length} unique campaigns to display.`);

  if (uniqueCampaigns.length === 0) {
    SpreadsheetApp.getUi().alert('No campaigns found in the "Campaign" column.');
    Logger.log('No campaigns found. Halting execution.');
    return;
  }

  Logger.log('--> ACTION: Displaying campaign selection dialog to user.');
  const htmlTemplate = HtmlService.createTemplateFromFile('CampaignDialog');
  htmlTemplate.campaigns = uniqueCampaigns;
  const html = htmlTemplate.evaluate().setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Run a Test Campaign');
}

function QAGEN_executeCampaign(campaignName) {
  Logger.log(`--- STARTING CAMPAIGN RUN: "${campaignName}" ---`);
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  const configSheet = ss.getSheetByName('_generator_config');
  if (!configSheet) {
      Logger.log('!!! CRITICAL ERROR: "_generator_config" sheet not found.');
      throw new Error('Critical Error: "_generator_config" sheet not found.');
  }

  const headerRange = configSheet.getRange(1, 1, 1, configSheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  const dataRange = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, headers.length);
  const allData = dataRange.getValues();

  const campaignColIndex = headers.indexOf('Campaign');
  const permanentIdColIndex = headers.indexOf('YouTrack_Permanent_ID');
  
  Logger.log(`Found ${allData.length} total jobs. Filtering for campaign "${campaignName}"...`);

  const jobsToProcess = allData.map((row, index) => ({ row, originalIndex: index }))
    .filter((item, index) => {
      const rowIndexForLog = index + 2;
      const campaignCell = item.row[campaignColIndex];
      const permanentIdValue = item.row[permanentIdColIndex];
      Logger.log(`--> Evaluating Row ${rowIndexForLog}:`);
      Logger.log(`----> CONTEXT: Campaign Cell is "${campaignCell}". Permanent_ID is "${permanentIdValue}".`);

      const ticketExists = permanentIdValue !== '';
      if(ticketExists){
        Logger.log('------> DECISION: Skipping because ticket already exists.');
        return false;
      }

      let includesCampaign = false;
      if (campaignCell && typeof campaignCell === 'string') {
        const campaignsInJob = campaignCell.split(',').map(c => c.trim());
        includesCampaign = campaignsInJob.includes(campaignName);
      }
      Logger.log(`------> DECISION: Does job belong to campaign "${campaignName}"? ${includesCampaign}`);
      
      const finalDecision = !ticketExists && includesCampaign;
      Logger.log(`------> OUTCOME for Row ${rowIndexForLog}: Keep this job for processing? ${finalDecision}`);
      return finalDecision;
    });

  if (jobsToProcess.length > 0) {
     Logger.log(`FILTERING COMPLETE: Found ${jobsToProcess.length} job(s) to process for this campaign.`);
    _processJobs(jobsToProcess, allData, dataRange, headers);
  } else {
    Logger.log(`FILTERING COMPLETE: No new jobs to process for campaign "${campaignName}".`);
  }
   Logger.log(`--- CAMPAIGN RUN COMPLETE: "${campaignName}" ---`);
}

function QAGEN_runTestGeneration() {
    Logger.log('--- STARTING MANUAL TEST GENERATION ---');
    const startTime = new Date();

    const ss = SpreadsheetApp.openById(MASTER_DB_ID);
    const configSheet = ss.getSheetByName('_generator_config');
    if (!configSheet) {
        Logger.log('!!! CRITICAL ERROR: "_generator_config" sheet not found.');
        return;
    }

    const headerRange = configSheet.getRange(1, 1, 1, configSheet.getLastColumn());
    const headers = headerRange.getValues()[0]; 
    const dataRange = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, headers.length);
    const allData = dataRange.getValues();
    
    Logger.log(`Found ${allData.length} total jobs in the sheet. Beginning filtering process.`);

    const permanentIdColIndex = headers.indexOf('YouTrack_Permanent_ID');
    const subCategoryColIndex = headers.indexOf('Sub_Category');
    const generateColIndex = headers.indexOf('Generate?');

    Logger.log(`--> Column Index Mapping: Generate?=${generateColIndex}, Permanent_ID=${permanentIdColIndex}, Sub_Category=${subCategoryColIndex}`);

    const jobsToProcess = allData.map((row, index) => ({ row, originalIndex: index }))
        .filter((item, index) => {
            const rowIndexForLog = index + 2;
            Logger.log(`--> Evaluating Row ${rowIndexForLog}:`);

            const generateValue = item.row[generateColIndex];
            const permanentIdValue = item.row[permanentIdColIndex];
            const subCategoryValue = item.row[subCategoryColIndex];

            Logger.log(`----> CONTEXT: Generate? is "${generateValue}" (Type: ${typeof generateValue}). Permanent_ID is "${permanentIdValue}" (Type: ${typeof permanentIdValue}).`);

            const shouldRun = (generateValue === true || String(generateValue).toUpperCase() === 'TRUE');
            Logger.log(`------> DECISION 1 (shouldRun): Is "Generate?" checkbox TRUE? Result: ${shouldRun}`);

            const ticketExists = permanentIdValue !== '';
            Logger.log(`------> DECISION 2 (ticketExists): Is "Permanent_ID" NOT empty? Result: ${ticketExists}`);
            
            if (shouldRun && ticketExists) {
                Logger.log(`------> ACTION: SKIPPING job '${subCategoryValue}' on row ${rowIndexForLog} because a ticket already exists.`);
            }

            const finalDecision = shouldRun && !ticketExists;
            Logger.log(`------> OUTCOME for Row ${rowIndexForLog}: Keep this job for processing? ${finalDecision}`);
            
            return finalDecision;
        });

    if (jobsToProcess.length > 0) {
      Logger.log(`FILTERING COMPLETE: Found ${jobsToProcess.length} new job(s) to process.`);
      _processJobs(jobsToProcess, allData, dataRange, headers);
    } else {
      Logger.log('FILTERING COMPLETE: No new jobs to process.');
    }

    const endTime = new Date();
    Logger.log(`--- MANUAL TEST GENERATION COMPLETE in ${(endTime - startTime) / 1000} seconds ---`);
}

/**
 * The core processing engine. Iterates through a pre-filtered list of jobs,
 * validates them, generates test cases, creates YouTrack tickets, and writes back the results.
 * @param {Array<object>} jobsToProcess An array of job items to be processed. Each item contains the row data and original index.
 * @param {Array<Array<any>>} allData A 2D array of all data from the config sheet, used for dependency lookups.
 * @param {GoogleAppsScript.Spreadsheet.Range} dataRange The full data range of the config sheet.
 * @param {Array<string>} headers An array of header names from the config sheet.
 * @param {boolean} [shouldUpdateLastRun=false] If true, the 'Last_Run' column will be updated for successfully processed jobs.
 */
function _processJobs(jobsToProcess, allData, dataRange, headers, shouldUpdateLastRun = false) {
    Logger.log(`--- Entering Core Processing Engine for ${jobsToProcess.length} job(s) ---`);
    const schemaMap = QAGEN_updateYouTrackSchemaCache();
    if (!schemaMap) {
        Logger.log('!!! CRITICAL ERROR: YouTrack schema could not be retrieved. Aborting processing engine.');
        throw new Error("Aborting: YouTrack schema could not be retrieved.");
    }

    const configSheet = dataRange.getSheet();
    const masterDb = SpreadsheetApp.openById(MASTER_DB_ID);
    const librarySheet = masterDb.getSheetByName('_Master_Test_Library');
    const libraryData = librarySheet.getDataRange().getValues();
    const libraryHeaders = libraryData.shift();
    const libraryHeaderMap = libraryHeaders.reduce((acc, h, i) => (acc[h] = i, acc), {});

    // Get column indices once for efficiency
    const jobIdColIndex = headers.indexOf('Job_ID');
    const statusColIndex = headers.indexOf('YouTrack_Status');
    const idReadableColIndex = headers.indexOf('YouTrack_Ticket_ID');
    const permanentIdColIndex = headers.indexOf('YouTrack_Permanent_ID');
    const urlColIndex = headers.indexOf('YouTrack_Ticket_URL');
    const lastRunColIndex = headers.indexOf('Last_Run');

    const allJobsMap = new Map(allData.map((row, index) => [row[jobIdColIndex], { row, originalIndex: index }]));

    jobsToProcess.forEach(item => {
        const job = headers.reduce((obj, header, index) => { obj[header] = item.row[index]; return obj; }, {});
        const jobID = job['Job_ID'];
        const rowIndex = item.originalIndex + 2;
        Logger.log(`--> Processing JobID '${jobID}' on row ${rowIndex}`);

        let dependenciesAreMet = true;
        const dependsOnString = job['Depends_On'];
        if (dependsOnString && dependsOnString.trim() !== '') {
            const dependencies = dependsOnString.split(',').map(d => d.split('|')[0].trim());
            Logger.log(`------> CONTEXT: Checking ${dependencies.length} dependenc(ies): [${dependencies.join(', ')}]`);
            for (const depId of dependencies) {
                const depJobItem = allJobsMap.get(depId);
                if (depJobItem) {
                    const depStatus = depJobItem.row[statusColIndex];
                    if (depStatus !== 'Verified' && depStatus !== 'Complete' && depStatus !== 'Fixed') {
                        dependenciesAreMet = false;
                        Logger.log(`--------> ❌ FAILED: Dependency '${depId}' has status "${depStatus}".`);
                        break;
                    }
                    Logger.log(`--------> ✅ PASSED: Dependency '${depId}' has status "${depStatus}".`);
                } else {
                    dependenciesAreMet = false;
                    Logger.log(`--------> ❌ FAILED: Could not find dependency JobID '${depId}' in the sheet.`);
                    break;
                }
            }
        }

        if (dependenciesAreMet) {
            Logger.log('------> ✅ All dependencies met. Proceeding with Parameter Safety Gate.');

            let paramsAreValid = true;
            const ruleRow = libraryData.find(r => r[libraryHeaderMap['Test_Type']] === job.Test_Type && r[libraryHeaderMap['Category']] === job.Category && r[libraryHeaderMap['Sub_Category']] === job.Sub_Category);
            
            if (ruleRow) {
                const blueprintText = (ruleRow[libraryHeaderMap['Description']] || '') + (ruleRow[libraryHeaderMap['Steps']] || '') + (ruleRow[libraryHeaderMap['Expected_Result']] || '');
                const requiredParams = new Set((blueprintText.match(/\$\{(.*?)\}/g) || []).map(p => p.slice(2, -1)));
                
                if (requiredParams.size > 0) {
                    const fullParamMap = _buildParameterMap(job);
                    requiredParams.forEach(reqParam => {
                        if (!fullParamMap.has(reqParam)) {
                            Logger.log(`------> ❌ FAILED Parameter Safety Gate. Missing required parameter: [${reqParam}]. Halting generation for Job_ID: [${jobID}].`);
                            paramsAreValid = false;
                        }
                    });
                }
            } else {
                Logger.log(`------> ❌ FAILED: Could not find a matching blueprint for JobID '${jobID}'. Halting generation.`);
                paramsAreValid = false;
            }

            if (paramsAreValid) {
                Logger.log(`------> ✅ PASSED Parameter Safety Gate. Generating test case(s).`);
                const generatedTests = QAGEN_generateTestCasesFromLibrary(job);

                if (generatedTests && generatedTests.length > 0) {
                    const testCase = generatedTests[0];

                    if (testCase.isSuite) {
                        Logger.log(`--------> ACTION: Creating parent meta-task for suite "${jobID}".`);
                        const parentTicket = QAGEN_createYouTrackIssue(testCase, schemaMap);

                        if (parentTicket && parentTicket.idReadable) {
                            Logger.log(`--------> OUTCOME: ✅ Successfully created parent ticket ${parentTicket.idReadable}.`);
                            const url = `${YOUTRACK_BASE_URL}/issue/${parentTicket.idReadable}`;
                            configSheet.getRange(rowIndex, idReadableColIndex + 1).setValue(parentTicket.idReadable);
                            configSheet.getRange(rowIndex, permanentIdColIndex + 1).setValue(parentTicket.id);
                            configSheet.getRange(rowIndex, urlColIndex + 1).setFormula(`=HYPERLINK("${url}","View Suite")`);
                            configSheet.getRange(rowIndex, statusColIndex + 1).setValue('Open');

                            Logger.log(`--------> ACTION: Creating and linking ${testCase.subTasks.length} sub-task(s).`);
                            testCase.subTasks.forEach((subTask, i) => {
                                const childTicket = QAGEN_createYouTrackIssue(subTask, schemaMap);
                                if (childTicket && childTicket.idReadable) {
                                    Logger.log(`-----------> ✅ Created sub-task ${i+1}: ${childTicket.idReadable}. Linking...`);
                                    QAGEN_linkYouTrackIssues_(childTicket.idReadable, parentTicket.idReadable, "subtask of");
                                } else {
                                    Logger.log(`-----------> ❌ FAILED to create sub-task ${i+1}.`);
                                }
                            });

                        } else {
                            Logger.log(`--------> ❌ FAILED: Parent ticket creation returned a null response for JobID '${jobID}'.`);
                        }

                    } else { // It's a single task
                        Logger.log(`--------> ACTION: Creating single ticket for job "${jobID}".`);
                        const newTicket = QAGEN_createYouTrackIssue(testCase, schemaMap);

                        if (newTicket && newTicket.idReadable) {
                            Logger.log(`--------> OUTCOME: ✅ Successfully created YouTrack ticket ${newTicket.idReadable} (Permanent ID: ${newTicket.id}).`);
                            const url = `${YOUTRACK_BASE_URL}/issue/${newTicket.idReadable}`;
                            Logger.log(`--------> ACTION: Performing surgical write-back to row ${rowIndex}.`);
                            configSheet.getRange(rowIndex, idReadableColIndex + 1).setValue(newTicket.idReadable);
                            configSheet.getRange(rowIndex, permanentIdColIndex + 1).setValue(newTicket.id);
                            configSheet.getRange(rowIndex, urlColIndex + 1).setFormula(`=HYPERLINK("${url}","View Ticket")`);
                            configSheet.getRange(rowIndex, statusColIndex + 1).setValue('Open');
                        } else {
                            Logger.log(`--------> ❌ FAILED: Ticket creation returned a null response for JobID '${jobID}'.`);
                        }
                    }

                    if (shouldUpdateLastRun) {
                        Logger.log(`--------> ACTION: Updating 'Last_Run' timestamp for autonomous job.`);
                        configSheet.getRange(rowIndex, lastRunColIndex + 1).setValue(new Date());
                    }

                } else {
                    Logger.log(`------> WARNING: Generation function for JobID '${jobID}' returned no test cases. Nothing to create.`);
                }
            }
        } else {
            Logger.log(`------> ❌ Dependencies not met for JobID '${jobID}'. Halting processing for this job.`);
        }
    });
    Logger.log('--- Exiting Core Processing Engine ---');
}

function QAGEN_evaluateTriggers() {
  Logger.log('--- STARTING AUTONOMOUS TRIGGER EVALUATION ---');
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  const configSheet = ss.getSheetByName('_generator_config');
  if (!configSheet) {
    Logger.log('!!! CRITICAL ERROR: "_generator_config" sheet not found.');
    return;
  }
  
  const headerRange = configSheet.getRange(1, 1, 1, configSheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  const dataRange = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, headers.length);
  const allData = dataRange.getValues();

  const generateColIndex = headers.indexOf('Generate?');
  const triggerTypeColIndex = headers.indexOf('Trigger_Type');
  const triggerValueColIndex = headers.indexOf('Trigger_Value');
  const lastRunColIndex = headers.indexOf('Last_Run');
  const permanentIdColIndex = headers.indexOf('YouTrack_Permanent_ID');
  
  const jobsToEvaluate = allData.map((row, index) => ({ row, originalIndex: index }))
    .filter(item => {
        const generateValue = item.row[generateColIndex];
        const hasPermanentId = item.row[permanentIdColIndex] !== '';
        const hasTriggerType = !!item.row[triggerTypeColIndex];
        return (generateValue === true || String(generateValue).toUpperCase() === 'TRUE') && !hasPermanentId && hasTriggerType;
    });
  
  Logger.log(`Found ${jobsToEvaluate.length} enabled job(s) with a trigger type to evaluate.`);
  const jobsToProcess = [];
  const today = new Date();
  today.setHours(0,0,0,0);
  const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const todayDayName = dayNames[today.getDay()];

  jobsToEvaluate.forEach(item => {
    const job = item.row;
    const jobID = job[headers.indexOf('Job_ID')];
    Logger.log(`--> Evaluating Row ${item.originalIndex + 2} (JobID '${jobID}')`);
    
    const triggerType = job[triggerTypeColIndex];
    const triggerValue = job[triggerValueColIndex];
    const lastRunRaw = job[lastRunColIndex];
    const lastRun = lastRunRaw ? new Date(lastRunRaw) : null;
    if(lastRun) lastRun.setHours(0,0,0,0);
    
    let shouldTrigger = false;
    Logger.log(`----> CONTEXT: Trigger Type is "${triggerType}", Value is "${triggerValue}", Last Run was "${lastRun ? lastRun.toISOString().split('T')[0] : 'never'}"`);

    switch (triggerType) {
      case 'On Date':
        if (!lastRun && triggerValue) {
          const triggerDate = new Date(triggerValue);
          triggerDate.setHours(0,0,0,0);
          if (triggerDate <= today) shouldTrigger = true;
        }
        break;
      case 'Recurring':
        if (lastRun && triggerValue && !isNaN(parseInt(triggerValue))) {
          const daysToWait = parseInt(triggerValue);
          const timeSinceLastRun = today.getTime() - lastRun.getTime();
          const daysSinceLastRun = timeSinceLastRun / (1000 * 3600 * 24);
          if (daysSinceLastRun >= daysToWait) shouldTrigger = true;
        }
        break;
      case 'Weekly':
        if (triggerValue && triggerValue.toLowerCase() === todayDayName.toLowerCase()) {
          if (!lastRun || lastRun.getTime() !== today.getTime()) shouldTrigger = true;
        }
        break;
    }
    
    Logger.log(`------> OUTCOME: Should this job trigger? ${shouldTrigger}`);
    if (shouldTrigger) {
        jobsToProcess.push(item);
    }
  });
  
  if (jobsToProcess.length > 0) {
    Logger.log(`EVALUATION COMPLETE: Found ${jobsToProcess.length} scheduled job(s) to trigger.`);
    _processJobs(jobsToProcess, allData, dataRange, headers, true);
  } else {
    Logger.log('EVALUATION COMPLETE: No scheduled jobs are due to be triggered.');
  }
  Logger.log('--- AUTONOMOUS TRIGGER EVALUATION COMPLETE ---');
}

function QAGEN_syncYouTrackStatuses() {
    Logger.log('--- STARTING YOUTRACK STATUS SYNC ---');
    const startTime = new Date();
    const ss = SpreadsheetApp.openById(MASTER_DB_ID);
    const configSheet = ss.getSheetByName('_generator_config');
    if (!configSheet) { 
        Logger.log('!!! CRITICAL ERROR: "_generator_config" sheet not found.');
        return; 
    }
    
    const dataRange = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, configSheet.getLastColumn());
    const allData = dataRange.getValues();

    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const permanentIdColIndex = headers.indexOf('YouTrack_Permanent_ID');
    const statusColIndex = headers.indexOf('YouTrack_Status');
    const jobIdColIndex = headers.indexOf('Job_ID');
    
    if (permanentIdColIndex === -1 || statusColIndex === -1) {
        Logger.log('!!! CRITICAL ERROR: "YouTrack_Permanent_ID" or "YouTrack_Status" column is missing.');
        return;
    }

    let updatedCount = 0;
    Logger.log(`Found ${allData.length} total jobs to evaluate for status sync.`);

    allData.forEach((row, index) => {
        const ticketId = row[permanentIdColIndex];
        const currentSheetStatus = row[statusColIndex];
        const jobID = row[jobIdColIndex];
        const rowIndexForLog = index + 2;

        if (ticketId && ticketId.trim() !== '') {
            Logger.log(`--> Evaluating Row ${rowIndexForLog} (JobID '${jobID}', PermanentID '${ticketId}')`);
            
            Logger.log('----> ACTION: Querying YouTrack API for current status...');
            const newStatus = QAGEN_getYouTrackIssueDetails_(ticketId);
            
            Logger.log(`------> CONTEXT: Sheet Status is "${currentSheetStatus}", YouTrack Status is "${newStatus}"`);
            
            if (newStatus && currentSheetStatus !== newStatus) {
                Logger.log('--------> OUTCOME: Status has changed. Updating cell.');
                const targetCell = configSheet.getRange(rowIndexForLog, statusColIndex + 1);
                targetCell.setValue(newStatus);
                updatedCount++;
                Utilities.sleep(500); // Be kind to the API
            } else if (newStatus) {
                Logger.log('--------> OUTCOME: Status is already up-to-date. No action taken.');
            } else {
                Logger.log('--------> WARNING: Received null status from YouTrack. Skipping update.');
            }
        }
    });

    if (updatedCount > 0) {
        Logger.log(`\nSYNC COMPLETE: Updated ${updatedCount} ticket status(es).`);
    } else {
        Logger.log("\nSYNC COMPLETE: No status changes found.");
    }

    const endTime = new Date();
    Logger.log(`--- YOUTRACK STATUS SYNC finished in ${(endTime - startTime) / 1000} seconds. ---`);
}

function QAGEN_applyDataValidationRules() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.openById(MASTER_DB_ID);
    const configSheet = ss.getSheetByName('_generator_config');
    const dataRefSheet = ss.getSheetByName('_data_ref');
    if (!configSheet || !dataRefSheet) {
      throw new Error('Could not find "_generator_config" or "_data_ref" sheet.');
    }
    const columnMapping = {
      'Category': 'Category',
      'Sub_Category': 'Sub_Category',
      'Priority': 'Priority',
      'Team_Members': 'Assignee',
      'YouTrack_Type': 'YouTrack_Type',
      'YouTrack_Component': 'Components',
      'Impact_Level': 'Impact_Level' // Added new validation
    };
    const configHeaders = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    const dataRefHeaders = dataRefSheet.getRange(1, 1, 1, dataRefSheet.getLastColumn()).getValues()[0];
    for (const sourceHeader in columnMapping) {
      const targetHeader = columnMapping[sourceHeader];
      const sourceColIndex = dataRefHeaders.indexOf(sourceHeader);
      const targetColIndex = configHeaders.indexOf(targetHeader);
      if (sourceColIndex !== -1 && targetColIndex !== -1) {
        const sourceColumnLetter = String.fromCharCode('A'.charCodeAt(0) + sourceColIndex);
        const validationRange = dataRefSheet.getRange(`${sourceColumnLetter}2:${sourceColumnLetter}${dataRefSheet.getLastRow()}`);
        const targetRange = configSheet.getRange(2, targetColIndex + 1, configSheet.getMaxRows() - 1, 1);
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(validationRange, true)
          .setAllowInvalid(false)
          .setHelpText(`Please select a valid ${targetHeader} from the list.`)
          .build();
        targetRange.setDataValidation(rule);
        console.log(`Applied data validation for "${targetHeader}" column.`);
      } else {
        console.warn(`Could not find matching columns for mapping: ${sourceHeader} -> ${targetHeader}`);
      }
    }
    ui.alert('✅ Data validation rules have been applied successfully to the _generator_config sheet.');
  } catch (e) {
    console.error(`Failed to apply data validation rules: ${e.message}`);
    ui.alert(`❌ An error occurred: ${e.message}`);
  }
}