// YouTrackSchemaManager.js - reads the project's custom fields and their allowed values; caches them only when they change.

/**
 * Main "start button" to query YouTrack for its project schema and cache it.
 * This function now includes fetching all valid values for enum-type fields.
 * @returns {Map<string, object>|null} A map of the schema, or null on critical failure.
 */
function QAGEN_updateYouTrackSchemaCache() {
  try {
    console.log('Checking for YouTrack schema updates...');
    const masterDb = SpreadsheetApp.openById(MASTER_DB_ID);
    const cacheSheet = masterDb.getSheetByName('_youtrack_schema_cache');
    if (!cacheSheet) throw new Error('Cache sheet "_youtrack_schema_cache" not found.');

    // Step 1: Always get the latest schema from YouTrack to calculate the hash.
    let projectFields = QAGEN_getProjectFields_();
    projectFields = projectFields.filter(field => field && field.name);
    // Sort fields alphabetically for a consistent hash
    projectFields.sort((a, b) => a.name.localeCompare(b.name));
    // Sort the values within each field for a consistent hash
    projectFields.forEach(field => {
      if (field.values) {
        field.values = field.values.filter(value => typeof value === 'string');
        field.values.sort();
      }
    });

    const schemaJson = JSON.stringify(projectFields);
    const newHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, schemaJson)
      .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
    const oldHash = PropertiesService.getScriptProperties().getProperty('YOUTRACK_SCHEMA_HASH');
    const isCacheEmpty = cacheSheet.getLastRow() < 2;

    const schemaMap = new Map(projectFields.map(field => [field.name, { 
        isRequired: field.isRequired, 
        type: field.$type, 
        values: field.values // Now includes the list of valid values
    }]));

    // Step 2: Decide if a sheet update is needed.
    if (newHash === oldHash && !isCacheEmpty) {
      console.log('✅ YouTrack schema is unchanged. No write needed. Returning map from live data.');
      return schemaMap;
    }

    // Step 3: If an update IS needed, write the new schema to the sheet.
    console.log(isCacheEmpty ? 'Cache is empty. Populating...' : 'Schema has changed. Updating cache...');
    
    cacheSheet.clear();
    cacheSheet.appendRow(['FieldName', 'IsRequired', 'Type', 'Values (JSON)']); // Added new column header
    projectFields.forEach(field => {
      const valuesJson = field.values ? JSON.stringify(field.values) : '';
      cacheSheet.appendRow([field.name, field.isRequired, field.$type, valuesJson]);
    });
    
    SpreadsheetApp.flush(); // Wait for the write to complete.
    PropertiesService.getScriptProperties().setProperty('YOUTRACK_SCHEMA_HASH', newHash);
    console.log('✅ Successfully updated YouTrack schema cache and hash.');
    
    // Step 4: Return the newly built schema map from memory.
    return schemaMap;

  } catch (e) {
    console.error(`❌ Failed to update YouTrack schema cache: ${e.message}`);
    return null;
  }
}


/**
 * Private helper to get all custom fields, now including their value bundles.
 */
function QAGEN_getProjectFields_() {
  // --- ENHANCED QUERY ---
  // We now ask for the 'bundle' which contains the 'values' for any enum fields.
  const fieldsQuery = 'fields=field(name),isRequired,$type,bundle(values(name))';
  const endpoint = `${YOUTRACK_BASE_URL}/api/admin/projects/${YOUTRACK_PROJECT_ID}/customFields?${fieldsQuery}`;
  const headers = QAGEN_getYouTrackHeaders_();
  const options = { headers: headers, muteHttpExceptions: true };
  const response = UrlFetchApp.fetch(endpoint, options);
  if (response.getResponseCode() !== 200) {
    throw new Error(`Could not fetch project fields. Status: ${response.getResponseCode()}`);
  }
  const rawFields = JSON.parse(response.getContentText());
  
  // We now map the response correctly, extracting the values from the nested bundle.
  return rawFields.map(f => ({
    name: f.field.name,
    isRequired: f.isRequired,
    $type: f.$type,
    // If a bundle with values exists, map them to a simple array of strings. Otherwise, null.
    values: f.bundle && f.bundle.values ? f.bundle.values.map(v => v.name) : null
  }));
}

/**
 * A utility function to manually clear the stored schema hash.
 */
function QAGEN_clearSchemaHash() {
  PropertiesService.getScriptProperties().deleteProperty('YOUTRACK_SCHEMA_HASH');
  console.log('✅ Stored YouTrack schema hash has been cleared. The cache will now be rebuilt.');
}