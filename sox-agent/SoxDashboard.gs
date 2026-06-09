// =============================================================================
// SOX Dashboard - Google Apps Script
// Paste the full contents of this file into the Apps Script editor of the
// SOX Dashboard Launcher spreadsheet (Extensions -> Apps Script).
// =============================================================================
//
// FIRST-TIME SETUP
// 1. Paste this file -> Save -> Run onOpen (or reload the sheet).
// 2. Click SOX Dashboard ->  Configure -> enter your Gemini API key.
//    Get a free key at: https://aistudio.google.com/apikey
// 3. Enter the Google Drive folder ID for SOX evidence
//    (open the SOX_Dashboards folder in Drive, copy the ID from the URL).
// 4. Click SOX Dashboard ->  Run SOX Tests.
// =============================================================================

// -- Menu ---------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SOX Dashboard')
    .addItem('Run SOX Tests', 'runSoxTests')
    .addSeparator()
    .addItem('Configure (API key & folder)', 'showConfig')
    .addItem('About', 'showAbout')
    .addToUi();
}

// -- Configuration dialog ------------------------------------------------------

function showConfig() {
  const props = PropertiesService.getScriptProperties();
  const currentKey    = props.getProperty('GEMINI_API_KEY')    || '';
  const currentFolder = props.getProperty('SOX_DRIVE_FOLDER_ID') || '';
  const currentTab    = props.getProperty('SOX_SHEET_TAB')     || 'Sheet1';

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; font-size: 13px; padding: 12px; }
      label { display: block; margin-top: 12px; font-weight: bold; }
      input  { width: 100%; box-sizing: border-box; padding: 5px; margin-top: 4px;
               border: 1px solid #ccc; border-radius: 3px; }
      .hint  { color: #888; font-size: 11px; margin-top: 2px; }
      button { margin-top: 16px; padding: 7px 18px; background: #1967D2;
               color: #fff; border: none; border-radius: 4px; cursor: pointer; }
    </style>
    <form>
      <label>Gemini API Key</label>
      <input id="key" type="password" value="${escHtml(currentKey)}" placeholder="AIza...">
      <div class="hint">Get a free key at <a href="https://aistudio.google.com/apikey" target="_blank">aistudio.google.com/apikey</a></div>

      <label>SOX Drive Folder ID</label>
      <input id="folder" value="${escHtml(currentFolder)}" placeholder="1AbCdEfGhIjKlMnOpQrStUvWxYz">
      <div class="hint">Open the SOX_Dashboards folder in Drive, copy the ID from the URL</div>

      <label>Sheet Tab Name</label>
      <input id="tab" value="${escHtml(currentTab)}" placeholder="Sheet1">

      <button onclick="save()">Save</button>
    </form>
    <script>
      function escHtml(s){ return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/"/g,'&quot;'); }
      function save(){
        const key    = document.getElementById('key').value.trim();
        const folder = document.getElementById('folder').value.trim();
        const tab    = document.getElementById('tab').value.trim();
        google.script.run.withSuccessHandler(()=>{ google.script.host.close(); })
          .saveConfig(key, folder, tab);
      }
    </script>
  `)
    .setTitle('SOX Dashboard - Configuration')
    .setWidth(400)
    .setHeight(340);

  SpreadsheetApp.getUi().showModalDialog(html, 'SOX Dashboard - Configuration');
}

function escHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/"/g, '&quot;');
}

function saveConfig(apiKey, driveFolderId, sheetTab) {
  const props = PropertiesService.getScriptProperties();
  if (apiKey)       props.setProperty('GEMINI_API_KEY',       apiKey);
  if (driveFolderId) props.setProperty('SOX_DRIVE_FOLDER_ID', driveFolderId);
  if (sheetTab)     props.setProperty('SOX_SHEET_TAB',        sheetTab);
}

function showAbout() {
  SpreadsheetApp.getUi().alert(
    'SOX Dashboard Agent\n\n' +
    'For each row with Run? = Yes:\n' +
    '  1. Reads evidence files from Drive (SOX_Dashboards / <Control ID> / input /)\n' +
    '  2. Sends them to Gemini AI for analysis\n' +
    '  3. Writes Pass/Fail, gaps, and timestamp back to the sheet\n\n' +
    'Configure via SOX Dashboard ->  Configure'
  );
}

// -- Main runner ---------------------------------------------------------------

function runSoxTests() {
  const props = PropertiesService.getScriptProperties();
  const apiKey    = props.getProperty('GEMINI_API_KEY');
  const folderId  = props.getProperty('SOX_DRIVE_FOLDER_ID');
  const tabName   = props.getProperty('SOX_SHEET_TAB') || 'Sheet1';

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('ERROR: Gemini API key is not set.\nGo to SOX Dashboard ->  Configure.');
    return;
  }
  if (!folderId) {
    SpreadsheetApp.getUi().alert('ERROR: SOX Drive folder ID is not set.\nGo to SOX Dashboard ->  Configure.');
    return;
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('ERROR: Sheet tab "' + tabName + '" not found.');
    return;
  }

  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) {
    SpreadsheetApp.getUi().alert('Sheet appears to be empty.');
    return;
  }

  const headers = allData[0].map(h => String(h).toLowerCase().trim());

  // -- Column detection (same fuzzy logic as the Python script) -------------
  function findCol() {
    const keywords = Array.from(arguments);
    for (let i = 0; i < headers.length; i++) {
      if (keywords.every(kw => headers[i].includes(kw.toLowerCase()))) return i;
    }
    return -1;
  }

  const C = {
    controlId  : findCol('control', 'id'),
    name       : findCol('control', 'name') >= 0 ? findCol('control', 'name') : findCol('name'),
    objective  : findCol('objective') >= 0 ? findCol('objective') : findCol('description'),
    run        : findCol('run'),
    result     : findCol('result') >= 0 ? findCol('result') :
                 (findCol('pass') >= 0 ? findCol('pass') : findCol('verdict')),
    gaps       : findCol('gap') >= 0 ? findCol('gap') :
                 (findCol('notes') >= 0 ? findCol('notes') : findCol('finding')),
    lastRun    : findCol('last run') >= 0 ? findCol('last run') :
                 (findCol('timestamp') >= 0 ? findCol('timestamp') : findCol('date')),
  };

  if (C.controlId < 0) {
    SpreadsheetApp.getUi().alert('ERROR: Could not find a "Control ID" column in row 1.\nHeaders found: ' + allData[0].join(', '));
    return;
  }
  if (C.run < 0) {
    SpreadsheetApp.getUi().alert('ERROR: Could not find a "Run?" column in row 1.\nHeaders found: ' + allData[0].join(', '));
    return;
  }

  // -- Process rows ----------------------------------------------------------
  let processed = 0;
  let failed    = 0;

  for (let i = 1; i < allData.length; i++) {
    const row      = allData[i];
    const runFlag  = String(row[C.run] || '').toLowerCase().trim();
    if (runFlag !== 'yes' && runFlag !== 'y') continue;

    const controlId  = String(row[C.controlId]  || '').trim();
    const ctrlName   = C.name      >= 0 ? String(row[C.name]      || '').trim() : '';
    const objective  = C.objective >= 0 ? String(row[C.objective] || '').trim() : '';

    if (!controlId) continue;

    // Toast notification - visible in the sheet while running
    ss.toast('Analyzing ' + controlId + ' (' + ctrlName + ')...', 'SOX Agent', -1);

    // Get evidence files
    const evidence = getEvidenceFiles(folderId, controlId);

    // Call Gemini
    let result;
    try {
      result = analyzeControl(apiKey, controlId, ctrlName, objective, evidence);
    } catch (e) {
      result = { verdict: 'Failed', gaps: ['Agent error: ' + e.message], summary: String(e) };
    }

    const rowNum = i + 1; // 1-indexed

    // Write verdict
    if (C.result >= 0) {
      const cell = sheet.getRange(rowNum, C.result + 1);
      cell.setValue(result.verdict);
      cell.setBackground(result.verdict === 'Passed' ? '#d9ead3' : '#fce5cd');
    }

    // Write gaps
    if (C.gaps >= 0 && result.gaps && result.gaps.length) {
      const gapText = result.gaps.map(g => '- ' + g).join('\n');
      sheet.getRange(rowNum, C.gaps + 1).setValue(
        result.summary ? gapText + '\n\n' + result.summary : gapText
      );
    } else if (C.gaps >= 0) {
      sheet.getRange(rowNum, C.gaps + 1).setValue(result.summary || '');
    }

    // Write timestamp
    if (C.lastRun >= 0) {
      sheet.getRange(rowNum, C.lastRun + 1).setValue(new Date().toLocaleString());
    }

    processed++;
    if (result.verdict !== 'Passed') failed++;
  }

  ss.toast('', '', 1); // clear toast
  SpreadsheetApp.getUi().alert(
    'Done! Done!\n\n' +
    'Controls tested: ' + processed + '\n' +
    'Passed: ' + (processed - failed) + '\n' +
    'Failed / has gaps: ' + failed
  );
}

// -- Google Drive helpers ------------------------------------------------------

function getEvidenceFiles(rootFolderId, controlId) {
  const evidence = [];
  try {
    const root = DriveApp.getFolderById(rootFolderId);

    // Find subfolder matching Control ID
    const ctrlFolders = root.getFoldersByName(controlId);
    if (!ctrlFolders.hasNext()) {
      Logger.log('No Drive folder for control: ' + controlId);
      return evidence;
    }
    const ctrlFolder = ctrlFolders.next();

    // Find "input" subfolder inside it
    const inputFolders = ctrlFolder.getFoldersByName('input');
    if (!inputFolders.hasNext()) {
      Logger.log('No "input" subfolder for control: ' + controlId);
      return evidence;
    }
    const inputFolder = inputFolders.next();

    // Collect all files
    const files = inputFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name     = file.getName();
      const mimeType = file.getMimeType();
      const sizeBytes = file.getSize();

      // Skip very large files (>8 MB - Gemini inline limit)
      if (sizeBytes > 8 * 1024 * 1024) {
        Logger.log('Skipping large file: ' + name + ' (' + sizeBytes + ' bytes)');
        continue;
      }

      let blob = file.getBlob();

      // Export Google Workspace formats to Office-compatible formats
      if (mimeType === 'application/vnd.google-apps.spreadsheet') {
        blob = file.exportLinks
          ? Drive.Files.export(file.getId(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
          : UrlFetchApp.fetch(
              'https://docs.google.com/spreadsheets/d/' + file.getId() + '/export?format=xlsx',
              { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }
            ).getBlob();
      } else if (mimeType === 'application/vnd.google-apps.document') {
        blob = UrlFetchApp.fetch(
          'https://docs.google.com/document/d/' + file.getId() + '/export?format=docx',
          { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }
        ).getBlob();
      }

      const bytes = blob.getBytes();
      const base64data = Utilities.base64Encode(bytes);

      evidence.push({
        name: name,
        mimeType: resolveEffectiveMime(name, mimeType),
        data: base64data,
      });

      Logger.log('Added evidence: ' + name + ' (' + bytes.length + ' bytes)');
    }
  } catch (e) {
    Logger.log('Drive error for ' + controlId + ': ' + e.message);
  }
  return evidence;
}

function resolveEffectiveMime(filename, driveMime) {
  const ext = filename.split('.').pop().toLowerCase();
  const extMap = {
    pdf: 'application/pdf',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    xls: 'application/vnd.ms-excel',
    csv: 'text/csv',
    txt: 'text/plain',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    doc: 'application/msword',
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
  };
  if (extMap[ext]) return extMap[ext];
  // Google Workspace exports
  if (driveMime === 'application/vnd.google-apps.spreadsheet')
    return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  if (driveMime === 'application/vnd.google-apps.document')
    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  return driveMime || 'application/octet-stream';
}

// -- Gemini analysis ------------------------------------------------------------

function analyzeControl(apiKey, controlId, controlName, controlObjective, evidence) {
  const GEMINI_MODEL = 'gemini-2.0-flash';
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
              GEMINI_MODEL + ':generateContent?key=' + apiKey;

  const promptText =
    'You are a SOX (Sarbanes-Oxley) compliance auditor.\n\n' +
    'CONTROL DETAILS\n' +
    'Control ID:        ' + controlId + '\n' +
    'Control Name:      ' + controlName + '\n' +
    'Control Objective: ' + (controlObjective || '(not specified)') + '\n\n' +
    'INSTRUCTIONS\n' +
    '1. Review all attached evidence files carefully.\n' +
    '2. Determine whether the evidence shows the control is operating effectively.\n' +
    '3. If failing, list specific, actionable gaps - not vague statements.\n\n' +
    'RESPONSE FORMAT - respond with valid JSON only, no markdown fences:\n' +
    '{\n' +
    '  "verdict": "Passed" or "Failed",\n' +
    '  "gaps": ["gap 1", "gap 2"],\n' +
    '  "summary": "One or two sentence summary"\n' +
    '}\n\n' +
    'If no evidence files were provided:\n' +
    '{"verdict":"Failed","gaps":["No evidence files found in the input folder"],' +
    '"summary":"Unable to test - no evidence uploaded"}';

  // Build parts array: text prompt + inline file data
  const parts = [{ text: promptText }];
  for (const ef of evidence) {
    parts.push({ inlineData: { mimeType: ef.mimeType, data: ef.data } });
  }

  const payload = {
    contents: [{ role: 'user', parts: parts }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0 },
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const body = response.getContentText();

  if (statusCode !== 200) {
    Logger.log('Gemini API error ' + statusCode + ': ' + body.slice(0, 500));
    throw new Error('Gemini API returned HTTP ' + statusCode + ': ' + body.slice(0, 200));
  }

  const json = JSON.parse(body);
  let raw = (json.candidates || [])[0]?.content?.parts?.[0]?.text || '';

  // Strip markdown fences just in case
  const fenceMatch = raw.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (fenceMatch) raw = fenceMatch[1].trim();

  try {
    return JSON.parse(raw);
  } catch (e) {
    Logger.log('Non-JSON from Gemini: ' + raw.slice(0, 300));
    return {
      verdict: 'Failed',
      gaps: ['Agent error: could not parse model response'],
      summary: raw.slice(0, 200),
    };
  }
}
