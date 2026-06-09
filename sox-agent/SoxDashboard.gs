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
    .addItem('Diagnose (show column map)', 'diagnose')
    .addItem('About', 'showAbout')
    .addToUi();
}

function diagnose() {
  var props   = PropertiesService.getScriptProperties();
  var tabName = props.getProperty('SOX_SHEET_TAB') || 'Sheet1';
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(tabName);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet tab "' + tabName + '" not found.\nAvailable tabs: ' +
      ss.getSheets().map(function(s){ return s.getName(); }).join(', '));
    return;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                  .map(function(h){ return String(h).trim(); });

  function findCol() {
    var kws = Array.prototype.slice.call(arguments);
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i].toLowerCase();
      if (kws.every(function(kw){ return h.indexOf(kw.toLowerCase()) >= 0; })) return i;
    }
    return -1;
  }

  var C = {
    controlId : findCol('control', 'id'),
    name      : findCol('control', 'activity') >= 0 ? findCol('control', 'activity') :
                (findCol('control', 'name') >= 0 ? findCol('control', 'name') : findCol('name')),
    objective : findCol('objective') >= 0 ? findCol('objective') :
                (findCol('test plan') >= 0 ? findCol('test plan') :
                (findCol('testing') >= 0 ? findCol('testing') : findCol('activity'))),
    run       : findCol('run'),
    result    : findCol('result') >= 0 ? findCol('result') : (findCol('pass') >= 0 ? findCol('pass') : findCol('verdict')),
    gaps      : findCol('gap') >= 0 ? findCol('gap') : (findCol('notes') >= 0 ? findCol('notes') : findCol('finding')),
    lastRun   : findCol('last run') >= 0 ? findCol('last run') : (findCol('timestamp') >= 0 ? findCol('timestamp') : findCol('date')),
  };

  // Count Run=Yes rows
  var data = sheet.getDataRange().getValues();
  var yesCount = 0;
  var runValues = [];
  for (var i = 1; i < Math.min(data.length, 6); i++) {
    if (C.run >= 0) runValues.push('"' + data[i][C.run] + '"');
  }
  if (C.run >= 0) {
    for (var i = 1; i < data.length; i++) {
      var v = String(data[i][C.run]).toLowerCase().trim();
      if (v === 'yes' || v === 'y') yesCount++;
    }
  }

  // Check Drive folder matches
  var driveFolderId = PropertiesService.getScriptProperties().getProperty('SOX_DRIVE_FOLDER_ID') || '';
  var driveCheck = '';
  if (driveFolderId && C.controlId >= 0) {
    try {
      var rootFolder = DriveApp.getFolderById(driveFolderId);
      var folderNames = [];
      var iter = rootFolder.getFolders();
      while (iter.hasNext()) folderNames.push(iter.next().getName());

      driveCheck = '\nDrive folder check (root has ' + folderNames.length + ' subfolders):\n';
      for (var i = 1; i < data.length; i++) {
        var cid = String(data[i][C.controlId] || '').trim();
        if (!cid) continue;
        var match = folderNames.some(function(fn){ return fn.trim().toLowerCase() === cid.toLowerCase(); });
        driveCheck += '  ' + cid + ': ' + (match ? 'folder found' : 'NO FOLDER FOUND') + '\n';
      }
      driveCheck += '\nActual Drive subfolders:\n' + folderNames.map(function(n){ return '  "' + n + '"'; }).join('\n');
    } catch(e) {
      driveCheck = '\nDrive check failed: ' + e.message;
    }
  }

  var msg = 'Tab: "' + tabName + '" (' + (data.length - 1) + ' data rows)\n\n' +
    'Column mapping:\n' +
    '  Control ID  : ' + (C.controlId >= 0  ? headers[C.controlId]  + ' (col ' + (C.controlId+1)  + ')' : 'NOT FOUND') + '\n' +
    '  Control Name: ' + (C.name >= 0       ? headers[C.name]       + ' (col ' + (C.name+1)       + ')' : 'NOT FOUND') + '\n' +
    '  Objective   : ' + (C.objective >= 0  ? headers[C.objective]  + ' (col ' + (C.objective+1)  + ')' : 'NOT FOUND') + '\n' +
    '  Run?        : ' + (C.run >= 0        ? headers[C.run]        + ' (col ' + (C.run+1)        + ')' : 'NOT FOUND') + '\n' +
    '  Result      : ' + (C.result >= 0     ? headers[C.result]     + ' (col ' + (C.result+1)     + ')' : 'NOT FOUND') + '\n' +
    '  Gaps/Notes  : ' + (C.gaps >= 0       ? headers[C.gaps]       + ' (col ' + (C.gaps+1)       + ')' : 'NOT FOUND') + '\n' +
    '  Last Run    : ' + (C.lastRun >= 0    ? headers[C.lastRun]    + ' (col ' + (C.lastRun+1)    + ')' : 'NOT FOUND') + '\n\n' +
    'Rows with Run = Yes: ' + yesCount + '\n' +
    'First Run column values: ' + (runValues.length ? runValues.join(', ') : 'n/a') +
    driveCheck;

  SpreadsheetApp.getUi().alert(msg);
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
    // Prefer "Control Activity" over "Control Owner Name" for the description
    name       : findCol('control', 'activity') >= 0 ? findCol('control', 'activity') :
                 (findCol('control', 'name') >= 0 ? findCol('control', 'name') : findCol('name')),
    // Use test plan / testing procedures as the objective if no explicit objective column
    objective  : findCol('objective') >= 0 ? findCol('objective') :
                 (findCol('test plan') >= 0 ? findCol('test plan') :
                 (findCol('testing') >= 0 ? findCol('testing') :
                 (findCol('activity') >= 0 ? findCol('activity') : findCol('description')))),
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
    'Done!\n\n' +
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

    // Find the control subfolder — case-insensitive, trimmed match.
    // Supports two structures:
    //   A) rootFolder / {ControlID} / files          (Control ID folders directly in root)
    //   B) rootFolder / {ControlID} / input / files  (evidence in a nested "input" subfolder)
    const controlIdNorm = controlId.trim().toLowerCase();
    let ctrlFolder = null;
    const allSubfolders = root.getFolders();
    while (allSubfolders.hasNext()) {
      const f = allSubfolders.next();
      if (f.getName().trim().toLowerCase() === controlIdNorm) {
        ctrlFolder = f;
        break;
      }
    }
    if (!ctrlFolder) {
      Logger.log('No Drive folder found for control: "' + controlId + '"');
      return evidence;
    }

    // Prefer an "input" subfolder if it exists; otherwise read files directly
    // from the Control ID folder itself.
    let inputFolder = ctrlFolder;
    const inputFolders = ctrlFolder.getFoldersByName('input');
    if (inputFolders.hasNext()) {
      inputFolder = inputFolders.next();
    }

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

      // Convert spreadsheets to CSV text (Gemini does not accept xlsx/xls inline)
      const isSpreadsheet =
        mimeType === 'application/vnd.google-apps.spreadsheet' ||
        mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        mimeType === 'application/vnd.ms-excel' ||
        name.match(/\.(xlsx|xls)$/i);

      const isDocument =
        mimeType === 'application/vnd.google-apps.document' ||
        mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
        mimeType === 'application/msword' ||
        name.match(/\.(docx|doc)$/i);

      if (isSpreadsheet) {
        try {
          // For binary xlsx/xls files, SpreadsheetApp.openById() won't work directly.
          // We create a temporary Google Sheets copy via the Drive REST API, read it,
          // then immediately trash the temp copy.
          let ssId = file.getId();
          let tempId = null;

          if (mimeType !== 'application/vnd.google-apps.spreadsheet') {
            const copyResp = UrlFetchApp.fetch(
              'https://www.googleapis.com/drive/v3/files/' + ssId + '/copy',
              {
                method: 'post',
                headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
                contentType: 'application/json',
                payload: JSON.stringify({
                  name: '_sox_tmp_' + ssId,
                  mimeType: 'application/vnd.google-apps.spreadsheet',
                }),
                muteHttpExceptions: true,
              }
            );
            if (copyResp.getResponseCode() !== 200) {
              Logger.log('Drive copy failed for ' + name + ': ' + copyResp.getContentText().slice(0, 200));
              continue;
            }
            tempId = JSON.parse(copyResp.getContentText()).id;
            ssId = tempId;
          }

          const ss = SpreadsheetApp.openById(ssId);
          let csv = '';
          ss.getSheets().forEach(function(s) {
            csv += '=== Sheet: ' + s.getName() + ' ===\n';
            s.getDataRange().getValues().forEach(function(row) {
              csv += row.map(function(c) {
                const v = String(c === null || c === undefined ? '' : c);
                return v.indexOf(',') >= 0 || v.indexOf('"') >= 0 ? '"' + v.replace(/"/g,'""') + '"' : v;
              }).join(',') + '\n';
            });
            csv += '\n';
          });

          if (tempId) {
            try { DriveApp.getFileById(tempId).setTrashed(true); } catch(te) {}
          }

          evidence.push({
            name: name + '.txt',
            mimeType: 'text/plain',
            data: Utilities.base64Encode(Utilities.newBlob(csv, 'text/plain; charset=utf-8').getBytes()),
          });
          Logger.log('Converted spreadsheet to text: ' + name + ' (' + csv.length + ' chars)');
        } catch (e) {
          Logger.log('Failed to convert spreadsheet ' + name + ': ' + e.message);
        }
        continue;
      }

      if (isDocument) {
        try {
          const doc = DocumentApp.openById(file.getId());
          const txt = doc.getBody().getText();
          evidence.push({
            name: name + '.txt',
            mimeType: 'text/plain',
            data: Utilities.base64Encode(Utilities.newBlob(txt, 'text/plain; charset=utf-8').getBytes()),
          });
          Logger.log('Converted document to text: ' + name + ' (' + txt.length + ' chars)');
        } catch (e) {
          Logger.log('Failed to convert document ' + name + ': ' + e.message);
        }
        continue;
      }

      // PDF and images — send as-is (Gemini supports these natively)
      const geminiMime = resolveGeminiMime(name, mimeType);
      if (!geminiMime) {
        Logger.log('Skipping unsupported type: ' + name + ' (' + mimeType + ')');
        continue;
      }

      const bytes = file.getBlob().getBytes();
      const base64data = Utilities.base64Encode(bytes);
      evidence.push({ name: name, mimeType: geminiMime, data: base64data });
      Logger.log('Added evidence: ' + name + ' (' + bytes.length + ' bytes)');
    }
  } catch (e) {
    Logger.log('Drive error for ' + controlId + ': ' + e.message);
  }
  return evidence;
}

// Only return MIME types Gemini's inline API actually accepts.
// Spreadsheets and documents are converted to text before this is called.
function resolveGeminiMime(filename, driveMime) {
  const ext = (filename.split('.').pop() || '').toLowerCase();
  const supported = {
    pdf: 'application/pdf',
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    webp: 'image/webp',
    txt: 'text/plain',
    csv: 'text/plain',
    md: 'text/plain',
  };
  if (supported[ext]) return supported[ext];
  if (driveMime === 'application/pdf') return 'application/pdf';
  if (driveMime && driveMime.startsWith('image/')) return driveMime;
  if (driveMime === 'text/plain' || driveMime === 'text/csv') return 'text/plain';
  return null; // unsupported — caller will skip
}

// -- Gemini analysis ------------------------------------------------------------

function analyzeControl(apiKey, controlId, controlName, controlObjective, evidence) {
  const GEMINI_MODEL = 'gemini-2.5-flash';
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
  // Hard whitelist — Gemini inline API only accepts these types.
  const GEMINI_OK = { 'text/plain':true, 'application/pdf':true,
    'image/png':true, 'image/jpeg':true, 'image/gif':true, 'image/webp':true };
  for (const ef of evidence) {
    if (!GEMINI_OK[ef.mimeType]) {
      Logger.log('Blocked unsupported MIME before Gemini call: ' + ef.mimeType + ' (' + ef.name + ')');
      continue;
    }
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
