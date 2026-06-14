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
    .addItem('Open SOX Drive Folder', 'openSoxDriveFolder')
    .addSeparator()
    .addItem('Configure (API key & folder)', 'showConfig')
    .addItem('Setup Instructions & Colors', 'setupInstructions')
    .addItem('Diagnose (show column map)', 'diagnose')
    .addItem('Test File Reading', 'testFileReading')
    .addItem('About', 'showAbout')
    .addToUi();
}

function openSoxDriveFolder() {
  const folderId = PropertiesService.getScriptProperties().getProperty('SOX_DRIVE_FOLDER_ID');
  if (!folderId) {
    SpreadsheetApp.getUi().alert('Drive folder ID is not configured.\nGo to SOX Dashboard > Configure first.');
    return;
  }
  const url = 'https://drive.google.com/drive/folders/' + folderId;
  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '","_blank"); google.script.host.close();</script>'
  ).setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Drive folder...');
}

// =============================================================================
// SETUP: Instructions tab + Sheet1 header color-coding + Run History rebuild
// =============================================================================

function setupInstructions() {
  const props   = PropertiesService.getScriptProperties();
  const tabName = props.getProperty('SOX_SHEET_TAB') || 'Sheet1';
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(tabName);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "' + tabName + '" not found. Run Configure first.');
    return;
  }

  // ── Detect columns (same logic as runSoxTests) ──────────────────────────
  const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers    = rawHeaders.map(h => String(h).trim());
  const headersLow = headers.map(h => h.toLowerCase());

  function findCol() {
    const kws = Array.from(arguments);
    for (let i = 0; i < headersLow.length; i++) {
      if (kws.every(k => headersLow[i].includes(k.toLowerCase()))) return i;
    }
    return -1;
  }

  const C = {
    controlId  : findCol('control', 'id'),
    name       : findCol('control', 'activity') >= 0 ? findCol('control', 'activity') :
                 (findCol('control', 'name') >= 0 ? findCol('control', 'name') : findCol('name')),
    objective  : findCol('objective') >= 0 ? findCol('objective') :
                 (findCol('test plan') >= 0 ? findCol('test plan') : findCol('description')),
    procedures : findCol('procedure') >= 0 ? findCol('procedure') :
                 (findCol('testing procedure') >= 0 ? findCol('testing procedure') :
                 (findCol('testing') >= 0 ? findCol('testing') : -1)),
    period     : findCol('testing period') >= 0 ? findCol('testing period') :
                 (findCol('test period') >= 0 ? findCol('test period') :
                 (findCol('period') >= 0 ? findCol('period') : -1)),
    publishDate: findCol('publishing report') >= 0 ? findCol('publishing report') :
                 (findCol('publish') >= 0 ? findCol('publish') : -1),
    run        : findCol('run'),
    result     : findCol('result') >= 0 ? findCol('result') :
                 (findCol('pass') >= 0 ? findCol('pass') : findCol('verdict')),
    gaps       : findCol('gap') >= 0 ? findCol('gap') :
                 (findCol('notes') >= 0 ? findCol('notes') : findCol('finding')),
    lastRun    : findCol('last run') >= 0 ? findCol('last run') :
                 (findCol('timestamp') >= 0 ? findCol('timestamp') : findCol('date')),
  };

  // ── Color palette ────────────────────────────────────────────────────────
  const COLOR = {
    gemini    : { bg: '#c9daf8', text: '#1155cc', label: 'Sent to Gemini (AI input)'       },
    output    : { bg: '#d9ead3', text: '#38761d', label: 'Written by the script (output)'  },
    control   : { bg: '#fce5cd', text: '#b45309', label: 'Controls processing (trigger)'   },
    unused    : { bg: '#f3f3f3', text: '#666666', label: 'Other columns (not used)'         },
    sectionHdr: '#1e3a5f',
    white     : '#ffffff',
    titleBg   : '#1e3a5f',
  };

  // Columns by group (0-indexed)
  const geminiCols  = [C.controlId, C.name, C.objective, C.procedures, C.period, C.publishDate].filter(c => c >= 0);
  const outputCols  = [C.result, C.gaps, C.lastRun].filter(c => c >= 0);
  const controlCols = [C.run].filter(c => c >= 0);

  // ── 1. Color-code Sheet1 header row ─────────────────────────────────────
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setBackground(COLOR.unused.bg).setFontColor(COLOR.unused.text).setFontWeight('normal');

  function colorCols(cols, bg, textColor) {
    cols.forEach(c => {
      const cell = sheet.getRange(1, c + 1);
      cell.setBackground(bg).setFontColor(textColor).setFontWeight('bold');
    });
  }
  colorCols(geminiCols,  COLOR.gemini.bg,  COLOR.gemini.text);
  colorCols(outputCols,  COLOR.output.bg,  COLOR.output.text);
  colorCols(controlCols, COLOR.control.bg, COLOR.control.text);

  // ── 2. Rebuild Instructions tab ─────────────────────────────────────────
  let instrSheet = ss.getSheetByName('Instructions');
  if (instrSheet) ss.deleteSheet(instrSheet);
  instrSheet = ss.insertSheet('Instructions');

  // Helper: write a styled row
  function writeRow(rowIdx, values, bgColor, textColor, bold, fontSize, merged) {
    const r = instrSheet.getRange(rowIdx, 1, 1, values.length);
    r.setValues([values]);
    if (bgColor)   r.setBackground(bgColor);
    if (textColor) r.setFontColor(textColor);
    if (bold)      r.setFontWeight('bold');
    if (fontSize)  r.setFontSize(fontSize);
    if (merged && values.length > 1) instrSheet.getRange(rowIdx, 1, 1, 5).merge();
  }

  function sectionHeader(rowIdx, title) {
    instrSheet.getRange(rowIdx, 1, 1, 5).merge()
      .setValue(title)
      .setBackground(COLOR.sectionHdr)
      .setFontColor(COLOR.white)
      .setFontWeight('bold')
      .setFontSize(11);
  }

  function colName(idx) {
    return idx >= 0 ? '"' + headers[idx] + '" (col ' + (idx + 1) + ')' : 'Not found';
  }

  let row = 1;

  // Title
  instrSheet.getRange(row, 1, 1, 5).merge()
    .setValue('SOX Dashboard Agent - Instructions')
    .setBackground(COLOR.titleBg)
    .setFontColor(COLOR.white)
    .setFontWeight('bold')
    .setFontSize(16);
  instrSheet.setRowHeight(row, 40);
  row++;

  instrSheet.getRange(row, 1, 1, 5).merge()
    .setValue('Last updated: ' + new Date().toLocaleString() + '  |  Main data tab: "' + tabName + '"')
    .setBackground('#3c5f91')
    .setFontColor('#d9e8ff')
    .setFontSize(9)
    .setFontStyle('italic');
  row += 2;

  // ── HOW IT WORKS ─────────────────────────────────────────────────────────
  sectionHeader(row, '  HOW IT WORKS'); row++;
  const howItWorks = [
    ['1.', 'Read sheet', 'The script reads every row in "' + tabName + '" where the Run? column = "Yes".'],
    ['2.', 'Find evidence files', 'For each row it looks up the Drive folder with the same name as the Control ID (case-insensitive). All files inside that folder (and subfolders) are collected.'],
    ['3.', 'Convert files', 'XLSX / Google Sheets files are converted to text (all tabs included). The header row of each sheet is always kept. If a file is large, the most recent rows are kept and older ones are omitted with a note. Word/Docs are extracted as plain text. PDFs and images are sent as-is.'],
    ['4.', 'Call Gemini AI', 'The control details + evidence files are sent to Google Gemini (gemini-2.5-pro). Gemini checks whether every step in the Testing Procedures has adequate, in-period evidence.'],
    ['5.', 'Write results', 'The verdict (Passed / Failed), gaps, and timestamp are written back to the row. A record is also appended to the Run History tab.'],
  ];
  howItWorks.forEach(([num, title, desc]) => {
    instrSheet.getRange(row, 1).setValue(num).setFontWeight('bold');
    instrSheet.getRange(row, 2).setValue(title).setFontWeight('bold');
    instrSheet.getRange(row, 3, 1, 3).merge().setValue(desc);
    row++;
  });
  row++;

  // ── PASS / FAIL RULES ────────────────────────────────────────────────────
  sectionHeader(row, '  PASS / FAIL RULES'); row++;
  const rules = [
    ['PASSED', 'ALL testing procedure steps have clear evidence within the Testing Period, and the review/sign-off date is on or before the Publishing Report Date.'],
    ['FAILED', 'Any procedure step is missing evidence, evidence is outside the Testing Period, or the review date is after the Publishing Report Date.'],
    ['Grace period', 'Review/sign-off dates (auditor approval, supervisor sign-off on working papers) that fall between the Testing Period end and the Publishing Report Date are acceptable.'],
  ];
  rules.forEach(([verdict, desc]) => {
    instrSheet.getRange(row, 1).setValue(verdict).setFontWeight('bold')
      .setBackground(verdict === 'PASSED' ? '#d9ead3' : verdict === 'FAILED' ? '#fce5cd' : '#f3f3f3');
    instrSheet.getRange(row, 2, 1, 4).merge().setValue(desc);
    row++;
  });
  row++;

  // ── DRIVE FOLDER STRUCTURE ───────────────────────────────────────────────
  sectionHeader(row, '  DRIVE FOLDER STRUCTURE'); row++;
  const driveLines = [
    ['Root folder', '(configured in SOX Dashboard > Configure - paste the folder ID from the Drive URL)'],
    ['', ''],
    ['', 'SOX_Dashboards/                    <- Root folder'],
    ['', '   IL.FSCP.04/                      <- Folder name must match Control ID exactly'],
    ['', '      evidence_file.xlsx            <- Files directly in control folder'],
    ['', '      another_file.pdf'],
    ['', '   IL.FSCP.06/'],
    ['', '      input/                         <- OR files in an "input" subfolder'],
    ['', '         evidence.xlsx'],
  ];
  driveLines.forEach(([label, desc]) => {
    instrSheet.getRange(row, 1).setValue(label).setFontWeight(label ? 'bold' : 'normal');
    instrSheet.getRange(row, 2, 1, 4).merge().setValue(desc)
      .setFontFamily(label ? null : 'Courier New').setFontSize(label ? 10 : 9);
    row++;
  });
  row++;

  // ── COLUMN REFERENCE ─────────────────────────────────────────────────────
  sectionHeader(row, '  COLUMN REFERENCE - "' + tabName + '" tab'); row++;

  // Legend row
  const legendCols = [COLOR.gemini, COLOR.control, COLOR.output];
  ['Color', 'Group', 'Description', '', ''].forEach((h, i) => {
    instrSheet.getRange(row, i + 1).setValue(h).setFontWeight('bold')
      .setBackground('#e8eaf6').setFontColor('#333333');
  });
  row++;
  legendCols.forEach(g => {
    instrSheet.getRange(row, 1).setBackground(g.bg);
    instrSheet.getRange(row, 2, 1, 4).merge().setValue(g.label)
      .setBackground(g.bg).setFontColor(g.text).setFontWeight('bold');
    row++;
  });
  instrSheet.getRange(row, 1).setBackground(COLOR.unused.bg);
  instrSheet.getRange(row, 2, 1, 4).merge().setValue(COLOR.unused.label)
    .setBackground(COLOR.unused.bg).setFontColor(COLOR.unused.text);
  row += 2;

  // Column table header
  ['Column #', 'Header in sheet', 'Role', 'Color group', 'Notes'].forEach((h, i) => {
    instrSheet.getRange(row, i + 1).setValue(h).setFontWeight('bold')
      .setBackground('#e8eaf6');
  });
  row++;

  function colRow(idx, role, colorGroup, notes) {
    const g = COLOR[colorGroup] || COLOR.unused;
    const cells = instrSheet.getRange(row, 1, 1, 5);
    cells.setValues([[
      idx >= 0 ? 'Col ' + (idx + 1) : 'N/A',
      idx >= 0 ? headers[idx] : '(not detected)',
      role,
      g.label,
      notes || '',
    ]]);
    cells.setBackground(g.bg);
    if (colorGroup !== 'unused') cells.setFontColor(g.text);
    row++;
  }

  colRow(C.controlId,   'Control ID',              'gemini',  'Must match the Drive folder name');
  colRow(C.name,        'Control Description',     'gemini',  'Describes what the control does');
  colRow(C.objective,   'Test Plan',                'gemini',  'High-level testing objective');
  colRow(C.procedures,  'Testing Procedures',       'gemini',  'Checklist - every step must have evidence');
  colRow(C.period,      'Testing Period',           'gemini',  'Date range evidence must fall within');
  colRow(C.publishDate, 'Publishing Report Date',   'gemini',  'Review/sign-off must be on or before this date');
  row++;
  colRow(C.run,         'Run? (Yes / No)',          'control', 'Set to "Yes" to include this row in the next run');
  row++;
  colRow(C.result,      'Passed / Failed',          'output',  'Written by script - green = Passed, orange = Failed');
  colRow(C.gaps,        'Gaps / Notes',             'output',  'Written by script - lists each gap with step reference');
  colRow(C.lastRun,     'Last Run date and Time',   'output',  'Written by script - timestamp of last run');
  row += 2;

  // ── CONFIGURATION ────────────────────────────────────────────────────────
  sectionHeader(row, '  CONFIGURATION  (SOX Dashboard > Configure)'); row++;
  const configRows = [
    ['Gemini API Key', 'Get a free key at https://aistudio.google.com/apikey - required for AI analysis'],
    ['Drive Folder ID', 'Open the root SOX evidence folder in Drive, copy the ID from the URL (the part after /folders/)'],
    ['Sheet Tab Name', 'Name of the tab that contains the control list (default: Sheet1)'],
  ];
  configRows.forEach(([setting, desc]) => {
    instrSheet.getRange(row, 1).setValue(setting).setFontWeight('bold').setBackground('#f8f9fa');
    instrSheet.getRange(row, 2, 1, 4).merge().setValue(desc).setBackground('#f8f9fa');
    row++;
  });

  // ── Format Instructions sheet ────────────────────────────────────────────
  instrSheet.setColumnWidth(1, 120);
  instrSheet.setColumnWidth(2, 220);
  instrSheet.setColumnWidth(3, 320);
  instrSheet.setColumnWidth(4, 200);
  instrSheet.setColumnWidth(5, 260);
  instrSheet.setTabColor('#1e3a5f');
  instrSheet.setFrozenRows(1);

  // ── 3. Run History tab — create if missing, preserve data if it exists ───
  const RUN_HISTORY_TAB = 'Run History';
  const histHeaders = ['Run Date & Time', 'Run By', 'Control ID', 'Control Name',
                       'Testing Period', 'Verdict', 'Gaps / Notes'];
  let histSheet = ss.getSheetByName(RUN_HISTORY_TAB);
  const histIsNew = !histSheet;
  if (histIsNew) {
    histSheet = ss.insertSheet(RUN_HISTORY_TAB);
    histSheet.appendRow(histHeaders);
  }
  // Always reformat the header row and column widths (no data is touched)
  const hdr = histSheet.getRange(1, 1, 1, histHeaders.length);
  hdr.setValues([histHeaders])
     .setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  histSheet.setFrozenRows(1);
  histSheet.setColumnWidth(1, 165);
  histSheet.setColumnWidth(2, 200);
  histSheet.setColumnWidth(3, 110);
  histSheet.setColumnWidth(4, 220);
  histSheet.setColumnWidth(5, 150);
  histSheet.setColumnWidth(6, 90);
  histSheet.setColumnWidth(7, 450);
  histSheet.setTabColor('#38761d');

  // Move sheets to logical order: main sheet, Instructions, Run History
  ss.setActiveSheet(instrSheet);
  ss.moveActiveSheet(2);
  ss.setActiveSheet(histSheet);
  ss.moveActiveSheet(3);
  ss.setActiveSheet(sheet);

  SpreadsheetApp.getUi().alert(
    'Done!\n\n' +
    'Instructions tab created with column reference.\n' +
    'Sheet1 header row color-coded:\n' +
    '  Blue   = ' + geminiCols.length  + ' columns sent to Gemini\n' +
    '  Orange = ' + controlCols.length + ' column controls processing\n' +
    '  Green  = ' + outputCols.length  + ' columns written by script\n\n' +
    (histIsNew ? 'Run History tab created.' : 'Run History tab reformatted (existing data preserved).')
  );
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
    controlId  : findCol('control', 'id'),
    name       : findCol('control', 'activity') >= 0 ? findCol('control', 'activity') :
                 (findCol('control', 'name') >= 0 ? findCol('control', 'name') : findCol('name')),
    objective  : findCol('objective') >= 0 ? findCol('objective') :
                 (findCol('test plan') >= 0 ? findCol('test plan') : findCol('description')),
    procedures : findCol('procedure') >= 0 ? findCol('procedure') :
                 (findCol('testing procedure') >= 0 ? findCol('testing procedure') :
                 (findCol('testing') >= 0 ? findCol('testing') : -1)),
    period      : findCol('testing period') >= 0 ? findCol('testing period') :
                  (findCol('test period') >= 0 ? findCol('test period') :
                  (findCol('period') >= 0 ? findCol('period') : -1)),
    publishDate : findCol('publishing report') >= 0 ? findCol('publishing report') :
                  (findCol('publish') >= 0 ? findCol('publish') : -1),
    run        : findCol('run'),
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
    '  Control ID        : ' + (C.controlId  >= 0 ? headers[C.controlId]  + ' (col ' + (C.controlId+1)  + ')' : 'NOT FOUND') + '\n' +
    '  Control Desc      : ' + (C.name       >= 0 ? headers[C.name]       + ' (col ' + (C.name+1)       + ')' : 'NOT FOUND') + '\n' +
    '  Test Plan         : ' + (C.objective  >= 0 ? headers[C.objective]  + ' (col ' + (C.objective+1)  + ')' : 'NOT FOUND') + '\n' +
    '  Testing Procedures: ' + (C.procedures >= 0 ? headers[C.procedures] + ' (col ' + (C.procedures+1) + ')' : 'NOT FOUND') + '\n' +
    '  Testing Period    : ' + (C.period      >= 0 ? headers[C.period]      + ' (col ' + (C.period+1)      + ')' : 'NOT FOUND') + '\n' +
    '  Publishing Date   : ' + (C.publishDate >= 0 ? headers[C.publishDate] + ' (col ' + (C.publishDate+1) + ')' : 'NOT FOUND') + '\n' +
    '  Run?              : ' + (C.run         >= 0 ? headers[C.run]         + ' (col ' + (C.run+1)         + ')' : 'NOT FOUND') + '\n' +
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

// -- Test File Reading diagnostic ------------------------------------------------
// Lets you pick a Control ID and see exactly what happens when the script tries
// to read the evidence files — including any Drive API errors.

function testFileReading() {
  const props    = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('SOX_DRIVE_FOLDER_ID');
  if (!folderId) {
    SpreadsheetApp.getUi().alert('Drive folder ID not set. Run Configure first.');
    return;
  }

  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt('Test File Reading',
    'Enter Control ID to test (e.g. IL.FSCP.04):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const controlId = res.getResponseText().trim();
  if (!controlId) return;

  let msg = 'Testing: "' + controlId + '"\n';
  msg += '=====================================\n\n';

  try {
    const root = DriveApp.getFolderById(folderId);
    msg += 'Root folder: ' + root.getName() + '\n\n';

    // Find matching subfolder (case-insensitive)
    const norm = controlId.toLowerCase();
    let ctrlFolder = null;
    const subs = root.getFolders();
    const allNames = [];
    while (subs.hasNext()) {
      const f = subs.next();
      allNames.push(f.getName());
      if (f.getName().trim().toLowerCase() === norm) ctrlFolder = f;
    }
    if (!ctrlFolder) {
      msg += 'ERROR: No folder matching "' + controlId + '"\n';
      msg += 'Drive has: ' + allNames.join(', ');
      SpreadsheetApp.getUi().alert(msg);
      return;
    }
    msg += 'Folder found: "' + ctrlFolder.getName() + '"\n\n';

    // Scan for an "input" subfolder, otherwise use ctrlFolder directly
    // Also scan any other subfolders
    const foldersToScan = [];
    const subFolders = ctrlFolder.getFolders();
    while (subFolders.hasNext()) {
      foldersToScan.push(subFolders.next());
    }
    if (foldersToScan.length === 0) {
      msg += 'No subfolders — reading files directly from control folder\n\n';
      foldersToScan.push(ctrlFolder); // treat ctrlFolder itself
    } else {
      // Also include the root control folder in case files are there too
      foldersToScan.unshift(ctrlFolder);
      msg += 'Subfolders found: ' + foldersToScan.slice(1).map(f => f.getName()).join(', ') + '\n\n';
    }

    let totalFiles = 0;
    for (const folder of foldersToScan) {
      const inFolder = folder.getId() === ctrlFolder.getId()
        ? 'Control folder root' : 'Subfolder: ' + folder.getName();
      const fileIter = folder.getFiles();
      const filesHere = [];
      while (fileIter.hasNext()) filesHere.push(fileIter.next());

      if (filesHere.length === 0) {
        msg += '[' + inFolder + ']: no files\n';
        continue;
      }

      msg += '[' + inFolder + ']: ' + filesHere.length + ' file(s)\n';
      for (const file of filesHere) {
        totalFiles++;
        const name     = file.getName();
        const mime     = file.getMimeType();
        const kb       = (file.getSize() / 1024).toFixed(1);
        msg += '\n  File: ' + name + '\n';
        msg += '  MIME: ' + mime + '\n';
        msg += '  Size: ' + kb + ' KB\n';

        if (file.getSize() > 8 * 1024 * 1024) {
          msg += '  STATUS: SKIPPED - too large (> 8 MB)\n';
          continue;
        }

        const isXlsx = mime === 'application/vnd.google-apps.spreadsheet' ||
                       mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                       mime === 'application/vnd.ms-excel' || name.match(/\.(xlsx|xls)$/i);
        const isDoc  = mime === 'application/vnd.google-apps.document' ||
                       mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
                       mime === 'application/msword' || name.match(/\.(docx|doc)$/i);

        if (isXlsx) {
          if (mime === 'application/vnd.google-apps.spreadsheet') {
            msg += '  Type: Native Google Sheet\n';
            try {
              const ss = SpreadsheetApp.openById(file.getId());
              const sheets = ss.getSheets();
              msg += '  Tabs: ' + sheets.map(s => s.getName() + '(' + s.getLastRow() + ' rows)').join(', ') + '\n';
              msg += '  STATUS: OK - readable directly\n';
            } catch(e) {
              msg += '  STATUS: FAILED to open - ' + e.message + '\n';
            }
          } else {
            msg += '  Type: XLSX/XLS - needs Drive copy conversion\n';
            msg += '  Token prefix: ' + ScriptApp.getOAuthToken().substring(0,20) + '...\n';
            try {
              const copyResp = UrlFetchApp.fetch(
                'https://www.googleapis.com/drive/v3/files/' + file.getId() + '/copy?supportsAllDrives=true',
                {
                  method: 'post',
                  headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
                  contentType: 'application/json',
                  payload: JSON.stringify({
                    name: '_sox_diag_' + file.getId(),
                    mimeType: 'application/vnd.google-apps.spreadsheet',
                  }),
                  muteHttpExceptions: true,
                }
              );
              const httpCode = copyResp.getResponseCode();
              msg += '  Drive copy HTTP: ' + httpCode + '\n';
              if (httpCode === 200) {
                const tempId = JSON.parse(copyResp.getContentText()).id;
                msg += '  Temp file created: ' + tempId + '\n';
                Utilities.sleep(4000);
                try {
                  const ss    = SpreadsheetApp.openById(tempId);
                  const tabs  = ss.getSheets();
                  msg += '  Tabs: ' + tabs.map(s => s.getName() + '(' + s.getLastRow() + 'r)').join(', ') + '\n';
                  msg += '  STATUS: OK - conversion works\n';
                } catch(oe) {
                  msg += '  STATUS: FAILED to open copy - ' + oe.message + '\n';
                }
                try { DriveApp.getFileById(tempId).setTrashed(true); } catch(te) {}
              } else {
                msg += '  Copy error body: ' + copyResp.getContentText().slice(0, 250) + '\n';
                msg += '  STATUS: COPY FAILED\n';
              }
            } catch(ce) {
              msg += '  STATUS: EXCEPTION - ' + ce.message + '\n';
            }
          }
        } else if (isDoc) {
          msg += '  Type: Word/Doc\n';
          try {
            DocumentApp.openById(file.getId());
            msg += '  STATUS: OK - readable\n';
          } catch(e) {
            msg += '  STATUS: FAILED - ' + e.message + '\n';
          }
        } else if (mime === 'application/pdf') {
          msg += '  Type: PDF - sent directly to Gemini\n  STATUS: OK\n';
        } else if (mime.startsWith('image/')) {
          msg += '  Type: Image\n  STATUS: OK\n';
        } else if (mime === 'text/plain') {
          msg += '  Type: Text\n  STATUS: OK\n';
        } else {
          msg += '  STATUS: UNSUPPORTED MIME - would be skipped\n';
        }
      }
      msg += '\n';
    }

    if (totalFiles === 0) {
      msg += '\nNO FILES FOUND IN ANY FOLDER!\n';
    } else {
      msg += '\nTotal files scanned: ' + totalFiles + '\n';
    }

  } catch(e) {
    msg += 'OUTER ERROR: ' + e.message + '\n';
  }

  // Show in chunks if too long (Apps Script alert has ~4000 char limit)
  const CHUNK = 1800;
  for (let i = 0; i < msg.length; i += CHUNK) {
    SpreadsheetApp.getUi().alert(msg.slice(i, i + CHUNK));
  }
}

// -- About -----------------------------------------------------------------------

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
    // "Test plan" column — high-level description of what should be tested
    objective  : findCol('objective') >= 0 ? findCol('objective') :
                 (findCol('test plan') >= 0 ? findCol('test plan') : findCol('description')),
    // "Testing procedures Q1, 2026" — specific quarterly testing steps (sent separately to Gemini)
    procedures : findCol('procedure') >= 0 ? findCol('procedure') :
                 (findCol('testing procedure') >= 0 ? findCol('testing procedure') :
                 (findCol('testing') >= 0 ? findCol('testing') : -1)),
    // "Testing period" — the date range evidence must fall within
    period      : findCol('testing period') >= 0 ? findCol('testing period') :
                  (findCol('test period') >= 0 ? findCol('test period') :
                  (findCol('period') >= 0 ? findCol('period') : -1)),
    // "Publishing report date" — review/sign-off must be on or before this date
    publishDate : findCol('publishing report') >= 0 ? findCol('publishing report') :
                  (findCol('publish') >= 0 ? findCol('publish') : -1),
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

  // -- Run History tab -------------------------------------------------------
  const RUN_HISTORY_TAB = 'Run History';
  const runnerEmail = Session.getActiveUser().getEmail() || 'unknown';
  let historySheet = ss.getSheetByName(RUN_HISTORY_TAB);
  if (!historySheet) {
    historySheet = ss.insertSheet(RUN_HISTORY_TAB);
    historySheet.appendRow(['Run Date', 'Control ID', 'Control Name', 'Verdict', 'Gaps / Notes', 'Run By']);
    historySheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    historySheet.setFrozenRows(1);
    historySheet.setColumnWidth(1, 160);
    historySheet.setColumnWidth(4, 80);
    historySheet.setColumnWidth(5, 400);
    historySheet.setColumnWidth(6, 200);
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
    const objective   = C.objective   >= 0 ? String(row[C.objective]   || '').trim() : '';
    const procedures  = C.procedures  >= 0 ? String(row[C.procedures]  || '').trim() : '';
    const period      = C.period      >= 0 ? String(row[C.period]      || '').trim() : '';
    const publishDate = C.publishDate >= 0 ? String(row[C.publishDate] || '').trim() : '';

    if (!controlId) continue;

    // Toast notification - visible in the sheet while running
    ss.toast('Analyzing ' + controlId + ' (' + ctrlName + ')...', 'SOX Agent', -1);

    // Get evidence files
    const evidence = getEvidenceFiles(folderId, controlId);

    // Call Gemini
    let result;
    try {
      result = analyzeControl(apiKey, controlId, ctrlName, objective, procedures, period, publishDate, evidence);
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

    const runTime = new Date();

    // Write timestamp to main sheet
    if (C.lastRun >= 0) {
      sheet.getRange(rowNum, C.lastRun + 1).setValue(runTime.toLocaleString());
    }

    // Append to Run History tab
    const gapSummary = result.gaps && result.gaps.length
      ? result.gaps.map(g => '- ' + g).join('\n')
      : (result.summary || '');
    historySheet.appendRow([
      runTime,
      runnerEmail,
      controlId,
      ctrlName,
      period,
      result.verdict,
      gapSummary,
    ]);
    const newRowNum = historySheet.getLastRow();
    historySheet.getRange(newRowNum, 6).setBackground(
      result.verdict === 'Passed' ? '#d9ead3' : '#fce5cd'
    );

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

    // Collect files from the control folder AND all its subfolders.
    // This handles: files in root, files in "input/", files in any other subfolder.
    const foldersToScan = [ctrlFolder];
    const subIter = ctrlFolder.getFolders();
    while (subIter.hasNext()) foldersToScan.push(subIter.next());

    const allFiles = [];
    for (const folder of foldersToScan) {
      const fi = folder.getFiles();
      while (fi.hasNext()) allFiles.push(fi.next());
    }
    Logger.log(controlId + ': found ' + allFiles.length + ' files across ' + foldersToScan.length + ' folder(s)');

    for (const file of allFiles) {
      const name     = file.getName();
      const mimeType = file.getMimeType();
      const sizeBytes = file.getSize();

      // Skip very large files (>8 MB - Gemini inline limit)
      if (sizeBytes > 8 * 1024 * 1024) {
        Logger.log('Skipping large file: ' + name + ' (' + sizeBytes + ' bytes)');
        evidence.push({
          name: name + '_note.txt',
          mimeType: 'text/plain',
          data: Utilities.base64Encode(Utilities.newBlob(
            'Evidence file present but too large to send (> 8 MB): ' + name +
            ' (' + (sizeBytes / 1024 / 1024).toFixed(1) + ' MB)',
            'text/plain; charset=utf-8').getBytes()),
        });
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
              'https://www.googleapis.com/drive/v3/files/' + ssId + '/copy?supportsAllDrives=true',
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
              const errMsg = 'Drive copy failed for "' + name + '": HTTP ' +
                             copyResp.getResponseCode() + ' — ' +
                             copyResp.getContentText().slice(0, 300);
              Logger.log(errMsg);
              // Still push a text note so Gemini knows the file exists
              evidence.push({
                name: name + '_note.txt',
                mimeType: 'text/plain',
                data: Utilities.base64Encode(
                  Utilities.newBlob(
                    'Evidence file present but could not be converted: ' + name + '\n' + errMsg,
                    'text/plain; charset=utf-8'
                  ).getBytes()
                ),
              });
              continue;
            }
            tempId = JSON.parse(copyResp.getContentText()).id;
            ssId = tempId;
            // Wait for Drive to finish creating the converted copy before opening it
            Utilities.sleep(4000);
          }

          const ss = SpreadsheetApp.openById(ssId);
          const sheets = ss.getSheets();

          // Budget: 300K chars split evenly across sheets (at least 50K per sheet)
          const TOTAL_BUDGET = 300000;
          const perSheetBudget = Math.max(50000, Math.floor(TOTAL_BUDGET / sheets.length));

          let csv = '';
          sheets.forEach(function(s) {
            const lastRow = s.getLastRow();
            const lastCol = s.getLastColumn();
            csv += '=== Sheet: ' + s.getName() + ' ===\n';
            if (lastRow < 1 || lastCol < 1) {
              csv += '(empty sheet)\n\n';
              return;
            }

            const allRows = s.getRange(1, 1, lastRow, lastCol).getValues();

            function rowToCsv(row) {
              return row.map(function(c) {
                const v = String(c === null || c === undefined ? '' : c);
                return v.indexOf(',') >= 0 || v.indexOf('"') >= 0
                  ? '"' + v.replace(/"/g, '""') + '"' : v;
              }).join(',') + '\n';
            }

            const headerLine = rowToCsv(allRows[0]);  // always include header row
            const dataRows   = allRows.slice(1);       // remaining rows

            // Build CSV within budget: header + as many rows as fit, preferring the most recent
            let sheetCsv = headerLine;
            let budget = perSheetBudget - headerLine.length;
            const included = [];
            // Walk from the end (most recent) back
            for (let r = dataRows.length - 1; r >= 0 && budget > 0; r--) {
              const line = rowToCsv(dataRows[r]);
              if (line.length <= budget) {
                included.unshift(line);   // prepend to keep chronological order
                budget -= line.length;
              }
            }
            const skipped = dataRows.length - included.length;
            if (skipped > 0) {
              sheetCsv += '[' + skipped + ' older rows omitted — showing the ' +
                           included.length + ' most recent rows]\n';
            }
            sheetCsv += included.join('');
            csv += sheetCsv + '\n';
          });

          if (tempId) {
            try { DriveApp.getFileById(tempId).setTrashed(true); } catch(te) {}
          }

          if (csv.trim().length > 0) {
            const finalCsv = csv;
            Logger.log('Converted spreadsheet: ' + name + ' (' + sheets.length + ' tabs, ' + csv.length + ' chars, budget ' + TOTAL_BUDGET + ')');
            evidence.push({
              name: name + '.txt',
              mimeType: 'text/plain',
              data: Utilities.base64Encode(Utilities.newBlob(finalCsv, 'text/plain; charset=utf-8').getBytes()),
            });
          }
        } catch (e) {
          Logger.log('Failed to convert spreadsheet ' + name + ': ' + e.message);
          // Push error note so Gemini knows the file exists
          evidence.push({
            name: name + '_note.txt',
            mimeType: 'text/plain',
            data: Utilities.base64Encode(
              Utilities.newBlob(
                'Evidence file present but conversion failed: ' + name + '\nError: ' + e.message,
                'text/plain; charset=utf-8'
              ).getBytes()
            ),
          });
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

function analyzeControl(apiKey, controlId, controlName, controlObjective, testingProcedures, testingPeriod, publishingReportDate, evidence) {
  const GEMINI_MODEL = 'gemini-2.5-pro';
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
              GEMINI_MODEL + ':generateContent?key=' + apiKey;

  const publishLine = publishingReportDate
    ? 'Publishing Report Date: ' + publishingReportDate + '\n'
    : '';

  const promptText =
    'You are a strict SOX (Sarbanes-Oxley) compliance auditor.\n\n' +
    'CONTROL DETAILS\n' +
    'Control ID:             ' + controlId + '\n' +
    'Control Description:    ' + (controlName      || '(not specified)') + '\n' +
    'Test Plan:              ' + (controlObjective || '(not specified)') + '\n' +
    'Testing Period:         ' + (testingPeriod    || '(not specified)') + '\n' +
    publishLine +
    'Testing Procedures:     ' + (testingProcedures || '(not specified)') + '\n\n' +

    'DATE RULES\n' +
    '  - CONTROL EVIDENCE (transactions, system logs, access lists, approvals of transactions):\n' +
    '    Must fall within the Testing Period. Evidence outside this period does NOT count.\n' +
    (publishingReportDate
      ? '  - WORKING PAPER REVIEW / AUDITOR SIGN-OFF / SUPERVISOR APPROVAL OF THE TEST:\n' +
        '    Must be on or before the Publishing Report Date (' + publishingReportDate + ').\n' +
        '    A review date after the Publishing Report Date is a gap.\n' +
        '    A review date between the Testing Period end and the Publishing Report Date is acceptable.\n'
      : '  - WORKING PAPER REVIEW / AUDITOR SIGN-OFF / SUPERVISOR APPROVAL OF THE TEST:\n' +
        '    These are expected after the period ends. Dates within a reasonable window after the\n' +
        '    Testing Period end are acceptable and must NOT be flagged as out-of-period.\n') +
    '  - If you are unsure whether a date is control-execution or review/sign-off,\n' +
    '    give the benefit of the doubt.\n\n' +

    'PASS CRITERIA — verdict must be "Passed" ONLY if ALL of the following are true:\n' +
    '  1. Every step in "Testing Procedures" has clear, supporting evidence.\n' +
    '  2. All control-execution evidence dates fall within the Testing Period.\n' +
    (publishingReportDate
      ? '  3. Review/sign-off date is on or before ' + publishingReportDate + '.\n'
      : '  3. Review/sign-off dates fall within a reasonable window after the Testing Period.\n') +
    '  4. No required procedure step is missing or has only partial evidence.\n' +
    'If ANY criterion is not met → verdict must be "Failed".\n\n' +

    'INSTRUCTIONS\n' +
    '1. Read each evidence file carefully, noting all dates.\n' +
    '2. Distinguish control-execution dates from review/sign-off dates.\n' +
    '3. For each step in "Testing Procedures", verify evidence exists and dates are in range.\n' +
    '4. List every gap as a separate entry. Be specific: name the step and reason\n' +
    '   (e.g. "Step 2: no approval log found for Q1 2026").\n' +
    '5. Do NOT flag review/sign-off dates as gaps if they are within the allowed window.\n\n' +

    'RESPONSE FORMAT — respond with valid JSON only, no markdown fences:\n' +
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

  // Cap total evidence at ~600K chars (~150K tokens) to stay well within Gemini's 1M token limit
  const MAX_TOTAL_CHARS = 600000;
  let totalChars = 0;

  for (const ef of evidence) {
    if (!GEMINI_OK[ef.mimeType]) {
      Logger.log('Blocked unsupported MIME before Gemini call: ' + ef.mimeType + ' (' + ef.name + ')');
      continue;
    }
    // Decode to check size, truncate if adding this file would exceed the total cap
    const decoded = Utilities.newBlob(Utilities.base64Decode(ef.data)).getDataAsString();
    if (totalChars + decoded.length > MAX_TOTAL_CHARS) {
      const remaining = MAX_TOTAL_CHARS - totalChars;
      if (remaining > 500) {
        const truncated = decoded.slice(0, remaining) +
          '\n\n[TRUNCATED: total evidence size limit reached — ' +
          Math.round((decoded.length - remaining) / 1000) + 'K chars omitted from this file]\n';
        parts.push({ inlineData: {
          mimeType: 'text/plain',
          data: Utilities.base64Encode(Utilities.newBlob(truncated, 'text/plain; charset=utf-8').getBytes()),
        }});
        Logger.log('Evidence total cap reached, partially included: ' + ef.name);
      } else {
        Logger.log('Evidence total cap reached, skipped: ' + ef.name);
      }
      break;
    }
    totalChars += decoded.length;
    parts.push({ inlineData: { mimeType: ef.mimeType, data: ef.data } });
  }
  Logger.log('Total evidence sent to Gemini: ' + totalChars + ' chars across ' + (parts.length - 1) + ' file(s)');

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
