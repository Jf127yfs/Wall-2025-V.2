

// ===== File: Pan_Analytics =====


/**
 * ============================================================================
 * PAN_ANALYTICS.GS - UNIFIED Analytics System
 * ============================================================================
 * 
 * PURPOSE:
 * - Orchestrates all analytics workflows (cleaning, DDD, distributions, correlations)
 * - Builds Pan_Master and Pan_Dict for statistical analysis
 * - Provides unified menu system for all analytics functions
 * 
 * SECTIONS:
 * 1. Configuration
 * 2. Menu System
 * 3. Orchestrator Functions (runAllAnalytics, etc.)
 * 4. Pan_Master Builder (buildPanSheets, buildMaster_, etc.)
 * 5. Utility Functions
 * ============================================================================
 */



// ============================================================================

// SECTION 1: CONFIGURATION

// ============================================================================



// Source and output sheets

const PA_SOURCE_SHEET = 'Form Responses 1';

const PA_CLEAN_SHEET = 'Form Responses (Clean)';

const PA_DISTS_SHEET = 'Distributions (Clean)';

const PAN_DICT_SHEET = 'Pan_Dict';

const PAN_MASTER_SHEET = 'Pan_Master';



// Required headers for validation

const PA_REQUIRED_HEADERS = [

  'Current 5 Digit Zip Code',

  'Self-Identified Gender',

  'Birthday (MM/DD)',

  'Screen Name',

  'UID',

  'Checked-In at Event',

  'Check-In Timestamp',

  'Photo URL'

];



// Headers for distribution analysis

const PA_DIST_HEADERS = [

  'Zodiac Sign',

  'Age Range',

  'Education Level',

  'Self Identified Ethnicity',

  'Self-Identified Gender',

  'Self-Identified Sexual Orientation',

  'Employment Information (Industry)',

  'Employment Information (Role)',

  'Do you know the Host(s)?',

  'Which host have you known the longest?',

  'Music Preference',

  'Which best describes your general social stance?'

];



// Variable specification for Pan_Master

const SPEC = [

  { key: 'timestamp', header: 'Timestamp', type: 'timestamp' },

  { key: 'birthday', header: 'Birthday (MM/DD)', type: 'birthday' },

  { key: 'zodiac', header: 'Zodiac Sign', type: 'single', 

    opts: ['Aries','Taurus','Gemini','Cancer','Leo','Virgo','Libra','Scorpio','Sagittarius','Capricorn','Aquarius','Pisces'] },

  { key: 'age_range', header: 'Age Range', type: 'single', 

    opts: ['21-24','25-29','30-34','35-39','40-44','45-49','50+'] },

  { key: 'education', header: 'Education Level', type: 'single', 

    opts: ['High School','Some College','Associates','Bachelors','Masters & Above'] },

  { key: 'zip', header: 'Current 5 Digit Zip Code', type: 'text_raw' },

  { key: 'ethnicity', header: 'Self Identified Ethnicity', type: 'single', 

    opts: ['Black / African American','Mixed / Multiracial','White','Asian / Pacific Islander','Hispanic / Latino','Native American','Not Listed','Prefer Not to Say'] },

  { key: 'gender', header: 'Self-Identified Gender', type: 'single', 

    opts: ['Man','Woman','Non-Binary','Other'] },

  { key: 'orientation', header: 'Self-Identified Sexual Orientation', type: 'single', 

    opts: ['Straight / Heterosexual','Bisexual','Gay','Lesbian','Pansexual','Queer','Asexual','Other'] },

  { key: 'industry', header: 'Employment Information (Industry)', type: 'single', 

    opts: ['Arts & Entertainment','Education','Finance / Business Services','Government / Military','Healthcare','Hospitality / Retail','Science / Research','Technology','Trades / Manufacturing','Other'] },

  { key: 'role', header: 'Employment Information (Role)', type: 'single', 

    opts: ['Creative / Designer / Artist','Educator / Instructor','Founder / Entrepreneur','Healthcare / Service Provider','Manager / Supervisor','Operations / Admin / Support','Researcher / Scientist','Sales / Marketing / Business Development','Student / Trainee','Technical / Engineer / Developer','Trades / Skilled Labor','Other'] },

  { key: 'know_hosts', header: 'Do you know the Host(s)?', type: 'single', 

    opts: ['No','Yes — less than 3 months','Yes — 3–12 months','Yes — 1–3 years','Yes — 3–5 years','Yes — 5–10 years','Yes — more than 10 years'] },

  { key: 'known_longest', header: 'Which host have you known the longest?', type: 'single', 

    opts: ['Jacob','Michael','Equal','Do Not Know Them'] },

  { key: 'know_score', header: 'If yes, how well do you know them?', type: 'number' },

  { key: 'interests', header: 'Your General Interests (Choose 3)', type: 'multi', 

    opts: ['Cooking','Music','Fashion','Travel','Fitness','Gaming','Reading','Art/Design','Photography','Hiking/Outdoors','Sports (general)','Volunteering','Health Sciences','TikTok (watching)','Other'] },

  { key: 'music_pref', header: 'Music Preference', type: 'single', 

    opts: ['Hip-hop','Pop','Indie/Alt','R&B','Rock','Country','Electronic','Jazz','Classical','A mix of all','Other'] },

  { key: 'fav_artist', header: 'Current Favorite Artist', type: 'text' },

  { key: 'song', header: 'Name one song you want to hear at the party.', type: 'text' },

  { key: 'recent_purchase', header: 'Recent purchase you're most happy about', type: 'single', 

    opts: ['Fashion/Clothing','Fitness gear','Tech gadget','Car/Motorcycle','Home/Kitchen','Pet item','Course/App','Other'] },

  { key: 'at_worst', header: 'At your worst you are…', type: 'single', 

    opts: ['Anxious','Distracted','Guarded','Impulsive','Jealous','Overly critical','Reckless','Self-conscious','Stubborn','Other'] },

  { key: 'social_stance', header: 'Which best describes your general social stance?', type: 'number' }

];



// ============================================================================

// SECTION 2: MENU SYSTEM

// ============================================================================



function mergeAnalyticsMenu(menu) {
  return menu.addSubMenu(SpreadsheetApp.getUi().createMenu('📊 Analytics')
      .addItem('Run All Analytics', 'runAllAnalytics')
      .addSeparator()
      .addItem('Build Pan Sheets', 'buildPanSheets')
      .addItem('Build Distributions', 'buildDistributions')
      .addItem('Build Cramers V', 'buildCramersVClean')
      .addSeparator()
      .addItem('Run DDD (All)', 'runDDDAll')
      .addItem('Run DDD (Checked-In Only)', 'runDDDCheckedIn'))
    .addSeparator()
    .addItem('🔄 Check Setup', 'checkSetup')
    .addItem('🌐 Open Display', 'testOpenDisplay')
    .addItem('📸 Fix Photos', 'testPhotoLoading');
}



// ============================================================================

// SECTION 3: ORCHESTRATOR FUNCTIONS

// ============================================================================



/**

 * Runs the complete analytics workflow end-to-end

 * WARNING: This runs 5 heavy processes and may take 2-5 minutes

 * 

 * Process:

 * 1. Clean data (cleanFormResponses)

 * 2. DDD analysis on all responses

 * 3. DDD analysis on checked-in guests only

 * 4. Build frequency distributions

 * 5. Calculate Cramér's V correlation matrix

 */

function runAllAnalytics() {

  var ui = SpreadsheetApp.getUi();

  

  // Confirmation dialog

  var response = ui.alert(

    'Run All Analytics',

    'This will run the complete analytics workflow:

' +

    '1. Clean data
' +

    '2. DDD analysis (all)
' +

    '3. DDD analysis (checked-in)
' +

    '4. Build distributions
' +

    '5. Build Cramér's V matrix

' +

    'This may take 2-5 minutes. Continue?',

    ui.ButtonSet.YES_NO

  );

  

  if (response !== ui.Button.YES) return;

  

  try {

    var ss = SpreadsheetApp.getActive();

    

    // 1) Clean data

    ss.toast('Step 1/5: Cleaning data...', 'Analytics', 10);

    if (typeof cleanFormResponses === 'function') {

      cleanFormResponses();

    } else {

      Logger.log('cleanFormResponses not found, skipping');

    }



    // 2) DDD over all cleaned rows

    ss.toast('Step 2/5: Running DDD (all responses)...', 'Analytics', 10);

    if (typeof detectDataDisruptors === 'function') {

      detectDataDisruptors();

    } else {

      Logger.log('detectDataDisruptors not found, skipping');

    }



    // 3) DDD over checked-in guests only

    ss.toast('Step 3/5: Running DDD (checked-in only)...', 'Analytics', 10);

    if (typeof detectDataDisruptorsForCheckedIn === 'function') {

      detectDataDisruptorsForCheckedIn();

    } else {

      Logger.log('detectDataDisruptorsForCheckedIn not found, skipping');

    }



    // 4) Distributions

    ss.toast('Step 4/5: Building distributions...', 'Analytics', 10);

    buildDistributions();



    // 5) Cramér's V

    ss.toast('Step 5/5: Building Cramers V matrix...', 'Analytics', 10);

    if (typeof buildCramersVClean === 'function') {

      buildCramersVClean();

    } else if (typeof buildVCramers === 'function') {

      buildVCramers();

    } else {

      Logger.log('No Cramers V function found, skipping');

    }



    ui.alert(

      '✅ Analytics Complete!',

      'All analytics have been generated:

' +

      '✓ Data cleaned
' +

      '✓ DDD reports
' +

      '✓ Distributions
' +

      '✓ Cramér's V matrix

' +

      'Check the respective sheets for results.',

      ui.ButtonSet.OK

    );

  } catch (e) {

    ui.alert('Error', 'Error running analytics:\n' + e.toString(), ui.ButtonSet.OK);

    Logger.log('Error in runAllAnalytics: ' + e);

  }

}



function runDDDAll() {

  if (typeof detectDataDisruptors !== 'function') {

    SpreadsheetApp.getUi().alert('detectDataDisruptors() not found. Ensure DDD file is present.');

    return;

  }

  detectDataDisruptors();

}



function runDDDCheckedIn() {

  if (typeof detectDataDisruptorsForCheckedIn !== 'function') {

    SpreadsheetApp.getUi().alert('detectDataDisruptorsForCheckedIn() not found. Ensure Automation file is present.');

    return;

  }

  detectDataDisruptorsForCheckedIn();

}



/**

 * Builds frequency distributions for categorical variables

 * Creates a summary sheet with counts for each category

 */

function buildDistributions() {

  var ss = SpreadsheetApp.getActive();

  var sh = ss.getSheetByName(PA_CLEAN_SHEET);

  

  if (!sh) {

    SpreadsheetApp.getUi().alert('Clean sheet "' + PA_CLEAN_SHEET + '" not found. Run cleanFormResponses() first.');

    return;

  }



  var data = sh.getDataRange().getValues();

  if (data.length < 2) {

    writeDistSheet_([]);

    return;

  }



  var headers = data[0].map(function(h) { return String(h || '').trim(); });

  var rows = data.slice(1);



  var sections = [];

  for (var i = 0; i < PA_DIST_HEADERS.length; i++) {

    var label = PA_DIST_HEADERS[i];

    var col = headers.indexOf(label);

    

    if (col === -1) {

      sections.push({ label: label, rows: [['(missing column)', 0]] });

      continue;

    }

    

    var counts = {};

    for (var r = 0; r < rows.length; r++) {

      var v = normCat_(rows[r][col]);

      if (!v) continue;

      counts[v] = (counts[v] || 0) + 1;

    }

    

    var table = [];

    for (var key in counts) {

      if (counts.hasOwnProperty(key)) {

        table.push([String(key), Number(counts[key])]);

      }

    }

    table.sort(function(a, b) { return b[1] - a[1]; });

    

    sections.push({ 

      label: label, 

      rows: table.length > 0 ? table : [['(no data)', 0]] 

    });

  }



  writeDistSheet_(sections);

  SpreadsheetApp.getActive().toast('Distributions built successfully', 'Success', 3);

}



function writeDistSheet_(sections) {

  var ss = SpreadsheetApp.getActive();

  var out = ss.getSheetByName(PA_DISTS_SHEET);

  if (!out) out = ss.insertSheet(PA_DISTS_SHEET);

  out.clear();



  if (!sections || sections.length === 0) {

    out.getRange(1, 1).setValue('No distributions available.');

    return;

  }



  var row = 1;

  for (var i = 0; i < sections.length; i++) {

    var sec = sections[i];

    

    // Header row - set value first, then format

    out.getRange(row, 1).setValue(sec.label);

    out.getRange(row, 1, 1, 2).mergeAcross().setFontWeight('bold').setBackground('#e3f2fd');

    row++;

    

    // Column labels

    var labelRange = out.getRange(row, 1, 1, 2);

    labelRange.setValues([['Value', 'Count']]);

    labelRange.setFontWeight('bold');

    row++;

    

    // Data rows - ensure proper 2D array structure

    if (sec.rows && sec.rows.length > 0) {

      // Verify each row has exactly 2 elements

      var validRows = [];

      for (var j = 0; j < sec.rows.length; j++) {

        var r = sec.rows[j];

        if (r && r.length >= 2) {

          validRows.push([r[0], r[1]]);

        } else if (r && r.length === 1) {

          validRows.push([r[0], '']);

        } else {

          validRows.push(['(invalid)', 0]);

        }

      }

      

      if (validRows.length > 0) {

        out.getRange(row, 1, validRows.length, 2).setValues(validRows);

        row += validRows.length;

      }

    }

    

    // Spacer

    row++;

  }



  out.autoResizeColumns(1, 2);

}



/**

 * Verifies sheets, columns, and functions exist

 */

function checkSetup() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

  var messages = [];



  // Check sheets

  var source = ss.getSheetByName(PA_SOURCE_SHEET);

  var clean = ss.getSheetByName(PA_CLEAN_SHEET);

  messages.push(source ? 'OK: Found source sheet "' + PA_SOURCE_SHEET + '"' : 'MISSING: Sheet "' + PA_SOURCE_SHEET + '"');

  messages.push(clean ? 'OK: Found clean sheet "' + PA_CLEAN_SHEET + '"' : 'MISSING: Sheet "' + PA_CLEAN_SHEET + '"');



  // Check columns in cleaned sheet

  if (clean) {

    try {

      var headers = clean.getRange(1, 1, 1, clean.getLastColumn()).getValues()[0];

      var headerSet = {};

      for (var i = 0; i < headers.length; i++) {

        headerSet[String(headers[i] || '').trim()] = true;

      }

      

      for (var j = 0; j < PA_REQUIRED_HEADERS.length; j++) {

        var h = PA_REQUIRED_HEADERS[j];

        messages.push(headerSet[h] ? 'OK: Column present - ' + h : 'MISSING: Column - ' + h);

      }

    } catch (e) {

      messages.push('ERROR reading headers: ' + e.message);

    }

  }



  // Check functions

  messages.push(typeof cleanFormResponses === 'function' ? 'OK: cleanFormResponses()' : 'MISSING: cleanFormResponses()');

  messages.push(typeof detectDataDisruptors === 'function' ? 'OK: detectDataDisruptors()' : 'MISSING: detectDataDisruptors()');

  messages.push(typeof detectDataDisruptorsForCheckedIn === 'function' ? 'OK: detectDataDisruptorsForCheckedIn()' : 'MISSING: detectDataDisruptorsForCheckedIn()');

  messages.push(typeof buildCramersVClean === 'function' || typeof buildVCramers === 'function' ? 'OK: Cramers V function' : 'MISSING: Cramers V function');



  Logger.log('Setup check:
' + messages.join('
'));

  ui.alert('Setup Check', messages.join('
'), ui.ButtonSet.OK);

}



// ============================================================================

// SECTION 4: PAN_MASTER BUILDER

// ============================================================================



/**

 * Builds Pan_Master and Pan_Dict from cleaned data

 * Pan_Master: Analysis dataset with encoded categorical variables

 * Pan_Dict: Data dictionary mapping values to codes

 */

function buildPanSheets() {

  var ss = SpreadsheetApp.getActive();

  var src = ss.getSheetByName(PA_CLEAN_SHEET);

  

  if (!src) {

    throw new Error('Source sheet "' + PA_CLEAN_SHEET + '" not found.');

  }



  var values = src.getDataRange().getValues();

  if (values.length < 2) {

    clearOrCreate_(PAN_DICT_SHEET);

    clearOrCreate_(PAN_MASTER_SHEET);

    SpreadsheetApp.getUi().alert('No data found in ' + PA_CLEAN_SHEET);

    return;

  }



  var headers = values[0];

  var idx = indexByHeader_(headers);

  var dictRows = buildDictRows_(idx);

  

  writeSheet_(PAN_DICT_SHEET, [['Key','Header','Type','Option','Code','Note']].concat(dictRows));

  

  var master = buildMaster_(values.slice(1), idx);

  writeSheet_(PAN_MASTER_SHEET, master);

  

  SpreadsheetApp.getUi().alert(

    'Pan Sheets Built!

' +

    'Pan_Dict: ' + dictRows.length + ' rows
' +

    'Pan_Master: ' + (master.length - 1) + ' guests (' + master[0].length + ' columns)'

  );

}



function buildDictRows_(idx) {

  var out = [];

  

  for (var i = 0; i < SPEC.length; i++) {

    var field = SPEC[i];

    var present = !!idx[norm(field.header)];

    var note = present ? (field.note || '') : 'Header not found';

    

    if (field.type === 'single' || field.type === 'multi') {

      var opts = field.opts || [];

      for (var j = 0; j < opts.length; j++) {

        out.push([field.key, field.header, field.type, opts[j], j + 1, note]);

      }

    } else if (field.type === 'number') {

      out.push([field.key, field.header, 'number', '', '', note]);

    } else if (field.type === 'timestamp') {

      out.push([field.key, field.header, 'timestamp', '', '', note]);

    } else if (field.type === 'birthday') {

      out.push([field.key, field.header, 'birthday', 'MM/DD', 'Birthday_MM/DD', note]);

    } else if (field.type === 'text_raw') {

      out.push([field.key, field.header, 'text', 'raw', '', note]);

    } else {

      out.push([field.key, field.header, 'text', '', '', note]);

    }

  }

  

  return out;

}



function buildMaster_(rows, idx) {

  var header = ['Screen Name', 'UID', 'Row', 'TimestampMs', 'Birthday_MM/DD'];

  

  // Add columns by type

  for (var i = 0; i < SPEC.length; i++) {

    if (SPEC[i].type === 'text_raw') {

      header.push(SPEC[i].key === 'zip' ? 'Zip' : SPEC[i].key);

    }

  }

  

  for (var i = 0; i < SPEC.length; i++) {

    if (SPEC[i].type === 'single') {

      header.push('code_' + SPEC[i].key);

    }

  }

  

  for (var i = 0; i < SPEC.length; i++) {

    if (SPEC[i].type === 'multi') {

      var opts = SPEC[i].opts || [];

      for (var j = 0; j < opts.length; j++) {

        header.push('oh_' + SPEC[i].key + '_' + opts[j]);

      }

    }

  }

  

  for (var i = 0; i < SPEC.length; i++) {

    if (SPEC[i].type === 'number') {

      header.push('code_' + SPEC[i].key);

    }

  }

  

  for (var i = 0; i < SPEC.length; i++) {

    if (SPEC[i].type === 'text') {

      header.push('has_' + SPEC[i].key);

    }

  }

  

  var out = [header];

  

  // Build code maps

  var codeMaps = {};

  for (var i = 0; i < SPEC.length; i++) {

    var f = SPEC[i];

    if (f.type === 'single' || f.type === 'multi') {

      var m = {};

      var opts = f.opts || [];

      for (var j = 0; j < opts.length; j++) {

        m[norm(opts[j])] = j + 1;

      }

      codeMaps[f.key] = m;

    }

  }

  

  // Process each row

  for (var i = 0; i < rows.length; i++) {

    var r = rows[i];

    if (!hasValue_(r[0])) continue;

    

    // NOTE: Form Responses (Clean) already contains only validated guests

    // No need to filter by check-in status here

    

    var outRow = [];

    var rowNum = i + 2;

    

    var snIdx = getIdx_(idx, 'Screen Name');

    var uidIdx = getIdx_(idx, 'UID');

    outRow.push(snIdx ? String(r[snIdx - 1] || '').trim() : '');

    outRow.push(uidIdx ? String(r[uidIdx - 1] || '').trim() : '');

    outRow.push(rowNum);

    

    var tsIdx = getIdx_(idx, 'Timestamp');

    outRow.push(toEpochMs_(tsIdx ? r[tsIdx - 1] : ''));

    

    var bIdx = getIdx_(idx, 'Birthday (MM/DD)');

    outRow.push(toMMDDString_(bIdx ? r[bIdx - 1] : ''));

    

    // Raw text fields

    for (var j = 0; j < SPEC.length; j++) {

      if (SPEC[j].type === 'text_raw') {

        var cIdx = getIdx_(idx, SPEC[j].header);

        outRow.push(String(cIdx ? r[cIdx - 1] || '' : '').trim());

      }

    }

    

    // Single-choice fields

    for (var j = 0; j < SPEC.length; j++) {

      if (SPEC[j].type === 'single') {

        var cIdx = getIdx_(idx, SPEC[j].header);

        var val = cIdx ? r[cIdx - 1] : '';

        var m = codeMaps[SPEC[j].key];

        var s = String(val || '').trim();

        outRow.push(s ? (m && m[norm(s)]) || '' : '');

      }

    }

    

    // Multi-choice fields

    for (var j = 0; j < SPEC.length; j++) {

      if (SPEC[j].type === 'multi') {

        var cIdx = getIdx_(idx, SPEC[j].header);

        var val = cIdx ? r[cIdx - 1] : '';

        var picks = splitMulti_(val);

        var set = {};

        for (var k = 0; k < picks.length; k++) {

          set[norm(picks[k])] = true;

        }

        var opts = SPEC[j].opts || [];

        for (var k = 0; k < opts.length; k++) {

          outRow.push(set[norm(opts[k])] ? 1 : 0);

        }

      }

    }

    

    // Numeric fields

    for (var j = 0; j < SPEC.length; j++) {

      if (SPEC[j].type === 'number') {

        var cIdx = getIdx_(idx, SPEC[j].header);

        var val = cIdx ? r[cIdx - 1] : '';

        var n = toNumber_(val);

        outRow.push(isFinite(n) ? n : '');

      }

    }

    

    // Text presence fields

    for (var j = 0; j < SPEC.length; j++) {

      if (SPEC[j].type === 'text') {

        var cIdx = getIdx_(idx, SPEC[j].header);

        var val = cIdx ? r[cIdx - 1] : '';

        outRow.push(hasText_(val) ? 1 : 0);

      }

    }

    

    out.push(outRow);

  }

  

  return out;

}



// ============================================================================

// SECTION 5: UTILITY FUNCTIONS

// ============================================================================



function norm(s) {

  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');

}



function normCat_(v) {

  if (v == null) return '';

  if (v instanceof Date) return '';

  var s = String(v).trim();

  return s ? s.replace(/\s+/g, ' ') : '';

}



function indexByHeader_(headers) {

  var map = {};

  for (var i = 0; i < headers.length; i++) {

    map[norm(headers[i])] = i + 1;

  }

  return map;

}



function getIdx_(idxMap, header) {

  return idxMap[norm(header)] || 0;

}



function splitMulti_(v) {

  if (v instanceof Array) return v;

  var s = String(v || '').trim();

  if (!s) return [];

  return s.split(/[;,]/).map(function(x) { return x.trim(); }).filter(function(x) { return x; });

}



function toEpochMs_(v) {

  if (v instanceof Date) return v.getTime();

  if (typeof v === 'number' && isFinite(v)) return v;

  if (typeof v === 'string' && v.trim()) {

    var d = new Date(v);

    if (!isNaN(d.getTime())) return d.getTime();

  }

  return '';

}



function toMMDDString_(v) {

  if (v instanceof Date) {

    var m = String(v.getMonth() + 1);

    var d = String(v.getDate());

    return (m.length < 2 ? '0' + m : m) + '/' + (d.length < 2 ? '0' + d : d);

  }

  var s = String(v || '').trim();

  if (!s) return '';

  var parts = s.split(/[\/.\-]/).map(function(p) { return p.trim(); }).filter(function(p) { return p; });

  if (parts.length >= 2) {

    var m = String(parts[0]);

    var d = String(parts[1]);

    return (m.length < 2 ? '0' + m : m) + '/' + (d.length < 2 ? '0' + d : d);

  }

  return '';

}



function toNumber_(v) {

  if (v === null || v === undefined || v === '') return NaN;

  if (typeof v === 'number') return v;

  return Number(String(v).trim());

}



function hasValue_(v) {

  if (v === null || v === undefined) return false;

  if (v instanceof Date) return true;

  if (typeof v === 'number') return true;

  if (typeof v === 'string') return v.trim() !== '';

  return true;

}



function hasText_(v) {

  return typeof v === 'string' ? v.trim().length > 0 : false;

}



function clearOrCreate_(name) {

  var ss = SpreadsheetApp.getActive();

  var sh = ss.getSheetByName(name);

  if (!sh) {

    sh = ss.insertSheet(name);

  } else {

    sh.clear();

  }

  return sh;

}



function writeSheet_(name, values) {

  var ss = SpreadsheetApp.getActive();

  var sh = ss.getSheetByName(name);

  if (!sh) sh = ss.insertSheet(name);

  

  sh.clear();

  if (!values || !values.length) return;

  

  // Verify all rows have the same number of columns

  var expectedCols = values[0].length;

  var validRows = [];

  

  for (var i = 0; i < values.length; i++) {

    var row = values[i];

    if (row.length === expectedCols) {

      validRows.push(row);

    } else {

      // Pad or trim row to match expected columns

      var fixedRow = [];

      for (var j = 0; j < expectedCols; j++) {

        fixedRow.push(j < row.length ? row[j] : '');

      }

      validRows.push(fixedRow);

      Logger.log('Warning: Row ' + i + ' had ' + row.length + ' columns, expected ' + expectedCols);

    }

  }

  

  if (validRows.length === 0) return;

  

  sh.getRange(1, 1, validRows.length, expectedCols).setValues(validRows);

  

  // Format header row

  if (validRows.length > 0) {

    var headerRange = sh.getRange(1, 1, 1, expectedCols);

    headerRange

      .setBackground('#434343')

      .setFontColor('#ffffff')

      .setFontWeight('bold')

      .setHorizontalAlignment('center');

    sh.setFrozenRows(1);

  }

}



// ===== File: Analytics_Report_Master =====


/**

 * ============================================================================

 * ANALYTICS_REPORT_MASTER.GS - Consolidated Guest Label & Report Generator

 * ============================================================================

 * * PURPOSE:

 * 1. Generates and calculates dynamic, fun labels (e.g., 'Old Soul') for guests 

 * in the Pan_Master sheet.

 * 2. Creates a clean, summarized 'Label_Report' sheet for hosts, breaking down 

 * guest demographics by these labels.

 * * WORKFLOW:

 * Reads Pan_Master → Calculates and writes 'Guest Label' column → Creates 

 * and formats the 'Label_Report' sheet.

 * * DEPENDENCIES:

 * - Requires the Pan_Master sheet to be built and present.

 * * AUTHOR: Gemini (for Halloween Party Analytics Team)

 * LAST UPDATED: 2025-10-19

 * ============================================================================

 */



// ============================================================================

// CONFIGURATION & CONSTANTS

// ============================================================================



const MASTER_SHEET = 'Pan_Master';

const REPORT_SHEET = 'Label_Report';

const LABEL_COLUMN_NAME = 'Guest Label';



// Party Date (The year doesn't matter for MM/DD comparisons)

const PARTY_MONTH = 10; 

const PARTY_DAY = 25;

const BIRTHDAY_WINDOW_DAYS = 10; // +/- 10 days from Oct 25



// Old Soul Criteria: Age 21-29 (Younger Group) AND Pop Music (Older Group's dominant preference)

// Age Codes for Young Cohort: 1:21-24, 2:25-29 (These are the base groups being analyzed)

const OLD_SOUL_AGE_CODES = [1, 2]; 

// Music Codes for Trait: Pop (1) is the preference correlated with 30-34 group.

const OLD_SOUL_MUSIC_CODES = [1]; // Pop music code



// Host VIP Criteria: Known Host for 3+ years

// Codes: 5: Yes—3–5 years, 6: Yes—5–10 years, 7: Yes—more than 10 years

const HOST_VIP_CODES = [5, 6, 7]; 



// Party Starter Criteria: High score on social stance scale (4 or 5 on 1-5 scale)

const PARTY_STARTER_THRESHOLD = 4; 



// Gaming Guru Criteria: Gaming Interest (1) AND specific music codes 

// Music Codes: Pop (1), Rock (2), Hip-hop (3), Electronic (4), Indie/Alt (6) (based on provided Pan_Dict)

const GAMING_MUSIC_CODES = [1, 2, 3, 4, 6]; 





// ============================================================================

// MAIN EXECUTION FUNCTION

// ============================================================================



/**

 * Creates custom menu when spreadsheet opens

 */

function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu('📊 Analytics')

    .addItem('🔨 Build Pan Sheets (Requires Pan_Analytics.gs)', 'PLACEHOLDER_FOR_PAN_SHEETS')

    .addItem('⭐ Generate Master Label Report', 'generateMasterLabelReport')

    .addToUi();

}





/**

 * Master function to calculate labels, apply them to Pan_Master, and generate the Label_Report.

 */

function generateMasterLabelReport() {

  const ss = SpreadsheetApp.getActive();

  const sh = ss.getSheetByName(MASTER_SHEET);

  const ui = SpreadsheetApp.getUi();



  if (!sh) {

    ui.alert('Error', `Source sheet "${MASTER_SHEET}" not found. Please ensure Pan_Master exists.`, ui.ButtonSet.OK);

    return;

  }



  // --- PHASE 1: CALCULATE AND APPLY LABELS TO PAN_MASTER ---

  ui.showSidebar(HtmlService.createHtmlOutput('Calculating labels...'));

  

  // NOTE: dataRange and values MUST be re-read if a column is added.

  let dataRange = sh.getDataRange();

  let values = dataRange.getValues();

  const headers = values[0];

  

  // 1. Get column indices (0-based)

  const idx = getColumnIndices_(headers);

  if (!checkRequiredColumns_(ui, idx)) return; // Check required columns



  // 2. Add or find the Label column

  let labelColIndex = headers.indexOf(LABEL_COLUMN_NAME);

  if (labelColIndex === -1) {

    // Insert column and re-read data if column was missing

    sh.insertColumnAfter(headers.length);

    sh.getRange(1, headers.length + 1).setValue(LABEL_COLUMN_NAME);

    labelColIndex = headers.length; 

    

    // Re-fetch all data to include the new column

    dataRange = sh.getDataRange();

    values = dataRange.getValues(); 

  }



  // 3. Process data rows and generate labels

  const partyDate = new Date(2000, PARTY_MONTH - 1, PARTY_DAY); 

  const labelsToApply = [];

  const reportRows = [];

  const labelSummary = {};

  

  for (let i = 1; i < values.length; i++) {

    const row = values[i];

    const labels = [];

    

    // --- LABEL GENERATION LOGIC ---

    if (isNearBirthday_(row[idx.birthday], partyDate)) {

      addLabel(labels, '🎂 Happy Birthday!');

    }

    if (isOldSoul_(row[idx.age], row[idx.musicPref])) {

      addLabel(labels, '👴 Old Soul');

    }

    if (isHostVIP_(row[idx.knowHosts])) {

      addLabel(labels, '⭐ Host VIP');

    }

    if (isPartyStarter_(row[idx.social])) {

      addLabel(labels, '🎉 Party Starter');

    }

    if (isGamingGuru_(row[idx.interests], row[idx.musicPref])) {

      addLabel(labels, '🎮 Gaming Guru');

    }

    // --- END LABEL GENERATION LOGIC ---



    const labelString = labels.join(', ');

    labelsToApply.push([labelString]);

    

    // Prepare data for report and summary (Screen Name, UID, Label String)

    const screenName = row[idx.screenName] || 'N/A';

    const uid = row[idx.uid] || 'N/A';

    reportRows.push([screenName, uid, labelString]);

    

    if (labelString) {

      labels.forEach(label => {

        labelSummary[label] = (labelSummary[label] || 0) + 1;

      });

    }

  }



  // 4. Write back the new column of labels to Pan_Master

  if (labelsToApply.length > 0) {

    sh.getRange(2, labelColIndex + 1, labelsToApply.length, 1).setValues(labelsToApply);

  }



  // --- PHASE 2: GENERATE LABEL REPORT SHEET ---

  const reportSh = clearOrCreateReportSheet_();

  const nextStartRow = writeSummarySection_(reportSh, labelSummary, reportRows.length);

  writeDetailedGuestSection_(reportSh, reportRows, nextStartRow);

  

  ui.alert('✅ Master Report Complete', 

    `1. 'Guest Label' column updated in Pan_Master.
` +

    `2. Detailed summary report created in "${REPORT_SHEET}".`, 

    ui.ButtonSet.OK

  );

}



// ============================================================================

// HELPER FUNCTIONS - LABEL GENERATION LOGIC

// ============================================================================



/**

 * Helper to ensure a label is added only once.

 */

function addLabel(labels, label) {

    if (!labels.includes(label)) {

        labels.push(label);

    }

}



/**

 * Checks if the birthday (MM/DD string) is near the party date (Oct 25).

 */

function isNearBirthday_(birthdayMMDD, partyDate) {

  if (typeof birthdayMMDD !== 'string' || birthdayMMDD.length < 4) return false;

  

  const parts = birthdayMMDD.split('/').map(s => parseInt(s, 10));

  if (parts.length < 2 || isNaN(parts[0]) || isNaN(parts[1])) return false;



  const month = parts[0];

  const day = parts[1];

  

  // Use year 2000 for non-leap year parity for all date comparisons

  const bDate = new Date(2000, month - 1, day); 

  

  // Calculate differences

  const msInDay = 1000 * 60 * 60 * 24;

  const diffDays = Math.abs((bDate.getTime() - partyDate.getTime()) / msInDay);

  

  // Account for wrap-around (e.g., Oct 25 vs Jan 1) by checking adjacent years

  const bDateLastYear = new Date(1999, month - 1, day);

  const bDateNextYear = new Date(2001, month - 1, day);

  

  const diffDaysWrapLast = Math.abs((bDateLastYear.getTime() - partyDate.getTime()) / msInDay);

  const diffDaysWrapNext = Math.abs((bDateNextYear.getTime() - partyDate.getTime()) / msInDay);



  // Check if any of the three year possibilities are within the window

  return diffDays <= BIRTHDAY_WINDOW_DAYS || 

         diffDaysWrapLast <= BIRTHDAY_WINDOW_DAYS || 

         diffDaysWrapNext <= BIRTHDAY_WINDOW_DAYS;

}



/**

 * Parses input value reliably into a number or returns NaN if it's invalid.

 */

function getNumericCode(value) {

    if (value === null || value === undefined || value === '') return NaN;

    // Attempt to convert strings (like '5') or keep numbers (like 5)

    const num = Number(String(value).trim());

    return isNaN(num) ? NaN : num;

}





/**

 * Checks if a guest is an "Old Soul" based on combined age and music preference.

 */

function isOldSoul_(ageCode, musicPrefCode) {

    const age = getNumericCode(ageCode);

    const music = getNumericCode(musicPrefCode);

    

    // Fail immediately if codes are not valid numbers

    if (isNaN(age) || isNaN(music)) return false;



    // Age must be 21-29 (Codes 1 or 2)

    const isYoungCohort = OLD_SOUL_AGE_CODES.includes(age);

    // Music must be Pop (Code 1) which correlates with 30-34 group

    const exhibitsOlderTrait = OLD_SOUL_MUSIC_CODES.includes(music);



    return isYoungCohort && exhibitsOlderTrait;

}



/**

 * Checks if the 'know hosts' code indicates a long-standing relationship (3+ years).

 */

function isHostVIP_(knowHostsCode) {

    const code = getNumericCode(knowHostsCode);

    

    if (isNaN(code)) return false;

    return HOST_VIP_CODES.includes(code);

}



/**

 * Checks if the social stance code indicates an extrovert (4 or higher on 1-5 scale).

 */

function isPartyStarter_(socialStanceCode) {

    const code = getNumericCode(socialStanceCode);

    

    if (isNaN(code)) return false;

    return code >= PARTY_STARTER_THRESHOLD; 

}



/**

 * Checks for Gaming interest AND specific music genres popular in gaming culture.

 */

function isGamingGuru_(gamingInterest, musicPrefCode) {

    // 'oh_interests_Gaming' is a binary column (1 or 0)

    const isGaming = getNumericCode(gamingInterest) === 1; 

    

    const musicCode = getNumericCode(musicPrefCode);

    if (isNaN(musicCode)) return false;

    

    // Music Codes: Pop, Rock, Hip-hop, Electronic, Indie/Alt

    const hasGamingMusic = GAMING_MUSIC_CODES.includes(musicCode); 



    return isGaming && hasGamingMusic;

}



// ============================================================================

// HELPER FUNCTIONS - REPORT GENERATION (From label_report_builder.gs)

// ============================================================================



/**

 * Locates required columns indices in the Pan_Master header.

 */

function getColumnIndices_(headers) {

  // NOTE: headers.indexOf returns -1 if not found.

  return {

    screenName: headers.indexOf('Screen Name'),

    uid: headers.indexOf('UID'),

    birthday: headers.indexOf('Birthday_MM/DD'),

    age: headers.indexOf('code_age_range'),

    knowHosts: headers.indexOf('code_know_hosts'),

    musicPref: headers.indexOf('code_music_pref'),

    interests: headers.indexOf('oh_interests_Gaming'),

    social: headers.indexOf('code_social_stance')

  };

}



/**

 * Checks if required columns exist and alerts the user if any are missing.

 */

function checkRequiredColumns_(ui, idx) {

    const required = ['Screen Name', 'UID', 'Birthday_MM/DD', 'code_age_range', 'code_know_hosts', 'code_music_pref', 'oh_interests_Gaming', 'code_social_stance'];

    const missing = [];

    

    // Loop through the named properties of the index map to check for -1

    for (const key in idx) {

        // Find the header name associated with the index key to check if it's one of the required fields

        // This is a complex check, simplifying to ensure all values are >= 0

        if (idx[key] === -1 && key !== 'uid') { // UID check is complex, focusing on the codes

            // We use a general check to see if the index is -1 for critical codes

            const headerNames = {

                'screenName': 'Screen Name', 'uid': 'UID', 'birthday': 'Birthday_MM/DD',

                'age': 'code_age_range', 'knowHosts': 'code_know_hosts', 'musicPref': 'code_music_pref',

                'interests': 'oh_interests_Gaming', 'social': 'code_social_stance'

            };

            missing.push(headerNames[key] || key);

        }

    }

    

    if (missing.length > 0) {

        // Final check on critical codes

        const criticalMissing = missing.filter(name => required.includes(name));

        if (criticalMissing.length > 0) {

            ui.alert('Fatal Error', `Missing required columns in Pan_Master: ${criticalMissing.join(', ')}. Please run Pan_Analytics build functions first.`, ui.ButtonSet.OK);

            return false;

        }

    }

    return true;

}



/**

 * Writes the summary table (Label, Count, Percentage) to the report sheet.

 */

function writeSummarySection_(sh, summary, totalGuests) {

  const summaryHeader = ['Label Archetype', 'Guest Count', 'Percentage (%)'];

  const summaryData = [];

  

  const sortedLabels = Object.keys(summary).sort((a, b) => summary[b] - summary[a]);



  sortedLabels.forEach(label => {

    const count = summary[label];

    const percentage = count / totalGuests;

    summaryData.push([label, count, percentage]);

  });

  

  const totalRow = ['TOTAL GUESTS PROCESSED', totalGuests, totalGuests > 0 ? 1 : 0];



  // Write title

  sh.getRange(1, 1).setValue('LABEL ARCHETYPE SUMMARY');

  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');



  // Write summary data

  const startRow = 3;

  sh.getRange(startRow, 1, 1, summaryHeader.length).setValues([summaryHeader]);

  sh.getRange(startRow + 1, 1, summaryData.length, summaryHeader.length).setValues(summaryData);

  

  // Write total row

  const totalRowIndex = startRow + summaryData.length + 1;

  sh.getRange(totalRowIndex, 1, 1, summaryHeader.length).setValues([totalRow]);

  sh.getRange(totalRowIndex, 1, 1, summaryHeader.length)

    .setFontWeight('bold')

    .setBackground('#eeeeee');



  // Apply formatting to count and percentage columns

  if (summaryData.length > 0) {

    sh.getRange(startRow + 1, 3, summaryData.length, 1).setNumberFormat('0.0%');

  }

  sh.getRange(totalRowIndex, 3).setNumberFormat('0.0%');

  

  // Format header

  sh.getRange(startRow, 1, 1, summaryHeader.length)

    .setBackground('#434343')

    .setFontColor('#ffffff')

    .setFontWeight('bold')

    .setHorizontalAlignment('center');

  

  return totalRowIndex + 2; // Return next available row for detailed section

}



/**

 * Writes the detailed guest list section to the report sheet.

 */

function writeDetailedGuestSection_(sh, rows, startRow) {

  const detailHeader = ['Screen Name', 'UID', LABEL_COLUMN_NAME];

  

  // Start the detailed section title at the provided row index

  sh.getRange(startRow, 1).setValue('GUEST BREAKDOWN BY LABEL');

  sh.getRange(startRow, 1).setFontSize(16).setFontWeight('bold');

  

  const dataStartRow = startRow + 2;

  

  // Write header

  sh.getRange(dataStartRow, 1, 1, detailHeader.length).setValues([detailHeader])

    .setBackground('#434343')

    .setFontColor('#ffffff')

    .setFontWeight('bold');



  // Write guest data

  sh.getRange(dataStartRow + 1, 1, rows.length, detailHeader.length).setValues(rows);



  // Set frozen rows to keep the detailed header visible 

  sh.setFrozenRows(dataStartRow);

  

  // Auto-resize columns

  sh.autoResizeColumns(1, detailHeader.length);

}



/**

 * Utility function to clear existing report sheet or create new one.

 */

function clearOrCreateReportSheet_() {

  const ss = SpreadsheetApp.getActive();

  let sh = ss.getSheetByName(REPORT_SHEET);

  if (!sh) {

    sh = ss.insertSheet(REPORT_SHEET);

  } else {

    sh.clear();

  }

  sh.setFrozenRows(1);

  return sh;

}



// ===== File: V_Cramer =====


/**

 * ============================================================================

 * REPORTS.GS - ALL PARTY SURVEY ANALYSIS REPORTS

 * ============================================================================

 * 

 * Contains all report generation functions for party survey analysis.

 * Each function creates a formatted output sheet with analysis results.

 * 

 * AVAILABLE REPORTS:

 * 1. buildDemographicsSummary() - Age, gender, ethnicity, education breakdown

 * 2. buildMusicInterestsReport() - Music preferences and general interests

 * 3. buildHostRelationshipReport() - Guest relationships with party hosts

 * 4. buildAttendeeAnalysis() - Compare attendees vs no-shows

 * 5. buildGuestProfiles() - Individual guest profile cards

 * 

 * USAGE:

 * Run any function from the script editor or create a custom menu

 */



// ============================================================================

// GLOBAL CONFIGURATION

// ============================================================================





// ============================================================================

// REPORT 1: DEMOGRAPHICS SUMMARY

// ============================================================================



/**

 * Generate comprehensive demographic overview

 * Shows distributions across age, gender, ethnicity, education, employment

 * Output: Demographics_Summary sheet

 */

function buildDemographicsSummary() {

  const ss = SpreadsheetApp.getActive();

  const cleanSheet = ss.getSheetByName(REPORTS_CLEAN_SHEET);

  

  if (!cleanSheet) {

    throw new Error(`Sheet "${REPORTS_CLEAN_SHEET}" not found.`);

  }



  const data = cleanSheet.getDataRange().getValues();

  if (data.length < 2) {

    writeSheet_('Demographics_Summary', [['No data available']]);

    return;

  }



  const header = data[0];

  const colIdx = getColumnMap_(header);



  // Calculate overview metrics

  const totalCount = data.length - 1;

  const checkedInCount = data.slice(1).filter(row => 

    String(row[colIdx['Checked-In at Event']] || '').trim().toUpperCase() === 'Y'

  ).length;



  // Build report sections

  const sections = [];



  // Title and overview

  sections.push(['PARTY SURVEY DEMOGRAPHICS SUMMARY']);

  sections.push(['']);

  sections.push(['Metric', 'Count', 'Percentage']);

  sections.push(['Total Responses', totalCount, '100.0%']);

  sections.push(['Checked-In Attendees', checkedInCount, ((checkedInCount / totalCount) * 100).toFixed(1) + '%']);

  sections.push(['No-Shows', totalCount - checkedInCount, (((totalCount - checkedInCount) / totalCount) * 100).toFixed(1) + '%']);

  sections.push(['']);



  // Add distribution sections

  const demographicFields = [

    ['Age Range', 'AGE RANGE DISTRIBUTION'],

    ['Self-Identified Gender', 'GENDER DISTRIBUTION'],

    ['Self Identified Ethnicity', 'ETHNICITY DISTRIBUTION'],

    ['Education Level', 'EDUCATION LEVEL DISTRIBUTION'],

    ['Self-Identified Sexual Orientation', 'SEXUAL ORIENTATION DISTRIBUTION'],

    ['Employment Information (Industry)', 'INDUSTRY DISTRIBUTION'],

    ['Employment Information (Role)', 'ROLE DISTRIBUTION']

  ];



  demographicFields.forEach(([colName, title]) => {

    sections.push(...buildDistribution_(data, colIdx, colName, title, 'Checked-In at Event'));

    sections.push(['']);

  });



  writeSheet_('Demographics_Summary', sections);

  formatDemographicsSheet_();

}



/**

 * Build distribution table for a demographic variable

 */

function buildDistribution_(data, colIdx, colName, title, checkedInCol) {

  const dataColIdx = colIdx[colName];

  const checkedInIdx = colIdx[checkedInCol];

  

  if (dataColIdx === undefined) {

    return [[title], ['Column not found']];

  }



  const totalCounts = {};

  const checkedInCounts = {};



  for (let i = 1; i < data.length; i++) {

    const value = String(data[i][dataColIdx] || '').trim();

    const isCheckedIn = String(data[i][checkedInIdx] || '').trim().toUpperCase() === 'Y';



    if (!value) continue;



    totalCounts[value] = (totalCounts[value] || 0) + 1;

    if (isCheckedIn) {

      checkedInCounts[value] = (checkedInCounts[value] || 0) + 1;

    }

  }



  const sortedValues = Object.keys(totalCounts).sort((a, b) => totalCounts[b] - totalCounts[a]);

  const totalSum = Object.values(totalCounts).reduce((a, b) => a + b, 0);

  const checkedInSum = Object.values(checkedInCounts).reduce((a, b) => a + b, 0);



  const rows = [

    [title],

    ['Category', 'Total', '% of Total', 'Checked-In', '% of Checked-In']

  ];



  sortedValues.forEach(value => {

    const totalCount = totalCounts[value];

    const checkedInCount = checkedInCounts[value] || 0;

    const totalPct = ((totalCount / totalSum) * 100).toFixed(1);

    const checkedInPct = checkedInSum > 0 ? ((checkedInCount / checkedInSum) * 100).toFixed(1) : '0.0';



    rows.push([value, totalCount, `${totalPct}%`, checkedInCount, `${checkedInPct}%`]);

  });



  rows.push(['TOTAL', totalSum, '100.0%', checkedInSum, '100.0%']);



  return rows;

}



function formatDemographicsSheet_() {

  const sheet = SpreadsheetApp.getActive().getSheetByName('Demographics_Summary');

  if (!sheet) return;



  const data = sheet.getDataRange().getValues();

  

  for (let row = 1; row <= data.length; row++) {

    const cellValue = String(data[row - 1][0] || '').trim();

    

    if (cellValue === 'PARTY SURVEY DEMOGRAPHICS SUMMARY') {

      sheet.getRange(row, 1, 1, 5).setFontWeight('bold').setFontSize(12)

        .setBackground('#1c4587').setFontColor('#ffffff');

    } else if (cellValue.toUpperCase() === cellValue && cellValue.includes('DISTRIBUTION')) {

      sheet.getRange(row, 1, 1, 5).setFontWeight('bold').setFontSize(11)

        .setBackground('#4a86e8').setFontColor('#ffffff');

    } else if (cellValue === 'Category' || cellValue === 'Metric') {

      sheet.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#d9d9d9');

    } else if (cellValue === 'TOTAL') {

      sheet.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#f3f3f3');

    }

  }



  for (let col = 1; col <= 5; col++) {

    sheet.autoResizeColumn(col);

  }

}



// ============================================================================

// REPORT 2: MUSIC & INTERESTS

// ============================================================================



/**

 * Analyze music preferences and general interests

 * Output: Music_Interests_Report sheet

 */

function buildMusicInterestsReport() {

  const ss = SpreadsheetApp.getActive();

  const cleanSheet = ss.getSheetByName(REPORTS_CLEAN_SHEET);

  

  if (!cleanSheet) {

    throw new Error(`Sheet "${REPORTS_CLEAN_SHEET}" not found.`);

  }



  const data = cleanSheet.getDataRange().getValues();

  if (data.length < 2) {

    writeSheet_('Music_Interests_Report', [['No data available']]);

    return;

  }



  const header = data[0];

  const colIdx = getColumnMap_(header);



  const sections = [];



  // Title

  sections.push(['MUSIC & INTERESTS ANALYSIS']);

  sections.push(['']);



  // Music Preferences

  sections.push(...buildMusicPreferences_(data, colIdx));

  sections.push(['']);



  // Top Artists

  sections.push(...buildTopArtists_(data, colIdx, 15));

  sections.push(['']);



  // Top Interests

  sections.push(...buildTopInterests_(data, colIdx));

  sections.push(['']);



  // Music by Age

  sections.push(...buildMusicByDemographic_(data, colIdx, 'Age Range', 'AGE GROUP'));

  sections.push(['']);



  // Music by Gender

  sections.push(...buildMusicByDemographic_(data, colIdx, 'Self-Identified Gender', 'GENDER'));

  sections.push(['']);



  // Song Requests

  sections.push(...buildSongRequestsSummary_(data, colIdx));



  writeSheet_('Music_Interests_Report', sections);

  formatGenericReport_('Music_Interests_Report');

}



function buildMusicPreferences_(data, colIdx) {

  const col = colIdx['Music Preference'];

  if (col === undefined) return [['MUSIC PREFERENCE DISTRIBUTION'], ['Column not found']];



  const counts = {};

  let total = 0;



  for (let i = 1; i < data.length; i++) {

    const value = String(data[i][col] || '').trim();

    if (!value) continue;

    counts[value] = (counts[value] || 0) + 1;

    total++;

  }



  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);

  const rows = [['MUSIC PREFERENCE DISTRIBUTION'], ['Genre', 'Count', 'Percentage']];



  sorted.forEach(([genre, count]) => {

    rows.push([genre, count, `${((count / total) * 100).toFixed(1)}%`]);

  });



  rows.push(['TOTAL', total, '100.0%']);

  return rows;

}



function buildTopArtists_(data, colIdx, limit) {

  const col = colIdx['Current Favorite Artist'];

  if (col === undefined) return [[`TOP ${limit} FAVORITE ARTISTS`], ['Column not found']];



  const counts = {};

  for (let i = 1; i < data.length; i++) {

    const value = String(data[i][col] || '').trim();

    if (!value) continue;

    counts[value] = (counts[value] || 0) + 1;

  }



  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, limit);

  const rows = [[`TOP ${limit} FAVORITE ARTISTS`], ['Rank', 'Artist', 'Mentions']];



  sorted.forEach(([artist, count], idx) => {

    rows.push([idx + 1, artist, count]);

  });



  return rows;

}



function buildTopInterests_(data, colIdx) {

  const col = colIdx['Your General Interests (Choose 3)'];

  if (col === undefined) return [['TOP INTERESTS'], ['Column not found']];



  const counts = {};

  let totalResponses = 0;



  for (let i = 1; i < data.length; i++) {

    const value = String(data[i][col] || '').trim();

    if (!value) continue;

    totalResponses++;



    const interests = value.split(',').map(s => s.trim()).filter(s => s);

    interests.forEach(interest => {

      counts[interest] = (counts[interest] || 0) + 1;

    });

  }



  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);

  const rows = [['TOP INTERESTS'], ['Interest', 'Count', '% of Guests']];



  sorted.forEach(([interest, count]) => {

    rows.push([interest, count, `${((count / totalResponses) * 100).toFixed(1)}%`]);

  });



  return rows;

}



function buildMusicByDemographic_(data, colIdx, demoCol, demoLabel) {

  const musicCol = colIdx['Music Preference'];

  const demoColIdx = colIdx[demoCol];

  

  if (musicCol === undefined || demoColIdx === undefined) {

    return [[`MUSIC PREFERENCE BY ${demoLabel}`], ['Required columns not found']];

  }



  const crosstab = {};

  for (let i = 1; i < data.length; i++) {

    const demo = String(data[i][demoColIdx] || '').trim();

    const music = String(data[i][musicCol] || '').trim();

    if (!demo || !music) continue;



    if (!crosstab[demo]) crosstab[demo] = {};

    crosstab[demo][music] = (crosstab[demo][music] || 0) + 1;

  }



  const allGenres = new Set();

  Object.values(crosstab).forEach(musicCounts => {

    Object.keys(musicCounts).forEach(genre => allGenres.add(genre));

  });

  const genres = Array.from(allGenres).sort();

  const demos = Object.keys(crosstab).sort();



  const rows = [[`MUSIC PREFERENCE BY ${demoLabel}`], [demoCol, ...genres, 'Total']];



  demos.forEach(demo => {

    const row = [demo];

    let rowTotal = 0;

    genres.forEach(genre => {

      const count = crosstab[demo][genre] || 0;

      row.push(count);

      rowTotal += count;

    });

    row.push(rowTotal);

    rows.push(row);

  });



  return rows;

}



function buildSongRequestsSummary_(data, colIdx) {

  const col = colIdx['Name one song you want to hear at the party.'];

  if (col === undefined) return [['SONG REQUESTS SUMMARY'], ['Column not found']];



  let totalResponses = 0, withSong = 0, withoutSong = 0;



  for (let i = 1; i < data.length; i++) {

    totalResponses++;

    const value = String(data[i][col] || '').trim();

    if (value) withSong++; else withoutSong++;

  }



  const withPct = ((withSong / totalResponses) * 100).toFixed(1);

  const withoutPct = ((withoutSong / totalResponses) * 100).toFixed(1);



  return [

    ['SONG REQUESTS SUMMARY'],

    ['Category', 'Count', 'Percentage'],

    ['Provided Song Request', withSong, `${withPct}%`],

    ['No Song Request', withoutSong, `${withoutPct}%`],

    ['Total Responses', totalResponses, '100.0%']

  ];

}



// ============================================================================

// REPORT 3: HOST RELATIONSHIPS

// ============================================================================



/**

 * Analyze guest relationships with party hosts

 * Output: Host_Relationships sheet

 */

function buildHostRelationshipReport() {

  const ss = SpreadsheetApp.getActive();

  const cleanSheet = ss.getSheetByName(REPORTS_CLEAN_SHEET);

  

  if (!cleanSheet) {

    throw new Error(`Sheet "${REPORTS_CLEAN_SHEET}" not found.`);

  }



  const data = cleanSheet.getDataRange().getValues();

  if (data.length < 2) {

    writeSheet_('Host_Relationships', [['No data available']]);

    return;

  }



  const header = data[0];

  const colIdx = getColumnMap_(header);



  const sections = [];



  sections.push(['HOST RELATIONSHIP ANALYSIS']);

  sections.push(['']);



  // Overview

  sections.push(...buildHostOverview_(data, colIdx));

  sections.push(['']);



  // Duration Distribution

  sections.push(...buildHostDuration_(data, colIdx));

  sections.push(['']);



  // Which Host Known

  sections.push(...buildWhichHost_(data, colIdx));

  sections.push(['']);



  // Closeness Scores

  sections.push(...buildClosenessScores_(data, colIdx));

  sections.push(['']);



  // Closeness by Host

  sections.push(...buildClosenessbyHost_(data, colIdx));



  writeSheet_('Host_Relationships', sections);

  formatGenericReport_('Host_Relationships');

}



function buildHostOverview_(data, colIdx) {

  const col = colIdx['Do you know the Host(s)?'];

  if (col === undefined) return [['OVERVIEW'], ['Column not found']];



  let total = data.length - 1, knowHosts = 0, unknownGuests = 0;



  for (let i = 1; i < data.length; i++) {

    const value = String(data[i][col] || '').trim().toLowerCase();

    if (value.includes('yes') || value.includes('—')) knowHosts++;

    else if (value.includes('no')) unknownGuests++;

  }



  return [

    ['OVERVIEW'],

    ['Metric', 'Count', 'Percentage'],

    ['Total Guests', total, '100.0%'],

    ['Know Host(s)', knowHosts, `${((knowHosts / total) * 100).toFixed(1)}%`],

    ['Unknown Guests', unknownGuests, `${((unknownGuests / total) * 100).toFixed(1)}%`]

  ];

}



function buildHostDuration_(data, colIdx) {

  const col = colIdx['Do you know the Host(s)?'];

  if (col === undefined) return [['FRIENDSHIP DURATION'], ['Column not found']];



  const counts = {};

  let total = 0;



  for (let i = 1; i < data.length; i++) {

    const value = String(data[i][col] || '').trim();

    if (!value) continue;

    counts[value] = (counts[value] || 0) + 1;

    total++;

  }



  const order = ['No', 'Yes — 3–12 months', 'Yes — 1–3 years', 'Yes — 3–5 years', 'Yes — 5–10 years', 'Yes — more than 10 years'];

  const sorted = Object.keys(counts).sort((a, b) => {

    const idxA = order.indexOf(a);

    const idxB = order.indexOf(b);

    if (idxA === -1 && idxB === -1) return a.localeCompare(b);

    if (idxA === -1) return 1;

    if (idxB === -1) return -1;

    return idxA - idxB;

  });



  const rows = [['FRIENDSHIP DURATION DISTRIBUTION'], ['Duration', 'Count', 'Percentage']];



  sorted.forEach(duration => {

    const count = counts[duration];

    rows.push([duration, count, `${((count / total) * 100).toFixed(1)}%`]);

  });



  rows.push(['TOTAL', total, '100.0%']);

  return rows;

}



function buildWhichHost_(data, colIdx) {

  const col = colIdx['Which host have you known the longest?'];

  if (col === undefined) return [['WHICH HOST KNOWN LONGEST'], ['Column not found']];



  const counts = {};

  let total = 0;



  for (let i = 1; i < data.length; i++) {

    const value = String(data[i][col] || '').trim();

    if (!value) continue;

    counts[value] = (counts[value] || 0) + 1;

    total++;

  }



  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);

  const rows = [['WHICH HOST KNOWN LONGEST'], ['Host', 'Count', 'Percentage']];



  sorted.forEach(([host, count]) => {

    rows.push([host, count, `${((count / total) * 100).toFixed(1)}%`]);

  });



  rows.push(['TOTAL', total, '100.0%']);

  return rows;

}



function buildClosenessScores_(data, colIdx) {

  const col = colIdx['If yes, how well do you know them?'];

  if (col === undefined) return [['CLOSENESS SCORE DISTRIBUTION'], ['Column not found']];



  const counts = {};

  let total = 0, sum = 0;



  for (let i = 1; i < data.length; i++) {

    const value = data[i][col];

    if (value === null || value === undefined || value === '') continue;

    const score = Number(value);

    if (!isFinite(score)) continue;

    

    counts[score] = (counts[score] || 0) + 1;

    sum += score;

    total++;

  }



  const avg = total > 0 ? (sum / total).toFixed(2) : 'N/A';

  const sorted = Object.keys(counts).sort((a, b) => Number(b) - Number(a));



  const rows = [

    ['CLOSENESS SCORE DISTRIBUTION (1=Acquaintance, 5=Close Friend)'],

    ['Score', 'Count', 'Percentage']

  ];



  sorted.forEach(score => {

    const count = counts[score];

    rows.push([score, count, `${((count / total) * 100).toFixed(1)}%`]);

  });



  rows.push(['TOTAL', total, '100.0%']);

  rows.push(['AVERAGE SCORE', avg, '']);

  return rows;

}



function buildClosenessbyHost_(data, colIdx) {

  const hostCol = colIdx['Which host have you known the longest?'];

  const scoreCol = colIdx['If yes, how well do you know them?'];

  

  if (hostCol === undefined || scoreCol === undefined) {

    return [['AVERAGE CLOSENESS BY HOST'], ['Required columns not found']];

  }



  const hostScores = {};



  for (let i = 1; i < data.length; i++) {

    const host = String(data[i][hostCol] || '').trim();

    const score = data[i][scoreCol];

    

    if (!host || score === null || score === undefined || score === '') continue;

    const scoreNum = Number(score);

    if (!isFinite(scoreNum)) continue;



    if (!hostScores[host]) hostScores[host] = { sum: 0, count: 0 };

    hostScores[host].sum += scoreNum;

    hostScores[host].count++;

  }



  const rows = [['AVERAGE CLOSENESS BY HOST'], ['Host', 'Count', 'Average Score']];



  Object.entries(hostScores).forEach(([host, data]) => {

    const avg = (data.sum / data.count).toFixed(2);

    rows.push([host, data.count, avg]);

  });



  return rows;

}



// ============================================================================

// REPORT 4: ATTENDEE VS NO-SHOW ANALYSIS

// ============================================================================



/**

 * Compare guests who attended vs those who didn't

 * Output: Attendee_Analysis sheet

 */

function buildAttendeeAnalysis() {

  const ss = SpreadsheetApp.getActive();

  const cleanSheet = ss.getSheetByName(REPORTS_CLEAN_SHEET);

  

  if (!cleanSheet) {

    throw new Error(`Sheet "${REPORTS_CLEAN_SHEET}" not found.`);

  }



  const data = cleanSheet.getDataRange().getValues();

  if (data.length < 2) {

    writeSheet_('Attendee_Analysis', [['No data available']]);

    return;

  }



  const header = data[0];

  const colIdx = getColumnMap_(header);



  const sections = [];



  sections.push(['ATTENDEE VS NO-SHOW ANALYSIS']);

  sections.push(['']);



  // Compare demographics

  const compareFields = [

    ['Age Range', 'AGE RANGE'],

    ['Self-Identified Gender', 'GENDER'],

    ['Education Level', 'EDUCATION'],

    ['Do you know the Host(s)?', 'HOST RELATIONSHIP']

  ];



  compareFields.forEach(([colName, label]) => {

    sections.push(...buildAttendeeComparison_(data, colIdx, colName, label));

    sections.push(['']);

  });



  writeSheet_('Attendee_Analysis', sections);

  formatGenericReport_('Attendee_Analysis');

}



function buildAttendeeComparison_(data, colIdx, colName, label) {

  const dataCol = colIdx[colName];

  const checkedInCol = colIdx['Checked-In at Event'];

  

  if (dataCol === undefined || checkedInCol === undefined) {

    return [[`${label} COMPARISON`], ['Required columns not found']];

  }



  const attendeeCounts = {};

  const noShowCounts = {};



  for (let i = 1; i < data.length; i++) {

    const value = String(data[i][dataCol] || '').trim();

    const isAttendee = String(data[i][checkedInCol] || '').trim().toUpperCase() === 'Y';



    if (!value) continue;



    if (isAttendee) {

      attendeeCounts[value] = (attendeeCounts[value] || 0) + 1;

    } else {

      noShowCounts[value] = (noShowCounts[value] || 0) + 1;

    }

  }



  const allValues = new Set([...Object.keys(attendeeCounts), ...Object.keys(noShowCounts)]);

  const sorted = Array.from(allValues).sort();



  const attendeeTotal = Object.values(attendeeCounts).reduce((a, b) => a + b, 0);

  const noShowTotal = Object.values(noShowCounts).reduce((a, b) => a + b, 0);



  const rows = [

    [`${label} COMPARISON: ATTENDEES VS NO-SHOWS`],

    ['Category', 'Attendees', '% Attendees', 'No-Shows', '% No-Shows']

  ];



  sorted.forEach(value => {

    const attendeeCount = attendeeCounts[value] || 0;

    const noShowCount = noShowCounts[value] || 0;

    const attendeePct = attendeeTotal > 0 ? ((attendeeCount / attendeeTotal) * 100).toFixed(1) : '0.0';

    const noShowPct = noShowTotal > 0 ? ((noShowCount / noShowTotal) * 100).toFixed(1) : '0.0';



    rows.push([value, attendeeCount, `${attendeePct}%`, noShowCount, `${noShowPct}%`]);

  });



  rows.push(['TOTAL', attendeeTotal, '100.0%', noShowTotal, '100.0%']);

  return rows;

}



// ============================================================================

// REPORT 5: GUEST PROFILE CARDS

// ============================================================================



/**

 * Create individual profile card for each guest

 * Output: Guest_Profiles sheet

 */

function buildGuestProfiles() {

  const ss = SpreadsheetApp.getActive();

  const cleanSheet = ss.getSheetByName(REPORTS_CLEAN_SHEET);

  const simSheet = ss.getSheetByName('Guest_Similarity');

  

  if (!cleanSheet) {

    throw new Error(`Sheet "${REPORTS_CLEAN_SHEET}" not found.`);

  }



  const data = cleanSheet.getDataRange().getValues();

  if (data.length < 2) {

    writeSheet_('Guest_Profiles', [['No data available']]);

    return;

  }



  const header = data[0];

  const colIdx = getColumnMap_(header);



  // Build profiles

  const profiles = [

    ['Screen Name', 'UID', 'Age', 'Gender', 'Orientation', 'Ethnicity', 'Education', 'Industry', 'Role', 

     'Top 3 Interests', 'Music Genre', 'Favorite Artist', 'Personality Trait', 'Host Known Longest', 

     'Closeness', 'Checked-In']

  ];



  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    

    // Parse interests

    const interestsStr = String(row[colIdx['Your General Interests (Choose 3)']] || '').trim();

    const interests = interestsStr.split(',').slice(0, 3).map(s => s.trim()).join(', ');



    profiles.push([

      row[colIdx['Screen Name']] || '',

      row[colIdx['UID']] || '',

      row[colIdx['Age Range']] || '',

      row[colIdx['Self-Identified Gender']] || '',

      row[colIdx['Self-Identified Sexual Orientation']] || '',

      row[colIdx['Self Identified Ethnicity']] || '',

      row[colIdx['Education Level']] || '',

      row[colIdx['Employment Information (Industry)']] || '',

      row[colIdx['Employment Information (Role)']] || '',

      interests,

      row[colIdx['Music Preference']] || '',

      row[colIdx['Current Favorite Artist']] || '',

      row[colIdx['At your worst you are…']] || '',

      row[colIdx['Which host have you known the longest?']] || '',

      row[colIdx['If yes, how well do you know them?']] || '',

      row[colIdx['Checked-In at Event']] || ''

    ]);

  }



  writeSheet_('Guest_Profiles', profiles);

  

  const sheet = ss.getSheetByName('Guest_Profiles');

  if (sheet) {

    sheet.getRange(1, 1, 1, profiles[0].length)

      .setFontWeight('bold')

      .setBackground('#4a86e8')

      .setFontColor('#ffffff');

    sheet.setFrozenRows(1);

    for (let col = 1; col <= profiles[0].length; col++) {

      sheet.autoResizeColumn(col);

    }

  }

}



// ============================================================================

// UTILITY FUNCTIONS

// ============================================================================



/**

 * Create column name to index mapping

 */

function getColumnMap_(header) {

  const map = {};

  header.forEach((col, idx) => {

    const colName = String(col || '').trim();

    if (colName) map[colName] = idx;

  });

  return map;

}



/**

 * Write data to a sheet

 */

function writeSheet_(sheetName, data) {

  const ss = SpreadsheetApp.getActive();

  let sheet = ss.getSheetByName(sheetName);

  

  if (!sheet) {

    sheet = ss.insertSheet(sheetName);

  }

  

  sheet.clear();



  if (!data || data.length === 0) {

    sheet.getRange(1, 1).setValue('No data to display');

    return;

  }



  // Find max columns

  const maxCols = Math.max(...data.map(row => Array.isArray(row) ? row.length : 1));



  // Pad rows to same length

  const paddedRows = data.map(row => {

    if (!Array.isArray(row)) return [row];

    const padded = [...row];

    while (padded.length < maxCols) padded.push('');

    return padded;

  });



  sheet.getRange(1, 1, paddedRows.length, maxCols).setValues(paddedRows);

}



/**

 * Apply generic formatting to report sheets

 */

function formatGenericReport_(sheetName) {

  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);

  if (!sheet) return;



  const data = sheet.getDataRange().getValues();

  const maxCols = data[0].length;



  for (let col = 1; col <= maxCols; col++) {

    sheet.autoResizeColumn(col);

  }



  for (let row = 1; row <= data.length; row++) {

    const cellValue = String(data[row - 1][0] || '').trim();

    

    // Main titles (ends with ANALYSIS)

    if (cellValue.includes('ANALYSIS')) {

      sheet.getRange(row, 1, 1, maxCols)

        .setFontWeight('bold')

        .setFontSize(14)

        .setBackground('#1c4587')

        .setFontColor('#ffffff');

    }

    // Section headers (all caps with keywords)

    else if (cellValue.toUpperCase() === cellValue && cellValue.length > 0 &&

             (cellValue.includes('DISTRIBUTION') || cellValue.includes('TOP ') || 

              cellValue.includes('BY ') || cellValue.includes('SUMMARY') ||

              cellValue.includes('COMPARISON') || cellValue.includes('OVERVIEW'))) {

      sheet.getRange(row, 1, 1, maxCols)

        .setFontWeight('bold')

        .setFontSize(11)

        .setBackground('#4a86e8')

        .setFontColor('#ffffff');

    }

    // Column headers

    else if (['Category', 'Metric', 'Genre', 'Rank', 'Interest', 'Host', 'Duration', 

              'Score', 'Age Range', 'Gender'].includes(cellValue)) {

      sheet.getRange(row, 1, 1, maxCols)

        .setFontWeight('bold')

        .setBackground('#d9d9d9');

    }

    // Total rows

    else if (cellValue === 'TOTAL' || cellValue.includes('AVERAGE')) {

      sheet.getRange(row, 1, 1, maxCols)

        .setFontWeight('bold')

        .setBackground('#f3f3f3');

    }

  }

}



// ===== File: Guest_Similarity =====


/**

 * Get compatibility matches for the MM.html matchmaker display

 * Reads from Edges_Top_Sim, enriches with guest details from Form Responses (Clean)

 * Only returns matches with similarity > 0.55, adds 10% for display effect

 * 

 * @return {Object} Match data for display

 */

function getCompatibilityMatches() {

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    

    // Get the top similarity edges

    const edgesSheet = ss.getSheetByName('Edges_Top_Sim');

    if (!edgesSheet) {

      Logger.log('❌ Edges_Top_Sim sheet not found');

      return { matches: [], totalGuests: 0 };

    }

    

    const edgesData = edgesSheet.getDataRange().getValues();

    const edgesHeaders = edgesData[0];

    

    // Find column indices in Edges_Top_Sim (lowercase: source, target, similarity)

    const sourceCol = edgesHeaders.indexOf('source');

    const targetCol = edgesHeaders.indexOf('target');

    const similarityCol = edgesHeaders.indexOf('similarity');

    

    if (sourceCol === -1 || targetCol === -1 || similarityCol === -1) {

      Logger.log('❌ Required columns not found in Edges_Top_Sim');

      Logger.log('Found headers: ' + edgesHeaders.join(', '));

      return { matches: [], totalGuests: 0 };

    }

    

    Logger.log('✅ Found Edges_Top_Sim with ' + (edgesData.length - 1) + ' rows');

    

    // Get Form Responses (Clean) for ALL guest details

    const cleanSheet = ss.getSheetByName('Form Responses (Clean)');

    if (!cleanSheet) {

      Logger.log('❌ Form Responses (Clean) sheet not found');

      return { matches: [], totalGuests: 0 };

    }

    

    const cleanData = cleanSheet.getDataRange().getValues();

    const cleanHeaders = cleanData[0];

    

    // Find columns in Form Responses (Clean) - based on Master_Desc

    const screenNameCol = 21; // Column 22 (0-indexed = 21) "Screen Name"

    const photoUrlCol = 25;   // Column 26 (0-indexed = 25) "Photo URL"

    const musicCol = 15;      // Column 16 (0-indexed = 15) "Music Preference"

    const zodiacCol = 2;      // Column 3 (0-indexed = 2) "Zodiac Sign"

    const interestsCol = 14;  // Column 15 (0-indexed = 14) "Your General Interests (Choose 3)"

    

    // Build guest lookup map from Form Responses (Clean)

    const guestMap = {};

    

    for (let i = 1; i < cleanData.length; i++) {

      const row = cleanData[i];

      const screenName = row[screenNameCol];

      

      if (screenName) {

        const interestsStr = row[interestsCol] || '';

        const interests = interestsStr ? interestsStr.split(',').map(s => s.trim()) : [];

        

        guestMap[screenName] = {

          screenName: screenName,

          photoUrl: processPhotoUrl(row[photoUrlCol] || ''),

          music: row[musicCol] || '---',

          zodiac: row[zodiacCol] || '---',

          interests: interests

        };

      }

    }

    

    const guestCount = Object.keys(guestMap).length;

    const photosCount = Object.values(guestMap).filter(g => g.photoUrl).length;

    

    Logger.log(`✅ Found ${guestCount} guests in Form Responses (Clean)`);

    Logger.log(`✅ ${photosCount} guests have photos`);

    

    // Process matches from Edges_Top_Sim

    const matches = [];

    let filteredOut = 0;

    

    for (let i = 1; i < edgesData.length; i++) {

      const row = edgesData[i];

      const screenName1 = row[sourceCol];

      const screenName2 = row[targetCol];

      const rawSimilarity = parseFloat(row[similarityCol]);

      

      // Filter: only show matches > 0.55

      if (!screenName1 || !screenName2 || isNaN(rawSimilarity)) {

        continue;

      }

      

      if (rawSimilarity <= 0.55) {

        filteredOut++;

        continue;

      }

      

      const person1 = guestMap[screenName1];

      const person2 = guestMap[screenName2];

      

      if (!person1 || !person2) {

        Logger.log(`⚠️ Missing guest data for: ${screenName1} or ${screenName2}`);

        continue;

      }

      

      // Find shared interests

      const interests1 = person1.interests || [];

      const interests2 = person2.interests || [];

      const sharedInterests = interests1.filter(int => interests2.includes(int));

      

      // Add common traits based on other attributes

      const commonTraits = [...sharedInterests];

      

      // Check music match

      if (person1.music && person2.music && 

          person1.music !== '---' && 

          person1.music === person2.music) {

        commonTraits.push(`Music: ${person1.music}`);

      }

      

      // Check zodiac match

      if (person1.zodiac && person2.zodiac && 

          person1.zodiac !== '---' && 

          person1.zodiac === person2.zodiac) {

        commonTraits.push(`Zodiac: ${person1.zodiac}`);

      }

      

      // Add 10% for visual effect (but cap at 1.0)

      const displaySimilarity = Math.min(rawSimilarity + 0.10, 1.0);

      

      matches.push({

        person1: {

          screenName: person1.screenName,

          photoUrl: person1.photoUrl || '',

          music: person1.music || '---',

          zodiac: person1.zodiac || '---',

          interests: interests1.slice(0, 5) // Limit to first 5

        },

        person2: {

          screenName: person2.screenName,

          photoUrl: person2.photoUrl || '',

          music: person2.music || '---',

          zodiac: person2.zodiac || '---',

          interests: interests2.slice(0, 5) // Limit to first 5

        },

        similarity: displaySimilarity, // Already includes +10% boost

        sharedInterests: commonTraits.slice(0, 8) // Limit to 8 shared traits

      });

    }

    

    // Sort by similarity (highest first)

    matches.sort((a, b) => b.similarity - a.similarity);

    

    Logger.log(`✅ Processed ${matches.length} matches (filtered out ${filteredOut} below 0.55)`);

    

    // Save to sheet for reference

    if (matches.length > 0) {

      saveMatchesToSheet(matches);

    }

    

    return {

      matches: matches,

      totalGuests: guestCount,

      matchesHeader: "The System Detects Compatibility, You Two Should Talk"

    };

    

  } catch (error) {

    Logger.log('❌ Error in getCompatibilityMatches: ' + error.toString());

    Logger.log('Stack: ' + error.stack);

    return { 

      matches: [], 

      totalGuests: 0,

      error: error.toString()

    };

  }

}



/**

 * Save matches to a sheet for reference

 */

function saveMatchesToSheet(matches) {

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let matchSheet = ss.getSheetByName('Recommended_Matches');

    

    // Create sheet if it doesn't exist

    if (!matchSheet) {

      matchSheet = ss.insertSheet('Recommended_Matches');

    }

    

    // Clear existing content

    matchSheet.clear();

    

    // Set header with title

    matchSheet.getRange(1, 1).setValue('THE SYSTEM DETECTS COMPATIBILITY, YOU TWO SHOULD TALK');

    matchSheet.getRange(1, 1, 1, 6).merge();

    matchSheet.getRange(1, 1).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');

    matchSheet.getRange(1, 1).setBackground('#FF6F00').setFontColor('#FFFFFF');

    

    // Set column headers

    const headers = [

      'Person 1',

      'Person 2',

      'Compatibility %',

      'Shared Interests',

      'Person 1 Photo',

      'Person 2 Photo'

    ];

    matchSheet.getRange(2, 1, 1, headers.length).setValues([headers]);

    matchSheet.getRange(2, 1, 1, headers.length).setFontWeight('bold').setBackground('#8D6E63').setFontColor('#FFFFFF');

    

    // Add data

    const rows = matches.map(match => [

      match.person1.screenName,

      match.person2.screenName,

      Math.round(match.similarity * 100) + '%',

      match.sharedInterests.join(', '),

      match.person1.photoUrl,

      match.person2.photoUrl

    ]);

    

    if (rows.length > 0) {

      matchSheet.getRange(3, 1, rows.length, headers.length).setValues(rows);

    }

    

    // Format

    matchSheet.autoResizeColumns(1, headers.length);

    matchSheet.setFrozenRows(2);

    

    // Add alternating row colors

    for (let i = 0; i < rows.length; i++) {

      const bgColor = i % 2 === 0 ? '#FFF8E1' : '#FFFFFF';

      matchSheet.getRange(3 + i, 1, 1, headers.length).setBackground(bgColor);

    }

    

    Logger.log(`✅ Saved ${matches.length} matches to Recommended_Matches sheet`);

    

  } catch (error) {

    Logger.log('❌ Error saving matches to sheet: ' + error.toString());

  }

}



/**

 * Test function to verify getCompatibilityMatches works

 */

function testGetCompatibilityMatches() {

  Logger.log('=== Starting Compatibility Matches Test ===
');

  

  const result = getCompatibilityMatches();

  

  Logger.log('
=== RESULTS ===');

  Logger.log('Total matches found: ' + result.matches.length);

  Logger.log('Total guests: ' + result.totalGuests);

  

  if (result.matches.length > 0) {

    Logger.log('
=== TOP 5 MATCHES ===');

    for (let i = 0; i < Math.min(5, result.matches.length); i++) {

      const match = result.matches[i];

      Logger.log(`
${i + 1}. ${match.person1.screenName} ⭐ ${match.person2.screenName}`);

      Logger.log(`   💕 Compatibility: ${Math.round(match.similarity * 100)}%`);

      Logger.log(`   🎵 Music: ${match.person1.music} / ${match.person2.music}`);

      Logger.log(`   ⭐ Zodiac: ${match.person1.zodiac} / ${match.person2.zodiac}`);

      Logger.log(`   ✨ Shared: ${match.sharedInterests.slice(0, 3).join(', ')}`);

      Logger.log(`   📸 Photos: ${match.person1.photoUrl ? '✅' : '❌'} / ${match.person2.photoUrl ? '✅' : '❌'}`);

    }

  } else {

    Logger.log('
⚠️ No matches found! Check:');

    Logger.log('  - Edges_Top_Sim has data with similarity > 0.55');

    Logger.log('  - Screen names in Edges_Top_Sim match Form Responses (Clean)');

  }

  

  SpreadsheetApp.getUi().alert(

    '🎃 Compatibility Matches Test

' +

    'Total Matches: ' + result.matches.length + '
' +

    'Total Guests: ' + result.totalGuests + '

' +

    (result.matches.length > 0 ? 

      'Top Match: ' + result.matches[0].person1.screenName + ' & ' + result.matches[0].person2.screenName + 

      ' (' + Math.round(result.matches[0].similarity * 100) + '%)

' : '') +

    'Check Recommended_Matches sheet and
execution log for full details!'

  );

}



/**

 * Get compatibility matches for the MM.html matchmaker display

 * Reads from Edges_Top_Sim, enriches with guest details from Form Responses (Clean)

 * Only returns matches with similarity > 0.55, adds 10% for display effect

 * 

 * @return {Object} Match data for display

 */

function getCompatibilityMatches() {

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    

    // Get the top similarity edges

    const edgesSheet = ss.getSheetByName('Edges_Top_Sim');

    if (!edgesSheet) {

      Logger.log('❌ Edges_Top_Sim sheet not found');

      return { matches: [], totalGuests: 0 };

    }

    

    const edgesData = edgesSheet.getDataRange().getValues();

    const edgesHeaders = edgesData[0];

    

    // Find column indices in Edges_Top_Sim (lowercase: source, target, similarity)

    const sourceCol = edgesHeaders.indexOf('source');

    const targetCol = edgesHeaders.indexOf('target');

    const similarityCol = edgesHeaders.indexOf('similarity');

    

    if (sourceCol === -1 || targetCol === -1 || similarityCol === -1) {

      Logger.log('❌ Required columns not found in Edges_Top_Sim');

      Logger.log('Found headers: ' + edgesHeaders.join(', '));

      return { matches: [], totalGuests: 0 };

    }

    

    Logger.log('✅ Found Edges_Top_Sim with ' + (edgesData.length - 1) + ' rows');

    

    // Get Form Responses (Clean) for ALL guest details

    const cleanSheet = ss.getSheetByName('Form Responses (Clean)');

    if (!cleanSheet) {

      Logger.log('❌ Form Responses (Clean) sheet not found');

      return { matches: [], totalGuests: 0 };

    }

    

    const cleanData = cleanSheet.getDataRange().getValues();

    const cleanHeaders = cleanData[0];

    

    // Find columns in Form Responses (Clean) - based on Master_Desc

    const screenNameCol = 21; // Column 22 (0-indexed = 21) "Screen Name"

    const photoUrlCol = 25;   // Column 26 (0-indexed = 25) "Photo URL"

    const musicCol = 15;      // Column 16 (0-indexed = 15) "Music Preference"

    const zodiacCol = 2;      // Column 3 (0-indexed = 2) "Zodiac Sign"

    const interestsCol = 14;  // Column 15 (0-indexed = 14) "Your General Interests (Choose 3)"

    

    // Build guest lookup map from Form Responses (Clean)

    const guestMap = {};

    

    for (let i = 1; i < cleanData.length; i++) {

      const row = cleanData[i];

      const screenName = row[screenNameCol];

      

      if (screenName) {

        const interestsStr = row[interestsCol] || '';

        const interests = interestsStr ? interestsStr.split(',').map(s => s.trim()) : [];

        

        guestMap[screenName] = {

          screenName: screenName,

          photoUrl: processPhotoUrl(row[photoUrlCol] || ''),

          music: row[musicCol] || '---',

          zodiac: row[zodiacCol] || '---',

          interests: interests

        };

      }

    }

    

    const guestCount = Object.keys(guestMap).length;

    const photosCount = Object.values(guestMap).filter(g => g.photoUrl).length;

    

    Logger.log(`✅ Found ${guestCount} guests in Form Responses (Clean)`);

    Logger.log(`✅ ${photosCount} guests have photos`);

    

    // Process matches from Edges_Top_Sim

    const matches = [];

    let filteredOut = 0;

    

    for (let i = 1; i < edgesData.length; i++) {

      const row = edgesData[i];

      const screenName1 = row[sourceCol];

      const screenName2 = row[targetCol];

      const rawSimilarity = parseFloat(row[similarityCol]);

      

      // Filter: only show matches > 0.55

      if (!screenName1 || !screenName2 || isNaN(rawSimilarity)) {

        continue;

      }

      

      if (rawSimilarity <= 0.55) {

        filteredOut++;

        continue;

      }

      

      const person1 = guestMap[screenName1];

      const person2 = guestMap[screenName2];

      

      if (!person1 || !person2) {

        Logger.log(`⚠️ Missing guest data for: ${screenName1} or ${screenName2}`);

        continue;

      }

      

      // Find shared interests

      const interests1 = person1.interests || [];

      const interests2 = person2.interests || [];

      const sharedInterests = interests1.filter(int => interests2.includes(int));

      

      // Add common traits based on other attributes

      const commonTraits = [...sharedInterests];

      

      // Check music match

      if (person1.music && person2.music && 

          person1.music !== '---' && 

          person1.music === person2.music) {

        commonTraits.push(`Music: ${person1.music}`);

      }

      

      // Check zodiac match

      if (person1.zodiac && person2.zodiac && 

          person1.zodiac !== '---' && 

          person1.zodiac === person2.zodiac) {

        commonTraits.push(`Zodiac: ${person1.zodiac}`);

      }

      

      // Add 10% for visual effect (but cap at 1.0)

      const displaySimilarity = Math.min(rawSimilarity + 0.10, 1.0);

      

      matches.push({

        person1: {

          screenName: person1.screenName,

          photoUrl: person1.photoUrl || '',

          music: person1.music || '---',

          zodiac: person1.zodiac || '---',

          interests: interests1.slice(0, 5) // Limit to first 5

        },

        person2: {

          screenName: person2.screenName,

          photoUrl: person2.photoUrl || '',

          music: person2.music || '---',

          zodiac: person2.zodiac || '---',

          interests: interests2.slice(0, 5) // Limit to first 5

        },

        similarity: displaySimilarity, // Already includes +10% boost

        sharedInterests: commonTraits.slice(0, 8) // Limit to 8 shared traits

      });

    }

    

    // Sort by similarity (highest first)

    matches.sort((a, b) => b.similarity - a.similarity);

    

    Logger.log(`✅ Processed ${matches.length} matches (filtered out ${filteredOut} below 0.55)`);

    

    // Save to sheet for reference

    if (matches.length > 0) {

      saveMatchesToSheet(matches);

    }

    

    return {

      matches: matches,

      totalGuests: guestCount,

      matchesHeader: "The System Detects Compatibility, You Two Should Talk"

    };

    

  } catch (error) {

    Logger.log('❌ Error in getCompatibilityMatches: ' + error.toString());

    Logger.log('Stack: ' + error.stack);

    return { 

      matches: [], 

      totalGuests: 0,

      error: error.toString()

    };

  }

}



/**

 * Save matches to a sheet for reference

 */

function saveMatchesToSheet(matches) {

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let matchSheet = ss.getSheetByName('Recommended_Matches');

    

    // Create sheet if it doesn't exist

    if (!matchSheet) {

      matchSheet = ss.insertSheet('Recommended_Matches');

    }

    

    // Clear existing content

    matchSheet.clear();

    

    // Set header with title

    matchSheet.getRange(1, 1).setValue('THE SYSTEM DETECTS COMPATIBILITY, YOU TWO SHOULD TALK');

    matchSheet.getRange(1, 1, 1, 6).merge();

    matchSheet.getRange(1, 1).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');

    matchSheet.getRange(1, 1).setBackground('#FF6F00').setFontColor('#FFFFFF');

    

    // Set column headers

    const headers = [

      'Person 1',

      'Person 2',

      'Compatibility %',

      'Shared Interests',

      'Person 1 Photo',

      'Person 2 Photo'

    ];

    matchSheet.getRange(2, 1, 1, headers.length).setValues([headers]);

    matchSheet.getRange(2, 1, 1, headers.length).setFontWeight('bold').setBackground('#8D6E63').setFontColor('#FFFFFF');

    

    // Add data

    const rows = matches.map(match => [

      match.person1.screenName,

      match.person2.screenName,

      Math.round(match.similarity * 100) + '%',

      match.sharedInterests.join(', '),

      match.person1.photoUrl,

      match.person2.photoUrl

    ]);

    

    if (rows.length > 0) {

      matchSheet.getRange(3, 1, rows.length, headers.length).setValues(rows);

    }

    

    // Format

    matchSheet.autoResizeColumns(1, headers.length);

    matchSheet.setFrozenRows(2);

    

    // Add alternating row colors

    for (let i = 0; i < rows.length; i++) {

      const bgColor = i % 2 === 0 ? '#FFF8E1' : '#FFFFFF';

      matchSheet.getRange(3 + i, 1, 1, headers.length).setBackground(bgColor);

    }

    

    Logger.log(`✅ Saved ${matches.length} matches to Recommended_Matches sheet`);

    

  } catch (error) {

    Logger.log('❌ Error saving matches to sheet: ' + error.toString());

  }

}



/**

 * Test function to verify getCompatibilityMatches works

 */

function testGetCompatibilityMatches() {

  Logger.log('=== Starting Compatibility Matches Test ===
');

  

  const result = getCompatibilityMatches();

  

  Logger.log('
=== RESULTS ===');

  Logger.log('Total matches found: ' + result.matches.length);

  Logger.log('Total guests: ' + result.totalGuests);

  

  if (result.matches.length > 0) {

    Logger.log('
=== TOP 5 MATCHES ===');

    for (let i = 0; i < Math.min(5, result.matches.length); i++) {

      const match = result.matches[i];

      Logger.log(`
${i + 1}. ${match.person1.screenName} ⭐ ${match.person2.screenName}`);

      Logger.log(`   💕 Compatibility: ${Math.round(match.similarity * 100)}%`);

      Logger.log(`   🎵 Music: ${match.person1.music} / ${match.person2.music}`);

      Logger.log(`   ⭐ Zodiac: ${match.person1.zodiac} / ${match.person2.zodiac}`);

      Logger.log(`   ✨ Shared: ${match.sharedInterests.slice(0, 3).join(', ')}`);

      Logger.log(`   📸 Photos: ${match.person1.photoUrl ? '✅' : '❌'} / ${match.person2.photoUrl ? '✅' : '❌'}`);

    }

  } else {

    Logger.log('
⚠️ No matches found! Check:');

    Logger.log('  - Edges_Top_Sim has data with similarity > 0.55');

    Logger.log('  - Screen names in Edges_Top_Sim match Form Responses (Clean)');

  }

  

  SpreadsheetApp.getUi().alert(

    '🎃 Compatibility Matches Test

' +

    'Total Matches: ' + result.matches.length + '
' +

    'Total Guests: ' + result.totalGuests + '

' +

    (result.matches.length > 0 ? 

      'Top Match: ' + result.matches[0].person1.screenName + ' & ' + result.matches[0].person2.screenName + 

      ' (' + Math.round(result.matches[0].similarity * 100) + '%)

' : '') +

    'Check Recommended_Matches sheet and
execution log for full details!'

  );

}



// ===== File: DDD =====


/**

 * DDD.gs - Data Disruptor Detection Script

 * 

 * DDD = Data Disruptor Detected

 * 

 * PURPOSE:

 * Analyzes "Form Responses (Clean)" to identify:

 * - Guests who didn't follow instructions

 * - Inconsistent or contradictory responses  

 * - Suspicious patterns suggesting dishonesty

 * - Joke responses or minimal effort

 * 

 * COMPATIBILITY:

 * - Works with updated Form Responses (Clean) structure (26 columns)

 * - Reads Screen Name and UID directly from Form Responses (Clean)

 * - Uses updated screen names (from web app changes)

 * 

 * CREATES:

 * - "DDD" sheet with all violation flags

 * - "DDD Report" sheet with detailed analysis

 * 

 * TO USE:

 * 1. Ensure "Form Responses (Clean)" exists (run DataClean.gs first)

 * 2. Go to Extensions > Apps Script

 * 3. Run detectDataDisruptors() function

 */



function detectDataDisruptors() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cleanSheet = ss.getSheetByName('Form Responses (Clean)');

  

  // Validation check

  if (!cleanSheet) {

    SpreadsheetApp.getUi().alert(

      'Error: "Form Responses (Clean)" sheet not found!

' +

      'Please run DataClean.gs first to create the cleaned dataset.'

    );

    return;

  }

  

  // Get data from clean sheet

  const cleanData = cleanSheet.getDataRange().getValues();

  const cleanHeaders = cleanData[0];

  const cleanRows = cleanData.slice(1);

  

  // Find Screen Name and UID columns in Form Responses (Clean)

  // These are now at columns 22-23 (1-indexed) or 21-22 (0-indexed)

  const screenNameCol = cleanHeaders.indexOf('Screen Name');

  const uidCol = cleanHeaders.indexOf('UID');

  

  if (screenNameCol === -1 || uidCol === -1) {

    SpreadsheetApp.getUi().alert(

      'Error: Could not find "Screen Name" or "UID" columns in Form Responses (Clean)!

' +

      'Please ensure DataClean.gs has been run successfully.'

    );

    return;

  }

  

  Logger.log('Screen Name column: ' + screenNameCol + ', UID column: ' + uidCol);

  

  // Define all DDD flag columns

  const dddHeaders = [

    'Screen Name',

    'UID',

    'DDD - Birthday Format',

    'DDD - More Than 3 Interests',

    'DDD - Unknown Guest',

    'DDD - No Song Request',

    'DDD - Vague Song Request',

    'DDD - Multiple Songs Listed',

    'DDD - No Artist Listed',

    'DDD - Barely Knows Host',

    'DDD - Fresh Acquaintance',

    'DDD - Host Ambiguity',

    'DDD - Age/Birthday Mismatch',

    'DDD - Education/Career Implausibility',

    'DDD - Meme/Joke Response',

    'DDD - Music/Artist Genre Mismatch',

    'DDD - Minimum Effort Detected',

    'DDD - Stranger Danger',

    'DDD - Relationship Contradiction',

    'DDD - Special Snowflake Syndrome',

    'DDD - Suspiciously Generic',

    'DDD - Composite Liar Score'

  ];

  

  // Column indices (0-based) for Form Responses (Clean)

  // These remain the same regardless of event day columns

  const COL = {

    TIMESTAMP: 0,

    BIRTHDAY: 1,

    ZODIAC: 2,

    AGE_RANGE: 3,

    EDUCATION: 4,

    ZIP: 5,

    ETHNICITY: 6,

    GENDER: 7,

    ORIENTATION: 8,

    INDUSTRY: 9,

    ROLE: 10,

    KNOW_HOST: 11,

    WHICH_HOST: 12,

    HOW_WELL: 13,

    INTERESTS: 14,

    MUSIC: 15,

    ARTIST: 16,

    SONG: 17,

    PURCHASE: 18,

    WORST_TRAIT: 19,

    SOCIAL_STANCE: 20

    // Columns 21-25 are Screen Name, UID, and event day data

    // We extract Screen Name and UID directly using their column indices

  };

  

  // Log for detailed reporting

  const detectionLog = [];

  let totalDisruptors = 0;

  

  // Process each row and generate DDD flags

  const dddResults = cleanRows.map((cleanRow, idx) => {

    const rowNum = idx + 2; // Account for header row

    

    // Extract Screen Name and UID from Form Responses (Clean)

    // NOTE: This uses the UPDATED screen names from the web app!

    const screenName = cleanRow[screenNameCol] || 'Unknown';

    const uid = cleanRow[uidCol] || 'No UID';

    

    // Analyze the clean row for violations

    const flags = analyzeSuspiciousPatterns(cleanRow, rowNum, COL, detectionLog, screenName);

    

    // Combine Screen Name, UID, and all DDD flags

    return [screenName, uid, ...flags];

  });

  

  // Create or clear DDD sheet

  let dddSheet = ss.getSheetByName('DDD');

  if (dddSheet) {

    dddSheet.clear();

  } else {

    dddSheet = ss.insertSheet('DDD');

  }

  

  // Write headers

  dddSheet.getRange(1, 1, 1, dddHeaders.length).setValues([dddHeaders]);

  

  // Write DDD results

  if (dddResults.length > 0) {

    dddSheet.getRange(2, 1, dddResults.length, dddHeaders.length).setValues(dddResults);

  }

  

  // Format DDD sheet

  formatDDDSheet(dddSheet, dddHeaders.length, dddResults.length);

  

  // Create detailed DDD report

  createDDDReport(ss, detectionLog, dddResults.length);

  

  // Count total disruptors (anyone with Liar Score > 0)

  dddResults.forEach(row => {

    const liarScore = row[row.length - 1]; // Last column

    if (liarScore > 0) {

      totalDisruptors++;

    }

  });

  

  SpreadsheetApp.getUi().alert(

    `🚨 Data Disruptor Detection Complete! 🚨

` +

    `Total Responses: ${dddResults.length}
` +

    `Data Disruptors Found: ${totalDisruptors}
` +

    `Clean Rate: ${((dddResults.length - totalDisruptors) / dddResults.length * 100).toFixed(1)}%

` +

    `Check the "DDD" and "DDD Report" sheets for details.`

  );

}



/**

 * Analyzes a single row for suspicious patterns and returns violation flags

 */

function analyzeSuspiciousPatterns(row, rowNum, COL, log, screenName) {

  const flags = [];

  let liarScore = 0;

  

  // Helper function to log violations

  function logViolation(violationType, description) {

    log.push({

      screenName: screenName,

      row: rowNum,

      violation: violationType,

      description: description

    });

  }

  

  // DDD 1: Birthday Format

  const birthday = String(row[COL.BIRTHDAY] || '').trim();

  const birthdayFlag = !birthday || !birthday.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/);

  flags.push(birthdayFlag ? 1 : 0);

  if (birthdayFlag) {

    liarScore += 1;

    logViolation('Birthday Format', `Invalid birthday format: "${birthday}"`);

  }

  

  // DDD 2: More Than 3 Interests

  const interests = String(row[COL.INTERESTS] || '').trim();

  const interestCount = interests ? interests.split(',').length : 0;

  const tooManyInterests = interestCount > 3;

  flags.push(tooManyInterests ? 1 : 0);

  if (tooManyInterests) {

    liarScore += 2;

    logViolation('More Than 3 Interests', `Listed ${interestCount} interests (max 3 allowed)`);

  }

  

  // DDD 3: Unknown Guest

  const knowHost = String(row[COL.KNOW_HOST] || '').toLowerCase().trim();

  const unknownGuest = knowHost === 'no' || knowHost === '';

  flags.push(unknownGuest ? 1 : 0);

  if (unknownGuest) {

    liarScore += 5;

    logViolation('Unknown Guest', 'Claims not to know the host');

  }

  

  // DDD 4: No Song Request

  const song = String(row[COL.SONG] || '').trim();

  const noSong = !song || song.toLowerCase() === 'none' || song.toLowerCase() === 'n/a';

  flags.push(noSong ? 1 : 0);

  if (noSong) {

    liarScore += 1;

    logViolation('No Song Request', 'Did not provide a song request');

  }

  

  // DDD 5: Vague Song Request

  const vagueSong = song && song.length > 0 && song.length < 10;

  flags.push(vagueSong ? 1 : 0);

  if (vagueSong) {

    liarScore += 1;

    logViolation('Vague Song Request', `Very short song request: "${song}"`);

  }

  

  // DDD 6: Multiple Songs Listed

  const multipleSongs = song.includes(',') || song.includes('/') || song.includes('&') || song.includes(' or ');

  flags.push(multipleSongs ? 1 : 0);

  if (multipleSongs) {

    liarScore += 1;

    logViolation('Multiple Songs Listed', `Listed multiple songs: "${song}"`);

  }

  

  // DDD 7: No Artist Listed

  const artist = String(row[COL.ARTIST] || '').trim();

  const noArtist = !artist || artist.toLowerCase() === 'none' || artist.toLowerCase() === 'n/a';

  flags.push(noArtist ? 1 : 0);

  if (noArtist && !noSong) {

    liarScore += 1;

    logViolation('No Artist Listed', 'Provided song but no artist');

  }

  

  // DDD 8: Barely Knows Host

  const howWell = String(row[COL.HOW_WELL] || '').toLowerCase().trim();

  const barelyKnows = howWell.includes('barely') || howWell.includes('not well') || howWell === 'acquaintance';

  flags.push(barelyKnows ? 1 : 0);

  if (barelyKnows) {

    liarScore += 2;

    logViolation('Barely Knows Host', `Barely knows host: "${howWell}"`);

  }

  

  // DDD 9: Fresh Acquaintance

  const freshAcquaintance = howWell.includes('just met') || howWell.includes('recently') || howWell.includes('new');

  flags.push(freshAcquaintance ? 1 : 0);

  if (freshAcquaintance) {

    liarScore += 3;

    logViolation('Fresh Acquaintance', `Just met the host: "${howWell}"`);

  }

  

  // DDD 10: Host Ambiguity

  const whichHost = String(row[COL.WHICH_HOST] || '').trim();

  const hostAmbiguous = !whichHost || whichHost.toLowerCase().includes('both') || whichHost.toLowerCase().includes('not sure');

  flags.push(hostAmbiguous ? 1 : 0);

  if (hostAmbiguous && knowHost === 'yes') {

    liarScore += 2;

    logViolation('Host Ambiguity', `Unclear which host they know: "${whichHost}"`);

  }

  

  // DDD 11: Age/Birthday Mismatch

  const ageRange = String(row[COL.AGE_RANGE] || '').trim();

  let ageMismatch = false;

  if (birthday.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {

    const birthYear = parseInt(birthday.split('/')[2]);

    const currentYear = new Date().getFullYear();

    const calculatedAge = currentYear - birthYear;

    

    if (ageRange.includes('18-24') && (calculatedAge < 18 || calculatedAge > 24)) ageMismatch = true;

    if (ageRange.includes('25-34') && (calculatedAge < 25 || calculatedAge > 34)) ageMismatch = true;

    if (ageRange.includes('35-44') && (calculatedAge < 35 || calculatedAge > 44)) ageMismatch = true;

    if (ageRange.includes('45-54') && (calculatedAge < 45 || calculatedAge > 54)) ageMismatch = true;

  }

  flags.push(ageMismatch ? 1 : 0);

  if (ageMismatch) {

    liarScore += 3;

    logViolation('Age/Birthday Mismatch', `Age range "${ageRange}" doesn't match birthday "${birthday}"`);

  }

  

  // DDD 12: Education/Career Implausibility

  const education = String(row[COL.EDUCATION] || '').toLowerCase();

  const role = String(row[COL.ROLE] || '').toLowerCase();

  const implausible = (education.includes('high school') && (role.includes('director') || role.includes('vp') || role.includes('ceo'))) ||

                       (education.includes('phd') && ageRange === '18-24');

  flags.push(implausible ? 1 : 0);

  if (implausible) {

    liarScore += 4;

    logViolation('Education/Career Implausibility', `Education "${education}" seems mismatched with role "${role}"`);

  }

  

  // DDD 13: Meme/Joke Response

  const worstTrait = String(row[COL.WORST_TRAIT] || '').toLowerCase();

  const memeKeywords = ['too awesome', 'too perfect', 'none', 'n/a', 'lol', 'lmao', 'haha', 'idk'];

  const isMeme = memeKeywords.some(keyword => worstTrait.includes(keyword));

  flags.push(isMeme ? 1 : 0);

  if (isMeme) {

    liarScore += 2;

    logViolation('Meme/Joke Response', `Joke response detected: "${worstTrait}"`);

  }

  

  // DDD 14: Music/Artist Genre Mismatch

  const musicGenre = String(row[COL.MUSIC] || '').toLowerCase();

  const genreMismatch = (musicGenre.includes('classical') && artist.toLowerCase().includes('drake')) ||

                        (musicGenre.includes('country') && artist.toLowerCase().includes('metallica')) ||

                        (musicGenre.includes('rap') && artist.toLowerCase().includes('mozart'));

  flags.push(genreMismatch ? 1 : 0);

  if (genreMismatch) {

    liarScore += 2;

    logViolation('Music/Artist Genre Mismatch', `Genre "${musicGenre}" doesn't match artist "${artist}"`);

  }

  

  // DDD 15: Minimum Effort Detected

  const totalChars = String(row.join('')).length;

  const minEffort = totalChars < 100 || interests.length < 5 || worstTrait.length < 5;

  flags.push(minEffort ? 1 : 0);

  if (minEffort) {

    liarScore += 2;

    logViolation('Minimum Effort Detected', `Very short responses (${totalChars} chars total)`);

  }

  

  // DDD 16: Stranger Danger

  const strangerDanger = unknownGuest && (barelyKnows || freshAcquaintance);

  flags.push(strangerDanger ? 1 : 0);

  if (strangerDanger) {

    liarScore += 5;

    logViolation('Stranger Danger', 'Multiple indicators of not knowing host');

  }

  

  // DDD 17: Relationship Contradiction

  const contradiction = (knowHost === 'yes' && (howWell.includes('never') || howWell.includes("don't"))) ||

                        (knowHost === 'no' && howWell.includes('well') || howWell.includes('close'));

  flags.push(contradiction ? 1 : 0);

  if (contradiction) {

    liarScore += 4;

    logViolation('Relationship Contradiction', `"${knowHost}" contradicts "${howWell}"`);

  }

  

  // DDD 18: Special Snowflake Syndrome

  const specialSnowflake = interests.toLowerCase().includes('unicorn') || 

                           interests.toLowerCase().includes('quantum') ||

                           worstTrait.toLowerCase().includes('care too much') ||

                           worstTrait.toLowerCase().includes('too nice');

  flags.push(specialSnowflake ? 1 : 0);

  if (specialSnowflake) {

    liarScore += 1;

    logViolation('Special Snowflake Syndrome', 'Overly unique or self-aggrandizing responses');

  }

  

  // DDD 19: Suspiciously Generic

  const genericKeywords = ['stuff', 'things', 'whatever', 'anything', 'nothing special'];

  const tooGeneric = genericKeywords.some(keyword => 

    interests.includes(keyword) || worstTrait.includes(keyword)

  );

  flags.push(tooGeneric ? 1 : 0);

  if (tooGeneric) {

    liarScore += 1;

    logViolation('Suspiciously Generic', 'Extremely vague or generic responses');

  }

  

  // DDD 20: Composite Liar Score

  flags.push(liarScore);

  

  return flags;

}



/**

 * Formats the DDD sheet for readability

 */

function formatDDDSheet(sheet, numCols, numRows) {

  // Freeze header row and first 2 columns

  sheet.setFrozenRows(1);

  sheet.setFrozenColumns(2);

  

  // Format header

  const headerRange = sheet.getRange(1, 1, 1, numCols);

  headerRange.setFontWeight('bold');

  headerRange.setBackground('#434343');

  headerRange.setFontColor('#ffffff');

  headerRange.setHorizontalAlignment('center');

  

  // Set column widths

  sheet.setColumnWidth(1, 150); // Screen Name

  sheet.setColumnWidth(2, 100); // UID

  for (let i = 3; i <= numCols - 1; i++) {

    sheet.setColumnWidth(i, 80); // DDD flags

  }

  sheet.setColumnWidth(numCols, 120); // Liar Score

  

  // Conditional formatting for flags

  if (numRows > 0) {

    for (let col = 3; col <= numCols; col++) {

      const range = sheet.getRange(2, col, numRows, 1);

      

      // Highlight violations in red

      const rule = SpreadsheetApp.newConditionalFormatRule()

        .whenNumberGreaterThan(0)

        .setBackground('#f4cccc')

        .setRanges([range])

        .build();

      

      const rules = sheet.getConditionalFormatRules();

      rules.push(rule);

      sheet.setConditionalFormatRules(rules);

    }

    

    // Special formatting for Liar Score (last column)

    const scoreRange = sheet.getRange(2, numCols, numRows, 1);

    scoreRange.setFontWeight('bold');

    

    // Color code liar scores

    const scoreRules = [

      SpreadsheetApp.newConditionalFormatRule()

        .whenNumberGreaterThanOrEqualTo(10)

        .setBackground('#cc0000')

        .setFontColor('#ffffff')

        .setRanges([scoreRange])

        .build(),

      SpreadsheetApp.newConditionalFormatRule()

        .whenNumberBetween(5, 9)

        .setBackground('#ff9900')

        .setRanges([scoreRange])

        .build(),

      SpreadsheetApp.newConditionalFormatRule()

        .whenNumberBetween(1, 4)

        .setBackground('#ffff00')

        .setRanges([scoreRange])

        .build()

    ];

    

    const allRules = sheet.getConditionalFormatRules();

    sheet.setConditionalFormatRules(allRules.concat(scoreRules));

  }

  

  // Auto-resize all columns

  sheet.autoResizeColumns(1, numCols);

}



/**

 * Creates a detailed DDD report sheet

 */

function createDDDReport(ss, log, totalResponses) {

  // Support direct runs without parameters

  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  log = Array.isArray(log) ? log : [];

  totalResponses = (typeof totalResponses === 'number') ? totalResponses : log.length;



  // Create or clear report sheet

  let reportSheet = ss.getSheetByName('DDD Report');

  if (reportSheet) {

    reportSheet.clear();

  } else {

    reportSheet = ss.insertSheet('DDD Report');

  }



  // Write report header - ensure every row has 4 columns

  const reportData = [

    ['🚨 DATA DISRUPTOR DETECTION REPORT 🚨', '', '', ''],

    ['Generated: ' + new Date().toLocaleString(), '', '', ''],

    ['Total Responses Analyzed: ' + totalResponses, '', '', ''],

    ['Total Violations Detected: ' + log.length, '', '', ''],

    ['', '', '', ''],

    ['VIOLATION LOG', '', '', ''],

    ['Screen Name', 'Row', 'Violation Type', 'Description']

  ];



  // Add violation details

  log.forEach(function(entry) {

    reportData.push([

      entry.screenName,

      entry.row,

      entry.violation,

      entry.description

    ]);

  });



  // Write data - now all rows have exactly 4 columns

  reportSheet.getRange(1, 1, reportData.length, 4).setValues(reportData);



  // Format report

  reportSheet.getRange(1, 1, 1, 4).merge()

    .setFontSize(16)

    .setFontWeight('bold')

    .setHorizontalAlignment('center')

    .setBackground('#cc0000')

    .setFontColor('#ffffff');



  reportSheet.getRange(2, 1, 4, 4).setFontStyle('italic');



  reportSheet.getRange(6, 1, 1, 4).merge()

    .setFontSize(14)

    .setFontWeight('bold')

    .setHorizontalAlignment('center')

    .setBackground('#434343')

    .setFontColor('#ffffff');



  reportSheet.getRange(7, 1, 1, 4)

    .setFontWeight('bold')

    .setBackground('#666666')

    .setFontColor('#ffffff');



  // Set column widths

  reportSheet.setColumnWidth(1, 150);

  reportSheet.setColumnWidth(2, 60);

  reportSheet.setColumnWidth(3, 200);

  reportSheet.setColumnWidth(4, 400);



  // Freeze header rows

  reportSheet.setFrozenRows(7);



  // Add alternating row colors

  if (log.length > 0) {

    const dataRange = reportSheet.getRange(8, 1, log.length, 4);

    const rule = SpreadsheetApp.newConditionalFormatRule()

      .whenFormulaSatisfied('=ISEVEN(ROW())')

      .setBackground('#f3f3f3')

      .setRanges([dataRange])

      .build();



    reportSheet.setConditionalFormatRules([rule]);

  }

}

/**

 * Creates a detailed DDD report sheet - FIXED VERSION

 */

function createDDDReport(ss, log, totalResponses) {

  // Create or clear report sheet

  let reportSheet = ss.getSheetByName('DDD Report');

  if (reportSheet) {

    reportSheet.clear();

  } else {

    reportSheet = ss.insertSheet('DDD Report');

  }

  

  // Write report header - ALL ROWS MUST HAVE 4 COLUMNS

  const reportData = [

    ['🚨 DATA DISRUPTOR DETECTION REPORT 🚨', '', '', ''],

    ['Generated: ' + new Date().toLocaleString(), '', '', ''],

    ['Total Responses Analyzed: ' + totalResponses, '', '', ''],

    ['Total Violations Detected: ' + log.length, '', '', ''],

    ['', '', '', ''],

    ['VIOLATION LOG', '', '', ''],

    ['Screen Name', 'Row', 'Violation Type', 'Description']

  ];

  

  // Add violation details

  log.forEach(function(entry) {

    reportData.push([

      entry.screenName,

      entry.row,

      entry.violation,

      entry.description

    ]);

  });

  

  // Write data - now all rows have exactly 4 columns

  reportSheet.getRange(1, 1, reportData.length, 4).setValues(reportData);

  

  // Format report

  reportSheet.getRange(1, 1, 1, 4).merge()

    .setFontSize(16)

    .setFontWeight('bold')

    .setHorizontalAlignment('center')

    .setBackground('#cc0000')

    .setFontColor('#ffffff');

  

  reportSheet.getRange(2, 1, 4, 4).setFontStyle('italic');

  

  reportSheet.getRange(6, 1, 1, 4).merge()

    .setFontSize(14)

    .setFontWeight('bold')

    .setHorizontalAlignment('center')

    .setBackground('#434343')

    .setFontColor('#ffffff');

  

  reportSheet.getRange(7, 1, 1, 4)

    .setFontWeight('bold')

    .setBackground('#666666')

    .setFontColor('#ffffff');

  

  // Set column widths

  reportSheet.setColumnWidth(1, 150);

  reportSheet.setColumnWidth(2, 60);

  reportSheet.setColumnWidth(3, 200);

  reportSheet.setColumnWidth(4, 400);

  

  // Freeze header rows

  reportSheet.setFrozenRows(7);

  

  // Add alternating row colors

  if (log.length > 0) {

    const dataRange = reportSheet.getRange(8, 1, log.length, 4);

    const rule = SpreadsheetApp.newConditionalFormatRule()

      .whenFormulaSatisfied('=ISEVEN(ROW())')

      .setBackground('#f3f3f3')

      .setRanges([dataRange])

      .build();

    

    reportSheet.setConditionalFormatRules([rule]);

  }

}



// ===== File: DataClean =====


/**

 * ============================================================================

 * DDD.gs - Data Disruptor Detection Script

 * ============================================================================

 * * DDD = Data Disruptor Detected

 * * PURPOSE:

 * Analyzes "Form Responses (Clean)" to identify:

 * - Guests who didn't follow instructions

 * - Inconsistent or contradictory responses  

 * - Suspicious patterns suggesting dishonesty

 * - Joke responses or minimal effort

 * * COMPATIBILITY:

 * - Works with updated Form Responses (Clean) structure (26 columns)

 * - Reads UID directly from Form Responses (Clean)

 * * CREATES:

 * - "DDD" sheet with all violation flags

 * - "DDD Report" sheet with detailed analysis (UID ONLY)

 * * TO USE:

 * 1. Ensure "Form Responses (Clean)" exists (run DataClean.gs first)

 * 2. Go to Extensions > Apps Script

 * 3. Run detectDataDisruptors() function

 */



function detectDataDisruptors() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cleanSheet = ss.getSheetByName('Form Responses (Clean)');

  

  // Validation check

  if (!cleanSheet) {

    SpreadsheetApp.getUi().alert(

      'Error: "Form Responses (Clean)" sheet not found!

' +

      'Please run DataClean.gs first to create the cleaned dataset.'

    );

    return;

  }

  

  // Get data from clean sheet

  const cleanData = cleanSheet.getDataRange().getValues();

  const cleanHeaders = cleanData[0];

  const cleanRows = cleanData.slice(1);

  

  // Find Screen Name and UID columns in Form Responses (Clean)

  const screenNameCol = cleanHeaders.indexOf('Screen Name');

  const uidCol = cleanHeaders.indexOf('UID');

  

  if (screenNameCol === -1 || uidCol === -1) {

    SpreadsheetApp.getUi().alert(

      'Error: Could not find "Screen Name" or "UID" columns in Form Responses (Clean)!

' +

      'Please ensure DataClean.gs has been run successfully.'

    );

    return;

  }

  

  Logger.log('Screen Name column: ' + screenNameCol + ', UID column: ' + uidCol);

  

  // Define all DDD flag columns

  const dddHeaders = [

    'UID', // CHANGED: UID is the primary identifier (Col 1)

    'DDD - Birthday Format',

    'DDD - More Than 3 Interests',

    'DDD - Unknown Guest',

    'DDD - No Song Request',

    'DDD - Vague Song Request',

    'DDD - Multiple Songs Listed',

    'DDD - No Artist Listed',

    'DDD - Barely Knows Host',

    'DDD - Fresh Acquaintance',

    'DDD - Host Ambiguity',

    'DDD - Age/Birthday Mismatch',

    'DDD - Education/Career Implausibility',

    'DDD - Meme/Joke Response',

    'DDD - Music/Artist Genre Mismatch',

    'DDD - Minimum Effort Detected',

    'DDD - Stranger Danger',

    'DDD - Relationship Contradiction',

    'DDD - Special Snowflake Syndrome',

    // 'DDD - Suspiciously Generic', // REMOVED TO BALANCE ARRAY SIZE

    'DDD - Too Many Violations Detected', // NEW: Flag for high composite score (Col 21)

    'DDD - Composite Liar Score'         // Col 22

  ];

  

  // Column indices (0-based) for Form Responses (Clean)

  const COL = {

    TIMESTAMP: 0,

    BIRTHDAY: 1,

    ZODIAC: 2,

    AGE_RANGE: 3,

    EDUCATION: 4,

    ZIP: 5,

    ETHNICITY: 6,

    GENDER: 7,

    ORIENTATION: 8,

    INDUSTRY: 9,

    ROLE: 10,

    KNOW_HOST: 11,

    WHICH_HOST: 12,

    HOW_WELL: 13,

    INTERESTS: 14,

    MUSIC: 15,

    ARTIST: 16,

    SONG: 17,

    PURCHASE: 18,

    WORST_TRAIT: 19,

    SOCIAL_STANCE: 20

    // Columns 21-25 are Screen Name, UID, and event day data

  };

  

  // Log for detailed reporting

  const detectionLog = [];

  let totalDisruptors = 0;

  

  // Process each row and generate DDD flags

  const dddResults = cleanRows.map((cleanRow, idx) => {

    const rowNum = idx + 2; // Account for header row

    

    // Extract UID from Form Responses (Clean)

    const uid = cleanRow[uidCol] || 'No UID';

    

    // Analyze the clean row for violations - PASSING UID INSTEAD OF SCREEN NAME

    const flags = analyzeSuspiciousPatterns(cleanRow, rowNum, COL, detectionLog, uid);

    

    // Combine UID and all DDD flags (flags array is 21 items long)

    // FIX: Using .concat() instead of spread operator for safer array merging

    return [uid].concat(flags); 

  });

  

  // Create or clear DDD sheet

  let dddSheet = ss.getSheetByName('DDD');

  if (dddSheet) {

    dddSheet.clear();

  } else {

    dddSheet = ss.insertSheet('DDD');

  }

  

  // Write headers

  // Line 138: No change needed here.

  dddSheet.getRange(1, 1, 1, dddHeaders.length).setValues([dddHeaders]);

  

  // Write DDD results

  if (dddResults.length > 0) {

    // Line 141: This range now correctly uses dddHeaders.length (22)

    dddSheet.getRange(2, 1, dddResults.length, dddHeaders.length).setValues(dddResults);

  }

  

  // Format DDD sheet

  formatDDDSheet(dddSheet, dddHeaders.length, dddResults.length);

  

  // Create detailed DDD report

  createDDDReport(ss, detectionLog, dddResults.length);

  

  // Count total disruptors (anyone with Liar Score > 0)

  dddResults.forEach(row => {

    const liarScore = row[row.length - 1]; // Last column

    if (liarScore > 0) {

      totalDisruptors++;

    }

  });

  

  SpreadsheetApp.getUi().alert(

    `🚨 Data Disruptor Detection Complete! 🚨

` +

    `Total Responses: ${dddResults.length}
` +

    `Data Disruptors Found: ${totalDisruptors}
` +

    `Clean Rate: ${((dddResults.length - totalDisruptors) / dddResults.length * 100).toFixed(1)}%

` +

    `Check the "DDD" and "DDD Report" sheets for details.`

  );

}



/**

 * Analyzes a single row for suspicious patterns and returns violation flags

 */

function analyzeSuspiciousPatterns(row, rowNum, COL, log, uid) {

  // FIX: Initialize flags array to 21 elements (0-20) to prevent size mismatch errors.

  // The structure is 20 binary flags (index 0-19) + 1 Liar Score (index 20).

  const flags = new Array(21).fill(0); 

  let liarScore = 0;

  

  // Helper function to log violations - NOW USING UID

  function logViolation(violationType, description) {

    log.push({

      uid: uid, // CHANGED: Using UID instead of screenName

      row: rowNum,

      violation: violationType,

      description: description

    });

  }

  

  // --- DDD 1: Birthday Format (Index 0) ---

  const birthday = String(row[COL.BIRTHDAY] || '').trim();

  const birthdayFlag = !birthday || !birthday.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/);

  flags[0] = birthdayFlag ? 1 : 0;

  if (birthdayFlag) {

    liarScore += 1;

    logViolation('Birthday Format', `Invalid birthday format: "${birthday}"`);

  }

  

  // --- DDD 2: More Than 3 Interests (Index 1) ---

  const interests = String(row[COL.INTERESTS] || '').trim();

  const interestCount = interests ? interests.split(/[;,]/).length : 0;

  const tooManyInterests = interestCount > 3;

  flags[1] = tooManyInterests ? 1 : 0;

  if (tooManyInterests) {

    liarScore += 2;

    logViolation('More Than 3 Interests', `Listed ${interestCount} interests (max 3 allowed)`);

  }

  

  // --- DDD 3: Unknown Guest (Index 2) ---

  const knowHost = String(row[COL.KNOW_HOST] || '').toLowerCase().trim();

  const unknownGuest = knowHost === 'no' || knowHost === '';

  flags[2] = unknownGuest ? 1 : 0;

  if (unknownGuest) {

    liarScore += 5;

    logViolation('Unknown Guest', 'Claims not to know the host');

  }

  

  // --- DDD 4: No Song Request (Index 3) ---

  const song = String(row[COL.SONG] || '').trim();

  const noSong = !song || song.toLowerCase() === 'none' || song.toLowerCase() === 'n/a';

  flags[3] = noSong ? 1 : 0;

  if (noSong) {

    liarScore += 1;

    logViolation('No Song Request', 'Did not provide a song request');

  }

  

  // --- DDD 5: Vague Song Request (Index 4) ---

  const vagueSong = song && song.length > 0 && song.length < 10;

  flags[4] = vagueSong ? 1 : 0;

  if (vagueSong) {

    liarScore += 1;

    logViolation('Vague Song Request', `Very short song request: "${song}"`);

  }

  

  // --- DDD 6: Multiple Songs Listed (Index 5) ---

  const multipleSongs = song.includes(',') || song.includes('/') || song.includes('&') || song.includes(' or ');

  flags[5] = multipleSongs ? 1 : 0;

  if (multipleSongs) {

    liarScore += 1;

    logViolation('Multiple Songs Listed', `Listed multiple songs: "${song}"`);

  }

  

  // --- DDD 7: No Artist Listed (Index 6) ---

  const artist = String(row[COL.ARTIST] || '').trim();

  const noArtist = !artist || artist.toLowerCase() === 'none' || artist.toLowerCase() === 'n/a';

  flags[6] = noArtist ? 1 : 0;

  if (noArtist && !noSong) {

    liarScore += 1;

    logViolation('No Artist Listed', 'Provided song but no artist');

  }

  

  // --- DDD 8: Barely Knows Host (Index 7) ---

  const howWell = String(row[COL.HOW_WELL] || '').toLowerCase().trim();

  const barelyKnows = howWell.includes('barely') || howWell.includes('not well') || howWell === 'acquaintance';

  flags[7] = barelyKnows ? 1 : 0;

  if (barelyKnows) {

    liarScore += 2;

    logViolation('Barely Knows Host', `Barely knows host: "${howWell}"`);

  }

  

  // --- DDD 9: Fresh Acquaintance (Index 8) ---

  const freshAcquaintance = howWell.includes('just met') || howWell.includes('recently') || howWell.includes('new');

  flags[8] = freshAcquaintance ? 1 : 0;

  if (freshAcquaintance) {

    liarScore += 3;

    logViolation('Fresh Acquaintance', `Just met the host: "${howWell}"`);

  }

  

  // --- DDD 10: Host Ambiguity (Index 9) ---

  const whichHost = String(row[COL.WHICH_HOST] || '').trim();

  const hostAmbiguous = !whichHost || whichHost.toLowerCase().includes('both') || whichHost.toLowerCase().includes('not sure');

  flags[9] = hostAmbiguous ? 1 : 0;

  if (hostAmbiguous && knowHost.includes('yes')) {

    liarScore += 2;

    logViolation('Host Ambiguity', `Unclear which host they know: "${whichHost}"`);

  }

  

  // --- DDD 11: Age/Birthday Mismatch (Index 10) ---

  const ageRange = String(row[COL.AGE_RANGE] || '').trim();

  let ageMismatch = false;

  if (birthday.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {

    const birthYear = parseInt(birthday.split('/')[2]);

    const currentYear = new Date().getFullYear();

    const calculatedAge = currentYear - birthYear;

    

    // Check if the calculated age falls outside the stated age range

    const rangeMatch = ageRange.match(/(\d+)-(\d+)/);

    if (rangeMatch) {

      const minAge = parseInt(rangeMatch[1]);

      const maxAge = parseInt(rangeMatch[2]);

      if (calculatedAge < minAge || calculatedAge > maxAge) ageMismatch = true;

    }

  }

  flags[10] = ageMismatch ? 1 : 0;

  if (ageMismatch) {

    liarScore += 3;

    logViolation('Age/Birthday Mismatch', `Age range "${ageRange}" doesn't match birthday "${birthday}"`);

  }

  

  // --- DDD 12: Education/Career Implausibility (Index 11) ---

  const education = String(row[COL.EDUCATION] || '').toLowerCase();

  const role = String(row[COL.ROLE] || '').toLowerCase();

  const implausible = (education.includes('high school') && (role.includes('director') || role.includes('vp') || role.includes('ceo'))) ||

    (education.includes('phd') && ageRange.includes('18-24')) ||

    (education.includes('masters') && ageRange.includes('18-24')); // Added Masters/young age check

  flags[11] = implausible ? 1 : 0;

  if (implausible) {

    liarScore += 4;

    logViolation('Education/Career Implausibility', `Education "${education}" seems mismatched with role "${role}"`);

  }

  

  // --- DDD 13: Meme/Joke Response (Index 12) ---

  const worstTrait = String(row[COL.WORST_TRAIT] || '').toLowerCase();

  const memeKeywords = ['too awesome', 'too perfect', 'none', 'n/a', 'lol', 'lmao', 'haha', 'idk'];

  const isMeme = memeKeywords.some(keyword => worstTrait.includes(keyword));

  flags[12] = isMeme ? 1 : 0;

  if (isMeme) {

    liarScore += 2;

    logViolation('Meme/Joke Response', `Joke response detected: "${worstTrait}"`);

  }

  

  // --- DDD 14: Music/Artist Genre Mismatch (Index 13) ---

  const musicGenre = String(row[COL.MUSIC] || '').toLowerCase();

  const genreMismatch = (musicGenre.includes('classical') && artist.toLowerCase().includes('drake')) ||

    (musicGenre.includes('country') && artist.toLowerCase().includes('metallica')) ||

    (musicGenre.includes('rap') && artist.toLowerCase().includes('mozart')) ||

    (musicGenre.includes('electronic') && artist.toLowerCase().includes('adele')); // Added one more mismatch

  flags[13] = genreMismatch ? 1 : 0;

  if (genreMismatch) {

    liarScore += 2;

    logViolation('Music/Artist Genre Mismatch', `Genre "${musicGenre}" doesn't match artist "${artist}"`);

  }

  

  // --- DDD 15: Minimum Effort Detected (Index 14) ---

  const totalChars = String(row.slice(COL.AGE_RANGE, COL.SOCIAL_STANCE + 1).join('')).length; // Focus on core questions

  const minEffort = totalChars < 150 || interests.length < 5 || worstTrait.length < 5; // Increased total char threshold

  flags[14] = minEffort ? 1 : 0;

  if (minEffort) {

    liarScore += 2;

    logViolation('Minimum Effort Detected', `Very short responses (${totalChars} chars total)`);

  }

  

  // --- DDD 16: Stranger Danger (Index 15) ---

  const strangerDanger = unknownGuest && (barelyKnows || freshAcquaintance);

  flags[15] = strangerDanger ? 1 : 0;

  if (strangerDanger) {

    liarScore += 5;

    logViolation('Stranger Danger', 'Multiple indicators of not knowing host');

  }

  

  // --- DDD 17: Relationship Contradiction (Index 16) ---

  const contradiction = (knowHost.includes('yes') && (howWell.includes('never') || howWell.includes("don't"))) ||

    (knowHost.includes('no') && howWell.includes('well') || howWell.includes('close'));

  flags[16] = contradiction ? 1 : 0; 

  if (contradiction) {

    liarScore += 4;

    logViolation('Relationship Contradiction', `"${knowHost}" contradicts "${howWell}"`);

  }

  

  // --- DDD 18: Special Snowflake Syndrome (Index 17) ---

  const specialSnowflake = interests.toLowerCase().includes('unicorn') || 

    interests.toLowerCase().includes('quantum') ||

    worstTrait.toLowerCase().includes('care too much') ||

    worstTrait.toLowerCase().includes('too nice');

  flags[17] = specialSnowflake ? 1 : 0;

  if (specialSnowflake) {

    liarScore += 1;

    logViolation('Special Snowflake Syndrome', 'Overly unique or self-aggrandizing responses');

  }

  

  // --- DDD 19: Suspiciously Generic (Index 18) ---

  const genericKeywords = ['stuff', 'things', 'whatever', 'anything', 'nothing special'];

  const tooGeneric = genericKeywords.some(keyword => 

    interests.includes(keyword) || worstTrait.includes(keyword) || song.includes(keyword)

  );

  flags[18] = tooGeneric ? 1 : 0;

  // NOTE: DDD 19 (Suspiciously Generic) is now index 18

  

  if (tooGeneric) {

    liarScore += 1;

    logViolation('Suspiciously Generic', 'Extremely vague or generic responses');

  }



  // --- DDD 20: Too Many Violations Detected (Index 19) ---

  const tooManyViolations = liarScore >= 5;

  flags[19] = tooManyViolations ? 1 : 0;

  if (tooManyViolations) {

    // Log this as a severe composite warning

    logViolation('Too Many Violations Detected', `Composite score of ${liarScore} (5+ is high-risk)`);

  }

  

  // --- DDD 21: Composite Liar Score (Index 20) ---

  flags[20] = liarScore;

  

  return flags;

}



/**

 * Formats the DDD sheet for readability

 */

function formatDDDSheet(sheet, numCols, numRows) {

  // Freeze header row and first column (UID)

  sheet.setFrozenRows(1);

  sheet.setFrozenColumns(1); // Now only freezes the UID column

  

  // Format header

  const headerRange = sheet.getRange(1, 1, 1, numCols);

  headerRange.setFontWeight('bold');

  headerRange.setBackground('#434343');

  headerRange.setFontColor('#ffffff');

  headerRange.setHorizontalAlignment('center');

  

  // Set column widths

  sheet.setColumnWidth(1, 100); // UID

  // The new column (21) is before the Liar Score (22)

  for (let i = 2; i <= numCols - 2; i++) { // DDD flags start at Col 2 and end before the last two columns

    sheet.setColumnWidth(i, 80); // DDD flags

  }

  sheet.setColumnWidth(numCols - 1, 120); // Too Many Violations

  sheet.setColumnWidth(numCols, 120); // Liar Score

  

  // Conditional formatting for flags

  if (numRows > 0) {

    // Apply conditional formatting to all flag columns (Col 2 up to numCols - 1, which is Liar Score)

    for (let col = 2; col <= numCols - 1; col++) { 

      const range = sheet.getRange(2, col, numRows, 1);

      

      // Highlight violations in red

      const rule = SpreadsheetApp.newConditionalFormatRule()

        .whenNumberGreaterThan(0)

        .setBackground('#f4cccc')

        .setRanges([range])

        .build();

      

      const rules = sheet.getConditionalFormatRules();

      rules.push(rule);

      sheet.setConditionalFormatRules(rules);

    }

    

    // Special formatting for Liar Score (last column)

    const scoreRange = sheet.getRange(2, numCols, numRows, 1);

    scoreRange.setFontWeight('bold');

    

    // Color code liar scores

    const scoreRules = [

      SpreadsheetApp.newConditionalFormatRule()

        .whenNumberGreaterThanOrEqualTo(10)

        .setBackground('#cc0000')

        .setFontColor('#ffffff')

        .setRanges([scoreRange])

        .build(),

      SpreadsheetApp.newConditionalFormatRule()

        .whenNumberBetween(5, 9)

        .setBackground('#ff9900')

        .setRanges([scoreRange])

        .build(),

      SpreadsheetApp.newConditionalFormatRule()

        .whenNumberBetween(1, 4)

        .setBackground('#ffff00')

        .setRanges([scoreRange])

        .build()

    ];

    

    const allRules = sheet.getConditionalFormatRules();

    sheet.setConditionalFormatRules(allRules.concat(scoreRules));

  }

  

  // Auto-resize all columns

  sheet.autoResizeColumns(1, numCols);

}



/**

 * Creates a detailed DDD report sheet - UPDATED TO USE UID

 */

function createDDDReport(ss, log, totalResponses) {

  // Create or clear report sheet

  let reportSheet = ss.getSheetByName('DDD Report');

  if (reportSheet) {

    reportSheet.clear();

  } else {

    reportSheet = ss.insertSheet('DDD Report');

  }

  

  // Write report header - ALL ROWS MUST HAVE 4 COLUMNS

  const reportData = [

    ['🚨 DATA DISRUPTOR DETECTION REPORT 🚨', '', '', ''],

    ['Generated: ' + new Date().toLocaleString(), '', '', ''],

    ['Total Responses Analyzed: ' + totalResponses, '', '', ''],

    ['Total Violations Detected: ' + log.length, '', '', ''],

    ['', '', '', ''],

    ['VIOLATION LOG', '', '', ''],

    ['UID', 'Row', 'Violation Type', 'Description'] // CHANGED: Header is now UID

  ];

  

  // Add violation details

  log.forEach(function(entry) {

    reportData.push([

      entry.uid, // Use entry.uid

      entry.row,

      entry.violation,

      entry.description

    ]);

  });

  

  // Write data - now all rows have exactly 4 columns

  reportSheet.getRange(1, 1, reportData.length, 4).setValues(reportData);

  

  // Format report

  reportSheet.getRange(1, 1, 1, 4).merge()

    .setFontSize(16)

    .setFontWeight('bold')

    .setHorizontalAlignment('center')

    .setBackground('#cc0000')

    .setFontColor('#ffffff');

  

  reportSheet.getRange(2, 1, 4, 4).setFontStyle('italic');

  

  reportSheet.getRange(6, 1, 1, 4).merge()

    .setFontSize(14)

    .setFontWeight('bold')

    .setHorizontalAlignment('center')

    .setBackground('#434343')

    .setFontColor('#ffffff');

  

  reportSheet.getRange(7, 1, 1, 4)

    .setFontWeight('bold')

    .setBackground('#666666')

    .setFontColor('#ffffff');

  

  // Set column widths

  reportSheet.setColumnWidth(1, 100); // UID

  reportSheet.setColumnWidth(2, 60);

  reportSheet.setColumnWidth(3, 200);

  reportSheet.setColumnWidth(4, 400);

  

  // Freeze header rows

  reportSheet.setFrozenRows(7);

  

  // Add alternating row colors

  if (log.length > 0) {

    const dataRange = reportSheet.getRange(8, 1, log.length, 4);

    const rule = SpreadsheetApp.newConditionalFormatRule()

      .whenFormulaSatisfied('=ISEVEN(ROW())')

      .setBackground('#f3f3f3')

      .setRanges([dataRange])

      .build();

    

    reportSheet.setConditionalFormatRules([rule]);

  }

}



// ===== File: Export_CSV =====


// ============================================

// ADMIN UTILITIES - EXPORT_CSV.GS

// ============================================

// Purpose: Contains functions for exporting spreadsheet data.

// - Creates an 'Admin Tools' menu on spreadsheet open.

// - Exports all sheets as CSV files to a FIXED Drive folder.

// ============================================



// GLOBAL CONSTANT

// The ID of the target Google Drive folder (ID extracted from the URL).

const TARGET_FOLDER_ID = '1iR4UQ1V8Iy4IIbqN5Ifzi6zUNt8bRueB'; 



// ============================================

// MENU AND SETUP

// ============================================



/**

 * Creates a custom menu in the spreadsheet UI for admin tasks.

 * This runs automatically when the spreadsheet is opened.

 */

function onOpen() {

  SpreadsheetApp.getUi()

      .createMenu('Admin Tools')

      .addItem('Export All Sheets as CSV', 'exportAllSheetsAsCsv')

      .addToUi();

}



/**

 * Retrieves the dedicated Drive folder using the fixed ID.

 * Throws an error if the folder is not found or not accessible.

 * @returns {GoogleAppsScript.Drive.Folder} The export folder.

 */

function getExportFolder_() {

  try {

    const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);

    Logger.log('Using fixed Drive folder: ' + folder.getName());

    return folder;

  } catch (e) {

    Logger.log('ERROR: Target folder not found or accessible with ID: ' + TARGET_FOLDER_ID + '. Error: ' + e.toString());

    throw new Error('Target Drive folder not found or accessible. Please check the ID and permissions.');

  }

}



// ============================================

// EXPORT LOGIC

// ============================================



/**

 * Exports all sheets in the active spreadsheet as individual CSV files 

 * to the fixed target folder in Google Drive. If a file exists, it is replaced,

 * ensuring the latest data is exported.

 * * @returns {void} The function now returns nothing for a cleaner exit.

 */

function exportAllSheetsAsCsv() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = ss.getSheets();

  

  // Use the updated function to get the fixed folder

  const exportFolder = getExportFolder_(); 

  

  let successCount = 0;

  let errorCount = 0;

  

  Logger.log('--- STARTING CSV EXPORT TO DRIVE FOLDER: ' + exportFolder.getName() + ' ---');



  sheets.forEach(sheet => {

    const sheetName = sheet.getName();

    const fileName = `${sheetName}.csv`;

    

    try {

      // 1. Get sheet data and generate CSV content

      const data = sheet.getDataRange().getValues();

      if (data.length === 0) {

        Logger.log(`Skipping sheet "${sheetName}": No data.`);

        return;

      }

      

      // Manually escape quotes and join data for robust CSV format

      const csvContent = data.map(row => 

        row.map(cell => {

          let value = String(cell).replace(/"/g, '""');

          // Add newline handling to ensure content within quotes is preserved

          if (value.includes('
')) {

              value = value.replace(/
/g, ' '); // Replace internal newlines with space

          }

          return `"${value}"`;

        }).join(',')

      ).join('
');

      

      // 2. Create Blob

      const blob = Utilities.newBlob(csvContent, MimeType.CSV, fileName);



      // 3. Check for existing file

      const files = exportFolder.getFilesByName(fileName);

      

      if (files.hasNext()) {

        // --- FIX IMPLEMENTED HERE: Use Drive API to update file content/MIME type ---

        const existingFile = files.next();

        

        // This is a more reliable way to update content and maintain MIME type

        // The Drive service is implicitly available if DriveApp is used.

        Drive.Files.update({ 

            mimeType: MimeType.CSV 

          }, 

          existingFile.getId(), 

          blob

        );

        Logger.log(`[UPDATED] Sheet "${sheetName}" saved to existing file: ${existingFile.getUrl()}`);

        // --------------------------------------------------------------------------

      } else {

        // Create new file

        exportFolder.createFile(blob);

        Logger.log(`[CREATED] Sheet "${sheetName}" saved to new file.`);

      }

      

      successCount++;

    } catch (e) {

      // Catch errors, particularly if the Drive API fails or an API call limit is hit

      Logger.log(`[ERROR] Failed to export sheet "${sheetName}": ${e.toString()}`);

      errorCount++;

    }

  });



  Logger.log('--- CSV EXPORT COMPLETE ---');

  

  // Final Alert to user

  SpreadsheetApp.getUi().alert(

    'CSV Export Complete', 

    `All sheets successfully exported to Drive:

` +

    `Folder: ${exportFolder.getName()}
` +

    `Successes: ${successCount}
` +

    `Failures: ${errorCount}

` +

    `Check the script logs (View > Logs) for links and details.`,

    SpreadsheetApp.getUi().ButtonSet.OK

  );

}



// ===== File: Pan_Validation =====


/**

 * ============================================================================

 * PAN_VALIDATION.GS - Setup Validator & Diagnostic Tools

 * ============================================================================

 * 

 * PURPOSE:

 * Validates that all required sheets and data structures are properly

 * configured before running Cramer's V correlation analysis

 * 

 * FUNCTIONS:

 * - validatePanSetup(): Full validation with detailed report

 * - isReadyForCramers(): Quick true/false readiness check

 * - generateColumnMapping(): Create detailed column mapping report

 * 

 * AUTHOR: Halloween Party Analytics Team

 * LAST UPDATED: 2025-10-19

 * ============================================================================

 */





// ============================================================================

// VALIDATION REPORT - Primary Diagnostic Function

// ============================================================================



/**

 * Comprehensive validation of Pan analytics setup

 * 

 * Checks:

 * 1. Form Responses (Clean) exists with required columns

 * 2. Pan_Master exists with properly encoded code_* columns

 * 3. Pan_Dict exists with correct structure

 * 4. buildVCramers() function is available

 * 5. V_Cramers output sheet status

 * 

 * Displays detailed report in popup dialog

 * 

 * @returns {boolean} True if all checks pass, false otherwise

 */

function validatePanSetup() {

  const ss = SpreadsheetApp.getActive();

  const report = [];

  let allGood = true;

  

  // ========== REPORT HEADER ==========

  report.push('🔍 CRAMER'S V ANALYSIS - VALIDATION REPORT');

  report.push('='.repeat(60));

  report.push(`Generated: ${new Date().toLocaleString()}`);

  report.push('');

  

  // ========== CHECK 1: SOURCE DATA ==========

  report.push('📋 STEP 1: Check Source Data');

  report.push('-'.repeat(60));

  

  const cleanSheet = ss.getSheetByName('Form Responses (Clean)');

  if (!cleanSheet) {

    report.push('❌ CRITICAL: "Form Responses (Clean)" sheet not found!');

    report.push('   → This is the required source for Pan_Master');

    allGood = false;

  } else {

    const rowCount = cleanSheet.getLastRow() - 1; // Exclude header

    report.push(`✅ Form Responses (Clean) found: ${rowCount} guests`);

    

    // Verify required categorical columns exist

    const cleanData = cleanSheet.getDataRange().getValues();

    if (cleanData.length < 2) {

      report.push('⚠️  Sheet exists but has no data rows');

      allGood = false;

    } else {

      const cleanHeaders = cleanData[0];

      

      // List of key categorical columns needed for analysis

      const requiredCols = [

        'Screen Name',

        'UID',

        'Zodiac Sign',

        'Age Range',

        'Education Level',

        'Self Identified Ethnicity',

        'Self-Identified Gender',

        'Self-Identified Sexual Orientation',

        'Employment Information (Industry)',

        'Employment Information (Role)',

        'Music Preference',

        'At your worst you are…'

      ];

      

      const missingCols = requiredCols.filter(col => 

        !cleanHeaders.some(h => String(h).trim() === col)

      );

      

      if (missingCols.length > 0) {

        report.push(`⚠️  Missing ${missingCols.length} required columns:`);

        missingCols.forEach(col => report.push(`   - ${col}`));

        allGood = false;

      } else {

        report.push(`✅ All ${requiredCols.length} key categorical columns present`);

      }

    }

  }

  report.push('');

  

  // ========== CHECK 2: PAN_MASTER ==========

  report.push('📊 STEP 2: Check Pan_Master (Analysis Dataset)');

  report.push('-'.repeat(60));

  

  const panMaster = ss.getSheetByName('Pan_Master');

  if (!panMaster) {

    report.push('❌ Pan_Master not found');

    report.push('   → Run "Build Pan Sheets" from Analytics menu');

    allGood = false;

  } else {

    const masterData = panMaster.getDataRange().getValues();

    if (masterData.length < 2) {

      report.push('⚠️  Pan_Master exists but is empty');

      report.push('   → Run "Build Pan Sheets" to populate it');

      allGood = false;

    } else {

      const masterHeaders = masterData[0];

      const codeColumns = masterHeaders.filter(h => String(h).startsWith('code_'));

      const ohColumns = masterHeaders.filter(h => String(h).startsWith('oh_'));

      const hasColumns = masterHeaders.filter(h => String(h).startsWith('has_'));

      const rowCount = masterData.length - 1;

      

      if (codeColumns.length === 0) {

        report.push('❌ No code_* columns found in Pan_Master!');

        report.push('   → These are required for Cramer's V analysis');

        allGood = false;

      } else {

        report.push(`✅ Pan_Master found: ${rowCount} guests, ${masterHeaders.length} columns`);

        report.push(`   - ${codeColumns.length} categorical variables (code_*)`);

        report.push(`   - ${ohColumns.length} one-hot columns (oh_*)`);

        report.push(`   - ${hasColumns.length} presence flags (has_*)`);

        report.push('');

        report.push(`   Categorical variables for correlation:`);

        // List first 10 variables to avoid cluttering report

        const varNames = codeColumns.slice(0, 10).map(c => c.replace('code_', ''));

        report.push(`   ${varNames.join(', ')}${codeColumns.length > 10 ? '...' : ''}`);

      }

    }

  }

  report.push('');

  

  // ========== CHECK 3: PAN_DICT ==========

  report.push('📚 STEP 3: Check Pan_Dict (Data Dictionary)');

  report.push('-'.repeat(60));

  

  const panDict = ss.getSheetByName('Pan_Dict');

  if (!panDict) {

    report.push('❌ Pan_Dict not found');

    report.push('   → Run "Build Pan Sheets" from Analytics menu');

    allGood = false;

  } else {

    const dictData = panDict.getDataRange().getValues();

    if (dictData.length < 2) {

      report.push('⚠️  Pan_Dict exists but is empty');

      allGood = false;

    } else {

      const dictHeaders = dictData[0];

      const requiredDictCols = ['Key', 'Header', 'Type', 'Option', 'Code'];

      const hasDictCols = requiredDictCols.every(col => 

        dictHeaders.some(h => String(h).trim() === col)

      );

      

      if (!hasDictCols) {

        report.push('❌ Pan_Dict missing required columns!');

        report.push(`   Required: ${requiredDictCols.join(', ')}`);

        allGood = false;

      } else {

        const singleTypes = dictData.slice(1).filter(row => row[2] === 'single');

        const multiTypes = dictData.slice(1).filter(row => row[2] === 'multi');

        const uniqueVars = new Set(singleTypes.map(row => row[0]));

        

        report.push(`✅ Pan_Dict found: ${dictData.length - 1} total rows`);

        report.push(`   - ${uniqueVars.size} single-choice variables`);

        report.push(`   - ${multiTypes.length > 0 ? 'Yes' : 'No'} multi-select variables`);

        report.push(`   - Structure: Key | Header | Type | Option | Code | Note`);

      }

    }

  }

  report.push('');

  

  // ========== CHECK 4: V_CRAMERS FUNCTION ==========

  report.push('🔧 STEP 4: Check V_Cramers Function');

  report.push('-'.repeat(60));

  

  if (typeof buildVCramers === 'function') {

    report.push('✅ buildVCramers() function is available');

    report.push('   → Ready to generate correlation matrix');

  } else {

    report.push('❌ buildVCramers() function not found!');

    report.push('   → Add V_Cramers.gs script file to your project');

    report.push('   → File should contain buildVCramers() function');

    allGood = false;

  }

  report.push('');

  

  // ========== CHECK 5: EXISTING OUTPUT ==========

  report.push('📈 STEP 5: Check Existing Output');

  report.push('-'.repeat(60));

  

  const vCramers = ss.getSheetByName('V_Cramers');

  if (vCramers) {

    const vData = vCramers.getDataRange().getValues();

    if (vData.length > 1) {

      const varCount = vData.length - 1; // Exclude header

      const lastModified = vCramers.getLastUpdated ? 

        vCramers.getLastUpdated().toLocaleString() : 'Unknown';

      

      report.push(`✅ V_Cramers sheet exists with ${varCount} × ${varCount} matrix`);

      report.push(`   → Contains correlation data`);

      report.push(`   → Can be rebuilt anytime using "Build V_Cramers Matrix"`);

    } else {

      report.push('⚠️  V_Cramers sheet exists but is empty');

      report.push('   → Run "Build V_Cramers Matrix" to generate');

    }

  } else {

    report.push('ℹ️  V_Cramers sheet not yet created');

    report.push('   → Will be created automatically on first run');

  }

  report.push('');

  

  // ========== FINAL SUMMARY ==========

  report.push('='.repeat(60));

  report.push('');

  

  if (allGood) {

    report.push('✅ ALL CHECKS PASSED! ');

    report.push('');

    report.push('Your setup is ready for Cramer's V correlation analysis.');

    report.push('');

    report.push('🚀 NEXT STEPS:');

    report.push('   1. If data changed: Analytics → Build Pan Sheets');

    report.push('   2. To generate correlations: Analytics → Build V_Cramers Matrix');

    report.push('   3. View results in the V_Cramers sheet');

    report.push('');

    report.push('💡 TIP: The correlation matrix shows associations between all');

    report.push('   categorical traits (0.0 = no association, 1.0 = perfect)');

  } else {

    report.push('❌ VALIDATION FAILED');

    report.push('');

    report.push('Please fix the issues above before running analysis.');

    report.push('');

    report.push('🔧 COMMON FIXES:');

    report.push('   • Missing Pan sheets? → Analytics → Build Pan Sheets');

    report.push('   • Missing V_Cramers.gs? → Add the script file (see doc)');

    report.push('   • Missing source data? → Check "Form Responses (Clean)"');

    report.push('   • Need help? → Review Master_Desc sheet for structure');

  }

  

  // ========== DISPLAY REPORT ==========

  const ui = SpreadsheetApp.getUi();

  const title = allGood ? '✅ Validation Passed' : '❌ Validation Failed';

  ui.alert(title, report.join('
'), ui.ButtonSet.OK);

  

  // Also log to console for debugging

  Logger.log('=== VALIDATION REPORT ===');

  Logger.log(report.join('
'));

  

  return allGood;

}





// ============================================================================

// QUICK READINESS CHECK - For Programmatic Use

// ============================================================================



/**

 * Quick validation check without UI popup

 * Useful for automated workflows or pre-flight checks

 * 

 * @returns {boolean} True if ready for Cramer's V analysis

 */

function isReadyForCramers() {

  const ss = SpreadsheetApp.getActive();

  

  // Check all required sheets exist

  const hasClean = !!ss.getSheetByName('Form Responses (Clean)');

  const hasMaster = !!ss.getSheetByName('Pan_Master');

  const hasDict = !!ss.getSheetByName('Pan_Dict');

  const hasFunction = typeof buildVCramers === 'function';

  

  if (!hasMaster || !hasDict || !hasFunction) {

    return false;

  }

  

  // Verify Pan_Master has categorical code columns

  const master = ss.getSheetByName('Pan_Master');

  const masterData = master.getDataRange().getValues();

  if (masterData.length < 2) return false; // No data

  

  const headers = masterData[0];

  const codeCount = headers.filter(h => String(h).startsWith('code_')).length;

  

  return hasClean && codeCount > 0;

}





// ============================================================================

// COLUMN MAPPING REPORT - Detailed Traceability

// ============================================================================



/**

 * Generate detailed mapping of source columns to Pan_Master columns

 * 

 * Creates new sheet "Column_Mapping" showing:

 * - Each column in Form Responses (Clean)

 * - Corresponding column(s) in Pan_Master

 * - Transformation type (ID, Categorical, Multi-select, etc.)

 * 

 * Useful for:

 * - Understanding data transformations

 * - Debugging encoding issues

 * - Documentation for stakeholders

 */

function generateColumnMapping() {

  const ss = SpreadsheetApp.getActive();

  const clean = ss.getSheetByName('Form Responses (Clean)');

  const master = ss.getSheetByName('Pan_Master');

  

  // Validate prerequisite sheets exist

  if (!clean) {

    SpreadsheetApp.getUi().alert(

      'Error',

      '"Form Responses (Clean)" sheet not found.

' +

      'Please ensure this sheet exists before generating mapping.',

      SpreadsheetApp.getUi().ButtonSet.OK

    );

    return;

  }

  

  if (!master) {

    SpreadsheetApp.getUi().alert(

      'Error',

      '"Pan_Master" sheet not found.

' +

      'Run "Build Pan Sheets" first to create Pan_Master.',

      SpreadsheetApp.getUi().ButtonSet.OK

    );

    return;

  }

  

  // Load headers

  const cleanHeaders = clean.getDataRange().getValues()[0];

  const masterHeaders = master.getDataRange().getValues()[0];

  

  // Build mapping report

  const report = [['Source Column (Clean)', 'Pan_Master Column(s)', 'Transformation Type', 'Notes']];

  

  cleanHeaders.forEach((cleanCol, i) => {

    const cleanName = String(cleanCol).trim();

    if (!cleanName) return; // Skip empty columns

    

    let masterCol = '';

    let colType = '';

    let notes = '';

    

    // Determine mapping based on column name and type

    if (cleanName === 'Screen Name' || cleanName === 'UID') {

      masterCol = cleanName;

      colType = 'Identity';

      notes = 'Copied as-is';

      

    } else if (cleanName === 'Timestamp') {

      masterCol = 'TimestampMs';

      colType = 'Temporal';

      notes = 'Converted to epoch milliseconds';

      

    } else if (cleanName === 'Birthday (MM/DD)') {

      masterCol = 'Birthday_MM/DD';

      colType = 'Temporal';

      notes = 'Formatted as MM/DD string';

      

    } else if (cleanName === 'Current 5 Digit Zip Code') {

      masterCol = 'Zip';

      colType = 'Text (raw)';

      notes = 'Preserved as text';

      

    } else if (cleanName === 'Checked-In at Event' || cleanName === 'Check-In Timestamp') {

      masterCol = '(not included)';

      colType = 'Filter';

      notes = 'Used for filtering only';

      

    } else if (cleanName === 'Photo URL') {

      masterCol = '(not included)';

      colType = 'Excluded';

      notes = 'Not needed for analysis';

      

    } else {

      // Try to find corresponding code_ column

      const normalized = cleanName.toLowerCase().replace(/[^a-z0-9]/g, '');

      const codeCol = masterHeaders.find(h => {

        const hNorm = String(h).replace('code_', '').toLowerCase().replace(/[^a-z0-9]/g, '');

        return hNorm === normalized;

      });

      

      if (codeCol) {

        masterCol = codeCol;

        colType = 'Categorical (coded)';

        notes = 'Options mapped to numeric codes (1, 2, 3...)';

      } else {

        // Check for multi-select (oh_* columns)

        const ohMatches = masterHeaders.filter(h => 

          String(h).startsWith('oh_') && 

          normalized.includes(String(h).split('_')[1].toLowerCase())

        );

        

        if (ohMatches.length > 0) {

          masterCol = `${ohMatches.length} columns (oh_*)`;

          colType = 'Multi-select (one-hot)';

          notes = 'One binary column per option (0/1)';

        } else {

          // Check for text presence flag

          const hasCol = masterHeaders.find(h => 

            String(h).startsWith('has_') && 

            normalized.includes(String(h).replace('has_', '').toLowerCase())

          );

          

          if (hasCol) {

            masterCol = hasCol;

            colType = 'Text (presence flag)';

            notes = '1 if text present, 0 if empty';

          } else {

            masterCol = '(not found)';

            colType = 'Unknown';

            notes = 'May not be in SPEC';

          }

        }

      }

    }

    

    report.push([cleanName, masterCol, colType, notes]);

  });

  

  // Write to new sheet

  let mapSheet = ss.getSheetByName('Column_Mapping');

  if (!mapSheet) {

    mapSheet = ss.insertSheet('Column_Mapping');

  }

  mapSheet.clear();

  

  // Write data

  mapSheet.getRange(1, 1, report.length, 4).setValues(report);

  

  // Format header row

  const headerRange = mapSheet.getRange(1, 1, 1, 4);

  headerRange

    .setBackground('#434343')

    .setFontColor('#ffffff')

    .setFontWeight('bold')

    .setHorizontalAlignment('center');

  

  // Auto-resize columns

  mapSheet.autoResizeColumns(1, 4);

  

  // Set reasonable column widths

  mapSheet.setColumnWidth(1, 280); // Source Column

  mapSheet.setColumnWidth(2, 280); // Pan_Master Column

  mapSheet.setColumnWidth(3, 180); // Type

  mapSheet.setColumnWidth(4, 320); // Notes

  

  // Freeze header

  mapSheet.setFrozenRows(1);

  

  // Add alternating row colors

  if (report.length > 1) {

    for (let row = 2; row <= report.length; row++) {

      if (row % 2 === 0) {

        mapSheet.getRange(row, 1, 1, 4).setBackground('#f3f3f3');

      }

    }

  }

  

  // Success message

  SpreadsheetApp.getUi().alert(

    '✅ Column Mapping Generated!',

    `Created "Column_Mapping" sheet with ${report.length - 1} source columns.

` +

    'This shows how each column in "Form Responses (Clean)" is
' +

    'transformed into Pan_Master for analysis.',

    SpreadsheetApp.getUi().ButtonSet.OK

  );

}





// ============================================================================

// REFRESH ALL ANALYTICS - One-Click Update

// ============================================================================



/**

 * Rebuild all analytics sheets in sequence

 * Useful after updating source data

 * 

 * Process:

 * 1. Rebuild Pan_Master and Pan_Dict from Form Responses (Clean)

 * 2. Rebuild V_Cramers correlation matrix

 * 3. Regenerate Master_Desc documentation

 */

function refreshAllAnalytics() {

  const ui = SpreadsheetApp.getUi();

  const ss = SpreadsheetApp.getActive();

  

  // Confirmation dialog

  const response = ui.alert(

    'Refresh All Analytics',

    'This will rebuild:
' +

    '• Pan_Master & Pan_Dict
' +

    '• V_Cramers correlation matrix
' +

    '• Master_Desc documentation

' +

    'This may take a few moments. Continue?',

    ui.ButtonSet.YES_NO

  );

  

  if (response !== ui.Button.YES) return;

  

  try {

    // Step 1: Build Pan sheets

    if (typeof buildPanSheets === 'function') {

      ss.toast('Building Pan_Master and Pan_Dict...', '⏳ Progress', 5);

      buildPanSheets();

      ss.toast('✓ Pan sheets rebuilt', 'Progress', 2);

      Utilities.sleep(1000); // Brief pause

    }

    

    // Step 2: Build V_Cramers

    if (typeof buildVCramers === 'function') {

      ss.toast('Computing Cramer's V correlations...', '⏳ Progress', 5);

      buildVCramers();

      ss.toast('✓ V_Cramers matrix rebuilt', 'Progress', 2);

      Utilities.sleep(1000);

    }

    

    // Step 3: Regenerate Master_Desc

    if (typeof generateMasterDesc === 'function') {

      ss.toast('Updating documentation...', '⏳ Progress', 5);

      generateMasterDesc();

      ss.toast('✓ Master_Desc updated', 'Progress', 2);

    }

    

    // Success!

    ui.alert(

      '✅ All Analytics Refreshed!',

      'Successfully updated:
' +

      '✓ Pan_Master & Pan_Dict
' +

      '✓ V_Cramers correlation matrix
' +

      '✓ Master_Desc documentation

' +

      'All sheets are now up to date with the latest data.',

      ui.ButtonSet.OK

    );

    

  } catch (e) {

    // Error handling

    Logger.log('Error in refreshAllAnalytics: ' + e.toString());

    ui.alert(

      '❌ Error',

      'An error occurred while refreshing analytics:

' +

      e.toString() + '

' +

      'Check the script logs (View → Logs) for details.',

      ui.ButtonSet.OK

    );

  }

}





// ============================================================================

// END OF PAN_VALIDATION.GS

// ============================================================================



// ===== File: Automation =====


/**

 * Automation.gs - Automated Analysis System

 * 

 * PURPOSE:

 * Automatically runs DDD analysis every 30 minutes during the event

 * to monitor checked-in guests in real-time.

 * 

 * FEATURES:

 * - Runs detectDataDisruptors() every 30 minutes

 * - Only analyzes guests who have ACTUALLY checked in at the event

 * - Can be started/stopped manually

 * - Logs all activity for debugging

 * 

 * SETUP:

 * 1. Run setupAutomation() once to create the triggers

 * 2. Automation will run every 30 minutes automatically

 * 3. Run stopAutomation() to turn it off after the event

 */



// ============================================

// SETUP AND CONTROL FUNCTIONS

// ============================================



function setupAutomation() {

  stopAutomation();

  

  ScriptApp.newTrigger('runAutomatedAnalysis')

    .timeBased()

    .everyMinutes(30)

    .create();

  

  Logger.log('✅ Automation setup complete!');

  Logger.log('Analysis will run every 30 minutes.');

  Logger.log('Run stopAutomation() to disable.');

  

  SpreadsheetApp.getUi().alert(

    '✅ Automation Started!

' +

    'DDD analysis will run automatically every 30 minutes.

' +

    'Only guests who have checked in at the event will be analyzed.

' +

    'To stop: Run stopAutomation() function.'

  );

}



function stopAutomation() {

  const triggers = ScriptApp.getProjectTriggers();

  let removedCount = 0;

  

  triggers.forEach(trigger => {

    if (trigger.getHandlerFunction() === 'runAutomatedAnalysis') {

      ScriptApp.deleteTrigger(trigger);

      removedCount++;

    }

  });

  

  Logger.log('✅ Removed ' + removedCount + ' automation trigger(s)');

  

  if (removedCount > 0) {

    SpreadsheetApp.getUi().alert(

      '✅ Automation Stopped!

' +

      'Removed ' + removedCount + ' trigger(s).
' +

      'Analysis will no longer run automatically.'

    );

  }

}



function checkAutomationStatus() {

  const triggers = ScriptApp.getProjectTriggers();

  const activeTriggers = triggers.filter(t => t.getHandlerFunction() === 'runAutomatedAnalysis');

  

  if (activeTriggers.length > 0) {

    Logger.log('✅ Automation is ACTIVE');

    Logger.log('Active triggers: ' + activeTriggers.length);

    

    SpreadsheetApp.getUi().alert(

      '✅ Automation Status: ACTIVE

' +

      'Analysis runs every 30 minutes.
' +

      activeTriggers.length + ' trigger(s) active.'

    );

  } else {

    Logger.log('❌ Automation is INACTIVE');

    SpreadsheetApp.getUi().alert(

      '❌ Automation Status: INACTIVE

' +

      'No triggers found.
' +

      'Run setupAutomation() to start.'

    );

  }

}



// ============================================

// AUTOMATED ANALYSIS FUNCTION

// ============================================



function runAutomatedAnalysis() {

  const startTime = new Date();

  Logger.log('=== AUTOMATED ANALYSIS STARTED at ' + startTime.toString() + ' ===');

  

  try {

    const checkedInCount = countCheckedInGuests();

    Logger.log('Guests checked in at event: ' + checkedInCount);

    

    if (checkedInCount === 0) {

      Logger.log('No guests checked in yet. Skipping analysis.');

      return;

    }

    

    Logger.log('Running DDD analysis on checked-in guests...');

    detectDataDisruptorsForCheckedIn();

    

    const endTime = new Date();

    const duration = (endTime - startTime) / 1000;

    Logger.log('=== ANALYSIS COMPLETE in ' + duration + ' seconds ===');

    Logger.log('Next run in 30 minutes');

    

  } catch (error) {

    Logger.log('ERROR in automated analysis: ' + error.toString());

    Logger.log('Stack: ' + error.stack);

  }

}



function manualRunAnalysis() {

  Logger.log('=== MANUAL ANALYSIS TRIGGERED ===');

  runAutomatedAnalysis();

  SpreadsheetApp.getUi().alert(

    '✅ Manual Analysis Complete!

' +

    'Check the "DDD (Checked-In Only)" sheet for updated results.
' +

    'View execution logs for details.'

  );

}



// ============================================

// HELPER FUNCTIONS

// ============================================



function countCheckedInGuests() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cleanSheet = ss.getSheetByName('Form Responses (Clean)');

  

  if (!cleanSheet) {

    Logger.log('ERROR: Form Responses (Clean) not found');

    return 0;

  }

  

  const data = cleanSheet.getDataRange().getValues();

  if (data.length < 2) return 0;

  

  const headers = data[0];

  const checkedInCol = headers.indexOf('Checked-In at Event');

  

  if (checkedInCol === -1) {

    Logger.log('ERROR: Checked-In at Event column not found');

    return 0;

  }

  

  let count = 0;

  for (let r = 1; r < data.length; r++) {

    const status = String(data[r][checkedInCol] || '').trim().toUpperCase();

    if (status === 'Y') count++;

  }

  

  return count;

}



function detectDataDisruptorsForCheckedIn() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cleanSheet = ss.getSheetByName('Form Responses (Clean)');

  

  if (!cleanSheet) {

    Logger.log('ERROR: Form Responses (Clean) not found');

    return;

  }

  

  const cleanData = cleanSheet.getDataRange().getValues();

  const cleanHeaders = cleanData[0];

  const allRows = cleanData.slice(1);

  

  const screenNameCol = cleanHeaders.indexOf('Screen Name');

  const uidCol = cleanHeaders.indexOf('UID');

  const checkedInCol = cleanHeaders.indexOf('Checked-In at Event');

  

  if (screenNameCol === -1 || uidCol === -1 || checkedInCol === -1) {

    Logger.log('ERROR: Required columns not found');

    return;

  }

  

  const checkedInRows = allRows.filter(row => {

    const status = String(row[checkedInCol] || '').trim().toUpperCase();

    return status === 'Y';

  });

  

  Logger.log('Analyzing ' + checkedInRows.length + ' checked-in guests (out of ' + allRows.length + ' total)');

  

  if (checkedInRows.length === 0) {

    Logger.log('No checked-in guests to analyze');

    return;

  }

  

  const dddHeaders = [

    'Screen Name',

    'UID',

    'DDD - Birthday Format',

    'DDD - More Than 3 Interests',

    'DDD - Unknown Guest',

    'DDD - No Song Request',

    'DDD - Vague Song Request',

    'DDD - Multiple Songs Listed',

    'DDD - No Artist Listed',

    'DDD - Barely Knows Host',

    'DDD - Fresh Acquaintance',

    'DDD - Host Ambiguity',

    'DDD - Age/Birthday Mismatch',

    'DDD - Education/Career Implausibility',

    'DDD - Meme/Joke Response',

    'DDD - Music/Artist Genre Mismatch',

    'DDD - Minimum Effort Detected',

    'DDD - Stranger Danger',

    'DDD - Relationship Contradiction',

    'DDD - Special Snowflake Syndrome',

    'DDD - Suspiciously Generic',

    'DDD - Composite Liar Score'

  ];

  

  const COL = {

    TIMESTAMP: 0,

    BIRTHDAY: 1,

    ZODIAC: 2,

    AGE_RANGE: 3,

    EDUCATION: 4,

    ZIP: 5,

    ETHNICITY: 6,

    GENDER: 7,

    ORIENTATION: 8,

    INDUSTRY: 9,

    ROLE: 10,

    KNOW_HOST: 11,

    WHICH_HOST: 12,

    HOW_WELL: 13,

    INTERESTS: 14,

    MUSIC: 15,

    ARTIST: 16,

    SONG: 17,

    PURCHASE: 18,

    WORST_TRAIT: 19,

    SOCIAL_STANCE: 20

  };

  

  const detectionLog = [];

  let totalDisruptors = 0;

  

  const dddResults = checkedInRows.map((row, idx) => {

    const rowNum = idx + 2;

    const screenName = row[screenNameCol] || 'Unknown';

    const uid = row[uidCol] || 'No UID';

    

    const flags = analyzeSuspiciousPatterns(row, rowNum, COL, detectionLog, screenName);

    

    return [screenName, uid, ...flags];

  });

  

  let dddSheet = ss.getSheetByName('DDD (Checked-In Only)');

  if (dddSheet) {

    dddSheet.clear();

  } else {

    dddSheet = ss.insertSheet('DDD (Checked-In Only)');

  }

  

  dddSheet.getRange(1, 1, 1, dddHeaders.length).setValues([dddHeaders]);

  if (dddResults.length > 0) {

    dddSheet.getRange(2, 1, dddResults.length, dddHeaders.length).setValues(dddResults);

  }

  

  formatDDDSheet(dddSheet, dddHeaders.length, dddResults.length);

  

  dddResults.forEach(row => {

    if (row[row.length - 1] > 0) totalDisruptors++;

  });

  

  Logger.log('Analysis complete: ' + totalDisruptors + ' disruptors found among checked-in guests');

  

  dddSheet.getRange(1, dddHeaders.length + 2).setValue('Last Updated:');

  dddSheet.getRange(1, dddHeaders.length + 3).setValue(new Date());

}



// ============================================

// ANALYSIS HELPER FUNCTIONS (from DDD.gs)

// ============================================



function analyzeSuspiciousPatterns(row, rowNum, COL, log, screenName) {

  const flags = [];

  let liarScore = 0;

  

  function logViolation(violationType, description) {

    log.push({

      screenName: screenName,

      row: rowNum,

      violation: violationType,

      description: description

    });

  }

  

  // DDD 1: Birthday Format

  const birthday = String(row[COL.BIRTHDAY] || '').trim();

  const birthdayFlag = !birthday || !birthday.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/);

  flags.push(birthdayFlag ? 1 : 0);

  if (birthdayFlag) {

    liarScore += 1;

    logViolation('Birthday Format', 'Invalid birthday format: "' + birthday + '"');

  }

  

  // DDD 2: More Than 3 Interests

  const interests = String(row[COL.INTERESTS] || '').trim();

  const interestCount = interests ? interests.split(',').length : 0;

  const tooManyInterests = interestCount > 3;

  flags.push(tooManyInterests ? 1 : 0);

  if (tooManyInterests) {

    liarScore += 2;

    logViolation('More Than 3 Interests', 'Listed ' + interestCount + ' interests (max 3 allowed)');

  }

  

  // DDD 3: Unknown Guest

  const knowHost = String(row[COL.KNOW_HOST] || '').toLowerCase().trim();

  const unknownGuest = knowHost === 'no' || knowHost === '';

  flags.push(unknownGuest ? 1 : 0);

  if (unknownGuest) {

    liarScore += 5;

    logViolation('Unknown Guest', 'Claims not to know the host');

  }

  

  // DDD 4: No Song Request

  const song = String(row[COL.SONG] || '').trim();

  const noSong = !song || song.toLowerCase() === 'none' || song.toLowerCase() === 'n/a';

  flags.push(noSong ? 1 : 0);

  if (noSong) {

    liarScore += 1;

    logViolation('No Song Request', 'Did not provide a song request');

  }

  

  // DDD 5: Vague Song Request

  const vagueSong = song && song.length > 0 && song.length < 10;

  flags.push(vagueSong ? 1 : 0);

  if (vagueSong) {

    liarScore += 1;

    logViolation('Vague Song Request', 'Very short song request: "' + song + '"');

  }

  

  // DDD 6: Multiple Songs Listed

  const multipleSongs = song.includes(',') || song.includes('/') || song.includes('&') || song.includes(' or ');

  flags.push(multipleSongs ? 1 : 0);

  if (multipleSongs) {

    liarScore += 1;

    logViolation('Multiple Songs Listed', 'Listed multiple songs: "' + song + '"');

  }

  

  // DDD 7: No Artist Listed

  const artist = String(row[COL.ARTIST] || '').trim();

  const noArtist = !artist || artist.toLowerCase() === 'none' || artist.toLowerCase() === 'n/a';

  flags.push(noArtist ? 1 : 0);

  if (noArtist && !noSong) {

    liarScore += 1;

    logViolation('No Artist Listed', 'Provided song but no artist');

  }

  

  // DDD 8: Barely Knows Host

  const howWell = String(row[COL.HOW_WELL] || '').toLowerCase().trim();

  const barelyKnows = howWell.includes('barely') || howWell.includes('not well') || howWell === 'acquaintance';

  flags.push(barelyKnows ? 1 : 0);

  if (barelyKnows) {

    liarScore += 2;

    logViolation('Barely Knows Host', 'Barely knows host: "' + howWell + '"');

  }

  

  // DDD 9: Fresh Acquaintance

  const freshAcquaintance = howWell.includes('just met') || howWell.includes('recently') || howWell.includes('new');

  flags.push(freshAcquaintance ? 1 : 0);

  if (freshAcquaintance) {

    liarScore += 3;

    logViolation('Fresh Acquaintance', 'Just met the host: "' + howWell + '"');

  }

  

  // DDD 10: Host Ambiguity

  const whichHost = String(row[COL.WHICH_HOST] || '').trim();

  const hostAmbiguous = !whichHost || whichHost.toLowerCase().includes('both') || whichHost.toLowerCase().includes('not sure');

  flags.push(hostAmbiguous ? 1 : 0);

  if (hostAmbiguous && knowHost === 'yes') {

    liarScore += 2;

    logViolation('Host Ambiguity', 'Unclear which host they know: "' + whichHost + '"');

  }

  

  // DDD 11: Age/Birthday Mismatch

  const ageRange = String(row[COL.AGE_RANGE] || '').trim();

  let ageMismatch = false;

  if (birthday.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {

    const birthYear = parseInt(birthday.split('/')[2]);

    const currentYear = new Date().getFullYear();

    const calculatedAge = currentYear - birthYear;

    

    if (ageRange.includes('18-24') && (calculatedAge < 18 || calculatedAge > 24)) ageMismatch = true;

    if (ageRange.includes('25-34') && (calculatedAge < 25 || calculatedAge > 34)) ageMismatch = true;

    if (ageRange.includes('35-44') && (calculatedAge < 35 || calculatedAge > 44)) ageMismatch = true;

    if (ageRange.includes('45-54') && (calculatedAge < 45 || calculatedAge > 54)) ageMismatch = true;

  }

  flags.push(ageMismatch ? 1 : 0);

  if (ageMismatch) {

    liarScore += 3;

    logViolation('Age/Birthday Mismatch', 'Age range "' + ageRange + '" doesn't match birthday "' + birthday + '"');

  }

  

  // DDD 12: Education/Career Implausibility

  const education = String(row[COL.EDUCATION] || '').toLowerCase();

  const role = String(row[COL.ROLE] || '').toLowerCase();

  const implausible = (education.includes('high school') && (role.includes('director') || role.includes('vp') || role.includes('ceo'))) ||

                       (education.includes('phd') && ageRange === '18-24');

  flags.push(implausible ? 1 : 0);

  if (implausible) {

    liarScore += 4;

    logViolation('Education/Career Implausibility', 'Education "' + education + '" seems mismatched with role "' + role + '"');

  }

  

  // DDD 13: Meme/Joke Response

  const worstTrait = String(row[COL.WORST_TRAIT] || '').toLowerCase();

  const memeKeywords = ['too awesome', 'too perfect', 'none', 'n/a', 'lol', 'lmao', 'haha', 'idk'];

  const isMeme = memeKeywords.some(keyword => worstTrait.includes(keyword));

  flags.push(isMeme ? 1 : 0);

  if (isMeme) {

    liarScore += 2;

    logViolation('Meme/Joke Response', 'Joke response detected: "' + worstTrait + '"');

  }

  

  // DDD 14: Music/Artist Genre Mismatch

  const musicGenre = String(row[COL.MUSIC] || '').toLowerCase();

  const genreMismatch = (musicGenre.includes('classical') && artist.toLowerCase().includes('drake')) ||

                        (musicGenre.includes('country') && artist.toLowerCase().includes('metallica')) ||

                        (musicGenre.includes('rap') && artist.toLowerCase().includes('mozart'));

  flags.push(genreMismatch ? 1 : 0);

  if (genreMismatch) {

    liarScore += 2;

    logViolation('Music/Artist Genre Mismatch', 'Genre "' + musicGenre + '" doesn't match artist "' + artist + '"');

  }

  

  // DDD 15: Minimum Effort Detected

  const totalChars = String(row.join('')).length;

  const minEffort = totalChars < 100 || interests.length < 5 || worstTrait.length < 5;

  flags.push(minEffort ? 1 : 0);

  if (minEffort) {

    liarScore += 2;

    logViolation('Minimum Effort Detected', 'Very short responses (' + totalChars + ' chars total)');

  }

  

  // DDD 16: Stranger Danger

  const strangerDanger = unknownGuest && (barelyKnows || freshAcquaintance);

  flags.push(strangerDanger ? 1 : 0);

  if (strangerDanger) {

    liarScore += 5;

    logViolation('Stranger Danger', 'Multiple indicators of not knowing host');

  }

  

  // DDD 17: Relationship Contradiction

  const contradiction = (knowHost === 'yes' && (howWell.includes('never') || howWell.includes("don't"))) ||

                        (knowHost === 'no' && howWell.includes('well') || howWell.includes('close'));

  flags.push(contradiction ? 1 : 0);

  if (contradiction) {

    liarScore += 4;

    logViolation('Relationship Contradiction', '"' + knowHost + '" contradicts "' + howWell + '"');

  }

  

  // DDD 18: Special Snowflake Syndrome

  const specialSnowflake = interests.toLowerCase().includes('unicorn') || 

                           interests.toLowerCase().includes('quantum') ||

                           worstTrait.toLowerCase().includes('care too much') ||

                           worstTrait.toLowerCase().includes('too nice');

  flags.push(specialSnowflake ? 1 : 0);

  if (specialSnowflake) {

    liarScore += 1;

    logViolation('Special Snowflake Syndrome', 'Overly unique or self-aggrandizing responses');

  }

  

  // DDD 19: Suspiciously Generic

  const genericKeywords = ['stuff', 'things', 'whatever', 'anything', 'nothing special'];

  const tooGeneric = genericKeywords.some(keyword => 

    interests.includes(keyword) || worstTrait.includes(keyword)

  );

  flags.push(tooGeneric ? 1 : 0);

  if (tooGeneric) {

    liarScore += 1;

    logViolation('Suspiciously Generic', 'Extremely vague or generic responses');

  }

  

  // DDD 20: Composite Liar Score

  flags.push(liarScore);

  

  return flags;

}



function formatDDDSheet(sheet, numCols, numRows) {

  sheet.setFrozenRows(1);

  sheet.setFrozenColumns(2);

  

  const headerRange = sheet.getRange(1, 1, 1, numCols);

  headerRange.setFontWeight('bold');

  headerRange.setBackground('#434343');

  headerRange.setFontColor('#ffffff');

  headerRange.setHorizontalAlignment('center');

  

  sheet.setColumnWidth(1, 150);

  sheet.setColumnWidth(2, 100);

  for (let i = 3; i <= numCols - 1; i++) {

    sheet.setColumnWidth(i, 80);

  }

  sheet.setColumnWidth(numCols, 120);

  

  if (numRows > 0) {

    for (let col = 3; col <= numCols; col++) {

      const range = sheet.getRange(2, col, numRows, 1);

      

      const rule = SpreadsheetApp.newConditionalFormatRule()

        .whenNumberGreaterThan(0)

        .setBackground('#f4cccc')

        .setRanges([range])

        .build();

      

      const rules = sheet.getConditionalFormatRules();

      rules.push(rule);

      sheet.setConditionalFormatRules(rules);

    }

    

    const scoreRange = sheet.getRange(2, numCols, numRows, 1);

    scoreRange.setFontWeight('bold');

    

    const scoreRules = [

      SpreadsheetApp.newConditionalFormatRule()

        .whenNumberGreaterThanOrEqualTo(10)

        .setBackground('#cc0000')

        .setFontColor('#ffffff')

        .setRanges([scoreRange])

        .build(),

      SpreadsheetApp.newConditionalFormatRule()

        .whenNumberBetween(5, 9)

        .setBackground('#ff9900')

        .setRanges([scoreRange])

        .build(),

      SpreadsheetApp.newConditionalFormatRule()

        .whenNumberBetween(1, 4)

        .setBackground('#ffff00')

        .setRanges([scoreRange])

        .build()

    ];

    

    const allRules = sheet.getConditionalFormatRules();

    sheet.setConditionalFormatRules(allRules.concat(scoreRules));

  }

  

  sheet.autoResizeColumns(1, numCols);

}



// ===== File: SheetDesc =====


/**

 * Master_Desc.gs - Automatic Workbook Documentation

 * Creates a comprehensive overview of all sheets, columns, and headers

 * Useful for providing context and understanding data structure

 */



const MASTER_DESC_SHEET = 'Master_Desc';



/**

 * Add menu item to Google Sheets UI

 * Run this automatically when spreadsheet opens

 */

function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu('📋 Documentation')

    .addItem('Generate Master_Desc', 'generateMasterDesc')

    .addItem('Refresh All Analytics', 'refreshAllAnalytics')

    .addToUi();

}



/**

 * Main function: Generate Master_Desc sheet

 * Documents all sheets, their columns, and sample data

 */

function generateMasterDesc() {

  const ss = SpreadsheetApp.getActive();

  const allSheets = ss.getSheets();

  

  // Prepare output data

  const output = [

    ['Sheet Name', 'Column #', 'Column Name', 'Data Type', 'Sample Values', 'Row Count', 'Notes']

  ];

  

  allSheets.forEach(sheet => {

    const sheetName = sheet.getName();

    

    // Skip Master_Desc itself to avoid recursion

    if (sheetName === MASTER_DESC_SHEET) return;

    

    const data = sheet.getDataRange().getValues();

    const rowCount = data.length;

    

    if (rowCount === 0) {

      output.push([sheetName, '', '(empty sheet)', '', '', 0, 'No data']);

      return;

    }

    

    const headers = data[0];

    const numCols = headers.length;

    

    // Document each column

    headers.forEach((header, colIdx) => {

      const colNum = colIdx + 1;

      const colName = String(header || `(Column ${colNum})`).trim();

      

      // Analyze column data type and get samples

      const analysis = analyzeColumn_(data, colIdx);

      

      // Notes about special columns

      let notes = '';

      if (colName.startsWith('code_')) notes = 'Categorical code';

      else if (colName.startsWith('oh_')) notes = 'One-hot encoded';

      else if (colName.startsWith('has_')) notes = 'Presence flag';

      else if (colName.toLowerCase().includes('timestamp')) notes = 'Timestamp';

      else if (colName === 'UID') notes = 'Unique identifier';

      

      output.push([

        sheetName,

        colNum,

        colName,

        analysis.type,

        analysis.samples,

        rowCount - 1, // Exclude header row

        notes

      ]);

    });

    

    // Add a blank row between sheets for readability

    output.push(['', '', '', '', '', '', '']);

  });

  

  // Write to Master_Desc sheet

  writeMasterDesc_(output);

  

  // Format the sheet

  formatMasterDesc_();

  

  SpreadsheetApp.getUi().alert(

    `✅ Master_Desc generated!

` +

    `Documented ${allSheets.length - 1} sheets with ${output.length - 1} total columns.`

  );

}



/**

 * Analyze a column to determine data type and get sample values

 */

function analyzeColumn_(data, colIdx) {

  if (data.length < 2) {

    return { type: 'Empty', samples: '' };

  }

  

  // Get up to 3 unique non-empty sample values (excluding header)

  const samples = [];

  const seen = new Set();

  

  for (let i = 1; i < data.length && samples.length < 3; i++) {

    const val = data[i][colIdx];

    if (val === null || val === undefined || val === '') continue;

    

    const valStr = String(val).trim();

    if (valStr && !seen.has(valStr)) {

      seen.add(valStr);

      samples.push(valStr);

    }

  }

  

  // Determine data type

  let type = 'Unknown';

  if (samples.length > 0) {

    const firstVal = data[1][colIdx];

    

    if (firstVal instanceof Date) {

      type = 'Date';

    } else if (typeof firstVal === 'number') {

      // Check if it's all 0s and 1s (binary flag)

      const allBinary = data.slice(1).every(row => {

        const v = row[colIdx];

        return v === 0 || v === 1 || v === '' || v === null;

      });

      type = allBinary ? 'Binary (0/1)' : 'Number';

    } else if (typeof firstVal === 'string') {

      // Check if it looks like codes (all numeric strings 1-10)

      const allCodes = data.slice(1, Math.min(20, data.length)).every(row => {

        const v = row[colIdx];

        if (v === '' || v === null) return true;

        const num = Number(v);

        return isFinite(num) && num >= 1 && num <= 50;

      });

      type = allCodes ? 'Code (1-N)' : 'Text';

    }

  }

  

  return {

    type: type,

    samples: samples.join('; ')

  };

}



/**

 * Write output to Master_Desc sheet

 */

function writeMasterDesc_(data) {

  const ss = SpreadsheetApp.getActive();

  let sheet = ss.getSheetByName(MASTER_DESC_SHEET);

  

  if (!sheet) {

    sheet = ss.insertSheet(MASTER_DESC_SHEET, 0); // Insert as first sheet

  } else {

    sheet.clear();

  }

  

  if (data.length === 0) return;

  

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

}



/**

 * Format Master_Desc sheet for better readability

 */

function formatMasterDesc_() {

  const ss = SpreadsheetApp.getActive();

  const sheet = ss.getSheetByName(MASTER_DESC_SHEET);

  if (!sheet) return;

  

  const lastRow = sheet.getLastRow();

  const lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) return;

  

  // Format header row

  const headerRange = sheet.getRange(1, 1, 1, lastCol);

  headerRange

    .setBackground('#434343')

    .setFontColor('#ffffff')

    .setFontWeight('bold')

    .setFontSize(11);

  

  // Freeze header row

  sheet.setFrozenRows(1);

  

  // Auto-resize all columns

  for (let i = 1; i <= lastCol; i++) {

    sheet.autoResizeColumn(i);

  }

  

  // Set column widths for specific columns

  sheet.setColumnWidth(1, 180); // Sheet Name

  sheet.setColumnWidth(2, 80);  // Column #

  sheet.setColumnWidth(3, 200); // Column Name

  sheet.setColumnWidth(4, 120); // Data Type

  sheet.setColumnWidth(5, 250); // Sample Values

  sheet.setColumnWidth(6, 100); // Row Count

  sheet.setColumnWidth(7, 150); // Notes

  

  // Add alternating row colors for each sheet

  if (lastRow > 1) {

    let currentSheet = '';

    let useGray = false;

    

    for (let row = 2; row <= lastRow; row++) {

      const sheetName = sheet.getRange(row, 1).getValue();

      

      if (sheetName && sheetName !== currentSheet) {

        currentSheet = sheetName;

        useGray = !useGray;

      }

      

      if (sheetName) { // Only color non-empty rows

        const rowRange = sheet.getRange(row, 1, 1, lastCol);

        if (useGray) {

          rowRange.setBackground('#f3f3f3');

        }

      }

    }

  }

  

  // Add borders

  if (lastRow > 1) {

    sheet.getRange(2, 1, lastRow - 1, lastCol)

      .setBorder(null, null, true, null, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID_THIN);

  }

  

  // Center align column numbers and row counts

  if (lastRow > 1) {

    sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment('center'); // Column #

    sheet.getRange(2, 6, lastRow - 1, 1).setHorizontalAlignment('center'); // Row Count

  }

}



/**

 * Optional: Refresh all analytics sheets

 * Calls buildPanSheets and buildVCramers if they exist

 */

function refreshAllAnalytics() {

  try {

    if (typeof buildPanSheets === 'function') {

      buildPanSheets();

      SpreadsheetApp.getActive().toast('✓ Pan sheets rebuilt', 'Progress', 3);

    }

  } catch (e) {

    Logger.log('buildPanSheets not available: ' + e);

  }

  

  try {

    if (typeof buildVCramers === 'function') {

      buildVCramers();

      SpreadsheetApp.getActive().toast('✓ V_Cramers rebuilt', 'Progress', 3);

    }

  } catch (e) {

    Logger.log('buildVCramers not available: ' + e);

  }

  

  // Regenerate Master_Desc to reflect changes

  generateMasterDesc();

  

  SpreadsheetApp.getUi().alert('✅ All analytics refreshed and Master_Desc updated!');

}

