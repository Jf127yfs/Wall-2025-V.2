
// ===== File: Config =====
// ============================================================================
// CENTRALIZED CONFIGURATION - Config.gs
// ============================================================================

// ============================================================================
// MENU SYSTEM
// ============================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('🎉 Party Wall');
  
  if (typeof mergeAnalyticsMenu === 'function') {
    menu = mergeAnalyticsMenu(menu);
  }
  
  menu.addItem('🔄 Check Setup', 'checkSetup')
      .addItem('🌐 Open Display', 'testOpenDisplay')
      .addItem('📸 Fix Photos', 'testPhotoLoading')
      .addToUi();
}

// ============================================================================
// CONFIGURATION
// ============================================================================

// SHEET NAMES
var SHEET_NAME = 'Form Responses (Clean)';
var WALL_CLEAN_SHEET = 'Form Responses (Clean)';
var REPORTS_CLEAN_SHEET = 'Form Responses (Clean)';
var DDD_SHEET_NAME = 'DDD (Checked-In Only)';
var DDD_REPORT_SHEET = 'DDD Report';
var FORM_RESPONSES_ORIGINAL = 'Form Responses 1';
var REPORTS_MASTER_SHEET = 'Pan_Master';

// DRIVE CONFIGURATION
var PHOTOS_FOLDER_NAME = 'Event Guest Photos';
var WALL_TARGET_ADDRESS = '5317 Charlotte St, Kansas City, MO 64110';

// FORM RESPONSES (CLEAN) - COLUMN INDICES
var TIMESTAMP_COL = 1;
var BIRTHDAY_COL = 2;
var ZODIAC_COL = 3;
var AGE_RANGE_COL = 4;
var EDUCATION_COL = 5;
var ZIP_COL = 6;
var ETHNICITY_COL = 7;
var GENDER_COL = 8;
var ORIENTATION_COL = 9;
var INDUSTRY_COL = 10;
var ROLE_COL = 11;
var HOST_KNOW_COL = 12;
var HOST_LONGEST_COL = 13;
var HOST_SCORE_COL = 14;
var INTERESTS_COL = 15;
var MUSIC_PREF_COL = 16;
var ARTIST_COL = 17;
var SONG_REQUEST_COL = 18;
var RECENT_PURCHASE_COL = 19;
var AT_WORST_COL = 20;
var SOCIAL_STANCE_COL = 21;
var SCREEN_NAME_COL = 22;
var UID_COL = 23;
var CHECKED_FLAG_COL = 24;
var CHECKED_TS_COL = 25;
var PHOTO_URL_COL = 26;

// DDD (CHECKED-IN ONLY) - COLUMN INDICES
var DDD_SCREEN_NAME_COL = 1;
var DDD_UID_COL = 2;
var DDD_BIRTHDAY_FORMAT_COL = 3;
var DDD_MORE_THAN_3_INTERESTS_COL = 4;
var DDD_UNKNOWN_GUEST_COL = 5;
var DDD_NO_SONG_REQUEST_COL = 6;
var DDD_VAGUE_SONG_REQUEST_COL = 7;
var DDD_MULTIPLE_SONGS_COL = 8;
var DDD_NO_ARTIST_COL = 9;
var DDD_BARELY_KNOWS_HOST_COL = 10;
var DDD_FRESH_ACQUAINTANCE_COL = 11;
var DDD_HOST_AMBIGUITY_COL = 12;
var DDD_AGE_BIRTHDAY_MISMATCH_COL = 13;
var DDD_EDUCATION_CAREER_COL = 14;
var DDD_MEME_JOKE_COL = 15;
var DDD_MUSIC_GENRE_MISMATCH_COL = 16;
var DDD_MINIMUM_EFFORT_COL = 17;
var DDD_STRANGER_DANGER_COL = 18;
var DDD_RELATIONSHIP_CONTRADICTION_COL = 19;
var DDD_SPECIAL_SNOWFLAKE_COL = 20;
var DDD_SUSPICIOUSLY_GENERIC_COL = 21;
var DDD_COMPOSITE_SCORE_COL = 22;
var DDD_ESTIMATED_FRAUD_COL = 23;

// ANALYSIS CATEGORIES FOR WALL
var ANALYSIS_CATEGORIES = [
  { field: 'Music Preference', name: 'MUSIC_GENRE' },
  { field: 'Self-Identified Gender', name: 'GENDER_IDENTITY' },
  { field: 'Self-Identified Sexual Orientation', name: 'ORIENTATION' },
  { field: 'Age Range', name: 'AGE_COHORT' },
  { field: 'Education Level', name: 'EDUCATION_LEVEL' },
  { field: 'Self Identified Ethnicity', name: 'ETHNIC_IDENTITY' },
  { field: 'Employment Information (Industry)', name: 'INDUSTRY_SECTOR' },
  { field: 'Employment Information (Role)', name: 'OCCUPATIONAL_ROLE' },
  { field: 'At your worst you are…', name: 'PERSONALITY_TRAIT' },
  { field: "Recent purchase you're most happy about", name: "CONSUMER_BEHAVIOR" },
  { field: 'Do you know the Host(s)?', name: 'HOST_RELATIONSHIP_DURATION' },
  { field: 'Which host have you known the longest?', name: 'PRIMARY_HOST_CONNECTION' }
];

// TIMING CONSTANTS
var CYCLE_INTERVAL = 20000;
var SECONDARY_CYCLE_INTERVAL = 15000;
var LINE_DRAW_TIME = 800;
var FADE_OUT_TIME = 500;
var CRITICAL_ANIMATION_SPEED = 3000;

// COLOR PALETTE
var COLOR_PALETTE = [
  'rgba(255, 0, 0, 1)', 'rgba(255, 105, 180, 1)', 'rgba(255, 165, 0, 1)',
  'rgba(255, 255, 0, 1)', 'rgba(0, 255, 0, 1)', 'rgba(0, 255, 127, 1)',
  'rgba(0, 255, 255, 1)', 'rgba(0, 191, 255, 1)', 'rgba(138, 43, 226, 1)',
  'rgba(255, 0, 255, 1)', 'rgba(255, 20, 147, 1)', 'rgba(255, 69, 0, 1)',
  'rgba(173, 255, 47, 1)', 'rgba(0, 206, 209, 1)', 'rgba(147, 112, 219, 1)',
  'rgba(255, 140, 0, 1)', 'rgba(50, 205, 50, 1)', 'rgba(255, 192, 203, 1)',
  'rgba(64, 224, 208, 1)', 'rgba(238, 130, 238, 1)'
];


// ===== File: Utilities =====

/**
 * Check in specific guests by UID for precise testing
 */
function mockCheckInSpecificGuests() {
  Logger.log('╔════════════════════════════════════════════════════════════════╗');
  Logger.log('║           🎯 CHECK IN SPECIFIC GUESTS (BY UID)                 ║');
  Logger.log('╚════════════════════════════════════════════════════════════════╝');
  
  var uidsToCheckIn = [
    'PS-438',  // PhantomShade 438
    'HJ-127',  // HollowJack 127
    'CM-903',  // CryptMuse 903
    'GD-521',  // GraveDancer 521
    'VS-664'   // VelvetSpecter 664
  ];
  
  Logger.log('📋 WHAT THIS DOES:');
  Logger.log('   • Checks in ONLY the guests you specify');
  Logger.log('   • Useful for testing specific scenarios');
  
  Logger.log('🎯 GUESTS TO CHECK IN:');
  for (var i = 0; i < uidsToCheckIn.length; i++) {
    var uid = uidsToCheckIn[i];
    Logger.log('   ' + (i + 1) + '. ' + uid);
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    Logger.log('❌ ERROR: Sheet "' + SHEET_NAME + '" not found');
    return;
  }
  
  Logger.log('✅ Found sheet: ' + SHEET_NAME);
  
  var data = sheet.getDataRange().getValues();
  var checkedInCount = 0;
  var notFoundCount = 0;
  
  Logger.log('🔍 Searching for specified guests...');
  
  var lock = LockService.getScriptLock();
  try {
    lock.tryLock(30000);
    Logger.log('   🔒 Lock acquired');
    
    for (var i = 1; i < data.length; i++) {
      var uid = String(data[i][UID_COL - 1] || '').trim();
      
      if (uidsToCheckIn.includes(uid)) {
        var row = i + 1;
        var screenName = data[i][SCREEN_NAME_COL - 1];
        
        sheet.getRange(row, CHECKED_FLAG_COL).setValue('Y');
        
        var now = new Date();
        var checkInTime = new Date(
          now.getFullYear(),
          now.getMonth(),
          now.getDate(),
          19 + Math.floor(Math.random() * 4),
          Math.floor(Math.random() * 60),
          0
        );
        sheet.getRange(row, CHECKED_TS_COL).setValue(checkInTime);
        
        var timeStr = Utilities.formatDate(checkInTime, Session.getScriptTimeZone(), 'HH:mm');
        Logger.log('   ✅ Found and checked in: ' + screenName + ' (' + uid + ') at ' + timeStr);
        checkedInCount++;
      }
    }
    
    SpreadsheetApp.flush();
    Logger.log('   💾 Changes saved');
    
  } finally {
    lock.releaseLock();
    Logger.log('   🔓 Lock released');
  }
  
  notFoundCount = uidsToCheckIn.length - checkedInCount;
  
  Logger.log('╔════════════════════════════════════════════════════════════════╗');
  Logger.log('║                    ✅ CHECK-IN COMPLETE                        ║');
  Logger.log('╚════════════════════════════════════════════════════════════════╝');
  
  Logger.log('📊 RESULTS:');
  Logger.log('   • Requested: ' + uidsToCheckIn.length + ' guests');
  Logger.log('   • Successfully checked in: ' + checkedInCount);
  if (notFoundCount > 0) {
    Logger.log('   • ⚠️  Not found: ' + notFoundCount);
  }
  Logger.log('');
  Logger.log('════════════════════════════════════════════════════════════════');
}

/**
 * Show current check-in status (diagnostic tool)
 */
function showCheckInStatus() {
  Logger.log('╔════════════════════════════════════════════════════════════════╗');
  Logger.log('║                📊 CHECK-IN STATUS REPORT                       ║');
  Logger.log('╚════════════════════════════════════════════════════════════════╝');
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    Logger.log('❌ ERROR: Sheet "' + SHEET_NAME + '" not found');
    return;
  }
  
  Logger.log('✅ Reading from: ' + SHEET_NAME);
  
  var data = sheet.getDataRange().getValues();
  var checkedInCount = 0;
  var notCheckedInCount = 0;
  
  for (var i = 1; i < data.length; i++) {
    var checkedIn = String(data[i][CHECKED_FLAG_COL - 1] || '').trim().toUpperCase();
    if (checkedIn === 'Y') {
      checkedInCount++;
    } else {
      notCheckedInCount++;
    }
  }
}

function testMMLoad() {
  try {
    // Try capital MM
    var html = HtmlService.createTemplateFromFile('MM').evaluate();
    Logger.log('✅ MM file found!');
  } catch (e) {
    Logger.log('❌ MM not found, trying lowercase...');
    try {
      var html = HtmlService.createTemplateFromFile('mm').evaluate();
      Logger.log('✅ mm file found!');
    } catch (e2) {
      Logger.log('❌ Neither MM nor mm found');
    }
  }
}

function testDisplayPage() {
  var result = serveDisplayPage();
  Logger.log('Display page title: ' + result.getTitle());
  Logger.log('Display page content length: ' + result.getContent().length);
}


// ===== File: MapCode =====

// Code.gs - Google Apps Script (SAME AS BEFORE)

function getZipCodesFromSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Form Responses 1');
    
    if (!sheet) {
      throw new Error('Sheet "Form Responses 1" not found');
    }
    
    var lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return {};
    }
    
    var zipRange = sheet.getRange(2, 5, lastRow - 1, 1);
    var zipValues = zipRange.getValues();
    
    // Count occurrences of each zip code
    var zipCounts = {};
    zipValues.forEach(row => {
      var zip = String(row[0]).trim();
      if (zip && zip !== '' && zip.length === 5) {
        zipCounts[zip] = (zipCounts[zip] || 0) + 1;
      }
    });
    
    return zipCounts;
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return {};
  }
}

function getAddressCoordinates(address) {
  try {
    var geocoder = Maps.newGeocoder();
    var location = geocoder.geocode(address);
    
    if (location.status === 'OK' && location.results.length > 0) {
      var result = location.results[0];
      return {
        lat: result.geometry.location.lat,
        lng: result.geometry.location.lng
      };
    }
    return null;
  } catch (error) {
    Logger.log('Error geocoding address: ' + error.toString());
    return null;
  }
}

function getZipCodeCoordinates(zipCode) {
  try {
    var geocoder = Maps.newGeocoder();
    var location = geocoder.geocode(zipCode + ', USA');
    
    if (location.status === 'OK' && location.results.length > 0) {
      var result = location.results[0];
      return {
        lat: result.geometry.location.lat,
        lng: result.geometry.location.lng,
        zip: zipCode
      };
    }
    return null;
  } catch (error) {
    Logger.log('Error geocoding ' + zipCode + ': ' + error.toString());
    return null;
  }
}


// ===== File: Code =====

// ============================================
// PARTY EVENT MANAGEMENT SYSTEM - Code.gs
// ============================================
// Purpose: Backend for guest check-in web application
// Created: 2025
// Last Modified: [Update when you make changes]
//
// FEATURES:
// - Guest check-in via web form (ZIP, Gender, DOB verification)
// - Screen name updates after check-in
// - Photo uploads to Google Drive
// - Real-time spreadsheet updates
//
// WORKFLOW:
// 1. Form Responses 1: Original form data with "Checked-In" pre-approval column
// 2. Run DataClean.gs: Filters guests with "Checked-In" = "Y" → creates Form Responses (Clean)
// 3. Form Responses (Clean): Has columns 24-26 (Checked-In at Event, Timestamp, Photo URL)
// 4. Web app: Searches Form Responses (Clean) and marks event day check-in
//
// IMPORTANT: Re-run DataClean.gs whenever Form Responses 1 is updated with new approvals
// ============================================

// ============================================
// GLOBAL CONSTANTS (DECLARED ONCE ONLY)
// ============================================



// ============================================
// WEB APP ENTRY POINT
// ============================================

/**
 * Main entry point for web app requests
 * 
 * DEPLOYMENT:
 * 1. Deploy → New deployment → Web app
 * 2. Execute as: Me
 * 3. Who has access: Anyone (or as needed)
 * 
 * URLs:
 * - https://script.google.com/.../exec → Display.html
 * - https://script.google.com/.../exec?page=wall → wall.html
 * - https://script.google.com/.../exec?page=mm → mm.html
 * - https://script.google.com/.../exec?page=map → map.html
 * - https://script.google.com/.../exec?page=checkin → CheckInInterface.html
 * 
 * @param {Object} e - Event object with query parameters
 * @return {HtmlOutput} Rendered HTML page
 */
function doGet(e) {
  try {
    // Handle case where e is undefined (when called from script editor)
    if (!e || !e.parameter) {
      e = { parameter: {} };
    }
    
    // Get page parameter from URL (?page=display)
    var page = (e.parameter.page || 'display').toLowerCase();
    
    Logger.log('doGet called with page: ' + page);
    
    // Route to appropriate page
    var template;
    switch(page) {
      case 'display':
        template = HtmlService.createTemplateFromFile('Display');
        break;
      case 'checkin':
        template = HtmlService.createTemplateFromFile('CheckInInterface');
        break;
      case 'wall':
        template = HtmlService.createTemplateFromFile('wall');
        break;
      case 'mm':
      case 'matchmaker':
        template = HtmlService.createTemplateFromFile('mm');
        break;
      case 'map':
        template = HtmlService.createTemplateFromFile('map');
        break;
      case 'intro':
        template = HtmlService.createTemplateFromFile('intro');
        break;
      default:
        template = HtmlService.createTemplateFromFile('Display');
    }
    
    // Set template variables if needed
    template.deploymentUrl = ScriptApp.getService().getUrl();
    
    // Evaluate template and set properties
    var output = template.evaluate()
      .setTitle('Party Wall System')
      .setFaviconUrl('https://www.google.com/images/favicon.ico')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    // Log successful page serve
    Logger.log('Served page: ' + page);
    return output;
    
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return HtmlService.createHtmlOutput(
      '<h1>Error</h1>' +
      '<p>Something went wrong: ' + error.message + '</p>'
    );
  }
}


// ============================================
// MAIN CHECK-IN FUNCTION
// ============================================

function checkInGuest(payload) {
  // STEP 1: Extract and validate input
  var zip = safeString_(payload && payload.zip);
  var genderInput = safeString_(payload && payload.gender);
  var dobInput = safeString_(payload && payload.dob);

  if (!zip || !genderInput || !dobInput) {
    Logger.log('Check-in failed: Missing required fields');
    return { ok: false, message: 'Missing zip, gender, or DOB.' };
  }

  Logger.log('Check-in attempt: ZIP=' + zip + ', Gender=' + genderInput + ', DOB=' + dobInput);

  // STEP 2: Access the spreadsheet and sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  
  if (!sh) {
    Logger.log('ERROR: Sheet not found: ' + SHEET_NAME);
    return { ok: false, message: 'Sheet "' + SHEET_NAME + '" not found.' };
  }

  // STEP 3: Get all data from sheet
  var data = sh.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('ERROR: Sheet is empty or has only headers');
    return { ok: false, message: 'No data to search.' };
  }

  Logger.log('Sheet has ' + (data.length - 1) + ' data rows');

  // STEP 4: Find gender and DOB columns dynamically
  var headers = [];
  for (var i = 0; i < data[0].length; i++) {
    headers.push(safeString_(data[0][i]));
  }
  var genderCol = findHeaderIndex_(headers, ['self-identified gender', 'gender', 'sex']);
  var dobCol = findHeaderIndex_(headers, ['birthday (mm/dd)', 'birthday', 'dob', 'date of birth']);
  
  if (genderCol === -1 || dobCol === -1) {
    Logger.log('ERROR: Required columns not found. Gender col: ' + genderCol + ', DOB col: ' + dobCol);
    return { ok: false, message: 'Required columns not found in sheet.' };
  }

  Logger.log('Found columns - Gender: ' + genderCol + ', DOB: ' + dobCol);

  // STEP 5: Normalize search criteria
  var wantZip = normalizeZip_(zip);
  var wantGD = normalizeGender_(genderInput);
  var wantMD = parseMonthDay_(dobInput);
  
  if (!wantMD) {
    Logger.log('ERROR: Invalid DOB format: ' + dobInput);
    return { ok: false, message: 'DOB must be MM/DD format (e.g., 03/15).' };
  }

  Logger.log('Searching for: ZIP=' + wantZip + ', Gender=' + wantGD + ', DOB=' + wantMD.m + '/' + wantMD.d);

  // STEP 6: Search through all rows for exact match
  var matches = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    
    var rowZip = normalizeZip_(safeString_(row[ZIP_COL - 1]));
    var rowGender = normalizeGender_(safeString_(row[genderCol]));
    var rowDOBRaw = row[dobCol];

    if (!rowZip || rowZip !== wantZip) continue;
    if (rowGender && wantGD && rowGender !== wantGD) continue;

    var md = monthDayFromCell_(rowDOBRaw);
    if (!md || md.m !== wantMD.m || md.d !== wantMD.d) continue;

    matches.push({ r, row });
    Logger.log('Match found at row ' + (r + 1));
  }

  // STEP 7: Validate match results
  if (matches.length === 0) {
    Logger.log('No matching guest found');
    return { ok: false, message: 'No matching guest found. Verify your information.' };
  }
  
  if (matches.length > 1) {
    Logger.log('ERROR: Multiple matches found (' + matches.length + ' guests)');
    return { ok: false, message: 'Multiple matches found. Contact event staff.' };
  }

  // STEP 8: Extract guest data from the single match
  var match = matches[0];
  var sheetRow = match.r + 1;
  var screenName = safeString_(match.row[SCREEN_NAME_COL - 1]);
  var uid = safeString_(match.row[UID_COL - 1]);
  var alreadyCheckedIn = safeString_(match.row[CHECKED_FLAG_COL - 1]).toUpperCase() === 'Y';

  // STEP 9: Check if already checked in
  if (alreadyCheckedIn) {
    Logger.log('Guest already checked in: ' + screenName);
    return {
      ok: true,
      alreadyCheckedIn: true,
      screenName: screenName || 'N/A',
      uid: uid || 'N/A',
      message: 'Welcome back, ' + screenName + '! You already checked in.'
    };
  }

  Logger.log('Check-in successful: ' + screenName + ' (' + uid + ') at row ' + sheetRow);

  // STEP 10: Mark guest as checked in with timestamp
  var lock = LockService.getScriptLock();
  try {
    lock.tryLock(10000);
    
    sh.getRange(sheetRow, CHECKED_FLAG_COL).setValue('Y');
    sh.getRange(sheetRow, CHECKED_TS_COL).setValue(new Date());
    
    SpreadsheetApp.flush();
    
    Logger.log('Sheet updated: Row ' + sheetRow + ' marked as checked in');
    
  } catch (error) {
    Logger.log('ERROR updating sheet: ' + error.toString());
    return { ok: false, message: 'Error during check-in: ' + error.toString() };
  } finally {
    lock.releaseLock();
  }

  // STEP 11: Return success with guest credentials
  return { 
    ok: true,
    alreadyCheckedIn: false,
    screenName: screenName || 'N/A',
    uid: uid || 'N/A',
    rowNumber: sheetRow,
    message: 'Check-in successful! Welcome, ' + screenName + '!'
  };
}

// ============================================
// SCREEN NAME UPDATE FUNCTION
// ============================================

function updateGuestScreenName(payload) {
  // STEP 1: Extract and validate input
  var uid = safeString_(payload && payload.uid);
  var newScreenName = safeString_(payload && payload.newScreenName);

  if (!uid || !newScreenName) {
    Logger.log('Update failed: Missing UID or screen name');
    return { ok: false, message: 'Missing UID or screen name.' };
  }

  Logger.log('Screen name update request: UID=' + uid + ', New name=' + newScreenName);

  // STEP 2: Validate screen name format
  if (newScreenName.length < 3) {
    Logger.log('Validation failed: Name too short');
    return { ok: false, message: 'Screen name must be at least 3 characters.' };
  }
  if (newScreenName.length > 50) {
    Logger.log('Validation failed: Name too long');
    return { ok: false, message: 'Screen name must be 50 characters or less.' };
  }

  // STEP 3: Access the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    Logger.log('ERROR: Sheet not found: ' + SHEET_NAME);
    return { ok: false, message: 'Sheet not found.' };
  }

  // STEP 4: Find guest by UID
  var data = sh.getDataRange().getValues();
  var targetRow = -1;
  
  for (var r = 1; r < data.length; r++) {
    if (safeString_(data[r][UID_COL - 1]) === uid) {
      targetRow = r + 1;
      Logger.log('Found guest at row ' + targetRow);
      break;
    }
  }

  if (targetRow === -1) {
    Logger.log('ERROR: Guest not found with UID: ' + uid);
    return { ok: false, message: 'Guest not found with UID: ' + uid };
  }

  // STEP 5: Update the screen name
  var lock = LockService.getScriptLock();
  try {
    lock.tryLock(10000);
    
    sh.getRange(targetRow, SCREEN_NAME_COL).setValue(newScreenName);
    SpreadsheetApp.flush();
    
    Logger.log('Screen name updated: ' + uid + ' → ' + newScreenName);
    
  } catch (error) {
    Logger.log('ERROR updating screen name: ' + error.toString());
    return { ok: false, message: 'Error updating: ' + error.toString() };
  } finally {
    lock.releaseLock();
  }

  return { 
    ok: true, 
    newScreenName: newScreenName,
    message: 'Screen name updated successfully!'
  };
}

// ============================================
// PHOTO UPLOAD FUNCTION
// ============================================

function uploadGuestPhoto(payload) {
  try {
    // STEP 1: Extract and validate input
    var uid = safeString_(payload && payload.uid);
    var fileName = safeString_(payload && payload.fileName);
    var mimeType = safeString_(payload && payload.mimeType);
    var base64Data = payload && payload.base64Data;

    if (!uid || !base64Data) {
      Logger.log('Upload failed: Missing UID or photo data');
      return { ok: false, message: 'Missing UID or photo data.' };
    }

    Logger.log('Photo upload started: UID=' + uid + ', File=' + fileName);

    // STEP 2: Decode base64 data to binary blob
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      mimeType || 'image/jpeg',
      fileName || 'photo.jpg'
    );

    Logger.log('Photo decoded: ' + blob.getBytes().length + ' bytes');

    // STEP 3: Get or create photos folder in Google Drive
    var folder = getOrCreatePhotosFolder_();
    Logger.log('Using Drive folder: ' + folder.getName() + ' (ID: ' + folder.getId() + ')');

    // STEP 4: Create unique filename
    var timestamp = new Date().getTime();
    var cleanFileName = (fileName || 'photo.jpg').replace(/[^a-zA-Z0-9._-]/g, '_');
    var uniqueFileName = uid + '_' + timestamp + '_' + cleanFileName;
    
    Logger.log('Saving as: ' + uniqueFileName);

    // STEP 5: Upload file to Drive
    var file = folder.createFile(blob);
    file.setName(uniqueFileName);
    file.setDescription('Photo for guest UID: ' + uid + ' uploaded at ' + new Date().toString());

    // STEP 6: Set file permissions
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileUrl = file.getUrl();
    var fileId = file.getId();
    
    Logger.log('File uploaded: ' + fileUrl);

    // STEP 7: Update spreadsheet with photo URL
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    
    if (!sh) {
      Logger.log('WARNING: Sheet not found, but photo was uploaded');
      return { 
        ok: true, 
        fileUrl: fileUrl,
        fileId: fileId,
        message: 'Photo uploaded but could not update spreadsheet.' 
      };
    }

    // STEP 8: Find guest by UID
    var data = sh.getDataRange().getValues();
    var targetRow = -1;
    
    for (var r = 1; r < data.length; r++) {
      if (safeString_(data[r][UID_COL - 1]) === uid) {
        targetRow = r + 1;
        break;
      }
    }

    if (targetRow === -1) {
      Logger.log('WARNING: Guest not found, but photo was uploaded');
      return { 
        ok: true, 
        fileUrl: fileUrl,
        fileId: fileId,
        message: 'Photo uploaded but guest not found in spreadsheet.' 
      };
    }

    // STEP 9: Save photo URL to spreadsheet
    var lock = LockService.getScriptLock();
    try {
      lock.tryLock(10000);
      
      sh.getRange(targetRow, PHOTO_URL_COL).setValue(fileUrl);
      SpreadsheetApp.flush();
      
      Logger.log('Photo URL saved to row ' + targetRow);
      
    } finally {
      lock.releaseLock();
    }

    // STEP 10: Return success
    return { 
      ok: true, 
      fileUrl: fileUrl,
      fileId: fileId,
      message: 'Photo uploaded successfully!'
    };

  } catch (error) {
    Logger.log('ERROR uploading photo: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    return { 
      ok: false, 
      message: 'Error uploading photo: ' + error.toString() 
    };
  }
}

// ============================================
// SHARED HELPER FUNCTIONS
// ============================================

function safeString_(val) {
  if (val == null) return '';
  return String(val).trim();
}

function findHeaderIndex_(headers, possibilities) {
  for (var p of possibilities) {
    var norm = p.toLowerCase().replace(/[^a-z0-9]/g, '');
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i].toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === norm) return i;
    }
  }
  return -1;
}

function normalizeZip_(zip) {
  var cleaned = zip.replace(/[^0-9]/g, '');
  return cleaned.length >= 5 ? cleaned.substring(0, 5) : cleaned;
}

function normalizeGender_(g) {
  var lower = g.toLowerCase();
  if (lower.includes('man') && !lower.includes('woman')) return 'man';
  if (lower.includes('woman')) return 'woman';
  if (lower.includes('non') || lower.includes('binary')) return 'nonbinary';
  if (lower.includes('other')) return 'other';
  return lower;
}

function parseMonthDay_(str) {
  var match = str.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if (!match) return null;
  return { m: parseInt(match[1], 10), d: parseInt(match[2], 10) };
}

// ============================================
// ADD THIS TO YOUR Code.gs
// Place it in the "SHARED HELPER FUNCTIONS" section
// Replace the existing monthDayFromCell_ function
// ============================================

/**
 * Includes and returns the content of an HTML file.
 * This allows modular HTML by including files in templates.
 * @param {string} filename - Name of the HTML file to include
 * @return {string} - The file contents
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Resolves a column index from the column map using the field name
 * @param {Object} colMap - Map of column headers to indices
 * @param {string} field - Field name to find
 * @return {number|undefined} - Column index or undefined if not found
 */
function resolveColumnIndex_(colMap, field) {
  var key;
  var keys = Object.keys(colMap);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase() === field.toLowerCase()) {
      key = keys[i];
      break;
    }
  }
  return key !== undefined ? colMap[key] : undefined;
}

/**
 * monthDayFromCell_() - Extracts month and day from various cell formats
 * Handles Date objects and string formats like "10/07" or "10/07/1990"
 * UPDATED to better handle Date objects from Google Sheets
 */

function monthDayFromCell_(val) {
  if (!val) return null;
  
  // Handle Date objects (most common in Google Sheets)
  if (val instanceof Date) {
    // Get month and day from Date object
    var month = val.getMonth() + 1; // getMonth() returns 0-11
    var day = val.getDate();
    return { m: month, d: day };
  }
  
  // Handle string formats
  var str = String(val).trim();
  
  // If it looks like a full date string from Date.toString()
  if (str.includes('GMT') || str.includes('UTC')) {
    try {
      var dateObj = new Date(str);
      if (!isNaN(dateObj.getTime())) {
        return { m: dateObj.getMonth() + 1, d: dateObj.getDate() };
      }
    } catch (e) {
      // Fall through to pattern matching
    }
  }
  
  // Pattern matching for various string formats
  var patterns = [
    /^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-]\d{2,4})?$/, // MM/DD or MM/DD/YYYY
    /^(\d{1,2})\/(\d{1,2})$/                         // MM/DD
  ];
  
  for (var pattern of patterns) {
    var match = str.match(pattern);
    if (match) {
      return { m: parseInt(match[1], 10), d: parseInt(match[2], 10) };
    }
  }
  
  return null;
}

function getOrCreatePhotosFolder_() {
  var folders = DriveApp.getFoldersByName(PHOTOS_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  Logger.log('Creating new Drive folder: ' + PHOTOS_FOLDER_NAME);
  return DriveApp.createFolder(PHOTOS_FOLDER_NAME);
}

// ============================================
// TEST AND DEBUG FUNCTIONS
// ============================================

function testSheetStructure() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  
  if (!sh) {
    Logger.log('ERROR: Sheet not found: ' + SHEET_NAME);
    return;
  }
  
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Logger.log('=== SHEET STRUCTURE ===');
  Logger.log('Sheet: ' + SHEET_NAME);
  Logger.log('Total columns: ' + headers.length);
  Logger.log('');
  Logger.log('Column mapping:');
  
  for (var i = 0; i < headers.length; i++) {
    var colNum = i + 1;
    var colLetter = String.fromCharCode(65 + (i % 26));
    Logger.log('  Column ' + colNum + ' (' + colLetter + '): ' + headers[i]);
  }
  
  Logger.log('');
  Logger.log('Current constants:');
  Logger.log('  ZIP_COL = ' + ZIP_COL + ' (' + headers[ZIP_COL - 1] + ')');
  Logger.log('  SCREEN_NAME_COL = ' + SCREEN_NAME_COL + ' (' + headers[SCREEN_NAME_COL - 1] + ')');
  Logger.log('  UID_COL = ' + UID_COL + ' (' + headers[UID_COL - 1] + ')');
  Logger.log('  CHECKED_FLAG_COL = ' + CHECKED_FLAG_COL + ' (' + headers[CHECKED_FLAG_COL - 1] + ')');
  Logger.log('  CHECKED_TS_COL = ' + CHECKED_TS_COL + ' (' + headers[CHECKED_TS_COL - 1] + ')');
  Logger.log('  PHOTO_URL_COL = ' + PHOTO_URL_COL + ' (' + headers[PHOTO_URL_COL - 1] + ')');
}

function testCheckIn() {
  Logger.log('=== TESTING CHECK-IN ===');
  
  var result = checkInGuest({
    zip: '64110',
    gender: 'man',
    dob: '10/07'
  });
  
  Logger.log('Result: ' + JSON.stringify(result, null, 2));
  
  if (result.ok) {
    Logger.log('✅ SUCCESS: Guest checked in');
  } else {
    Logger.log('❌ FAILED: ' + result.message);
  }
}

/**
 * COMPREHENSIVE TEST SUITE
 * Copy this into your Code.gs file at the bottom
 * Run testEverything() to validate your entire setup
 */

/**
 * Master test function - runs all tests
 * Run this to validate your entire system
 */
function testEverything() {
  Logger.log('========================================');
  Logger.log('🧪 COMPREHENSIVE SYSTEM TEST');
  Logger.log('========================================');
  
  var allPassed = true;
  
  // Test 1: Sheet Structure
  Logger.log('TEST 1: Sheet Structure');
  Logger.log('------------------------');
  if (!testSheetColumns()) {
    allPassed = false;
  }
  Logger.log('');
  
  // Test 2: Backend Functions
  Logger.log('TEST 2: Backend Functions');
  Logger.log('------------------------');
  if (!testBackendFunctions()) {
    allPassed = false;
  }
  Logger.log('');
  
  // Test 3: HTML/Code Alignment
  Logger.log('TEST 3: HTML/Code.gs Alignment');
  Logger.log('------------------------');
  if (!testHTMLAlignment()) {
    allPassed = false;
  }
  Logger.log('');
  
  // Test 4: Sample Check-In
  Logger.log('TEST 4: Sample Check-In (with real data)');
  Logger.log('------------------------');
  if (!testSampleCheckIn()) {
    allPassed = false;
  }
  Logger.log('');
  
  // Final Results
  Logger.log('========================================');
  if (allPassed) {
    Logger.log('✅ ALL TESTS PASSED!');
    Logger.log('Your system is ready to deploy!');
  } else {
    Logger.log('❌ SOME TESTS FAILED');
    Logger.log('Please fix the issues above before deploying.');
  }
  Logger.log('========================================');
  
  return allPassed;
}

/**
 * Test 1: Verify sheet has all required columns
 */
function testSheetColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  
  if (!sh) {
    Logger.log('❌ FAIL: Sheet "' + SHEET_NAME + '" not found');
    Logger.log('   → Run DataClean.gs first!');
    return false;
  }
  
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var totalCols = headers.length;
  
  Logger.log('Sheet: ' + SHEET_NAME);
  Logger.log('Total columns: ' + totalCols);
  
  // Check required columns
  var required = [
    {col: ZIP_COL, name: 'Current 5 Digit Zip Code'},
    {col: SCREEN_NAME_COL, name: 'Screen Name'},
    {col: UID_COL, name: 'UID'},
    {col: CHECKED_FLAG_COL, name: 'Checked-In at Event'},
    {col: CHECKED_TS_COL, name: 'Check-In Timestamp'},
    {col: PHOTO_URL_COL, name: 'Photo URL'}
  ];
  
  var allFound = true;
  
  for (var req of required) {
    var actualName = headers[req.col - 1];
    
    if (!actualName) {
      Logger.log('❌ FAIL: Column ' + req.col + ' (' + req.name + ') is MISSING');
      allFound = false;
    } else if (actualName !== req.name) {
      Logger.log('⚠️  WARN: Column ' + req.col + ' is "' + actualName + '" (expected "' + req.name + '")');
    } else {
      Logger.log('✅ Column ' + req.col + ': ' + req.name);
    }
  }
  
  if (!allFound) {
    Logger.log('💡 FIX: Run cleanFormResponses() in DataClean.gs to create missing columns');
    return false;
  }
  
  Logger.log('✅ All required columns exist!');
  return true;
}

/**
 * Test 2: Verify backend functions exist and are callable
 */
function testBackendFunctions() {
  var requiredFunctions = [
    'checkInGuest',
    'updateGuestScreenName',
    'uploadGuestPhoto',
    'doGet'
  ];
  
  var allExist = true;
  
  for (var funcName of requiredFunctions) {
    if (typeof this[funcName] === 'function') {
      Logger.log('✅ ' + funcName + '() exists');
    } else {
      Logger.log('❌ FAIL: ' + funcName + '() NOT FOUND');
      allExist = false;
    }
  }
  
  if (!allExist) {
    Logger.log('💡 FIX: Make sure Code.gs has all required functions');
    return false;
  }
  
  // Test doGet returns HTML
  try {
    var output = doGet();
    if (output && output.getContent) {
      Logger.log('✅ doGet() returns valid HTML output');
    } else {
      Logger.log('❌ FAIL: doGet() does not return HTML');
      return false;
    }
  } catch (e) {
    Logger.log('❌ FAIL: doGet() threw error: ' + e.message);
    return false;
  }
  
  Logger.log('✅ All backend functions exist!');
  return true;
}

/**
 * Test 4: Test actual check-in with sample data
 * FIXED VERSION - properly converts Date objects to MM/DD format
 */
function testSampleCheckIn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  
  if (!sh) {
    Logger.log('❌ FAIL: Cannot test - sheet not found');
    return false;
  }
  
  var data = sh.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('⚠️  SKIP: No guest data to test with');
    Logger.log('   Add some guests to "Form Responses (Clean)" first');
    return true; // Not a failure, just no data
  }
  
  // Get first guest's data
  var headers = data[0];
  var firstGuest = data[1];
  
  var zipCol = headers.indexOf('Current 5 Digit Zip Code');
  var genderCol = headers.indexOf('Self-Identified Gender');
  var dobCol = headers.indexOf('Birthday (MM/DD)');
  var screenNameCol = headers.indexOf('Screen Name');
  
  if (zipCol === -1 || genderCol === -1 || dobCol === -1) {
    Logger.log('❌ FAIL: Required columns not found in sheet');
    return false;
  }
  
  var testZip = String(firstGuest[zipCol] || '').trim();
  var testGender = String(firstGuest[genderCol] || '').trim().toLowerCase();
  var rawDOB = firstGuest[dobCol];
  var expectedName = String(firstGuest[screenNameCol] || '').trim();
  
  // CRITICAL FIX: Convert DOB to MM/DD format
  var formattedDOB;
  
  if (rawDOB instanceof Date) {
    // If it's a Date object, extract month and day
    var month = String(rawDOB.getMonth() + 1).padStart(2, '0');
    var day = String(rawDOB.getDate()).padStart(2, '0');
    formattedDOB = month + '/' + day;
    Logger.log('Converted Date object to MM/DD format');
  } else {
    // If it's a string, extract MM/DD only
    var dobStr = String(rawDOB).trim();
    var dobParts = dobStr.split('/');
    if (dobParts.length >= 2) {
      var month = dobParts[0].padStart(2, '0');
      var day = dobParts[1].padStart(2, '0');
      formattedDOB = month + '/' + day;
    } else {
      formattedDOB = dobStr;
    }
  }
  
  Logger.log('Testing with first guest in sheet:');
  Logger.log('  ZIP: ' + testZip);
  Logger.log('  Gender: ' + testGender);
  Logger.log('  DOB (raw): ' + rawDOB);
  Logger.log('  DOB (formatted): ' + formattedDOB);
  Logger.log('  Expected Name: ' + expectedName);
  
  if (!testZip || !testGender || !formattedDOB) {
    Logger.log('⚠️  SKIP: Guest data incomplete - cannot test');
    return true;
  }
  
  // Test check-in with properly formatted data
  try {
    var result = checkInGuest({
      zip: testZip,
      gender: testGender,
      dob: formattedDOB
    });
    
    Logger.log('
Check-in result:');
    Logger.log(JSON.stringify(result, null, 2));
    
    if (result.ok) {
      Logger.log('✅ Check-in successful!');
      Logger.log('   Screen Name: ' + result.screenName);
      Logger.log('   UID: ' + result.uid);
      
      if (result.screenName !== expectedName) {
        Logger.log('⚠️  Note: Screen name differs (may have been updated)');
      }
      
      return true;
    } else {
      Logger.log('❌ FAIL: Check-in failed');
      Logger.log('   Message: ' + result.message);
      return false;
    }
    
  } catch (e) {
    Logger.log('❌ FAIL: Check-in threw error');
    Logger.log('   Error: ' + e.message);
    Logger.log('   Stack: ' + e.stack);
    return false;
  }
}

/**
 * ALSO UPDATE testHTMLAlignment to be less strict
 */
function testHTMLAlignment() {
  var allAligned = true;
  
  // Test that HTML file exists
  try {
    var html = HtmlService.createHtmlOutputFromFile('CheckInInterface');
    var content = html.getContent();
    
    Logger.log('✅ CheckInInterface.html exists and is loadable');
    
    // Check if HTML contains required elements
    var requiredElements = [
      'checkInForm',
      'zip',
      'gender',
      'dob',
      'checkInBtn',
      'guestInfo',
      'photoInput'
    ];
    
    for (var elemId of requiredElements) {
      if (content.includes('id="' + elemId + '"')) {
        Logger.log('✅ HTML has element: ' + elemId);
      } else {
        Logger.log('❌ FAIL: HTML missing element: ' + elemId);
        allAligned = false;
      }
    }
    
    // Check if HTML calls correct backend functions (more flexible check)
    var requiredCalls = [
      {name: 'checkInGuest', pattern: /google\.script\.run[.\s\S]*checkInGuest/},
      {name: 'updateGuestScreenName', pattern: /google\.script\.run[.\s\S]*updateGuestScreenName/},
      {name: 'uploadGuestPhoto', pattern: /google\.script\.run[.\s\S]*uploadGuestPhoto/}
    ];
    
    for (var call of requiredCalls) {
      if (call.pattern.test(content)) {
        Logger.log('✅ HTML calls: ' + call.name);
      } else {
        Logger.log('⚠️  WARN: HTML might not call: ' + call.name);
        Logger.log('   (This is OK if the function name appears in HTML somewhere)');
        // Don't fail on this - just warn
        if (!content.includes(call.name)) {
          Logger.log('❌ FAIL: Function name "' + call.name + '" not found anywhere in HTML');
          allAligned = false;
        }
      }
    }
    
  } catch (e) {
    Logger.log('❌ FAIL: Cannot load CheckInInterface.html');
    Logger.log('   Error: ' + e.message);
    Logger.log('
💡 FIX: Make sure CheckInInterface.html exists with proper HTML content');
    return false;
  }
  
  if (!allAligned) {
    Logger.log('💡 FIX: Replace CheckInInterface.html with the HTML artifact provided');
    return false;
  }
  
  Logger.log('✅ HTML and Code.gs are aligned!');
  return true;
}

/**
 * Quick test with your own data
 * UPDATE THESE VALUES and run this function
 */
function testCheckInWithMyData() {
  Logger.log('========================================');
  Logger.log('🧪 CUSTOM CHECK-IN TEST');
  Logger.log('========================================');
  
  // ⬇️⬇️⬇️ UPDATE THESE WITH REAL DATA FROM YOUR SHEET ⬇️⬇️⬇️
  var testData = {
    zip: '64110',      // ← Change this
    gender: 'man',     // ← Change this (man/woman/nonbinary/other)
    dob: '10/07'       // ← Change this (MM/DD format)
  };
  // ⬆️⬆️⬆️ UPDATE THESE WITH REAL DATA FROM YOUR SHEET ⬆️⬆️⬆️
  
  Logger.log('Testing check-in with:');
  Logger.log('  ZIP: ' + testData.zip);
  Logger.log('  Gender: ' + testData.gender);
  Logger.log('  DOB: ' + testData.dob);
  Logger.log('');
  
  try {
    var result = checkInGuest(testData);
    
    Logger.log('Result:');
    Logger.log(JSON.stringify(result, null, 2));
    Logger.log('');
    
    if (result.ok) {
      Logger.log('✅ SUCCESS!');
      Logger.log('   Screen Name: ' + result.screenName);
      Logger.log('   UID: ' + result.uid);
      Logger.log('   Already Checked In: ' + (result.alreadyCheckedIn ? 'Yes' : 'No'));
    } else {
      Logger.log('❌ FAILED');
      Logger.log('   Message: ' + result.message);
      Logger.log('');
      Logger.log('💡 TROUBLESHOOTING:');
      Logger.log('   1. Verify the guest exists in "Form Responses (Clean)"');
      Logger.log('   2. Check ZIP is exact match (5 digits, no spaces)');
      Logger.log('   3. Gender must be lowercase: man/woman/nonbinary/other');
      Logger.log('   4. DOB must be MM/DD format (e.g., 03/15 not 3/15)');
    }
    
  } catch (e) {
    Logger.log('❌ ERROR');
    Logger.log('   ' + e.message);
    Logger.log('   Stack: ' + e.stack);
  }
  
  Logger.log('========================================');
}

/**
 * Test screen name update function
 */
function testScreenNameUpdate() {
  Logger.log('========================================');
  Logger.log('🧪 SCREEN NAME UPDATE TEST');
  Logger.log('========================================');
  
  // ⬇️⬇️⬇️ UPDATE THESE ⬇️⬇️⬇️
  var testUID = 'BW-902';           // ← Change to a real UID from your sheet
  var newName = 'TestName123';       // ← Change to your desired test name
  // ⬆️⬆️⬆️ UPDATE THESE ⬆️⬆️⬆️
  
  Logger.log('Testing screen name update:');
  Logger.log('  UID: ' + testUID);
  Logger.log('  New Name: ' + newName);
  Logger.log('');
  
  try {
    var result = updateGuestScreenName({
      uid: testUID,
      newScreenName: newName
    });
    
    Logger.log('Result:');
    Logger.log(JSON.stringify(result, null, 2));
    Logger.log('');
    
    if (result.ok) {
      Logger.log('✅ SUCCESS! Name updated to: ' + result.newScreenName);
    } else {
      Logger.log('❌ FAILED: ' + result.message);
    }
    
  } catch (e) {
    Logger.log('❌ ERROR: ' + e.message);
  }
  
  Logger.log('========================================');
}

/**
 * Debug Check-In - Shows exactly what's being compared
 * Add this to Code.gs and run it to see why guests aren't matching
 */
function debugCheckIn() {
  Logger.log('========================================');
  Logger.log('🔍 CHECK-IN DEBUG MODE');
  Logger.log('========================================');
  
  // ⬇️⬇️⬇️ ENTER THE DATA YOU'RE TRYING TO CHECK IN WITH ⬇️⬇️⬇️
  var searchData = {
    zip: '64110',      // ← What you entered in the form
    gender: 'man',     // ← What you selected
    dob: '05/18'       // ← What you entered (MM/DD)
  };
  // ⬆️⬆️⬆️ CHANGE THESE TO MATCH WHAT YOU TRIED ⬆️⬆️⬆️
  
  Logger.log('🔎 SEARCHING FOR:');
  Logger.log('  ZIP: "' + searchData.zip + '"');
  Logger.log('  Gender: "' + searchData.gender + '"');
  Logger.log('  DOB: "' + searchData.dob + '"');
  Logger.log('');
  
  // Normalize search data (same way the check-in function does)
  var wantZip = normalizeZip_(searchData.zip);
  var wantGender = normalizeGender_(searchData.gender);
  var wantDOB = parseMonthDay_(searchData.dob);
  
  Logger.log('🔧 AFTER NORMALIZATION:');
  Logger.log('  ZIP: "' + wantZip + '"');
  Logger.log('  Gender: "' + wantGender + '"');
  Logger.log('  DOB: Month=' + wantDOB.m + ', Day=' + wantDOB.d);
  Logger.log('');
  Logger.log('========================================');
  Logger.log('📊 CHECKING SHEET DATA:');
  Logger.log('========================================
');
  
  // Get sheet data
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  
  if (!sh) {
    Logger.log('❌ ERROR: Sheet not found!');
    return;
  }
  
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(safeString_);
  
  // Find columns
  var genderCol = findHeaderIndex_(headers, ['self-identified gender', 'gender', 'sex']);
  var dobCol = findHeaderIndex_(headers, ['birthday (mm/dd)', 'birthday', 'dob', 'date of birth']);
  
  Logger.log('Column positions:');
  Logger.log('  ZIP: Column ' + ZIP_COL + ' (0-indexed: ' + (ZIP_COL - 1) + ')');
  Logger.log('  Gender: Column ' + (genderCol + 1) + ' (0-indexed: ' + genderCol + ')');
  Logger.log('  DOB: Column ' + (dobCol + 1) + ' (0-indexed: ' + dobCol + ')');
  Logger.log('');
  
  // Check first 5 rows
  Logger.log('First 5 guests in sheet:');
  Logger.log('----------------------------------------');
  
  for (var r = 1; r < Math.min(6, data.length); r++) {
    var row = data[r];
    
    var rowZipRaw = row[ZIP_COL - 1];
    var rowZip = normalizeZip_(safeString_(rowZipRaw));
    
    var rowGenderRaw = row[genderCol];
    var rowGender = normalizeGender_(safeString_(rowGenderRaw));
    
    var rowDOBRaw = row[dobCol];
    var rowDOB = monthDayFromCell_(rowDOBRaw);
    
    var rowScreenName = row[SCREEN_NAME_COL - 1];
    var rowUID = row[UID_COL - 1];
    
    Logger.log('Row ' + (r + 1) + ': ' + rowScreenName + ' (' + rowUID + ')');
    Logger.log('  ZIP (raw): "' + rowZipRaw + '"');
    Logger.log('  ZIP (normalized): "' + rowZip + '"');
    Logger.log('  Gender (raw): "' + rowGenderRaw + '"');
    Logger.log('  Gender (normalized): "' + rowGender + '"');
    Logger.log('  DOB (raw): ' + rowDOBRaw);
    Logger.log('  DOB (parsed): Month=' + (rowDOB ? rowDOB.m : 'null') + ', Day=' + (rowDOB ? rowDOB.d : 'null'));
    
    // Check if this row matches
    var zipMatch = rowZip === wantZip;
    var genderMatch = rowGender === wantGender;
    var dobMatch = rowDOB && rowDOB.m === wantDOB.m && rowDOB.d === wantDOB.d;
    
    Logger.log('  MATCH CHECK:');
    Logger.log('    ZIP Match: ' + (zipMatch ? '✅' : '❌') + ' (' + rowZip + ' vs ' + wantZip + ')');
    Logger.log('    Gender Match: ' + (genderMatch ? '✅' : '❌') + ' (' + rowGender + ' vs ' + wantGender + ')');
    Logger.log('    DOB Match: ' + (dobMatch ? '✅' : '❌') + ' (' + (rowDOB ? rowDOB.m + '/' + rowDOB.d : 'null') + ' vs ' + wantDOB.m + '/' + wantDOB.d + ')');
    
    if (zipMatch && genderMatch && dobMatch) {
      Logger.log('  🎉 FULL MATCH! This guest should check in successfully.');
    }
    
    Logger.log('');
  }
  
  Logger.log('========================================');
  Logger.log('💡 TROUBLESHOOTING TIPS:');
  Logger.log('========================================');
  Logger.log('1. ZIP Code: Must be exactly 5 digits');
  Logger.log('2. Gender: Must match exactly after normalization');
  Logger.log('   - Sheet has: "Man" → normalized to "man"');
  Logger.log('   - You enter: "man" → normalized to "man" ✅');
  Logger.log('3. Birthday: Must match month AND day');
  Logger.log('   - If DOB shows "null", the format is wrong in the sheet');
  Logger.log('   - Run DataClean.gs to fix birthday formats');
  Logger.log('4. If no rows match, try:');
  Logger.log('   - Use data from Row 2 exactly as shown above');
  Logger.log('   - Re-run DataClean.gs to refresh the sheet');
  Logger.log('========================================');
}

/**
 * Quick function to show ALL guests in your sheet
 * Use this to find a guest to test with
 */
function showAllGuests() {
  Logger.log('========================================');
  Logger.log('📋 ALL GUESTS IN SHEET');
  Logger.log('========================================');
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  
  if (!sh) {
    Logger.log('❌ ERROR: Sheet not found!');
    return;
  }
  
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  
  var genderCol = headers.indexOf('Self-Identified Gender');
  var dobCol = headers.indexOf('Birthday (MM/DD)');
  
  Logger.log('Total guests: ' + (data.length - 1));
  Logger.log('');
  
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    
    var zip = row[ZIP_COL - 1];
    var gender = row[genderCol];
    var dob = row[dobCol];
    var screenName = row[SCREEN_NAME_COL - 1];
    var uid = row[UID_COL - 1];
    
    // Convert DOB if it's a Date object
    var dobFormatted;
    if (dob instanceof Date) {
      var month = String(dob.getMonth() + 1).padStart(2, '0');
      var day = String(dob.getDate()).padStart(2, '0');
      dobFormatted = month + '/' + day;
    } else {
      dobFormatted = String(dob);
    }
    
    Logger.log('Row ' + (r + 1) + ': ' + screenName + ' (' + uid + ')');
    Logger.log('  Use this to check in:');
    Logger.log('    ZIP: ' + zip);
    Logger.log('    Gender: ' + String(gender).toLowerCase());
    Logger.log('    DOB: ' + dobFormatted);
    Logger.log('');
  }
  
  Logger.log('========================================');
  Logger.log('Copy the data from Row 2 above and use it in the web form!');
  Logger.log('========================================');
}


/**
 * Get all checked-in guests for wall display
 * Returns: [{uid, screenName, checkedAt}, ...]
 */
function getWallData() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(WALL_CLEAN_SHEET);
    
    if (!sheet) {
      throw new Error('Clean sheet not found');
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    
    var headers = data[0];
    var colMap = {};
    headers.forEach(function(h, i) {
      colMap[String(h).trim()] = i;
    });
    
    var uidCol = colMap['UID'];
    var screenNameCol = colMap['Screen Name'];
    var checkedInCol = colMap['Checked-In at Event'];
    var timestampCol = colMap['Check-In Timestamp'];
    
    if (uidCol === undefined || screenNameCol === undefined) {
      throw new Error('Required columns not found');
    }
    
    var guests = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Only guests who actually checked in at event
      var checkedInAtEvent = String(row[checkedInCol] || '').trim().toUpperCase();
      if (checkedInAtEvent !== 'Y') {
        continue;
      }
      
      var uid = String(row[uidCol] || '').trim();
      var screenName = String(row[screenNameCol] || '').trim();
      
      if (!uid || !screenName) continue;
      
      // Format check-in time
      var checkedAt = 'PRESENT';
      if (timestampCol !== undefined && row[timestampCol]) {
        var timestamp = row[timestampCol];
        if (timestamp instanceof Date) {
          checkedAt = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'HH:mm');
        } else if (typeof timestamp === 'string') {
          var match = timestamp.match(/(\d{1,2}):(\d{2})/);
          if (match) {
            checkedAt = match[0];
          }
        }
      }
      
      guests.push({
        uid: uid,
        screenName: screenName,
        checkedAt: checkedAt
      });
    }
    
    return guests;
    
  } catch (error) {
    Logger.log('Error in getWallData: ' + error.toString());
    throw error;
  }
}

/**
 * Analyze all categories and generate connection pairs
 * Returns: { analyses: [...], topCommonalities: [...] }
 */
function getDetailedWallConnections() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(WALL_CLEAN_SHEET);
    
    if (!sheet) {
      return { error: 'Clean sheet not found' };
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { error: 'No data found' };
    }
    
    var headers = data[0];
    var colMap = {};
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i];
      colMap[String(h).trim()] = i;
    }
    
    // Get checked-in guests only
    var guests = [];
    var checkedInCol = colMap['Checked-In at Event'];
    var uidCol = colMap['UID'];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var checkedIn = String(row[checkedInCol] || '').trim().toUpperCase();
      if (checkedIn !== 'Y') continue;
      
      var uid = String(row[uidCol] || '').trim();
      if (!uid) continue;
      
      // Build guest with all demographics
      var guest = { uid: uid, demographics: {} };
      
      for (var j = 0; j < ANALYSIS_CATEGORIES.length; j++) {
        var category = ANALYSIS_CATEGORIES[j];
        var colIdx = resolveColumnIndex_(colMap, category.field);
        if (colIdx !== undefined) {
          var value = String(row[colIdx] || '').trim();
          if (value && value !== 'Not Provided' && value !== '') {
            guest.demographics[category.name] = value;
          }
        }
      }
      
      guests.push(guest);
    }
    
    if (guests.length < 2) {
      return { error: 'Insufficient guests for analysis' };
    }
    
    // Analyze each category
    var analyses = [];
    
    ANALYSIS_CATEGORIES.forEach(category => {
      // Group guests by value
      var valueGroups = {};
      
      guests.forEach(guest => {
        var value = guest.demographics[category.name];
        if (value) {
          if (!valueGroups[value]) {
            valueGroups[value] = [];
          }
          valueGroups[value].push(guest.uid);
        }
      });
      
      // Create connections for each value
      Object.keys(valueGroups).forEach(value => {
        var uids = valueGroups[value];
        
        if (uids.length < 2) return;
        
        // Create all possible pairs
        var connections = [];
        for (var i = 0; i < uids.length; i++) {
          for (var j = i + 1; j < uids.length; j++) {
            connections.push({
              uid1: uids[i],
              uid2: uids[j]
            });
          }
        }
        
        analyses.push({
          analysisType: category.name,
          traitValue: value,
          connections: connections
        });
      });
    });
    
    // Calculate top commonalities
    var valueCounts = {};
    
    guests.forEach(function(guest) {
      Object.keys(guest.demographics).forEach(function(category) {
        var value = guest.demographics[category];
        
        if (!valueCounts[value]) {
          valueCounts[value] = 0;
        }
        valueCounts[value]++;
      });
    });
    
    var keys = Object.keys(valueCounts);
    var topCommonalities = [];
    for (var i = 0; i < keys.length; i++) {
      topCommonalities.push({
        name: keys[i],
        count: valueCounts[keys[i]]
      });
    }
    topCommonalities.sort(function(a, b) {
      return b.count - a.count;
    });
    topCommonalities = topCommonalities.slice(0, 10);
    
    return {
      analyses: analyses,
      topCommonalities: topCommonalities
    };
    
  } catch (error) {
    Logger.log('Error in getDetailedWallConnections: ' + error.toString());
    return { error: error.toString() };
  }
}

/**
 * Get scrolling intro text for header ticker
 */
function getIntroText() {
  return 'NETWORK OPTIMIZATION SYSTEM ACTIVE // ANALYZING SOCIAL CONNECTIONS // ' +
         'DEMOGRAPHIC INTEGRATION PROTOCOL ENABLED // REAL-TIME PATTERN DETECTION // ' +
         'CURATED EXPERIENCE OPTIMIZATION IN PROGRESS';
}

/**
 * Get geographic distribution data for map view
 * Returns: { target: {...}, zips: [...], totalCount: n }
 */
function getAllZipData() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(WALL_CLEAN_SHEET);
    
    if (!sheet) {
      return { error: 'Clean sheet not found' };
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { error: 'No data found' };
    }
    
    var headers = data[0];
    var colMap = {};
    headers.forEach((h, i) => colMap[String(h).trim()] = i);
    
    var zipCol = colMap['Current 5 Digit Zip Code'];
    var checkedInCol = colMap['Checked-In at Event'];
    
    if (zipCol === undefined) {
      return { error: 'Zip code column not found' };
    }
    
    // Count guests per zip (checked-in only)
    var zipCounts = {};
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var checkedIn = String(row[checkedInCol] || '').trim().toUpperCase();
      if (checkedIn !== 'Y') continue;
      
      var zip = String(row[zipCol] || '').trim();
      if (zip && zip.length === 5 && /^\d{5}$/.test(zip)) {
        zipCounts[zip] = (zipCounts[zip] || 0) + 1;
      }
    }
    
    // Geocode target
    var targetGeo = geocodeAddress(WALL_TARGET_ADDRESS);
    
    // Geocode all guest zips
    var zipLocations = [];
    var zips = Object.keys(zipCounts);
    for (var i = 0; i < zips.length; i++) {
      var zip = zips[i];
      var geo = geocodeZip(zip);
      if (geo && geo.lat && geo.lng) {
        zipLocations.push({
          zip: zip,
          count: zipCounts[zip],
          lat: geo.lat,
          lng: geo.lng
        });
      }
    }
    
    return {
      target: {
        lat: targetGeo.lat,
        lng: targetGeo.lng,
        displayName: '5317 Charlotte',
        count: (function() {
          var sum = 0;
          for (var key in zipCounts) {
            if (zipCounts.hasOwnProperty(key)) {
              sum += zipCounts[key];
            }
          }
          return sum;
        })()
      },
      zips: zipLocations,
      totalCount: Object.keys(zipCounts).length
    };
    
  } catch (error) {
    Logger.log('Error in getAllZipData: ' + error.toString());
    return { error: error.toString() };
  }
}

/**
 * Geocode full address
 */
function geocodeAddress(address) {
  try {
    var response = Maps.newGeocoder().geocode(address);
    if (response.results && response.results.length > 0) {
      var location = response.results[0].geometry.location;
      return {
        lat: location.lat,
        lng: location.lng
      };
    }
  } catch (error) {
    Logger.log('Geocoding error for address ' + address + ': ' + error);
  }
  
  // Default to KC center
  return { lat: 39.0997, lng: -94.5786 };
}

/**
 * Geocode zip code
 */
function geocodeZip(zip) {
  try {
    var response = Maps.newGeocoder().geocode(zip + ', USA');
    if (response.results && response.results.length > 0) {
      var location = response.results[0].geometry.location;
      return {
        lat: location.lat,
        lng: location.lng
      };
    }
  } catch (error) {
    Logger.log('Geocoding error for zip ' + zip + ': ' + error);
  }
  
  return null;
}

// Test functions
function testWallData() {
  var guests = getWallData();
  Logger.log('Total guests: ' + guests.length);
  if (guests.length > 0) {
    Logger.log('Sample guest: ' + JSON.stringify(guests[0]));
  }
}

function testWallConnections() {
  var result = getDetailedWallConnections();
  if (result.error) {
    Logger.log('Error: ' + result.error);
  } else {
    Logger.log('Total analyses: ' + result.analyses.length);
    Logger.log('Top commonality: ' + JSON.stringify(result.topCommonalities[0]));
  }
}


// In your Google Apps Script (Code.gs)

function getDDDViolations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dddSheet = ss.getSheetByName('DDD (Checked-In Only)');
  
  if (!dddSheet) {
    Logger.log('DDD sheet not found');
    return [];
  }
  
  var data = dddSheet.getDataRange().getValues();
  var headers = data[0];
  var violations = [];
  
  // Column indices based on Master_Desc
  var uidCol = 1; // Column 2: UID
  var violationColumns = {
    'Birthday Format': 2,
    'More Than 3 Interests': 3,
    'Unknown Guest': 4, // This is "Do Not Know Host"
    'No Song Request': 5,
    'Vague Song Request': 6,
    'Multiple Songs Listed': 7,
    'No Artist Listed': 8,
    'Barely Knows Host': 9,
    'Fresh Acquaintance': 10,
    'Host Ambiguity': 11,
    'Age/Birthday Mismatch': 12,
    'Education/Career Implausibility': 13,
    'Meme/Joke Response': 14,
    'Music/Artist Genre Mismatch': 15,
    'Minimum Effort Detected': 16,
    'Stranger Danger': 17,
    'Relationship Contradiction': 18,
    'Special Snowflake Syndrome': 19,
    'Suspiciously Generic': 20
  };
  
  // Skip header row
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var uid = row[uidCol];
    
    if (!uid) continue;
    
    // Check each violation column
    Object.keys(violationColumns).forEach(violationType => {
      var colIndex = violationColumns[violationType];
      var hasViolation = row[colIndex] === 1 || row[colIndex] === '1';
      
      if (hasViolation) {
        violations.push({
          uid: uid,
          type: violationType,
          description: getViolationDescription(violationType)
        });
      }
    });
  }
  
  Logger.log('Found ' + violations.length + ' violations');
  return violations;
}

function getViolationDescription(type) {
  var descriptions = {
    'Birthday Format': 'Invalid birthday format detected',
    'More Than 3 Interests': 'Listed more than 3 interests (max 3 allowed)',
    'Unknown Guest': 'Does not know the host(s)',
    'No Song Request': 'Did not provide song request',
    'Vague Song Request': 'Song request too vague or unclear',
    'Multiple Songs Listed': 'Listed multiple songs instead of one',
    'No Artist Listed': 'Did not provide favorite artist',
    'Barely Knows Host': 'Knows host less than 3 months',
    'Fresh Acquaintance': 'Very new acquaintance with host',
    'Host Ambiguity': 'Unclear about which host they know',
    'Age/Birthday Mismatch': 'Age range does not match birthday',
    'Education/Career Implausibility': 'Education level inconsistent with career role',
    'Meme/Joke Response': 'Provided meme or joke responses',
    'Music/Artist Genre Mismatch': 'Artist does not match stated music preference',
    'Minimum Effort Detected': 'Minimal effort in form responses',
    'Stranger Danger': 'Multiple red flags indicating unknown guest',
    'Relationship Contradiction': 'Contradictory information about host relationship',
    'Special Snowflake Syndrome': 'Overly unique or attention-seeking responses',
    'Suspiciously Generic': 'Responses are too generic or template-like'
  };
  
  return descriptions[type] || 'Violation detected';
}

function getWallData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cleanSheet = ss.getSheetByName('Form Responses (Clean)');
  
  if (!cleanSheet) {
    Logger.log('Form Responses (Clean) sheet not found');
    return [];
  }
  
  var data = cleanSheet.getDataRange().getValues();
  var guests = [];
  
  // Based on Master_Desc: UID is column 23, Check-in Time is column 25
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var uid = row[22]; // Column 23 (0-indexed = 22)
    var checkedIn = row[23]; // Column 24
    var checkInTime = row[24]; // Column 25
    
    if (uid && checkedIn === 'Y') {
      guests.push({
        uid: uid,
        checkedAt: formatTime(checkInTime)
      });
    }
  }
  
  Logger.log('Found ' + guests.length + ' checked-in guests');
  return guests;
}

function formatTime(timestamp) {
  if (!timestamp) return 'N/A';
  var date = new Date(timestamp);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm');
}

/**
 * ============================================================================
 * REPORTS.GS - ALL PARTY SURVEY ANALYSIS REPORTS
 * ============================================================================
 * * Contains all report generation functions for party survey analysis.
 * Each function creates a formatted output sheet with analysis results.
 * * AVAILABLE REPORTS:
 * 1. buildDemographicsSummary() - Age, gender, ethnicity, education breakdown
 * 2. buildMusicInterestsReport() - Music preferences and general interests
 * 3. buildHostRelationshipReport() - Guest relationships with party hosts
 * 4. buildAttendeeAnalysis() - Compare attendees vs no-shows
 * 5. buildGuestProfiles() - Individual guest profile cards
 * * USAGE:
 * Run any function from the script editor or create a custom menu
 */

// ============================================================================
// REPORT 1: DEMOGRAPHICS SUMMARY
// ============================================================================

/**
 * Generate comprehensive demographic overview
 * Shows distributions across age, gender, ethnicity, education, employment
 * Output: Demographics_Summary sheet
 */
function buildDemographicsSummary() {
  var ss = SpreadsheetApp.getActive();
  var cleanSheet = ss.getSheetByName(REPORTS_CLEAN_SHEET);
  
  if (!cleanSheet) {
    throw new Error('Sheet "' + REPORTS_CLEAN_SHEET + '" not found.');
  }

  var data = cleanSheet.getDataRange().getValues();
  if (data.length < 2) {
    writeSheet_('Demographics_Summary', [['No data available']]);
    return;
  }

  var header = data[0];
  var colIdx = getColumnMap_(header);

  // Calculate overview metrics
  var totalCount = data.length - 1;
  var checkedInCount = data.slice(1).filter(row => 
    String(row[colIdx['Checked-In at Event']] || '').trim().toUpperCase() === 'Y'
  ).length;

  // Build report sections
  var sections = [];

  // Title and overview
  sections.push(['PARTY SURVEY DEMOGRAPHICS SUMMARY']);
  sections.push(['']);
  sections.push(['Metric', 'Count', 'Percentage']);
  sections.push(['Total Responses', totalCount, '100.0%']);
  sections.push(['Checked-In Attendees', checkedInCount, ((checkedInCount / totalCount) * 100).toFixed(1) + '%']);
  sections.push(['No-Shows', totalCount - checkedInCount, (((totalCount - checkedInCount) / totalCount) * 100).toFixed(1) + '%']);
  sections.push(['']);

  // Add distribution sections
  var demographicFields = [
    ['Age Range', 'AGE RANGE DISTRIBUTION'],
    ['Self-Identified Gender', 'GENDER DISTRIBUTION'],
    ['Self Identified Ethnicity', 'ETHNICITY DISTRIBUTION'],
    ['Education Level', 'EDUCATION LEVEL DISTRIBUTION'],
    ['Self-Identified Sexual Orientation', 'SEXUAL ORIENTATION DISTRIBUTION'],
    ['Employment Information (Industry)', 'INDUSTRY DISTRIBUTION'],
    ['Employment Information (Role)', 'ROLE DISTRIBUTION']
  ];

  for (var i = 0; i < demographicFields.length; i++) {
    var field = demographicFields[i];
    var colName = field[0];
    var title = field[1];
    Array.prototype.push.apply(sections, buildDistribution_(data, colIdx, colName, title, 'Checked-In at Event'));
    sections.push(['']);
  });

  writeSheet_('Demographics_Summary', sections);
  formatDemographicsSheet_();
}

/**
 * Build distribution table for a demographic variable
 */
function buildDistribution_(data, colIdx, colName, title, checkedInCol) {
  var dataColIdx = colIdx[colName];
  var checkedInIdx = colIdx[checkedInCol];
  
  if (dataColIdx === undefined) {
    return [[title], ['Column not found']];
  }

  var totalCounts = {};
  var checkedInCounts = {};

  for (var i = 1; i < data.length; i++) {
    var value = String(data[i][dataColIdx] || '').trim();
    var isCheckedIn = String(data[i][checkedInIdx] || '').trim().toUpperCase() === 'Y';

    if (!value) continue;

    totalCounts[value] = (totalCounts[value] || 0) + 1;
    if (isCheckedIn) {
      checkedInCounts[value] = (checkedInCounts[value] || 0) + 1;
    }
  }

  var sortedValues = Object.keys(totalCounts).sort((a, b) => totalCounts[b] - totalCounts[a]);
  var totalSum = 0;
  for (var key in totalCounts) {
    if (totalCounts.hasOwnProperty(key)) {
      totalSum += totalCounts[key];
    }
  }
  var checkedInSum = 0;
  for (var key in checkedInCounts) {
    if (checkedInCounts.hasOwnProperty(key)) {
      checkedInSum += checkedInCounts[key];
    }
  }

  var rows = [
    [title],
    ['Category', 'Total', '% of Total', 'Checked-In', '% of Checked-In']
  ];

  sortedValues.forEach(value => {
    var totalCount = totalCounts[value];
    var checkedInCount = checkedInCounts[value] || 0;
    var totalPct = ((totalCount / totalSum) * 100).toFixed(1);
    var checkedInPct = checkedInSum > 0 ? ((checkedInCount / checkedInSum) * 100).toFixed(1) : '0.0';

    rows.push([value, totalCount, totalPct + '%', checkedInCount, checkedInPct + '%']);
  });

  rows.push(['TOTAL', totalSum, '100.0%', checkedInSum, '100.0%']);

  return rows;
}

function formatDemographicsSheet_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Demographics_Summary');
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  
  for (var row = 1; row <= data.length; row++) {
    var cellValue = String(data[row - 1][0] || '').trim();
    
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

  for (var col = 1; col <= 5; col++) {
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
  var ss = SpreadsheetApp.getActive();
  var cleanSheet = ss.getSheetByName(REPORTS_CLEAN_SHEET);
  
  if (!cleanSheet) {
    throw new Error('Sheet "' + REPORTS_CLEAN_SHEET + '" not found.');
  }

  var data = cleanSheet.getDataRange().getValues();
  if (data.length < 2) {
    writeSheet_('Music_Interests_Report', [['No data available']]);
    return;
  }

  var header = data[0];
  var colIdx = getColumnMap_(header);

  var sections = [];

  // Title
  sections.push(['MUSIC & INTERESTS ANALYSIS']);
  sections.push(['']);

  // Music Preferences
  Array.prototype.push.apply(sections, buildMusicPreferences_(data, colIdx));
  sections.push(['']);

  // Top Artists
  Array.prototype.push.apply(sections, buildTopArtists_(data, colIdx, 15));
  sections.push(['']);

  // Top Interests
  Array.prototype.push.apply(sections, buildTopInterests_(data, colIdx));
  sections.push(['']);

  // Music by Age
  Array.prototype.push.apply(sections, buildMusicByDemographic_(data, colIdx, 'Age Range', 'AGE GROUP'));
  sections.push(['']);

  // Music by Gender
  Array.prototype.push.apply(sections, buildMusicByDemographic_(data, colIdx, 'Self-Identified Gender', 'GENDER'));
  sections.push(['']);

  // Song Requests
  Array.prototype.push.apply(sections, buildSongRequestsSummary_(data, colIdx));

  writeSheet_('Music_Interests_Report', sections);
  formatGenericReport_('Music_Interests_Report');
}

function buildMusicPreferences_(data, colIdx) {
  var col = colIdx['Music Preference'];
  if (col === undefined) return [['MUSIC PREFERENCE DISTRIBUTION'], ['Column not found']];

  var counts = {};
  var total = 0;

  for (var i = 1; i < data.length; i++) {
    var value = String(data[i][col] || '').trim();
    if (!value) continue;
    counts[value] = (counts[value] || 0) + 1;
    total++;
  }

  var sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  var rows = [['MUSIC PREFERENCE DISTRIBUTION'], ['Genre', 'Count', 'Percentage']];

  for (var i = 0; i < sorted.length; i++) {
    var genre = sorted[i][0];
    var count = sorted[i][1];
    rows.push([genre, count, ((count / total) * 100).toFixed(1) + '%']);
  });

  rows.push(['TOTAL', total, '100.0%']);
  return rows;
}

function buildTopArtists_(data, colIdx, limit) {
  var col = colIdx['Current Favorite Artist'];
  if (col === undefined) return [['TOP ' + limit + ' FAVORITE ARTISTS'], ['Column not found']];

  var counts = {};
  for (var i = 1; i < data.length; i++) {
    var value = String(data[i][col] || '').trim();
    if (!value) continue;
    counts[value] = (counts[value] || 0) + 1;
  }

  var sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, limit);
  var rows = [['TOP ' + limit + ' FAVORITE ARTISTS'], ['Rank', 'Artist', 'Mentions']];

  for (var idx = 0; idx < sorted.length; idx++) {
    var artist = sorted[idx][0];
    var count = sorted[idx][1];
    rows.push([idx + 1, artist, count]);
  });

  return rows;
}

function buildTopInterests_(data, colIdx) {
  var col = colIdx['Your General Interests (Choose 3)'];
  if (col === undefined) return [['TOP INTERESTS'], ['Column not found']];

  var counts = {};
  var totalResponses = 0;

  for (var i = 1; i < data.length; i++) {
    var value = String(data[i][col] || '').trim();
    if (!value) continue;
    totalResponses++;

    var interestParts = value.split(',');
    var interests = [];
    for (var j = 0; j < interestParts.length; j++) {
      var s = interestParts[j].trim();
      if (s) {
        interests.push(s);
      }
    }
    for (var k = 0; k < interests.length; k++) {
      var interest = interests[k];
      counts[interest] = (counts[interest] || 0) + 1;
    }
  }

  var sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  var rows = [['TOP INTERESTS'], ['Interest', 'Count', '% of Guests']];

  sorted.forEach(([interest, count]) => {
    rows.push([interest, count, ((count / totalResponses) * 100).toFixed(1) + '%']);
  });

  return rows;
}

function buildMusicByDemographic_(data, colIdx, demoCol, demoLabel) {
  var musicCol = colIdx['Music Preference'];
  var demoColIdx = colIdx[demoCol];
  
  if (musicCol === undefined || demoColIdx === undefined) {
    return [['MUSIC PREFERENCE BY ' + demoLabel], ['Required columns not found']];
  }

  var crosstab = {};
  for (var i = 1; i < data.length; i++) {
    var demo = String(data[i][demoColIdx] || '').trim();
    var music = String(data[i][musicCol] || '').trim();
    if (!demo || !music) continue;

    if (!crosstab[demo]) crosstab[demo] = {};
    crosstab[demo][music] = (crosstab[demo][music] || 0) + 1;
  }

  var allGenres = new Set();
  Object.values(crosstab).forEach(musicCounts => {
    Object.keys(musicCounts).forEach(genre => allGenres.add(genre));
  });
  var genres = [];
  allGenres.forEach(function(val) { genres.push(val); });
  genres.sort();
  var demos = Object.keys(crosstab).sort();

  var rows = [['MUSIC PREFERENCE BY ' + demoLabel], [demoCol].concat(genres).concat(['Total'])];

  demos.forEach(demo => {
    var row = [demo];
    var rowTotal = 0;
    genres.forEach(genre => {
      var count = crosstab[demo][genre] || 0;
      row.push(count);
      rowTotal += count;
    });
    row.push(rowTotal);
    rows.push(row);
  });

  return rows;
}

function buildSongRequestsSummary_(data, colIdx) {
  var col = colIdx['Name one song you want to hear at the party.'];
  if (col === undefined) return [['SONG REQUESTS SUMMARY'], ['Column not found']];

  var totalResponses = 0, withSong = 0, withoutSong = 0;

  for (var i = 1; i < data.length; i++) {
    totalResponses++;
    var value = String(data[i][col] || '').trim();
    if (value) withSong++; else withoutSong++;
  }

  var withPct = ((withSong / totalResponses) * 100).toFixed(1);
  var withoutPct = ((withoutSong / totalResponses) * 100).toFixed(1);

  return [
    ['SONG REQUESTS SUMMARY'],
    ['Category', 'Count', 'Percentage'],
    ['Provided Song Request', withSong, withPct + '%'],
    ['No Song Request', withoutSong, withoutPct + '%'],
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
  var ss = SpreadsheetApp.getActive();
  var cleanSheet = ss.getSheetByName(REPORTS_CLEAN_SHEET);
  
  if (!cleanSheet) {
    throw new Error('Sheet "' + REPORTS_CLEAN_SHEET + '" not found.');
  }

  var data = cleanSheet.getDataRange().getValues();
  if (data.length < 2) {
    writeSheet_('Host_Relationships', [['No data available']]);
    return;
  }

  var header = data[0];
  var colIdx = getColumnMap_(header);

  var sections = [];

  sections.push(['HOST RELATIONSHIP ANALYSIS']);
  sections.push(['']);

  // Overview
  Array.prototype.push.apply(sections, buildHostOverview_(data, colIdx));
  sections.push(['']);

  // Duration Distribution
  Array.prototype.push.apply(sections, buildHostDuration_(data, colIdx));
  sections.push(['']);

  // Which Host Known
  Array.prototype.push.apply(sections, buildWhichHost_(data, colIdx));
  sections.push(['']);

  // Closeness Scores
  Array.prototype.push.apply(sections, buildClosenessScores_(data, colIdx));
  sections.push(['']);

  // Closeness by Host
  Array.prototype.push.apply(sections, buildClosenessbyHost_(data, colIdx));

  writeSheet_('Host_Relationships', sections);
  formatGenericReport_('Host_Relationships');
}

function buildHostOverview_(data, colIdx) {
  var col = colIdx['Do you know the Host(s)?'];
  if (col === undefined) return [['OVERVIEW'], ['Column not found']];

  var total = data.length - 1, knowHosts = 0, unknownGuests = 0;

  for (var i = 1; i < data.length; i++) {
    var value = String(data[i][col] || '').trim().toLowerCase();
    if (value.includes('yes') || value.includes('—')) knowHosts++;
    else if (value.includes('no')) unknownGuests++;
  }

  return [
    ['OVERVIEW'],
    ['Metric', 'Count', 'Percentage'],
    ['Total Guests', total, '100.0%'],
    ['Know Host(s)', knowHosts, ((knowHosts / total) * 100).toFixed(1) + '%'],
    ['Unknown Guests', unknownGuests, ((unknownGuests / total) * 100).toFixed(1) + '%']
  ];
}

function buildHostDuration_(data, colIdx) {
  var col = colIdx['Do you know the Host(s)?'];
  if (col === undefined) return [['FRIENDSHIP DURATION'], ['Column not found']];

  var counts = {};
  var total = 0;

  for (var i = 1; i < data.length; i++) {
    var value = String(data[i][col] || '').trim();
    if (!value) continue;
    counts[value] = (counts[value] || 0) + 1;
    total++;
  }

  var order = ['No', 'Yes — 3–12 months', 'Yes — 1–3 years', 'Yes — 3–5 years', 'Yes — 5–10 years', 'Yes — more than 10 years'];
  var sorted = Object.keys(counts).sort((a, b) => {
    var idxA = order.indexOf(a);
    var idxB = order.indexOf(b);
    if (idxA === -1 && idxB === -1) return a.localeCompare(b);
    if (idxA === -1) return 1;
    if (idxB === -1) return -1;
    return idxA - idxB;
  });

  var rows = [['FRIENDSHIP DURATION DISTRIBUTION'], ['Duration', 'Count', 'Percentage']];

  sorted.forEach(duration => {
    var count = counts[duration];
    rows.push([duration, count, ((count / total) * 100).toFixed(1) + '%']);
  });

  rows.push(['TOTAL', total, '100.0%']);
  return rows;
}

function buildWhichHost_(data, colIdx) {
  var col = colIdx['Which host have you known the longest?'];
  if (col === undefined) return [['WHICH HOST KNOWN LONGEST'], ['Column not found']];

  var counts = {};
  var total = 0;

  for (var i = 1; i < data.length; i++) {
    var value = String(data[i][col] || '').trim();
    if (!value) continue;
    counts[value] = (counts[value] || 0) + 1;
    total++;
  }

  var sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  var rows = [['WHICH HOST KNOWN LONGEST'], ['Host', 'Count', 'Percentage']];

  sorted.forEach(([host, count]) => {
    rows.push([host, count, ((count / total) * 100).toFixed(1) + '%']);
  });

  rows.push(['TOTAL', total, '100.0%']);
  return rows;
}

function buildClosenessScores_(data, colIdx) {
  var col = colIdx['If yes, how well do you know them?'];
  if (col === undefined) return [['CLOSENESS SCORE DISTRIBUTION'], ['Column not found']];

  var counts = {};
  var total = 0, sum = 0;

  for (var i = 1; i < data.length; i++) {
    var value = data[i][col];
    if (value === null || value === undefined || value === '') continue;
    var score = Number(value);
    if (!isFinite(score)) continue;
    
    counts[score] = (counts[score] || 0) + 1;
    sum += score;
    total++;
  }

  var avg = total > 0 ? (sum / total).toFixed(2) : 'N/A';
  var sorted = Object.keys(counts).sort((a, b) => Number(b) - Number(a));

  var rows = [
    ['CLOSENESS SCORE DISTRIBUTION (1=Acquaintance, 5=Close Friend)'],
    ['Score', 'Count', 'Percentage']
  ];

  sorted.forEach(score => {
    var count = counts[score];
    rows.push([score, count, ((count / total) * 100).toFixed(1) + '%']);
  });

  rows.push(['TOTAL', total, '100.0%']);
  rows.push(['AVERAGE SCORE', avg, '']);
  return rows;
}

function buildClosenessbyHost_(data, colIdx) {
  var hostCol = colIdx['Which host have you known the longest?'];
  var scoreCol = colIdx['If yes, how well do you know them?'];
  
  if (hostCol === undefined || scoreCol === undefined) {
    return [['AVERAGE CLOSENESS BY HOST'], ['Required columns not found']];
  }

  var hostScores = {};

  for (var i = 1; i < data.length; i++) {
    var host = String(data[i][hostCol] || '').trim();
    var score = data[i][scoreCol];
    
    if (!host || score === null || score === undefined || score === '') continue;
    var scoreNum = Number(score);
    if (!isFinite(scoreNum)) continue;

    if (!hostScores[host]) hostScores[host] = { sum: 0, count: 0 };
    hostScores[host].sum += scoreNum;
    hostScores[host].count++;
  }

  var rows = [['AVERAGE CLOSENESS BY HOST'], ['Host', 'Count', 'Average Score']];

  Object.entries(hostScores).forEach(([host, data]) => {
    var avg = (data.sum / data.count).toFixed(2);
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
  var ss = SpreadsheetApp.getActive();
  var cleanSheet = ss.getSheetByName(REPORTS_CLEAN_SHEET);
  
  if (!cleanSheet) {
    throw new Error('Sheet "' + REPORTS_CLEAN_SHEET + '" not found.');
  }

  var data = cleanSheet.getDataRange().getValues();
  if (data.length < 2) {
    writeSheet_('Attendee_Analysis', [['No data available']]);
    return;
  }

  var header = data[0];
  var colIdx = getColumnMap_(header);

  var sections = [];

  sections.push(['ATTENDEE VS NO-SHOW ANALYSIS']);
  sections.push(['']);

  // Compare demographics
  var compareFields = [
    ['Age Range', 'AGE RANGE'],
    ['Self-Identified Gender', 'GENDER'],
    ['Education Level', 'EDUCATION'],
    ['Do you know the Host(s)?', 'HOST RELATIONSHIP']
  ];

  compareFields.forEach(([colName, label]) => {
    Array.prototype.push.apply(sections, buildAttendeeComparison_(data, colIdx, colName, label));
    sections.push(['']);
  });

  writeSheet_('Attendee_Analysis', sections);
  formatGenericReport_('Attendee_Analysis');
}

function buildAttendeeComparison_(data, colIdx, colName, label) {
  var dataCol = colIdx[colName];
  var checkedInCol = colIdx['Checked-In at Event'];
  
  if (dataCol === undefined || checkedInCol === undefined) {
    return [[label + ' COMPARISON'], ['Required columns not found']];
  }

  var attendeeCounts = {};
  var noShowCounts = {};

  for (var i = 1; i < data.length; i++) {
    var value = String(data[i][dataCol] || '').trim();
    var isAttendee = String(data[i][checkedInCol] || '').trim().toUpperCase() === 'Y';

    if (!value) continue;

    if (isAttendee) {
      attendeeCounts[value] = (attendeeCounts[value] || 0) + 1;
    } else {
      noShowCounts[value] = (noShowCounts[value] || 0) + 1;
    }
  }

  var allValues = new Set(Object.keys(attendeeCounts).concat(Object.keys(noShowCounts)));
  var sorted = [];
  allValues.forEach(function(val) { sorted.push(val); });
  sorted.sort();

  var attendeeTotal = Object.values(attendeeCounts).reduce((a, b) => a + b, 0);
  var noShowTotal = Object.values(noShowCounts).reduce((a, b) => a + b, 0);

  var rows = [
    [label + ' COMPARISON: ATTENDEES VS NO-SHOWS'],
    ['Category', 'Attendees', '% Attendees', 'No-Shows', '% No-Shows']
  ];

  sorted.forEach(value => {
    var attendeeCount = attendeeCounts[value] || 0;
    var noShowCount = noShowCounts[value] || 0;
    var attendeePct = attendeeTotal > 0 ? ((attendeeCount / attendeeTotal) * 100).toFixed(1) : '0.0';
    var noShowPct = noShowTotal > 0 ? ((noShowCount / noShowTotal) * 100).toFixed(1) : '0.0';

    rows.push([value, attendeeCount, attendeePct + '%', noShowCount, noShowPct + '%']);
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
  var ss = SpreadsheetApp.getActive();
  var cleanSheet = ss.getSheetByName(REPORTS_CLEAN_SHEET);
  // simSheet is not used in the function body, removed for clarity:
  // var simSheet = ss.getSheetByName('Guest_Similarity'); 
  
  if (!cleanSheet) {
    throw new Error('Sheet "' + REPORTS_CLEAN_SHEET + '" not found.');
  }

  var data = cleanSheet.getDataRange().getValues();
  if (data.length < 2) {
    writeSheet_('Guest_Profiles', [['No data available']]);
    return;
  }

  var header = data[0];
  var colIdx = getColumnMap_(header);

  // Build profiles
  var profiles = [
    ['Screen Name', 'UID', 'Age', 'Gender', 'Orientation', 'Ethnicity', 'Education', 'Industry', 'Role', 
     'Top 3 Interests', 'Music Genre', 'Favorite Artist', 'Personality Trait', 'Host Known Longest', 
     'Closeness', 'Checked-In']
  ];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Parse interests
    var interestsStr = String(row[colIdx['Your General Interests (Choose 3)']] || '').trim();
    var interests = interestsStr.split(',').slice(0, 3).map(s => s.trim()).join(', ');

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
  
  var sheet = ss.getSheetByName('Guest_Profiles');
  if (sheet) {
    sheet.getRange(1, 1, 1, profiles[0].length)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    for (var col = 1; col <= profiles[0].length; col++) {
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
  var map = {};
  header.forEach((col, idx) => {
    var colName = String(col || '').trim();
    if (colName) map[colName] = idx;
  });
  return map;
}

/**
 * Resolve a column index from a header map in a robust way.
 * Matches by normalized text (lowercase, alphanumeric only) to tolerate punctuation/encoding.
 */
function resolveColumnIndex_(colMap, targetName) {
  if (!colMap || !targetName) return undefined;
  // Exact match first
  if (Object.prototype.hasOwnProperty.call(colMap, targetName)) {
    return colMap[targetName];
  }
  var normalize = s => String(s || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
  var normTarget = normalize(targetName);
  for (var key in colMap) {
    if (!Object.prototype.hasOwnProperty.call(colMap, key)) continue;
    if (normalize(key) === normTarget) return colMap[key];
  }
  return undefined;
}

/**
 * Write data to a sheet
 */
function writeSheet_(sheetName, data) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  sheet.clear();

  if (!data || data.length === 0) {
    sheet.getRange(1, 1).setValue('No data to display');
    return;
  }

  // Find max columns
  var maxCols = Math.max.apply(null, data.map(row => Array.isArray(row) ? row.length : 1));

  // Pad rows to same length
  var paddedRows = data.map(row => {
    if (!Array.isArray(row)) return [row];
    var padded = row.slice();
    while (padded.length < maxCols) padded.push('');
    return padded;
  });

  sheet.getRange(1, 1, paddedRows.length, maxCols).setValues(paddedRows);
}

/**
 * Apply generic formatting to report sheets
 */
function formatGenericReport_(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  // Ensure data has at least one row to safely access data[0].
  if (data.length === 0) return;
  var maxCols = data[0].length;

  for (var col = 1; col <= maxCols; col++) {
    sheet.autoResizeColumn(col);
  }

  for (var row = 1; row <= data.length; row++) {
    var cellValue = String(data[row - 1][0] || '').trim();
    
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

// ============================================================================
// QUICK TEST FUNCTIONS
// ============================================================================
// Add these to the BOTTOM of your Code.gs file
// Run after creating Config.gs

/**
 * Quick test to verify Config.gs constants are loading
 * RUN THIS FIRST after creating Config.gs
 */
function testConstants() {
  Logger.log('════════════════════════════════════════════════════════════════');
  Logger.log('🧪 TESTING CONSTANTS FROM CONFIG.GS');
  Logger.log('════════════════════════════════════════════════════════════════
');
  
  try {
    Logger.log('✅ CORE SHEET NAMES:');
    Logger.log('   SHEET_NAME: ' + SHEET_NAME);
    Logger.log('   WALL_CLEAN_SHEET: ' + WALL_CLEAN_SHEET);
    Logger.log('   DDD_SHEET_NAME: ' + DDD_SHEET_NAME);
    Logger.log('');
    
    Logger.log('✅ FORM RESPONSES (CLEAN) COLUMNS:');
    Logger.log('   UID_COL: ' + UID_COL + ' (Column W)');
    Logger.log('   SCREEN_NAME_COL: ' + SCREEN_NAME_COL + ' (Column V)');
    Logger.log('   ZIP_COL: ' + ZIP_COL + ' (Column F)');
    Logger.log('   GENDER_COL: ' + GENDER_COL + ' (Column H)');
    Logger.log('   BIRTHDAY_COL: ' + BIRTHDAY_COL + ' (Column B)');
    Logger.log('   CHECKED_FLAG_COL: ' + CHECKED_FLAG_COL + ' (Column X)');
    Logger.log('   CHECKED_TS_COL: ' + CHECKED_TS_COL + ' (Column Y)');
    Logger.log('   PHOTO_URL_COL: ' + PHOTO_URL_COL + ' (Column Z)');
    Logger.log('');
    
    Logger.log('✅ DDD SHEET COLUMNS:');
    Logger.log('   DDD_UID_COL: ' + DDD_UID_COL + ' (Column B)');
    Logger.log('   DDD_SCREEN_NAME_COL: ' + DDD_SCREEN_NAME_COL + ' (Column A)');
    Logger.log('   DDD_UNKNOWN_GUEST_COL: ' + DDD_UNKNOWN_GUEST_COL + ' (Column E)');
    Logger.log('');
    
    Logger.log('✅ ANALYSIS CATEGORIES:');
    Logger.log('   Total categories: ' + ANALYSIS_CATEGORIES.length);
    Logger.log('   First 3:');
    ANALYSIS_CATEGORIES.slice(0, 3).forEach(cat => {
      Logger.log('   • ' + cat.name + ' → "' + cat.header + '"');
    });
    Logger.log('');
    
    Logger.log('════════════════════════════════════════════════════════════════');
    Logger.log('✅✅✅ ALL CONSTANTS LOADED SUCCESSFULLY! ✅✅✅');
    Logger.log('════════════════════════════════════════════════════════════════');
    Logger.log('✅ No duplicate declaration errors!');
    Logger.log('✅ Config.gs is working correctly!');
    Logger.log('✅ Ready to run completeSystemCheck()');
    Logger.log('════════════════════════════════════════════════════════════════
');
    
    return true;
    
  } catch (e) {
    Logger.log('════════════════════════════════════════════════════════════════');
    Logger.log('❌❌❌ ERROR LOADING CONSTANTS ❌❌❌');
    Logger.log('════════════════════════════════════════════════════════════════');
    Logger.log('Error message: ' + e.message);
    Logger.log('Error line: ' + e.lineNumber);
    Logger.log('');
    Logger.log('TROUBLESHOOTING STEPS:');
    Logger.log('1. ✓ Make sure Config.gs file exists in file list');
    Logger.log('2. ✓ Make sure Config.gs is saved (💾 icon)');
    Logger.log('3. ✓ Remove ALL "var" lines from Code.gs');
    Logger.log('4. ✓ Remove ALL "var" lines from Utilities.gs');
    Logger.log('5. ✓ Close and reopen Apps Script editor');
    Logger.log('6. ✓ Try running testConstants() again');
    Logger.log('════════════════════════════════════════════════════════════════
');
    
    return false;
  }
}

/**
 * Verify that duplicate constants have been removed
 * Run this to double-check your cleanup
 */
function verifyNoDuplicates() {
  Logger.log('════════════════════════════════════════════════════════════════');
  Logger.log('🔍 VERIFYING NO DUPLICATE CONSTANTS');
  Logger.log('════════════════════════════════════════════════════════════════
');
  
  Logger.log('This test will pass if Config.gs is the ONLY place with constants.
');
  
  Logger.log('MANUAL CHECK:');
  Logger.log('1. Open each .gs file in your project');
  Logger.log('2. Search (Ctrl+F) for: "var SHEET_NAME"');
  Logger.log('3. Should ONLY appear in Config.gs');
  Logger.log('4. If found elsewhere, delete that line');
  Logger.log('');
  Logger.log('FILES TO CHECK:');
  Logger.log('• Config.gs          → ✅ SHOULD have var declarations');
  Logger.log('• Code.gs            → ❌ SHOULD NOT have var declarations');
  Logger.log('• Utilities.gs       → ❌ SHOULD NOT have var declarations');
  Logger.log('• Reports.gs         → ❌ SHOULD NOT have var declarations');
  Logger.log('• DataClean.gs       → ❌ SHOULD NOT have var declarations');
  Logger.log('• Any other .gs file → ❌ SHOULD NOT have var declarations');
  Logger.log('');
  Logger.log('After cleanup, run testConstants() to verify!
');
  Logger.log('════════════════════════════════════════════════════════════════
');
}

/**
 * Show which constants you should remove from other files
 */
function showConstantsToRemove() {
  Logger.log('════════════════════════════════════════════════════════════════');
  Logger.log('🧹 CONSTANTS TO REMOVE FROM OTHER FILES');
  Logger.log('════════════════════════════════════════════════════════════════
');
  
  Logger.log('SEARCH for these in Code.gs, Utilities.gs, and other files:');
  Logger.log('Then DELETE any lines you find (except in Config.gs)
');
  
  var constantsToRemove = [
    'var SHEET_NAME',
    'var WALL_CLEAN_SHEET',
    'var REPORTS_CLEAN_SHEET',
    'var DDD_SHEET_NAME',
    'var PHOTOS_FOLDER_NAME',
    'var WALL_TARGET_ADDRESS',
    'var ZIP_COL',
    'var GENDER_COL',
    'var BIRTHDAY_COL',
    'var SCREEN_NAME_COL',
    'var UID_COL',
    'var CHECKED_FLAG_COL',
    'var CHECKED_TS_COL',
    'var PHOTO_URL_COL',
    'var DDD_UID_COL',
    'var ANALYSIS_CATEGORIES',
    'var CYCLE_INTERVAL',
    'var COLOR_PALETTE'
  ];
  
  Logger.log('DELETE THESE LINES from Code.gs and Utilities.gs:');
  Logger.log('─────────────────────────────────────────────────────────────');
  constantsToRemove.forEach((line, i) => {
    Logger.log((i + 1) + '. ' + line + ' = ...');
  });
  Logger.log('─────────────────────────────────────────────────────────────
');
  
  Logger.log('HOW TO SEARCH:');
  Logger.log('1. Open Code.gs');
  Logger.log('2. Press Ctrl+F (or Cmd+F on Mac)');
  Logger.log('3. Type: var SHEET_NAME');
  Logger.log('4. If found, DELETE the entire line');
  Logger.log('5. Repeat for each constant above');
  Logger.log('6. Do the same for Utilities.gs and other files');
  Logger.log('7. Save all files');
  Logger.log('8. Run testConstants() to verify
');
  
  Logger.log('════════════════════════════════════════════════════════════════
');
}

function checkConfig() {
  try {
    Logger.log('Testing SHEET_NAME...');
    Logger.log(SHEET_NAME);
    Logger.log('✅ Config loaded!');
  } catch (e) {
    Logger.log('❌ Config NOT loaded: ' + e.message);
  }
}

/**
 * Convert Google Drive view link to direct image link
 */
function convertDriveUrlToDirectLink(driveUrl) {
  if (!driveUrl || driveUrl === '') return '';
  
  try {
    // Extract file ID from Drive URL
    var match = driveUrl.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
    if (match) {
      var fileId = match[1];
      return 'https://drive.google.com/uc?export=view&id=' + fileId;
    }
    
    // If already a direct link, return as-is
    if (driveUrl.includes('drive.google.com/uc?')) {
      return driveUrl;
    }
    
    return driveUrl;
    
  } catch (error) {
    Logger.log('Error converting Drive URL: ' + error.toString());
    return driveUrl;
  }
}

/**
 * Sync all photos from Drive view links to direct image links
 */
function syncPhotosFromDrive() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cleanSheet = ss.getSheetByName('Form Responses (Clean)');
    
    if (!cleanSheet) {
      return { success: false, message: 'Form Responses (Clean) sheet not found' };
    }
    
    var data = cleanSheet.getDataRange().getValues();
    var photoUrlCol = 25; // Column Z (0-indexed = 25)
    
    var updatedCount = 0;
    
    for (var i = 1; i < data.length; i++) {
      var currentUrl = data[i][photoUrlCol];
      
      if (currentUrl && typeof currentUrl === 'string' && currentUrl.includes('drive.google.com')) {
        if (currentUrl.includes('/view') || currentUrl.includes('?usp=')) {
          var directUrl = convertDriveUrlToDirectLink(currentUrl);
          
          if (directUrl !== currentUrl) {
            cleanSheet.getRange(i + 1, photoUrlCol + 1).setValue(directUrl);
            updatedCount++;
          }
        }
      }
    }
    
    Logger.log('✅ Photo sync complete: ' + updatedCount + ' URLs converted');
    
    return {
      success: true,
      updated: updatedCount,
      message: 'Successfully converted ' + updatedCount + ' photo URLs'
    };
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Manual function to sync all photos - RUN THIS!
 */
function manualSyncAllPhotos() {
  var result = syncPhotosFromDrive();
  
  SpreadsheetApp.getUi().alert(
    '📸 Photo Sync Results

' +
    (result.success ?
      '✅ Success!\n\nConverted ' + result.updated + ' URLs to direct image links.\n\nRefresh MM.html to see photos.' :
      '❌ Failed\n\n' + result.message)
  );
}

/**
 * Helper function to extract the Google Drive File ID from the raw 'Photo URL'.
 */
function extractFileId(url) {
  if (!url || typeof url !== 'string') return '';
  // The 'uc?export=view' format: extract the ID after 'id='
  var match = url.match(/id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : '';
}

/**
 * Helper function to extract the Google Drive File ID from the raw 'Photo URL'.
 */
function extractFileId(url) {
  if (!url || typeof url !== 'string') return '';
  // The 'uc?export=view' format: extract the ID after 'id='
  var match = url.match(/id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : '';
}

/**
 * [CRUCIAL STEP] This function acts as the secure image server.
 * It is triggered when the deployed Web App URL is accessed with the '?fileId=' parameter.
 * It fetches the file securely using DriveApp and returns the raw image content.
 */



function getIntroText() {
  // EDIT THIS TEXT TO CHANGE THE ROLLING MESSAGE
  return "INPUT TEXT HERE";
}

function getZipCodesFromSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Form Responses 1');
    
    if (!sheet) {
      throw new Error('Sheet "Form Responses 1" not found');
    }
    
    var lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return {};
    }
    
    var zipRange = sheet.getRange(2, 5, lastRow - 1, 1);
    var zipValues = zipRange.getValues();
    
    // Count occurrences of each zip code
    var zipCounts = {};
    zipValues.forEach(row => {
      var zip = String(row[0]).trim();
      if (zip && zip !== '' && zip.length === 5) {
        zipCounts[zip] = (zipCounts[zip] || 0) + 1;
      }
    });
    
    return zipCounts;
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return {};
  }
}

function getAddressCoordinates(address) {
  try {
    var geocoder = Maps.newGeocoder();
    var location = geocoder.geocode(address);
    
    if (location.status === 'OK' && location.results.length > 0) {
      var result = location.results[0];
      return {
        lat: result.geometry.location.lat,
        lng: result.geometry.location.lng
      };
    }
    return null;
  } catch (error) {
    Logger.log('Error geocoding address: ' + error.toString());
    return null;
  }
}

function getZipCodeCoordinates(zipCode) {
  try {
    var geocoder = Maps.newGeocoder();
    var location = geocoder.geocode(zipCode + ', USA');
    
    if (location.status === 'OK' && location.results.length > 0) {
      var result = location.results[0];
      return {
        lat: result.geometry.location.lat,
        lng: result.geometry.location.lng,
        zip: zipCode
      };
    }
    return null;
  } catch (error) {
    Logger.log('Error geocoding ' + zipCode + ': ' + error.toString());
    return null;
  }
}

function getAllZipData() {
  var zipCounts = getZipCodesFromSheet();
  
  if (Object.keys(zipCounts).length === 0) {
    return { error: 'No zip codes found in Column E' };
  }
  
  var targetCount = zipCounts['64110'] || 0;
  
  // Geocode the specific address
  var targetCoords = getAddressCoordinates('5317 Charlotte St, Kansas City, MO 64110');
  if (!targetCoords) {
    return { error: 'Could not geocode target address' };
  }
  
  var target = {
    lat: targetCoords.lat,
    lng: targetCoords.lng,
    zip: '5317 Charlotte',
    count: targetCount,
    displayName: '5317 Charlotte'
  };
  
  var allZips = [];
  var totalRespondents = 0;
  
  Object.keys(zipCounts).forEach(zip => {
    totalRespondents += zipCounts[zip];
    
    if (zip !== '64110') {
      var coord = getZipCodeCoordinates(zip);
      if (coord) {
        coord.count = zipCounts[zip];
        allZips.push(coord);
      }
      Utilities.sleep(100);
    }
  });
  
  return {
    target: target,
    zips: allZips,
    totalCount: Object.keys(zipCounts).length,
    totalRespondents: totalRespondents
  };
}

/**
 * PHOTO LOADING FIX FOR PARTY DISPLAY SYSTEM
 * Add these functions to your Code.gs file
 */

/**
 * Get properly formatted photo data for display
 * Returns array of guest objects with working photo URLs
 */
function getDisplayPhotos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cleanSheet = ss.getSheetByName('Form Responses (Clean)');
    
    if (!cleanSheet) {
      Logger.log('ERROR: Form Responses (Clean) sheet not found');
      return [];
    }
    
    var data = cleanSheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    
    var headers = data[0];
    var colMap = {};
    headers.forEach((h, i) => colMap[String(h).trim()] = i);
    
    var screenNameCol = colMap['Screen Name'];
    var checkedInCol = colMap['Checked-In at Event'];
    var photoCol = colMap['Photo URL'];
    
    var guests = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Only include checked-in guests
      var checkedIn = String(row[checkedInCol] || '').trim().toUpperCase();
      if (checkedIn !== 'Y') continue;
      
      var screenName = String(row[screenNameCol] || '').trim();
      if (!screenName) continue;
      
      var rawPhotoUrl = String(row[photoCol] || '').trim();
      var photoUrl = processPhotoUrl(rawPhotoUrl);
      
      guests.push({
        screenName: screenName,
        photoUrl: photoUrl,
        hasPhoto: photoUrl !== ''
      });
    }
    
    Logger.log('Found ' + guests.length + ' checked-in guests');
    return guests;
    
  } catch (error) {
    Logger.log('Error in getDisplayPhotos: ' + error.toString());
    return [];
  }
}

/**
 * Process and validate photo URL
 * Converts various Drive URL formats to working image URLs
 */
function processPhotoUrl(rawUrl) {
  if (!rawUrl || rawUrl === '' || rawUrl === 'Photo URL') {
    return '';
  }
  try {
    if (rawUrl.includes('drive.google.com')) {
      var fileId = extractDriveFileId(rawUrl);
      if (fileId) {
        return 'https://drive.google.com/uc?export=view&id=' + fileId;
      }
    }
    if (rawUrl.startsWith('http') && !rawUrl.includes('drive.google.com')) {
      return rawUrl;
    }
    return '';
  } catch (error) {
    Logger.log('Error processing photo URL: ' + error.toString());
    return '';
  }
}


/**
 * Extract Google Drive file ID from various URL formats
 */
function extractDriveFileId(url) {
  if (!url) return null;
  
  // Format 1: /file/d/FILE_ID/
  var match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  
  // Format 2: ?id=FILE_ID
  match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  
  // Format 3: /open?id=FILE_ID
  match = url.match(/\/open\?id=([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  
  return null;
}

/**
 * Clean up all photo URLs in the sheet
 * Run this once to fix existing URLs
 */
function cleanupPhotoUrls() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cleanSheet = ss.getSheetByName('Form Responses (Clean)');
    
    if (!cleanSheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    var data = cleanSheet.getDataRange().getValues();
    var headers = data[0];
    var photoCol = headers.indexOf('Photo URL') + 1; // Convert to 1-indexed
    
    if (photoCol === 0) {
      return { success: false, message: 'Photo URL column not found' };
    }
    
    var fixedCount = 0;
    var clearedCount = 0;
    
    for (var i = 2; i <= data.length; i++) {
      var rawUrl = String(data[i - 1][photoCol - 1] || '').trim();
      
      // Clear invalid placeholder text
      if (rawUrl === 'Photo URL' || rawUrl === '') {
        cleanSheet.getRange(i, photoCol).setValue('');
        clearedCount++;
        continue;
      }
      
      // Convert Drive URLs to thumbnail format
      if (rawUrl.includes('drive.google.com')) {
        var fileId = extractDriveFileId(rawUrl);
        if (fileId) {
          var thumbnailUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
          cleanSheet.getRange(i, photoCol).setValue(thumbnailUrl);
          fixedCount++;
        }
      }
    }
    
    SpreadsheetApp.flush();
    
    var message = 'Photo cleanup complete!
' +
                   'Fixed: ' + fixedCount + ' URLs
' +
                   'Cleared: ' + clearedCount + ' invalid entries';
    
    Logger.log(message);
    
    return {
      success: true,
      fixed: fixedCount,
      cleared: clearedCount,
      message: message
    };
    
  } catch (error) {
    Logger.log('Error cleaning photo URLs: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Test function - run this to verify photo loading
 */
function testPhotoLoading() {
  Logger.log('=== TESTING PHOTO LOADING ===
');
  
  var guests = getDisplayPhotos();
  
  Logger.log('Total checked-in guests: ' + guests.length);
  Logger.log('Guests with photos: ' + guests.filter(g => g.hasPhoto).length);
  Logger.log('Guests without photos: ' + guests.filter(g => !g.hasPhoto).length);
  Logger.log('');
  
  if (guests.length > 0) {
    Logger.log('Sample guests:');
    guests.slice(0, 5).forEach(guest => {
      Logger.log('  • ' + guest.screenName);
      Logger.log('    Photo: ' + (guest.hasPhoto ? '✅ ' + guest.photoUrl : '❌ No photo'));
    });
  }
  
  Logger.log('
=== END TEST ===');
}

/**
 * Manual cleanup function - RUN THIS to fix all URLs
 */
function manualPhotoCleanup() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    '📸 Photo URL Cleanup',
    'This will:
' +
    '• Convert Drive URLs to thumbnail format
' +
    '• Clear invalid "Photo URL" text entries
' +
    '• Fix image loading issues

' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  var result = cleanupPhotoUrls();
  
  ui.alert(
    result.success ? '✅ Success!' : '❌ Error',
    result.message,
    ui.ButtonSet.OK
  );
}






// ===== File: Router =====

/**
 * ============================================================================
 * ROUTER.GS - Web App Entry Point and Routing
 * ============================================================================
 */

/**
 * Main entry point for web app requests
 * @param {Object} e - Event object with query parameters
 * @return {HtmlOutput} Rendered HTML page
 */
function doGet(e) {
  try {
    // Handle case where e is undefined (when called from script editor)
    if (!e || !e.parameter) {
      e = { parameter: {} };
    }
    
    // Get page parameter from URL (?page=display)
    var page = (e.parameter.page || 'display').toLowerCase();
    
    Logger.log('doGet called with page: ' + page);
    
    // Route to appropriate page
    switch(page) {
      // === MAIN DISPLAY SYSTEM ===
      case 'display':
        return serveDisplayPage();
      
      // === STANDALONE CHECK-IN ===
      case 'checkin':
        return serveCheckInPage();
      
      // === INDIVIDUAL ANALYTICS PAGES ===
      case 'intro':
        return serveIntroPage();
      
      case 'wall':
        return serveWallPage();
      
      case 'mm':
      case 'matchmaker':
        return serveMatchmakerPage();
      
      case 'msa':
        return serveMSAPage();
      
      case 'network':
        return serveNetworkPage();
      
      case 'map':
        return serveMapPage();
      
      // === DEFAULT ===
      default:
        Logger.log('Unknown page requested: ' + page + ', defaulting to display');
        return serveDisplayPage();
    }
    
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return createErrorPage(error);
  }
}

/**
 * Serve the main display rotation system
 * @return {HtmlOutput} Display orchestrator page
 */
function serveDisplayPage() {
  try {
    var html = HtmlService.createTemplateFromFile('Display')
      .evaluate()
      .setTitle('Panopticon Display System')
      .setFaviconUrl('https://ssl.gstatic.com/docs/script/images/favicon.ico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    return html;
  } catch (error) {
    Logger.log('Error serving display page: ' + error.toString());
    throw error;
  }
}

/**
 * Serve the check-in interface
 * @return {HtmlOutput} Check-in page
 */
function serveCheckInPage() {
  try {
    var html = HtmlService.createTemplateFromFile('CheckInInterface')
      .evaluate()
      .setTitle('Guest Check-In')
      .setFaviconUrl('https://ssl.gstatic.com/docs/script/images/favicon.ico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    return html;
  } catch (error) {
    Logger.log('Error serving check-in page: ' + error.toString());
    throw error;
  }
}

/**
 * Serve the intro/welcome page
 * @return {HtmlOutput} Intro page
 */
function serveIntroPage() {
  try {
    var html = HtmlService.createTemplateFromFile('intro')
      .evaluate()
      .setTitle('Welcome to Panopticon')
      .setFaviconUrl('https://ssl.gstatic.com/docs/script/images/favicon.ico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    return html;
  } catch (error) {
    Logger.log('Error serving intro page: ' + error.toString());
    return createPlaceholderPage('intro', 'Welcome to the Panopticon');
  }
}

/**
 * Serve the network visualization wall
 * @return {HtmlOutput} Wall HTML page
 */
function serveWallPage() {
  try {
    var html = HtmlService.createTemplateFromFile('wall')
      .evaluate()
      .setTitle('Network Visualization Wall')
      .setFaviconUrl('https://ssl.gstatic.com/docs/script/images/favicon.ico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    return html;
  } catch (error) {
    Logger.log('Error serving wall page: ' + error.toString());
    return createPlaceholderPage('wall', 'Network Visualization Wall');
  }
}

/**
 * Serve the matchmaker page
 * @return {HtmlOutput} Matchmaker page
 */
function serveMatchmakerPage() {
  try {
    var html = HtmlService.createTemplateFromFile('mm')  // ← Changed to lowercase 'mm'
      .evaluate()
      .setTitle('Guest Matchmaker')
      .setFaviconUrl('https://ssl.gstatic.com/docs/script/images/favicon.ico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    return html;
  } catch (error) {
    Logger.log('Error serving matchmaker page: ' + error.toString());
    return createPlaceholderPage('mm', 'Guest Similarity Matchmaker');
  }
}

/**
 * Serve the MSA analysis page
 * @return {HtmlOutput} MSA page
 */
function serveMSAPage() {
  try {
    var html = HtmlService.createTemplateFromFile('msa')
      .evaluate()
      .setTitle('MSA Analysis')
      .setFaviconUrl('https://ssl.gstatic.com/docs/script/images/favicon.ico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    return html;
  } catch (error) {
    Logger.log('Error serving MSA page: ' + error.toString());
    return createPlaceholderPage('msa', 'MSA Analysis');
  }
}

/**
 * Serve the network graph page
 * @return {HtmlOutput} Network page
 */
function serveNetworkPage() {
  try {
    var html = HtmlService.createTemplateFromFile('network')
      .evaluate()
      .setTitle('Network Graph')
      .setFaviconUrl('https://ssl.gstatic.com/docs/script/images/favicon.ico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    return html;
  } catch (error) {
    Logger.log('Error serving network page: ' + error.toString());
    return createPlaceholderPage('network', 'Network Graph Visualization');
  }
}

/**
 * Serve the zip code map page
 * @return {HtmlOutput} Map page
 */
function serveMapPage() {
  try {
    var html = HtmlService.createTemplateFromFile('map')
      .evaluate()
      .setTitle('Geographic Distribution')
      .setFaviconUrl('https://ssl.gstatic.com/docs/script/images/favicon.ico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    return html;
  } catch (error) {
    Logger.log('Error serving map page: ' + error.toString());
    return createPlaceholderPage('map', 'Geographic Distribution Map');
  }
}

/**
 * Include helper - allows including CSS/JS files in HTML
 * Usage in HTML: <?!= include('styles') ?>
 * @param {string} filename - Name of file to include (without .html extension)
 * @return {string} File contents
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Create a placeholder page when the actual HTML file doesn't exist yet
 * @param {string} pageName - Name of the page
 * @param {string} title - Page title
 * @return {HtmlOutput} Placeholder page
 */
function createPlaceholderPage(pageName, title) {
  var html = '    <!DOCTYPE html>\n' +
    '    <html>\n' +
    '    <head>\n' +
    '      <meta charset="UTF-8">\n' +
    '      <meta name="viewport" content="width=device-width, initial-scale=1.0">\n' +
    '      <title>' + title + '</title>\n' +
    '      <style>\n' +
    '        * {\n' +
    '          margin: 0;\n' +
    '          padding: 0;\n' +
    '          box-sizing: border-box;\n' +
    '        }\n' +
    '        body {\n' +
    '          font-family: \'Arial\', sans-serif;\n' +
    '          background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);\n' +
    '          color: white;\n' +
    '          display: flex;\n' +
    '          flex-direction: column;\n' +
    '          justify-content: center;\n' +
    '          align-items: center;\n' +
    '          height: 100vh;\n' +
    '          text-align: center;\n' +
    '          padding: 20px;\n' +
    '        }\n' +
    '        h1 {\n' +
    '          font-size: 48px;\n' +
    '          margin-bottom: 20px;\n' +
    '          text-shadow: 0 2px 10px rgba(0,0,0,0.3);\n' +
    '        }\n' +
    '        p {\n' +
    '          font-size: 20px;\n' +
    '          opacity: 0.9;\n' +
    '          max-width: 600px;\n' +
    '          line-height: 1.6;\n' +
    '        }\n' +
    '        .icon {\n' +
    '          font-size: 80px;\n' +
    '          margin-bottom: 30px;\n' +
    '          opacity: 0.7;\n' +
    '        }\n' +
    '        .status {\n' +
    '          margin-top: 40px;\n' +
    '          padding: 15px 30px;\n' +
    '          background: rgba(255, 255, 255, 0.1);\n' +
    '          border-radius: 10px;\n' +
    '          backdrop-filter: blur(10px);\n' +
    '        }\n' +
    '      </style>\n' +
    '    </head>\n' +
    '    <body>\n' +
    '      <div class="icon">🚧</div>\n' +
    '      <h1>' + title + '</h1>\n' +
    '      <p>This page is currently under construction.</p>\n' +
    '      <div class="status">\n' +
    '        <strong>Page:</strong> ' + pageName + '.html<br>\n' +
    '        <strong>Status:</strong> Coming Soon\n' +
    '      </div>\n' +
    '    </body>\n' +
    '    </html>\n';

  return HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Create an error page
 * @param {Error} error - The error object
 * @return {HtmlOutput} Error page
 */
function createErrorPage(error) {
  var html = '    <!DOCTYPE html>\n' +
    '    <html>\n' +
    '    <head>\n' +
    '      <meta charset="UTF-8">\n' +
    '      <title>Error</title>\n' +
    '      <style>\n' +
    '        body {\n' +
    '          font-family: Arial, sans-serif;\n' +
    '          background: #1a1a1a;\n' +
    '          color: #ff4444;\n' +
    '          padding: 40px;\n' +
    '          text-align: center;\n' +
    '        }\n' +
    '        h1 { margin-bottom: 20px; }\n' +
    '        pre {\n' +
    '          background: #2a2a2a;\n' +
    '          padding: 20px;\n' +
    '          border-radius: 5px;\n' +
    '          text-align: left;\n' +
    '          overflow: auto;\n' +
    '          color: #ffaa00;\n' +
    '        }\n' +
    '        a {\n' +
    '          display: inline-block;\n' +
    '          margin-top: 20px;\n' +
    '          padding: 10px 20px;\n' +
    '          background: #333;\n' +
    '          color: white;\n' +
    '          text-decoration: none;\n' +
    '          border-radius: 5px;\n' +
    '        }\n' +
    '        a:hover { background: #444; }\n' +
    '      </style>\n' +
    '    </head>\n' +
    '    <body>\n' +
    '      <h1>⚠️ Error Loading Page</h1>\n' +
    '      <pre>' + error.toString() + '</pre>\n' +
    '      <a href="?page=display">Go to Display System</a>\n' +
    '      <a href="?page=checkin">Go to Check-In</a>\n' +
    '    </body>\n' +
    '    </html>\n';

  return HtmlService.createHtmlOutput(html);
}

/**
 * Test function - opens display system in new tab from script editor
 * Run this to quickly test your deployment
 */
function testOpenDisplay() {
  var url = ScriptApp.getService().getUrl();
  Logger.log('Web App URL: ' + url);
  Logger.log('Display URL: ' + url + '?page=display');
  Logger.log('Check-in URL: ' + url + '?page=checkin');
  
  SpreadsheetApp.getUi().alert(
    '🚀 Web App URLs:

' +
    'Display System:
' + url + '?page=display

' +
    'Check-In Portal:
' + url + '?page=checkin

' +
    'Copy these URLs for your event!'
  );
}

/**
 * Test function - verify all pages load
 */
function testAllPages() {
  var pages = ['display', 'checkin', 'intro', 'wall', 'mm', 'msa', 'network', 'map'];
  var results = [];
  
  pages.forEach(page => {
    try {
      var mockEvent = { parameter: { page: page } };
      var result = doGet(mockEvent);
      results.push('✅ ' + page + ': OK');
    } catch (error) {
      results.push('❌ ' + page + ': ' + error.toString());
    }
  });
  
  Logger.log('Page Test Results:
' + results.join('
'));
  SpreadsheetApp.getUi().alert('Page Test Results:

' + results.join('
'));
}
