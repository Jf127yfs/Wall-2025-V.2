// ============================================
// MENU SYSTEM
// ============================================

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