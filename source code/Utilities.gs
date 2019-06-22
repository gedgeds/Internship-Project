//----------------------------------------------------------------------------------------//
//                                      DATABASES                                         //
//----------------------------------------------------------------------------------------//

function createDatabase(activeSpreadsheet, activeSheet) {
  
  var databaseSheetName = activeSheet.getName() + ' database';
  var databaseSheet = activeSpreadsheet.insertSheet(databaseSheetName);
  
  copyDatabase(databaseSheetName);
  activeSpreadsheet.setActiveSheet(activeSheet);
  
}


function updateDatabase() {
  
  // check if function is called from media plan sheet
  var currentSheetName = SpreadsheetApp.getActiveSheet().getName();
  if(currentSheetName.indexOf(getMediaPlanSheetName()) < 0){
    showInvalidDatabaseUpdateAlert();
    return;
  }
  
  var spreadsheet = SpreadsheetApp.getActive();
  var currentSheet = spreadsheet.getSheetByName(getMediaPlanSheetName());
  var mediaPlanDatabaseSheet = spreadsheet.getSheetByName(getMediaPlanSheetName() + ' database');
  
  if (!mediaPlanDatabaseSheet) {
    createDatabase(spreadsheet, currentSheet);
  }
  else if (mediaPlanDatabaseSheet) {
    mediaPlanDatabaseSheet.clear();
    copyDatabase(getMediaPlanSheetName() + ' database');
  }
}


function copyDatabase(localDatabaseName) {
  
  var externalSpreadsheet = SpreadsheetApp.openById(getDatabaseSpreadsheetID());
  var externalSheet = externalSpreadsheet.getSheetByName(getDatabaseSheetName());
  var currentSpreadsheet = SpreadsheetApp.getActive();
  var localDatabaseSheet = currentSpreadsheet.getSheetByName(localDatabaseName);
  
  var rowNumber = externalSheet.getLastRow();
  var colNumber = externalSheet.getLastColumn();
  var externalData = externalSheet.getRange(1, 1, rowNumber, colNumber).getValues();
  
  localDatabaseSheet.clearContents();
  localDatabaseSheet.getRange(1, 1, rowNumber, colNumber).setValues(externalData);  
  localDatabaseSheet.hideSheet();
  
}


//----------------------------------------------------------------------------------------//
//                                      RECORDS                                           //
//----------------------------------------------------------------------------------------//

function generateRecords() {
   
  // check if function is called from media plan sheet
  var currentSheetName = SpreadsheetApp.getActiveSheet().getName();
  if (currentSheetName.indexOf(getMediaPlanSheetName()) < 0) {return;}
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Kurti áraðus', 'Áraðø skaièius:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var records = Number(response.getResponseText());
    PropertiesService.getScriptProperties().setProperty('amountOfRecords', records);
    if (records > 0) {
      createDropdownLists(records);
    }
  } 
  else if (response.getSelectedButton() == ui.Button.CANCEL) {
    // if cancel
  } 
  else {
    // if close dialog
  }
}


function createDropdownLists(records) {
  
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = currentSpreadsheet.getSheetByName(getMediaPlanSheetName());
  var databaseSheet = currentSpreadsheet.getSheetByName(currentSheet.getName() + ' database');
  var constants = getMediaPlanConstantsEnum();
  
  //--------- POPULATING DROPDOWN LISTS --------//
  
  var mediaPlanRangeChannel = currentSheet.getRange('A14:A'+ (14 + records - 1).toString() );
  var databaseRangeChannel = databaseSheet.getRange('A2:A'); 
  var ruleChannel = SpreadsheetApp.newDataValidation().requireValueInRange(databaseRangeChannel).build();
  
  mediaPlanRangeChannel.setDataValidation(ruleChannel);
    
  if (Number(PropertiesService.getScriptProperties().getProperty('amountOfDays')) > 0) {
    generateCalendarFields(currentSheet, records, constants);
  }
  setRecordValuesFormulas(records);
  
}


function setRecordValuesFormulas(records) {
  
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = currentSpreadsheet.getSheetByName(getMediaPlanSheetName());
  var style = getStyleConstantsEnum();
  
  var rowPosition = 14;  
  var colQuantity = 9;
  var colTotal = 11;
  var colUnitGrossPrice = 12;
  var colUnitNettoPrice = 13;
  var colGrossPrice = 14;
  var colDiscount = 15;
  var colNettoPrice = 16;
  var colIndications = 17;
  var colCTR = 18;
  var colRedirects = 19;
  
  var wholeRecordsRange = currentSheet.getRange(rowPosition, 1, records, 19);
  var num = Number(PropertiesService.getScriptProperties().getProperty('amountOfDays')) + 25;
  
  for (var i=1; i<=records; i++ ) {
    
    currentSheet.getRange(rowPosition, colQuantity).setFormula('=sum(T' + rowPosition.toString() + ':' + colName(num) + rowPosition.toString() + ')');
    currentSheet.getRange(rowPosition, colTotal).setFormula('=H' + rowPosition.toString() + '*I' + rowPosition.toString());
    currentSheet.getRange(rowPosition, colUnitNettoPrice).setFormula('=+L' + rowPosition.toString() + '-L' + rowPosition.toString() + '*O' + rowPosition.toString()).setNumberFormat(getPriceFormat());
    currentSheet.getRange(rowPosition, colGrossPrice).setFormula('=L' + rowPosition.toString() + '*I' + rowPosition.toString()).setNumberFormat(getPriceFormat());
    currentSheet.getRange(rowPosition, colNettoPrice).setFormula('=M' + rowPosition.toString() + '*I' + rowPosition.toString()).setNumberFormat(getPriceFormat());
    currentSheet.getRange(rowPosition, colIndications).setFormula('=H' + rowPosition.toString() + '*I' + rowPosition.toString());
    currentSheet.getRange(rowPosition, colRedirects).setFormula('=Q' + rowPosition.toString() + '*R' + rowPosition.toString());
    
    currentSheet.getRange(rowPosition, colUnitGrossPrice).setNumberFormat(getPriceFormat());
    currentSheet.getRange(rowPosition, colDiscount).setNumberFormat(getPercentFormat());
    currentSheet.getRange(rowPosition, colCTR).setNumberFormat(getPercentFormat());
        
    rowPosition++;  
    
  }
        
  //----------- SETTING CELLS STYLES -----------//
  
  // Conditional formatting
  if (Number(PropertiesService.getScriptProperties().getProperty('amountOfDays')) > 0) {
    setCalendarConditionalFormatting(currentSheet);
  }  
  // Allignment
  wholeRecordsRange.setHorizontalAlignment('center');
  wholeRecordsRange.setVerticalAlignment('middle');
    
  // Cell borders
  wholeRecordsRange.setBorder(true, true, true, true, true, true);
  
  // Font family
  currentSheet.getRange(1, 1, currentSheet.getLastRow(), currentSheet.getLastColumn()).setFontFamily(style.GLOBAL_FONT_FAMILY);
    
}


//----------------------------------------------------------------------------------------//
//                                      CALENDAR                                          //
//----------------------------------------------------------------------------------------//

function generateCalendar() {
  
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var currentSheetName = currentSheet.getName();
  var style = getStyleConstantsEnum();
  
  var isCalledFromMedia = currentSheetName.indexOf(getMediaPlanSheetName());
  var isCalledFromOrder = currentSheetName.indexOf('Order');
  
  // check if function is called from media plan sheet or order sheet
  if (isCalledFromMedia >= 0) {
    var constants = getMediaPlanConstantsEnum();
    var records = Number(PropertiesService.getScriptProperties().getProperty('amountOfRecords'));
  }
  else if (isCalledFromOrder >= 0) {
    var constants = getOrderConstantsEnum();
    var records = Number(PropertiesService.getScriptProperties().getProperty('amountOfOrderRecords')) - 1;
  }
  else {return;}
  
  var calendarLang = PropertiesService.getScriptProperties().getProperty('calendarLanguage');
  
  clearCalendar(currentSheet);
   
  if (calendarLang === 'LT') {
    var monthsEnum = getMonthsEnumLT();
    var weekDaysEnum = getWeekDaysEnumLT();
  } 
  else if (calendarLang === 'EN') {
    var monthsEnum = getMonthsEnumEN();
    var weekDaysEnum = getWeekDaysEnumEN();
  }
  
  var datePeriodValue = currentSheet.getRange(5, 2).getValue();  
  var periodStartDate = datePeriodValue.substring(0, 10);
  var periodEndDate = datePeriodValue.substring(11);
  
  periodStartDate = periodStartDate.replace(/\./g, "-");
  periodEndDate = periodEndDate.replace(/\./g, "-");
  
  var periodStartMs = new Date(periodStartDate).valueOf();
  var periodEndMs = new Date(periodEndDate).valueOf();
  
  if (periodEndMs <= periodStartMs) {
    showWrongDateAlert();
    return;
  }
  
  var msInDay = 24 * 60 * 60 * 1000;
  var diffInMs = periodEndMs - periodStartMs;
  
  if (diffInMs >= 100 * 24 * 60 * 60 * 1000) {
    if (showOverflowDateAlert()) {
      return;
    }
  }
  
  var daysRange = Math.floor(diffInMs/msInDay);
  PropertiesService.getScriptProperties().setProperty('amountOfDays', daysRange);
  
  var periodMonthRange   = currentSheet.getRange( constants.CALENDAR_MONTH_ROW_POSITION,    constants.CALENDAR_START_COL_POSITION, 1, daysRange + 7);
  var periodWeekRange    = currentSheet.getRange( constants.CALENDAR_WEEK_ROW_POSITION,     constants.CALENDAR_START_COL_POSITION, 1, daysRange + 7);
  var periodWeekdayRange = currentSheet.getRange( constants.CALENDAR_WEEK_DAY_ROW_POSITION, constants.CALENDAR_START_COL_POSITION, 1, daysRange + 7);
  var periodDayRange     = currentSheet.getRange( constants.CALENDAR_DAY_ROW_POSITION,      constants.CALENDAR_START_COL_POSITION, 1, daysRange + 7);
    
  var date = new Date(periodStartDate);
  
  
  //----------- SETTING CELLS FORMAT -----------//
  
  // Allignment
  periodMonthRange.setHorizontalAlignment('center');
  periodMonthRange.setVerticalAlignment('middle');
  periodWeekRange.setHorizontalAlignment('center');
  periodWeekRange.setVerticalAlignment('middle');
  periodWeekdayRange.setHorizontalAlignment('center');
  periodWeekdayRange.setVerticalAlignment('middle');
  periodDayRange.setHorizontalAlignment('center');
  periodDayRange.setVerticalAlignment('middle');
  
  // Cell size
  currentSheet.setColumnWidths(constants.CALENDAR_START_COL_POSITION, daysRange + 7, 32);
  
  // Cell color
  periodMonthRange.setBackground(style.HEADERS_BG_COLOR);
  periodWeekRange.setBackground(style.HEADERS_BG_COLOR);
  periodWeekdayRange.setBackground(style.HEADERS_BG_COLOR);
    
  // Cell borders
  periodMonthRange.setBorder(false, true, true, false, true, true);
  periodWeekdayRange.setBorder(true, true, true, true, true, true);
  periodDayRange.setBorder(true, true, true, true, true, true);
  
  
  //--------- SETTING CALENDAR VALUES ---------//
  
  for (var i=-3; i<=daysRange+3; i++) {
    
    var iterativeDate = new Date(date.getTime() + i * 3600000 * 24);
    
    var dateMonth = Utilities.formatDate(iterativeDate, "GMT","MM");
    var dateWeek = Utilities.formatDate(iterativeDate, "GMT-1","w");
    var dateWeekDay = iterativeDate.getDay();
    var dateDay = Utilities.formatDate(iterativeDate, "GMT","dd");
    
    currentSheet.getRange( constants.CALENDAR_MONTH_ROW_POSITION,    constants.CALENDAR_START_COL_POSITION + 3 + i).setValue(monthsEnum[dateMonth]);
    currentSheet.getRange( constants.CALENDAR_WEEK_ROW_POSITION,     constants.CALENDAR_START_COL_POSITION + 3 + i).setValue(dateWeek);
    currentSheet.getRange( constants.CALENDAR_WEEK_DAY_ROW_POSITION, constants.CALENDAR_START_COL_POSITION + 3 + i).setValue(weekDaysEnum[dateWeekDay]);
    currentSheet.getRange( constants.CALENDAR_DAY_ROW_POSITION,      constants.CALENDAR_START_COL_POSITION + 3 + i).setValue(dateDay);
  }
  
  
  //----------- SETTING CELLS STYLES -----------//
  for (var i=constants.CALENDAR_START_COL_POSITION; i<=constants.CALENDAR_START_COL_POSITION + 6; i++) {
    var iterativeWeekDayRange = currentSheet.getRange(constants.CALENDAR_WEEK_DAY_ROW_POSITION, i).getValue();
    if (iterativeWeekDayRange == 'Ðt' || iterativeWeekDayRange == 'Sat') {
      for (var j=0; j<=daysRange+6; j+=7) {
          currentSheet.getRange(constants.CALENDAR_WEEK_DAY_ROW_POSITION, i+j).setBackground(style.WEEKENDS_BG_COLOR);
      }
    }
    if (iterativeWeekDayRange == 'Sk' || iterativeWeekDayRange == 'Sun') {
      for (var j=0; j<=daysRange+6; j+=7) {
          currentSheet.getRange(constants.CALENDAR_WEEK_DAY_ROW_POSITION, i+j).setBackground(style.WEEKENDS_BG_COLOR);
      }
    }
  } 
  clearCalendarOverflow(currentSheet);
      
  // Merge cells
  mergeDuplicateColumns(constants.CALENDAR_MONTH_ROW_POSITION, constants.CALENDAR_START_COL_POSITION, daysRange + 7, currentSheet);
  mergeDuplicateColumns(constants.CALENDAR_WEEK_ROW_POSITION,  constants.CALENDAR_START_COL_POSITION, daysRange + 7, currentSheet);
  
  if (records > 0) {
    generateCalendarFields(currentSheet, records, constants);
  }
  
  if (records > 0 && currentSheetName.indexOf('Order') < 0) {
    setRecordValuesFormulas(records);
  }
}


function generateCalendarFields(currentSheet, records, constants) {
  
  var daysRange = Number(PropertiesService.getScriptProperties().getProperty('amountOfDays'));
  
  var calendarFieldsInnerRange = currentSheet.getRange(constants.RECORDS_START_ROW_POSITION, constants.CALENDAR_START_COL_POSITION, records, daysRange + 7);
  var calendarFieldsOuterRange = currentSheet.getRange(constants.RECORDS_START_ROW_POSITION, constants.CALENDAR_START_COL_POSITION, records, daysRange + 7);
  
  calendarFieldsInnerRange.setBorder(null, null, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.DOTTED);
  calendarFieldsOuterRange.setBorder(true, true, null, null, null, null);
  
}


function setCalendarConditionalFormatting(currentSheet) {
  
  var records = Number(PropertiesService.getScriptProperties().getProperty('amountOfRecords'));
  var daysRange = Number(PropertiesService.getScriptProperties().getProperty('amountOfDays'));
  var style = getStyleConstantsEnum();
  
  var range = currentSheet.getRange(14, 20, records, daysRange + 7);
  range.setHorizontalAlignment('center');
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground(style.CALENDAR_BG_COLOR)
    .setRanges([range])
    .build();
  var rules = currentSheet.getConditionalFormatRules();
  rules.pop();
  rules.push(rule);
  currentSheet.setConditionalFormatRules(rules);
  
}


function clearCalendar(currentSheet) {
  
  var days = Number(PropertiesService.getScriptProperties().getProperty('amountOfDays'));
  if (days > 0) {
    var calendarRange = currentSheet.getRange(10, 20, 4, days + 13);
  }
  else {
    var calendarRange = currentSheet.getRange(10, 20, 4, currentSheet.getLastColumn());
  }
  
  calendarRange.breakApart();
  calendarRange.clear();
  
}

function clearCalendarOverflow(currentSheet) {
  
  var days = Number(PropertiesService.getScriptProperties().getProperty('amountOfDays'));
  var calendarRange = currentSheet.getRange(12, days + 27, 1, 7);
  calendarRange.clear();
  
}


//----------------------------------------------------------------------------------------//
//                                         OTHER                                          //
//----------------------------------------------------------------------------------------//

function insertImage(sheetName, imgWidth, imgHeight) {
  
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var getImageBlob = DriveApp.getFileById(getImageID()).getBlob();
  
  var img = currentSheet.insertImage(getImageBlob,3,1,1,1);
  img.setWidth(imgWidth).setHeight(imgHeight);
  
}


function mergeDuplicateColumns(startRow, startCol, colRange, currentSheet) {

  var cell = {};
  var k = "";
  var offset = 0;

  // Get data from sheet
  var data = currentSheet
  .getRange(startRow, startCol, 1, colRange)
  .getValues()
  .filter(String)[0];

  // Count duplicates
  data.forEach(function(e) {
    cell[e] = cell[e] ? cell[e] + 1 : 1;
  });

  // Merge duplicate cells
  data.forEach(function(e) {
    if (e != k) {
      currentSheet.getRange(startRow, startCol + offset, 1, cell[e]).merge();
      offset += cell[e];
    }
    k = e;
  });
}


function colName(n) {
  
  var ordA = 'a'.charCodeAt(0);
  var ordZ = 'z'.charCodeAt(0);
  var len = ordZ - ordA + 1;
      
  var s = "";
  while(n >= 0) {
    s = String.fromCharCode(n % len + ordA) + s;
    n = Math.floor(n / len) - 1;
  }
  return s;
}