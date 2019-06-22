function prepareOrderTemplate(paramLanguage, paramChannel) {
   
  
  var style = getStyleConstantsEnum();
  var constant = getOrderConstantsEnum();
  var spreadsheet = SpreadsheetApp.getActive();
  var orderSheet = spreadsheet.getSheetByName(getOrderSheetName() + ' ' + paramChannel);
  
  // check if order sheet already exists
  if (!orderSheet) {
    var orderSheet = spreadsheet.insertSheet(getOrderSheetName() + ' ' + paramChannel);    
  } 
  else if (orderSheet) {
    showOrderDuplicateAlert(paramChannel);
    return;
  }
    
  //----------- INSTANTIATING DATA -----------//
  
  if (paramLanguage === 'LT') {
    var metaHeaders = getMetaHeadersOrderLT();
    var headersMid = getHeadersMidLT();
    PropertiesService.getScriptProperties().setProperty('calendarLanguage', 'LT');
  } 
  else if (paramLanguage === 'EN') {
    var metaHeaders = getMetaHeadersOrderEN();
    var headersMid = getHeadersMidEN();
    PropertiesService.getScriptProperties().setProperty('calendarLanguage', 'EN');
  }
    
  //----------- SETTING DATA RANGES -----------//
  var metaHeadersRange = orderSheet.getRange(1, 1, constant.AMOUNT_OF_META_HEADERS, 1);
  var metaValuesRange = orderSheet.getRange(1, 2, constant.AMOUNT_OF_META_HEADERS, 1);
  var headersRange = orderSheet.getRange(constant.HEADERS_START_ROW_POSITION, 1, 1, constant.AMOUNT_OF_MAIN_COLUMNS);
  
  //----------- SETTING DATA VALUES -----------//
  metaHeadersRange.setValues(metaHeaders);
  headersRange.setValues([headersMid]);
    
  //----------- SETTING CELLS STYLES -----------//
  // Font family
  orderSheet.getRange(1, 1, orderSheet.getLastRow(), orderSheet.getLastColumn()).setFontFamily(style.GLOBAL_FONT_FAMILY);
  
  // Font weight
  metaHeadersRange.setFontWeight('bold');
  metaValuesRange.setFontWeight('bold');
  
  // Font size
  metaHeadersRange.setFontSize(constant.META_HEADERS_FONT_SIZE);
  metaValuesRange.setFontSize(constant.META_VALUES_FONT_SIZE);
  
  // Merge cells
  for (var i=1; i<=constant.AMOUNT_OF_MAIN_COLUMNS; i++) {
    orderSheet.getRange(constant.HEADERS_START_ROW_POSITION, i, 3, 1).merge();
  }
  
  // Allignment
  metaHeadersRange.setHorizontalAlignment('left');
  metaHeadersRange.setVerticalAlignment('middle');
  metaValuesRange.setHorizontalAlignment('left');
  metaValuesRange.setVerticalAlignment('middle');
  headersRange.setHorizontalAlignment('center');
  headersRange.setVerticalAlignment('middle');
    
  // Cell color
  metaHeadersRange.setBackground(style.HEADERS_BG_COLOR);
  metaValuesRange.setBackground(style.HEADERS_BG_COLOR);
  headersRange.setBackground(style.HEADERS_BG_COLOR);
    
  // Cell borders
  headersRange.setBorder(true, true, false, true, true, true);
     
  // Other
  insertImage(getOrderSheetName() + ' ' + paramChannel, constant.BPN_LOGO_WIDTH, constant.BPN_LOGO_HEIGHT);
  orderSheet.setHiddenGridlines(true);
  orderSheet.setFrozenColumns(constant.AMOUNT_OF_FROZEN_COLUMNS);
  
  copyOrderMetaValues(); 
  generateCalendar();
  copyOrderRecords(paramLanguage, paramChannel); 
  
  // Cell size
  orderSheet.autoResizeColumns(1, constant.AMOUNT_OF_TOTAL_COLUMNS);
  orderSheet.autoResizeRows(1, constant.AMOUNT_OF_META_HEADERS);
  
}

function copyOrderMetaValues() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  var currentSheet = spreadsheet.getActiveSheet();
  var mediaPlanSheet = spreadsheet.getSheetByName(getMediaPlanSheetName());
  
  var sourceRecordsRange = mediaPlanSheet.getRange(1, 2, 6, 1);
  var targetRecordsRange = currentSheet.getRange(1, 2, 6, 1);
  sourceRecordsRange.copyTo(targetRecordsRange);
  
}

function copyOrderRecords(paramLanguage, paramChannel) {
  
  var constant = getOrderConstantsEnum();
  var constantMP = getMediaPlanConstantsEnum();
  var spreadsheet = SpreadsheetApp.getActive();
  var currentSheet = spreadsheet.getActiveSheet();
  var mediaPlanSheet = spreadsheet.getSheetByName(getMediaPlanSheetName());
  
  var records = Number(PropertiesService.getScriptProperties().getProperty('amountOfRecords'));
  var days = Number(PropertiesService.getScriptProperties().getProperty('amountOfDays'));
  var orderRowsToCopy = [];
  
  for (var i=0; i<records; i++) {
    if (mediaPlanSheet.getRange(constantMP.RECORDS_START_ROW_POSITION + i, 1).getValue() == paramChannel) {
      orderRowsToCopy.push(constantMP.RECORDS_START_ROW_POSITION + i);
    }
  }
  
  for (var i=0; i<orderRowsToCopy.length; i++) {
    var sourceRecordsRange = mediaPlanSheet.getRange(orderRowsToCopy[i], 1, 1, constant.AMOUNT_OF_TOTAL_COLUMNS);
    var targetRecordsRange = currentSheet.getRange(constant.RECORDS_START_ROW_POSITION + i, 1, 1, constant.AMOUNT_OF_TOTAL_COLUMNS);
    sourceRecordsRange.copyTo(targetRecordsRange);
  }
    
  PropertiesService.getScriptProperties().setProperty('amountOfOrderRecords', orderRowsToCopy.length);
  generateOrderTotals(orderRowsToCopy.length, paramLanguage);
  copyCalendarValues(orderRowsToCopy);
  
}

function copyCalendarValues(orderRowsToCopy) {
  
  var spreadsheet = SpreadsheetApp.getActive();
  var currentSheet = spreadsheet.getActiveSheet();
  var mediaPlanSheet = spreadsheet.getSheetByName(getMediaPlanSheetName());
  
  var days = Number(PropertiesService.getScriptProperties().getProperty('amountOfDays'));
  
  for (var i=0; i<orderRowsToCopy.length; i++) {
    var sourceRecordsRange = mediaPlanSheet.getRange(orderRowsToCopy[i], 20, 1, days+7);
    var targetRecordsRange = currentSheet.getRange(12 + i, 17, 1, days+7);
    sourceRecordsRange.copyTo(targetRecordsRange);
  }
  
}

function generateOrderTotals(orderRecords, paramLanguage) {
  
  var style = getStyleConstantsEnum();
  var constant = getOrderConstantsEnum();
  
  var spreadsheet = SpreadsheetApp.getActive();
  var currentSheet = spreadsheet.getActiveSheet();
  
  var totalsHeadersRange = currentSheet.getRange(constant.RECORDS_START_ROW_POSITION + orderRecords, 1, 1, constant.AMOUNT_OF_TOTAL_COLUMNS);
  var totalsValuesRange = currentSheet.getRange(constant.RECORDS_START_ROW_POSITION + orderRecords + 1, 1, 1, constant.AMOUNT_OF_TOTAL_COLUMNS);
  
  if (paramLanguage === 'LT') {
    var totalsHeaders = getOrderTotalsHeadersLT();
  } 
  else if (paramLanguage === 'EN') {
    var totalsHeaders = getOrderTotalsHeadersEN();
  } 
  totalsHeadersRange.setValues([totalsHeaders]);
  
  var colGrossPrice = 14;
  var colDiscount = 15;
  var colNettoPrice = 16;
  
  currentSheet.getRange(12 + orderRecords + 1, colGrossPrice).setFormula('=sum(' + colName(colGrossPrice-1) + '12' + ':' + colName(colGrossPrice-1) + (12 + orderRecords - 1).toString() + ')');
  currentSheet.getRange(12 + orderRecords + 1, colDiscount).setFormula('1-' + colName(colNettoPrice-1) + (12+1+orderRecords).toString() + '/' + colName(colGrossPrice-1) + (12+1+orderRecords).toString());
  currentSheet.getRange(12 + orderRecords + 1, colNettoPrice).setFormula('=sum(' + colName(colNettoPrice-1) + '12' + ':' + colName(colNettoPrice-1) + (12 + orderRecords - 1).toString() + ')');
  currentSheet.getRange(13 + orderRecords, colDiscount).setNumberFormat(getPercentFormat());
  
  totalsHeadersRange.setBackground(style.ORDER_TOTALS_BG_COLOR);
  totalsValuesRange.setBackground(style.ORDER_TOTALS_BG_COLOR);
  
}

