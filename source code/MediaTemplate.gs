function prepareMediaPlanTemplate(paramLanguage) {
  
  
  var style = getStyleConstantsEnum();  
  var constant = getMediaPlanConstantsEnum();
  var spreadsheet = SpreadsheetApp.getActive();
  var mediaPlanSheet = spreadsheet.getSheetByName(getMediaPlanSheetName());
  var mediaPlanDatabaseSheet = spreadsheet.getSheetByName(getMediaPlanSheetName() + ' database');
  
  
  //---------- CHECK IF MEDIA PLAN SHEET AND DATABASE IS ALREADY CREATED ----------//  
  
  if (!mediaPlanSheet && !mediaPlanDatabaseSheet) {
    var mediaPlanSheet = spreadsheet.insertSheet(getMediaPlanSheetName());
    createDatabase(spreadsheet, mediaPlanSheet);
    PropertiesService.getScriptProperties().setProperty('amountOfDays', 0);
    PropertiesService.getScriptProperties().setProperty('amountOfRecords', 0); 
  } 
  else if (!mediaPlanSheet && mediaPlanDatabaseSheet) {
    var mediaPlanSheet = spreadsheet.insertSheet(getMediaPlanSheetName());
    PropertiesService.getScriptProperties().setProperty('amountOfDays', 0);
    PropertiesService.getScriptProperties().setProperty('amountOfRecords', 0);   
  } 
  else if (mediaPlanSheet && !mediaPlanDatabaseSheet) {
    showMediaPlanDuplicateAlert();
    createDatabase(spreadsheet, mediaPlanSheet);
    return;  
  } 
  else if (mediaPlanSheet && mediaPlanDatabaseSheet) {
    showMediaPlanDuplicateAlert();
    return;
  }
    
  //----------- INSTANTIATING DATA -----------//
  
  if (paramLanguage === 'LT') {
    var metaHeaders = getMetaHeadersPlanLT();
    var headersMid = getHeadersMidLT();
    PropertiesService.getScriptProperties().setProperty('calendarLanguage', 'LT');
  } 
  else if (paramLanguage === 'EN') {
    var metaHeaders = getMetaHeadersPlanEN();
    var headersMid = getHeadersMidEN();
    PropertiesService.getScriptProperties().setProperty('calendarLanguage', 'EN');
  }
  var metaValues = getMetaValues();
  
  
  //----------- SETTING DATA RANGES -----------//
  var metaHeadersRange = mediaPlanSheet.getRange(1, 1, constant.AMOUNT_OF_META_HEADERS, 1);
  var metaValuesRange = mediaPlanSheet.getRange(1, 2, constant.AMOUNT_OF_META_HEADERS, 1);
  var headersRange = mediaPlanSheet.getRange(constant.HEADERS_START_ROW_POSITION, 1, 1, constant.AMOUNT_OF_MAIN_COLUMNS);
  
  
  //----------- SETTING DATA VALUES -----------//
  metaHeadersRange.setValues(metaHeaders);
  metaValuesRange.setValues(metaValues);
  headersRange.setValues([headersMid]);
  generateMediaPlanPredictions(paramLanguage);
  
  
  //----------- SETTING CELLS STYLES -----------//
  
  // Font family
  mediaPlanSheet.getRange(1, 1, mediaPlanSheet.getLastRow(), mediaPlanSheet.getLastColumn()).setFontFamily(style.GLOBAL_FONT_FAMILY);
  
  // Font weight
  metaHeadersRange.setFontWeight('bold');
  metaValuesRange.setFontWeight('bold');
  
  // Font size
  metaHeadersRange.setFontSize(constant.META_HEADERS_FONT_SIZE);
  metaValuesRange.setFontSize(constant.META_VALUES_FONT_SIZE);
  
  // Merge cells
  for (var i=1; i<=constant.AMOUNT_OF_MAIN_COLUMNS; i++) {
    mediaPlanSheet.getRange(constant.HEADERS_START_ROW_POSITION, i, 3, 1).merge();
  }
  
  // Allignment
  metaHeadersRange.setHorizontalAlignment('left');
  metaHeadersRange.setVerticalAlignment('middle');
  metaValuesRange.setHorizontalAlignment('left');
  metaValuesRange.setVerticalAlignment('middle');
  headersRange.setHorizontalAlignment('center');
  headersRange.setVerticalAlignment('middle');
  
  // Cell size
  mediaPlanSheet.autoResizeColumns(1, constant.AMOUNT_OF_TOTAL_COLUMNS);
  mediaPlanSheet.autoResizeRows(1, constant.AMOUNT_OF_META_HEADERS);
  
  // Cell color
  metaHeadersRange.setBackground(style.HEADERS_BG_COLOR);
  metaValuesRange.setBackground(style.HEADERS_BG_COLOR);
  headersRange.setBackground(style.HEADERS_BG_COLOR);
    
  // Cell borders
  headersRange.setBorder(false, true, false, true, true, true);
  
  // Other
  insertImage(getMediaPlanSheetName(), constant.BPN_LOGO_WIDTH, constant.BPN_LOGO_HEIGHT);
  mediaPlanSheet.setHiddenGridlines(true);
  
}


function generateMediaPlanPredictions(paramLanguage) {
  
  var style = getStyleConstantsEnum();
  var mediaPlanSheet = SpreadsheetApp.getActive().getSheetByName(getMediaPlanSheetName()); 
  
  if (paramLanguage === 'LT') {
    var predictionsHeadersTop = getPredictionsHeadersTopLT();
    var predictionsHeadersMid = getPredictionsHeadersMidLT();
  } 
  else if (paramLanguage === 'EN') {
    var predictionsHeadersTop = getPredictionsHeadersTopEN();
    var predictionsHeadersMid = getPredictionsHeadersMidEN();
  }
  
  var predictionsHeadersTopRange = mediaPlanSheet.getRange('Q10:S10');
  var predictionsHeadersMidRange = mediaPlanSheet.getRange('Q11:S11');
  var predictionsHeadersBotRange = mediaPlanSheet.getRange('Q12:S12');
  
  predictionsHeadersTopRange.setValues([predictionsHeadersTop]);
  predictionsHeadersMidRange.setValues([predictionsHeadersMid]);
  
  
  //----------- SETTING CELLS STYLES -----------//  
  
  // Allignment
  predictionsHeadersTopRange.setHorizontalAlignment('center');
  predictionsHeadersTopRange.setVerticalAlignment('middle');
  predictionsHeadersMidRange.setHorizontalAlignment('center');
  predictionsHeadersMidRange.setVerticalAlignment('middle');
  predictionsHeadersBotRange.setHorizontalAlignment('center');
  predictionsHeadersBotRange.setVerticalAlignment('middle');
  
  // Cell color
  predictionsHeadersTopRange.setBackground(style.HEADERS_BG_COLOR);
  predictionsHeadersMidRange.setBackground(style.HEADERS_BG_COLOR);
  predictionsHeadersBotRange.setBackground(style.HEADERS_BG_COLOR);
    
  // Cell borders
  predictionsHeadersTopRange.setBorder(false, true, true, true, true, true);
  predictionsHeadersMidRange.setBorder(true, true, true, true, true, true);
  predictionsHeadersBotRange.setBorder(true, true, true, true, true, true);
  
  // Merge cells
  mediaPlanSheet.getRange('Q10:S10').merge();
  mediaPlanSheet.getRange('Q11:Q12').merge();
  mediaPlanSheet.getRange('R11:R12').merge();
  mediaPlanSheet.getRange('S11:S12').merge();
  
}


function finalizeMediaPlanTemplate() {
  
  // check if function is called from media plan sheet
  var currentSheetName = SpreadsheetApp.getActiveSheet().getName();
  if(currentSheetName.indexOf(getMediaPlanSheetName()) < 0) {return;}
  
  var style = getStyleConstantsEnum();
  var spreadsheet = SpreadsheetApp.getActive();
  var mediaPlanSheet = spreadsheet.getSheetByName(getMediaPlanSheetName());  
  var signaturePosition = mediaPlanSheet.getLastRow() + 5;
  
  var signatureNameBorderRange = mediaPlanSheet.getRange(signaturePosition + 4, 3, 1, 4);
  var signatureDateBorderRange = mediaPlanSheet.getRange(signaturePosition + 9, 3, 1, 4);
  var signatureSignBorderRange = mediaPlanSheet.getRange(signaturePosition + 11, 3, 1, 4);
  
  var confirmationDateRange = mediaPlanSheet.getRange(signaturePosition, 1, 1, 6);
  var signatureRange = mediaPlanSheet.getRange(signaturePosition + 2, 1, 11, 6);
  var deadlineDateCell = mediaPlanSheet.getRange(signaturePosition, 6);
  
  var mediaPlanSheetLang = PropertiesService.getScriptProperties().getProperty('calendarLanguage');
  
  if(mediaPlanSheetLang === 'LT'){
    var confirmationDateValues = getConfirmationDateValuesLT();
    var signatureValues = getSignatureValuesLT();
  }
  else if(mediaPlanSheetLang === 'EN'){
    var confirmationDateValues = getConfirmationDateValuesEN();
    var signatureValues = getSignatureValuesEN();
  }
  
  var datePeriodValue = mediaPlanSheet.getRange(5, 2).getValue();
  var periodStartDate = datePeriodValue.substring(0, 10);
  periodStartDate = periodStartDate.replace(/\./g, "-");
  var date = new Date(periodStartDate);
  var deadlineDate = new Date(date.getTime() - 3600000 * 24);
  
  signatureRange.setValues(signatureValues);
  confirmationDateRange.setValues(confirmationDateValues);
  deadlineDateCell.setValue(deadlineDate);
 
  
  //----------- SETTING CELLS STYLES -----------//  
  
  // Font family
  mediaPlanSheet.getRange(1, 1, mediaPlanSheet.getLastRow(), mediaPlanSheet.getLastColumn()).setFontFamily(style.GLOBAL_FONT_FAMILY);
  
  // Font weight
  deadlineDateCell.setFontWeight('bold');
  
  // Font color
  deadlineDateCell.setFontColor(style.ATTENTION_BG_COLOR);
  
  // Cell color
  signatureRange.setBackground(style.HEADERS_BG_COLOR);
  confirmationDateRange.setBackground(style.HEADERS_BG_COLOR);
  
  // Cell borders
  signatureRange.setBorder(true, true, true, true, false, false);
  signatureNameBorderRange.setBorder(false, false, true, true, false, false);
  signatureDateBorderRange.setBorder(false, false, true, true, false, false);
  signatureSignBorderRange.setBorder(false, false, true, true, false, false);
  confirmationDateRange.setBorder(true, true, true, true, false, false);
  
  mediaPlanSheet.autoResizeColumns(1, 19);
  
}

