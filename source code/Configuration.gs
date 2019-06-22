
//----------------------------------------------------------------------------------------//
//                                       MEDIA PLAN                                       //
//----------------------------------------------------------------------------------------//

var getMediaPlanSheetName = (function () {
  var mediaPlanSheetName = 'Media planas';
  return function () {return mediaPlanSheetName}
})();

var getMetaHeadersPlanLT = (function () {
  var metaHeadersPlanLT = [['Uþsakovas:'],['Klientas:'],['Produktas:'],['Kampanija:'],['Periodas:'],['Ðalis:'],['Kampanijos ID:'],['Kampanijos pavadinimas:']];
  return function () {return metaHeadersPlanLT}
})();

var getMetaHeadersPlanEN = (function () {
  var metaHeadersPlanEN = [['Agency:'],['Client:'],['Product:'],['Campaign:'],['Period:'],['Country:'],['Campaign ID:'],['Campaign Name:']];
  return function () {return metaHeadersPlanEN}
})();

var getMetaValues = (function () {
  var metaValues = [['BPN LT'],['KLIENTAS'],['KLIENTAS'],[''],['2017.11.02-2017.11.30'],[''],[''],['']];
  return function () {return metaValues}
})();
    
var getMediaPlanConstantsEnum = (function () {
  var mediaPlanConstants = {
    HEADERS_START_ROW_POSITION: 10,
    RECORDS_START_ROW_POSITION: 14,
    AMOUNT_OF_META_HEADERS: 8,
    AMOUNT_OF_MAIN_COLUMNS: 16,
    AMOUNT_OF_TOTAL_COLUMNS: 19,
    META_HEADERS_FONT_SIZE: 8,
    META_VALUES_FONT_SIZE: 8,
    BPN_LOGO_WIDTH: 287,
    BPN_LOGO_HEIGHT: 164,
    CALENDAR_START_COL_POSITION: 20,
    CALENDAR_MONTH_ROW_POSITION: 10,
    CALENDAR_WEEK_ROW_POSITION: 11,
    CALENDAR_WEEK_DAY_ROW_POSITION: 12,
    CALENDAR_DAY_ROW_POSITION: 13,
};
  return function () {return mediaPlanConstants}
})();



//----------------------------------------------------------------------------------------//
//                                       ORDERS                                           //
//----------------------------------------------------------------------------------------//

var getOrderSheetName = (function () {
  var orderSheetName = 'Order';
  return function () {return orderSheetName}
})();
    
var getMetaHeadersOrderLT = (function () {
  var metaHeadersOrderLT = [['Uþsakovas:'],['Klientas:'],['Produktas:'],['Kampanija:'],['Periodas:'],['Ðalis:']];
  return function () {return metaHeadersOrderLT}
})();

var getMetaHeadersOrderEN = (function () {
  var metaHeadersOrderEN = [['Agency:'],['Client:'],['Product:'],['Campaign:'],['Period:'],['Country:']];
  return function () {return metaHeadersOrderEN}
})();

var getOrderTotalsHeadersLT = (function () {
  var orderTotalsValuesLT = ['','','','','','','','','','','','','','Gross kaina:','Nuolaida:','Neto kaina:'];
  return function () {return orderTotalsValuesLT}
})();

var getOrderTotalsHeadersEN = (function () {
  var orderTotalsValuesEN = ['','','','','','','','','','','','','','Gross price:','Discount:','Neto price:'];
  return function () {return orderTotalsValuesEN}
})();

var getOrderConstantsEnum = (function () {
  var mediaPlanConstants = {
    HEADERS_START_ROW_POSITION: 8,
    RECORDS_START_ROW_POSITION: 12,
    AMOUNT_OF_META_HEADERS: 6,
    AMOUNT_OF_MAIN_COLUMNS: 16,
    AMOUNT_OF_TOTAL_COLUMNS: 16,
    AMOUNT_OF_FROZEN_COLUMNS: 2,
    META_HEADERS_FONT_SIZE: 8,
    META_VALUES_FONT_SIZE: 8,
    BPN_LOGO_WIDTH: 216,
    BPN_LOGO_HEIGHT: 123,
    CALENDAR_START_COL_POSITION: 17,
    CALENDAR_MONTH_ROW_POSITION: 8,
    CALENDAR_WEEK_ROW_POSITION: 9,
    CALENDAR_WEEK_DAY_ROW_POSITION: 10,
    CALENDAR_DAY_ROW_POSITION: 11,
};
  return function () {return mediaPlanConstants}
})();



//----------------------------------------------------------------------------------------//
//                                       GENERAL                                          //
//----------------------------------------------------------------------------------------//

var getImageID = (function () {
  var imageID = '14cl5sdMVa_yTaQ-aVWWaGjA6AoUbGaGs';
  return function () {return imageID}
})();

var getLanguageOptions = (function () {
  var languages = ['LT','EN'];
  return function () {return languages}
})();

var getMonthsEnumLT = (function () {
  var monthsEnum = {'01': 'Sausis', '02': 'Vasaris', '03': 'Kovas', '04': 'Balandis', '05': 'Geguþë', '06': 'Birþelis', '07': 'Liepa', '08': 'Rugpjûtis', '09': 'Rugsëjis', 10: 'Spalis', 11: 'Lapkritis', 12: 'Gruodis'};
  return function () {return monthsEnum}
})();

var getMonthsEnumEN = (function () {
  var monthsEnum = {'01': 'January', '02': 'February', '03': 'March', '04': 'April', '05': 'May', '06': 'June', '07': 'July', '08': 'August', '09': 'September', 10: 'October', 11: 'November', 12: 'December'};
  return function () {return monthsEnum}
})();

var getWeekDaysEnumLT = (function () {
  var weekDaysEnum = {0:'Sk', 1:'Pr', 2:'An', 3:'Tr', 4:'Kt', 5:'Pn', 6:'Ðt'};
  return function () {return weekDaysEnum}
})();

var getWeekDaysEnumEN = (function () {
  var weekDaysEnum = {0:'Sun', 1:'Mon', 2:'Tue', 3:'Wed', 4:'Thu', 5:'Fri', 6:'Sat'};
  return function () {return weekDaysEnum}
})();

var getHeadersMidLT = (function () {
  var headersMidLT = ['Tiekëjas','Tinklapis','Aplinka','Nukreipimas','Ribojimas','Skydelis, px','Skydelio tipas','Vienetas','Kiekis','Pirkimo tipas','Viso','Vieneto gross kaina','Vieneto neto kaina','Gross kaina (EUR)','Nuolaida (%)','Neto kaina (EUR)'];
  return function () {return headersMidLT}
})();

var getHeadersMidEN = (function () {
  var headersMidEN = ['Channel','Website','Position','Targeting','Capping','Banner size, px','Animation','Unit','Quantity','Buying type','Total','Gross price per unit','Neto price per unit','Gross price (EUR)','Discount, %','Neto price (EUR)'];
  return function () {return headersMidEN}
})();

var getPredictionsHeadersTopLT = (function () {
  var predictionsHeadersTop = ['','Prognozës',''];
  return function () {return predictionsHeadersTop}
})();

var getPredictionsHeadersTopEN = (function () {
  var predictionsHeadersTop = ['','Predictions',''];
  return function () {return predictionsHeadersTop}
})();

var getPredictionsHeadersMidLT = (function () {
  var predictionsHeadersMid = ['Parodymai','CTR','Nukreipimai'];
  return function () {return predictionsHeadersMid}
})();

var getPredictionsHeadersMidEN = (function () {
  var predictionsHeadersMid = ['Evidence','CTR','Redirects'];
  return function () {return predictionsHeadersMid}
})();

var getConfirmationDateValuesLT = (function () {
  var confirmationDateValues = [['Vëliausias patvirtinimas iki:','','','','','']];
  return function () {return confirmationDateValues}
})();

var getConfirmationDateValuesEN = (function () {
  var confirmationDateValues = [['Latest confirmation until:','','','','','']];
  return function () {return confirmationDateValues}
})();

var getSignatureValuesLT = (function () {
  var signatureValues = [
    ['KLIENTO PATVIRTINIMAS','','','','',''],
    ['','','','','',''],
    ['Vardas, pavardë:','','','','',''],
    ['','','','','',''],
    ['','','','','',''],
    ['','','','','',''],
    ['','','','','',''],
    ['Data:','','','','',''],
    ['','','','','',''],
    ['Paraðas:','','','','',''],
    ['','','','','','']
  ];
  return function () {return signatureValues}
})();

var getSignatureValuesEN = (function () {
  var signatureValues = [
    ['CLIENT CONFIRMATION','','','','',''],
    ['','','','','',''],
    ['First name, last name:','','','','',''],
    ['','','','','',''],
    ['','','','','',''],
    ['','','','','',''],
    ['','','','','',''],
    ['Date:','','','','',''],
    ['','','','','',''],
    ['Signature:','','','','',''],
    ['','','','','','']
  ];
  return function () {return signatureValues}
})();

var getPriceFormat = (function () {
  var priceFormat = '€#,##0.00';
  return function () {return priceFormat}
})();

var getPercentFormat = (function () {
  var percentFormat = '0.00%';
  return function () {return percentFormat}
})();

var getStyleConstantsEnum = (function () {
  var styleConstants = {
    HEADERS_BG_COLOR:      '#F1F1F1',
    WEEKENDS_BG_COLOR:     '#D2D2D2',
    ATTENTION_BG_COLOR:    '#CC0000',
    CALENDAR_BG_COLOR:     '#FFFF99',
    ORDER_TOTALS_BG_COLOR: '#F1F1F1',
    GLOBAL_FONT_FAMILY:    'Tahoma',
};
  return function () {return styleConstants}
})();



//----------------------------------------------------------------------------------------//
//                                       DATABASE                                         //
//----------------------------------------------------------------------------------------//

var getDatabaseSpreadsheetID = (function () {
  var databaseSpreadsheetID = '1lMB0rzyDpqXEUNYSHT-zv_uGx2jzqKXvT1YFGxQxv_E';
  return function () {return databaseSpreadsheetID}
})();

var getDatabaseSheetName = (function () {
  var databaseSheetName = 'database';
  return function () {return databaseSheetName}
})();



//----------------------------------------------------------------------------------------//
//                                         ALERTS                                         //
//----------------------------------------------------------------------------------------//

function showWrongDateAlert() {
  
  var ui = SpreadsheetApp.getUi();
  ui.alert(
     'Logical failure!',
     'Beginning of the period must be earlier than the end of period!',
      ui.ButtonSet.OK
  );  
}

function showOverflowDateAlert() {
  
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
     'Attention!',
     'Period is longer than 100 days. Continue?',
      ui.ButtonSet.YES_NO);

  if (result == ui.Button.YES) {
    return false;
  } 
  if (result == ui.Button.NO) {
    return true;
  } else {
    return true;
  }
}

function showMediaPlanDuplicateAlert() {
  
  var ui = SpreadsheetApp.getUi();
  ui.alert(
     'Attention!',
     'Please rename current Media Plan sheet',
      ui.ButtonSet.OK
  );
}

function showOrderDuplicateAlert(orderChannel) {
  
  var ui = SpreadsheetApp.getUi();
  ui.alert(
     'Attention!',
     'Please rename current ' + orderChannel + ' order sheet name',
      ui.ButtonSet.OK
  );
}

function showInvalidOrderCreationAlert() {
  
  var ui = SpreadsheetApp.getUi();
  var ui = SpreadsheetApp.getUi();
  ui.alert(
     'Attention!',
     'Order creation can only be made from Media Plan sheet',
      ui.ButtonSet.OK
  );
}

function showInvalidDatabaseUpdateAlert() {
  
  var ui = SpreadsheetApp.getUi();
  var ui = SpreadsheetApp.getUi();
  ui.alert(
     'Attention!',
     'Database can only be updated from Media Plan sheet',
      ui.ButtonSet.OK
  );
}