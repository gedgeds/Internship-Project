function openMediaPlanModalDialog() {
  
  var htmlDlg = HtmlService.createHtmlOutputFromFile('ModalMediaPlanCreation')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(250)
      .setHeight(150);
  SpreadsheetApp.getUi() 
      .showModalDialog(htmlDlg, 'Kurti Media planà');
}

function openOrderModalDialog() {
  
  // check if function is called from media plan sheet
  var currentSheetName = SpreadsheetApp.getActiveSheet().getName();
  if(currentSheetName.indexOf(getMediaPlanSheetName()) < 0){
    showInvalidOrderCreationAlert();
    return;
  }
  
  var htmlDlg = HtmlService.createHtmlOutputFromFile('ModalOrderCreation')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(200)
      .setHeight(200);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlDlg, 'Kurti uþsakymà');
}

function getLanguageOptionsHTML() {
  
  return getLanguageOptions();
  
}

function getChannelOptionsHTML() {
  
  var currentSheet = SpreadsheetApp.getActive().getActiveSheet();
  var channelsRange = currentSheet.getRange('A14:A');
  var channelColValues = channelsRange.getValues();
  var allChannels = [];
  
  for(var i=0; i<channelColValues.length; i++){
    allChannels.push(channelColValues[i][0]);
  }
  
  var uniqueChannels = allChannels.unique();
  
  var index = uniqueChannels.indexOf('');
  if (index !== -1) uniqueChannels.splice(index, 1);
    
  return uniqueChannels;
  
}

Array.prototype.unique = function() {
  
  return this.filter(function (value, index, self) { 
    return self.indexOf(value) === index;
  });
}