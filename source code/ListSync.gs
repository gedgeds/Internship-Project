function onEdit(event) 
{

  // Change Settings:
  //--------------------------------------------------------------------------------------
  var mediaPlanSheetName = getMediaPlanSheetName();                 // name of sheet with data validation
  var databaseSheetName = getMediaPlanSheetName() + ' database';    // name of sheet with data
  var numOfDropdowns = 6;                                           // number of levels of data validation
  var lcol = 1;                                                     // number of column where validation starts; A = 1, B = 2, etc.
  var lrow = 14;                                                    // number of row where validation starts
  var offsets = [1,1,1,1,2,1];                                      // offsets for levels
  
  // =====================================================================================
  SmartDataValidation(event, mediaPlanSheetName, databaseSheetName, numOfDropdowns, lcol, lrow, offsets);
  
  var currentSheet = event.source.getActiveSheet();
  currentSheet.autoResizeColumns(1, 19);

}



function SmartDataValidation(event, mediaPlanSheetName, databaseSheetName, numOfDropdowns, lcol, lrow, offsets) 
{
  //--------------------------------------------------------------------------------------
  // The event handler, adds data validation for the input parameters
  //--------------------------------------------------------------------------------------
  
  var FormulaSplitter = ';'; // depends on regional setting, ';' or ',' works for US
  //--------------------------------------------------------------------------------------
  
  //	===================================   key variables	 =================================
  //
  //		ss			    sheet we change (mediaPlanSheet)
  //		selectedRange	range to change
  //		selectedCol		number of column to edit
  //		selectedRow		number of row to edit	
  //		CurrentLevel	level of drop-down, which we change
  //		HeadLevel		main level
  //		r				current cell, which was changed by user
  //		X         		number of levels could be checked on the right
  //
  //		databaseSheet	Data sheet (databaseSheet)
  //
  //    ======================================================================================

  // Checks
  var currentSheet = event.source.getActiveSheet();
  var currentSheetName = currentSheet.getName(); 
  if (currentSheetName !== mediaPlanSheetName) { return -1;  } // not main sheet
  
  // Test if range fits
  var selectedRange = event.range;                // selected range
  var selectedCol = selectedRange.getColumn();    // column number in which the change is made
  var selectedRow = selectedRange.getRow()        // row number in which the change is made
  var amountOfCols = selectedRange.getWidth();    // how many columns selected
  
  if ((selectedCol + amountOfCols - 1) < lcol) { return -2; }  // columns... 
  if (selectedRow < lrow) { return -3; } // rows
  
  // Test range is in levels
  var columnsLevels = getColumnsOffset_(offsets, lcol); // Columns for all levels	
  var CurrentLevel = getCurrentLevel_(amountOfCols, selectedRange, selectedCol, columnsLevels);
  if(CurrentLevel === 1) { return -4; } // out of data validations
  if(CurrentLevel > numOfDropdowns) { return -5; } // last level	
  
  // Constants
  var ReplaceCommas = true;  // ReplaceCommas = true if locale uses commas to separate decimals
  var databaseSheet = SpreadsheetApp.getActive().getSheetByName(databaseSheetName); // Data sheet       				         
  var amountOfRows = selectedRange.getHeight();
  /* 	Adjust the range 'selectedRange' 
  ???       !
  xxx       x
  xxx       x 
  xxx  =>   x
  xxx       x
  xxx       x
  */	
  selectedRange = currentSheet.getRange(selectedRange.getRow(), columnsLevels[CurrentLevel - 2], amountOfRows); 
  
  // Levels
  var HeadLevel = CurrentLevel - 1; // main level
  var X = numOfDropdowns - CurrentLevel + 1; // number of levels left       
  
  // determine columns on the sheet "Data"
  var KudaCol = numOfDropdowns + 2;
  var KudaNado = databaseSheet.getRange(1, KudaCol);  // 1 place for a formula
  var lastRow = databaseSheet.getLastRow();
  var ChtoNado = databaseSheet.getRange(1, KudaCol, lastRow, KudaCol); // the range with list, returned by a formula
  
  // ============================================================================= > loop >
  var CurrLevelBase = CurrentLevel; // remember the first current level
  
  
  
  for (var j = 1; j <= amountOfRows; j++) // [01] loop rows start
  {    
    // refresh first val  
    var currentRow = selectedRange.getCell(j, 1).getRow();      
    loopColumns_(HeadLevel, X, currentRow, numOfDropdowns, CurrLevelBase, lastRow, FormulaSplitter, CurrLevelBase, columnsLevels, selectedRange, KudaNado, ChtoNado, ReplaceCommas, currentSheet);
  } // [01] loop rows end
    
  // currentSheet.autoResizeColumns(1, 19);

}


function getColumnsOffset_(offsets, lefColumn)
{
	// Columns for all levels
	var columnsLevels = [];
	var totalOffset = 0;	
	for (var i = 0, l = offsets.length; i < l; i++)
	{	
		totalOffset += offsets[i];
		columnsLevels.push(totalOffset + lefColumn - 1);
	}	
	
	return columnsLevels;
	
}


function getCurrentLevel_(amountOfCols, selectedRange, selectedCol, columnsLevels)
{
	var colPlus = 2; // const
	if (amountOfCols === 1) { return columnsLevels.indexOf(selectedCol) + colPlus; }
	var CurrentLevel = -1;
	var level = 0;
	var column = 0;
	for (var i = 0; i < amountOfCols; i++ )
	{
		column = selectedRange.offset(0, i).getColumn();
		level = columnsLevels.indexOf(column) + colPlus;
		if (level > CurrentLevel) { CurrentLevel = level; }
	}
	return CurrentLevel;
}



function loopColumns_(HeadLevel, X, currentRow, numOfDropdowns, CurrentLevel, lastRow, FormulaSplitter, CurrLevelBase, columnsLevels, selectedRange, KudaNado, ChtoNado, ReplaceCommas, currentSheet)
{
  for (var k = 1; k <= X; k++)
  {   
    HeadLevel = HeadLevel + k - 1; 
    CurrentLevel = CurrLevelBase + k - 1;
	var r = currentSheet.getRange(currentRow, columnsLevels[CurrentLevel - 2]);
	var SearchText = r.getValue(); // searched text 
	X = loopColumn_(X, SearchText, HeadLevel, HeadLevel, currentRow, numOfDropdowns, CurrentLevel, lastRow, FormulaSplitter, CurrLevelBase, columnsLevels, selectedRange, KudaNado, ChtoNado, ReplaceCommas, currentSheet);
  } 
}


function loopColumn_(X, SearchText, HeadLevel, HeadLevel, currentRow, numOfDropdowns, CurrentLevel, lastRow, FormulaSplitter, CurrLevelBase, columnsLevels, selectedRange, KudaNado, ChtoNado, ReplaceCommas, currentSheet)
{
    
  // if nothing is chosen!
  if (SearchText === '') // condition value =''
  {
    // kill extra data validation if there were 
    // columns on the right
    if (CurrentLevel <= numOfDropdowns) 
    {
      for (var f = 0; f < X; f++) 
      {
        var cell = currentSheet.getRange(currentRow, columnsLevels[CurrentLevel + f - 1]);		  
        // clean & get rid of validation
        cell.clear({contentsOnly: true});              
        cell.clear({validationsOnly: true});
        // exit columns loop  
      }
    }
    return 0;	// end loop this row	
  }
  
  
  // formula for values
  var formula = getDVListFormula_(CurrentLevel, currentRow, columnsLevels, lastRow, ReplaceCommas, FormulaSplitter, currentSheet);  
  KudaNado.setFormula(formula);
    
  
  // get response
  var Response = getResponse_(ChtoNado, lastRow, ReplaceCommas);
  var Variants = Response.length;
    

  // build data validation rule
  if (Variants === 0.0) // empty is found
  {
    return;
  }  
  if(Variants >= 1.0) // if some variants were found
  {
    
    var cell = currentSheet.getRange(currentRow, columnsLevels[CurrentLevel - 1]);
    var rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(Response, true)
    .setAllowInvalid(false)
    .build();
    // set validation rule
    cell.setDataValidation(rule);
  }    
  if (Variants === 1.0) // // set the only value
  {      
    cell.setValue(Response[0]);
    SearchText = null;
    Response = null;
    return X; // continue doing DV
  } // the only value
  
  return 0; // end DV in this row
  
}


function getDVListFormula_(CurrentLevel, currentRow, columnsLevels, lastRow, ReplaceCommas, FormulaSplitter, currentSheet)
{
  
  var checkVals = [];
  var Offs = CurrentLevel - 2;
  var values = [];
  // get values and display values for a formula
  for (var s = 0; s <= Offs; s++)
  {
    var checkR = currentSheet.getRange(currentRow, columnsLevels[s]);
    values.push(checkR.getValue());
  } 		  
  
  var LookCol = colName(CurrentLevel-1); // gets column name "A,B,C..."
  var formula = '=unique(filter(' + LookCol + '2:' + LookCol + lastRow; // =unique(filter(A2:A84

  var mathOpPlusVal = ''; 
  var value = '';

  // loop levels for multiple conditions  
  for (var i = 0; i < CurrentLevel - 1; i++) {            
    formula += FormulaSplitter; // =unique(filter(A2:A84;
    LookCol = colName(i);
    		
    value = values[i];

    mathOpPlusVal = getValueAndMathOpForFunction_(value, FormulaSplitter, ReplaceCommas); // =unique(filter(A2:A84;B2:B84="Text"
    
    if ( Array.isArray(mathOpPlusVal) )
    {
      formula += mathOpPlusVal[0];
      formula += LookCol + '2:' + LookCol + lastRow; // =unique(filter(A2:A84;ROUND(B2:B84
      formula += mathOpPlusVal[1];
    }
    else
    {
      formula += LookCol + '2:' + LookCol + lastRow; // =unique(filter(A2:A84;B2:B84
      formula += mathOpPlusVal;
    }
    
    
  }  
  
  formula += "))"; //=unique(filter(A2:A84;B2:B84="Text"))

  return formula;
}


function getValueAndMathOpForFunction_(value, FormulaSplitter, ReplaceCommas)
{
  var result = '';
  var splinter = '';	
	
  var type = typeof value;
  
 
  // strings
  if (type === 'string') return '="' + value + '"';
  // date
  if(value instanceof Date)
  {
	return ['ROUND(', FormulaSplitter +'5)=ROUND(DATE(' + value.getFullYear() + FormulaSplitter + (value.getMonth() + 1) + FormulaSplitter + value.getDate() + ')' + '+' 
	      + 'TIME(' + value.getHours() + FormulaSplitter + value.getMinutes() + FormulaSplitter + value.getSeconds() + ')' + FormulaSplitter + '5)'];	  
  }  
  // numbers
  if (type === 'number')
  {
	if (ReplaceCommas)
	{
		return '+0=' + value.toString().replace('.', ',');		
	}
	else
	{
		return '+0=' + value;
	}
  }
  // booleans
  if (type === 'boolean')
  {
	  return '=' + value;
  }  
  // other
  return '=' + value;
	
}


function getResponse_(allRange, l, ReplaceCommas)
{
  var data = allRange.getValues();
  var data_ = allRange.getDisplayValues();
  
  var response = [];
  var val = '';
  for (var i = 0; i < l; i++)
  {
    val = data[i][0];
    if (val !== '') 
    {
      var type = typeof val;
      if (type === 'boolean' || val instanceof Date) val = String(data_[i][0]);
      if (type === 'number' && ReplaceCommas) val = val.toString().replace('.', ',')
      response.push(val);  
    }
  }
  
  return response;  
}
