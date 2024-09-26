function entete() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
};

function entete_planning_kholle() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:E1').activate()
  .mergeAcross();
  spreadsheet.getActiveRangeList().setFontSize(13)
  .setFontSize(14)
  .setFontSize(15)
  .setFontWeight('bold')
  .setHorizontalAlignment('center')
  .setBackground('#9fc5e8')
  .setFontWeight(null)
  .setFontWeight('bold');
};

function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H3').activate();
  spreadsheet.getRange('H3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Option 1', 'Option 2'], true)
  .build());
  spreadsheet.getRange('H3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Option 2'], true)
  .build());
  spreadsheet.getRange('H3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['P', 'Option 2'], true)
  .build());
  spreadsheet.getRange('H3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Prof', 'Option 2'], true)
  .build());
  spreadsheet.getRange('H3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Prof ', 'Option 2'], true)
  .build());
  spreadsheet.getRange('H3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Prof 1', 'Option 2'], true)
  .build());
  spreadsheet.getRange('H3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Prof 1'], true)
  .build());
  spreadsheet.getRange('H3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Prof 1', 'P'], true)
  .build());
  spreadsheet.getRange('H3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Prof 1', 'Prof '], true)
  .build());
  spreadsheet.getRange('H3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Prof 1', 'Prof 2'], true)
  .build());
};

function MERGE() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B10:B16').activate()
  .mergeVertically();
  spreadsheet.getActiveRangeList().setVerticalAlignment('middle');
};

function centre() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A3:A9').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
};

function nameranged() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.setNamedRange('Année', spreadsheet.getRange('A1'));
};

function center() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C7:E7').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
};

function mef() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2:E2').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('E2'));
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('A1:E23').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
};

function cl() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:E23').activate();
  spreadsheet.getRange('A1:E23').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  var banding = spreadsheet.getRange('A1:E23').getBandings()[0];
  banding.setHeaderRowColor('#5b95f9')
  .setFirstRowColor('#ffffff')
  .setSecondRowColor('#e8f0fe')
  .setFooterRowColor(null);
};

function autofit() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
};

function autofit1() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveSheet().autoResizeColumns(1, 26);
};

function bordure_epaisse() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:E6').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  spreadsheet.getRange('A2:E2').activate();
  spreadsheet.getActiveRangeList().setFontWeight('bold')
  .setFontSize(11);
};

function autof() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveSheet().autoResizeColumns(1, 26);
};

function tri() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort([{column: 3, ascending: true}, {column: 1, ascending: true}, {column: 2, ascending: true}, {column: 7, ascending: true}]);
};

function tri_kholleur() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:J111').activate();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort([{column: 5, ascending: true}, {column: 6, ascending: true}, {column: 1, ascending: true}, {column: 2, ascending: true}, {column: 7, ascending: true}]);
};

function alternance() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:H44').activate();
  spreadsheet.getRange('A1:H44').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  var banding = spreadsheet.getRange('A1:H44').getBandings()[0];
  banding.setHeaderRowColor('#5b95f9')
  .setFirstRowColor('#ffffff')
  .setSecondRowColor('#e8f0fe')
  .setFooterRowColor(null);
};

function case_cochee() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E3').activate();
  spreadsheet.getCurrentCell().setValue('TRUE');
};

function coul() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C12').activate();
  spreadsheet.getActiveRangeList().setBackground('#00ff00');
};

function nocoul() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C12').activate();
  spreadsheet.getActiveRangeList().setBackground(null);
};

function vert() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('I2').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('M. LECOUTEUX Guilhem : 0/8 notes - 0/8 commentaires\n PERSANDA Inconnu : 8/8 notes - 8/8 commentaires\n RUIZ Inconnu : 0/8 notes - 0/8 commentaires')
  .setTextStyle(52, 101, SpreadsheetApp.newTextStyle()
  .setForegroundColor('#00ff00')
  .build())
  .build());
};

function printSelectedRange() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var range = spreadsheet.getActiveSheet().getActiveRange();

  // Customize print settings (optional)
  var printOptions = {
    //orientation: SpreadsheetApp.Orientation.PORTRAIT,
    margins: { top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 },
    scale: 1.2,
    header: {
      odd: "My Report",
      even: "My Report"
    },
    footer: {
      odd: "Page [page] of [pages]",
      even: "Page [page] of [pages]"
    }
  };

  // Print the range with the specified options
  spreadsheet.printRange(range, printOptions);
}

function sup() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E6').activate();
  spreadsheet.getCurrentCell().setValue('TRUE');
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

function creee_tcd(feuille_donnees,range_donnees,feuille_TCD) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceData = spreadsheet.getSheetByName(feuille_donnees).getRange(range_donnees);
  spreadsheet.insertSheet(spreadsheet.getActiveSheet().getIndex() + 1).activate();
  var existingSheet = spreadsheet.getSheetByName(feuille_TCD);
  if (existingSheet) 
    spreadsheet.deleteSheet(existingSheet);
  spreadsheet.getActiveSheet().setName(feuille_TCD);
  spreadsheet.getActiveSheet().setHiddenGridlines(true);
  var pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  var pivotGroup = pivotTable.addRowGroup(3);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup.showTotals(false);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup.showTotals(false);
  pivotGroup = pivotTable.addRowGroup(11);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup.showTotals(false)
  .showRepeatedLabels();
  pivotGroup = pivotTable.addRowGroup(11);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup.showTotals(false)
  .showRepeatedLabels();
  pivotGroup = pivotTable.addRowGroup(11);
  pivotGroup.showTotals(false);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  var pivotValue = pivotTable.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup.showTotals(false)
  .showRepeatedLabels();
  pivotGroup = pivotTable.addRowGroup(11);
  pivotGroup.showTotals(false);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotValue = pivotTable.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup.showTotals(false)
  .showRepeatedLabels();
  pivotGroup = pivotTable.addRowGroup(11);
  pivotGroup.showTotals(false);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotValue = pivotTable.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup.showTotals(false)
  .showRepeatedLabels();
  pivotGroup = pivotTable.addRowGroup(11);
  pivotGroup.showTotals(false);
};



function nbexa() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B2').activate();
};

function nbhm() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C2').activate();
  spreadsheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireNumberBetween(1, 6)
  .build());
  spreadsheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .setHelpText('Saisissez un nombre compris entre 1 et 6')
  .requireNumberBetween(1, 6)
  .build());
  spreadsheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .setHelpText('Saisissez un nombre d\'examinateurs entre 1 et 6')
  .requireNumberBetween(1, 6)
  .build());
  spreadsheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .setHelpText('Saisissez un nombre d\'examinateurs  entre 1 et 6')
  .requireNumberBetween(1, 6)
  .build());
};

function dureesi() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C3').activate();
  spreadsheet.getRange('C3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Option 1', 'Option 2'], true)
  .build());
  spreadsheet.getRange('C3').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInRange(spreadsheet.getRange('Durees_interrogation'), true)
  .build());
};

function hhmm() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C8').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('hh":"mm');
};

function coul1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B7:C19').activate();
  spreadsheet.getActiveRangeList().setBackground(null)
  .setBackground('#b6d7a8')
  .setBackground('#d9ead3');
  spreadsheet.getRange('F12').activate();
};

function rule() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G5').activate();
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Option 1', 'Option 2'], true)
  .build());
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireFormulaSatisfied()
  .build());
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireFormulaSatisfied()
  .build());
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireFormulaSatisfied()
  .build());
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireFormulaSatisfied()
  .build());
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireFormulaSatisfied()
  .build());
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireFormulaSatisfied()
  .build());
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireFormulaSatisfied()
  .build());
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireFormulaSatisfied()
  .build());
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireFormulaSatisfied()
  .build());
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireFormulaSatisfied()
  .build());
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireFormulaSatisfied()
  .build());
};

function titre() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K1').activate();
  
};

function coul2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A8:D8').activate();
  spreadsheet.getActiveRangeList().setFontSize(11)
  .setFontSize(12)
  .setFontSize(13)
  .setBackground('#f4cccc');
};

function centrehm() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D:D').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
};