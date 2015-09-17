function ageAverage(){
  var BIRTHDAY_INDEX = 3;
  
  var spread = SpreadsheetApp.getActiveSpreadsheet();
  
  // feuille active du tableur
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // recuperation de toutes les lignes
  var rows = sheet.getDataRange();
  
  // nombre de lignes
  var numRows = rows.getNumRows();
  
  // contenu des lignes
  var values = rows.getValues();
  var total = 0;
  // pour toutes les lignes sauf la première ( entêtes )
  for (var i = 1; i < numRows ; i++) {

    // recuperation de la ligne
    var row = values[i];
    var bd = values[i][BIRTHDAY_INDEX]
    var today = new Date()

    var age = (today.getFullYear() - bd.getFullYear());
    Logger.log("age : " + age );
    total += age;

  }
    Logger.log("moyenne : " + total / (numRows - 1) );
}