
/**
* ajoute le bouton d'export dans la feuille de calcul
*/
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Rows to Docs",
    functionName : "exportMultiDocs"
  }];
  spreadsheet.addMenu("Export", entries);
}

/**
* parcours la feuille de calcul générée à partir des réponses à un questionnaire Google Form,
* Exporte chaque "ligne" ( participant ) dans un nouveau fichier "Texte" ( google Doc )
*/

function exportMultiDocs(){
  
  var SHARE_WITH_EDITORS = false;
  var FILENAME_PREFIX = "ExportRow_";
  
  var spread = SpreadsheetApp.getActiveSpreadsheet();
  
  // feuille active du tableur
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // recuperation de toutes les lignes
  var rows = sheet.getDataRange();
  
  // nombre de lignes
  var numRows = rows.getNumRows();
  
  // contenu des lignes
  var values = rows.getValues();

  // creation d'un dossier d'export
  var formattedDate = Utilities.formatDate(new Date(), "GMT + 2:00", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var dirName = "Export_" + spread.getName() + formattedDate;
  var dir = DriveApp.createFolder(dirName)
  
  // en tête des colonnes ( champs, questions... )
  var dataFields = values[0];
  // Logger.log("num questions : " + dataFields.length);
  
  // pour toutes les lignes sauf la première ( entêtes )
  for (var i = 1; i < numRows ; i++) {

    // recuperation de la ligne
    var row = values[i];

    // creation du fichier Doc
    var doc = DocumentApp.create( FILENAME_PREFIX + i );
    // recupère contenu ( vide ) éditable du fichier
    var body = doc.getBody();
    
    // insère chaque titre de colonne avec contenu saisi par participant
    for(var j=1; j < dataFields.length ; j++){
      var title = values[0][j];
      var rep = values[i][j];
      
      var titleParagraph = body.appendParagraph( title );
      titleParagraph.setBold(true);
      titleParagraph.setItalic(true);
      
      var repParagraph = body.appendParagraph( rep + "\n" );
      repParagraph.setBold(false);      
      repParagraph.setItalic(false);
    }
    
    // enregistrement et fermeture du document
    doc.saveAndClose();
    
    // déplacement d'une copie du fichier généré vers le dossier d'export
    var file = DriveApp.getFileById( doc.getId() );
    file.makeCopy( doc.getName(), dir);
    
    // suppression du fichier généré
    file.setTrashed(true);
  }  

  // partage le dossier généré avec tous les admins du tableurs
  if( SHARE_WITH_EDITORS ){
    
    // liste des admins de la feuille de calcul ( le dossier d'export  )
    var admins = spread.getEditors();
    
    for(var k = 1; k < admins.length ; k++){
      dir.addEditor(admins[k]);
    }
  } 
}
