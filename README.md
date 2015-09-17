# appsscrx - Google apps scripts utils

Pour installer SpreadsheetToDoc

1. Dans la feuille de calcul : Ouvrir le menu "Modules complémentaires" > "Télécharger des Modules complémentaires"

2. Choisir "Pour Simplon.co" dans la liste déroulante`

3. Installer "SpreadsheetToDoc" ( bouton "+ Gratuit" )
4. Autoriser les permissions requises5

Après l'installation, recharger le document : Le menu "Modules complémentaires" de vos spreadsheets contiendra désormais un sous-menu "SpreadsheetToDoc".

Pour intégrer l'utilitaire dans une nouvelle feuille de calcul :
1. "Modules complémentaires" > "Gérer les Modules complémentaires"

2. "Gérer" > "Utiliser dans ce document"

3. Recharger le doc

4. "Modules complémentaires" > "SpreadsheetToDoc" > "Rows to Docs"
5. l'export crée un dossier portant le nom du spreadsheet, et contenant les fichiers doc générés. 

Ttrois scripts sont dispo : 
- un export standard indépendant de l'ordre des colonnes
- un export spécial candidature, qui récupère des données précises à partir de l'index de certaines colonnes ( prénom, nom, mail, téléphone ). Le nom et le prénom sont notamment utilisés pour le nom du doc.
- un calcul de la moyenne d'âge des candidats ( le résultat est visible dans les logs )
