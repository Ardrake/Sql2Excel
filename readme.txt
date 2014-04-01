Sql2Excel - Python 

Dumps Sql query results in an excel spreadsheet

Uses an ini file passed as an argument to the script for connexion info and query

by default uses the REQUETE.INI file.

______________________________________________________________________________________

Requires : pyodbc, pandas and ConfigParser
______________________________________________________________________________________

Ini file example

[SQL2EXCEL]
temp_folder = C:/Temp/     <---  Temp folder 
nom_export = fichier       <---  Spreadsheet filename
nom_feuille = page 1       <---  Spreadsheet sheet/page name

requete_sql = 	SELECT top 50 *        
		FROM CLIENT                    <--- Sql Query
		WHERE CLADR4 = 'Québec'
		
[INFO_CNXN]
CNXN_SERVER = SERVER      <--- Server address
CNXN_DATABASE = DATABASE  <--- Database name
USER_ID = USER            <--- User
USER_PSWD = PASS		  <--- Password



Author : André Cooke
Couriel : andrecooke@hotmail.com



