# -*- coding: utf-8 -*-
import pyodbc
import pandas as ps 
import pandas.io.sql as pssql
from ConfigParser import SafeConfigParser
import sys

try:
    arg_requete = sys.argv[1]
except:
    arg_requete = 'REQUETE'

parser = SafeConfigParser()
parser.read(arg_requete+'.ini')

script = parser.get('SQL2EXCEL', 'requete_sql') 
temp_folder = parser.get('SQL2EXCEL', 'temp_folder')
nom_fichier = parser.get('SQL2EXCEL', 'nom_export')  
nom_feuille = parser.get('SQL2EXCEL', 'nom_feuille') 
CNXN_SERVER = parser.get('INFO_CNXN', 'CNXN_SERVER')
CNXN_DATABASE = parser.get('INFO_CNXN', 'CNXN_DATABASE')
CNXN_TRUSTED = parser.get('INFO_CNXN', 'CNXN_TRUSTED')
USER_ID = parser.get('INFO_CNXN', 'USER_ID')
USER_PSWD = parser.get('INFO_CNXN', 'USER_PSWD')


if CNXN_TRUSTED == 'O':
    cnxn = pyodbc.connect(r'DRIVER={SQL Server};Server=CABANONS00006v;Database=Autofab;Trusted_Connection=yes;',unicode_results=True)
else:
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+CNXN_SERVER+';DATABASE='+CNXN_DATABASE+';UID='+USER_ID+';PWD='+USER_PSWD,unicode_results=True)
    
cursor = cnxn.cursor()

df = pssql.read_frame(script, cnxn)
writer = ps.ExcelWriter(temp_folder+'/'+nom_fichier+'.xlsx', encoding='iso-8859-1')
df.to_excel(writer, sheet_name=nom_feuille)
writer.save()
print ("Export complété")