
# DATABASE IN LOCALE
database_name = 'MAPPA_database.xlsx' # nome ed estensione del file relativo al database
database_path = f'../database/{str(database_name)}' # percorso dove si trova il file

# DATABASE ONLINE
db_url='https://unipiit.sharepoint.com/:x:/s/PianteOfficinali-DB733/ERsCSmQ1rKVFo0h70WewmXMBjvrEjtFyyb66MyrvYGzWTA?e=mt4as5'

# lista delle tabelle del database
# (fogli nella cartella di lavoro Excel)
# l'ordine non è importante
list_of_tables = ['box_tossicologico', # new
                  'dato_EU_commission', # new
                  'nomenclatura_botanica', # ex 'specie_vegetale', ex 'pianta'
                  'dati_mercato',
                  'dato_bio_farmacologico',
                  'dato_etnobotanico',
                  'WHO_OMS',
                  'REACH',
                  'procedure_controllo_qualita',
                  'principi_attivi_markers',
                  'Ph_Eur',
                  'campione_riferimento', # ex 'orto'
                  'min_sal_ita',
                  'ISO',
                  'FUI',
                  'ESCOP',
                  'EMA_HMPC',
                  'EFSA',
                  'EFSA_2',
                  'KOME', # new
                  'IARC', # new
                  'HCM'] # new
