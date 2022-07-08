'''
This file extract all worksheets (tables) from Excel workbook (database)
'''

import pandas as pd
from utilities import read_tables
from params import *

proc = 'Process 1'
print(f'{str(proc)}: read database and extract each table.')

# list of tables to extract
sheet = list_of_tables # in 'params' file

# sostituire 'database_path' con 'socket_' se si vuole usare il database online
#socket_ = urllib.request.urlopen(db_url)
xls = pd.ExcelFile(database_path)

# extraction
extracted_tables = read_tables(xls, sheet)

print('Successfully extracted', len(extracted_tables), 'tables from database.')

print(f'{str(proc)} successfully terminated.')