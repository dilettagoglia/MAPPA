import pandas as pd
from params import list_of_tables
import win32com
import os
import sys
import comtypes.client

# FUNZIONI UTILI

def read_tables(path, tables):
    list_df = []
    for table in tables:
        df = pd.read_excel(path, table)
        list_df.append(df)
    #for (dfr, col) in zip(list_df, list_of_tables):
        #dfr.add_prefix(f"{str(col)}") # not working
        #dfr.columns = f'{str(col)}_' + dfr.columns.values # add prefix to cols
    return list_df

def docx_to_pdf(src, dst):
    #word = win32com.client.Dispatch("Word.Application")
    word = comtypes.client.CreateObject('Word.Application')
    wdFormatPDF = 17
    doc = word.Documents.Open(src)
    doc.SaveAs(dst, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

