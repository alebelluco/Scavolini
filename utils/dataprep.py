# raggruppamento utility
# 26-06-2024


import pandas as pd 
import streamlit as st 
from io import BytesIO
import xlsxwriter


def unisci_colonne(df, colonne, new):
    new_col = [] 
    for col in colonne:
        try:
            df[col]=df[col].fillna('')
            new_col.append(col)
        except:
            pass
  
    df[new] = None
    for i in range(len(df)):
        key = []
        for col in new_col:
            key.append(str(df[col].iloc[i]))
        df[new].iloc[i] = ''.join(list(set(key))) #list-set-list serve per eliminare eventuali valori doppi popolati su due colonne

    return df


def crea_chiave(df, dic_key):
    df['key']=None
    for i in range(len(df)):
        fornitore = df['Intestatario'].iloc[i]
        colonne = dic_key[fornitore]
        key = []
        for col in colonne:
            key.append(str(df[col].iloc[i]))
        df['key'].iloc[i] = ''.join((key))
    return df


def scarica_excel(df, filename):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1',index=False)
    writer.close()
    st.download_button(
        label="Download Excel workbook",
        data=output.getvalue(),
        file_name=filename,
        mime="application/vnd.ms-excel"
    )





