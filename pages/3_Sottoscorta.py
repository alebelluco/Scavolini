import streamlit as st 
import pandas as pd 
from utils import dataprep as dp 
from datetime import datetime as dt


st.set_page_config(layout='wide')
st.title('Sottoscorta')

layout = {
    'output':['Ragione sociale','Materiale','Definizione',
              'OdA Passato','OdA a 1mese','OdA a 2mesi',
              'OdA a 3mesi','OdA oltre','OdA Totale',
              'Avv.B2B','Data rischedulazione x forn.',
              'Altezza effettiva','Larghezza effettiva','Spessore']
}

path = st.file_uploader('Caricare ZMM28')
if not path:
    st.stop()

zmm28 = pd.read_excel(path)
zmm28 = zmm28[zmm28['CTL Stock']=='X']
zmm28 = zmm28[zmm28['Kanban']!='X']
zmm28 = zmm28[zmm28['Approv.']!= 'X']

# aggiustamento formato data
for i in range(len(zmm28)):
    if str((zmm28['Data rischedulazione x forn.'].iloc[i])) != 'NaT':
        zmm28['Data rischedulazione x forn.'].iloc[i] = dt.date(zmm28['Data rischedulazione x forn.'].iloc[i])
    else:
        zmm28['Data rischedulazione x forn.'].iloc[i] = ''

st.dataframe(zmm28[layout['output']])

fornitori = list(zmm28['Ragione sociale'].unique())
st.subheader('Dowload file', divider = 'red')

df_dict = {}
i=0
for fornitore in fornitori:
    i+=1
    df_fil = zmm28[zmm28['Ragione sociale'] == fornitore][layout['output']]
    df_dict[f'{fornitore}.xlsx']= dp.create_excel_file(df_fil,f'{fornitore}.xlsx')
    
zip_data = dp.create_zip_file(df_dict)
st.download_button(
    label="Scarica file zip",
    data=zip_data,
    file_name='files.zip',
    mime='application/zip'
)
