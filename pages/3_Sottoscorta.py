import streamlit as st 
import pandas as pd 
from utils import dataprep as dp 
from datetime import datetime as dt


st.set_page_config(layout='wide')
st.title('Sottoscorta')
st.write('Output: elenco delle righe suddivise per fornitore')


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
st.subheader('Dowload file excel', divider = 'red')

# propone un pulsante per ogni fornitore per scaricare il file excel
for forn in fornitori:
    name = f'{forn}.xlsx'
    df = zmm28[zmm28['Ragione sociale'] == forn]
    df = df[layout['output']]
    st.write(f'{forn}')
    dp.scarica_excel(df,name)
    st.divider()
