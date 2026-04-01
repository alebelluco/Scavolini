import streamlit as st 
import pandas as pd 
import numpy as np
from io import BytesIO
import xlsxwriter
from utils import dataprep as dp 

st.set_page_config(layout='wide')
st.title('Prezzi fuori misura')

layout = { 
    'Work':['Numero','Posizione','Materiale','colore','superficie','key'],
    'Work2':['Numero','Posizione','Materiale','colore','superficie','VAL.MINIMO','PREZZO AL MQ','Soglia','prezzo'],
    'Output':['Numero','Posizione','prezzo'],
}




path2 = st.file_uploader('Caricare db prezzi')
if not path2:
    st.stop()

path = st.file_uploader('Caricare estrazione')
if not path:
    st.stop()    

df = pd.read_excel(path)
db_prezzi = pd.read_excel(path2)
config = pd.read_excel(path2, sheet_name='Configurazione')

config['key']=[str.split(campo, ',') for campo in config.Campi]

legame_chiavi = dict(zip(config['Ragione Sociale'], config.key))


db_prezzi['Mag_fissa']  = db_prezzi['Mag_fissa'].fillna(0)
db_prezzi['Mag_var']  = db_prezzi['Mag_var'].fillna(0)
db_prezzi['Soglia_mag']  = db_prezzi['Soglia_mag'].fillna(9999)

# Unione colonne
    
finitura = [
    'C_FIN-Finitura',
    'C_FINP1-Finitura pannello 1',
    'C_FINPANN-Finitura pannello',
    'C_FINANTACOMM-Finitura anta commerciale'
]

colore = [
    'C_COL1-Colore frontale',
    'C_COL2-Colore',
    'C_COLANTA-Colore anta / pannello esterno',
    'C_COLRIPSPALLA-Colore ripiano spalla',
    'C_COLANTAINT-Colore anta / pannello inte',
    'C_COLFASC-Colore Fascia',
    'C_COLTELAIO-Colore telaio'
]

altezza = [
    'C_HE-Altezza effettiva',
    'C_HEA-Altezza effettiva acquisto',
    'C_ALTE-Altezza effettiva'
]

larghezza = [
    'C_LE-Larghezza effettiva',
    'C_LEA-Larghezza effettiva acquisto',
    'C_LARGHE-Larghezza effettiva'

]

dp.unisci_colonne(df,colore,'colore')
dp.unisci_colonne(df,finitura,'finitura')
dp.unisci_colonne(df,altezza,'altezza')
dp.unisci_colonne(df,larghezza,'larghezza')

df['colore'] = df['colore'].fillna('').str.replace('ZZ_Non Definito','', regex=False)

# Eliminazione mm e modifica virgola in punto

df['altezza'] = df['altezza'].fillna('').str.replace(' mm','', regex=False)
df['larghezza'] = df['larghezza'].fillna('').str.replace(' mm','', regex=False)

df['altezza'] = df['altezza'].str.replace(',','.', regex=False)
df['larghezza'] = df['larghezza'].str.replace(',','.', regex=False)

df['altezza'] = pd.to_numeric(df['altezza'], errors='coerce')

try:
    df['larghezza'] = pd.to_numeric(df['larghezza'], errors='coerce')
except:
    pass

df['superficie'] = [(df['altezza'].iloc[i]*df['larghezza'].iloc[i])/1000000 for i in range(len(df))]

try:
    df['mat'] = df['C_MATP1-Materiale pannello 1']
except:
    pass

try:
    df['spessore'] = df['C_SPESSORE-Spessore']
    df['spessore'] =  df['spessore'].astype(str)
    df['spessore'] = df['spessore'].fillna('').str.replace(' mm','', regex=False)
    df['spessore'] = df['spessore'].str.replace(',','.', regex=False)
    df['spessore'] = pd.to_numeric(df['spessore'], errors='coerce')
    df['spessore'] = df['spessore'].astype('Int64')
except:
    pass

dp.crea_chiave(df,legame_chiavi)
db_prezzi['key']=[str(db_prezzi.Fornitore.iloc[i])+str(db_prezzi.CONCATENA.iloc[i]) for i in range(len(db_prezzi))]
df = df.merge(db_prezzi[['key','VAL.MINIMO','PREZZO AL MQ','Soglia','Mag_fissa','Mag_var','Soglia_mag']], how='left', left_on='key', right_on='key')


df['prezzo'] = np.where(df['superficie']<df['Soglia'],df['VAL.MINIMO']+df['Mag_fissa'],df['PREZZO AL MQ']*df['superficie']+df['Mag_fissa']).round(2)
df['prezzo'] = np.where((df['altezza']>df['Soglia_mag']),df['prezzo']+df['Mag_var'],df['prezzo'])
df['prezzo'] = np.where((df['larghezza']>df['Soglia_mag']),df['prezzo']+df['Mag_var'],df['prezzo'])
df['prezzo'] = df['prezzo'].astype(str)

df['prezzo'] = df['prezzo'].fillna('').str.replace('.',',', regex=False)

#st.write('df', df)
#df = df.drop(columns=['key'])

sx, dx = st.columns([1,1])

with sx:
    st.subheader('Output')
    st.dataframe(df[layout['Work2']])

with dx:
    st.subheader('Dati mancanti')
    mancanti = df[df.prezzo.astype(str) == 'nan']
    st.dataframe(mancanti[layout['Work2']])

dp.scarica_excel(df[layout['Output']],'output.xlsx')






