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


legame_chiavi ={
    "400020 TERENZI VITTORIO &C.SNC DI T.S":['Intestatario','Materiale','colore'],
    "400089 FAB SRL":['Intestatario','Materiale','colore'],
    "400642 M.B.F. S.R.L.":['Intestatario','Materiale','colore','spessore'], 
    "400592 BECA GROUP S.R.L.":['Intestatario','Materiale'],
    "400545 VETROTEC SRL":['Intestatario','Materiale'],
    "400763 ARTI DEL VETRO SRL":['Intestatario','Materiale','mat'], 
    "400516 PANTAREI SRL":['Intestatario','Materiale'],
    "400844 STIVAL SRL":['Intestatario','Materiale'],
    "400624 FULIGNA & SENSOLI S.R.L.":['Intestatario','Materiale'],
    "400817 PESARO GLASS S.R.L.":['Intestatario','Materiale'],
    "400058 SCILM SPA":['Intestatario','Materiale'],
    "400782 ERREBIELLE COMPONENTS SRL":['Intestatario','Materiale','colore'],
    "400789 L.G. S.R.L.":['Intestatario','Materiale','colore','finitura'], # la colonna finitura va aggiunta
    "400525 VETR. ARTIST.ARTIGIANA GLASS DI LUC":['Intestatario','Materiale','colore','finitura'],
    "400510 G. & D. S.P.A.":['Intestatario','Materiale','colore','finitura'],
    "400423 ATLANTIS SRL":['Intestatario','Materiale','colore','finitura'],
    "400109 DMM SPA":['Intestatario','Materiale'], # non era disponibile, utilizzato il valore più frequente
    "400119 VITEMPER SRL":['Intestatario','Materiale']
}



path2 = st.file_uploader('Caricare db prezzi')
if not path2:
    st.stop()

path = st.file_uploader('Caricare estrazione')
if not path:
    st.stop()    

df = pd.read_excel(path)
db_prezzi = pd.read_excel(path2)

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

df['colore'] = [str.replace(df['colore'].iloc[i],'ZZ_Non Definito','') for i in range(len(df))]

# Eliminazione mm e modifica virgola in punto

df['altezza'] = [str.replace(df['altezza'].iloc[i],' mm','') for i in range(len(df))]
df['larghezza'] = [str.replace(df['larghezza'].iloc[i],' mm','') for i in range(len(df))]

df['altezza'] = [str.replace(df['altezza'].iloc[i],',','.') for i in range(len(df))]
df['larghezza'] = [str.replace(df['larghezza'].iloc[i],',','.') for i in range(len(df))]

df['altezza'] = df['altezza'].astype(float)

try:
    df['larghezza'] = df['larghezza'].astype(float)
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
    df['spessore'] = [str.replace(df['spessore'].iloc[i],' mm','') for i in range(len(df))]
    df['spessore'] = [str.replace(df['spessore'].iloc[i],',','.') for i in range(len(df))]
    df['spessore'] = df['spessore'].astype(float)
    df['spessore'] = df['spessore'].astype(int)
except:
    pass

dp.crea_chiave(df,legame_chiavi)
db_prezzi['key']=[str(db_prezzi.Fornitore.iloc[i])+str(db_prezzi.CONCATENA.iloc[i]) for i in range(len(db_prezzi))]
df = df.merge(db_prezzi[['key','VAL.MINIMO','PREZZO AL MQ','Soglia','Mag_fissa','Mag_var','Soglia_mag']], how='left', left_on='key', right_on='key')


df['prezzo'] = np.where(df['superficie']<df['Soglia'],df['VAL.MINIMO']+df['Mag_fissa'],df['PREZZO AL MQ']*df['superficie']+df['Mag_fissa']).round(2)
df['prezzo'] = np.where((df['altezza']>df['Soglia_mag']),df['prezzo']+df['Mag_var'],df['prezzo'])
df['prezzo'] = np.where((df['larghezza']>df['Soglia_mag']),df['prezzo']+df['Mag_var'],df['prezzo'])
df['prezzo'] = df['prezzo'].astype(str)

df['prezzo'] = [str.replace(df['prezzo'].iloc[i],'.',',') for i in range(len(df))]

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



