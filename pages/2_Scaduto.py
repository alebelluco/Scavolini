import streamlit as st 
import pandas as pd 
import numpy as np
from io import BytesIO
import xlsxwriter

st.set_page_config(layout='wide')
st.title('Scaduto')

path = st.file_uploader('Caricare ZMM11')
if not path:
    st.stop()

path2 = st.file_uploader('ZSD67')
if not path2:
    st.stop()

zmm11 = pd.read_excel(path) # serve per il campo b2b
zsd67 = pd.read_excel(path2) # file principale

zmm11['key'] = [str(zmm11['Doc.acquisti'].iloc[i])+str(zmm11['Pos'].iloc[i]) for i in range(len(zmm11)) ]
zsd67['key'] = [str(zsd67['Numero'].iloc[i])+str(zsd67['Posizione'].iloc[i]) for i in range(len(zsd67))]
zsd67 = zsd67.merge(zmm11[['key','Qtà B2B']], how='left', left_on = 'key', right_on = 'key')
zsd67 = zsd67[zsd67['Qtà B2B']==0]



# unione colonna colore
# ----------------------------------

colonne_colore = [
    'C_COL1-Colore frontale',
    'C_COL2-Colore',
    'C_COLANTA-Colore anta / pannello esterno',
    'C_COLANTAINT-Colore anta / pannello inte',
    'C_COLTELAIO-Colore telaio',
    'C_COLMCM-Colore mensola/cornice',
    'C_COLPANN1-Colore mat.struttura faccia 1',
    'C_COLPANN2-Colore mat.struttura faccia 2',
    'C_COLPANNSCH-Colore Pannello',
    'C_COLSCHSPALLA-Colore schienale per spal',
    'C_COLRIPSPALLA-Colore ripiano spalla'
]

colore = []
for col in colonne_colore:
    try:
        zsd67[col]=zsd67[col].fillna('')
        colore.append(col)
    except:
        pass

st.write(colore)

zsd67['colore'] = None
for i in range(len(zsd67)):
    key_col = []
    for col in colore:
        key_col.append(str(zsd67[col].iloc[i]))
    zsd67['colore'].iloc[i] = ''.join(list(set(key_col))) #list-set-list serve per eliminare eventuali valori doppi popolati su due colonne


# Unione colonna finitura
# ------------------------------


colonne_finitura = [
    'C_FIN-Finitura',
    'C_FINMC-Finitura mensola cornice',
    'C_FINP1-Finitura pannello 1',
    'C_FINPANN-Finitura pannello',
    'C_FINPANNSCH-Finitura Pannello'
]

finitura = [] 
for col in colonne_finitura:
    try:
        zsd67[col]=zsd67[col].fillna('')
        finitura.append(col)
    except:
        pass

st.write(finitura)

zsd67['finitura'] = None
for i in range(len(zsd67)):
    key_fin = []
    for col in finitura:
        key_fin.append(str(zsd67[col].iloc[i]))
    zsd67['finitura'].iloc[i] = ''.join(list(set(key_fin))) #list-set-list serve per eliminare eventuali valori doppi popolati su due colonne


st.subheader('ZDS67')
st.dataframe(zsd67)







