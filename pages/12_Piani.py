import streamlit as st
import pandas as pd
from utils import dataprep as dp

st.set_page_config(layout='wide')
st.title('Elaborazione piani')

# ─────────────────────────────────────────────
# 1. CARICAMENTO FILE
# ─────────────────────────────────────────────

path_zsd67 = st.file_uploader('Caricare ZSD67')
if not path_zsd67:
    st.stop()


# ─────────────────────────────────────────────
# 2. ELABORAZIONE ZSD67
# ─────────────────────────────────────────────

@st.cache_data
def carica_zsd67(file):
    return pd.read_excel(file)


zsd67 = carica_zsd67(path_zsd67)


colonne_keep = [
    'Materiale',
    'Descrizione mat.',
    'UM',
    'Quantità',
    'Valore netto',
    'Numero',
    'Posizione',
    'Tp.Doc',
    'Data documento',
    'Data consegna',
    'Intestatario',
    'Numero OdV',
    'Pos. OdV',
    'Dt. produzione',
    'C_COLALZ-Colore alzatina',
    'C_COLMUR-Colore muretto',
    'C_COLPANNSCH-Colore Pannello',
    'C_COLPIANO-Colore piano',
    'C_COLSUP1TAV-Colore Superficie 1 tavolo',
    'C_EL1MONT-Elettrodomestico montato',
    'C_EL2MONT-Secondo Elettrodom. montato',
    'C_FINMUR-Finitura muretto',
    'C_FINPANNSCH-Finitura Pannello',
    'C_FINPIANO-Finitura piano',
    'C_FINPTAV-Finitura Piano Tavolo',
    'C_LAVPB03-Addebito lavabo integrato',
    'C_MATPROFALZ-Materiale profilo alzatina',
    'C_MODPTAV-Modello Piano Tavolo',
    'C_PROFPIANO-Profilo piano',
    'C_SPESSCH-Spessore schienale'
]

# Colonne da unire per ricavare colore
colonne_colore = [
    'C_COLALZ-Colore alzatina',
    'C_COLMUR-Colore muretto',
    'C_COLPANNSCH-Colore Pannello',
    'C_COLPIANO-Colore piano',
    'C_COLSUP1TAV-Colore Superficie 1 tavolo',
]


colonne_finitura = [
    'C_FINMUR-Finitura muretto',
    'C_FINPANNSCH-Finitura Pannello',
    'C_FINPIANO-Finitura piano',
    'C_FINPTAV-Finitura Piano Tavolo'
]









zsd67 = zsd67[colonne_keep].copy()

# Pulizia colonne colore: rimuove "ZZ_Non Definito" e spazi prima di unire
for c in colonne_colore:
    zsd67[c] = (
        zsd67[c]
        .fillna('')
        .astype(str)
        .str.strip()
        .replace({'nan': '', 'None': '', 'ZZ_Non Definito': ''})
    )


# Costruisce la colonna Colore unendo i valori non vuoti (senza duplicati)
zsd67['Colore'] = zsd67[colonne_colore].apply(
    lambda r: ' | '.join(dict.fromkeys([v for v in r.tolist() if v])),
    axis=1
)


for f in colonne_finitura:
    zsd67[f] = (
        zsd67[f]
        .fillna('')
        .astype(str)
        .str.strip()
        .replace({'nan': '', 'None': ''})
    )   

zsd67['Finitura'] = zsd67[colonne_finitura].apply(
    lambda r: ' | '.join(dict.fromkeys([v for v in r.tolist() if v])),
    axis=1
)   




zsd67 = zsd67.drop(columns=colonne_colore + colonne_finitura)

ordine_col=[
'Materiale',
'Descrizione mat.',
'UM',
'Quantità',
'Valore netto',
'Tp.Doc',
'Numero',
'Posizione',
'Data documento',
'Data consegna',
'Numero OdV',
'Pos. OdV',
'Dt. produzione',
'Colore',
'Finitura',
'C_PROFPIANO-Profilo piano',
'C_MATPROFALZ-Materiale profilo alzatina',
'C_SPESSCH-Spessore schienale',
'C_MODPTAV-Modello Piano Tavolo',
'C_LAVPB03-Addebito lavabo integrato',
'C_EL1MONT-Elettrodomestico montato',
'C_EL2MONT-Secondo Elettrodom. montato',
'Intestatario'
]

zsd67 = zsd67[ordine_col].copy()



st.dataframe(zsd67)
dp.scarica_excel(zsd67, 'zsd67_elaborato.xlsx')
