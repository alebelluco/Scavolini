
import streamlit as st
import pandas as pd
from utils import dataprep as dp

st.set_page_config(layout='wide')
st.title('MTO to MTS')

# ─────────────────────────────────────────────
# 1. CARICAMENTO FILE
# ─────────────────────────────────────────────

path_zsd67 = st.file_uploader('Caricare ZSD67')
if not path_zsd67:
    st.stop()

path_trans = st.file_uploader('Caricare Transcodifica')
if not path_trans:
    st.stop()

path_zmm28 = st.file_uploader('Caricare ZMM28')
if not path_zmm28:
    st.stop()

# ─────────────────────────────────────────────
# 2. ELABORAZIONE ZSD67
# ─────────────────────────────────────────────

@st.cache_data
def carica_zsd67(file):
    return pd.read_excel(file)

@st.cache_data
def carica_transcodifica(file):
    return file

@st.cache_data
def carica_zmm28(file):
    return pd.read_excel(file)

zsd67 = carica_zsd67(path_zsd67)
zmm28 = carica_zmm28(path_zmm28)

# Colonne da unire per ricavare colore e finitura (stessa logica delle altre pages)
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
    'C_COLRIPSPALLA-Colore ripiano spalla',
    'C_COLBORPTAV-Colore Bordo Tavolo',
    'C_COLCAPPACAMINF-Colore cappa camino inf',
    'C_COLCOPATT-Colore Coperchiet.Attaccagli',
    'C_COLPIANO-Colore piano',
    'C_COLPREZZO-Colore prezzo',
    'C_COLSTRU-Colore Struttura',
    'C_COLSUP1TAV-Colore Superficie 1 tavolo',
    'C_COLSUP2TAV-Colore Superficie 2 tavolo',
    'C_BOCCETTA-Colore boccetta'
]

colonne_finitura = [
    'C_FIN-Finitura',
    'C_FINMC-Finitura mensola cornice',
    'C_FINP1-Finitura pannello 1',
    'C_FINPANN-Finitura pannello',
    'C_FINPANNSCH-Finitura Pannello',
    'C_FINPIANO-Finitura piano',
    'C_ESTCOPRIF-Estetica coprifianco'
]

# Estrazione colore e finitura
dp.unisci_colonne(zsd67, colonne_colore, 'Colore')
dp.unisci_colonne(zsd67, colonne_finitura, 'Finitura')

# Pulizia valore non definito
zsd67['Colore'] = zsd67['Colore'].fillna('').str.replace('ZZ_Non Definito', '', regex=False)
zsd67['Finitura'] = zsd67['Finitura'].fillna('')

# ─────────────────────────────────────────────
# 3. FILTRO PER COLORE E FINITURA (dropdown separati)
# ─────────────────────────────────────────────

colori_disponibili = sorted(zsd67['Colore'].dropna().unique().tolist())
finiture_disponibili = sorted(zsd67['Finitura'].dropna().unique().tolist())

col1, col2 = st.columns(2)
with col1:
    colore_sel = st.multiselect('Filtra per Colore', options=colori_disponibili)
with col2:
    finitura_sel = st.multiselect('Filtra per Finitura', options=finiture_disponibili)

# Applicazione filtri (lista vuota = nessun filtro attivo)
df = zsd67.copy()
if colore_sel:
    df = df[df['Colore'].isin(colore_sel)]
if finitura_sel:
    df = df[df['Finitura'].isin(finitura_sel)]

# Mantenimento sole colonne necessarie di ZSD67
df = df[['Materiale', 'Descrizione mat.', 'Numero', 'Numero OdV', 'Pos. OdV', 'Quantità', 'Colore', 'Finitura']].copy()

# ─────────────────────────────────────────────
# 4. ELABORAZIONE TRANSCODIFICA
#    - Tutti i fogli vengono accorpati
#    - Il nome del foglio diventa la colonna "tipologia"
#    - Primo "Oggetto" (sx) = codice MTS, secondo = codice MTO
#      (pandas rinomina automaticamente i duplicati: "Oggetto" e "Oggetto.1")
# ─────────────────────────────────────────────

xls = pd.ExcelFile(path_trans)

# Leggo solo il foglio "Conversione Vecchio-Nuovo"
# header=1 perché la riga 0 contiene le etichette di sezione "MTS"/"MTO" (merged cells)
trans = pd.read_excel(xls, sheet_name='Conversione Vecchio-nuovo', header=1)

# Pandas rinomina automaticamente il secondo "Oggetto" in "Oggetto.1"
# Primo (sx) = codice MTS, secondo = codice MTO
trans = trans.rename(columns={'Oggetto': 'cod_MTS', 'Oggetto.1': 'cod_MTO'})
# ─────────────────────────────────────────────
# 5. MERGE ZSD67 + TRANSCODIFICA
#    Chiave: Materiale (ZSD67) = cod_MTO (transcodifica)
# ─────────────────────────────────────────────

df = df.merge(
    trans[['cod_MTS', 'cod_MTO']],
    how='left',
    left_on='Materiale',
    right_on='cod_MTO'
)

# ─────────────────────────────────────────────
# 6. MERGE CON ZMM28
#    Chiave: cod_MTS (dalla transcodifica) = Materiale (ZMM28)
# ─────────────────────────────────────────────

df = df.merge(
    zmm28[['Materiale', 'Qta. Stock', 'Imp. Totale']],
    how='left',
    left_on='cod_MTS',
    right_on='Materiale',
    suffixes=('', '_zmm28')
)

# Rimozione colonna Materiale duplicata proveniente da ZMM28
df = df.drop(columns=['Materiale_zmm28'], errors='ignore')

# ─────────────────────────────────────────────
# 7. OUTPUT
# ─────────────────────────────────────────────

colonne_output = ['Materiale', 'Descrizione mat.','Numero', 'Numero OdV', 'Pos. OdV', 'Quantità',
                  'Colore', 'Finitura', 'cod_MTS', 'Qta. Stock', 'Imp. Totale']

st.subheader('Risultato')
st.dataframe(df[colonne_output].dropna(subset=['Qta. Stock']), use_container_width=True)

st.subheader('Download', divider='red')
dp.scarica_excel(df[colonne_output].dropna(subset=['Qta. Stock']), 'MTO_to_MTS.xlsx')
