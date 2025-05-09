import streamlit as st 
import pandas as pd 
import numpy as np
from io import BytesIO
import xlsxwriter
from utils import dataprep as dp 
from datetime import datetime as dt


st.set_page_config(layout='wide')
st.title('Cambio fornitori ordini a commessa')

path = st.file_uploader('Caricare ZSD67')
if not path:
    st.stop()

zsd67 = pd.read_excel(path)

flat = pd.read_excel('/Users/Alessandro/Documents/AB/Clienti/ADI!/Scavolini/Acquisti/Estrazione/Luca Bozzi/ordine_commesse_new/flat.xlsx')
flat
codici_speciali = list(flat['Articolo'])

layout = {
    'output' : ['Materiale','Descrizione mat.','UM','Quantità','Numero',
                'Posizione','Tp.Doc','Data documento','Data consegna','Intestatario',
                'Numero OdV','Pos. OdV','Dt. consegna OdV','colore','finitura','altezza','larghezza','spessore','testo'],
    'output2' : ['Materiale','Descrizione mat.','UM','Quantità','Numero',
                'Posizione','Tp.Doc','Data documento','Data consegna','Intestatario',
                'Numero OdV','Pos. OdV','Dt. consegna OdV','colore','finitura','altezza','larghezza','spessore','testo','Fornitore']
}

# unione colonne

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
    'C_COLPREZZO-Colore prezzo'
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
    #'C_FINANTACOMM-Finitura anta commerciale',
    'C_FINPIANO-Finitura piano',
    'C_ESTCOPRIF-Estetica coprifianco'
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

spessore = [
    'C_SPESSCH-Spessore schienale',
    'C_SPESSORE-Spessore'
]

note_testo = [
    'C_NOTATESTO1-Nota testo 1',
    'C_NOTATESTO2-Nota testo 2',
    'C_NOTATESTO3-Nota testo 3',
    'C_NOTATESTO4-Nota testo 4'
]


dp.unisci_colonne(zsd67,colonne_colore,'colore')
dp.unisci_colonne(zsd67,colonne_finitura,'finitura')
dp.unisci_colonne(zsd67,altezza,'altezza')
dp.unisci_colonne(zsd67,larghezza,'larghezza')
dp.unisci_colonne(zsd67,spessore, 'spessore')
dp.unisci_colonne(zsd67,note_testo,'testo_appoggio')

zsd67['colore'] = [str.replace(zsd67['colore'].iloc[i],'ZZ_Non Definito','') for i in range(len(zsd67))]

# aggiustamento formati data
zsd67['Data documento'] = [dt.date(zsd67['Data documento'].iloc[i]).strftime("%d-%m-%Y") for i in range(len(zsd67))]
zsd67['Data consegna'] = [dt.date(zsd67['Data consegna'].iloc[i]).strftime("%d-%m-%Y") for i in range(len(zsd67))]
zsd67['Dt. consegna OdV'] = [dt.date(zsd67['Dt. consegna OdV'].iloc[i]).strftime("%d-%m-%Y") for i in range(len(zsd67))]

# aggiunta colonna testo dove colore mancante
zsd67['testo'] = np.where(zsd67['colore'].astype(str)=='', zsd67['testo_appoggio'],'')

spec_work = zsd67[[any(codice == test for codice in codici_speciali) for test in zsd67.Materiale.astype(str)]]
odv_spec =list(set(spec_work['Numero OdV']))

#spec_work
fornitori_change = spec_work[['Numero OdV','Intestatario']].drop_duplicates()
fornitori_change = fornitori_change.rename(columns={'Intestatario':'Fornitore'})

#st.write(codici_speciali)
with st.expander('OdV con codici speciali'):
    #st.write(spec_work)
    st.write((odv_spec))

st.subheader('Output: OdV con fornitori aggiornati')
st.subheader()

spec_all = zsd67[[any(odv == test for odv in odv_spec) for test in zsd67['Numero OdV']]]
spec_all = spec_all.merge(fornitori_change, how='left', left_on='Numero OdV', right_on='Numero OdV')
spec_all = spec_all[layout['output2']]

st.write(spec_all)

dp.scarica_excel(spec_all[layout['output2']],'Ordine_commesse_speciali.xlsx')