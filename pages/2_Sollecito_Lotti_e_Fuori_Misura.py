
import streamlit as st 
import pandas as pd 
import numpy as np
from io import BytesIO
import xlsxwriter
from utils import dataprep as dp 
from datetime import datetime as dt

st.set_page_config(layout='wide')
st.title('Sollecito Lotti e Fuori Misura')

path = st.file_uploader('Caricare ZMM11')
if not path:
    st.stop()

path2 = st.file_uploader('ZSD67')
if not path2:
    st.stop()

zmm11 = pd.read_excel(path) # serve per il campo b2b
zsd67 = pd.read_excel(path2) # file principale

zmm11['key'] = [str(zmm11['Doc.acquisti'].iloc[i])+str(zmm11['Pos'].iloc[i]) for i in range(len(zmm11))]
zsd67['key'] = [str(zsd67['Numero'].iloc[i])+str(zsd67['Posizione'].iloc[i]) for i in range(len(zsd67))]
zsd67 = zsd67.merge(zmm11[['key','Qtà B2B']], how='left', left_on = 'key', right_on = 'key')
zsd67 = zsd67[zsd67['Qtà B2B']==0]

# Layout di esportazione

layout = {
    'output' : ['Materiale','Descrizione doc.','UM','Quantità','Numero',
                'Posizione','Tp.Doc','Data documento','Data consegna','Intestatario',
                'Numero OdV','Pos. OdV','Dt. consegna OdV','colore','finitura','altezza','larghezza','spessore','testo']
}

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

st.subheader('ZSD67')
st.dataframe(zsd67[layout['output']])

fornitori = list(zsd67['Intestatario'].unique())
st.subheader('Dowload file', divider = 'red')

# Download massivo

fornitori = list(zsd67['Intestatario'].unique())
df_dict = {}
i=0
for fornitore in fornitori:
    i+=1
    df_fil = zsd67[zsd67.Intestatario == fornitore][layout['output']]
    df_dict[f'{fornitore}.xlsx']= dp.create_excel_file(df_fil,f'{fornitore}.xlsx')
    
zip_data = dp.create_zip_file(df_dict)
st.download_button(
    label="Scarica file zip",
    data=zip_data,
    file_name='files.zip',
    mime='application/zip'
)


