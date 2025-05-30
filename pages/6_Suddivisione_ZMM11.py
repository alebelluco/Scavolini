import streamlit as st 
import pandas as pd 
import numpy as np
from io import BytesIO
import xlsxwriter
from utils import dataprep as dp 
from datetime import datetime as dt

st.set_page_config(layout='wide')
st.title('Suddivisione fornitori ZMM11')

path = st.file_uploader('Caricare ZMM11')
if not path:
    st.stop()

zmm11 = pd.read_excel(path) 
mto = False
if st.checkbox('MTO'):
    mto=True
    layout = ["Buyer",
        "Fornitore",
        "Ragione sociale",
        "Tp doc.",
        "Doc.acquisti",
        "Pos",
        "Materiale",
        "Definizione",
        "Articolo fornitore",
        "UM",
        "N° Ordine",
        "Posizione",
        "Data ordine",
        "Qtà ordine",
        "Qtà cons.",
        "Qtà residua",
        "Qtà B2B",
        "Data consegna",
        "Data Cons."]
else:
    layout = ["Buyer",
        "Fornitore",
        "Ragione sociale",
        "Tp doc.",
        "Doc.acquisti",
        "Pos",
        "Materiale",
        "Definizione",
        "Articolo fornitore",
        #"C_MODACQ",
        #"Kanban",
        "Data ordine",
        "Qtà ordine",
        "Qtà cons.",
        "Qtà residua",
        "Qtà B2B",
        "Data consegna"]


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

zmm11['Data ordine'] = [dt.strftime(data, format='%d/%m/%Y') for data in zmm11['Data ordine']]
zmm11['Data consegna'] = [dt.strftime(data, format='%d/%m/%Y') for data in zmm11['Data consegna']]
if mto:
   zmm11['Data Cons.'] = [dt.strftime(data, format='%d/%m/%Y') for data in zmm11['Data Cons.'] if len(str(data))>5]
            


zmm11
now = str((dt.now().date())).replace('-','')
# Download massivo

fornitori = list(zmm11['Ragione sociale'].unique())
df_dict = {}
i=0
for fornitore in fornitori:
    i+=1
    df_fil = zmm11[zmm11['Ragione sociale'] == fornitore][layout]
    df_dict[f'{fornitore}.xlsx']= dp.create_excel_file(df_fil,f'{fornitore}.xlsx')
    
zip_data = dp.create_zip_file(df_dict)
st.download_button(
    label="Scarica file zip",
    data=zip_data,
    file_name=f'{now}_ZMM11.zip',
    mime='application/zip'
)
