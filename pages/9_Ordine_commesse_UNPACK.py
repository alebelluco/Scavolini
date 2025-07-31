# versione aggiornata il 31/07/2025
# aggiunta la possibilità di scaricare l'output suddiviso per fornitore
# inserite regole per LG

import streamlit as st 
import pandas as pd 
import numpy as np
from io import BytesIO
import xlsxwriter
from utils import dataprep as dp 
from datetime import datetime as dt


st.set_page_config(layout='wide')
st.title('Ordine Commesse')

path = st.file_uploader('Caricare ZSD67')
if not path:
    st.stop()

zsd67 = pd.read_excel(path)

layout = {
    'output' : ['Materiale','Descrizione mat.','UM','Quantità','Numero',
                'Posizione','Tp.Doc','Data documento','Data consegna','Intestatario',
                'Numero OdV','Pos. OdV','Dt. consegna OdV','colore','finitura','altezza','larghezza','spessore','testo','T'],
    'output2' : ['categoria','Materiale','Descrizione mat.','UM','Quantità','Numero',
                'Posizione','Tp.Doc','Data documento','Data consegna','Intestatario',
                'Numero OdV','Pos. OdV','Dt. consegna OdV','colore','finitura','altezza','larghezza','spessore','testo','T'],
    'output_gd' : ['Materiale','Descrizione mat.','UM','Quantità','Numero',
                'Posizione','Tp.Doc','Data documento','Data consegna','Intestatario',
                'Numero OdV','Pos. OdV','Dt. consegna OdV','colore','finitura','altezza','larghezza','spessore','testo'],

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
    'C_BOCCETTA-Colore boccetta',
    'C_COLFASC-Colore Fascia'
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



# unpack GD



def dividi_categorie_gd(zsd67):
    struttura = ['FIA','DIV','SCH','RIP','CIE','FON','ZOC']
    zsd67['categoria']=None
    for i in range(len(zsd67)):

        testo = zsd67['Descrizione mat.'].iloc[i]
        codice = zsd67['Materiale'].iloc[i]

        # prima condizione: fuori misura
        
        if zsd67['Tp.Doc'].iloc[i] == 'ZLAC':
            # ante: descrizione inizia con FR
            if testo[:2]=='FR':
                zsd67.categoria.iloc[i] = 'Ante'
            elif any(testo[:3] == voce for voce in struttura) and (str(codice)[:1]=='2') :
                zsd67.categoria.iloc[i] = 'Fianchi + struttura'
            elif ('GIO' in testo) and (str(codice)[:3]!='211'):
                zsd67.categoria.iloc[i] = 'Elementi struttura pensile giorno'
            elif ('GIO' in testo) and (str(codice)[:3]=='211'):
                zsd67.categoria.iloc[i] = 'Pensili giorno'
            elif testo[:3]=='BOC':
                zsd67.categoria.iloc[i] = 'Boccette'
            elif testo[:3]=='COP':
                zsd67.categoria.iloc[i] = 'Coprifianchi'
            elif testo[:3]=='TAP':
                zsd67.categoria.iloc[i] = 'Pensili giorno'
            elif (testo[:3]=='SCH') and  (str(codice)[:1]=='7'):
                zsd67.categoria.iloc[i] = 'Schienali'
            elif (zsd67['C/lav'].iloc[i]=='L') and (str(codice)[:3]=='205'):
                zsd67.categoria.iloc[i] = 'Mensole contolavoro'
            elif (testo[:4] == 'MENS') and (zsd67['C/lav'].iloc[i]!='L') :
                zsd67.categoria.iloc[i] = 'Mensole'
            elif (zsd67['C/lav'].iloc[i]=='L') and (str(codice)[:3]=='203'):
                zsd67.categoria.iloc[i] = 'Schienali contolavoro'
            elif testo[:3]=='PAN':
                zsd67.categoria.iloc[i] = 'Pannelli'

   
        else: # MTO
            if testo[:2]=='FR':
                zsd67.categoria.iloc[i] = 'Ante fuori misura'
            elif any(testo[:3] == voce for voce in struttura) and (str(codice)[:1]=='2'):
                zsd67.categoria.iloc[i] = 'Fianchi + struttura fuori misura'
            elif ('GIO' in testo) and (str(codice)[:3]!='211'):
                zsd67.categoria.iloc[i] = 'Elementi struttura pensile giorno fuori misura'
            elif (zsd67['C/lav'].iloc[i]=='L') and (str(codice)[:3]=='205'):
                zsd67.categoria.iloc[i] = 'Mensole contolavoro fuori misura'
            elif (zsd67['C/lav'].iloc[i]=='L') and (str(codice)[:3]=='203'):
                zsd67.categoria.iloc[i] = 'Schienali contolavoro fuori misura'
            elif ('GIO' in testo) and (str(codice)[:3]=='211') :
                zsd67.categoria.iloc[i] = 'Pensili giorno fuori misura'
            elif testo[:3]=='PAN':
                zsd67.categoria.iloc[i] = 'Pannelli fuori misura'
            elif testo[:3]=='COP':
                zsd67.categoria.iloc[i] = 'Coprifianchi fuori misura'
            elif testo[:3]=='TAP':
                zsd67.categoria.iloc[i] = 'Pensili giorno fuori misura'
            elif (testo[:3]=='SCH') and  (str(codice)[:1]=='7'):
                zsd67.categoria.iloc[i] = 'Schienali fuori misura'
    return zsd67


codici_carrellino = [
    '20388051',
    '20388052',
    '20388053',
    '20388054',
    '20388055',
    '20388056',
    '20388057'
]

def dividi_categorie_lg(zsd67, codici_carrellino):

    struttura = ['FIA','DIV','SCH','RIP','CIE','FON','ZOC']

    zsd67['categoria']=None
    for i in range(len(zsd67)):

        testo = zsd67['Descrizione mat.'].iloc[i]
        codice = zsd67['Materiale'].iloc[i]

        # prima condizione: fuori misura
        
        if zsd67['Tp.Doc'].iloc[i] == 'ZLAC':

            if testo[:3]=='BOC':
                zsd67.categoria.iloc[i] = 'Boccette'

        
            if codice[:3] == '211' or testo[:3]=='TAP':
                zsd67.categoria.iloc[i] = 'Pensili giorno'

            if any(testo[:3] == voce for voce in struttura) and (str(codice)[:1]=='2') :
                zsd67.categoria.iloc[i] = 'Fianchi + struttura'

            if (testo[:2]=='FR' or testo[:3]=='FAS' or testo[:3]=='COP') and zsd67['C/lav'].iloc[i] != 'L' and 'RIGAT' not in testo:
                zsd67.categoria.iloc[i] = 'Ante'

            if 'RIGAT' in testo:
                zsd67.categoria.iloc[i] = 'Dogato'

            if ('JMT' in testo) and zsd67['C/lav'].iloc[i] == 'L':

                zsd67.categoria.iloc[i] = 'Jeometrica contolavoro'

            if ('JMT' not in testo) and zsd67['C/lav'].iloc[i] == 'L':

                zsd67.categoria.iloc[i] = 'Contolavoro'

            if (((testo[:3]=='SCH' or testo[:3]=='PIA') and  (str(codice)[:1]=='7')) or str(codice)[:1]=='7' or testo[:3]=='PAN') and 'DOG' not in testo:
                zsd67.categoria.iloc[i] = 'Schienali Piani e Pannelli'
            
            if (((testo[:3]=='SCH' or testo[:3]=='PIA') and  (str(codice)[:1]=='7')) or str(codice)[:1]=='7' or testo[:3]=='PAN') and 'DOG' in testo:
                zsd67.categoria.iloc[i] = 'Schienali Piani e Pannelli dogati'

            if ('GIO' in testo) and ('GRIGIO' not in testo) and ('SOGGIORNO' not in testo) and (str(codice)[:3]!='211'):
                zsd67.categoria.iloc[i] = 'Elementi struttura pensile giorno'

            if ('GIO' in testo) and ('GRIGIO' not in testo) and ('SOGGIORNO' not in testo) and (str(codice)[:3]=='211'):
                zsd67.categoria.iloc[i] = 'Pensili giorno'

            #if (testo[:3]=='SCH' or testo[:3]=='PIA') and  (str(codice)[:1]=='7'):
                #zsd67.categoria.iloc[i] = 'Schienali e Piani'
            if any( cod == codice for cod in codici_carrellino):
                            zsd67.categoria.iloc[i] = 'Carrellino'

            pass


            
        elif zsd67['Tp.Doc'].iloc[i] == 'ZMTO': # MTO

            if codice[:3] == '211' or testo[:3]=='TAP':
                zsd67.categoria.iloc[i] = 'Pensili giorno fuori misura'

            if any(testo[:3] == voce for voce in struttura) and (str(codice)[:1]=='2'):
                zsd67.categoria.iloc[i] = 'Fianchi + struttura fuori misura'

            if (testo[:2]=='FR' or testo[:3]=='PAN' or testo[:3]=='FAS' or testo[:3]=='COP') and zsd67['C/lav'].iloc[i] != 'L' and 'RIGAT' not in testo :
                zsd67.categoria.iloc[i] = 'Ante e pannelli fuori misura'

            if 'RIGAT' in testo:
                zsd67.categoria.iloc[i] = 'Dogato'

            if (testo[:3]=='SCH' or testo[:3]=='PIA') and  (str(codice)[:1]=='7'):
                zsd67.categoria.iloc[i] = 'Schienali e Piani'

            if (((testo[:3]=='SCH' or testo[:3]=='PIA') and  (str(codice)[:1]=='7')) or str(codice)[:1]=='7' or testo[:3]=='PAN') and 'DOG' not in testo:
                zsd67.categoria.iloc[i] = 'Schienali Piani e Pannelli fuori misura'
            
            if (((testo[:3]=='SCH' or testo[:3]=='PIA') and  (str(codice)[:1]=='7')) or str(codice)[:1]=='7' or testo[:3]=='PAN') and 'DOG' in testo:
                zsd67.categoria.iloc[i] = 'Schienali Piani e Pannelli fuori misura dogati'

            if ('GIO' in testo) and ('GRIGIO' not in testo) and ('SOGGIORNO' not in testo) and (str(codice)[:3]!='211'):
                zsd67.categoria.iloc[i] = 'Elementi struttura pensile giorno fuori misura'

            if ('GIO' in testo) and ('GRIGIO' not in testo) and ('SOGGIORNO' not in testo) and (str(codice)[:3]=='211'):
                zsd67.categoria.iloc[i] = 'Pensili giorno fuori misura'


            pass

        elif zsd67['Tp.Doc'] == 'ZCLA':

            if ('JMT' in testo) and zsd67['C/lav'].iloc[i] == 'L':

                zsd67.categoria.iloc[i] = 'Jeometrica contolavoro'

            if ('JMT' not in testo) and zsd67['C/lav'].iloc[i] == 'L':

                zsd67.categoria.iloc[i] = 'Contolavoro'

        else:
            zsd67.categotia.iloc[i] = 'Just in time'


    return zsd67


# RUN divisione categorie ============================================================

if st.radio('Fornitore', options=['LG','G&D']) == 'G&D':
    st.subheader(':red[STAI UTILIZZANDO LE IMPOSTAZIONI PER G&D]')
    zsd67 = dividi_categorie_gd(zsd67)
    
    mobil_giorno_smontati =["20395057","20395058","20395059","20395060","20395061","20395062","20395063","20395064","20395065","20395066","20395067",
                            "20395068","20395069","20395070","20395071","20395072","20395073","20395074","20395075","20395076","20395077","20395078",
                            "20395079","20395080","20395081","20395082","20395083","20395084","20395085","20395086","20395087","20395088","20395089",
                            "20395090","20395091","20395092","20395093","20395094","20395095","20395096","20395097","20395098","20395099","20395100",
                            "20395101","20395102","20395103","20395142","20395143","20395144","20395145","20395146","20395147","20395148","20395149",
                            "20395150","20395151","20395152","20395153","20395154","20395155","20395156","20395157","20395158","20395159","20395160",
                            "20395161","20395162","20395163","20395164","20395165","20395166","20395167","20395168","20395169","20395170","20395171",
                            "20395172","20395173","20395174","20395175","20395176","20395177","20395178","20395179","20395180","20395181","20395182",
                            "20395183","20395184","20395185","20395186","20395187","20395188","20395104","20395105","20395106","20395107","20395108",
                            "20395109","20395110","20395111","20395112","20395113","20395114","20395115","20395116","20395117","20395118","20395119",
                            "20395120","20395121","20395122","20395123","20395124","20395125","20395126","20395127","20395129","20395130","20395131",
                            "20395132","20395133","20395134","20395135","20395136","20395137","20395138","20395139","20395140","20395141","20396130",
                            "20398848","20395189","20395190","20395191","20395192","20395193","20395194","20395195","20395196","20395197","20395198",
                            "20395199","20395200","20395201","20395202","20395203","20395204","20395205","20395206","20395207","20395208","20395209",
                            "20395210","20395211","20395212","20395214","20395215","20395216","20395217","20395218","20395219","20395220","20395221",
                            "20395222","20395223","20395224","20395225","20395226","20396131","20398849"]

    for i in range(len(zsd67)):
        art_check = zsd67['Materiale'].iloc[i]
        if any(codice == art_check for codice in mobil_giorno_smontati) :
            zsd67.categoria.iloc[i] = 'Mobiletti giorno smontati'
            
    st.write(zsd67[['Descrizione mat.','Materiale','Tp.Doc','C/lav','categoria']])

    df_dict = {}
    categorie = list(zsd67['categoria'].unique())
    i=0
    for categoria in categorie:
        i+=1
        df_fil = zsd67[zsd67['categoria'] == categoria][layout['output_gd']]
        df_dict[f'{categoria}.xlsx']= dp.create_excel_file(df_fil,f'{categoria}.xlsx')


    zip_data = dp.create_zip_file(df_dict)


    st.subheader('Download Zip Controllo Lotti', divider='red')
    st.write('Viene creata una cartella contenente un file excel per ogni fornitore')
    st.download_button(
        label="Scarica file zip",
        data=zip_data,
        file_name='files.zip',
        mime='application/zip'
    )



else:
    st.subheader(':red[STAI UTILIZZANDO LE IMPOSTAZIONI PER LG]')
    path_colori = st.file_uploader('caricare il file di abbinamento colori - T')

    if not path_colori:
        st.stop()

    flat_colori = pd.read_excel(path_colori)

    zsd67 = dividi_categorie_lg(zsd67, codici_carrellino)

    # recupero info T da file colori

    for i in range(len(zsd67)):
        check_col = zsd67.colore.iloc[i]
        if check_col == '':
            zsd67.colore.iloc[i] = 'Non trovato'
    
    zsd67 = zsd67.merge(flat_colori, how='left', left_on='colore', right_on='Colore')
    zsd67['T'] = zsd67['T'].fillna('T20')

    # raggruppamento codici Jeometrica

    # 1. estrazione degli ordini Jeometrica

    zsdjeo = zsd67[['JMT' in testo for testo in zsd67['Descrizione mat.']]]
    ordini_jeo = list(zsdjeo['Numero OdV'].unique())

    for i in range(len(zsd67)):
        odv_check = zsd67['Numero OdV'].iloc[i]

        if any(ordine== odv_check for ordine in ordini_jeo) and zsd67['C/lav'].iloc[i] != 'L':
            zsd67.categoria.iloc[i] = 'Jeometrica vendita'

    st.write(zsd67[['Descrizione mat.','Materiale','Tp.Doc','C/lav','categoria']])

    dp.scarica_excel(zsd67[layout['output2']], 'LGdivisione.xlsx')

    t_unique = list(zsd67['T'].unique())

    df_dict = {}
    for t in t_unique:
        zsd67_work = zsd67[zsd67['T']==t].copy()
        categorie = list(zsd67_work['categoria'].unique())
        i=0
        for categoria in categorie:
            i+=1
            df_fil = zsd67_work[zsd67_work['categoria'] == categoria][layout['output']]
            df_dict[f'{t}-{categoria}.xlsx']= dp.create_excel_file(df_fil,f'{categoria}.xlsx')


    zip_data = dp.create_zip_file(df_dict)


    st.subheader('Download Zip Controllo Lotti', divider='red')
    st.write('Viene creata una cartella contenente un file excel per ogni fornitore')
    st.download_button(
        label="Scarica file zip",
        data=zip_data,
        file_name='files.zip',
        mime='application/zip'
    )
