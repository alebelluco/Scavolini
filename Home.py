import streamlit as  st 
import pandas as pd 


st.set_page_config(layout='wide')

sx, cx, dx =  st.columns([1,2,5])
with sx:
    st.image('/Users/Alessandro/Desktop/APP/Scavolini/Immagine1.png')

with dx:    
    st.title('Toolbox Acquisti')

st.divider()

st.subheader('Applicazioni')