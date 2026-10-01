#Bibliotecas
import pandas as pd
import numpy as np
from datetime import date
from openpyxl import load_workbook
import streamlit as st

st.title("Control de Pedidos", text_alignment="center")

st.divider()
st.subheader("Seleccionar el numero de bases activas")
numero_bases_activas = st.number_input("Numero de Bases activas?", min_value=1, max_value=10)

st.divider()
st.header("Archvios necesarios para el procesamiento:")
lista_de_materiales = st.file_uploader("Seleccionar archivo con la LISTA DE MATERIALES necesarios en el archivos de pedidos:", type=["xlsx","xls"])
inventario_sae_ = st.file_uploader("Seleccionar INVENTARIO DE SAE:", type=["xlsx","xls"])
inventario_homologado = st.file_uploader("Seleccionar INVENTARIO HOMOLOGADO SEMANAL:", type=["xlsx","xls", "xlsm"])
st.divider()


if lista_de_materiales is not None and inventario_sae_ is not None and inventario_homologado is not None:
  # Leer el archivo Excel de Lista de Materiales e Inventario SAE
  df = pd.read_excel(lista_de_materiales)
  inventario_sae = pd.read_excel(inventario_sae_)

  st.write("### Lista de Materiales")
  st.dataframe(df)

  st.write("### Inventario SAE")
  st.dataframe(inventario_sae)

  # Ingenieria de variables para posterior lectura del archivo de Inventario homologado semanal
  nombre_bases_activas = pd.ExcelFile(inventario_homologado).sheet_names[1:numero_bases_activas+1]
  columnas = list(np.array(pd.read_excel(inventario_homologado)[3:4])[0][:2]) + list(np.array(pd.read_excel(inventario_homologado)[2:3])[0][2:])
  #Agregamos la base a las columnas especificas de cada base
  col=1
  for i in nombre_bases_activas:
      for j in range(1,5):
          columnas[col+j] = columnas[col+j] + f' {i}'
      col=j+1

  #Leemos el archivo de Inventario homologado semanal
  homologado_inventarios_bases = pd.read_excel(inventario_homologado)[4:].dropna(subset=['Unnamed: 0'])
  homologado_inventarios_bases.columns = columnas
  homologado_inventarios_bases.rename(columns={'Total\n+\nCotizaciones':'Total en Bases'}, inplace=True)
  homologado_inventarios_bases.reset_index(drop=True, inplace=True)
  
  st.write("### Inventario Homologado semanal")
  st.dataframe(homologado_inventarios_bases)

  #Creamos un diccionario con las tablas de los inventarios de cada base activa
  tablas_bases = {}
  c=2
  for i in nombre_bases_activas:
      tablas_bases[i]=homologado_inventarios_bases[['NO. DE PARTE ', homologado_inventarios_bases.columns[c], homologado_inventarios_bases.columns[c+1], homologado_inventarios_bases.columns[c+2], homologado_inventarios_bases.columns[c+3]]]
      c+=4

  st.write(tablas_bases)










