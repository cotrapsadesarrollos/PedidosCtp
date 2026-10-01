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
  # Leer el archivo Excel del inventario de la plataforma e Inventario SAE
  df = pd.read_excel(lista_de_materiales)
  inventario_sae = pd.read_excel(inventario_sae_)

  st.write("### Lista de Materiales")
  st.dataframe(df)

  st.write("### Inventario SAE")
  st.dataframe(inventario_sae)

  nombre_bases_activas = pd.ExcelFile(inventario_homologado).sheet_names[1:numero_bases_activas+1]
  st.write(nombre_bases_activas)
