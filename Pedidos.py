#Bibliotecas
import pandas as pd
import numpy as np
from datetime import date
from openpyxl import load_workbook
import streamlit as st

st.title("Control de Pedidos", text_alignment="center")

st.divider()
st.header("Archvios necesarios para el procesamiento:")
inventario_plataforma = st.file_uploader("Seleccionar archivo con la lista de materiales necesarios en el archivos de pedidos:", type=["xlsx","xls", "xlsm"])
