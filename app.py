import streamlit as st
import pandas as pd
import zipfile
import tempfile
import os
import re
import unicodedata
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(
    page_title="Consolidador Prova Paulista",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Consolidador Prova Paulista")

def normalizar(texto):

    if texto is None:
        return ""

    texto = str(texto)

    texto = unicodedata.normalize(
        'NFKD',
        texto
    ).encode(
        'ASCII',
        'ignore'
    ).decode(
        'utf-8'
    )

    texto = texto.upper()

    texto = re.sub(
        r'\s+',
        ' ',
        texto
    )

    return texto.strip()

def calcular_status(valor):

    if valor < 0.50:
        return "Abaixo do Básico"

    elif valor < 0.70:
        return "Básico"

    elif valor < 0.90:
        return "Adequado"

    return "Proficiente"

def cor_status(status):

    cores = {

        "Abaixo do Básico": "F4CCCC",
        "Básico": "FFF2CC",
        "Adequado": "D9EAD3",
        "Proficiente": "CFE2F3"

    }

    return cores.get(
        status,
        "FFFFFF"
    )

xlsm = st.file_uploader(
    "Planilha XLSM",
    type=["xlsm"]
)

zip_file = st.file_uploader(
    "Arquivo ZIP",
    type=["zip"]
)
if st.button("Gerar Consolidado"):

    st.info(
        "Versão de consolidação em preparação..."
    )
