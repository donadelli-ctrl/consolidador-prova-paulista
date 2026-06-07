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


def normalizar(txt):

    if txt is None:
        return ""

    txt = str(txt)

    txt = unicodedata.normalize(
        "NFKD",
        txt
    ).encode(
        "ASCII",
        "ignore"
    ).decode(
        "utf-8"
    )

    txt = txt.upper()

    txt = re.sub(
        r"\s+",
        " ",
        txt
    )

    return txt.strip()


def turma_curta(texto):

    texto = normalizar(texto)

    numero = re.search(r"(\d+)", texto)

    letra = re.search(
        r"\b([A-Z])\b",
        texto
    )

    if numero and letra:
        return f"{numero.group(1)}{letra.group(1)}"

    return texto


def status(valor):

    if valor < 0.50:
        return "Abaixo do Básico"

    elif valor < 0.70:
        return "Básico"

    elif valor < 0.90:
        return "Adequado"

    return "Proficiente"


def cor_status(txt):

    cores = {

        "Abaixo do Básico": "F4CCCC",
        "Básico": "FFF2CC",
        "Adequado": "D9EAD3",
        "Proficiente": "CFE2F3"

    }

    return cores.get(
        txt,
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

    if not xlsm:
        st.error("Selecione o XLSM.")
        st.stop()

    if not zip_file:
        st.error("Selecione o ZIP.")
        st.stop()

    resultados = {}

    with tempfile.TemporaryDirectory() as pasta:

        caminho_zip = os.path.join(
            pasta,
            zip_file.name
        )

        with open(
            caminho_zip,
            "wb"
        ) as f:

            f.write(
                zip_file.getbuffer()
            )

        with zipfile.ZipFile(
            caminho_zip,
            "r"
        ) as z:

            z.extractall(pasta)

        arquivos_excel = []

        for raiz, _, arquivos in os.walk(pasta):

            for arq in arquivos:

                if arq.lower().endswith(".xlsx"):

                    arquivos_excel.append(
                        os.path.join(
                            raiz,
                            arq
                        )
                    )

        for arquivo in arquivos_excel:

            try:

                turma = os.path.splitext(
                    os.path.basename(
                        arquivo
                    )
                )[0]

                turma = normalizar(
                    turma
                )

                df = pd.read_excel(
                    arquivo
                )

                if "Nome" not in df.columns:
                    continue

                for _, row in df.iterrows():

                    nome = normalizar(
                        row["Nome"]
                    )

                    resultados[
                        (
                            turma,
                            nome
                        )
                    ] = {

                        "LP": row["PORT"],
                        "MAT": row["MAT"]

                    }

            except:
                pass

    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".xlsm"
    ) as tmp:

        tmp.write(
            xlsm.getbuffer()
        )

        caminho_xlsm = tmp.name

    wb = load_workbook(
        caminho_xlsm,
        keep_vba=True
    )

    encontrados = 0
    nao_encontrados = 0

    for ws in wb.worksheets:

        for linha in range(
            4,
            ws.max_row + 1
        ):

            turma_original = ws.cell(
                linha,
                1
            ).value

            estudante = ws.cell(
                linha,
                2
            ).value

            if not estudante:
                continue

            turma = turma_curta(
                turma_original
            )

            nome = normalizar(
                estudante
            )

            chave = (
                turma,
                nome
            )

            if chave not in resultados:

                nao_encontrados += 1
                continue

            reg = resultados[
                chave
            ]

            lp = reg["LP"]
            mat = reg["MAT"]

            ws.cell(
                linha,
                7
            ).value = lp

            ws.cell(
                linha,
                7
            ).number_format = "0.0%"

            st_lp = status(lp)

            ws.cell(
                linha,
                8
            ).value = st_lp

            ws.cell(
                linha,
                8
            ).fill = PatternFill(
                "solid",
                fgColor=cor_status(st_lp)
            )

            ws.cell(
                linha,
                9
            ).value = mat

            ws.cell(
                linha,
                9
            ).number_format = "0.0%"

            st_mat = status(mat)

            ws.cell(
                linha,
                10
            ).value = st_mat

            ws.cell(
                linha,
                10
            ).fill = PatternFill(
                "solid",
                fgColor=cor_status(st_mat)
            )

            encontrados += 1

    saida_xlsm = BytesIO()

    wb.save(
        saida_xlsm
    )

    saida_xlsm.seek(0)

    st.success(
        f"Encontrados: {encontrados}"
    )

    st.warning(
        f"Não encontrados: {nao_encontrados}"
    )

    st.download_button(
        "📥 Baixar XLSM",
        data=saida_xlsm.getvalue(),
        file_name="CONSOLIDADO.xlsm",
        mime="application/vnd.ms-excel"
    )
