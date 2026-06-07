import streamlit as st
import pandas as pd
import zipfile
import tempfile
import os

from openpyxl import load_workbook

st.set_page_config(
    page_title="Consolidador Prova Paulista",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Consolidador Prova Paulista")

st.write(
    "Versão de Diagnóstico - Leitura dos Arquivos"
)

xlsm = st.file_uploader(
    "Selecione a planilha XLSM",
    type=["xlsm"]
)

zip_file = st.file_uploader(
    "Selecione o arquivo ZIP",
    type=["zip"]
)

if st.button("Analisar Arquivos"):

    if not xlsm:
        st.error("Selecione o XLSM.")
        st.stop()

    if not zip_file:
        st.error("Selecione o ZIP.")
        st.stop()

    try:

        st.subheader("📁 Analisando XLSM")

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsm"
        ) as tmp_xlsm:

            tmp_xlsm.write(
                xlsm.getbuffer()
            )

            caminho_xlsm = tmp_xlsm.name

        wb = load_workbook(
            caminho_xlsm,
            keep_vba=True
        )

        abas = wb.sheetnames

        st.success(
            f"Abas encontradas: {len(abas)}"
        )

        for aba in abas:
            st.write("•", aba)

        st.divider()

        st.subheader("📦 Analisando ZIP")

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

            for raiz, dirs, arquivos in os.walk(pasta):

                for arq in arquivos:

                    if arq.lower().endswith(
                        ".xlsx"
                    ):
                        arquivos_excel.append(
                            os.path.join(
                                raiz,
                                arq
                            )
                        )

            st.success(
                f"Arquivos Excel encontrados: {len(arquivos_excel)}"
            )

            total_alunos = 0

            for arq in arquivos_excel:

                try:

                    df = pd.read_excel(
                        arq
                    )

                    total_alunos += len(df)

                    st.write(
                        f"📄 {os.path.basename(arq)} - {len(df)} alunos"
                    )

                except Exception as erro:

                    st.error(
                        f"Erro ao ler {os.path.basename(arq)}"
                    )

            st.divider()

            st.subheader("📊 Resultado")

            st.metric(
                "Arquivos Excel",
                len(arquivos_excel)
            )

            st.metric(
                "Total de Registros",
                total_alunos
            )

            st.success(
                "Leitura concluída com sucesso."
            )

    except Exception as erro:

        st.error(str(erro))
