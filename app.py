import streamlit as st

st.set_page_config(
    page_title="Consolidador Prova Paulista",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Consolidador Prova Paulista")

st.markdown("""
Faça upload da planilha da escola e do arquivo ZIP da Prova Paulista.
""")

xlsm = st.file_uploader(
    "Selecione a planilha XLSM",
    type=["xlsm"]
)

zip_file = st.file_uploader(
    "Selecione o arquivo ZIP",
    type=["zip"]
)

if st.button("🚀 Gerar Consolidado"):

    if not xlsm:
        st.error("Selecione o arquivo XLSM.")
        st.stop()

    if not zip_file:
        st.error("Selecione o arquivo ZIP.")
        st.stop()

    st.success("Arquivos recebidos com sucesso!")

    st.info(
        "Na próxima etapa será integrado o processamento automático."
    )
