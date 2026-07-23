import streamlit as st
import pandas as pd
from datetime import datetime
import os

st.set_page_config(
    page_title="Fiscalização de Obras",
    layout="wide"
)

st.title("🔌 Fiscalização de Obras")
st.markdown("---")

# Carrega composição
df_comp = pd.read_excel(
    "composicoes.xlsx"
)

# Dados da obra
col1, col2 = st.columns(2)

with col1:
    obra = st.text_input(
        "Número da Obra"
    )

with col2:
    municipio = st.text_input(
        "Município"
    )

col1, col2 = st.columns(2)

with col1:
    estrutura = st.selectbox(
        "Estrutura",
        sorted(df_comp["Estrutura"].unique())
    )

with col2:
    acao = st.selectbox(
        "Ação",
        ["Instalação", "Retirada"]
    )

# Filtra composição
resultado = df_comp[
    (df_comp["Estrutura"] == estrutura)
    &
    (df_comp["Acao"] == acao)
]

st.markdown("---")

st.subheader("Mão de Obra")

if len(resultado) > 0:

    cod_mo = resultado.iloc[0]["Cod_MO"]
    desc_mo = resultado.iloc[0]["Descricao_MO"]

    st.success(
        f"{cod_mo} - {desc_mo}"
    )

    st.subheader("Materiais")

    st.dataframe(
        resultado[["Material", "Quantidade"]],
        use_container_width=True
    )

else:
    st.warning(
        "Composição não encontrada."
    )

st.markdown("---")

st.subheader("Fotos")

foto = st.camera_input(
    "Tire uma foto da estrutura"
)

st.subheader("Observações")

observacao = st.text_area(
    "Digite suas observações"
)

salvar = st.button(
    "Salvar Fiscalização"
)

if salvar:

    data = {
        "Data": datetime.now(),
        "Obra": obra,
        "Municipio": municipio,
        "Estrutura": estrutura,
        "Acao": acao,
        "Cod_MO": resultado.iloc[0]["Cod_MO"] if len(resultado) > 0 else "",
        "Descricao_MO": resultado.iloc[0]["Descricao_MO"] if len(resultado) > 0 else "",
        "Observacao": observacao
    }

    novo = pd.DataFrame([data])

    arquivo = "fiscalizacoes.xlsx"

    if os.path.exists(arquivo):
        existente = pd.read_excel(arquivo)
        final = pd.concat(
            [existente, novo],
            ignore_index=True
        )
    else:
        final = novo

    final.to_excel(
        arquivo,
        index=False
    )

    if foto:
        os.makedirs(
            "fotos",
            exist_ok=True
        )
        nome_foto = (
            f"fotos/{obra}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        )
        with open(nome_foto, "wb") as f:
            f.write(foto.getbuffer())

    st.success(
        "Fiscalização salva com sucesso!"
    )

# ==========================================
# NOVO BLOCO: DOWNLOAD DA PLANILHA GERADA
# ==========================================
st.markdown("---")
st.subheader("📥 Download dos Dados")

arquivo_gerado = "fiscalizacoes.xlsx"

if os.path.exists(arquivo_gerado):
    with open(arquivo_gerado, "rb") as f:
        st.download_button(
            label="Baixar Planilha de Fiscalizações",
            data=f,
            file_name="fiscalizacoes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Nenhuma fiscalização salva ainda. Preencha o formulário acima para gerar a planilha!")
