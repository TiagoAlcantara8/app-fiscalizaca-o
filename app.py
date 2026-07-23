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

# 1. DADOS GERAIS
col1, col2, col3 = st.columns(3)

with col1:
    obra = st.text_input("Número da Obra")
with col2:
    municipio = st.text_input("Município")
with col3:
    poste = st.text_input("Identificação do Poste (Ex: P01, P02)")

# 2. ESTRUTURA E AÇÃO
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
    (df_comp["Estrutura"] == estrutura) & (df_comp["Acao"] == acao)
]

st.markdown("---")

# 3. MOSTRAR MATERIAIS E MÃO DE OBRA
st.subheader("Mão de Obra e Materiais")

if len(resultado) > 0:
    cod_mo = resultado.iloc[0]["Cod_MO"]
    desc_mo = resultado.iloc[0]["Descricao_MO"]
    st.success(f"{cod_mo} - {desc_mo}")
    st.dataframe(resultado[["Material", "Quantidade"]], use_container_width=True)
else:
    st.warning("Composição não encontrada.")

st.markdown("---")

# 4. FOTOS
st.subheader("📸 Fotos")

opcao_foto = st.radio(
    "Como deseja enviar a foto?",
    ["Tirar foto com a Câmera", "Anexar arquivo da Galeria", "Não enviar foto"],
    horizontal=True
)

foto_arquivo = None

if opcao_foto == "Tirar foto com a Câmera":
    foto_arquivo = st.camera_input("Tire uma foto da estrutura")
elif opcao_foto == "Anexar arquivo da Galeria":
    foto_arquivo = st.file_uploader("Escolha uma imagem", type=["jpg", "jpeg", "png"])

# 5. OBSERVAÇÕES E BOTÃO
st.markdown("---")
st.subheader("📝 Observações")
observacao = st.text_area("Digite suas observações")

salvar = st.button("Salvar Fiscalização")

# 6. LÓGICA DE SALVAR (UM ARQUIVO POR OBRA)
if salvar:
    nome_poste = poste.strip().upper() if poste.strip() != "" else "GERAL"
    nome_obra_arquivo = obra.strip() if obra.strip() != "" else "SEM_NUMERO"
    
    # Criar uma linha para a Mão de Obra
    dados_mo = {
        "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Obra": obra,
        "Municipio": municipio,
        "Poste": nome_poste,
        "Estrutura": estrutura,
        "Acao": acao,
        "Tipo_Item": "Mão de Obra",
        "Codigo": resultado.iloc[0]["Cod_MO"] if len(resultado) > 0 else "",
        "Descricao": resultado.iloc[0]["Descricao_MO"] if len(resultado) > 0 else "",
        "Quantidade": 1 if len(resultado) > 0 else 0,
        "Observacao": observacao
    }
    
    linhas_para_salvar = [dados_mo]
    
    # Criar linhas para os Materiais (se houver)
    if len(resultado) > 0:
        for index, row in resultado.iterrows():
            dados_mat = {
                "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "Obra": obra,
                "Municipio": municipio,
                "Poste": nome_poste,
                "Estrutura": estrutura,
                "Acao": acao,
                "Tipo_Item": "Material",
                "Codigo": "", 
                "Descricao": row["Material"],
                "Quantidade": row["Quantidade"],
                "Observacao": observacao
            }
            linhas_para_salvar.append(dados_mat)

    novo_df = pd.DataFrame(linhas_para_salvar)
    
    # Nomeia o arquivo de acordo com a obra
    arquivo = f"Fiscalizacao_Obra_{nome_obra_arquivo}.xlsx"

    # Salva na mesma aba contínua
    if os.path.exists(arquivo):
        existente = pd.read_excel(arquivo)
        final = pd.concat([existente, novo_df], ignore_index=True)
    else:
        final = novo_df

    final.to_excel(arquivo, index=False)

    # Salva a Foto
    if foto_arquivo is not None:
        os.makedirs("fotos", exist_ok=True)
        nome_foto = f"fotos/Obra_{nome_obra_arquivo}_Poste_{nome_poste}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        with open(nome_foto, "wb") as f:
            f.write(foto_arquivo.getbuffer())

    st.success(f"Fiscalização do poste {nome_poste} salva na Obra {nome_obra_arquivo}!")

# 7. BOTÃO DE DOWNLOAD (Dinâmico por Obra)
st.markdown("---")
st.subheader("📥 Download dos Dados")

obra_atual = obra.strip() if obra.strip() != "" else "SEM_NUMERO"
arquivo_gerado = f"Fiscalizacao_Obra_{obra_atual}.xlsx"

if os.path.exists(arquivo_gerado):
    with open(arquivo_gerado, "rb") as f:
        st.download_button(
            label=f"Baixar Planilha da Obra {obra_atual}",
            data=f,
            file_name=arquivo_gerado,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info(f"Nenhum dado salvo para a obra '{obra_atual}' ainda.")
