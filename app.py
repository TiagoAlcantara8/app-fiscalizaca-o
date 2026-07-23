import streamlit as st
import pandas as pd
from datetime import datetime
import os

# Configuração inicial da página
st.set_page_config(
    page_title="Fiscalização de Obras",
    layout="wide",
    page_icon="🔌"
)

# ==========================================
# FUNÇÕES PARA LIMPAR OS DADOS
# ==========================================
# Inicializa as variáveis na memória se elas não existirem
if "input_obra" not in st.session_state:
    st.session_state["input_obra"] = ""
if "input_municipio" not in st.session_state:
    st.session_state["input_municipio"] = ""
if "input_poste" not in st.session_state:
    st.session_state["input_poste"] = ""
if "input_obs" not in st.session_state:
    st.session_state["input_obs"] = ""

# Função que zera todos os campos (Vai para o topo)
def limpar_tudo():
    st.session_state["input_obra"] = ""
    st.session_state["input_municipio"] = ""
    st.session_state["input_poste"] = ""
    st.session_state["input_obs"] = ""

# Função que zera apenas a observação (Fica lá embaixo)
def limpar_obs():
    st.session_state["input_obs"] = ""

# ==========================================
# CABEÇALHO 
# ==========================================
st.title("🔌 Fiscalização de Obras")
st.markdown("---")

# Carrega composição
try:
    df_comp = pd.read_excel("composicoes.xlsx")
except Exception:
    st.error("Erro: O arquivo 'composicoes.xlsx' não foi encontrado na pasta.")
    st.stop()

# ==========================================
# 1. DADOS GERAIS & BOTÃO LIMPAR TUDO NO TOPO
# ==========================================
col_top1, col_top2 = st.columns([4, 1])
with col_top2:
    st.button("🧹 Limpar Todos os Dados", on_click=limpar_tudo, type="secondary", use_container_width=True)

col1, col2, col3 = st.columns(3)

with col1:
    obra = st.text_input("Número da Obra", key="input_obra")
with col2:
    municipio = st.text_input("Município", key="input_municipio")
with col3:
    poste = st.text_input("Identificação do Poste (Ex: P01, P02)", key="input_poste")

# ==========================================
# 2. ESTRUTURA E AÇÃO
# ==========================================
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

# ==========================================
# 3. MOSTRAR MATERIAIS E MÃO DE OBRA
# ==========================================
st.subheader("🛠️ Mão de Obra e Materiais")

if len(resultado) > 0:
    cod_mo = resultado.iloc[0]["Cod_MO"]
    desc_mo = resultado.iloc[0]["Descricao_MO"]
    st.success(f"**{cod_mo}** - {desc_mo}")
    st.dataframe(resultado[["Material", "Quantidade"]], use_container_width=True)
else:
    st.warning("Composição não encontrada.")

st.markdown("---")

# ==========================================
# 4. FOTOS
# ==========================================
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

# ==========================================
# 5. OBSERVAÇÕES E BOTÃO SALVAR
# ==========================================
st.markdown("---")
st.subheader("📝 Observações")
observacao = st.text_area("Digite suas observações", key="input_obs")

st.markdown("<br>", unsafe_allow_html=True) 

# Botões Lado a Lado (Salvar e Limpar Observação)
col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    salvar = st.button("✅ Salvar Fiscalização", use_container_width=True)
with col_btn2:
    st.button("🗑️ Limpar Apenas Observação", on_click=limpar_obs, type="secondary", use_container_width=True)

# ==========================================
# 6. LÓGICA DE SALVAR
# ==========================================
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
    
    arquivo = f"Fiscalizacao_Obra_{nome_obra_arquivo}.xlsx"

    # Salva na planilha
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

    st.success(f"Fiscalização do poste {nome_poste} salva na Obra {nome_obra_arquivo} com sucesso!")

# ==========================================
# 7. BOTÃO DE DOWNLOAD
# ==========================================
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
