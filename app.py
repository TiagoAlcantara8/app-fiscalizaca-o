import streamlit as st
import pandas as pd
from datetime import datetime
import os
import shutil

# Configuração inicial da página
st.set_page_config(
    page_title="Fiscalização de Obras",
    layout="wide",
    page_icon="🔌"
)

# ==========================================
# FUNÇÕES PARA LIMPAR OS DADOS
# ==========================================
if "input_obra" not in st.session_state:
    st.session_state["input_obra"] = ""
if "input_municipio" not in st.session_state:
    st.session_state["input_municipio"] = ""
if "input_poste" not in st.session_state:
    st.session_state["input_poste"] = ""
if "input_obs" not in st.session_state:
    st.session_state["input_obs"] = ""

def limpar_tudo():
    st.session_state["input_obra"] = ""
    st.session_state["input_municipio"] = ""
    st.session_state["input_poste"] = ""
    st.session_state["input_obs"] = ""

def limpar_obs():
    st.session_state["input_obs"] = ""

# ==========================================
# CABEÇALHO E VISUAL (IMAGENS)
# ==========================================
try:
    st.image("imagem.png", use_container_width=True)
except Exception:
    st.title("🔌 Fiscalização de Obras")

st.markdown("---")

# ==========================================
# CARREGAMENTO DO BANCO DE DADOS
# ==========================================
try:
    xls = pd.ExcelFile("BANCO.xlsx")
    tb_mao_obra = pd.read_excel(xls, "tb_MaoObra")
    tb_composicao = pd.read_excel(xls, "tb_Composicao")
    tb_material = pd.read_excel(xls, "tb_Material")
    
    # Limpa espaços invisíveis nos nomes das colunas
    tb_mao_obra.columns = tb_mao_obra.columns.str.strip()
    tb_composicao.columns = tb_composicao.columns.str.strip()
    tb_material.columns = tb_material.columns.str.strip()

    # Deleta a coluna de ID do Power Apps para não atrapalhar
    for df in [tb_mao_obra, tb_composicao, tb_material]:
        if "__PowerAppsId__" in df.columns:
            df.drop(columns=["__PowerAppsId__"], inplace=True)
            
except Exception:
    st.error("Erro: O arquivo 'BANCO.xlsx' não foi encontrado na pasta. Faça o upload no GitHub.")
    st.stop()

# ==========================================
# 1. DADOS GERAIS
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
# 2. DADOS DO POSTE
# ==========================================
st.markdown("---")
st.subheader("🗼 Dados do Poste Principal")

incluir_poste = st.checkbox("Incluir instalação/retirada de Poste Limpo nesta fiscalização?")

cod_poste = ""
desc_poste = ""

if incluir_poste:
    col_p1, col_p2 = st.columns(2)
    
    with col_p1:
        altura = st.selectbox("Altura do Poste (m)", [10, 11, 12, 13, 14, 15, 16])
    
    with col_p2:
        esforco = st.selectbox("Esforço (daN)", [300, 600, 1000])

    if altura <= 12 and esforco <= 600:
        cod_poste = "1015"
        desc_poste = "Poste Limpo (sem mat. ou equip.) 12 >=P<= 600"
    elif altura <= 12 and esforco > 600:
        cod_poste = "1016"
        desc_poste = "Poste Limpo (sem mat. ou equip.) 12 >=P> 600"
    elif 12 < altura <= 16 and esforco >= 600:
        cod_poste = "1017"
        desc_poste = "Poste Limpo (sem mat. ou equip.) 12<P<=16 e P>=600"
    else:
        cod_poste = "1018"
        desc_poste = "Poste Limpo (sem mat. ou equip.) Configuração Especial"

    st.info(f"Código Mão de Obra do Poste gerado: **{cod_poste}** - {desc_poste}")

# ==========================================
# 3. ESTRUTURA E AÇÃO (COM FILTRO MT/BT)
# ==========================================
st.markdown("---")
st.subheader("⚙️ Estruturas e Equipamentos")

# Botões para escolher entre Média Tensão ou Baixa Tensão
tipo_rede = st.radio("Selecione o Tipo de Rede da Estrutura:", ["MT", "BT"], horizontal=True)

# Filtra a aba de composições pelo tipo de rede escolhido
tb_comp_filtrada = tb_composicao[tb_composicao["Rede"] == tipo_rede]

col_e1, col_e2 = st.columns(2)

# Busca as estruturas únicas APENAS da rede selecionada (MT ou BT)
estruturas_disponiveis = sorted(tb_comp_filtrada["Estrutura"].dropna().unique())

with col_e1:
    estrutura = st.selectbox("Estrutura Adicional", estruturas_disponiveis)

with col_e2:
    acao = st.selectbox("Ação da Estrutura", ["Instalação", "Retirada"])

# ==========================================
# 4. TABELA DE MATERIAIS E MÃO DE OBRA
# ==========================================
st.info("💡 **Dica:** Você pode alterar as quantidades, apagar linhas ou adicionar novos materiais na tabela abaixo.")

# 4.1 Busca a Mão de Obra correspondente na tb_MaoObra
mo_info = tb_mao_obra[tb_mao_obra["Estrutura"] == estrutura]

if len(mo_info) > 0:
    cod_mo = mo_info.iloc[0]["Cod_MO"]
    desc_mo = mo_info.iloc[0]["Descricao_MO"]
else:
    cod_mo = "MO_EXTRA"
    desc_mo = f"Mão de Obra para {estrutura}"

st.success(f"**{cod_mo}** - {desc_mo} ({acao})")

# 4.2 Busca os Materiais correspondentes na planilha já filtrada
resultado_materiais = tb_comp_filtrada[tb_comp_filtrada["Estrutura"] == estrutura]

if len(resultado_materiais) > 0:
    # Seleciona Código, Nome e Quantidade para o editor
    df_materiais_base = resultado_materiais[["CodMaterial", "Material", "Quantidade"]].copy()
    df_materiais_base.rename(columns={"CodMaterial": "Codigo"}, inplace=True)
else:
    st.warning("Materiais não encontrados para esta estrutura. Insira manualmente abaixo:")
    df_materiais_base = pd.DataFrame(columns=["Codigo", "Material", "Quantidade"])

# Gera a tabela interativa
df_editavel = st.data_editor(
    df_materiais_base,
    num_rows="dynamic", 
    use_container_width=True,
    key="editor_materiais"
)

# ==========================================
# 5. FOTOS E OBSERVAÇÕES
# ==========================================
st.markdown("---")
st.subheader("📸 Fotos")
opcao_foto = st.radio("Como deseja enviar a foto?", ["Tirar foto com a Câmera", "Anexar arquivo da Galeria", "Não enviar foto"], horizontal=True)

foto_arquivo = None
if opcao_foto == "Tirar foto com a Câmera":
    foto_arquivo = st.camera_input("Tire uma foto da estrutura")
elif opcao_foto == "Anexar arquivo da Galeria":
    foto_arquivo = st.file_uploader("Escolha uma imagem", type=["jpg", "jpeg", "png"])

st.markdown("---")
st.subheader("📝 Observações")
observacao = st.text_area("Digite suas observações", key="input_obs")
st.markdown("<br>", unsafe_allow_html=True) 

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
    
    linhas_para_salvar = []

    # 1. Salva a Mão de Obra do Poste
    if incluir_poste:
        dados_poste = {
            "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Obra": obra,
            "Municipio": municipio,
            "Poste": nome_poste,
            "Estrutura": "Poste Limpo",
            "Acao": "Instalação", 
            "Tipo_Item": "Mão de Obra",
            "Codigo": cod_poste,
            "Descricao": desc_poste,
            "Quantidade": 1,
            "Observacao": observacao
        }
        linhas_para_salvar.append(dados_poste)

    # 2. Salva a Mão de Obra da Estrutura
    dados_mo_est = {
        "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Obra": obra,
        "Municipio": municipio,
        "Poste": nome_poste,
        "Estrutura": estrutura,
        "Acao": acao,
        "Tipo_Item": "Mão de Obra",
        "Codigo": cod_mo,
        "Descricao": desc_mo,
        "Quantidade": 1,
        "Observacao": observacao
    }
    linhas_para_salvar.append(dados_mo_est)
    
    # 3. Salva os Materiais da Tabela Editada
    if len(df_editavel) > 0:
        for index, row in df_editavel.iterrows():
            if pd.notna(row["Material"]) and str(row["Material"]).strip() != "":
                dados_mat = {
                    "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "Obra": obra,
                    "Municipio": municipio,
                    "Poste": nome_poste,
                    "Estrutura": estrutura,
                    "Acao": acao,
                    "Tipo_Item": "Material",
                    "Codigo": row.get("Codigo", ""), 
                    "Descricao": row["Material"],
                    "Quantidade": row["Quantidade"],
                    "Observacao": observacao
                }
                linhas_para_salvar.append(dados_mat)

    novo_df = pd.DataFrame(linhas_para_salvar)
    arquivo = f"Fiscalizacao_Obra_{nome_obra_arquivo}.xlsx"

    if os.path.exists(arquivo):
        existente = pd.read_excel(arquivo)
        final = pd.concat([existente, novo_df], ignore_index=True)
    else:
        final = novo_df

    final.to_excel(arquivo, index=False)

    # 4. Salva a Foto
    if foto_arquivo is not None:
        os.makedirs("fotos", exist_ok=True)
        nome_foto = f"fotos/Obra_{nome_obra_arquivo}_Poste_{nome_poste}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        with open(nome_foto, "wb") as f:
            f.write(foto_arquivo.getbuffer())

    st.success(f"Fiscalização salva com sucesso na Obra {nome_obra_arquivo}!")

# ==========================================
# 7. BOTÕES DE DOWNLOAD (EXCEL E FOTOS)
# ==========================================
st.markdown("---")
st.subheader("📥 Download dos Dados e Fotos")

obra_atual = obra.strip() if obra.strip() != "" else "SEM_NUMERO"
arquivo_gerado = f"Fiscalizacao_Obra_{obra_atual}.xlsx"
col_down1, col_down2 = st.columns(2)

with col_down1:
    if os.path.exists(arquivo_gerado):
        with open(arquivo_gerado, "rb") as f:
            st.download_button(label=f"📊 Baixar Planilha da Obra {obra_atual}", data=f, file_name=arquivo_gerado, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    else:
        st.info(f"Nenhum dado salvo para a obra '{obra_atual}'.")

with col_down2:
    if os.path.exists("fotos") and len(os.listdir("fotos")) > 0:
        shutil.make_archive("fotos_backup", 'zip', "fotos")
        with open("fotos_backup.zip", "rb") as f_zip:
            st.download_button(label="📦 Baixar Todas as Fotos (ZIP)", data=f_zip, file_name=f"Fotos_Fiscalizacao_{datetime.now().strftime('%Y%m%d')}.zip", mime="application/zip", use_container_width=True)
    else:
        st.info("Nenhuma foto salva ainda.")
