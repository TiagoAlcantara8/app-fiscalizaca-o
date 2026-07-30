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

# Função auxiliar para garantir que valores vazios não quebrem o cálculo
def tratar_valor(val):
    if pd.isna(val):
        return 0.0
    try:
        return float(val)
    except:
        return 0.0

# ==========================================
# INICIALIZAÇÃO DA MEMÓRIA (ESTADOS)
# ==========================================
if "input_obra" not in st.session_state:
    st.session_state["input_obra"] = ""
if "input_municipio" not in st.session_state:
    st.session_state["input_municipio"] = ""
if "input_poste" not in st.session_state:
    st.session_state["input_poste"] = ""
if "input_obs" not in st.session_state:
    st.session_state["input_obs"] = ""
if "carrinho" not in st.session_state:
    st.session_state["carrinho"] = []

def limpar_tudo():
    st.session_state["input_obra"] = ""
    st.session_state["input_municipio"] = ""
    st.session_state["input_poste"] = ""
    st.session_state["input_obs"] = ""
    st.session_state["carrinho"] = []

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
    
    tb_mao_obra.columns = tb_mao_obra.columns.str.strip()
    tb_composicao.columns = tb_composicao.columns.str.strip()
    tb_material.columns = tb_material.columns.str.strip()

    for df in [tb_mao_obra, tb_composicao, tb_material]:
        if "__PowerAppsId__" in df.columns:
            df.drop(columns=["__PowerAppsId__"], inplace=True)
            
except Exception:
    st.error("Erro: O arquivo 'BANCO.xlsx' não foi encontrado. Faça o upload no GitHub.")
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
# 3. ESTRUTURAS (CARRINHO DE COMPRAS)
# ==========================================
st.markdown("---")
st.subheader("⚙️ Seleção de Estruturas (MT e BT)")

col_rad1, col_rad2 = st.columns(2)
with col_rad1:
    tipo_rede = st.radio("Selecione o Tipo de Rede:", ["MT", "BT"], horizontal=True)
with col_rad2:
    tipo_execucao = st.radio("Método de Execução:", ["Linha Morta (LM)", "Linha Viva (LV)"], horizontal=True)

tb_comp_filtrada = tb_composicao[tb_composicao["Rede"] == tipo_rede]

col_e1, col_e2 = st.columns(2)
estruturas_disponiveis = sorted(tb_comp_filtrada["Estrutura"].dropna().unique())

with col_e1:
    estrutura = st.selectbox("Estrutura Adicional", estruturas_disponiveis)
with col_e2:
    acao = st.selectbox("Ação da Estrutura", ["Instalação", "Retirada"])

# Lógica para Mão de Obra (LV vs LM)
sigla_exec = "LV" if tipo_execucao == "Linha Viva (LV)" else "LM"
mo_info = tb_mao_obra[(tb_mao_obra["Estrutura"] == estrutura) & (tb_mao_obra["Tipo"] == sigla_exec)]

valor_previsto = 0.0

if len(mo_info) > 0:
    cod_mo = mo_info.iloc[0]["Cod_MO"]
    desc_mo = mo_info.iloc[0]["Descricao_MO"]
    col_valor = "Valor_Instalar" if acao == "Instalação" else "Valor_Retirar"
    valor_previsto = tratar_valor(mo_info.iloc[0][col_valor])
else:
    mo_info_fallback = tb_mao_obra[tb_mao_obra["Estrutura"] == estrutura]
    if len(mo_info_fallback) > 0:
        cod_mo = mo_info_fallback.iloc[0]["Cod_MO"]
        desc_mo = mo_info_fallback.iloc[0]["Descricao_MO"]
        col_valor = "Valor_Instalar" if acao == "Instalação" else "Valor_Retirar"
        valor_previsto = tratar_valor(mo_info_fallback.iloc[0][col_valor])
    else:
        cod_mo = "MO_EXTRA"
        desc_mo = f"Mão de Obra para {estrutura}"

st.success(f"**Mão de Obra identificada:** {cod_mo} - {desc_mo} | Valor Previsto: R$ {valor_previsto:,.2f}")

# Busca Materiais
resultado_materiais = tb_comp_filtrada[tb_comp_filtrada["Estrutura"] == estrutura]
if len(resultado_materiais) > 0:
    df_materiais_base = resultado_materiais[["CodMaterial", "Material", "Quantidade"]].copy()
    df_materiais_base.rename(columns={"CodMaterial": "Codigo"}, inplace=True)
else:
    st.warning("Materiais não encontrados para esta estrutura. Insira manualmente abaixo:")
    df_materiais_base = pd.DataFrame(columns=["Codigo", "Material", "Quantidade"])

df_editavel = st.data_editor(
    df_materiais_base,
    num_rows="dynamic", 
    use_container_width=True,
    hide_index=True,
    key=f"editor_{estrutura}"
)

if st.button("➕ Adicionar Esta Estrutura ao Poste", type="primary"):
    item_carrinho = {
        "estrutura": estrutura,
        "acao": acao,
        "rede": tipo_rede,
        "execucao": sigla_exec,
        "cod_mo": cod_mo,
        "desc_mo": desc_mo,
        "valor_mo": valor_previsto,
        "materiais": df_editavel.to_dict('records')
    }
    st.session_state["carrinho"].append(item_carrinho)
    st.rerun()

# ==========================================
# 4. RESUMO DO POSTE
# ==========================================
st.markdown("---")
st.subheader("🛒 Estruturas Adicionadas neste Poste")

if len(st.session_state["carrinho"]) == 0:
    st.info("Nenhuma estrutura adicionada ainda. Preencha acima e clique em 'Adicionar Esta Estrutura'.")
else:
    for i, item in enumerate(st.session_state["carrinho"]):
        st.write(f"**{i+1}. {item['estrutura']}** ({item['rede']} - {item['execucao']}) | R$ {item['valor_mo']:,.2f} | *{len(item['materiais'])} materiais*")
    
    if st.button("🗑️ Esvaziar Estruturas do Poste"):
        st.session_state["carrinho"] = []
        st.rerun()

# ==========================================
# 5. FOTOS E OBSERVAÇÕES
# ==========================================
st.markdown("---")
st.subheader("📸 Fotos")
opcao_foto = st.radio("Como deseja enviar a foto?", ["Tirar foto com a Câmera", "Anexar arquivo da Galeria", "Não enviar foto"], horizontal=True)

foto_arquivo = None
if opcao_foto == "Tirar foto com a Câmera":
    foto_arquivo = st.camera_input("Tire uma foto do poste completo")
elif opcao_foto == "Anexar arquivo da Galeria":
    foto_arquivo = st.file_uploader("Escolha uma imagem do poste completo", type=["jpg", "jpeg", "png"])

st.markdown("---")
st.subheader("📝 Observações Gerais do Poste")
observacao = st.text_area("Digite suas observações", key="input_obs")

# ==========================================
# 6. SALVAMENTO FINAL DA PLANILHA OFICIAL
# ==========================================
st.markdown("<br>", unsafe_allow_html=True) 
salvar = st.button("✅ Salvar Fiscalização Completa do Poste", use_container_width=True, type="primary")

if salvar:
    if len(st.session_state["carrinho"]) == 0 and not incluir_poste:
        st.error("⚠️ Adicione pelo menos uma estrutura ou marque a inclusão do poste limpo antes de salvar!")
    else:
        nome_poste = poste.strip().upper() if poste.strip() != "" else "GERAL"
        nome_obra_arquivo = obra.strip() if obra.strip() != "" else "SEM_NUMERO"
        
        linhas_para_salvar = []

        if incluir_poste:
            valor_poste = 0.0
            poste_match = tb_mao_obra[tb_mao_obra["Cod_MO"].astype(str) == str(cod_poste)]
            if len(poste_match) > 0:
                valor_poste = tratar_valor(poste_match.iloc[0]["Valor_Instalar"])

            dados_poste = {
                "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "Obra": obra,
                "Municipio": municipio,
                "Poste": nome_poste,
                "Rede": "MT/BT", 
                "Estrutura": "Poste Limpo",
                "Acao": "Instalação",
                "Execucao": "LM",
                "Tipo_Item": "Mão de Obra",
                "Codigo": cod_poste,
                "Descricao": desc_poste,
                "Quantidade": 1,
                "Valor_Unitario": round(valor_poste, 2),
                "Valor_Total": round(valor_poste, 2),
                "Observacao": observacao
            }
            linhas_para_salvar.append(dados_poste)

        for item in st.session_state["carrinho"]:
            dados_mo_est = {
                "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "Obra": obra,
                "Municipio": municipio,
                "Poste": nome_poste,
                "Rede": item["rede"], 
                "Estrutura": item["estrutura"],
                "Acao": item["acao"],
                "Execucao": item["execucao"],
                "Tipo_Item": "Mão de Obra",
                "Codigo": item["cod_mo"],
                "Descricao": item["desc_mo"],
                "Quantidade": 1,
                "Valor_Unitario": round(item["valor_mo"], 2),
                "Valor_Total": round(item["valor_mo"], 2),
                "Observacao": observacao
            }
            linhas_para_salvar.append(dados_mo_est)
            
            for mat in item["materiais"]:
                if pd.notna(mat.get("Material")) and str(mat.get("Material")).strip() != "":
                    dados_mat = {
                        "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                        "Obra": obra,
                        "Municipio": municipio,
                        "Poste": nome_poste,
                        "Rede": item["rede"], 
                        "Estrutura": item["estrutura"],
                        "Acao": item["acao"],
                        "Execucao": item["execucao"],
                        "Tipo_Item": "Material",
                        "Codigo": mat.get("Codigo", ""),
                        "Descricao": mat.get("Material", ""),
                        "Quantidade": mat.get("Quantidade", 0),
                        "Valor_Unitario": 0.0, 
                        "Valor_Total": 0.0,
                        "Observacao": observacao
                    }
                    linhas_para_salvar.append(dados_mat)

        novo_df = pd.DataFrame(linhas_para_salvar)
        arquivo = f"Fiscalizacao_Obra_{nome_obra_arquivo}.xlsx"

        if os.path.exists(arquivo):
            # Lê SOMENTE a aba 0 (Dados) para evitar puxar as informações da aba de Resumo
            existente = pd.read_excel(arquivo, sheet_name=0)
            
            # Caso o usuário tenha salvo o arquivo usando o código anterior, limpamos a coluna velha
            if "Gasto_Previsto_Obra" in existente.columns:
                existente = existente.drop(columns=["Gasto_Previsto_Obra"])
                
            final = pd.concat([existente, novo_df], ignore_index=True)
        else:
            final = novo_df

        # Garante que os valores numéricos estão corretos
        final["Valor_Total"] = pd.to_numeric(final["Valor_Total"], errors="coerce").fillna(0.0)

        # ========================================================
        # NOVA LÓGICA: CRIA A TABELA DE RESUMO EM UMA ABA SEPARADA
        # ========================================================
        # Agrupa tudo pelo nome da obra e soma
        df_resumo = final.groupby("Obra", as_index=False)["Valor_Total"].sum()
        df_resumo.rename(columns={"Valor_Total": "VALOR PREVISTO DA OBRA"}, inplace=True)

        # Salva o arquivo usando o ExcelWriter para criar múltiplas abas (Planilhas)
        with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
            final.to_excel(writer, sheet_name="Dados", index=False)
            df_resumo.to_excel(writer, sheet_name="Resumo", index=False)

        if foto_arquivo is not None:
            os.makedirs("fotos", exist_ok=True)
            nome_foto = f"fotos/Obra_{nome_obra_arquivo}_Poste_{nome_poste}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
            with open(nome_foto, "wb") as f:
                f.write(foto_arquivo.getbuffer())

        st.session_state["carrinho"] = []
        st.success(f"Fiscalização completa do poste {nome_poste} salva na Obra {nome_obra_arquivo}!")

# ==========================================
# 7. BOTÕES DE DOWNLOAD (EXCEL E FOTOS)
# ==========================================
st.markdown("---")
st.subheader("📥 Download dos Dados Oficiais")

obra_atual = obra.strip() if obra.strip() != "" else "SEM_NUMERO"
arquivo_gerado = f"Fiscalizacao_Obra_{obra_atual}.xlsx"
col_down1, col_down2 = st.columns(2)

with col_down1:
    if os.path.exists(arquivo_gerado):
        with open(arquivo_gerado, "rb") as f:
            st.download_button(label=f"📊 Baixar Planilha COMPLETA da Obra {obra_atual}", data=f, file_name=arquivo_gerado, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    else:
        st.info(f"Nenhum dado salvo para a obra '{obra_atual}'. Use o botão verde acima para salvar primeiro.")

with col_down2:
    if os.path.exists("fotos") and len(os.listdir("fotos")) > 0:
        shutil.make_archive("fotos_backup", 'zip', "fotos")
        with open("fotos_backup.zip", "rb") as f_zip:
            st.download_button(label="📦 Baixar Todas as Fotos (ZIP)", data=f_zip, file_name=f"Fotos_Fiscalizacao_{datetime.now().strftime('%Y%m%d')}.zip", mime="application/zip", use_container_width=True)
    else:
        st.info("Nenhuma foto salva ainda.")
