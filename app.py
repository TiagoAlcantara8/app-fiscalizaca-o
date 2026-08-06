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

# Função auxiliar
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
if "carrinho_cabos" not in st.session_state:
    st.session_state["carrinho_cabos"] = []

def limpar_tudo():
    st.session_state["input_obra"] = ""
    st.session_state["input_municipio"] = ""
    st.session_state["input_poste"] = ""
    st.session_state["input_obs"] = ""
    st.session_state["carrinho"] = []
    st.session_state["carrinho_cabos"] = []

def limpar_obs():
    st.session_state["input_obs"] = ""

# ==========================================
# CABEÇALHO E VISUAL
# ==========================================
try:
    st.image("imagem.png", use_container_width=True)
except Exception:
    st.title("🔌 Fiscalização de Obras")

st.markdown("---")

# ==========================================
# CARREGAMENTO DOS BANCOS DE DADOS
# ==========================================
try:
    xls = pd.ExcelFile("BANCO.xlsx")
    tb_mao_obra = pd.read_excel(xls, "tb_MaoObra")
    tb_composicao = pd.read_excel(xls, "tb_Composicao")
    
    tb_mao_obra.columns = tb_mao_obra.columns.str.strip()
    tb_composicao.columns = tb_composicao.columns.str.strip()

    for df in [tb_mao_obra, tb_composicao]:
        if "__PowerAppsId__" in df.columns:
            df.drop(columns=["__PowerAppsId__"], inplace=True)
            
except Exception:
    st.error("Erro: O arquivo 'BANCO.xlsx' não foi encontrado. Faça o upload no GitHub.")
    st.stop()

# Carregamento da aba de Cabos (Kit Técnico)
try:
    df_cabos = pd.read_excel("Kit Tecnico - Copy.xlsx", sheet_name="Peso cabos", header=1)
    df_cabos.rename(columns={'Unnamed: 1': 'Descricao'}, inplace=True)
    df_cabos = df_cabos.dropna(subset=['Descricao', 'FATOR kg/m'])
except Exception:
    df_cabos = pd.DataFrame() 

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
    poste = st.text_input("Identificação (Ex: P01 ao P10)", key="input_poste")

# ==========================================
# 2. DADOS DO POSTE
# ==========================================
st.markdown("---")
st.subheader("🗼 Dados do Poste Principal")

incluir_poste = st.checkbox("Incluir instalação/retirada de Poste Limpo nesta fiscalização?")
cod_poste = ""
desc_poste = ""
qtd_poste = 1 # Variável padrão caso ele não marque

if incluir_poste:
    col_p1, col_p2, col_p3 = st.columns([2, 2, 1])
    with col_p1:
        altura = st.selectbox("Altura do Poste (m)", [10, 11, 12, 13, 14, 15, 16])
    with col_p2:
        esforco = st.selectbox("Esforço (daN)", [300, 600, 1000])
    with col_p3:
        qtd_poste = st.number_input("Qtd Postes", min_value=1, value=1, step=1)

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
    st.info(f"Código Mão de Obra gerado: **{cod_poste}** - {desc_poste} | Multiplicador: {qtd_poste}x")

# ==========================================
# 3. ESTRUTURAS
# ==========================================
st.markdown("---")
st.subheader("⚙️ Seleção de Estruturas (MT e BT)")

col_rad1, col_rad2 = st.columns(2)
with col_rad1:
    tipo_rede = st.radio("Selecione o Tipo de Rede:", ["MT", "BT"], horizontal=True)
with col_rad2:
    tipo_execucao = st.radio("Método de Execução:", ["Linha Morta (LM)", "Linha Viva (LV)"], horizontal=True)

tb_comp_filtrada = tb_composicao[tb_composicao["Rede"] == tipo_rede]
col_e1, col_e2, col_e3 = st.columns([2, 2, 1])

with col_e1:
    estrutura = st.selectbox("Estrutura Adicional", sorted(tb_comp_filtrada["Estrutura"].dropna().unique()))
with col_e2:
    acao = st.selectbox("Ação da Estrutura", ["Instalação", "Retirada"])
with col_e3:
    qtd_estrutura = st.number_input("Qtd (Multiplicador)", min_value=1, value=1, step=1)

sigla_exec = "LV" if tipo_execucao == "Linha Viva (LV)" else "LM"
mo_info = tb_mao_obra[(tb_mao_obra["Estrutura"] == estrutura) & (tb_mao_obra["Tipo"] == sigla_exec)]
valor_previsto_unitario = 0.0

if len(mo_info) > 0:
    cod_mo = mo_info.iloc[0]["Cod_MO"]
    desc_mo = mo_info.iloc[0]["Descricao_MO"]
    col_valor = "Valor_Instalar" if acao == "Instalação" else "Valor_Retirar"
    valor_previsto_unitario = tratar_valor(mo_info.iloc[0][col_valor])
else:
    mo_info_fallback = tb_mao_obra[tb_mao_obra["Estrutura"] == estrutura]
    if len(mo_info_fallback) > 0:
        cod_mo = mo_info_fallback.iloc[0]["Cod_MO"]
        desc_mo = mo_info_fallback.iloc[0]["Descricao_MO"]
        col_valor = "Valor_Instalar" if acao == "Instalação" else "Valor_Retirar"
        valor_previsto_unitario = tratar_valor(mo_info_fallback.iloc[0][col_valor])
    else:
        cod_mo = "MO_EXTRA"
        desc_mo = f"Mão de Obra para {estrutura}"

st.success(f"**Mão de Obra:** {cod_mo} - {desc_mo} | Unitário: R$ {valor_previsto_unitario:,.2f} | **Total do Trecho: R$ {valor_previsto_unitario * qtd_estrutura:,.2f}**")

resultado_materiais = tb_comp_filtrada[tb_comp_filtrada["Estrutura"] == estrutura]
if len(resultado_materiais) > 0:
    df_materiais_base = resultado_materiais[["CodMaterial", "Material", "Quantidade"]].copy()
    df_materiais_base.rename(columns={"CodMaterial": "Codigo"}, inplace=True)
else:
    df_materiais_base = pd.DataFrame(columns=["Codigo", "Material", "Quantidade"])

st.info(f"⚠️ A tabela abaixo mostra os materiais de **UMA** estrutura. Ao salvar, eles serão multiplicados por {qtd_estrutura}.")
df_editavel = st.data_editor(df_materiais_base, num_rows="dynamic", use_container_width=True, hide_index=True, key=f"editor_{estrutura}")

if st.button("➕ Adicionar Esta Estrutura (em Lote)", type="primary"):
    st.session_state["carrinho"].append({
        "estrutura": estrutura, "acao": acao, "rede": tipo_rede, "execucao": sigla_exec,
        "cod_mo": cod_mo, "desc_mo": desc_mo, "valor_mo_unit": valor_previsto_unitario, 
        "qtd_estrutura": qtd_estrutura, "materiais": df_editavel.to_dict('records')
    })
    st.rerun()

# ==========================================
# 4. LANÇAMENTO DE CABOS
# ==========================================
if not df_cabos.empty:
    st.markdown("---")
    st.subheader("🧵 Lançamento e Retirada de Cabos")
    
    col_c1, col_c2, col_c3 = st.columns(3)
    with col_c1:
        cabo_selecionado = st.selectbox("Selecione o Cabo (Bitola/Tipo):", df_cabos["Descricao"].tolist())
    with col_c2:
        acao_cabo = st.selectbox("Ação do Cabo", ["Instalação", "Retirada"], key="acao_cabo_sel")
    with col_c3:
        rede_cabo = st.radio("Rede do Cabo:", ["MT", "BT"], horizontal=True, key="rede_cabo_sel")
        
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        tipo_medida = st.radio("Informar quantidade em:", ["Metros (m)", "Quilos (kg)"], horizontal=True)
    with col_v2:
        qte_cabo = st.number_input("Quantidade Total (Tamanho ou Peso):", min_value=0.0, step=1.0)
        
    cabo_info = df_cabos[df_cabos["Descricao"] == cabo_selecionado].iloc[0]
    fator = tratar_valor(cabo_info["FATOR kg/m"])
    cod_cabo = str(cabo_info["CÓDIGO"]) if pd.notna(cabo_info["CÓDIGO"]) else "S/C"
    
    metros_calc = 0.0
    quilos_calc = 0.0
    
    if tipo_medida == "Metros (m)":
        metros_calc = qte_cabo
        quilos_calc = qte_cabo * fator
    else:
        quilos_calc = qte_cabo
        metros_calc = qte_cabo / fator if fator > 0 else 0.0
        
    if qte_cabo > 0:
        st.info(f"⚖️ Conversão Automática do Kit Técnico: **{metros_calc:,.2f} m** equivalem a **{quilos_calc:,.2f} kg**")
        
    if st.button("➕ Adicionar Cabo à Obra"):
        if qte_cabo > 0:
            st.session_state["carrinho_cabos"].append({
                "codigo": cod_cabo,
                "descricao": cabo_selecionado,
                "acao": acao_cabo,
                "rede": rede_cabo,
                "metros": metros_calc,
                "quilos": quilos_calc
            })
            st.rerun()
        else:
            st.error("Informe uma quantidade maior que zero.")

# ==========================================
# 5. RESUMO VISUAL DO CARRINHO
# ==========================================
st.markdown("---")
st.subheader("🛒 Resumo das Inserções do Trecho")

if len(st.session_state["carrinho"]) > 0:
    st.write("🔧 **Estruturas (Com Multiplicador):**")
    for i, item in enumerate(st.session_state["carrinho"]):
        valor_total_item = item['valor_mo_unit'] * item['qtd_estrutura']
        st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;{i+1}. **{item['qtd_estrutura']}x {item['estrutura']}** ({item['rede']} - {item['execucao']}) | R$ {valor_total_item:,.2f} Total | *{len(item['materiais'])} materiais-base*")

if len(st.session_state["carrinho_cabos"]) > 0:
    st.write("🧵 **Cabos (Condutores):**")
    for i, item in enumerate(st.session_state["carrinho_cabos"]):
        st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;{i+1}. {item['descricao']} ({item['rede']} - {item['acao']}) | *{item['metros']:,.2f} m ({item['quilos']:,.2f} kg)*")

if len(st.session_state["carrinho"]) == 0 and len(st.session_state["carrinho_cabos"]) == 0:
    st.info("Nenhuma estrutura ou cabo adicionado ainda.")
else:
    if st.button("🗑️ Limpar Carrinhos"):
        st.session_state["carrinho"] = []
        st.session_state["carrinho_cabos"] = []
        st.rerun()

# ==========================================
# 6. FOTOS E OBSERVAÇÕES
# ==========================================
st.markdown("---")
st.subheader("📸 Fotos do Poste/Vão")

opcao_foto = st.radio("Como deseja enviar a foto?", ["Tirar foto com a Câmera", "Anexar arquivo da Galeria", "Não enviar foto"], horizontal=True)

foto_camera = None
fotos_galeria = []

if opcao_foto == "Tirar foto com a Câmera":
    st.warning("⚠️ Nota: Fotos tiradas por aqui podem ter qualidade reduzida pelo navegador.")
    foto_camera = st.camera_input("Tire uma foto do poste")
elif opcao_foto == "Anexar arquivo da Galeria":
    st.info("💡 Dica: Pelo celular, escolha 'Tirar Foto' ao anexar para usar a qualidade máxima. Você pode selecionar VÁRIAS fotos de uma vez!")
    fotos_galeria = st.file_uploader("Anexe as fotos (Obra, Poste, etc.)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

st.markdown("---")
st.subheader("📝 Observações Gerais do Poste/Vão")
observacao = st.text_area("Digite suas observações", key="input_obs_geral")

# ==========================================
# 7. SALVAMENTO FINAL DA PLANILHA OFICIAL E FOTOS
# ==========================================
st.markdown("<br>", unsafe_allow_html=True) 
salvar = st.button("✅ Salvar Lançamento do Trecho", use_container_width=True, type="primary")

if salvar:
    if len(st.session_state["carrinho"]) == 0 and len(st.session_state["carrinho_cabos"]) == 0 and not incluir_poste:
        st.error("⚠️ Adicione pelo menos uma estrutura, cabo ou poste limpo antes de salvar!")
    else:
        nome_poste = poste.strip().upper() if poste.strip() != "" else "GERAL"
        nome_obra_arquivo = obra.strip() if obra.strip() != "" else "SEM_NUMERO"
        
        linhas_para_salvar = []

        if incluir_poste:
            valor_poste_unit = 0.0
            poste_match = tb_mao_obra[tb_mao_obra["Cod_MO"].astype(str) == str(cod_poste)]
            if len(poste_match) > 0:
                valor_poste_unit = tratar_valor(poste_match.iloc[0]["Valor_Instalar"])

            linhas_para_salvar.append({
                "Data": datetime.now().strftime("%d/%m/%Y %H:%M"), "Obra": obra, "Municipio": municipio, "Poste": nome_poste,
                "Rede": "MT/BT", "Estrutura": "Poste Limpo", "Acao": "Instalação", "Execucao": "LM", "Tipo_Item": "Mão de Obra",
                "Codigo": cod_poste, "Descricao": desc_poste, "Quantidade": qtd_poste, 
                "Valor_Unitario": round(valor_poste_unit, 2), "Valor_Total": round(valor_poste_unit * qtd_poste, 2), "Observacao": observacao
            })

        for item in st.session_state["carrinho"]:
            # Adiciona a Mão de Obra multiplicada
            linhas_para_salvar.append({
                "Data": datetime.now().strftime("%d/%m/%Y %H:%M"), "Obra": obra, "Municipio": municipio, "Poste": nome_poste,
                "Rede": item["rede"], "Estrutura": item["estrutura"], "Acao": item["acao"], "Execucao": item["execucao"],
                "Tipo_Item": "Mão de Obra", "Codigo": item["cod_mo"], "Descricao": item["desc_mo"], "Quantidade": item["qtd_estrutura"],
                "Valor_Unitario": round(item["valor_mo_unit"], 2), "Valor_Total": round(item["valor_mo_unit"] * item["qtd_estrutura"], 2), 
                "Observacao": observacao
            })
            
            # Adiciona os Materiais multiplicados
            for mat in item["materiais"]:
                if pd.notna(mat.get("Material")) and str(mat.get("Material")).strip() != "":
                    qtd_mat_base = tratar_valor(mat.get("Quantidade", 0))
                    qtd_mat_total = qtd_mat_base * item["qtd_estrutura"] # Mágica da multiplicação em lote!
                    
                    linhas_para_salvar.append({
                        "Data": datetime.now().strftime("%d/%m/%Y %H:%M"), "Obra": obra, "Municipio": municipio, "Poste": nome_poste,
                        "Rede": item["rede"], "Estrutura": item["estrutura"], "Acao": item["acao"], "Execucao": item["execucao"],
                        "Tipo_Item": "Material", "Codigo": mat.get("Codigo", ""), "Descricao": mat.get("Material", ""),
                        "Quantidade": qtd_mat_total, "Valor_Unitario": 0.0, "Valor_Total": 0.0, "Observacao": observacao
                    })

        for cabo in st.session_state["carrinho_cabos"]:
            linhas_para_salvar.append({
                "Data": datetime.now().strftime("%d/%m/%Y %H:%M"), "Obra": obra, "Municipio": municipio, "Poste": nome_poste,
                "Rede": cabo["rede"], "Estrutura": "Lançamento de Condutor", "Acao": cabo["acao"], "Execucao": "LM",
                "Tipo_Item": "Material (Condutor)", "Codigo": cabo["codigo"], "Descricao": cabo["descricao"], 
                "Quantidade": f"{cabo['metros']:.2f} m ({cabo['quilos']:.2f} kg)", 
                "Valor_Unitario": 0.0, "Valor_Total": 0.0, "Observacao": observacao
            })

        novo_df = pd.DataFrame(linhas_para_salvar)
        arquivo = f"Fiscalizacao_Obra_{nome_obra_arquivo}.xlsx"

        if os.path.exists(arquivo):
            existente = pd.read_excel(arquivo, sheet_name=0)
            if "Gasto_Previsto_Obra" in existente.columns:
                existente = existente.drop(columns=["Gasto_Previsto_Obra"])
            final = pd.concat([existente, novo_df], ignore_index=True)
        else:
            final = novo_df

        final["Valor_Total"] = pd.to_numeric(final["Valor_Total"], errors="coerce").fillna(0.0)
        df_resumo = final.groupby("Obra", as_index=False)["Valor_Total"].sum()
        df_resumo.rename(columns={"Valor_Total": "VALOR PREVISTO DA OBRA"}, inplace=True)

        with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
            final.to_excel(writer, sheet_name="Dados", index=False)
            df_resumo.to_excel(writer, sheet_name="Resumo", index=False)

        # ----------------------------------------------------
        # SALVAMENTO DAS FOTOS
        # ----------------------------------------------------
        os.makedirs("fotos", exist_ok=True)
        
        if foto_camera is not None:
            data_hora = datetime.now().strftime('%Y%m%d_%H%M%S')
            nome_foto = f"fotos/Obra_{nome_obra_arquivo}_Poste_{nome_poste}_{data_hora}_Camera.jpg"
            with open(nome_foto, "wb") as f:
                f.write(foto_camera.getbuffer())

        if fotos_galeria:
            for idx, foto in enumerate(fotos_galeria):
                data_hora = datetime.now().strftime('%Y%m%d_%H%M%S')
                nome_foto = f"fotos/Obra_{nome_obra_arquivo}_Poste_{nome_poste}_{data_hora}_Galeria_Foto{idx+1}.jpg"
                with open(nome_foto, "wb") as f:
                    f.write(foto.getbuffer())

        st.session_state["carrinho"] = []
        st.session_state["carrinho_cabos"] = []
        st.success(f"Fiscalização do trecho {nome_poste} salva na Obra {nome_obra_arquivo} com sucesso!")

# ==========================================
# 8. BOTÕES DE DOWNLOAD (EXCEL E FOTOS)
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
