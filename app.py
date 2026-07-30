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
# NOVA MEMÓRIA: O Carrinho de Compras do Poste
if "carrinho" not in st.session_state:
    st.session_state["carrinho"] = []

def limpar_tudo():
    st.session_state["input_obra"] = ""
    st.session_state["input_municipio"] = ""
    st.session_state["input_poste"] = ""
    st.session_state["input_obs"] = ""
    st.session_state["carrinho"] = [] # Limpa o carrinho também

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

tipo_rede = st.radio("Selecione o Tipo de Rede da Estrutura:", ["MT", "BT"], horizontal=True)
tb_comp_filtrada = tb_composicao[tb_composicao["Rede"] == tipo_rede]

col_e1, col_e2 = st.columns(2)
estruturas_disponiveis = sorted(tb_comp_filtrada["Estrutura"].dropna().unique())

with col_e1:
    estrutura = st.selectbox("Estrutura Adicional", estruturas_disponiveis)
with col_e2:
    acao = st.selectbox("Ação da Estrutura", ["Instalação", "Retirada"])

# Busca a Mão de Obra
mo_info = tb_mao_obra[tb_mao_obra["Estrutura"] == estrutura]
if len(mo_info) > 0:
    cod_mo = mo_info.iloc[0]["Cod_MO"]
    desc_mo = mo_info.iloc[0]["Descricao_MO"]
else:
    cod_mo = "MO_EXTRA"
    desc_mo = f"Mão de Obra para {estrutura}"

st.success(f"**Mão de Obra identificada:** {cod_mo} - {desc_mo} ({acao})")

# Busca Materiais
resultado_materiais = tb_comp_filtrada[tb_comp_filtrada["Estrutura"] == estrutura]
if len(resultado_materiais) > 0:
    df_materiais_base = resultado_materiais[["CodMaterial", "Material", "Quantidade"]].copy()
    df_materiais_base.rename(columns={"CodMaterial": "Codigo"}, inplace=True)
else:
    st.warning("Materiais não encontrados para esta estrutura. Insira manualmente abaixo:")
    df_materiais_base = pd.DataFrame(columns=["Codigo", "Material", "Quantidade"])

st.info("💡 **Dica:** Revise/altere os materiais abaixo e clique em 'Adicionar Esta Estrutura' para guardar no poste.")

df_editavel = st.data_editor(
    df_materiais_base,
    num_rows="dynamic", 
    use_container_width=True,
    key=f"editor_{estrutura}" # Key dinâmica para não dar conflito ao mudar a seleção
)

# --- BOTÃO DE ADICIONAR AO CARRINHO ---
if st.button("➕ Adicionar Esta Estrutura ao Poste", type="primary"):
    item_carrinho = {
        "estrutura": estrutura,
        "acao": acao,
        "rede": tipo_rede,
        "cod_mo": cod_mo,
        "desc_mo": desc_mo,
        "materiais": df_editavel.to_dict('records') # Transforma o DataFrame editado em uma lista de dicionários
    }
    st.session_state["carrinho"].append(item_carrinho)
    st.rerun() # Atualiza a tela para mostrar o carrinho cheio

# ==========================================
# 4. RESUMO DO POSTE (O CARRINHO VISUAL)
# ==========================================
st.markdown("---")
st.subheader("🛒 Estruturas Adicionadas neste Poste")

if len(st.session_state["carrinho"]) == 0:
    st.info("Nenhuma estrutura adicionada ainda. Preencha acima e clique em 'Adicionar Esta Estrutura'.")
else:
    # Mostra os itens que o fiscal já adicionou
    for i, item in enumerate(st.session_state["carrinho"]):
        st.write(f"**{i+1}. Estrutura {item['estrutura']}** ({item['rede']} - {item['acao']}) | *{len(item['materiais'])} materiais listados*")
    
    # Botão para esvaziar o carrinho caso o fiscal tenha errado
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
# 6. SALVAMENTO FINAL DO POSTE INTEIRO
# ==========================================
st.markdown("<br>", unsafe_allow_html=True) 
salvar = st.button("✅ Salvar Fiscalização Completa do Poste", use_container_width=True, type="primary")

if salvar:
    # Se ele não marcou o poste e não tem nada no carrinho, avisa
    if len(st.session_state["carrinho"]) == 0 and not incluir_poste:
        st.error("⚠️ Adicione pelo menos uma estrutura ou marque a inclusão do poste limpo antes de salvar!")
    else:
        nome_poste = poste.strip().upper() if poste.strip() != "" else "GERAL"
        nome_obra_arquivo = obra.strip() if obra.strip() != "" else "SEM_NUMERO"
        
        linhas_para_salvar = []

        # 1. Adiciona a linha da Mão de Obra do Poste (Se a caixinha estiver marcada)
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

        # 2. Percorre o Carrinho e adiciona tudo que está nele
        for item in st.session_state["carrinho"]:
            
            # 2.1 Mão de Obra da Estrutura atual do carrinho
            dados_mo_est = {
                "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "Obra": obra,
                "Municipio": municipio,
                "Poste": nome_poste,
                "Estrutura": item["estrutura"],
                "Acao": item["acao"],
                "Tipo_Item": "Mão de Obra",
                "Codigo": item["cod_mo"],
                "Descricao": item["desc_mo"],
                "Quantidade": 1,
                "Observacao": observacao
            }
            linhas_para_salvar.append(dados_mo_est)
            
            # 2.2 Materiais da Estrutura atual do carrinho
            for mat in item["materiais"]:
                # Pula linhas vazias
                if pd.notna(mat.get("Material")) and str(mat.get("Material")).strip() != "":
                    dados_mat = {
                        "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                        "Obra": obra,
                        Achei a ideia do **"Carrinho de Compras" (Opção 1) fantástica e definitivamente o melhor caminho a seguir.** 

Você tocou no ponto exato: para um aplicativo de fiscalização, a usabilidade em campo é tão importante quanto a lógica do código. O fiscal geralmente está sob o sol, segurando um tablet ou celular, e minimizar a quantidade de cliques e recarregamentos da página faz uma diferença enorme na produtividade.

Aqui estão os principais motivos pelos quais a Opção 1 é superior para o seu cenário no setor de distribuição elétrica:

### 1. Escalabilidade Dinâmica
Postes no mundo real podem ser imprevisíveis. Embora o padrão possa ser uma estrutura de Média Tensão (MT) e uma de Baixa Tensão (BT), o fiscal pode se deparar com cruzamentos, derivações ou equipamentos adicionais (como chaves fusíveis, religadores ou iluminação pública) no mesmo poste. 
* Se você usar a Opção 2 (blocos fixos), ficará limitado a 2 ou 3 espaços. 
* Com o "carrinho", o fiscal pode adicionar 1, 3 ou 5 estruturas ao mesmo poste sem poluir a interface.

### 2. Integridade dos Dados
Ao agrupar tudo antes de salvar, você garante que os dados do "cabeçalho" (Obra, Município, ID do Poste) sejam amarrados perfeitamente a todas as estruturas daquele poste de uma só vez na sua planilha. Isso evita que o fiscal esqueça de salvar uma das estruturas ou altere acidentalmente o número do poste no meio do processo.

### 3. Melhoria na Interface (UX)
Você pode exibir um pequeno resumo visual (uma tabela ou uma lista) logo abaixo dos botões de seleção, mostrando o que já está no "carrinho" daquele poste (ex: `[MT - N4]`, `[BT - SI3]`). Isso dá ao fiscal a confirmação visual de que ele mapeou tudo corretamente antes de enviar para o banco de dados ou Excel.

---

### O Segredo no Streamlit: `st.session_state`
Para que essa Opção 1 funcione perfeitamente no Streamlit, o segredo é utilizar o gerenciamento de estado da sessão (`st.session_state`). Como o Streamlit recarrega o código a cada interação, você precisará de uma variável na sessão (por exemplo, uma lista vazia chamada `st.session_state['estruturas_poste']`) para guardar temporariamente a N4 e a SI3 até que o botão final de "Salvar Fiscalização" seja clicado, descarregando essa lista na sua planilha e limpando o carrinho para o próximo poste.

Você já tem familiaridade com o uso do `st.session_state` no Streamlit para criar esse agrupamento temporário, ou gostaria que eu montasse um esboço do código Python mostrando como implementar essa lógica do carrinho?
