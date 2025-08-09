import streamlit as st
import pandas as pd
import os
from PIL import Image, ImageOps

# --- Configuração da Página ---
st.set_page_config(layout="wide", page_title="Estoque de Materiais")
st.title("Visão Geral do Estoque da Segurança Ocupacional")

# --- Funções Auxiliares ---

# --- ALTERAÇÃO IMPORTANTE: REMOVEMOS O CACHE DAQUI ---
# A função agora simplesmente lê o arquivo, sem guardar na memória de longo prazo.
def carregar_dados():
    caminho_excel = 'Controle de Materiais Estoque.xlsx'
    try:
        df = pd.read_excel(caminho_excel)
        if 'Classe' not in df.columns: df['Classe'] = 'Sem Classe'
        else: df['Classe'] = df['Classe'].fillna('Sem Classe')
        if 'MRP' not in df.columns: df['MRP'] = 'N/A'
        else: df['MRP'] = df['MRP'].fillna('N/A')
        if 'Última Atualização' not in df.columns: df['Última Atualização'] = "Não disponível"
        else: df['Última Atualização'] = df['Última Atualização'].fillna("Não disponível")
        return df
    except FileNotFoundError:
        st.error(f"ERRO: A planilha de dados '{caminho_excel}' não foi encontrada no repositório.")
        return None
    except Exception as e:
        st.error(f"ERRO ao carregar a planilha: {e}")
        return None

def padronizar_imagem(caminho, tamanho_final=(220, 220)):
    # ... (código igual, sem alterações)
    placeholder_url = f"https://placehold.co/{tamanho_final[0]}x{tamanho_final[1]}/333333/FAFAFA?text=Sem+Foto"
    if not caminho or not os.path.exists(caminho): return placeholder_url
    try:
        img = Image.open(caminho)
        img.thumbnail(tamanho_final)
        img_padronizada = ImageOps.pad(img, tamanho_final, color='#262730')
        return img_padronizada
    except Exception: return placeholder_url

def criar_cartao_material(item):
    # ... (código igual, sem alterações)
    with st.container(border=True, height=420):
        st.image(item['imagem_objeto'], use_container_width=True)
        st.markdown(f"<strong>{item['Descrição do Material']}</strong>", unsafe_allow_html=True)
        col_info, col_zoom = st.columns([4, 1])
        with col_info:
            st.caption(f"NM: {item['NM']} | MRP: {item['MRP']}")
            estoque_html = f"""<p style="font-size: 0.9em; color: #FAFAFA; margin-bottom: 0;"><strong>Estoque:</strong> {formatar_estoque(item['Saldo do Estoque'])} {item['Unidade de Medida']}</p>"""
            st.markdown(estoque_html, unsafe_allow_html=True)
        with col_zoom:
            if st.button("🔍", key=f"zoom_{item['NM']}", help="Ampliar imagem"):
                st.session_state.item_para_zoom = item['NM']
                st.rerun()

def formatar_estoque(numero):
    # ... (código igual, sem alterações)
    if isinstance(numero, float) and numero.is_integer(): return f"{int(numero)}"
    if isinstance(numero, float): return str(numero).replace('.', ',')
    return str(numero)

# --- Lógica Principal do Dashboard ---

# --- NOVA LÓGICA DE CACHE COM SESSION_STATE ---
# Verifica se os dados já foram carregados nesta sessão.
if 'df_principal' not in st.session_state:
    # Se não foram, chama a função para ler o arquivo Excel e guarda na memória da sessão.
    st.session_state.df_principal = carregar_dados()

# Usa o dataframe da memória da sessão para o resto do app.
df = st.session_state.df_principal

# Inicializa o estado do zoom
if 'item_para_zoom' not in st.session_state:
    st.session_state.item_para_zoom = None
    
if df is not None:
    df = df.sort_values(by='Descrição do Material').reset_index(drop=True)
    df['NM'] = df['NM'].astype(str)
    caminho_base_imagens = "Imagens"
    def obter_caminho_real(nm):
        nome_arquivo = str(nm).replace('.', '-')
        caminho_jpg = os.path.join(caminho_base_imagens, f"{nome_arquivo}.jpg")
        if os.path.exists(caminho_jpg): return caminho_jpg
        caminho_png = os.path.join(caminho_base_imagens, f"{nome_arquivo}.png")
        if os.path.exists(caminho_png): return caminho_png
        return None
    df['caminho_original'] = df['NM'].apply(obter_caminho_real)
    df['imagem_objeto'] = df['caminho_original'].apply(padronizar_imagem)

    if st.session_state.item_para_zoom:
        item_selecionado = df[df['NM'] == st.session_state.item_para_zoom].iloc[0]
        st.header(f"Detalhe: {item_selecionado['Descrição do Material']}")
        if st.button("⬅️ Voltar para a Galeria"):
            st.session_state.item_para_zoom = None
            st.rerun()
        st.image(item_selecionado['caminho_original'], width=1200)
    else:
        with st.sidebar:
            try:
                logo = Image.open("petrobras_logo.png")
                st.image(logo, use_container_width=True)
            except FileNotFoundError: st.error("Logo não encontrada.")
            if not df.empty:
                ultima_atualizacao = df['Última Atualização'].iloc[0]
                st.caption(f"Dados atualizados em:  \n{ultima_atualizacao}")
            st.header("Filtros")
            termo_busca = st.text_input("Buscar por Descrição:")
            st.subheader("Filtrar por Classe")
            classes_unicas = sorted(df['Classe'].unique())
            with st.expander("Selecionar Classes", expanded=True):
                selecionar_todas_classes = st.checkbox("Selecionar Todas", value=True, key='select_all_classes')
                classes_selecionadas = [cls for cls in classes_unicas if st.checkbox(cls, value=selecionar_todas_classes, key=f"check_{cls}")]
            st.subheader("Filtrar por MRP")
            df_filtrado_por_classe = df[df['Classe'].isin(classes_selecionadas)] if classes_selecionadas else df
            mrps_disponiveis = sorted(df_filtrado_por_classe['MRP'].unique())
            with st.expander("Selecionar MRPs", expanded=True):
                selecionar_todos_mrps = st.checkbox("Selecionar Todos", value=True, key='select_all_mrps')
                mrps_selecionados = [mrp for mrp in mrps_disponiveis if st.checkbox(mrp, value=selecionar_todos_mrps, key=f"check_{mrp}")]
        
        df_filtrado = df
        if classes_selecionadas: df_filtrado = df_filtrado[df_filtrado['Classe'].isin(classes_selecionadas)]
        if mrps_selecionados: df_filtrado = df_filtrado[df_filtrado['MRP'].isin(mrps_selecionados)]
        if termo_busca: df_filtrado = df_filtrado[df_filtrado['Descrição do Material'].str.contains(termo_busca, case=False)]
        
        st.caption(f"Exibindo {len(df_filtrado)} de {len(df)} itens.")
        st.divider()

        if df_filtrado.empty:
            st.warning("Nenhum item corresponde aos filtros selecionados.")
        else:
            classes_para_exibir = sorted(df_filtrado['Classe'].unique())
            for classe in classes_para_exibir:
                with st.expander(f"**Classe: {classe}** ({len(df_filtrado[df_filtrado['Classe'] == classe])} itens)", expanded=True):
                    df_da_classe = df_filtrado[df_filtrado['Classe'] == classe]
                    num_colunas = 7
                    cols = st.columns(num_colunas)
                    for index, item in df_da_classe.reset_index(drop=True).iterrows():
                        col_atual = cols[index % num_colunas]
                        with col_atual:
                            criar_cartao_material(item)
else:
    st.warning("Aguardando o carregamento dos dados. Por favor, gere e envie a planilha 'Controle de Materiais Estoque.xlsx' para o repositório.")