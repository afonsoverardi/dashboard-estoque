import streamlit as st
import pandas as pd
import os
from PIL import Image, ImageOps

# --- Configuração da Página ---
st.set_page_config(layout="wide", page_title="Estoque de Materiais")

# --- Título Principal ---
st.title("Visão Geral do Estoque")

# --- Inicialização do Estado da Sessão ---
if 'item_para_zoom' not in st.session_state:
    st.session_state.item_para_zoom = None

# --- Funções Auxiliares ---
@st.cache_data
def carregar_dados():
    caminho_excel = 'Controle de Materiais Estoque.xlsx'
    try:
        df = pd.read_excel(caminho_excel)
        if 'Classe' not in df.columns: df['Classe'] = 'Sem Classe'
        else: df['Classe'] = df['Classe'].fillna('Sem Classe')
        if 'MRP' not in df.columns: df['MRP'] = 'N/A'
        else: df['MRP'] = df['MRP'].fillna('N/A')
        return df
    except FileNotFoundError:
        st.error(f"ERRO: A planilha de dados '{caminho_excel}' não foi encontrada no repositório.")
        return None
    except Exception as e:
        st.error(f"ERRO ao carregar a planilha: {e}")
        return None

def padronizar_imagem(caminho, tamanho_final=(220, 220)):
    placeholder_url = f"https://placehold.co/{tamanho_final[0]}x{tamanho_final[1]}/333333/FAFAFA?text=Sem+Foto"
    if not caminho or not os.path.exists(caminho): return placeholder_url
    try:
        img = Image.open(caminho)
        img.thumbnail(tamanho_final)
        img_padronizada = ImageOps.pad(img, tamanho_final, color='#262730')
        return img_padronizada
    except Exception: return placeholder_url

def criar_cartao_material(item):
    """Cria o cartão com o ícone de zoom."""
    with st.container(border=True, height=420):
        st.image(item['imagem_objeto'], use_container_width=True)
        st.markdown(f"<strong>{item['Descrição do Material']}</strong>", unsafe_allow_html=True)

        col_info, col_zoom = st.columns([4, 1])
        with col_info:
            st.caption(f"NM: {item['NM']} | MRP: {item['MRP']}\n**Estoque:** {item['Saldo do Estoque']} {item['Unidade de Medida']}")
        with col_zoom:
            if st.button("🔍", key=f"zoom_{item['NM']}", help="Ampliar imagem"):
                st.session_state.item_para_zoom = item['NM']
                st.rerun()

# --- Lógica Principal do Dashboard ---
df = carregar_dados()

if df is not None:
    # Preparação dos dados de imagem
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

    # LÓGICA DE EXIBIÇÃO: ZOOM OU GALERIA
    if st.session_state.item_para_zoom:
        item_selecionado = df[df['NM'] == st.session_state.item_para_zoom].iloc[0]
        
        st.header(f"Detalhe: {item_selecionado['Descrição do Material']}")
        
        if st.button("⬅️ Voltar para a Galeria"):
            st.session_state.item_para_zoom = None
            st.rerun()
            
        st.image(item_selecionado['caminho_original'], width=1200)

    else:
        # --- FILTROS E LOGO NA BARRA LATERAL ---
        with st.sidebar:
            try:
                logo = Image.open("petrobras_logo.png")
                # --- CORREÇÃO APLICADA AQUI ---
                st.image(logo, use_container_width=True)
            except FileNotFoundError:
                st.error("Logo não encontrada.")
            
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
        
        # Aplicação dos filtros
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
                    num_colunas = 6
                    cols = st.columns(num_colunas)
                    for index, item in df_da_classe.reset_index(drop=True).iterrows():
                        col_atual = cols[index % num_colunas]
                        with col_atual:
                            criar_cartao_material(item)
else:
    st.warning("Aguardando o carregamento dos dados.")