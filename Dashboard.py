import streamlit as st
import pandas as pd
import os
from PIL import Image, ImageOps

# --- Configuração da Página ---
st.set_page_config(layout="wide", page_title="Estoque de Materiais")
st.title("Visão Geral do Estoque")

# --- Funções Auxiliares ---
@st.cache_data
def carregar_dados():
    # --- CORREÇÃO APLICADA AQUI ---
    # Trocamos o caminho absoluto por um caminho relativo.
    # Agora, o script procurará o arquivo na mesma pasta onde ele está rodando no servidor.
    caminho_excel = 'Controle de Materiais Estoque.xlsx'
    
    try:
        df = pd.read_excel(caminho_excel)
        if 'Classe' not in df.columns: df['Classe'] = 'Sem Classe'
        else: df['Classe'] = df['Classe'].fillna('Sem Classe')
        if 'MRP' not in df.columns: df['MRP'] = 'N/A'
        else: df['MRP'] = df['MRP'].fillna('N/A')
        return df
    except FileNotFoundError:
        # A mensagem de erro agora mostrará o nome do arquivo correto que ele está procurando
        st.error(f"ERRO: A planilha de dados '{caminho_excel}' não foi encontrada no repositório do GitHub.")
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
    with st.container(border=True, height=420): 
        st.image(item['imagem_objeto'], use_container_width=True)
        st.markdown(f"<strong>{item['Descrição do Material']}</strong>", unsafe_allow_html=True)
        st.caption(f"NM: {item['NM']} | MRP: {item['MRP']}\n**Estoque:** {item['Saldo do Estoque']} {item['Unidade de Medida']}")

# --- Lógica Principal do Dashboard ---
df = carregar_dados()

if df is not None:
    df['NM'] = df['NM'].astype(str)
    caminho_base_imagens = "Imagens"
    
    def obter_caminho_real(nm):
        nome_arquivo = str(nm).replace('.', '-')
        caminho_jpg = os.path.join(caminho_base_imagens, f"{nome_arquivo}.jpg")
        if os.path.exists(caminho_jpg): return caminho_jpg
        caminho_png = os.path.join(caminho_base_imagens, f"{nome_arquivo}.png")
        if os.path.exists(caminho_png): return caminho_png
        return None

    df['imagem_objeto'] = df['NM'].apply(obter_caminho_real).apply(padronizar_imagem)
    
    st.sidebar.header("Filtros")
    termo_busca = st.sidebar.text_input("Buscar por Descrição:")
    st.sidebar.subheader("Filtrar por Classe")
    classes_unicas = sorted(df['Classe'].unique())
    with st.sidebar.expander("Selecionar Classes", expanded=True):
        selecionar_todas_classes = st.checkbox("Selecionar Todas", value=True, key='select_all_classes')
        classes_selecionadas = [cls for cls in classes_unicas if st.checkbox(cls, value=selecionar_todas_classes, key=f"check_{cls}")]
    
    st.sidebar.subheader("Filtrar por MRP")
    df_filtrado_por_classe = df[df['Classe'].isin(classes_selecionadas)] if classes_selecionadas else df
    mrps_disponiveis = sorted(df_filtrado_por_classe['MRP'].unique())
    with st.sidebar.expander("Selecionar MRPs", expanded=True):
        selecionar_todos_mrps = st.checkbox("Selecionar Todos", value=True, key='select_all_mrps')
        mrps_selecionados = [mrp for mrp in mrps_disponiveis if st.checkbox(mrp, value=selecionar_todos_mrps, key=f"check_{mrp}")]
    
    df_filtrado = df
    if classes_selecionadas: df_filtrado = df_filtrado[df_filtrado['Classe'].isin(classes_selecionadas)]
    if mrps_selecionados: df_filtrado = df_filtrado[df_filtrado['MRP'].isin(mrps_selecionados)]
    if termo_busca: df_filtrado = df_filtrado[df_filtrado['Descrição do Material'].str.contains(termo_busca, case=False)]
    
    st.info(f"Exibindo {len(df_filtrado)} de {len(df)} itens.")
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