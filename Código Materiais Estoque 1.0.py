import pandas as pd
import os
import re
from datetime import datetime

def processar_estoque_completo(caminho_entrada, caminho_saida, caminho_referencia):
    """
    Processa a planilha de estoque usando groupby para garantir a agregação correta dos saldos.
    """
    try:
        print("ETAPA 1: Lendo a planilha de estoque bruta...")
        df_raw = pd.read_excel(caminho_entrada, skiprows=5, header=None)
        
        # Atribui nomes temporários às colunas para facilitar o acesso
        num_colunas = len(df_raw.columns)
        df_raw.columns = [f'col_{i}' for i in range(num_colunas)]

        print("ETAPA 2: Preparando e limpando os dados...")
        # Cria uma coluna 'NM' identificando as linhas que são "cabeçalho" de material
        df_raw['NM'] = df_raw['col_1'].where(df_raw['col_1'].astype(str).str.match(r'^\d{2}\.\d{3}\.\d{3}$'))
        
        # Propaga o último NM válido para as linhas de lote abaixo dele
        df_raw['NM'].ffill(inplace=True)
        
        # Cria a coluna 'Descrição do Material' da mesma forma
        df_raw['Descrição do Material'] = df_raw['col_6'].where(df_raw['NM'] == df_raw['col_1'])
        df_raw['Descrição do Material'].ffill(inplace=True)

        # Filtra para manter apenas as linhas que parecem ser de estoque (ignora linhas BRL e outras)
        # Uma linha de estoque válida tem um valor na coluna 4 que NÃO é 'BRL' e tem um saldo numérico na coluna 3
        df_dados = df_raw[df_raw['col_4'].astype(str).str.upper().str.strip() != 'BRL'].copy()
        
        # Converte a coluna de saldo para um formato numérico, tratando erros
        saldo_str = df_dados['col_3'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        df_dados['Saldo'] = pd.to_numeric(saldo_str, errors='coerce')
        
        # Remove linhas que não puderam ser convertidas para um saldo numérico (ex: linhas de cabeçalho)
        df_dados.dropna(subset=['Saldo'], inplace=True)
        
        # Renomeia a coluna da unidade de medida
        df_dados.rename(columns={'col_4': 'Unidade de Medida'}, inplace=True)

        print("ETAPA 3: Agrupando e somando os saldos...")
        # Define as regras de agregação: somar os saldos e pegar a primeira unidade de medida encontrada
        aggregation_rules = {
            'Saldo': 'sum',
            'Unidade de Medida': 'first' # Pega a primeira unidade válida (UN, L, etc.)
        }
        # A mágica do groupby: agrupa por NM e Descrição e aplica as regras
        df_estoque = df_dados.groupby(['NM', 'Descrição do Material']).agg(aggregation_rules).reset_index()
        df_estoque.rename(columns={'Saldo': 'Saldo do Estoque'}, inplace=True)

        print("ETAPA 4: Juntando com a planilha de referência...")
        df_referencia = pd.read_excel(caminho_referencia)
        df_estoque['NM'] = df_estoque['NM'].astype(str)
        df_referencia['NM'] = df_referencia['NM'].astype(str)
        
        df_final = pd.merge(df_estoque, df_referencia, on='NM', how='outer', suffixes=('', '_ref'))

        # Preenche e limpa os dados finais
        df_final['Descrição do Material'].fillna(df_final['Descrição do Material_ref'], inplace=True)
        df_final['Saldo do Estoque'].fillna(0, inplace=True)
        df_final['Unidade de Medida'].fillna('UN', inplace=True)
        for col in ['MRP', 'Classe']:
            if col + '_ref' in df_final.columns:
                 df_final[col] = df_final[col].fillna(df_final[col + '_ref'])
            if col not in df_final.columns: df_final[col] = ''
            df_final[col] = df_final[col].fillna('')
        
        # Adiciona data de atualização e finaliza
        timestamp_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        df_final['Última Atualização'] = timestamp_str
        colunas_finais = ['NM', 'Descrição do Material', 'Saldo do Estoque', 'Unidade de Medida', 'MRP', 'Classe', 'Última Atualização']
        df_final = df_final[colunas_finais].sort_values(by='Descrição do Material').reset_index(drop=True)

        df_final.to_excel(caminho_saida, index=False)
        print(f"SUCESSO: Arquivo '{caminho_saida}' foi criado/atualizado com sucesso e sem duplicatas.")

    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")

# --- PONTO DE EXECUÇÃO DO SCRIPT ---
if __name__ == '__main__':
    arquivo_entrada_local = r'C:\Users\E5QV\OneDrive - PETROBRAS\Documentos\GitHub\dashboard-estoque\Materiais.xlsx'
    arquivo_referencia_local = r'C:\Users\E5QV\OneDrive - PETROBRAS\Documentos\GitHub\dashboard-estoque\NM materiais do SMS SI.xlsx'
    pasta_saida = r'C:\Users\E5QV\OneDrive - PETROBRAS\Documentos\GitHub\dashboard-estoque'
    nome_arquivo_saida = 'Controle de Materiais Estoque.xlsx'
    arquivo_saida_local = os.path.join(pasta_saida, nome_arquivo_saida)
    
    processar_estoque_completo(arquivo_entrada_local, arquivo_saida_local, arquivo_referencia_local)