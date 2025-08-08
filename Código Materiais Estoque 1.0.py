import pandas as pd
import os
import re
from datetime import datetime

def processar_estoque_completo(caminho_entrada, caminho_saida, caminho_referencia):
    """
    Processa a planilha de estoque e a de referência para gerar o arquivo de controle final,
    adicionando uma coluna com a data e hora da atualização.
    """
    try:
        print("ETAPA 1: Lendo e processando a planilha de estoque...")
        df_estoque_raw = pd.read_excel(caminho_entrada, skiprows=5, header=None)
        
        # Lógica de processamento (sem alterações)
        dados_processados = []
        current_nm = None; current_desc = None; current_saldo_total = 0; current_unidade = ''
        padrao_nm = re.compile(r'^\d{2}\.\d{3}\.\d{3}$')
        for index, row in df_estoque_raw.iterrows():
            valor_col_b = str(row.get(1, '')).strip()
            if padrao_nm.match(valor_col_b):
                if current_nm:
                    dados_processados.append({'NM': current_nm, 'Descrição do Material': current_desc, 'Saldo do Estoque': current_saldo_total, 'Unidade de Medida': current_unidade})
                current_nm = valor_col_b; current_desc = row.get(6, ''); current_saldo_total = 0; current_unidade = ''
            elif current_nm:
                identificador_linha = str(row.get(4, '')).strip().upper()
                if identificador_linha and identificador_linha != 'BRL':
                    if not current_unidade: current_unidade = identificador_linha
                    valor_saldo_texto = str(row.get(3, '0')).replace('.', '').replace(',', '.')
                    saldo_linha = pd.to_numeric(valor_saldo_texto, errors='coerce')
                    if pd.notna(saldo_linha): current_saldo_total += saldo_linha
        if current_nm:
            dados_processados.append({'NM': current_nm, 'Descrição do Material': current_desc, 'Saldo do Estoque': current_saldo_total, 'Unidade de Medida': current_unidade})
        df_estoque = pd.DataFrame(dados_processados)

        print("ETAPA 2: Lendo a planilha de referência e juntando os dados...")
        df_referencia = pd.read_excel(caminho_referencia)

        df_estoque['NM'] = df_estoque['NM'].astype(str)
        df_referencia['NM'] = df_referencia['NM'].astype(str)
        
        df_final = pd.merge(df_estoque, df_referencia, on='NM', how='outer', suffixes=('', '_ref'))

        df_final['Descrição do Material'] = df_final['Descrição do Material'].fillna(df_final['Descrição do Material_ref'])
        df_final['Saldo do Estoque'] = df_final['Saldo do Estoque'].fillna(0)
        df_final['Unidade de Medida'] = df_final['Unidade de Medida'].fillna('UN')

        for col in ['MRP', 'Classe']:
            if col + '_ref' in df_final.columns:
                 df_final[col] = df_final[col].fillna(df_final[col + '_ref'])
            if col not in df_final.columns:
                 df_final[col] = ''
            df_final[col] = df_final[col].fillna('')
        
        agora = datetime.now()
        timestamp_str = agora.strftime("%d/%m/%Y %H:%M:%S")
        df_final['Última Atualização'] = timestamp_str

        colunas_finais = ['NM', 'Descrição do Material', 'Saldo do Estoque', 'Unidade de Medida', 'MRP', 'Classe', 'Última Atualização']
        df_final = df_final[colunas_finais].sort_values(by='Descrição do Material').reset_index(drop=True)

        df_final.to_excel(caminho_saida, index=False)
        print(f"SUCESSO: Arquivo '{caminho_saida}' foi criado/atualizado com sucesso.")

    except FileNotFoundError as e:
        print(f"ERRO: Arquivo não encontrado - {e.filename}. Verifique se o caminho e o nome do arquivo estão corretos.")
    except KeyError as e:
        print(f"ERRO: A coluna {e} não foi encontrada na sua planilha de referência. Verifique se o nome da coluna está escrito corretamente no arquivo Excel.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")

# --- PONTO DE EXECUÇÃO DO SCRIPT ---
if __name__ == '__main__':
    # --- CAMINHOS ABSOLUTOS AJUSTADOS CONFORME SOLICITADO ---
    
    # --- CORREÇÃO APLICADA AQUI ---
    # Caminho para o arquivo de entrada bruto, com o nome correto 'Materiais.xlsx'
    arquivo_entrada_local = r'C:\Users\E5QV\OneDrive - PETROBRAS\Documentos\GitHub\dashboard-estoque\Materiais.xlsx'
    
    # Caminho para o arquivo de referência
    arquivo_referencia_local = r'C:\Users\E5QV\OneDrive - PETROBRAS\Documentos\GitHub\dashboard-estoque\NM materiais do SMS SI.xlsx'
    
    # Caminho completo para o arquivo de saída
    pasta_saida = r'C:\Users\E5QV\OneDrive - PETROBRAS\Documentos\GitHub\dashboard-estoque'
    nome_arquivo_saida = 'Controle de Materiais Estoque.xlsx'
    arquivo_saida_local = os.path.join(pasta_saida, nome_arquivo_saida)
    
    processar_estoque_completo(arquivo_entrada_local, arquivo_saida_local, arquivo_referencia_local)