import pandas as pd
import os
import re

def processar_estoque_completo(caminho_entrada, caminho_saida, caminho_referencia):
    """
    Processa a planilha de estoque e a de referência para gerar o arquivo de controle final.
    Versão corrigida para o erro de KeyError e os FutureWarning.
    """
    try:
        print("ETAPA 1: Lendo e processando a planilha de estoque...")
        df_estoque_raw = pd.read_excel(caminho_entrada, skiprows=5, header=None)
        
        # Lógica de processamento do arquivo de materiais (sem alterações)
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

        # --- CORREÇÃO APLICADA AQUI ---
        # Corrigimos a lógica de preenchimento de dados e os avisos do pandas
        
        # Preenche a Descrição do Material (esta coluna pode existir em ambos os arquivos)
        df_final['Descrição do Material'] = df_final['Descrição do Material'].fillna(df_final['Descrição do Material_ref'])

        # Preenche os valores para itens que só existem na referência
        df_final['Saldo do Estoque'] = df_final['Saldo do Estoque'].fillna(0)
        df_final['Unidade de Medida'] = df_final['Unidade de Medida'].fillna('UN')

        # Garante que as colunas MRP e Classe existam e preenche valores nulos com texto vazio
        for col in ['MRP', 'Classe']:
            # Se a coluna de referência existir, usa-a para preencher NAs
            if col + '_ref' in df_final.columns:
                 df_final[col] = df_final[col].fillna(df_final[col + '_ref'])
            # Se a coluna ainda não existir, cria uma vazia
            if col not in df_final.columns:
                 df_final[col] = ''
            # Preenche quaisquer NAs restantes com vazio
            df_final[col] = df_final[col].fillna('')
        
        # Seleciona e ordena as colunas finais para o dashboard
        colunas_finais = ['NM', 'Descrição do Material', 'Saldo do Estoque', 'Unidade de Medida', 'MRP', 'Classe']
        df_final = df_final[colunas_finais].sort_values(by='NM').reset_index(drop=True)

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
    arquivo_entrada_local = os.path.join('Planilha Base', 'Materiais.xlsx')
    arquivo_referencia_local = 'NM materiais do SMS SI.xlsx'
    arquivo_saida_local = 'Controle de Materiais Estoque.xlsx'
    
    processar_estoque_completo(arquivo_entrada_local, arquivo_saida_local, arquivo_referencia_local)