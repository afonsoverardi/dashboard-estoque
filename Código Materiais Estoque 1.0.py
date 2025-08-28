import pandas as pd
import os
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from getpass import getpass # Para pedir a senha de forma segura
import time

def extrair_dados_sap_bw(url_query, usuario, senha):
    """
    Abre um navegador Chrome, faz login no portal SAP e extrai a tabela de dados da query.
    """
    print("Iniciando a automação do navegador...")
    
    # Configura o navegador Chrome para ser controlado pelo script
    # O webdriver-manager cuida de baixar o driver correto automaticamente
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    
    try:
        # Abre a URL da query. O portal SAP geralmente redireciona para a página de login.
        driver.get(url_query)
        
        # Espera um pouco para a página de login carregar
        print("Aguardando página de login...")
        time.sleep(5) # Pausa de 5 segundos

        # --- SEÇÃO DE LOGIN ---
        # ATENÇÃO: Os 'IDs' dos campos ('logonuidfield', 'logonpassfield') podem variar.
        # Se o login falhar, esta é a parte que precisa ser ajustada.
        print(f"Preenchendo usuário: {usuario}")
        driver.find_element("id", "logonuidfield").send_keys(usuario)
        
        print("Preenchendo senha...")
        driver.find_element("id", "logonpassfield").send_keys(senha)
        
        # Clica no botão de logon
        driver.find_element("id", "logon_button").click()
        
        # Espera o carregamento da página da query após o login
        print("Login realizado. Aguardando o carregamento dos dados da query...")
        time.sleep(10) # Pausa maior para garantir que a tabela de dados carregue

        # --- SEÇÃO DE EXTRAÇÃO DE DADOS ---
        # O pandas consegue ler as tabelas diretamente da página HTML que o navegador está vendo
        print("Lendo tabelas da página...")
        tabelas_na_pagina = pd.read_html(driver.page_source)
        
        # Geralmente, a tabela principal de dados é uma das primeiras. 
        # Pode ser necessário ajustar o índice [0] se houver outras tabelas na página.
        if tabelas_na_pagina:
            print("Tabela encontrada! Extraindo dados...")
            df_bruto_sap = tabelas_na_pagina[0] # Pega a primeira tabela encontrada
            return df_bruto_sap
        else:
            print("ERRO: Nenhuma tabela de dados foi encontrada na página após o login.")
            return None

    finally:
        # Garante que o navegador seja fechado no final, mesmo que ocorra um erro
        print("Fechando o navegador.")
        driver.quit()


def processar_estoque_completo(df_estoque_raw, caminho_saida, caminho_referencia):
    # ... (O resto do seu script de processamento continua aqui,
    # mas agora ele recebe os dados brutos como um argumento)
    try:
        print("ETAPA 2: Lendo a planilha de referência e juntando os dados...")
        df_referencia = pd.read_excel(caminho_referencia)
        
        # A partir daqui, a lógica é a mesma, mas adaptada para a estrutura da tabela do SAP BW.
        # **ESTA PARTE PRECISA SER AJUSTADA QUANDO TIVERMOS UMA AMOSTRA DA TABELA DO BW**
        # Por enquanto, vou manter a lógica antiga como um placeholder.
        
        # Exemplo: df_estoque = adaptar_dados_bw(df_estoque_raw)
        # Onde 'adaptar_dados_bw' seria uma nova função para limpar e somar os saldos
        # da tabela extraída. Sem ver a tabela, é impossível escrever essa função.
        
        # **Por enquanto, vamos assumir que a tabela do BW já vem semi-pronta**
        # E que as colunas se chamam 'NM', 'Descrição do Material', 'Saldo do Estoque', etc.
        # Esta é uma suposição que precisará ser validada.
        
        df_estoque = df_estoque_raw # Placeholder
        
        df_estoque['NM'] = df_estoque['NM'].astype(str)
        df_referencia['NM'] = df_referencia['NM'].astype(str)
        df_final = pd.merge(df_estoque, df_referencia, on='NM', how='outer', suffixes=('', '_ref'))
        
        # ... (Resto da lógica de preenchimento e salvamento) ...
        # ... (esta parte provavelmente precisará de ajustes mínimos) ...
        df_final['Descrição do Material'].fillna(df_final['Descrição do Material_ref'], inplace=True)
        df_final['Saldo do Estoque'].fillna(0, inplace=True)
        # ... etc ...

        print(f"SUCESSO: Arquivo '{caminho_saida}' foi criado/atualizado com sucesso.")

    except Exception as e:
        print(f"Ocorreu um erro inesperado durante o processamento: {e}")


# --- PONTO DE EXECUÇÃO DO SCRIPT ---
if __name__ == '__main__':
    # URL da sua query do BEx Web
    url_bw = "https://bi.petrobras.com.br/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=003YQP9UVRYL7EMJ3FV31M4CK"
    
    # Pede as credenciais de forma segura
    usuario_sap = input("Digite seu usuário de rede Petrobras (chave): ")
    senha_sap = getpass("Digite sua senha de rede: ")
    
    # ETAPA 1: Extrai os dados brutos do SAP BW
    df_bruto = extrair_dados_sap_bw(url_bw, usuario_sap, senha_sap)
    
    if df_bruto is not None:
        # ETAPA 2: Processa os dados extraídos
        caminho_referencia_local = r'C:\Users\E5QV\OneDrive - PETROBRAS\Documentos\GitHub\dashboard-estoque\NM materiais do SMS SI.xlsx'
        pasta_saida = r'C:\Users\E5QV\OneDrive - PETROBRAS\Documentos\GitHub\dashboard-estoque'
        nome_arquivo_saida = 'Controle de Materiais Estoque.xlsx'
        caminho_saida_local = os.path.join(pasta_saida, nome_arquivo_saida)
        
        # Por enquanto, vamos apenas salvar os dados brutos extraídos para análise
        print("Salvando dados brutos extraídos para análise em 'dados_brutos_sap.xlsx'...")
        df_bruto.to_excel("dados_brutos_sap.xlsx", index=False)
        print("Arquivo de dados brutos salvo! Por favor, verifique-o para ajustarmos a lógica de processamento.")
        
        # A linha abaixo está comentada. Precisamos primeiro ver como são os dados brutos.
        # processar_estoque_completo(df_bruto, caminho_saida_local, caminho_referencia_local)