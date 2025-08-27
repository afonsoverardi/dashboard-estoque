@echo off
:: Muda o diretorio para a pasta do projeto
cd "C:\Users\E5QV\OneDrive - PETROBRAS\Documentos\GitHub\dashboard-estoque"

:: Ativa o ambiente virtual
echo Ativando o ambiente virtual...
call venv\Scripts\activate.bat

:: Executa o script usando o servidor do Streamlit
echo Iniciando o servidor do Streamlit...
streamlit run Dashboard.py