import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import io
import os

st.set_page_config(page_title="Transferência de Dados de Câmbio", layout="wide")
st.title("Transferência de Dados de Câmbio")

# Função para ler os dados do arquivo de origem
def ler_dados_cambio(arquivo_cambio, data_inicial, data_final):
    try:
        # Carregar o arquivo Excel
        # Usar o pandas diretamente para simplificar
        df = pd.read_excel(
            arquivo_cambio, 
            sheet_name="BGP e BGX Cambio",
            engine="openpyxl"
        )
        
        # Verificar se as colunas necessárias existem
        try:
            # Tentar usar os nomes de colunas
            dados = pd.DataFrame({
                "Data": df.iloc[:, 1],  # Coluna B (index 1)
                "Cliente": df.iloc[:, 19],  # Coluna T (index 19)
                "Receita_BGX": df.iloc[:, 47]  # Coluna AV (index 47)
            })
        except:
            st.warning("As colunas não foram encontradas pelos índices esperados. Usando leitura alternativa.")
            # Ler por índice como fallback
            dados = pd.DataFrame({
                "Data": df.iloc[:, 1],
                "Cliente": df.iloc[:, 19],
                "Receita_BGX": df.iloc[:, 47]
            })
        
        # Remover linhas com data vazia
        dados = dados.dropna(subset=["Data"])
        
        # Converter a coluna Data para datetime
        dados["Data"] = pd.to_datetime(dados["Data"], errors='coerce')
        dados = dados.dropna(subset=["Data"])  # Remover linhas com datas inválidas
        
        # Filtrar por data
        if data_inicial and data_final:
            data_inicial = pd.to_datetime(data_inicial)
            data_final = pd.to_datetime(data_final)
            dados = dados[(dados["Data"] >= data_inicial) & (dados["Data"] <= data_final)]
        
        return dados
    
    except Exception as e:
        st.error(f"Erro ao ler o arquivo de câmbio: {str(e)}")
        return None

# Função para preparar um dataframe para o arquivo de destino
def preparar_df_destino(dados_cambio, df_destino=None):
    """
    Prepara um dataframe para exportação sem depender de macros
    """
    if df_destino is None:
        # Criar um novo dataframe se não for fornecido
        df_destino = pd.DataFrame(columns=["Data", "Cliente", "Receita_BGX"])
    
    # Preparar os dados para inserção
    novos_dados = pd.DataFrame({
        "Data": dados_cambio["Data"],
        "Cliente": dados_cambio["Cliente"],
        "Receita_BGX": dados_cambio["Receita_BGX"]
    })
    
    # Concatenar com os dados existentes
    df_combinado = pd.concat([df_destino, novos_dados], ignore_index=True)
    
    return df_combinado

# Interface Streamlit
st.header("Selecione os Arquivos")

# Upload de arquivos
arquivo_cambio_upload = st.file_uploader("Selecione o arquivo de Operações de câmbio", type=["xlsm", "xlsx"])

# Opção para também importar o arquivo de destino para referência
importar_destino = st.checkbox("Importar arquivo de destino para referência", value=False)
arquivo_nf_upload = None
if importar_destino:
    arquivo_nf_upload = st.file_uploader("Selecione o arquivo de Notas Fiscais", type=["xlsm", "xlsx"])

# Seleção de datas
col1, col2 = st.columns(2)
with col1:
    data_inicial = st.date_input("Data Inicial", value=None)
with col2:
    data_final = st.date_input("Data Final", value=None)

if arquivo_cambio_upload:
    if st.button("Processar Dados"):
        with st.spinner("Processando..."):
            # Ler os dados do arquivo de câmbio
            dados_cambio = ler_dados_cambio(arquivo_cambio_upload, data_inicial, data_final)
            
            if dados_cambio is not None and not dados_cambio.empty:
                st.success(f"Dados lidos com sucesso! {len(dados_cambio)} registros encontrados.")
                
                # Mostrar os dados que serão transferidos
                st.subheader("Dados a serem transferidos:")
                st.dataframe(dados_cambio)
                
                # Criar um novo arquivo Excel
                df_final = preparar_df_destino(dados_cambio)
                
                # Criar um buffer para o download
                buffer = io.BytesIO()
                
                # Criar um ExcelWriter
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    # Escrever os dados na planilha
                    df_final.to_excel(writer, sheet_name='Dados_Cambio', index=False)
                
                buffer.seek(0)
                
                # Oferecer o download do arquivo Excel (formato normal .xlsx)
                st.download_button(
                    label="Baixar Dados Extraídos (XLSX)",
                    data=buffer,
                    file_name="Dados_Cambio_Extraidos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Também oferecer download como CSV para máxima compatibilidade
                csv_buffer = io.StringIO()
                df_final.to_csv(csv_buffer, index=False)
                
                st.download_button(
                    label="Baixar Dados como CSV",
                    data=csv_buffer.getvalue(),
                    file_name="Dados_Cambio_Extraidos.csv",
                    mime="text/csv"
                )
                
                st.success("Processo concluído com sucesso!")
                
                st.info("""
                **Nota importante**: Esta versão alternativa extrai os dados e disponibiliza-os em um novo arquivo Excel 
                (.xlsx) ou CSV, sem macros. Estes dados podem ser copiados manualmente para o arquivo de destino.
                
                Se precisar manter o formato original com macros, use a outra versão do aplicativo.
                """)
            else:
                st.warning("Nenhum dado encontrado para o período selecionado.")

# Instruções de uso
st.sidebar.header("Instruções de Uso")
st.sidebar.markdown("""
1. Selecione o arquivo de operações de câmbio: "Operações de câmbio BRA.xlsm"
2. Escolha o intervalo de datas para filtrar os dados
3. Clique em "Processar Dados"
4. Verifique os dados que serão extraídos
5. Baixe o arquivo com os dados extraídos (XLSX ou CSV)
6. Importe esses dados para o arquivo de destino

Esta versão alternativa gera um novo arquivo sem macros, o que evita problemas de compatibilidade.
""")

st.sidebar.markdown("---")
st.sidebar.info("Este aplicativo extrai dados de câmbio para um novo arquivo Excel ou CSV.")
