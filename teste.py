import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import io
import os
from pathlib import Path

st.set_page_config(page_title="Transferência de Dados de Câmbio", layout="wide")

st.title("Transferência de Dados de Câmbio")

# Função para ler os dados do arquivo de origem
def ler_dados_cambio(arquivo_cambio, data_inicial, data_final):
    try:
        # Carregar o arquivo Excel com openpyxl para preservar formatação
        wb = openpyxl.load_workbook(arquivo_cambio, data_only=True)
        
        # Selecionar a planilha BGP e BGX Cambio
        ws = wb["BGP e BGX Cambio"]
        
        # Converter os dados da planilha para um DataFrame
        data = []
        for row in ws.iter_rows(min_row=2, values_only=True):  # Começar da segunda linha para pular o cabeçalho
            if row[1] is not None:  # Coluna B (Data)
                data.append({
                    "Data": row[1],  # Coluna B
                    "Cliente": row[19],  # Coluna T
                    "Receita_BGX": row[47]  # Coluna AV
                })
        
        df = pd.DataFrame(data)
        
        # Converter a coluna Data para datetime se ainda não estiver nesse formato
        if not pd.api.types.is_datetime64_any_dtype(df["Data"]):
            df["Data"] = pd.to_datetime(df["Data"], errors='coerce')
        
        # Filtrar por data
        if data_inicial and data_final:
            data_inicial = pd.to_datetime(data_inicial)
            data_final = pd.to_datetime(data_final)
            df = df[(df["Data"] >= data_inicial) & (df["Data"] <= data_final)]
        
        return df
    
    except Exception as e:
        st.error(f"Erro ao ler o arquivo de câmbio: {str(e)}")
        return None

# Função para atualizar o arquivo de destino
def atualizar_notas_fiscais(arquivo_nf, dados_cambio):
    try:
        # Carregar o arquivo de destino
        wb = openpyxl.load_workbook(arquivo_nf)
        
        # Selecionar a aba correta
        ws = wb["Todas as Op - Câmbio"]
        
        # Encontrar a última linha com dados na tabela
        ultima_linha = 1  # Começar da linha 1
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
            if row[0].value is not None:
                ultima_linha = row[0].row
            else:
                break
        
        # Adicionar os novos dados a partir da última linha + 1
        linha_atual = ultima_linha + 1
        
        # Inserir os dados nas colunas corretas
        for _, row in dados_cambio.iterrows():
            ws.cell(row=linha_atual, column=1).value = row["Data"]  # Coluna A - Data
            ws.cell(row=linha_atual, column=9).value = row["Cliente"]  # Coluna I - Cliente
            ws.cell(row=linha_atual, column=5).value = row["Receita_BGX"]  # Coluna E - Receita BGX
            linha_atual += 1
        
        # Criar um buffer de bytes para salvar o arquivo atualizado
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return output
    
    except Exception as e:
        st.error(f"Erro ao atualizar o arquivo de notas fiscais: {str(e)}")
        return None

# Interface Streamlit
st.header("Selecione os Arquivos")

# Upload de arquivos
arquivo_cambio_upload = st.file_uploader("Selecione o arquivo de Operações de câmbio", type=["xlsm", "xlsx"])
arquivo_nf_upload = st.file_uploader("Selecione o arquivo de Notas Fiscais", type=["xlsm", "xlsx"])

# Seleção de datas
col1, col2 = st.columns(2)
with col1:
    data_inicial = st.date_input("Data Inicial", value=None)
with col2:
    data_final = st.date_input("Data Final", value=None)

if arquivo_cambio_upload and arquivo_nf_upload:
    if st.button("Processar Dados"):
        with st.spinner("Processando..."):
            # Salvar os arquivos carregados temporariamente
            temp_cambio = "temp_cambio.xlsm"
            temp_nf = "temp_nf.xlsm"
            
            with open(temp_cambio, 'wb') as f:
                f.write(arquivo_cambio_upload.getvalue())
            
            with open(temp_nf, 'wb') as f:
                f.write(arquivo_nf_upload.getvalue())
            
            # Ler os dados do arquivo de câmbio
            dados_cambio = ler_dados_cambio(temp_cambio, data_inicial, data_final)
            
            if dados_cambio is not None and not dados_cambio.empty:
                st.success(f"Dados lidos com sucesso! {len(dados_cambio)} registros encontrados.")
                
                # Mostrar os dados que serão transferidos
                st.subheader("Dados a serem transferidos:")
                st.dataframe(dados_cambio)
                
                # Atualizar o arquivo de notas fiscais
                arquivo_atualizado = atualizar_notas_fiscais(temp_nf, dados_cambio)
                
                if arquivo_atualizado:
                    # Oferecer o download do arquivo atualizado
                    nome_arquivo = "Operacoes_Atualizadas.xlsm"
                    st.download_button(
                        label="Baixar Arquivo Atualizado",
                        data=arquivo_atualizado,
                        file_name=nome_arquivo,
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12"
                    )
                    
                    st.success("Processo concluído com sucesso!")
            else:
                st.warning("Nenhum dado encontrado para o período selecionado.")
            
            # Limpar arquivos temporários
            try:
                os.remove(temp_cambio)
                os.remove(temp_nf)
            except:
                pass

# Instruções de uso
st.sidebar.header("Instruções de Uso")
st.sidebar.markdown("""
1. Selecione o arquivo de operações de câmbio: "Operações de câmbio BRA.xlsm"
2. Selecione o arquivo de notas fiscais: "01. Operações.xlsm"
3. Escolha o intervalo de datas para filtrar os dados
4. Clique em "Processar Dados"
5. Verifique os dados que serão transferidos
6. Baixe o arquivo atualizado
""")

st.sidebar.markdown("---")
st.sidebar.info("Este aplicativo transfere dados de câmbio entre arquivos Excel.")
