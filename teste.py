import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from io import BytesIO
import os

st.set_page_config(page_title="Extração de Dados de Câmbio", layout="wide")

def main():
    st.title("Extração de Dados de Câmbio")
    
    # Definir os caminhos dos arquivos
    cambio_path = st.text_input(
        "Caminho do arquivo 'Operações de câmbio BRA.xlsm'",
        value=r"C:\Users\VitorMedina\Bluegreen\Finance - Documents\OTC\Operações de câmbio BRA.xlsm"
    )
    
    op_path = st.text_input(
        "Caminho do arquivo 'Operações.xlsm'",
        value=r"C:\Users\VitorMedina\Bluegreen\Finance - Documents\Notas Fiscais\4. Notas Fiscais\01. Operações.xlsm"
    )
    
    # Seleção de datas
    st.subheader("Filtrar por Data")
    col1, col2 = st.columns(2)
    with col1:
        data_inicial = st.date_input("Data inicial", datetime.now())
    with col2:
        data_final = st.date_input("Data final", datetime.now())
    
    if st.button("Extrair e Transferir Dados"):
        with st.spinner("Processando..."):
            try:
                # Carregar o arquivo de origem
                st.write("Carregando arquivo de câmbio...")
                cambio_df = pd.read_excel(
                    cambio_path,
                    sheet_name="BGP e BGX Cambio",
                    engine="openpyxl"
                )
                
                # Verificar se a tabela existe
                if "Tabela_Câmbio" in cambio_df:
                    # Se a tabela for nomeada, extrair apenas os dados dessa tabela
                    cambio_df = cambio_df[cambio_df["Tabela_Câmbio"]]
                
                # Filtrar as colunas desejadas
                st.write("Extraindo colunas específicas...")
                # Assumindo que as colunas têm os nomes Data, Cliente e Receita BGX
                # Se não tiverem, precisamos usar os índices das colunas
                try:
                    extracted_df = cambio_df[["Data", "Cliente", "Receita BGX"]]
                except KeyError:
                    # Usar índices de coluna se os nomes não estiverem disponíveis
                    extracted_df = cambio_df.iloc[:, [1, 19, 47]]  # Colunas B, T, AV
                    extracted_df.columns = ["Data", "Cliente", "Receita BGX"]
                
                # Filtrar por data
                st.write("Aplicando filtro de datas...")
                extracted_df["Data"] = pd.to_datetime(extracted_df["Data"])
                mask = (extracted_df["Data"] >= pd.Timestamp(data_inicial)) & (extracted_df["Data"] <= pd.Timestamp(data_final))
                filtered_df = extracted_df.loc[mask]
                
                if filtered_df.empty:
                    st.error("Não foram encontrados dados para o período selecionado.")
                    return
                
                # Mostrar dados extraídos
                st.write(f"Dados extraídos ({len(filtered_df)} registros):")
                st.dataframe(filtered_df)
                
                # Abrir o arquivo de destino com openpyxl para preservar fórmulas e formatação
                st.write("Abrindo arquivo de destino...")
                wb = openpyxl.load_workbook(op_path, keep_vba=True)
                
                # Selecionar a aba desejada
                try:
                    ws = wb["Todas as Op - Câmbio"]
                except KeyError:
                    st.error("Aba 'Todas as Op - Câmbio' não encontrada no arquivo de destino.")
                    return
                
                # Encontrar a última linha com dados na tabela
                st.write("Procurando última linha disponível na tabela de destino...")
                last_row = 1
                for row in ws.iter_rows(min_row=1, max_col=1):
                    if row[0].value is not None:
                        last_row = row[0].row
                    else:
                        break
                
                # Adicionar dados na tabela de destino
                st.write("Inserindo dados na tabela de destino...")
                for i, row in filtered_df.iterrows():
                    last_row += 1
                    # Adicionar data (coluna A)
                    ws.cell(row=last_row, column=1, value=row["Data"])
                    # Adicionar receita BGX (coluna E)
                    ws.cell(row=last_row, column=5, value=row["Receita BGX"])
                    # Adicionar cliente (coluna I)
                    ws.cell(row=last_row, column=9, value=row["Cliente"])
                
                # Salvar em um buffer de memória para download
                st.write("Preparando arquivo para download...")
                buffer = BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                
                # Opção para salvar no destino original
                if st.checkbox("Também salvar no arquivo original", value=False):
                    try:
                        wb.save(op_path)
                        st.success(f"Arquivo salvo com sucesso em {op_path}")
                    except Exception as e:
                        st.error(f"Erro ao salvar no arquivo original: {str(e)}")
                
                # Preparar para download
                st.download_button(
                    label="Baixar arquivo atualizado",
                    data=buffer,
                    file_name="Operacoes_Atualizadas.xlsm",
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12"
                )
                
                st.success(f"Processamento concluído! {len(filtered_df)} registros transferidos.")
                
            except Exception as e:
                st.error(f"Ocorreu um erro: {str(e)}")
                st.error("Detalhes do erro para debug:")
                st.exception(e)

if __name__ == "__main__":
    main()
