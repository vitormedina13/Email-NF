import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, timedelta
from io import BytesIO
import os
from dateutil.relativedelta import relativedelta

# Configuração da página
st.set_page_config(
    page_title="Extrator de Dados de Câmbio",
    page_icon="💱",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilo CSS personalizado
st.markdown("""
    <style>
    .main {
        padding: 1rem 2rem;
    }
    .stButton button {
        width: 100%;
        height: 3rem;
        font-size: 1.1rem;
        font-weight: bold;
        background-color: #4CAF50;
        color: white;
    }
    .success-message {
        background-color: #e6f4ea;
        color: #137333;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #137333;
        margin: 1rem 0;
    }
    .error-message {
        background-color: #fce8e6;
        color: #c5221f;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #c5221f;
        margin: 1rem 0;
    }
    .upload-area {
        border: 2px dashed #cccccc;
        border-radius: 5px;
        padding: 20px;
        text-align: center;
        margin: 10px 0;
    }
    </style>
    """, unsafe_allow_html=True)

def main():
    st.title("🔄 Extrator de Dados de Câmbio")
    st.markdown("### Transferência de dados entre planilhas de câmbio")
    
    with st.expander("ℹ️ Como usar", expanded=False):
        st.markdown("""
        **Este aplicativo extrai dados da planilha de operações de câmbio e transfere para a planilha de Notas Fiscais.**
        
        1. Faça upload dos arquivos de origem e destino
        2. Selecione o período desejado para filtrar os dados
        3. Clique em "Extrair e Transferir Dados"
        4. Verifique o resultado e baixe o arquivo atualizado
        
        **Campos extraídos:**
        - Data (coluna B)
        - Cliente (coluna T)
        - Receita BGX (coluna AV)
        
        **Destino dos dados:**
        - Data → coluna A
        - Receita BGX → coluna E
        - Cliente → coluna I
        """)
    
    # Upload de arquivos
    st.subheader("📁 Upload de Arquivos")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Arquivo de Origem:**")
        cambio_file = st.file_uploader(
            "Faça upload do arquivo de operações de câmbio (.xlsm)",
            type=["xlsm", "xlsx"],
            help="Arquivo que contém os dados a serem extraídos"
        )
    
    with col2:
        st.markdown("**Arquivo de Destino:**")
        op_file = st.file_uploader(
            "Faça upload do arquivo de operações (.xlsm)",
            type=["xlsm", "xlsx"],
            help="Arquivo onde os dados serão inseridos"
        )
    
    # Verificação de upload dos arquivos
    files_uploaded = cambio_file is not None and op_file is not None
    
    if not files_uploaded:
        st.warning("⚠️ Por favor, faça upload de ambos os arquivos para continuar.")
    
    # Seção de seleção de período
    st.subheader("📅 Período")
    
    # Obter o primeiro e último dia do mês atual
    today = datetime.now()
    first_day_of_month = today.replace(day=1)
    next_month = first_day_of_month + relativedelta(months=1)
    last_day_of_month = next_month - timedelta(days=1)
    
    col1, col2, col3 = st.columns([2, 2, 3])
    
    with col1:
        data_inicial = st.date_input(
            "Data inicial",
            value=first_day_of_month,
            help="Selecione a data inicial do período"
        )
    
    with col2:
        data_final = st.date_input(
            "Data final",
            value=last_day_of_month,
            help="Selecione a data final do período"
        )
    
    with col3:
        st.markdown("**Períodos predefinidos:**")
        cols = st.columns(4)
        
        # Função para atualizar as datas
        if cols[0].button("📆 Este mês"):
            data_inicial = first_day_of_month
            data_final = last_day_of_month
            st.experimental_rerun()
            
        if cols[1].button("◀️ Mês anterior"):
            data_inicial = (first_day_of_month - relativedelta(months=1))
            data_final = first_day_of_month - timedelta(days=1)
            st.experimental_rerun()
            
        if cols[2].button("📆 Ano atual"):
            data_inicial = today.replace(month=1, day=1)
            data_final = today.replace(month=12, day=31)
            st.experimental_rerun()
            
        if cols[3].button("🔄 Últimos 30 dias"):
            data_inicial = today - timedelta(days=30)
            data_final = today
            st.experimental_rerun()
    
    # Botão de processamento
    st.markdown("### 🚀 Executar")
    process_button = st.button("Extrair e Transferir Dados", disabled=not files_uploaded)
    
    # Log de processamento
    log_container = st.container()
    
    if process_button and files_uploaded:
        with st.spinner("Processando dados..."):
            log_messages = []
            def add_log(message):
                log_messages.append(message)
                with log_container:
                    st.text_area("Log de Processamento", value="\n".join(log_messages), height=300, key="log_area")
            
            try:
                # Início do processamento
                add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando processamento...")
                add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Período: {data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}")
                
                # Carregar o arquivo de origem
                add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Carregando arquivo de câmbio...")
                try:
                    cambio_df = pd.read_excel(
                        cambio_file,
                        sheet_name="BGP e BGX Cambio",
                        engine="openpyxl"
                    )
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Arquivo carregado com sucesso. Encontradas {len(cambio_df)} linhas.")
                except Exception as e:
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] ERRO ao carregar o arquivo: {str(e)}")
                    st.error(f"Erro ao carregar o arquivo de câmbio: {str(e)}")
                    return
                
                # Extrair colunas específicas
                add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Extraindo colunas específicas...")
                try:
                    # Usar índices de coluna (B=1, T=19, AV=47 em base 0)
                    extracted_df = cambio_df.iloc[:, [1, 19, 47]]
                    extracted_df.columns = ["Data", "Cliente", "Receita BGX"]
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Colunas extraídas com sucesso.")
                except Exception as e:
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] ERRO ao extrair colunas: {str(e)}")
                    st.error(f"Erro ao extrair colunas: {str(e)}")
                    return
                
                # Filtrar por data
                add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Aplicando filtro de datas...")
                try:
                    extracted_df["Data"] = pd.to_datetime(extracted_df["Data"], errors='coerce')
                    mask = (extracted_df["Data"] >= pd.Timestamp(data_inicial)) & (extracted_df["Data"] <= pd.Timestamp(data_final))
                    filtered_df = extracted_df.loc[mask].copy()
                    
                    if filtered_df.empty:
                        add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Não foram encontrados dados para o período selecionado.")
                        st.warning("⚠️ Não foram encontrados dados para o período selecionado.")
                        return
                    
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Filtro aplicado. {len(filtered_df)} registros encontrados.")
                except Exception as e:
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] ERRO ao aplicar filtro de datas: {str(e)}")
                    st.error(f"Erro ao aplicar filtro de datas: {str(e)}")
                    return
                
                # Exibir amostra dos dados extraídos
                st.subheader("📊 Dados Extraídos")
                st.dataframe(filtered_df, use_container_width=True)
                
                # Abrir o arquivo de destino com openpyxl
                add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Abrindo arquivo de destino...")
                try:
                    wb = openpyxl.load_workbook(op_file, keep_vba=True)
                    
                    # Selecionar a aba desejada
                    try:
                        ws = wb["Todas as Op - Câmbio"]
                        add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Aba 'Todas as Op - Câmbio' encontrada.")
                    except KeyError:
                        add_log(f"[{datetime.now().strftime('%H:%M:%S')}] ERRO: Aba 'Todas as Op - Câmbio' não encontrada!")
                        st.error("Erro: Aba 'Todas as Op - Câmbio' não encontrada no arquivo de destino.")
                        return
                    
                    # Encontrar a última linha com dados na tabela
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Procurando última linha disponível na tabela...")
                    last_row = 0
                    for row in ws.iter_rows(min_row=1, max_col=1):
                        if row[0].value is not None:
                            last_row = row[0].row
                        else:
                            if last_row > 0:  # Já encontramos pelo menos uma linha com dados
                                break
                    
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Última linha ocupada: {last_row}")
                    
                    # Adicionar dados na tabela de destino
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Inserindo {len(filtered_df)} registros na tabela...")
                    for i, (_, row) in enumerate(filtered_df.iterrows()):
                        last_row += 1
                        # Adicionar data (coluna A)
                        ws.cell(row=last_row, column=1, value=row["Data"].date())
                        # Adicionar receita BGX (coluna E)
                        ws.cell(row=last_row, column=5, value=row["Receita BGX"])
                        # Adicionar cliente (coluna I)
                        ws.cell(row=last_row, column=9, value=row["Cliente"])
                    
                    # Salvar em buffer para download
                    buffer = BytesIO()
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] Salvando arquivo...")
                    wb.save(buffer)
                    buffer.seek(0)
                    
                    # Exibir sucesso
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] ✅ PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
                    
                    # Mensagem de sucesso estilizada
                    st.markdown(f"""
                    <div class="success-message">
                        <h3>✅ Operação concluída com sucesso!</h3>
                        <p>Total de {len(filtered_df)} registros transferidos.</p>
                        <p>Clique no botão abaixo para baixar o arquivo atualizado.</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Botão de download
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    file_name = os.path.basename(op_file.name)
                    download_name = f"Operacoes_Atualizadas_{timestamp}.xlsm"
                    
                    st.download_button(
                        label="📥 Baixar Arquivo Atualizado",
                        data=buffer,
                        file_name=download_name,
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                        key="download_button"
                    )
                    
                except Exception as e:
                    add_log(f"[{datetime.now().strftime('%H:%M:%S')}] ERRO ao processar arquivo de destino: {str(e)}")
                    st.error(f"Erro ao processar arquivo de destino: {str(e)}")
                    return
                    
            except Exception as e:
                add_log(f"[{datetime.now().strftime('%H:%M:%S')}] ERRO geral: {str(e)}")
                st.markdown(f"""
                <div class="error-message">
                    <h3>❌ Erro durante o processamento</h3>
                    <p>{str(e)}</p>
                </div>
                """, unsafe_allow_html=True)
                return

# Executar a aplicação
if __name__ == "__main__":
    main()
