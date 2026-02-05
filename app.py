import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Extrator RFB (Com Zerados)", layout="wide", page_icon="🗂️")

# --- 1. FUNÇÃO DE EXTRAÇÃO (Com Resgate de Arquivos Zerados) ---
def extrair_tabelas_brutas(uploaded_files):
    df_consolidado = pd.DataFrame()
    
    progresso_texto = st.empty()
    barra = st.progress(0)
    total_arquivos = len(uploaded_files)
    
    for i, pdf_file in enumerate(uploaded_files):
        progresso_texto.text(f"Processando {i+1}/{total_arquivos}: {pdf_file.name}")
        
        arquivo_teve_dados = False # Flag de controle
        
        try:
            with pdfplumber.open(pdf_file) as pdf:
                # --- TENTATIVA 1: Extração por Tabela (Para quem tem dívida) ---
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table: continue
                        
                        df_temp = pd.DataFrame(table)
                        if df_temp.empty: continue
                        
                        # Converte header para string maiúscula para busca
                        header_row = str(df_temp.iloc[0].values).upper()
                        
                        # Palavras-chave de tabela válida
                        keywords = ["PROCESSO", "MODALIDADE", "SALDO DEVEDOR", "CNPJ VINCULADO"]
                        if any(k in header_row for k in keywords):
                            
                            # Se a tabela tem apenas o cabeçalho e mais nada (caso raro de tabela vazia desenhada)
                            if len(df_temp) <= 1:
                                continue 

                            df_temp.columns = df_temp.iloc[0] # Define header
                            df_temp = df_temp[1:] # Remove linha do header dos dados
                            df_temp["Arquivo Origem"] = pdf_file.name
                            df_consolidado = pd.concat([df_consolidado, df_temp], ignore_index=True)
                            arquivo_teve_dados = True

                # --- TENTATIVA 2: Resgate (Para arquivos zerados/sem tabela) ---
                if not arquivo_teve_dados:
                    # Lê o texto da primeira página para pegar metadados
                    texto_pag1 = pdf.pages[0].extract_text() or ""
                    
                    # Regex para capturar Município e CNPJ
                    municipio_match = re.search(r"MUNICIPIO DE\s+(.*)", texto_pag1)
                    municipio_nome = municipio_match.group(1).strip() if municipio_match else "DESCONHECIDO"
                    
                    cnpj_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_pag1)
                    cnpj_num = cnpj_match.group(1) if cnpj_match else ""

                    # Cria uma linha "fake" para constar no relatório
                    df_resgate = pd.DataFrame([{
                        "Arquivo Origem": pdf_file.name,
                        "CNPJ VINCULADO": cnpj_num, # Usa nome da coluna padrão da RFB
                        "PROCESSO/ID PARCELAMENTO": "SEM PARCELAMENTO IDENTIFICADO",
                        "MODALIDADE": "Nada Consta",
                        "SISTEMA": "-",
                        "SALDO DEVEDOR": "0,00"
                    }])
                    
                    df_consolidado = pd.concat([df_consolidado, df_resgate], ignore_index=True)

        except Exception as e:
            st.error(f"Erro no arquivo {pdf_file.name}: {e}")
        
        barra.progress((i + 1) / total_arquivos)
    
    progresso_texto.empty()
    barra.empty()
    
    return df_consolidado

# --- 2. FUNÇÃO DE LIMPEZA E ORGANIZAÇÃO ---
def organizar_dados(df_bruto):
    if df_bruto.empty:
        return pd.DataFrame()

    df = df_bruto.copy()

    # A. Normalização de Nomes de Colunas
    df.columns = [str(c).replace('\n', ' ').strip().upper() if c is not None else f"COL_{idx}" for idx, c in enumerate(df.columns)]
    
    # B. Mapeamento de Colunas
    col_map = {}
    for col in df.columns:
        if "PROCESSO" in col: col_map[col] = "Processo"
        elif "CNPJ" in col: col_map[col] = "CNPJ"
        elif "MODALIDADE" in col: col_map[col] = "Modalidade"
        elif "SISTEMA" in col: col_map[col] = "Sistema"
        elif "SALDO" in col and "DEVEDOR" in col: col_map[col] = "Valor Original"
        elif "ARQUIVO" in col: col_map[col] = "Arquivo"

    df = df.rename(columns=col_map)

    # C. Deduplicação de Colunas (Segurança)
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique(): 
        cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols

    # D. Filtros de Linhas Inválidas
    # Remove repetições de cabeçalho
    if 'CNPJ' in df.columns:
        df = df[~df['CNPJ'].astype(str).str.contains("CNPJ", case=False, na=False)]
    
    # Remove Totais, mas MANTÉM as linhas de "SEM PARCELAMENTO" (que tem valor 0,00)
    # A lógica: Se tiver "TOTAL" na linha E o valor não for 0,00 vindo do nosso resgate.
    mask_total = df.astype(str).apply(lambda x: x.str.contains('TOTAL', case=False)).any(axis=1)
    df = df[~mask_total]

    # E. Limpeza de Texto
    for col in df.columns:
        if col in df.columns and df[col].dtype == object:
            df[col] = df[col].astype(str).str.replace('\n', ' ').str.strip()

    # F. Conversão Numérica
    if 'Valor Original' in df.columns:
        def converter_valor(x):
            try:
                x_str = str(x).replace(' ', '').replace('.', '').replace(',', '.')
                if re.match(r'^-?\d+(\.\d+)?$', x_str):
                    return float(x_str)
                return 0.0
            except:
                return 0.0

        df['Valor Numérico'] = df['Valor Original'].apply(converter_valor)

    # G. Extração de Município do Nome do Arquivo
    def extrair_municipio(row):
        nome_arq = str(row.get('Arquivo', ''))
        # Se for do resgate, talvez já tenhamos pego o município via regex? 
        # Mas para padronizar, vamos tentar pegar do arquivo primeiro.
        try:
            partes = nome_arq.split('-')
            if len(partes) >= 2:
                return partes[1].strip().upper()
            return nome_arq
        except:
            return "DESCONHECIDO"

    df.insert(0, 'Município', df.apply(extrair_municipio, axis=1))

    # H. Seleção Final
    cols_desejadas = ['Município', 'CNPJ', 'Processo', 'Modalidade', 'Sistema', 'Valor Numérico', 'Valor Original', 'Arquivo']
    cols_finais = [c for c in cols_desejadas if c in df.columns]
    
    return df[cols_finais]

# --- 3. INTERFACE PRINCIPAL ---
st.title("🗂️ Extrator RFB (Lista Completa)")
st.markdown("Extrai parcelamentos. Se o município não tiver débitos, ele aparecerá na lista como 'Sem Parcelamento' (R$ 0,00).")

uploaded_files = st.file_uploader("Arraste os PDFs aqui", type="pdf", accept_multiple_files=True)

if uploaded_files:
    with st.spinner("Analisando arquivos..."):
        df_bruto = extrair_tabelas_brutas(uploaded_files)
    
    if not df_bruto.empty:
        try:
            df_limpo = organizar_dados(df_bruto)
            
            st.success(f"Processamento concluído! {len(df_limpo)} registros gerados.")
            
            # Abas
            tab1, tab2 = st.tabs(["✅ Lista Refinada (Gestão)", "🔍 Dados Brutos (Auditoria)"])
            
            with tab1:
                st.dataframe(df_limpo, use_container_width=True)
                if 'Valor Numérico' in df_limpo.columns:
                    total = df_limpo['Valor Numérico'].sum()
                    st.metric("Soma Total dos Débitos Encontrados", f"R$ {total:,.2f}")
            
            with tab2:
                st.dataframe(df_bruto, use_container_width=True)
            
            # Download
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Aba Organizada
                df_limpo.to_excel(writer, index=False, sheet_name='Refinado')
                wb = writer.book
                ws = writer.sheets['Refinado']
                fmt_money = wb.add_format({'num_format': '#,##0.00'})
                
                if "Valor Numérico" in df_limpo.columns:
                    idx_val = df_limpo.columns.get_loc("Valor Numérico")
                    ws.set_column(idx_val, idx_val, 18, fmt_money)
                
                ws.set_column(0, 0, 25) # Município
                ws.set_column(2, 3, 30) # Processo/Modalidade

                # Aba Bruta
                df_bruto.to_excel(writer, index=False, sheet_name='Bruto')
                
            output.seek(0)
            st.download_button(
                "⬇️ Baixar Excel Completo (.xlsx)",
                output,
                "Relatorio_RFB_Completo.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Erro na organização: {e}")
            st.dataframe(df_bruto)
    else:
        st.warning("Nenhum dado encontrado nos arquivos.")
