import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Extrator RFB (Robusto)", layout="wide", page_icon="🛡️")

# --- 1. FUNÇÃO DE EXTRAÇÃO (Motor Reforçado) ---
def extrair_tabelas_brutas(uploaded_files):
    df_consolidado = pd.DataFrame()
    
    progresso_texto = st.empty()
    barra = st.progress(0)
    total_arquivos = len(uploaded_files)
    
    for i, pdf_file in enumerate(uploaded_files):
        progresso_texto.text(f"Lendo {i+1}/{total_arquivos}: {pdf_file.name}")
        
        arquivo_teve_dados = False 
        
        try:
            with pdfplumber.open(pdf_file) as pdf:
                # --- TENTATIVA 1: Caçar Tabelas (Modo Agressivo) ---
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table: continue
                        
                        df_temp = pd.DataFrame(table)
                        if df_temp.empty: continue
                        
                        # TRATAMENTO DE CABEÇALHO (Aqui estava o problema de Sonora)
                        # Pega a primeira linha, converte para string, remove quebras de linha e espaços extras
                        first_row_values = [str(x).upper().replace('\n', ' ').strip() for x in df_temp.iloc[0].values]
                        header_string = " ".join(first_row_values)
                        
                        # Lista expandida de palavras-chave para aceitar tabelas diferentes
                        keywords = [
                            "PROCESSO", "MODALIDADE", "SALDO", "DEVEDOR", 
                            "CNPJ", "VINCULADO", "VALOR", "TOTAL", 
                            "SITUAÇÃO", "NATUREZA", "TRIBUTO"
                        ]
                        
                        # Se encontrar QUALQUER palavra-chave no cabeçalho
                        if any(k in header_string for k in keywords):
                            
                            # Se a tabela tiver só o cabeçalho, ignora
                            if len(df_temp) <= 1: continue 

                            # Ajusta o cabeçalho do DataFrame
                            df_temp.columns = first_row_values 
                            df_temp = df_temp[1:] # Dados começam na linha 2
                            df_temp["Arquivo Origem"] = pdf_file.name
                            
                            df_consolidado = pd.concat([df_consolidado, df_temp], ignore_index=True)
                            arquivo_teve_dados = True

                # --- TENTATIVA 2: Resgate (Arquivos sem tabela ou Nada Consta) ---
                if not arquivo_teve_dados:
                    # Tenta ler o texto bruto da primeira página
                    texto_pag1 = ""
                    if len(pdf.pages) > 0:
                        texto_pag1 = pdf.pages[0].extract_text() or ""
                    
                    # Se tiver texto, tenta extrair CNPJ
                    if texto_pag1.strip():
                        cnpj_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_pag1)
                        cnpj_num = cnpj_match.group(1) if cnpj_match else ""

                        # Cria o registro "Nada Consta"
                        df_resgate = pd.DataFrame([{
                            "Arquivo Origem": pdf_file.name,
                            "CNPJ VINCULADO": cnpj_num,
                            "PROCESSO/ID PARCELAMENTO": "-",   
                            "MODALIDADE": "Nada Consta",       
                            "SISTEMA": "-",
                            "SALDO DEVEDOR": "-"               
                        }])
                        df_consolidado = pd.concat([df_consolidado, df_resgate], ignore_index=True)
                    else:
                        # Se não tiver texto nenhum (imagem pura)
                        st.warning(f"⚠️ O arquivo '{pdf_file.name}' parece ser uma imagem digitalizada (não tem texto selecionável). O extrator não consegue ler imagens.")

        except Exception as e:
            st.error(f"Erro ao processar {pdf_file.name}: {e}")
        
        barra.progress((i + 1) / total_arquivos)
    
    progresso_texto.empty()
    barra.empty()
    
    return df_consolidado

# --- 2. FUNÇÃO DE LIMPEZA E ORGANIZAÇÃO ---
def organizar_dados(df_bruto):
    if df_bruto.empty:
        return pd.DataFrame()

    df = df_bruto.copy()

    # A. Normalização de Colunas (Upper + Trim)
    df.columns = [str(c).replace('\n', ' ').strip().upper() if c is not None else f"COL_{idx}" for idx, c in enumerate(df.columns)]
    
    # B. Mapeamento Flexível
    col_map = {}
    for col in df.columns:
        if "PROCESSO" in col: col_map[col] = "Processo"
        elif "CNPJ" in col: col_map[col] = "CNPJ"
        elif "MODALIDADE" in col or "NATUREZA" in col: col_map[col] = "Modalidade"
        elif "SISTEMA" in col: col_map[col] = "Sistema"
        elif "SALDO" in col or "VALOR" in col: col_map[col] = "Valor Original"
        elif "ARQUIVO" in col: col_map[col] = "Arquivo"

    df = df.rename(columns=col_map)

    # C. Deduplicação de Colunas
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique(): 
        cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols

    # D. Filtros
    if 'CNPJ' in df.columns:
        # Remove linhas onde o conteúdo é igual ao título
        df = df[~df['CNPJ'].astype(str).str.contains("CNPJ", case=False, na=False)]
    
    # Remove linhas de Totais
    mask_total = df.astype(str).apply(lambda x: x.str.contains('TOTAL', case=False)).any(axis=1)
    df = df[~mask_total]

    # E. Limpeza de Texto (\n -> espaço)
    for col in df.columns:
        if col in df.columns and df[col].dtype == object:
            df[col] = df[col].astype(str).str.replace('\n', ' ').str.strip()

    # F. Conversão Numérica
    if 'Valor Original' in df.columns:
        def converter_valor(x):
            try:
                x_str = str(x).strip()
                if x_str == "-" or x_str == "": return 0.0
                
                # Remove R$, espaços, pontos milhar
                x_str = x_str.upper().replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
                if re.match(r'^-?\d+(\.\d+)?$', x_str):
                    return float(x_str)
                return 0.0
            except:
                return 0.0

        df['Valor Numérico'] = df['Valor Original'].apply(converter_valor)

    # G. Extração de Município (Title Case)
    def extrair_municipio(row):
        nome_arq = str(row.get('Arquivo', ''))
        try:
            partes = nome_arq.split('-')
            if len(partes) >= 2:
                return partes[1].strip().title() 
            return nome_arq.title()
        except:
            return "Desconhecido"

    df.insert(0, 'Município', df.apply(extrair_municipio, axis=1))

    # H. Seleção Final
    cols_desejadas = ['Município', 'CNPJ', 'Processo', 'Modalidade', 'Sistema', 'Valor Numérico', 'Valor Original', 'Arquivo']
    cols_finais = [c for c in cols_desejadas if c in df.columns]
    
    return df[cols_finais]

# --- 3. INTERFACE PRINCIPAL ---
st.title("🗂️ Extrator RFB (Robusto)")
st.markdown("Extração aprimorada para tabelas quebradas ou formatos atípicos (ex: Sonora).")

uploaded_files = st.file_uploader("Arraste os PDFs aqui", type="pdf", accept_multiple_files=True)

if uploaded_files:
    with st.spinner("Analisando estrutura dos arquivos..."):
        df_bruto = extrair_tabelas_brutas(uploaded_files)
    
    if not df_bruto.empty:
        try:
            df_limpo = organizar_dados(df_bruto)
            
            st.success(f"✅ Processamento concluído! {len(df_limpo)} registros encontrados.")
            
            tab1, tab2 = st.tabs(["📊 Lista Gerencial", "🔍 Dados Brutos"])
            
            with tab1:
                st.dataframe(df_limpo, use_container_width=True)
                if 'Valor Numérico' in df_limpo.columns:
                    total = df_limpo['Valor Numérico'].sum()
                    st.metric("Total da Seleção", f"R$ {total:,.2f}")
            
            with tab2:
                st.write("Dados extraídos diretamente do PDF (sem tratamento):")
                st.dataframe(df_bruto, use_container_width=True)
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Aba Refinada
                df_limpo.to_excel(writer, index=False, sheet_name='Refinado')
                wb = writer.book
                ws = writer.sheets['Refinado']
                
                # Formatação Contábil
                fmt_contabil = wb.add_format({'num_format': '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'})
                
                if "Valor Numérico" in df_limpo.columns:
                    idx_val = df_limpo.columns.get_loc("Valor Numérico")
                    ws.set_column(idx_val, idx_val, 18, fmt_contabil)
                
                ws.set_column(0, 0, 25) # Município
                ws.set_column(2, 3, 30) # Processo/Modalidade

                # Aba Bruta
                df_bruto.to_excel(writer, index=False, sheet_name='Bruto')
                
            output.seek(0)
            st.download_button(
                "⬇️ Baixar Excel (.xlsx)",
                output,
                "Relatorio_RFB_Final.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Erro na organização dos dados: {e}")
            st.dataframe(df_bruto)
    else:
        st.warning("Nenhum dado encontrado. Verifique se os arquivos não são imagens digitalizadas.")
