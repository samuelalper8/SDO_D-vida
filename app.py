import streamlit as st
import pdfplumber
import fitz  # PyMuPDF (Mais robusto para texto solto)
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Extrator RFB (Inteligente)", layout="wide", page_icon="🧠")

# --- 1. FUNÇÃO DE EXTRAÇÃO (Híbrida: Tabela + Texto) ---
def extrair_dados_pdf(uploaded_files):
    df_consolidado = pd.DataFrame()
    
    progresso = st.progress(0)
    status_text = st.empty()
    total = len(uploaded_files)
    
    for i, pdf_file in enumerate(uploaded_files):
        status_text.text(f"Analisando {i+1}/{total}: {pdf_file.name}")
        
        dados_encontrados = False
        
        # --- ETAPA 1: TENTATIVA VIA TABELA (PDFPLUMBER) ---
        # Ideal para documentos com grades/linhas bem definidas
        try:
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table: continue
                        df_temp = pd.DataFrame(table)
                        if df_temp.empty: continue
                        
                        # Limpa cabeçalho para busca
                        header_row = str(df_temp.iloc[0].values).upper().replace('\n', ' ')
                        keywords = ["PROCESSO", "MODALIDADE", "SALDO", "CNPJ", "VALOR", "SITUAÇÃO"]
                        
                        if any(k in header_row for k in keywords):
                            if len(df_temp) > 1:
                                df_temp.columns = df_temp.iloc[0] # Define header
                                df_temp = df_temp[1:] # Dados
                                df_temp["Arquivo Origem"] = pdf_file.name
                                df_temp["Metodo"] = "Tabela"
                                df_consolidado = pd.concat([df_consolidado, df_temp], ignore_index=True)
                                dados_encontrados = True
        except:
            pass # Se falhar, vai para a próxima etapa

        # --- ETAPA 2: TENTATIVA VIA TEXTO BRUTO (PYMUPDF) ---
        # Ideal para Sonora e tabelas sem bordas (Invisible Tables)
        if not dados_encontrados:
            try:
                # Reabre o arquivo com PyMuPDF (fitz) que é melhor para texto
                pdf_file.seek(0) # Reseta ponteiro
                doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
                texto_completo = ""
                for page in doc:
                    texto_completo += page.get_text()
                
                # Regex para encontrar o CNPJ do cabeçalho (para referência)
                cnpj_header_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_completo)
                cnpj_header = cnpj_header_match.group(1) if cnpj_header_match else ""

                linhas = texto_completo.split('\n')
                for linha in linhas:
                    # Lógica de Sonora: A linha tem CNPJ e um Valor Monetário?
                    # Regex busca: CNPJ ... espaço ... Processo ... espaço ... Valor
                    # Exemplo Sonora: "24.651.234/0001-67","641993919",... "505.961,53"
                    
                    # 1. Tem CNPJ na linha?
                    if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", linha):
                        # 2. Tem valor monetário no fim?
                        if re.search(r"\d+,\d{2}", linha):
                            
                            # Extração por Regex
                            # Remove caracteres de "aspas" que vi no seu arquivo de Sonora
                            linha_limpa = linha.replace('"', '').replace("'", "")
                            
                            # Captura Processo (número longo que não é o CNPJ)
                            # Remove o CNPJ da linha para não confundir
                            linha_sem_cnpj = re.sub(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", "", linha_limpa)
                            processos = re.findall(r"\b\d{7,}\b", linha_sem_cnpj)
                            processo_final = processos[0] if processos else "-"
                            
                            # Captura Valor (último número monetário da linha)
                            valores = re.findall(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha_limpa)
                            valor_final = valores[-1] if valores else "0,00"
                            
                            # Captura Modalidade (Texto entre Processo e Valor - aproximado)
                            # Pega palavras maiúsculas/números típicos de leis
                            modalidade = "Verificar (Extração Texto)"
                            if "Lei" in linha_limpa or "EC" in linha_limpa or "Simplificado" in linha_limpa:
                                modalidade_match = re.search(r"(Lei.*?|EC.*?|Parcelamento.*?)(\d+,)", linha_limpa)
                                if modalidade_match:
                                    modalidade = modalidade_match.group(1).replace(valor_final, "").strip()

                            df_consolidado = pd.concat([df_consolidado, pd.DataFrame([{
                                "Arquivo Origem": pdf_file.name,
                                "CNPJ VINCULADO": cnpj_header, # Usa o do cabeçalho para garantir
                                "PROCESSO/ID PARCELAMENTO": processo_final,
                                "MODALIDADE": modalidade[:50], # Corta se for muito longo
                                "SISTEMA": "Texto Extraído",
                                "SALDO DEVEDOR": valor_final,
                                "Metodo": "Texto"
                            }])], ignore_index=True)
                            dados_encontrados = True

                # --- ETAPA 3: NADA CONSTA ---
                if not dados_encontrados:
                    # Se leu o texto mas não achou linhas de débito, assume sem dívida
                    if cnpj_header:
                        df_consolidado = pd.concat([df_consolidado, pd.DataFrame([{
                            "Arquivo Origem": pdf_file.name,
                            "CNPJ VINCULADO": cnpj_header,
                            "PROCESSO/ID PARCELAMENTO": "-",
                            "MODALIDADE": "Nada Consta",
                            "SISTEMA": "-",
                            "SALDO DEVEDOR": "-",
                            "Metodo": "Resgate"
                        }])], ignore_index=True)
                    else:
                        # Caso extremo: Imagem pura sem texto nenhum
                         st.warning(f"⚠️ Arquivo '{pdf_file.name}' parece ser uma imagem sem texto. OCR seria necessário aqui.")

            except Exception as e:
                st.error(f"Erro na leitura de texto de {pdf_file.name}: {e}")

        progresso.progress((i + 1) / total)

    status_text.empty()
    progresso.empty()
    return df_consolidado

# --- 2. ORGANIZAÇÃO E LIMPEZA ---
def organizar_dados(df_bruto):
    if df_bruto.empty: return pd.DataFrame()
    df = df_bruto.copy()

    # Normalização de Colunas
    df.columns = [str(c).replace('\n', ' ').strip().upper() if c else f"C{i}" for i, c in enumerate(df.columns)]
    
    # Mapeamento
    col_map = {}
    for c in df.columns:
        if "PROCESSO" in c: col_map[c] = "Processo"
        elif "CNPJ" in c: col_map[c] = "CNPJ"
        elif "MODALIDADE" in c: col_map[c] = "Modalidade"
        elif "SISTEMA" in c: col_map[c] = "Sistema"
        elif "SALDO" in c or "VALOR" in c: col_map[c] = "Valor Original"
        elif "ARQUIVO" in c: col_map[c] = "Arquivo"
    df = df.rename(columns=col_map)

    # Deduplicação de Colunas
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique(): 
        cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols

    # Filtros
    if 'CNPJ' in df.columns:
        df = df[~df['CNPJ'].astype(str).str.contains("CNPJ", case=False, na=False)]
    mask_total = df.astype(str).apply(lambda x: x.str.contains('TOTAL', case=False)).any(axis=1)
    df = df[~mask_total]

    # Conversão Numérica
    if 'Valor Original' in df.columns:
        def conv(x):
            s = str(x).strip().replace(' ', '').replace('.', '').replace(',', '.')
            if s == "-" or s == "": return 0.0
            if re.match(r'^-?\d+(\.\d+)?$', s): return float(s)
            return 0.0
        df['Valor Numérico'] = df['Valor Original'].apply(conv)

    # Município Title Case
    def get_city(row):
        f = str(row.get('Arquivo', ''))
        try:
            parts = f.split('-')
            return parts[1].strip().title() if len(parts) >= 2 else f
        except: return "Desconhecido"
    
    df.insert(0, 'Município', df.apply(get_city, axis=1))

    # Colunas Finais
    target = ['Município', 'CNPJ', 'Processo', 'Modalidade', 'Sistema', 'Valor Numérico', 'Valor Original', 'Arquivo', 'Metodo']
    exist = [c for c in target if c in df.columns]
    return df[exist]

# --- 3. UI ---
st.title("🧠 Extrator RFB Híbrido")
st.markdown("""
Esta versão usa **Inteligência de Texto** quando a tabela visual falha (caso de Sonora).
1. Tenta extrair Tabela.
2. Se falhar, varre o texto procurando linhas com 'CNPJ + Processo + Valor'.
3. Se não achar nada, marca como 'Nada Consta'.
""")

files = st.file_uploader("Arraste os PDFs", type="pdf", accept_multiple_files=True)

if files:
    with st.spinner("Processando..."):
        raw = extrair_dados_pdf(files)
    
    if not raw.empty:
        clean = organizar_dados(raw)
        
        st.success(f"Concluído! {len(clean)} registros.")
        
        tab1, tab2 = st.tabs(["📊 Relatório Final", "🔍 Auditoria (Bruto)"])
        
        with tab1:
            st.dataframe(clean, use_container_width=True)
            if 'Valor Numérico' in clean.columns:
                st.metric("Total", f"R$ {clean['Valor Numérico'].sum():,.2f}")
        
        with tab2:
            st.dataframe(raw, use_container_width=True)
        
        out = BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            clean.to_excel(writer, index=False, sheet_name='Final')
            
            wb = writer.book
            ws = writer.sheets['Final']
            fmt = wb.add_format({'num_format': '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'})
            
            if "Valor Numérico" in clean.columns:
                idx = clean.columns.get_loc("Valor Numérico")
                ws.set_column(idx, idx, 18, fmt)
            
            ws.set_column(0, 0, 25)
            ws.set_column(2, 3, 30)
            
            raw.to_excel(writer, index=False, sheet_name='Bruto_Audit')
            
        out.seek(0)
        st.download_button("⬇️ Baixar Excel", out, "Relatorio_RFB_Smart.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("Não foi possível extrair dados. Verifique se os arquivos são PDFs válidos.")
