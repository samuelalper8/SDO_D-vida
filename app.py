import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Extrator RFB (Universal)", layout="wide", page_icon="💎")

# --- MOTOR 1: TABELAS PADRÃO (PDFPLUMBER) ---
def extrair_via_tabela(pdf_obj):
    try:
        # Tenta modo 'lattice' (bordas) e depois 'stream' (espaçamento)
        for flavor in ['lattice', 'stream']:
            with pdfplumber.open(pdf_obj) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"}) if flavor == 'stream' else page.extract_tables()
                    
                    for table in tables:
                        if not table: continue
                        df = pd.DataFrame(table)
                        if df.empty: continue
                        
                        # Limpa cabeçalho
                        header = str(df.iloc[0].values).upper().replace('\n', ' ')
                        if "PROCESSO" in header and ("SALDO" in header or "VALOR" in header):
                            if len(df) > 1:
                                df.columns = df.iloc[0]
                                return df[1:], "Tabela"
    except: pass
    return pd.DataFrame(), None

# --- MOTOR 2: TEXTO CSV / ASPAS (ESPECÍFICO PARA SONORA) ---
def extrair_via_texto_csv(pdf_bytes):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        texto = ""
        for page in doc: texto += page.get_text()
        
        # Procura padrão: "CNPJ","PROCESSO","MODALIDADE","SISTEMA","VALOR"
        # O arquivo de Sonora tem esse formato exato escondido no texto
        
        # Regex para capturar linhas que tenham formato "dado","dado"
        # Captura linhas que começam com aspas, tem um CNPJ no meio e terminam com aspas
        linhas_csv = re.findall(r'\"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\"\s*,\s*\"(.*?)\"\s*,\s*\"(.*?)\"\s*,\s*\"(.*?)\"\s*,\s*\"(.*?)\"', texto)
        
        if linhas_csv:
            lista_dicts = []
            for lin in linhas_csv:
                # lin é uma tupla: (CNPJ, Processo, Modalidade, Sistema, Valor)
                lista_dicts.append({
                    "CNPJ VINCULADO": lin[0],
                    "PROCESSO/ID PARCELAMENTO": lin[1].replace('\n', ''),
                    "MODALIDADE": lin[2].replace('\n', ''),
                    "SISTEMA": lin[3].replace('\n', ''),
                    "SALDO DEVEDOR": lin[4].replace('\n', '')
                })
            return pd.DataFrame(lista_dicts), "Texto CSV"
            
    except: pass
    return pd.DataFrame(), None

# --- MOTOR 3: VARREDURA GENÉRICA (ÚLTIMO RECURSO) ---
def extrair_via_varredura(pdf_bytes, cnpj_header):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        texto = ""
        for page in doc: texto += page.get_text()
        
        lista = []
        # Procura linhas soltas com CNPJ e Valor
        linhas = texto.split('\n')
        for i, linha in enumerate(linhas):
            if cnpj_header in linha and re.search(r'\d+,\d{2}', linha):
                # Tenta isolar o processo (número grande)
                processos = re.findall(r'\b\d{7,}\b', linha.replace(cnpj_header, ''))
                proc = processos[0] if processos else "-"
                
                # Tenta isolar valor (último monetário)
                valores = re.findall(r'(\d{1,3}(?:\.\d{3})*,\d{2})', linha)
                val = valores[-1] if valores else "0,00"
                
                lista.append({
                    "CNPJ VINCULADO": cnpj_header,
                    "PROCESSO/ID PARCELAMENTO": proc,
                    "MODALIDADE": "Varredura Texto",
                    "SISTEMA": "-",
                    "SALDO DEVEDOR": val
                })
        
        if lista: return pd.DataFrame(lista), "Varredura"
    except: pass
    return pd.DataFrame(), None

# --- CONTROLADOR PRINCIPAL ---
def processar_arquivos(files):
    consolidado = pd.DataFrame()
    
    progresso = st.progress(0)
    msg = st.empty()
    
    for i, f in enumerate(files):
        msg.text(f"Processando {f.name}...")
        
        # 1. Pega CNPJ do Cabeçalho (fundamental para o Nada Consta)
        f.seek(0)
        doc_temp = fitz.open(stream=f.read(), filetype="pdf")
        txt_full = "".join([p.get_text() for p in doc_temp])
        cnpj_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", txt_full)
        cnpj_header = cnpj_match.group(1) if cnpj_match else ""
        
        # RESET PONTEIRO
        f.seek(0)
        
        # TENTATIVA 1: TABELAS
        df, metodo = extrair_via_tabela(f)
        
        # TENTATIVA 2: CSV (SONORA)
        if df.empty:
            f.seek(0)
            df, metodo = extrair_via_texto_csv(f.read())
            
        # TENTATIVA 3: VARREDURA GENÉRICA
        if df.empty and cnpj_header:
            f.seek(0)
            df, metodo = extrair_via_varredura(f.read(), cnpj_header)
            
        # RESULTADO FINAL: DADOS OU NADA CONSTA
        if not df.empty:
            df["Arquivo"] = f.name
            df["Metodo"] = metodo
            # Normaliza nomes de colunas
            col_map = {}
            for c in df.columns:
                c_upper = str(c).upper()
                if "PROCESSO" in c_upper: col_map[c] = "Processo"
                elif "MODALIDADE" in c_upper: col_map[c] = "Modalidade"
                elif "SISTEMA" in c_upper: col_map[c] = "Sistema"
                elif "SALDO" in c_upper or "VALOR" in c_upper: col_map[c] = "Valor Original"
                elif "CNPJ" in c_upper: col_map[c] = "CNPJ"
            
            df = df.rename(columns=col_map)
            consolidado = pd.concat([consolidado, df], ignore_index=True)
        else:
            # Nada Consta
            consolidado = pd.concat([consolidado, pd.DataFrame([{
                "Arquivo": f.name,
                "CNPJ": cnpj_header,
                "Processo": "-",
                "Modalidade": "Nada Consta",
                "Sistema": "-",
                "Valor Original": "-",
                "Metodo": "Nada Consta"
            }])], ignore_index=True)
            
        progresso.progress((i+1)/len(files))
        
    msg.empty()
    progresso.empty()
    return consolidado

# --- INTERFACE ---
st.title("💎 Extrator RFB Universal")
st.markdown("Algoritmo triplo: Tabela Padrão + CSV Embutido (Sonora) + Varredura.")

files = st.file_uploader("PDFs", type="pdf", accept_multiple_files=True)

if files:
    df_final = processar_arquivos(files)
    
    if not df_final.empty:
        # Tratamento Final
        
        # 1. Limpeza de Valor
        def limpar_valor(x):
            s = str(x).replace(' ', '').replace('.', '').replace(',', '.')
            if s in ["-", ""]: return 0.0
            try: return float(s)
            except: return 0.0
            
        if "Valor Original" in df_final.columns:
            df_final["Valor Numérico"] = df_final["Valor Original"].apply(limpar_valor)
        
        # 2. Município Title Case
        def get_mun(row):
            try: return str(row["Arquivo"]).split('-')[1].strip().title()
            except: return str(row["Arquivo"])
        
        if "Arquivo" in df_final.columns:
            df_final.insert(0, "Município", df_final.apply(get_mun, axis=1))

        # 3. Colunas Finais
        cols = ["Município", "CNPJ", "Processo", "Modalidade", "Sistema", "Valor Numérico", "Valor Original", "Arquivo", "Metodo"]
        cols = [c for c in cols if c in df_final.columns]
        df_exibir = df_final[cols].copy()

        st.success("Extração Concluída!")
        st.dataframe(df_exibir, use_container_width=True)
        
        # Download
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_exibir.to_excel(writer, index=False, sheet_name="Dados")
            wb = writer.book
            ws = writer.sheets['Dados']
            fmt = wb.add_format({'num_format': '#,##0.00'})
            if "Valor Numérico" in df_exibir.columns:
                idx = df_exibir.columns.get_loc("Valor Numérico")
                ws.set_column(idx, idx, 15, fmt)
            ws.set_column(0, 0, 25) # Mun
            ws.set_column(2, 3, 25) # Proc/Mod
            
        buffer.seek(0)
        st.download_button("⬇️ Baixar Excel", buffer, "Relatorio_RFB_Universal.xlsx")
