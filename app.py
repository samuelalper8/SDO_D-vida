import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Extrator RFB (100% pdfplumber)", layout="wide", page_icon="🧩")

# --- 1. MOTOR TABELA (Para arquivos com linhas visíveis) ---
def extrair_tabela(pdf):
    for page in pdf.pages:
        # Tenta extração padrão
        tables = page.extract_tables()
        for table in tables:
            if not table: continue
            df = pd.DataFrame(table)
            
            # Limpa cabeçalho
            header = str(df.iloc[0].values).upper().replace('\n', ' ')
            if "PROCESSO" in header and ("SALDO" in header or "VALOR" in header):
                if len(df) > 1:
                    df.columns = df.iloc[0]
                    return df[1:], "Tabela"
    return pd.DataFrame(), None

# --- 2. MOTOR TEXTO CSV (Específico para SONORA) ---
def extrair_texto_csv(pdf):
    # Concatena texto de todas as páginas
    texto_completo = ""
    for page in pdf.pages:
        texto_completo += page.extract_text() or ""
    
    # O PDF de Sonora tem dados entre aspas separados por vírgula, muitas vezes com quebras de linha
    # Ex: "CNPJ","PROCESSO"
    
    # Regex para capturar linhas CSV: "XX.XXX.XXX/XXXX-XX","XXXXX","XXXX","XXXX","XXXX,XX"
    # O flag re.DOTALL permite que o ponto (.) pegue quebras de linha dentro das aspas
    padrao = r'\"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\"\s*,\s*\"(.*?)\"\s*,\s*\"(.*?)\"\s*,\s*\"(.*?)\"\s*,\s*\"(.*?)\"'
    
    matches = re.findall(padrao, texto_completo, re.DOTALL)
    
    if matches:
        lista = []
        for m in matches:
            lista.append({
                "CNPJ VINCULADO": m[0],
                "PROCESSO/ID PARCELAMENTO": m[1].replace('\n', '').strip(),
                "MODALIDADE": m[2].replace('\n', ' ').strip(),
                "SISTEMA": m[3].replace('\n', '').strip(),
                "SALDO DEVEDOR": m[4].replace('\n', '').strip()
            })
        return pd.DataFrame(lista), "Texto CSV (Sonora)"
        
    return pd.DataFrame(), None

# --- 3. MOTOR NADA CONSTA (Varredura Simples) ---
def extrair_nada_consta(pdf, filename):
    texto = pdf.pages[0].extract_text() or ""
    cnpj_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    cnpj = cnpj_match.group(1) if cnpj_match else ""
    
    if cnpj:
        return pd.DataFrame([{
            "CNPJ VINCULADO": cnpj,
            "PROCESSO/ID PARCELAMENTO": "-",
            "MODALIDADE": "Nada Consta",
            "SISTEMA": "-",
            "SALDO DEVEDOR": "-"
        }]), "Nada Consta"
    return pd.DataFrame(), None

# --- PROCESSAMENTO PRINCIPAL ---
def processar(files):
    final = pd.DataFrame()
    
    progresso = st.progress(0)
    status = st.empty()
    
    for i, f in enumerate(files):
        status.text(f"Lendo {f.name}...")
        
        try:
            with pdfplumber.open(f) as pdf:
                # 1. Tenta Tabela (Melhor qualidade)
                df, metodo = extrair_tabela(pdf)
                
                # 2. Se falhar, tenta o formato "Sonora" (Aspas/CSV)
                if df.empty:
                    df, metodo = extrair_texto_csv(pdf)
                
                # 3. Se falhar, assume Nada Consta (mas pega o CNPJ do texto)
                if df.empty:
                    df, metodo = extrair_nada_consta(pdf, f.name)
                
                if not df.empty:
                    df["Arquivo"] = f.name
                    df["Metodo"] = metodo
                    
                    # Padronização de Colunas
                    col_map = {}
                    for c in df.columns:
                        cu = str(c).upper()
                        if "PROCESSO" in cu: col_map[c] = "Processo"
                        elif "MODALIDADE" in cu: col_map[c] = "Modalidade"
                        elif "SISTEMA" in cu: col_map[c] = "Sistema"
                        elif "SALDO" in cu or "VALOR" in cu: col_map[c] = "Valor Original"
                        elif "CNPJ" in cu: col_map[c] = "CNPJ"
                    
                    df = df.rename(columns=col_map)
                    final = pd.concat([final, df], ignore_index=True)
                    
        except Exception as e:
            st.error(f"Erro em {f.name}: {e}")
            
        progresso.progress((i+1)/len(files))
        
    status.empty()
    progresso.empty()
    return final

# --- UI ---
st.title("🧩 Extrator RFB (Compatível)")
st.markdown("Extrai dados de Tabelas e Textos CSV (Sonora) sem depender de bibliotecas externas complexas.")

files = st.file_uploader("PDFs", type="pdf", accept_multiple_files=True)

if files:
    df_result = processar(files)
    
    if not df_result.empty:
        # Tratamento Final
        
        # Valor Numérico
        def to_num(x):
            s = str(x).replace(' ', '').replace('.', '').replace(',', '.')
            if s in ["-", ""]: return 0.0
            try: return float(s)
            except: return 0.0
            
        if "Valor Original" in df_result.columns:
            df_result["Valor Numérico"] = df_result["Valor Original"].apply(to_num)
            
        # Município Title Case
        def get_mun(row):
            try: return str(row["Arquivo"]).split('-')[1].strip().title()
            except: return str(row["Arquivo"])
            
        if "Arquivo" in df_result.columns:
            df_result.insert(0, "Município", df_result.apply(get_mun, axis=1))

        # Colunas Finais
        cols = ["Município", "CNPJ", "Processo", "Modalidade", "Sistema", "Valor Numérico", "Valor Original", "Arquivo", "Metodo"]
        cols = [c for c in cols if c in df_result.columns]
        df_show = df_result[cols].copy()

        st.success("Concluído!")
        st.dataframe(df_show, use_container_width=True)
        
        # Download
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            df_show.to_excel(writer, index=False, sheet_name="Dados")
            wb = writer.book
            ws = writer.sheets['Dados']
            fmt = wb.add_format({'num_format': '#,##0.00'})
            
            if "Valor Numérico" in df_show.columns:
                idx = df_show.columns.get_loc("Valor Numérico")
                ws.set_column(idx, idx, 15, fmt)
            
            ws.set_column(0, 0, 25)
            ws.set_column(2, 3, 30)
            
        buf.seek(0)
        st.download_button("⬇️ Baixar Excel", buf, "Relatorio_RFB.xlsx")
