import streamlit as st
import pandas as pd
import pdfplumber
import pytesseract
from pdf2image import convert_from_bytes
from io import BytesIO
import re

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Extrator RFB (OCR Limpo)", layout="wide", page_icon="👁️")

# --- FUNÇÕES DE EXTRAÇÃO ---

def limpar_texto_ocr(texto):
    """Limpa ruídos comuns de OCR e aspas do formato Sonora"""
    # Remove aspas duplas, simples e pipes que o OCR confunde
    texto = texto.replace('|', ' ').replace("'", "").replace('"', '')
    return texto

def extrair_via_ocr(pdf_bytes):
    """Converte PDF em imagem e extrai texto via Tesseract"""
    try:
        images = convert_from_bytes(pdf_bytes)
        texto_completo = ""
        for img in images:
            texto_pagina = pytesseract.image_to_string(img, lang='por')
            texto_completo += texto_pagina + "\n"
        return texto_completo
    except Exception as e:
        return ""

def processar_texto_extraido(texto, nome_arquivo, metodo):
    """Processa o texto e tenta isolar a Modalidade removendo o resto"""
    dados = []
    
    # 1. Busca CNPJ no cabeçalho
    cnpj_header_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    cnpj_header = cnpj_header_match.group(1) if cnpj_header_match else ""
    
    # 2. Divide em linhas
    linhas = texto.split('\n')
    
    for linha in linhas:
        # Limpeza inicial da linha
        linha_original = linha
        linha = limpar_texto_ocr(linha)
        
        # Padrão: Tem CNPJ e tem Valor Monetário na mesma linha?
        if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", linha) and re.search(r"\d+,\d{2}", linha):
            
            # --- Extração de Dados ---
            
            # 1. CNPJ da linha (usado para localizar, mas removemos depois)
            match_cnpj = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", linha)
            cnpj_linha = match_cnpj.group(0) if match_cnpj else ""
            
            # 2. Valor (último monetário da linha)
            valores = re.findall(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
            valor = valores[-1] if valores else "0,00"
            
            # 3. Processo (número longo >6 dígitos que NÃO seja parte do CNPJ)
            linha_sem_cnpj = linha.replace(cnpj_linha, "")
            processos = re.findall(r"\b\d{6,}\b", linha_sem_cnpj)
            # Filtra números que possam ser restos de data ou ano (ex: 2025)
            processos_validos = [p for p in processos if len(p) > 6] 
            processo = processos_validos[0] if processos_validos else "-"
            
            # 4. MODALIDADE (O que sobrou)
            # Estratégia: Removemos CNPJ, Processo e Valor da linha. O que sobrar é texto.
            resto = linha
            resto = resto.replace(cnpj_linha, "")
            resto = resto.replace(processo, "")
            resto = resto.replace(valor, "")
            
            # Limpa caracteres de pontuação que sobram (, - . ; :)
            modalidade_limpa = re.sub(r'[.,;:\-]', ' ', resto).strip()
            # Remove espaços duplos
            modalidade_limpa = re.sub(r'\s+', ' ', modalidade_limpa)
            
            # Se o que sobrou for muito curto (ex: "SISTEMA"), provavelmente é lixo do OCR
            if len(modalidade_limpa) < 3:
                modalidade_limpa = "-"
            
            dados.append({
                "Arquivo": nome_arquivo,
                "CNPJ": cnpj_header,
                "Processo": processo,
                "Modalidade": modalidade_limpa, # Agora vem limpo ou "-"
                "Sistema": "-", 
                "Valor Original": valor,
                "Metodo": metodo
            })
            
    # Se não achou linhas, mas tem CNPJ -> Nada Consta
    if not dados and cnpj_header:
        dados.append({
            "Arquivo": nome_arquivo,
            "CNPJ": cnpj_header,
            "Processo": "-",
            "Modalidade": "Nada Consta",
            "Sistema": "-",
            "Valor Original": "-",
            "Metodo": f"{metodo} (Vazio)"
        })
        
    return pd.DataFrame(dados)

# --- INTERFACE ---
st.title("👁️ Extrator RFB - OCR (Limpo)")
st.markdown("Extração visual via OCR. O texto 'Extraído via OCR' foi removido.")

# Checkbox
usar_ocr_sempre = st.checkbox("Forçar OCR em todos os arquivos", value=False)

uploaded_files = st.file_uploader("PDFs", type="pdf", accept_multiple_files=True)

if uploaded_files:
    df_final = pd.DataFrame()
    bar = st.progress(0)
    status = st.empty()
    
    for i, f in enumerate(uploaded_files):
        status.text(f"Processando {f.name}...")
        
        texto_extraido = ""
        metodo_usado = ""
        
        # 1. Tentativa Rápida (Texto Nativo)
        if not usar_ocr_sempre:
            try:
                with pdfplumber.open(f) as pdf:
                    for page in pdf.pages:
                        texto_extraido += page.extract_text() or ""
                metodo_usado = "Nativo"
            except:
                texto_extraido = ""
        
        # 2. Verifica se pegou dados úteis
        tem_dados = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto_extraido) and re.search(r"\d+,\d{2}", texto_extraido)
        
        # 3. Fallback para OCR
        if usar_ocr_sempre or not tem_dados:
            status.text(f"Aplicando OCR em {f.name}...")
            f.seek(0)
            texto_extraido = extrair_via_ocr(f.read())
            metodo_usado = "OCR"
            
        # 4. Processa
        df_temp = processar_texto_extraido(texto_extraido, f.name, metodo_usado)
        df_final = pd.concat([df_final, df_temp], ignore_index=True)
        
        bar.progress((i+1)/len(uploaded_files))
        
    status.empty()
    bar.empty()
    
    if not df_final.empty:
        # Tratamento Numérico e Formatação
        def to_num(x):
            s = str(x).replace(' ', '').replace('.', '').replace(',', '.')
            if s in ["-", ""]: return 0.0
            try: return float(s)
            except: return 0.0
        
        if "Valor Original" in df_final.columns:
            df_final["Valor Numérico"] = df_final["Valor Original"].apply(to_num)
            
        def get_mun(row):
            try: return str(row["Arquivo"]).split('-')[1].strip().title()
            except: return str(row["Arquivo"])
        
        if "Arquivo" in df_final.columns:
            df_final.insert(0, "Município", df_final.apply(get_mun, axis=1))
            
        st.success("Extração Concluída!")
        
        # Exibe tabela
        cols = ["Município", "CNPJ", "Processo", "Modalidade", "Sistema", "Valor Numérico", "Valor Original", "Arquivo", "Metodo"]
        cols = [c for c in cols if c in df_final.columns]
        st.dataframe(df_final[cols], use_container_width=True)
        
        # Download
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            df_final[cols].to_excel(writer, index=False, sheet_name="Dados")
            wb = writer.book
            ws = writer.sheets['Dados']
            fmt = wb.add_format({'num_format': '#,##0.00'})
            
            if "Valor Numérico" in df_final.columns:
                idx = df_final.columns.get_loc("Valor Numérico") # Correção aqui (busca no df_final original ou filtrado)
                # Como filtramos cols, vamos pegar o índice certo da lista cols
                try:
                    idx_excel = cols.index("Valor Numérico")
                    ws.set_column(idx_excel, idx_excel, 18, fmt)
                except: pass
            
            ws.set_column(0, 0, 25)
            ws.set_column(2, 3, 30)
            
        buf.seek(0)
        st.download_button("⬇️ Baixar Excel", buf, "Relatorio_RFB.xlsx")
    else:
        st.error("Não foi possível extrair dados.")
