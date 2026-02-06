import streamlit as st
import pandas as pd
import pdfplumber
import pytesseract
from pdf2image import convert_from_bytes
from io import BytesIO
import re

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Extrator RFB com OCR", layout="wide", page_icon="👁️")

# --- FUNÇÕES DE EXTRAÇÃO ---

def limpar_texto_ocr(texto):
    """Limpa ruídos comuns de OCR"""
    texto = texto.replace('|', '').replace("'", "").replace('"', '')
    return texto

def extrair_via_ocr(pdf_bytes):
    """Converte PDF em imagem e extrai texto via Tesseract"""
    try:
        # Converte páginas do PDF em imagens
        images = convert_from_bytes(pdf_bytes)
        texto_completo = ""
        
        for img in images:
            # Configuração para português e layout de bloco
            texto_pagina = pytesseract.image_to_string(img, lang='por')
            texto_completo += texto_pagina + "\n"
            
        return texto_completo
    except Exception as e:
        return f"Erro no OCR: {str(e)}"

def processar_texto_extraido(texto, nome_arquivo, metodo):
    """Processa o texto (seja do OCR ou do PDF nativo) e busca os padrões"""
    dados = []
    
    # 1. Busca CNPJ no cabeçalho (para referência)
    cnpj_header_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    cnpj_header = cnpj_header_match.group(1) if cnpj_header_match else ""
    
    # 2. Divide em linhas
    linhas = texto.split('\n')
    
    for linha in linhas:
        linha = limpar_texto_ocr(linha)
        
        # Padrão Genérico: Tem CNPJ e tem Valor Monetário na mesma linha?
        if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", linha) and re.search(r"\d+,\d{2}", linha):
            
            # Remove o CNPJ para não confundir com o processo
            linha_sem_cnpj = re.sub(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", "", linha)
            
            # Busca Processo (número longo, >6 dígitos)
            processos = re.findall(r"\b\d{6,}\b", linha_sem_cnpj)
            processo = processos[0] if processos else "-"
            
            # Busca Valor (último valor monetário da linha)
            valores = re.findall(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
            valor = valores[-1] if valores else "0,00"
            
            # Busca Modalidade (Texto que sobrou)
            # Removemos os números e caracteres conhecidos, o resto é modalidade
            # Essa é uma limpeza aproximada
            dados.append({
                "Arquivo": nome_arquivo,
                "CNPJ": cnpj_header,
                "Processo": processo,
                "Modalidade": "Extraído via OCR/Texto", # OCR dificulta pegar o texto exato da modalidade limpo
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
st.title("👁️ Extrator RFB - OCR Ativado")
st.markdown("""
**Modo Avançado:** Se a leitura padrão falhar, o sistema converterá o PDF em imagem e lerá visualmente (OCR).
*Requer configuração de `packages.txt` no Streamlit Cloud.*
""")

# Opção para forçar OCR em todos
usar_ocr_sempre = st.checkbox("Forçar OCR em todos os arquivos (Mais lento, porém mais garantido para arquivos problemáticos)", value=False)

uploaded_files = st.file_uploader("PDFs", type="pdf", accept_multiple_files=True)

if uploaded_files:
    df_final = pd.DataFrame()
    bar = st.progress(0)
    status = st.empty()
    
    for i, f in enumerate(uploaded_files):
        status.text(f"Processando {f.name}...")
        
        # Estratégia: Tenta pdfplumber primeiro (rápido). Se falhar ou se "Forçar OCR" estiver on, usa OCR.
        
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
                texto_extraido = "" # Falhou
        
        # 2. Validação: O texto nativo tem dados úteis? (Pelo menos um CNPJ e um Valor?)
        tem_dados = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto_extraido) and re.search(r"\d+,\d{2}", texto_extraido)
        
        # 3. Tentativa OCR (Se forçado ou se o nativo falhou em achar dados)
        if usar_ocr_sempre or not tem_dados:
            status.text(f"Aplicando OCR em {f.name} (pode demorar)...")
            f.seek(0) # Reseta ponteiro
            texto_extraido = extrair_via_ocr(f.read())
            metodo_usado = "OCR"
            
        # 4. Processamento do Texto Final
        df_temp = processar_texto_extraido(texto_extraido, f.name, metodo_usado)
        df_final = pd.concat([df_final, df_temp], ignore_index=True)
        
        bar.progress((i+1)/len(uploaded_files))
        
    status.empty()
    bar.empty()
    
    if not df_final.empty:
        # Tratamento Final de Dados
        
        # Numérico
        def to_num(x):
            s = str(x).replace(' ', '').replace('.', '').replace(',', '.')
            if s in ["-", ""]: return 0.0
            try: return float(s)
            except: return 0.0
        
        if "Valor Original" in df_final.columns:
            df_final["Valor Numérico"] = df_final["Valor Original"].apply(to_num)
            
        # Município
        def get_mun(row):
            try: return str(row["Arquivo"]).split('-')[1].strip().title()
            except: return str(row["Arquivo"])
        
        if "Arquivo" in df_final.columns:
            df_final.insert(0, "Município", df_final.apply(get_mun, axis=1))
            
        st.success("Extração Concluída!")
        st.dataframe(df_final, use_container_width=True)
        
        # Download Excel
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name="Dados")
            wb = writer.book
            ws = writer.sheets['Dados']
            fmt = wb.add_format({'num_format': '#,##0.00'})
            
            if "Valor Numérico" in df_final.columns:
                idx = df_final.columns.get_loc("Valor Numérico")
                ws.set_column(idx, idx, 15, fmt)
            
            ws.set_column(0, 0, 25)
            ws.set_column(2, 3, 30)
            
        buf.seek(0)
        st.download_button("⬇️ Baixar Excel (OCR)", buf, "Relatorio_RFB_OCR.xlsx")
    else:
        st.error("Não foi possível extrair dados de nenhum arquivo.")
