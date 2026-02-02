import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Extrator de Dados RFB", layout="wide", page_icon="📊")

# --- FUNÇÃO DE EXTRAÇÃO (O Motor) ---
def extrair_dados_pdf(pdf_bytes, nome_arquivo):
    """Lê o PDF e retorna os dados brutos."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto_completo = ""
    for pagina in doc:
        texto_completo += pagina.get_text()
    
    # 1. Regex para Município e CNPJ
    municipio_match = re.search(r"MUNICIPIO DE\s+(.*)", texto_completo)
    municipio = municipio_match.group(1).strip() if municipio_match else "DESCONHECIDO"
    
    cnpj_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_completo)
    cnpj = cnpj_match.group(1) if cnpj_match else ""

    # 2. Lógica de extração da tabela (Parcelamentos)
    parcelamentos = []
    linhas = [l.strip() for l in texto_completo.split('\n') if l.strip()]
    
    for i, linha in enumerate(linhas):
        if cnpj and linha.startswith(cnpj):
            match_processo = re.search(r"\s+(\d{9,})", linha)
            match_valor = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
            
            if match_processo:
                valor = match_valor.group(1) if match_valor else "Verificar"
                # Se não achou valor na linha, tenta na próxima (quebra de linha do PDF)
                if not match_valor and i + 1 < len(linhas):
                    prox_valor = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linhas[i+1])
                    if prox_valor: valor = prox_valor.group(1)
                
                parcelamentos.append({"Processo": match_processo.group(1), "Saldo": valor})

    # 3. Fallback (Se não achou parcelas, pega total)
    if not parcelamentos:
        saldo_total = re.search(r"SALDO DEVEDOR TOTAL\s+([\d.,]+)", texto_completo)
        valor_total = saldo_total.group(1) if saldo_total else "0,00"
        dossie = re.search(r"No Processo/Dossiê\s+([\d./-]+)", texto_completo)
        proc_ref = dossie.group(1) if dossie else "Consolidado"
        parcelamentos.append({"Processo": proc_ref, "Saldo": valor_total})

    return {"Arquivo": nome_arquivo, "Município": municipio, "CNPJ": cnpj, "Parcelamentos": parcelamentos}

# --- INTERFACE VISUAL (O Painel) ---
st.title("📊 Extrator de Saldos - Receita Federal")
st.markdown("Faça upload dos PDFs para gerar uma planilha com os saldos devedores.")

uploaded_files = st.file_uploader("Arraste os arquivos PDF aqui", type="pdf", accept_multiple_files=True)

if uploaded_files:
    lista_final = []
    
    with st.spinner("Processando arquivos..."):
        for f in uploaded_files:
            # Chama o motor de extração
            dados = extrair_dados_pdf(f.read(), f.name)
            
            # "Achata" os dados para o formato de tabela (Excel)
            # Se tiver 3 processos, cria 3 linhas para o mesmo município
            for p in dados['Parcelamentos']:
                lista_final.append({
                    "Arquivo": dados['Arquivo'],
                    "Município": dados['Município'],
                    "CNPJ": dados['CNPJ'],
                    "Processo/Dossiê": p['Processo'],
                    "Saldo Devedor (R$)": p['Saldo']
                })
    
    # Cria o DataFrame
    if lista_final:
        df = pd.DataFrame(lista_final)
        
        st.success(f"✅ Sucesso! {len(uploaded_files)} arquivos processados.")
        
        # Mostra a tabela na tela
        st.dataframe(df, use_container_width=True)
        
        # Botão de Download para Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Saldos_RFB')
            
        output.seek(0)
        
        st.download_button(
            label="⬇️ Baixar Planilha Excel (.xlsx)",
            data=output,
            file_name="Relatorio_Saldos_RFB.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nenhum dado encontrado nos arquivos enviados.")
