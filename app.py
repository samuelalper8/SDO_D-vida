import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

def extrair_dados_rfb(pdf_stream):
    doc = fitz.open(stream=pdf_stream, filetype="pdf")
    texto_completo = ""
    for pagina in doc:
        texto_completo += pagina.get_text()
    
    # Regex para capturar os campos específicos baseados nos seus documentos
    municipio = re.search(r"MUNICIPIO DE\s+(.*)", texto_completo)
    cnpj = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_completo)
    processo = re.search(r"No Processo/Dossiê\s+([\d./-]+)", texto_completo)
    
    # Captura o saldo devedor total (procurando o padrão numérico após o título da tabela)
    saldo_total = re.search(r"SALDO DEVEDOR TOTAL\s+([\d.,]+)", texto_completo)
    
    return {
        "Município": municipio.group(1).strip() if municipio else "Não encontrado",
        "CNPJ": cnpj.group(1) if cnpj else "Não encontrado",
        "Processo": processo.group(1) if processo else "Não encontrado",
        "Saldo em 31/12/2025": saldo_total.group(1) if saldo_total else "0,00"
    }

st.set_page_config(page_title="Extrator Fisco Federal", layout="wide")

st.title("📂 Processador de Saldos RFB - ConPrev")
st.subheader("Automação para Balanço Patrimonial")

arquivos = st.file_uploader("Arraste os PDFs de saldo devedor aqui", type="pdf", accept_multiple_files=True)

if arquivos:
    resultados = []
    
    for arq in arquivos:
        dados = extrair_dados_rfb(arq.read())
        resultados.append(dados)
    
    df = pd.DataFrame(resultados)
    
    st.write("### Dados Extraídos")
    st.dataframe(df, use_container_width=True)
    
    # Exportação para Excel/CSV para seu controle
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button("Baixar Planilha de Controle (CSV)", csv, "saldos_extraidos.csv", "text/csv")

    st.info("💡 Esses dados podem agora ser injetados no seu template de Ofício DFF.")
