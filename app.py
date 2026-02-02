import fitz  # PyMuPDF
import re

def extrair_dados_pdf(caminho_ou_bytes_pdf, nome_arquivo="Desconhecido"):
    """
    Função pura para extrair dados de Saldo Devedor da RFB.
    Pode ser usada em qualquer script (Streamlit, Flask, Terminal, Automação).
    """
    
    # Abre o PDF (aceita caminho do arquivo ou bytes na memória)
    if isinstance(caminho_ou_bytes_pdf, (str, bytes)):
        # Se for bytes (upload) usa stream, se for string (caminho local) abre direto
        if isinstance(caminho_ou_bytes_pdf, bytes):
            doc = fitz.open(stream=caminho_ou_bytes_pdf, filetype="pdf")
        else:
            doc = fitz.open(caminho_ou_bytes_pdf)
    else:
        # Caso já seja um objeto fitz aberto
        doc = caminho_ou_bytes_pdf

    texto_completo = ""
    for pagina in doc:
        texto_completo += pagina.get_text()
    
    # 1. Regex para Município
    municipio_match = re.search(r"MUNICIPIO DE\s+(.*)", texto_completo)
    municipio = municipio_match.group(1).strip() if municipio_match else "DESCONHECIDO"
    
    # 2. Regex para CNPJ
    cnpj_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_completo)
    cnpj = cnpj_match.group(1) if cnpj_match else ""

    # 3. Lógica de extração da tabela
    parcelamentos = []
    linhas = [l.strip() for l in texto_completo.split('\n') if l.strip()]
    
    for i, linha in enumerate(linhas):
        # A linha de dados geralmente começa com o CNPJ
        if cnpj and linha.startswith(cnpj):
            # Procura processo (sequência longa de números)
            match_processo = re.search(r"\s+(\d{9,})", linha)
            # Procura valor monetário
            match_valor = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
            
            if match_processo:
                valor = match_valor.group(1) if match_valor else "Verificar"
                
                # Fallback: Se não achou valor na mesma linha, olha a próxima
                if not match_valor and i + 1 < len(linhas):
                    prox_valor = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linhas[i+1])
                    if prox_valor: valor = prox_valor.group(1)
                
                parcelamentos.append({"Processo": match_processo.group(1), "Saldo": valor})

    # 4. Fallback se não encontrar parcelas individuais (Pega o total geral)
    if not parcelamentos:
        saldo_total = re.search(r"SALDO DEVEDOR TOTAL\s+([\d.,]+)", texto_completo)
        valor_total = saldo_total.group(1) if saldo_total else "0,00"
        
        # Tenta pegar o número do dossiê no cabeçalho se não houver parcelas
        dossie = re.search(r"No Processo/Dossiê\s+([\d./-]+)", texto_completo)
        proc_ref = dossie.group(1) if dossie else "Consolidado"
        
        parcelamentos.append({"Processo": proc_ref, "Saldo": valor_total})

    return {
        "Arquivo": nome_arquivo,
        "Município": municipio,
        "CNPJ": cnpj,
        "Parcelamentos": parcelamentos
    }
