import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Extrator RFB Detalhado", layout="wide", page_icon="📊")

# --- FUNÇÃO DE EXTRAÇÃO (Lógica Aprimorada) ---
def extrair_dados_pdf(pdf_bytes, nome_arquivo):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto_completo = ""
    for pagina in doc:
        texto_completo += pagina.get_text()
    
    # 1. Regex para Município e CNPJ
    municipio_match = re.search(r"MUNICIPIO DE\s+(.*)", texto_completo)
    municipio = municipio_match.group(1).strip().upper() if municipio_match else "DESCONHECIDO"
    
    cnpj_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_completo)
    cnpj = cnpj_match.group(1) if cnpj_match else ""

    # 2. Extração da Tabela (Linha a Linha)
    parcelamentos = []
    linhas = [l.strip() for l in texto_completo.split('\n') if l.strip()]
    
    for i, linha in enumerate(linhas):
        # Verifica se o CNPJ aparece na linha (não precisa ser no início)
        if cnpj and cnpj in linha:
            # Tenta capturar o ID do Parcelamento/Processo (Sequência numérica longa, geralmente 9+ dígitos)
            # Ignora o próprio CNPJ na busca de números longos
            numeros_na_linha = re.findall(r"\d{9,}", linha)
            
            # Remove números que pareçam ser parte do CNPJ (limpeza básica)
            numeros_limpos = [n for n in numeros_na_linha if n not in cnpj.replace('.', '').replace('/', '').replace('-', '')]
            
            # O processo geralmente é o primeiro número longo encontrado após o CNPJ
            processo = numeros_limpos[0] if numeros_limpos else "Não identificado"

            # Tenta capturar o Valor (Formato R$ com vírgula: 1.000,00)
            match_valor = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
            
            if match_valor:
                valor = match_valor.group(1)
                
                # ADICIONA O ITEM ENCONTRADO
                parcelamentos.append({
                    "Origem": "Linha Detalhada",
                    "Processo": processo, 
                    "Saldo": valor
                })
            else:
                # Se achou o processo mas o valor quebrou para a linha de baixo
                if i + 1 < len(linhas):
                    prox_valor = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linhas[i+1])
                    if prox_valor:
                        parcelamentos.append({
                            "Origem": "Linha Detalhada (Quebra)",
                            "Processo": processo, 
                            "Saldo": prox_valor.group(1)
                        })

    # 3. Fallback: Se a lista estiver vazia, pega o Total Geral e o Número do Dossiê
    if not parcelamentos:
        saldo_total = re.search(r"SALDO DEVEDOR TOTAL\s+([\d.,]+)", texto_completo)
        valor_total = saldo_total.group(1) if saldo_total else "0,00"
        
        # Pega o número do processo principal (Dossiê) no cabeçalho
        dossie_match = re.search(r"No Processo/Dossiê\s+([\d./-]+)", texto_completo)
        proc_ref = dossie_match.group(1) if dossie_match else "Consolidado"
        
        parcelamentos.append({
            "Origem": "Consolidado (Total)",
            "Processo": proc_ref, 
            "Saldo": valor_total
        })

    return {"Arquivo": nome_arquivo, "Município": municipio, "CNPJ": cnpj, "Parcelamentos": parcelamentos}

# --- INTERFACE ---
st.title("📊 Extrator de Parcelamentos - RFB")
st.markdown("Extrai cada linha de negociação individualmente.")

uploaded_files = st.file_uploader("Arraste os PDFs aqui", type="pdf", accept_multiple_files=True)

if uploaded_files:
    lista_final = []
    
    with st.spinner("Processando..."):
        for f in uploaded_files:
            dados = extrair_dados_pdf(f.read(), f.name)
            
            for p in dados['Parcelamentos']:
                lista_final.append({
                    "Arquivo": dados['Arquivo'],
                    "Município": dados['Município'],
                    "CNPJ": dados['CNPJ'],
                    "Tipo Extração": p['Origem'],
                    "Processo/Negociação": p['Processo'],
                    "Saldo Devedor (R$)": p['Saldo']
                })
    
    if lista_final:
        df = pd.DataFrame(lista_final)
        st.success(f"✅ Processado! Foram encontradas {len(df)} linhas de débitos.")
        st.dataframe(df, use_container_width=True)
        
        output = BytesIO()
        # Engine ajustada para xlsxwriter (lembre-se do requirements.txt)
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Saldos_Detalhados')
            
        output.seek(0)
        st.download_button("⬇️ Baixar Excel Detalhado", output, "Saldos_RFB_Detalhados.xlsx")
    else:
        st.warning("Nenhum dado encontrado.")
