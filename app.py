import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Extrator RFB (Alta Fidelidade)", layout="wide", page_icon="🎯")

def extrair_dados_pdf(pdf_bytes, nome_arquivo):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto_completo = ""
    for pagina in doc:
        # Tenta preservar o layout físico para evitar que colunas se misturem
        texto_completo += pagina.get_text("text") + "\n"
    
    # 1. Cabeçalhos (Metadados)
    municipio_match = re.search(r"MUNICIPIO DE\s+(.*)", texto_completo)
    municipio = municipio_match.group(1).strip().upper() if municipio_match else "DESCONHECIDO"
    
    # Captura CNPJ padrão do cabeçalho
    cnpj_header_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_completo)
    cnpj_header = cnpj_header_match.group(1) if cnpj_header_match else ""

    parcelamentos = []
    
    # Divide linhas e remove vazias
    linhas = [l.strip() for l in texto_completo.split('\n') if l.strip()]
    
    for i, linha in enumerate(linhas):
        # Critério: A linha deve conter o CNPJ (mesmo que com espaçamento zoado) ou partes dele
        # Mas para garantir fidedignidade, vamos buscar o padrão de CNPJ na linha
        cnpj_na_linha = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", linha)
        
        # Se achou um CNPJ na linha e tem cara de linha de tabela (tem valor monetário no final)
        if cnpj_na_linha and re.search(r"\d+,\d{2}", linha):
            
            # Remove o CNPJ da string para ele não ser confundido com o processo
            # (Substitui por vazio)
            linha_sem_cnpj = linha.replace(cnpj_na_linha.group(0), "")
            
            # --- ESTRATÉGIA DE CAPTURA DO PROCESSO ---
            processo_encontrado = "N/D"
            
            # Prioridade 1: Buscar formato NUP (Ex: 10265.438351/2022-56)
            match_nup = re.search(r"\d{5}\.\d{6}/\d{4}-\d{2}", linha_sem_cnpj)
            
            if match_nup:
                processo_encontrado = match_nup.group(0)
            else:
                # Prioridade 2: Buscar ID Numérico Longo (Ex: 10120729679201251 ou 641617569)
                # Buscamos números com 7 ou mais dígitos que sobraram na linha
                ids_numericos = re.findall(r"\b\d{7,}\b", linha_sem_cnpj)
                if ids_numericos:
                    # Pega o primeiro número longo encontrado (geralmente é o processo)
                    processo_encontrado = ids_numericos[0]

            # --- ESTRATÉGIA DE CAPTURA DO VALOR ---
            # Pega todos os valores monetários (X.XXX,XX)
            match_valores = re.findall(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
            valor = match_valores[-1] if match_valores else "0,00" # O último costuma ser o saldo
            
            parcelamentos.append({
                "Processo": processo_encontrado,
                "Saldo": valor,
                "Origem": "Tabela Detalhada"
            })

    # Fallback: Se não achou linhas na tabela, busca o Total Geral e o Dossiê do Cabeçalho
    if not parcelamentos:
        saldo_total = re.search(r"SALDO DEVEDOR TOTAL\s+([\d.,]+)", texto_completo)
        valor_total = saldo_total.group(1) if saldo_total else "0,00"
        
        # Busca dossiê no cabeçalho
        dossie_header = re.search(r"No Processo/Dossiê\s+([\d./-]+)", texto_completo)
        proc_ref = dossie_header.group(1) if dossie_header else "Consolidado"
        
        if valor_total != "0,00":
            parcelamentos.append({
                "Processo": proc_ref,
                "Saldo": valor_total,
                "Origem": "Valor Consolidado"
            })
        else:
             parcelamentos.append({
                "Processo": "-",
                "Saldo": "0,00",
                "Origem": "Sem Débitos"
            })

    return {
        "Arquivo": nome_arquivo,
        "Município": municipio, 
        "CNPJ": cnpj_header, 
        "Parcelamentos": parcelamentos
    }

# --- INTERFACE ---
st.title("🎯 Extrator RFB - Alta Precisão")
st.markdown("Extrai IDs numéricos e Processos Administrativos (NUP) com formatação correta.")

uploaded_files = st.file_uploader("Arraste os PDFs aqui", type="pdf", accept_multiple_files=True)

if uploaded_files:
    lista_excel = []
    
    with st.spinner("Analisando documentos..."):
        for f in uploaded_files:
            dados = extrair_dados_pdf(f.read(), f.name)
            
            for p in dados['Parcelamentos']:
                lista_excel.append({
                    "Arquivo": dados['Arquivo'],
                    "Município": dados['Município'],
                    "CNPJ": dados['CNPJ'],
                    "Processo / Negociação": p['Processo'], # Agora vem formatado
                    "Saldo Devedor": p['Saldo']
                })
    
    if lista_excel:
        df = pd.DataFrame(lista_excel)
        st.success(f"Extração concluída! {len(df)} registros encontrados.")
        st.dataframe(df, use_container_width=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Salva o DataFrame no Excel
            df.to_excel(writer, index=False, sheet_name='Saldos')
            
            # --- FORMATAÇÃO AVANÇADA DO EXCEL ---
            workbook = writer.book
            worksheet = writer.sheets['Saldos']
            
            # Formato de Texto para a coluna de Processo (Evita notação científica 1.02E+16)
            formato_texto = workbook.add_format({'num_format': '@'})
            
            # Aplica formatação nas colunas
            worksheet.set_column('A:A', 30) # Arquivo
            worksheet.set_column('B:B', 30) # Município
            worksheet.set_column('C:C', 20) # CNPJ
            worksheet.set_column('D:D', 25, formato_texto) # Processo como TEXTO
            worksheet.set_column('E:E', 15) # Valor
            
        output.seek(0)
        
        st.download_button(
            "⬇️ Baixar Excel (.xlsx)",
            data=output,
            file_name="Saldos_RFB_Fidedignos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nenhum dado encontrado.")
