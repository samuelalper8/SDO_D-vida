import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Extrator Detalhado RFB", layout="wide", page_icon="📑")

# --- FUNÇÃO DE EXTRAÇÃO (Lógica Linha a Linha) ---
def extrair_dados_pdf(pdf_bytes, nome_arquivo):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto_completo = ""
    for pagina in doc:
        texto_completo += pagina.get_text()
    
    # 1. Identificação do Município e CNPJ (Cabeçalho)
    municipio_match = re.search(r"MUNICIPIO DE\s+(.*)", texto_completo)
    municipio = municipio_match.group(1).strip().upper() if municipio_match else "DESCONHECIDO"
    
    # Busca o CNPJ padrão no cabeçalho
    cnpj_header_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_completo)
    cnpj_header = cnpj_header_match.group(1) if cnpj_header_match else ""

    parcelamentos = []
    
    # Quebra o texto em linhas para analisar uma a uma
    linhas = [l.strip() for l in texto_completo.split('\n') if l.strip()]
    
    # 2. Varredura da Tabela
    # A lógica aqui é: Toda linha de parcelamento na RFB contém o CNPJ do devedor.
    # Vamos caçar todas as ocorrências do CNPJ e extrair os dados ao redor.
    
    for i, linha in enumerate(linhas):
        # Verifica se o CNPJ (ou parte dele) está na linha
        # Usamos apenas os números do CNPJ para evitar problemas com pontuação diferente
        cnpj_limpo = cnpj_header.replace('.', '').replace('/', '').replace('-', '')
        linha_limpa = linha.replace('.', '').replace('/', '').replace('-', '')
        
        # Se a linha contém o CNPJ e parece ser uma linha de dados (tem valor monetário)
        if cnpj_header and (cnpj_header in linha) and re.search(r"\d+,\d{2}", linha):
            
            # --- Extração do Processo/Negociação ---
            # Removemos o CNPJ da linha para não confundir
            linha_sem_cnpj = linha.replace(cnpj_header, "")
            
            # Buscamos sequências numéricas longas (Processos geralmente tem > 7 dígitos)
            # Ex: 10120729679201251 ou 620240890
            match_processos = re.findall(r"\b\d{7,}\b", linha_sem_cnpj)
            
            # O primeiro número longo que sobrar geralmente é o processo
            processo = match_processos[0] if match_processos else "N/D"

            # --- Extração do Valor ---
            # Busca formato monetário brasileiro: X.XXX,XX ou apenas XXX,XX
            match_valor = re.findall(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
            
            # O valor do saldo devedor é geralmente o ÚLTIMO valor monetário da linha
            valor = match_valor[-1] if match_valor else "0,00"
            
            # Identificação do Sistema (SIPADE / SICOB / PARCWEB) - Opcional, ajuda a validar
            sistema = "Outros"
            if "SIPADE" in linha: sistema = "SIPADE"
            elif "SICOB" in linha or "PARCWEB" in linha: sistema = "PARCWEB/SICOB"
            
            parcelamentos.append({
                "Processo/Negociação": processo,
                "Sistema": sistema,
                "Saldo": valor,
                "Linha Original": linha # Debug se precisar conferir
            })

    # 3. Fallback (Caso de Amaralina ou tabelas vazias)
    if not parcelamentos:
        saldo_total = re.search(r"SALDO DEVEDOR TOTAL\s+([\d.,]+)", texto_completo)
        valor_total = saldo_total.group(1) if saldo_total else "0,00"
        
        # Se o valor for 0,00, adiciona uma linha indicando que não há dívida
        obs = "Sem parcelamentos listados"
        if valor_total == "0,00":
            parcelamentos.append({
                "Processo/Negociação": "-",
                "Sistema": "-",
                "Saldo": "0,00",
                "Linha Original": obs
            })
        else:
            # Se tem saldo total mas não achou linhas, pega o processo do cabeçalho
            proc_header = re.search(r"No Processo/Dossiê\s+([\d./-]+)", texto_completo)
            proc_ref = proc_header.group(1) if proc_header else "Consolidado"
            
            parcelamentos.append({
                "Processo/Negociação": proc_ref,
                "Sistema": "Consolidado",
                "Saldo": valor_total,
                "Linha Original": "Extração pelo Total Geral"
            })

    return {
        "Arquivo": nome_arquivo,
        "Município": municipio, 
        "CNPJ": cnpj_header, 
        "Parcelamentos": parcelamentos
    }

# --- INTERFACE ---
st.title("📊 Extrator Detalhado de Parcelamentos RFB")
st.markdown("""
Esta ferramenta extrai **cada linha** da tabela de parcelamentos individualmente.
Se houver múltiplos processos para o mesmo município, cada um aparecerá em uma linha na planilha.
""")

uploaded_files = st.file_uploader("Arraste seus PDFs aqui", type="pdf", accept_multiple_files=True)

if uploaded_files:
    lista_para_excel = []
    
    with st.spinner("Lendo cada linha dos arquivos..."):
        for f in uploaded_files:
            dados = extrair_dados_pdf(f.read(), f.name)
            
            # "Explode" a lista de parcelamentos para criar várias linhas no Excel
            for p in dados['Parcelamentos']:
                lista_para_excel.append({
                    "Arquivo Origem": dados['Arquivo'],
                    "Município": dados['Município'],
                    "CNPJ": dados['CNPJ'],
                    "Processo / Negociação": p['Processo/Negociação'],
                    "Sistema": p['Sistema'],
                    "Saldo Devedor (R$)": p['Saldo']
                })
    
    if lista_para_excel:
        df = pd.DataFrame(lista_para_excel)
        
        st.success(f"✅ Processamento Concluído! Extraídas {len(df)} linhas de débitos.")
        st.dataframe(df, use_container_width=True)
        
        # Botão Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Analitico_Dividas')
            
            # Ajuste de largura das colunas (estética)
            workbook = writer.book
            worksheet = writer.sheets['Analitico_Dividas']
            format_currency = workbook.add_format({'num_format': '#,##0.00'})
            worksheet.set_column('F:F', 15, format_currency) # Coluna de valor
            worksheet.set_column('B:B', 30) # Município
            worksheet.set_column('D:D', 25) # Processo
            
        output.seek(0)
        
        st.download_button(
            label="⬇️ Baixar Planilha Detalhada (.xlsx)",
            data=output,
            file_name="Relatorio_Detalhado_Parcelamentos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nenhuma informação encontrada nos arquivos.")
