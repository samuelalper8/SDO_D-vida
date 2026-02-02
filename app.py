import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import RGBColor
import zipfile
from datetime import datetime

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Ofícios DFF", layout="wide")

# --- FUNÇÕES DE EXTRAÇÃO ---
def extrair_dados_pdf(pdf_stream, filename):
    doc = fitz.open(stream=pdf_stream, filetype="pdf")
    texto_completo = ""
    for pagina in doc:
        texto_completo += pagina.get_text()
    
    # 1. Identificar Município e CNPJ
    municipio_match = re.search(r"MUNICIPIO DE\s+(.*)", texto_completo)
    municipio = municipio_match.group(1).strip() if municipio_match else "DESCONHECIDO"
    
    cnpj_match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_completo)
    cnpj = cnpj_match.group(1) if cnpj_match else ""

    # 2. Tentar extrair linhas da tabela de parcelamentos
    # A estratégia é buscar linhas que comecem com o CNPJ, pois é o padrão da tabela RFB
    parcelamentos = []
    
    # Divide por linhas e remove vazias
    linhas = [l.strip() for l in texto_completo.split('\n') if l.strip()]
    
    capturando_tabela = False
    for i, linha in enumerate(linhas):
        # Início da tabela geralmente tem cabeçalhos, mas o dado começa com o CNPJ
        if cnpj and linha.startswith(cnpj):
            # Tenta capturar: CNPJ | Processo | ... | Valor
            # Como o texto pode quebrar de formas variadas, usamos regex na linha ou nas próximas
            # Padrão esperado: CNPJ (espaço) PROCESSO (espaço) ... VALOR
            # Exemplo Bela Vista: 01.005.917/0001-41 10120729679201251 PARC ESPECIAL... 39.234,24
            
            # Regex busca o processo (números longos) e o valor (final da linha ou próximo numérico)
            match_processo = re.search(r"\s+(\d{9,})", linha) # Busca sequencia longa de numeros
            match_valor = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linha)
            
            if match_processo and match_valor:
                parcelamentos.append({
                    "Processo": match_processo.group(1),
                    "Saldo": match_valor.group(1)
                })
            elif match_processo: 
                # Se achou o processo mas o valor quebrou a linha, tenta pegar na próxima
                # (Lógica simplificada, pode precisar de ajuste fino dependendo do PDF real)
                if i + 1 < len(linhas):
                    prox_valor = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", linhas[i+1])
                    if prox_valor:
                        parcelamentos.append({
                            "Processo": match_processo.group(1),
                            "Saldo": prox_valor.group(1)
                        })

    # 3. Fallback: Se não achou linhas individuais, pega o Total Geral
    if not parcelamentos:
        saldo_total = re.search(r"SALDO DEVEDOR TOTAL\s+([\d.,]+)", texto_completo)
        valor_total = saldo_total.group(1) if saldo_total else "0,00"
        processo_match = re.search(r"No Processo/Dossiê\s+([\d./-]+)", texto_completo)
        proc_geral = processo_match.group(1) if processo_match else "Consolidado"
        
        parcelamentos.append({
            "Processo": proc_geral,
            "Saldo": valor_total
        })

    return {
        "Arquivo": filename,
        "Município": municipio,
        "CNPJ": cnpj,
        "Parcelamentos": parcelamentos # Lista de dicts
    }

# --- FUNÇÃO DE GERAÇÃO DO WORD ---
def gerar_docx(dados, nome_prefeito):
    doc = Document()
    
    # Configurar fonte base
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    # Data
    data_atual = datetime.now().strftime("%d de %B de %Y")
    # Tenta traduzir mês tosco (opcional, ou use locale)
    meses = {
        "January": "janeiro", "February": "fevereiro", "March": "março", "April": "abril",
        "May": "maio", "June": "junho", "July": "julho", "August": "agosto",
        "September": "setembro", "October": "outubro", "November": "novembro", "December": "dezembro"
    }
    for ing, pt in meses.items():
        data_atual = data_atual.replace(ing, pt)
        
    p_data = doc.add_paragraph(f"Goiânia, {data_atual}.")
    p_data.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_paragraph() # Espaço

    # Cabeçalho Destinatário
    p_dest = doc.add_paragraph()
    p_dest.add_run("EXCELENTÍSSIMO SENHOR\n").bold = True
    p_dest.add_run(f"{nome_prefeito.upper()}\n").bold = True
    p_dest.add_run(f"PREFEITO MUNICIPAL DE {dados['Município'].upper()} – GO") # Assumindo GO pelos arquivos

    doc.add_paragraph(f"Assunto: Ficam apresentados os valores e a documentação comprobatória dos saldos de débitos existentes em 31 de dezembro de 2025, destinados à composição do Balanço Patrimonial.")

    p_intro = doc.add_paragraph("Senhor Prefeito,")
    
    texto_corpo = (
        "Ao tempo em que lhe cumprimento, na qualidade de assessoria do Município para assuntos relacionados a atos de pessoal e ao fisco federal, "
        "no âmbito das ações de conformidade administrativa, venho, por meio do presente, apresentar os valores e a documentação requisitados por esta assessoria especializada, "
        "referentes aos saldos de débitos destinados à composição do Balanço Patrimonial."
    )
    p_corpo = doc.add_paragraph(texto_corpo)
    p_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_paragraph("Nesse contexto, discriminam-se abaixo o órgão de origem, o número do processo e os respectivos valores apurados em 31/12/2025, estando anexa a documentação comprobatória pertinente:")

    # --- TABELA ---
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Órgão'
    hdr_cells[1].text = 'Processo / Documento'
    hdr_cells[2].text = 'Saldo em 31/12/2025'
    
    # Formata cabeçalho negrito
    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True

    # Preenche dados (Receita Federal)
    for p in dados['Parcelamentos']:
        row_cells = table.add_row().cells
        row_cells[0].text = "Receita Federal do Brasil"
        row_cells[1].text = str(p['Processo'])
        row_cells[2].text = str(p['Saldo'])
    
    # Linha PGFN (Vazia conforme modelo)
    row_pgfn = table.add_row().cells
    row_pgfn[0].text = "Procuradoria da Fazenda Nacional"
    row_pgfn[1].text = "-"
    row_pgfn[2].text = "-"

    doc.add_paragraph()
    
    texto_final = (
        "Solicita-se, por oportuno, que a referida documentação seja encaminhada ao setor contábil, a fim de que sejam adotadas as providências e registros contábeis cabíveis, conforme as normas aplicáveis.\n"
        "Esta consultoria agradece a confiança depositada e permanece à disposição para quaisquer esclarecimentos adicionais que se fizerem necessários.\n"
        "Sem mais para o momento, reitero protestos de elevada estima e consideração."
    )
    p_final = doc.add_paragraph(texto_final)
    p_final.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_paragraph("Atenciosamente,")
    doc.add_paragraph()
    doc.add_paragraph()

    # --- ASSINATURAS ---
    # Usando tabela invisível para alinhar assinaturas se necessário, ou parágrafos centralizados
    
    p_sig1 = doc.add_paragraph()
    p_sig1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p_sig1.add_run("Rubens Pires Malaquias\n")
    run1.bold = True
    p_sig1.add_run("Diretor Técnico e Consultor junto ao Fisco Federal\nCRA/GO 6-007-48")
    
    doc.add_paragraph()

    p_sig2 = doc.add_paragraph()
    p_sig2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p_sig2.add_run("Glayzer Antônio Gomes da Silva\n")
    run2.bold = True
    p_sig2.add_run("Advogado Especialista em Direito Público\nOAB/GO 28.315")

    doc.add_paragraph()

    p_sig3 = doc.add_paragraph()
    p_sig3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run3 = p_sig3.add_run("Samuel Almeida\n")
    run3.bold = True
    p_sig3.add_run("Líder de Equipe – Fisco Federal GO")

    # Salva em memória
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFACE PRINCIPAL ---

st.title("📄 Gerador Automático de Ofícios de Dívida (DFF)")
st.markdown("Extrai saldos de PDFs da RFB e gera os documentos `.docx` para Balanço Patrimonial.")

uploaded_files = st.file_uploader("Selecione os PDFs da Receita Federal", type="pdf", accept_multiple_files=True)

if uploaded_files:
    # 1. Extração Inicial
    lista_dados = []
    with st.spinner("Lendo arquivos..."):
        for f in uploaded_files:
            dados = extrair_dados_pdf(f.read(), f.name)
            lista_dados.append(dados)
    
    # 2. Preparação para Edição (Usuário precisa inserir Prefeitos)
    df_preparacao = []
    for d in lista_dados:
        df_preparacao.append({
            "Arquivo": d["Arquivo"],
            "Município": d["Município"],
            "Prefeito (Preencher)": ""  # Coluna vazia para input
        })
    
    st.info("👇 **Atenção:** Os PDFs não contêm o nome do prefeito. Preencha a coluna 'Prefeito' abaixo antes de gerar os arquivos.")
    
    df_editado = st.data_editor(
        pd.DataFrame(df_preparacao),
        column_config={
            "Prefeito (Preencher)": st.column_config.TextColumn("Nome do Prefeito", help="Ex: Fulano de Tal", required=True)
        },
        use_container_width=True,
        hide_index=True,
        num_rows="fixed"
    )

    # 3. Botão de Geração
    if st.button("Gerar Ofícios (.docx)"):
        zip_buffer = BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            progress_bar = st.progress(0)
            
            for i, row in df_editado.iterrows():
                # Encontra os dados originais correspondentes (incluindo parcelas)
                dados_originais = next(d for d in lista_dados if d["Arquivo"] == row["Arquivo"])
                nome_prefeito = row["Prefeito (Preencher)"]
                
                if not nome_prefeito:
                    nome_prefeito = "A/C GESTOR MUNICIPAL" # Fallback se deixar vazio
                
                # Gera o DOCX
                docx_buffer = gerar_docx(dados_originais, nome_prefeito)
                
                # Nome do arquivo de saída
                nome_municipio_limpo = row["Município"].replace(" ", "_").upper()
                nome_arquivo = f"OFICIO_DIVIDA_{nome_municipio_limpo}.docx"
                
                zf.writestr(nome_arquivo, docx_buffer.getvalue())
                progress_bar.progress((i + 1) / len(df_editado))
        
        zip_buffer.seek(0)
        
        st.success("✅ Processamento concluído!")
        st.download_button(
            label="⬇️ Baixar Todos os Ofícios (ZIP)",
            data=zip_buffer,
            file_name="Oficios_Divida_2025.zip",
            mime="application/zip"
        )
