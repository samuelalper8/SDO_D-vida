import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Extrator RFB (Bruto & Limpo)", layout="wide", page_icon="🗂️")

# --- 1. FUNÇÃO DE EXTRAÇÃO (Motor pdfplumber) ---
def extrair_tabelas_brutas(uploaded_files):
    df_consolidado = pd.DataFrame()
    
    # Barra de progresso para visualização
    progresso_texto = st.empty()
    barra = st.progress(0)
    
    total_arquivos = len(uploaded_files)
    
    for i, pdf_file in enumerate(uploaded_files):
        progresso_texto.text(f"Lendo arquivo {i+1}/{total_arquivos}: {pdf_file.name}")
        
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    if not table: continue
                    
                    df_temp = pd.DataFrame(table)
                    
                    # Verifica se é uma tabela válida de débitos (procura palavras-chave na 1ª linha)
                    header_row = str(df_temp.iloc[0].values).upper()
                    
                    # Palavras-chave típicas da tabela RFB
                    keywords = ["PROCESSO", "MODALIDADE", "SALDO DEVEDOR", "CNPJ VINCULADO"]
                    if any(k in header_row for k in keywords):
                        
                        # Define a 1ª linha como cabeçalho
                        df_temp.columns = df_temp.iloc[0]
                        df_temp = df_temp[1:]
                        
                        # Adiciona coluna de origem
                        df_temp["Arquivo Origem"] = pdf_file.name
                        
                        # Concatena
                        df_consolidado = pd.concat([df_consolidado, df_temp], ignore_index=True)
        
        barra.progress((i + 1) / total_arquivos)
    
    progresso_texto.empty()
    barra.empty()
    
    return df_consolidado

# --- 2. FUNÇÃO DE LIMPEZA E ORGANIZAÇÃO ---
def organizar_dados(df_bruto):
    if df_bruto.empty:
        return pd.DataFrame()

    df = df_bruto.copy()

    # A. Normalização de Nomes de Colunas (Remove espaços e quebras)
    df.columns = [str(c).replace('\n', ' ').strip().upper() for c in df.columns]
    
    # B. Identificação das Colunas (Para lidar com variações)
    col_map = {}
    for col in df.columns:
        if "PROCESSO" in col: col_map[col] = "Processo"
        elif "CNPJ" in col: col_map[col] = "CNPJ"
        elif "MODALIDADE" in col: col_map[col] = "Modalidade"
        elif "SISTEMA" in col: col_map[col] = "Sistema"
        elif "SALDO" in col: col_map[col] = "Valor Original"
        elif "ARQUIVO" in col: col_map[col] = "Arquivo"

    df = df.rename(columns=col_map)

    # C. Filtragem de Lixo (Linhas repetidas de cabeçalho e Totais)
    # Remove linhas onde o conteúdo da coluna CNPJ é igual ao título da coluna
    if 'CNPJ' in df.columns:
        df = df[~df['CNPJ'].astype(str).str.contains("CNPJ", case=False, na=False)]
    
    # Remove linhas de Totais (geralmente tem "TOTAL" em alguma coluna)
    # Convertemos tudo para string e procuramos "TOTAL"
    mask_total = df.astype(str).apply(lambda x: x.str.contains('TOTAL', case=False)).any(axis=1)
    df = df[~mask_total]

    # D. Limpeza de Texto (\n -> espaço)
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.replace('\n', ' ').str.strip()

    # E. Conversão Numérica (Brasil R$ -> Float)
    if 'Valor Original' in df.columns:
        # Remove pontos de milhar e troca vírgula por ponto
        df['Valor Numérico'] = df['Valor Original'].apply(
            lambda x: float(x.replace('.', '').replace(',', '.')) if re.match(r'[\d.,]+', x) else 0.0
        )

    # F. Extração Inteligente do Município (Baseado no nome do arquivo padrão "GO - CIDADE - ...")
    # Tenta pegar o texto entre o primeiro e o segundo hífen
    def extrair_municipio(nome_arq):
        partes = nome_arq.split('-')
        if len(partes) >= 2:
            return partes[1].strip().upper()
        return nome_arq # Fallback

    if 'Arquivo' in df.columns:
        df.insert(0, 'Município', df['Arquivo'].apply(extrair_municipio))

    # G. Seleção e Ordem Final das Colunas
    cols_desejadas = ['Município', 'CNPJ', 'Processo', 'Modalidade', 'Sistema', 'Valor Numérico', 'Valor Original', 'Arquivo']
    cols_finais = [c for c in cols_desejadas if c in df.columns]
    
    return df[cols_finais]

# --- 3. INTERFACE PRINCIPAL ---
st.title("🗂️ Extrator e Organizador RFB")
st.markdown("Extrai tabelas de PDFs, gera uma aba bruta (auditoria) e uma aba limpa (gestão).")

uploaded_files = st.file_uploader("Arraste os PDFs aqui", type="pdf", accept_multiple_files=True)

if uploaded_files:
    # 1. Extração
    df_bruto = extrair_tabelas_brutas(uploaded_files)
    
    if not df_bruto.empty:
        # 2. Organização
        df_limpo = organizar_dados(df_bruto)
        
        st.success("Processamento concluído!")
        
        # 3. Visualização em Abas
        tab1, tab2 = st.tabs(["📂 Dados Organizados (Limpo)", "🔍 Dados Originais (Bruto)"])
        
        with tab1:
            st.dataframe(df_limpo, use_container_width=True)
            if 'Valor Numérico' in df_limpo.columns:
                total = df_limpo['Valor Numérico'].sum()
                st.metric("Total Consolidado da Seleção", f"R$ {total:,.2f}")
        
        with tab2:
            st.dataframe(df_bruto, use_container_width=True)
            
        # 4. Botão de Download (Excel Multi-Aba)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Aba Organizada
            df_limpo.to_excel(writer, index=False, sheet_name='Organizado')
            wb = writer.book
            ws_org = writer.sheets['Organizado']
            
            # Formatação Moeda
            fmt_money = wb.add_format({'num_format': '#,##0.00'})
            col_idx_valor = df_limpo.columns.get_loc("Valor Numérico")
            ws_org.set_column(col_idx_valor, col_idx_valor, 18, fmt_money)
            ws_org.set_column(0, 0, 25) # Largura Município
            ws_org.set_column(2, 3, 20) # Largura Processo/Modalidade

            # Aba Bruta
            df_bruto.to_excel(writer, index=False, sheet_name='Bruto_Original')
            
        output.seek(0)
        
        st.download_button(
            label="⬇️ Baixar Excel (Bruto + Limpo)",
            data=output,
            file_name="Relatorio_Dividas_RFB_Completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nenhuma tabela encontrada. Verifique se os arquivos são PDFs pesquisáveis (não escaneados).")
