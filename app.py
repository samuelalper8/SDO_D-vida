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
        
        try:
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    
                    for table in tables:
                        if not table: continue
                        
                        df_temp = pd.DataFrame(table)
                        
                        # Verifica se é uma tabela válida de débitos
                        # Converte a primeira linha para string para buscar palavras-chave
                        if df_temp.empty: continue
                        
                        header_row = str(df_temp.iloc[0].values).upper()
                        
                        # Palavras-chave típicas da tabela RFB
                        keywords = ["PROCESSO", "MODALIDADE", "SALDO DEVEDOR", "CNPJ VINCULADO", "SISTEMA"]
                        if any(k in header_row for k in keywords):
                            
                            # Define a 1ª linha como cabeçalho
                            df_temp.columns = df_temp.iloc[0]
                            df_temp = df_temp[1:]
                            
                            # Adiciona coluna de origem
                            df_temp["Arquivo Origem"] = pdf_file.name
                            
                            # Concatena
                            df_consolidado = pd.concat([df_consolidado, df_temp], ignore_index=True)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo {pdf_file.name}: {e}")
        
        barra.progress((i + 1) / total_arquivos)
    
    progresso_texto.empty()
    barra.empty()
    
    return df_consolidado

# --- 2. FUNÇÃO DE LIMPEZA E ORGANIZAÇÃO (CORRIGIDA) ---
def organizar_dados(df_bruto):
    if df_bruto.empty:
        return pd.DataFrame()

    df = df_bruto.copy()

    # A. Normalização de Nomes de Colunas (Remove espaços e quebras)
    # Garante que os nomes sejam strings
    df.columns = [str(c).replace('\n', ' ').strip().upper() if c is not None else f"COL_{i}" for i, c in enumerate(df.columns)]
    
    # B. Identificação das Colunas (Mapeamento Seguro)
    col_map = {}
    for col in df.columns:
        if "PROCESSO" in col: col_map[col] = "Processo"
        elif "CNPJ" in col: col_map[col] = "CNPJ"
        elif "MODALIDADE" in col: col_map[col] = "Modalidade"
        elif "SISTEMA" in col: col_map[col] = "Sistema"
        elif "SALDO" in col and "DEVEDOR" in col: col_map[col] = "Valor Original" # Mais específico
        elif "ARQUIVO" in col: col_map[col] = "Arquivo"

    df = df.rename(columns=col_map)

    # --- FIX CRÍTICO: DEDUPLICAÇÃO DE COLUNAS ---
    # Se houver duas colunas com o mesmo nome (ex: "Processo"), o Pandas falha no loop seguinte.
    # Esta rotina renomeia duplicatas para Processo, Processo_1, etc.
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique(): 
        cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    # ---------------------------------------------

    # C. Filtragem de Lixo
    if 'CNPJ' in df.columns:
        # Garante que é string antes de usar .str
        df = df[~df['CNPJ'].astype(str).str.contains("CNPJ", case=False, na=False)]
    
    # Remove linhas de Totais
    mask_total = df.astype(str).apply(lambda x: x.str.contains('TOTAL', case=False)).any(axis=1)
    df = df[~mask_total]

    # D. Limpeza de Texto (\n -> espaço)
    for col in df.columns:
        # Verifica se a coluna existe e se o tipo é objeto (texto)
        if col in df.columns and df[col].dtype == object:
            df[col] = df[col].astype(str).str.replace('\n', ' ').str.strip()

    # E. Conversão Numérica
    if 'Valor Original' in df.columns:
        def converter_valor(x):
            try:
                # Remove espaços, remove pontos de milhar, troca vírgula decimal por ponto
                limpo = str(x).replace(' ', '').replace('.', '').replace(',', '.')
                # Regex para garantir que sobrou um número válido
                if re.match(r'^-?\d+(\.\d+)?$', limpo):
                    return float(limpo)
                return 0.0
            except:
                return 0.0

        df['Valor Numérico'] = df['Valor Original'].apply(converter_valor)

    # F. Extração Inteligente do Município
    def extrair_municipio(nome_arq):
        try:
            partes = str(nome_arq).split('-')
            if len(partes) >= 2:
                return partes[1].strip().upper()
            return nome_arq
        except:
            return "DESCONHECIDO"

    if 'Arquivo' in df.columns:
        df.insert(0, 'Município', df['Arquivo'].apply(extrair_municipio))

    # G. Seleção Final (Garante que as colunas existam)
    cols_desejadas = ['Município', 'CNPJ', 'Processo', 'Modalidade', 'Sistema', 'Valor Numérico', 'Valor Original', 'Arquivo']
    cols_finais = [c for c in cols_desejadas if c in df.columns]
    
    return df[cols_finais]

# --- 3. INTERFACE PRINCIPAL ---
st.title("🗂️ Extrator e Organizador RFB")
st.markdown("Extrai tabelas de PDFs, gera uma aba bruta (auditoria) e uma aba limpa (gestão).")

uploaded_files = st.file_uploader("Arraste os PDFs aqui", type="pdf", accept_multiple_files=True)

if uploaded_files:
    # 1. Extração
    with st.spinner("Processando arquivos..."):
        df_bruto = extrair_tabelas_brutas(uploaded_files)
    
    if not df_bruto.empty:
        # 2. Organização
        try:
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
                
            # 4. Botão de Download
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Aba Organizada
                df_limpo.to_excel(writer, index=False, sheet_name='Organizado')
                wb = writer.book
                ws_org = writer.sheets['Organizado']
                
                # Formatação Moeda
                fmt_money = wb.add_format({'num_format': '#,##0.00'})
                
                # Aplica formatação se a coluna existir
                if "Valor Numérico" in df_limpo.columns:
                    col_idx_valor = df_limpo.columns.get_loc("Valor Numérico")
                    ws_org.set_column(col_idx_valor, col_idx_valor, 18, fmt_money)
                
                ws_org.set_column(0, 0, 25) # Largura Município
                
                # Aba Bruta
                df_bruto.to_excel(writer, index=False, sheet_name='Bruto_Original')
                
            output.seek(0)
            
            st.download_button(
                label="⬇️ Baixar Excel (Bruto + Limpo)",
                data=output,
                file_name="Relatorio_Dividas_RFB_Completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erro durante a organização dos dados: {e}")
            st.write("Dados brutos recuperados (para debug):")
            st.dataframe(df_bruto)
            
    else:
        st.warning("Nenhuma tabela encontrada. Verifique se os arquivos são PDFs pesquisáveis (não escaneados).")
