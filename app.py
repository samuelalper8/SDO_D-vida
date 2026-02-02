import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Extrator de Tabelas RFB", layout="wide", page_icon="🗃️")

def extrair_tabelas_exatas(pdf_file, nome_arquivo):
    tabelas_encontradas = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Extrai todas as tabelas da página
            tables = page.extract_tables()
            
            for table in tables:
                # table é uma lista de listas: [['Header1', 'Header2'], ['Val1', 'Val2']]
                if not table: continue
                
                # Vamos converter para DataFrame para analisar
                df_temp = pd.DataFrame(table)
                
                # Limpeza: Remove quebras de linha (\n) dentro das células que sujam o Excel
                df_temp = df_temp.replace(r'\n', ' ', regex=True)
                
                # CRITÉRIO DE SEGURANÇA:
                # Verificamos se é a tabela de débitos procurando palavras-chave do cabeçalho
                # Convertemos a primeira linha para string para buscar as palavras
                header_row = str(table[0]).upper()
                
                if "PROCESSO" in header_row or "MODALIDADE" in header_row or "SALDO DEVEDOR" in header_row:
                    # Se achou, adiciona metadados (Arquivo/Página)
                    df_temp.insert(0, "Arquivo Origem", nome_arquivo)
                    tabelas_encontradas.append(df_temp)

    return tabelas_encontradas

# --- INTERFACE ---
st.title("🗃️ Extrator de Tabelas Integrais (RFB)")
st.markdown("Extrai a estrutura exata da tabela do PDF (todas as colunas: Modalidade, Sistema, Valor).")

uploaded_files = st.file_uploader("Arraste os PDFs (Barro Alto, Pilar, etc)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    df_consolidado = pd.DataFrame()
    
    with st.spinner("Mapeando grades das tabelas..."):
        for f in uploaded_files:
            # pdfplumber exige o arquivo, não apenas bytes. O Streamlit file buffer funciona.
            tabelas = extrair_tabelas_exatas(f, f.name)
            
            for df_tabela in tabelas:
                # Define a primeira linha como cabeçalho
                novo_header = df_tabela.iloc[0]
                df_tabela = df_tabela[1:] # Pega os dados da linha 1 em diante
                df_tabela.columns = novo_header # Renomeia colunas
                
                # Adiciona coluna do nome do arquivo (que ficou sem nome após o header reset)
                # O índice 0 do novo_header geralmente é 'Arquivo Origem' que inserimos, ou lixo.
                # Vamos garantir que a coluna de arquivo exista
                if "Arquivo Origem" not in df_tabela.columns:
                     df_tabela["Arquivo Origem"] = f.name
                
                df_consolidado = pd.concat([df_consolidado, df_tabela], ignore_index=True)

    if not df_consolidado.empty:
        st.success(f"✅ Extração completa! Tabelas recuperadas com sucesso.")
        
        # Mostra prévia
        st.dataframe(df_consolidado, use_container_width=True)
        
        # Botão Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_consolidado.to_excel(writer, index=False, sheet_name='Tabelas_Completas')
            
            # Ajuste fino de colunas
            worksheet = writer.sheets['Tabelas_Completas']
            worksheet.set_column('A:A', 30) # Arquivo
            worksheet.set_column('B:Z', 20) # Outras colunas
            
        output.seek(0)
        st.download_button("⬇️ Baixar Tabela Original (.xlsx)", output, "Tabelas_RFB_Exatas.xlsx")
        
    else:
        st.warning("Nenhuma tabela padrão encontrada. Verifique se o PDF é imagem (scaneado) ou texto selecionável.")
