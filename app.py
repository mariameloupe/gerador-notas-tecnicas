import streamlit as st
import pandas as pd
import os
from docx import Document
from datetime import datetime
from docx.shared import Inches, Pt

# Caminho fixo da logo
LOGO_PATH = "logo.png"  # Certifique-se de que a logo está salva neste caminho

# Função para carregar a planilha de pagamentos
def carregar_dados(uploaded_file):
    if uploaded_file is not None:
        xls = pd.ExcelFile(uploaded_file)
        
        df_inf_gerais = pd.read_excel(xls, "INF GERAIS", skiprows=5, dtype=str)
        df_inf_gerais.columns = df_inf_gerais.columns.str.strip()
        df_inf_gerais = df_inf_gerais[["ORDEM", "MUNICÍPIO", "PROJETO", "PROJETO DETALHADO", "TETO FEM", "STATUS PTM", "STATUS OBRA", "RESSALVA"]]
        
        df_resumo = pd.read_excel(xls, "RESUMO", skiprows=5, dtype=str)
        df_resumo.columns = df_resumo.columns.str.strip()
        df_resumo = df_resumo[["ORDEM", "MUNICÍPIO", "PROJETO", "STATUS PTM", "STATUS OBRA", "TETO FEM", "DATA ÚLTIMO PAGAMENTO", "REPASSE_VÁLIDO"]]
        
        df_inf_gerais["ORDEM"] = df_inf_gerais["ORDEM"].astype(str)
        df_resumo["ORDEM"] = df_resumo["ORDEM"].astype(str)
        df = pd.merge(df_inf_gerais, df_resumo, on="ORDEM", how="left")
        return df
    return None

# Layout no Streamlit
st.set_page_config(page_title="Gerador de Notas Técnicas", layout="wide")

st.markdown("<h3 style='text-align: right;'>SEPLAG-SEDRC</h3>", unsafe_allow_html=True)

st.title("📄 Gerador de Notas Técnicas")
st.markdown("Este aplicativo gera automaticamente notas técnicas a partir das planilhas de controle do FEM.")

if os.path.exists(LOGO_PATH):
    st.sidebar.image(LOGO_PATH, use_container_width=True)

st.sidebar.header("Upload do Arquivo")
uploaded_files = st.sidebar.file_uploader("Faça upload dos arquivos Excel", type=["xlsm"], accept_multiple_files=True)

if uploaded_files:
    dfs = []
    for uploaded_file in uploaded_files:
        df_temp = carregar_dados(uploaded_file)
        if df_temp is not None:
            dfs.append(df_temp)
    if dfs:
        df = pd.concat(dfs, ignore_index=True)
        df['ANO'] = df['ORDEM'].astype(str).str[:4]  # Extrai o ano dos 4 primeiros dígitos da coluna ORDEM
    
    df.columns = df.columns.str.strip()
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("📊 Visualização da Planilha")
        st.dataframe(df, height=300)
    
    coluna_municipio = 'MUNICÍPIO_x' if 'MUNICÍPIO_x' in df.columns else 'MUNICÍPIO_y'
    municipios_disponiveis = df[coluna_municipio].dropna().unique()
    municipio_selecionado = st.selectbox("🌍 Selecione um Município", municipios_disponiveis)
    df_filtrado = df[df[coluna_municipio] == municipio_selecionado]
    anos_disponiveis = df_filtrado['ANO'].unique()
    data_atualizacao = st.text_input("📅 Atualizado em (Digite a data e hora)", value=datetime.today().strftime('%d/%m/%Y %H:%M'))
    
    if not df_filtrado.empty and st.button("📝 Gerar Nota Técnica"):
        caminho_nota = "nota_tecnica.docx"
        doc = Document()
        doc.styles['Normal'].font.name = 'Arial'
        doc.styles['Normal'].font.size = Pt(12)
        
        if os.path.exists(LOGO_PATH):
            paragraph = doc.add_paragraph()
            run = paragraph.add_run()
            run.add_picture(LOGO_PATH, width=Inches(1.5))
            paragraph.alignment = 4  # Alinha à direita
        
        title = doc.add_paragraph()
        title_run = title.add_run('NOTA TÉCNICA')
        title_run.bold = True
        title.alignment = 1  # Centraliza o texto
        municipio_paragraph = doc.add_paragraph()
        municipio_run = municipio_paragraph.add_run(f"Município de {municipio_selecionado}")
        municipio_run.bold = True
        municipio_run.font.color.rgb = None
        doc.add_paragraph(f"Atualizado em: {data_atualizacao}")
        
        for ano in sorted(anos_disponiveis):
            df_ano = df_filtrado[df_filtrado['ANO'] == ano]
            doc.add_paragraph(f"Ano do Registro: {ano}", style='Heading 3')
            
            colunas_tabela = ["PROJETO DETALHADO", 'TETO FEM_x' if 'TETO FEM_x' in df_ano.columns else 'TETO FEM_y', "REPASSE_VÁLIDO", "DATA ÚLTIMO PAGAMENTO", 'STATUS OBRA_x' if 'STATUS OBRA_x' in df_ano.columns else 'STATUS OBRA_y']
            
            tabela = doc.add_table(rows=1, cols=len(colunas_tabela))
            tabela.style = 'Table Grid'
            
            hdr_cells = tabela.rows[0].cells
            for i, coluna in enumerate(colunas_tabela):
                hdr_cells[i].text = coluna
            
            for _, linha in df_ano.iterrows():
                row_cells = tabela.add_row().cells
                for i, coluna in enumerate(colunas_tabela):
                    valor = linha.get(coluna, '0')
                    if coluna in ['TETO FEM_x', 'TETO FEM_y', 'REPASSE_VÁLIDO']:
                        try:
                            valor = f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                        except ValueError:
                            valor = 'R$ 0,00'
                    if coluna == 'DATA ÚLTIMO PAGAMENTO':
                        valor = str(valor)
                        try:
                            valor = datetime.strptime(valor, '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y')
                        except ValueError:
                            pass
                    row_cells[i].text = str(valor)
        caminho_nota = "nota_tecnica.docx"
        doc = Document()
        doc.styles['Normal'].font.name = 'Arial'
        doc.styles['Normal'].font.size = Pt(12)
        
        if os.path.exists(LOGO_PATH):
            paragraph = doc.add_paragraph()
            run = paragraph.add_run()
            run.add_picture(LOGO_PATH, width=Inches(1.5))
            paragraph.alignment = 4  # Alinha à direita
        
        title = doc.add_paragraph()
        title_run = title.add_run('NOTA TÉCNICA')
        title_run.bold = True
        title.alignment = 1  # Centraliza o texto
        municipio_paragraph = doc.add_paragraph()
        municipio_run = municipio_paragraph.add_run(f"Município de {municipio_selecionado}")
        
        ano = df_filtrado.iloc[0]['ORDEM'][:4] if not df_filtrado.empty else "Desconhecido"
        doc.add_paragraph(f"Ano do Registro: {ano}", style='Heading 3')
        municipio_run.bold = True
        municipio_run.font.color.rgb = None
        doc.add_paragraph(f"Atualizado em: {data_atualizacao}")
        colunas_tabela = ["PROJETO DETALHADO", 'TETO FEM_x' if 'TETO FEM_x' in df_filtrado.columns else 'TETO FEM_y', "REPASSE_VÁLIDO", "DATA ÚLTIMO PAGAMENTO", 'STATUS OBRA_x' if 'STATUS OBRA_x' in df_filtrado.columns else 'STATUS OBRA_y']
        
        tabela = doc.add_table(rows=1, cols=len(colunas_tabela))
        tabela.style = 'Table Grid'
        
        hdr_cells = tabela.rows[0].cells
        for i, coluna in enumerate(colunas_tabela):
            hdr_cells[i].text = coluna
        
        for _, linha in df_filtrado.iterrows():
            row_cells = tabela.add_row().cells
            for i, coluna in enumerate(colunas_tabela):
                valor = linha.get(coluna, '0')
                if coluna in ['TETO FEM_x', 'TETO FEM_y', 'REPASSE_VÁLIDO']:
                    try:
                        valor = f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    except ValueError:
                        valor = 'R$ 0,00'
                if coluna == 'DATA ÚLTIMO PAGAMENTO':
                    valor = str(valor)  # Garante que o valor seja tratado como string antes da conversão
                    try:
                        valor = datetime.strptime(valor, '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y')
                    except ValueError:
                        pass
                row_cells[i].text = str(valor)
        
        doc.save(caminho_nota)
        st.success("✅ Nota Técnica gerada com sucesso!")
        with open(caminho_nota, "rb") as file:
            st.download_button("📥 Baixar Nota Técnica (DOCX)", file, file_name="nota_tecnica.docx")