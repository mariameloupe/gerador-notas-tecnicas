import streamlit as st
import pandas as pd
import os
from docx import Document
from datetime import datetime
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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

# Função para adicionar parágrafos formatados
def adicionar_paragrafo_formatado(doc, texto, fonte='Arial', tamanho=12, negrito=False, cor=None):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(texto)
    run.font.name = fonte
    run.font.size = Pt(tamanho)
    run.bold = negrito
    if cor:
        run.font.color.rgb = cor
    return paragraph

# Função para formatar valores como moeda
def formatar_moeda(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except ValueError:
        return 'R$ 0,00'

# Função para aplicar fonte Arial tamanho 8 em todas as células da tabela
def formatar_tabela(tabela):
    for row in tabela.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(8)

# Função para definir a cor de fundo de uma célula
def definir_cor_fundo_celula(celula, cor_hex):
    """
    Define a cor de fundo de uma célula da tabela.
    cor_hex: Cor em formato hexadecimal (ex: #75B4FF).
    """
    cor_rgb = tuple(int(cor_hex.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
    tcPr = celula._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), cor_hex.lstrip('#'))
    tcPr.append(shd)

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
            adicionar_paragrafo_formatado(doc, f"Ano do Registro: {ano}", fonte='Arial', tamanho=12, negrito=False, cor=None)
            df_ano = df_filtrado[df_filtrado['ANO'] == ano]
            
            # Define os novos cabeçalhos
            colunas_tabela = ["PTM", 'VALOR TOTAL FEM', "VALOR REPASSADO", "DATA ÚLTIMO PAGAMENTO", 'STATUS']
            
            tabela = doc.add_table(rows=1, cols=len(colunas_tabela))
            tabela.style = 'Table Grid'
            
            # Aplica a cor de fundo no cabeçalho
            hdr_cells = tabela.rows[0].cells
            for i, coluna in enumerate(colunas_tabela):
                hdr_cells[i].text = coluna
                definir_cor_fundo_celula(hdr_cells[i], '#75B4FF')  # Cor de fundo #75B4FF
            
            # Adiciona as linhas da tabela
            for _, linha in df_ano.iterrows():
                row_cells = tabela.add_row().cells
                for i, coluna in enumerate(colunas_tabela):
                    if coluna == "PTM":
                        valor = linha.get("PROJETO DETALHADO", '0')
                    elif coluna == "VALOR TOTAL FEM":
                        valor = linha.get('TETO FEM_x' if 'TETO FEM_x' in df_ano.columns else 'TETO FEM_y', '0')
                        valor = formatar_moeda(valor)
                    elif coluna == "VALOR REPASSADO":
                        valor = linha.get("REPASSE_VÁLIDO", '0')
                        valor = formatar_moeda(valor)
                    elif coluna == "DATA ÚLTIMO PAGAMENTO":
                        valor = str(linha.get("DATA ÚLTIMO PAGAMENTO", '0'))
                        try:
                            valor = datetime.strptime(valor, '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y')
                        except ValueError:
                            pass
                    elif coluna == "STATUS":
                        valor = linha.get('STATUS OBRA_x' if 'STATUS OBRA_x' in df_ano.columns else 'STATUS OBRA_y', '0')
                    row_cells[i].text = str(valor)
            
            # Adiciona a linha de total
            total_row = tabela.add_row().cells
            total_row[0].text = "VALOR TOTAL"
            total_row[0].paragraphs[0].runs[0].bold = True  # Formatação em negrito
            
            # Calcula e formata os totais para VALOR TOTAL FEM e VALOR REPASSADO
            coluna_teto_fem = 'TETO FEM_x' if 'TETO FEM_x' in df_ano.columns else 'TETO FEM_y'
            total_teto_fem = df_ano[coluna_teto_fem].replace('', '0').astype(float).sum()
            total_row[1].text = formatar_moeda(total_teto_fem)
            
            total_repasse_valido = df_ano['REPASSE_VÁLIDO'].replace('', '0').astype(float).sum()
            total_row[2].text = formatar_moeda(total_repasse_valido)
            
            # Aplica a formatação da tabela (fonte Arial tamanho 8)
            formatar_tabela(tabela)
        
        doc.save(caminho_nota)
        st.success("✅ Nota Técnica gerada com sucesso!")
        with open(caminho_nota, "rb") as file:
            st.download_button("📥 Baixar Nota Técnica (DOCX)", file, file_name="nota_tecnica.docx")
