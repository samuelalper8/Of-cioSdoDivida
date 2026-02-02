import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import io
import zipfile
import os
from datetime import datetime

# --- CONFIGURAÇÕES DE CAMINHO ---
CAMINHO_MODELO_LOCAL = r"C:\Users\Estação-reserva\OneDrive\Área de Trabalho\Z_APP_SaldoDivida\Modelo_Saldo_Divida.docx"
NOME_ARQUIVO_MODELO = "Modelo_Saldo_Divida.docx"

st.set_page_config(page_title="Gerador de Ofícios (Word)", layout="wide")

# ================= 1. FUNÇÕES DE MANIPULAÇÃO WORD =================

def replace_everywhere(doc: Document, old: str, new: str) -> None:
    """Substitui texto em parágrafos, tabelas e cabeçalhos."""
    def repl(par):
        if old in par.text:
            # Tenta substituir preservando formatação (runs)
            for run in par.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)
            # Fallback: se a palavra estiver quebrada em runs diferentes
            if old in par.text:
                par.text = par.text.replace(old, new)

    for p in doc.paragraphs: repl(p)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs: repl(p)
    for s in doc.sections:
        for h in [s.header, s.first_page_header, s.footer, s.first_page_footer]:
            if h:
                for p in h.paragraphs: repl(p)

def mover_tabela_para_placeholder(doc, table, placeholder_text):
    """
    Encontra o parágrafo que tem o placeholder (ex: {{TABELA_DEBITOS}}),
    move a tabela para logo após ele e apaga o texto do placeholder.
    Isso permite que a tabela fique entre o texto e as assinaturas.
    """
    target_p = None
    for p in doc.paragraphs:
        if placeholder_text in p.text:
            target_p = p
            break
            
    if target_p:
        # Lógica XML para mover a tabela (table._tbl) para depois do parágrafo (target_p._p)
        target_p._p.addnext(table._tbl)
        target_p.text = "" # Limpa o placeholder
        return True
    return False

def criar_tabela_divida(doc, df_municipio):
    """Cria a tabela e retorna o objeto table para ser posicionado."""
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False 
    
    # Define larguras relativas (opcional, ajustável via XML se necessário)
    # python-docx não lida bem com larguras fixas em cm sem hacks, 
    # mas o autofit costuma funcionar bem para texto.

    # Cabeçalho
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Órgão'
    hdr_cells[1].text = 'Processo / Documento'
    hdr_cells[2].text = 'Saldo em 31/12/2025'
    
    for cell in hdr_cells:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            run.font.bold = True
            run.font.size = Pt(10)

    # Dados
    for index, row in df_municipio.iterrows():
        row_cells = table.add_row().cells
        
        # Lógica do Órgão
        orgao = "Receita Federal do Brasil"
        if "PGFN" in str(row.get('Sistema', '')): orgao = "Procuradoria da Fazenda Nacional"
            
        processo = str(row['Processo'])
        val = row['Valor Original']
        
        # Formatação Monetária
        if isinstance(val, (int, float)):
            valor_str = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            valor_str = str(val)

        row_cells[0].text = orgao
        row_cells[1].text = processo
        row_cells[2].text = valor_str
        
        for cell in row_cells:
            cell.vertical_alignment = 1
            for p in cell.paragraphs:
                p.runs[0].font.size = Pt(10)
                
    return table

def carregar_prefeitos():
    """Lê prefeitos.csv"""
    arquivo = "prefeitos.csv"
    dic = {}
    if os.path.exists(arquivo):
        try:
            try: df = pd.read_csv(arquivo, encoding='utf-8')
            except: df = pd.read_csv(arquivo, encoding='latin-1', sep=',')
            
            for _, row in df.iterrows():
                if 'Município' in df.columns:
                    dic[str(row['Município']).strip().upper()] = str(row['Prefeito']).strip()
        except: pass
    return dic

# ================= 2. INTERFACE =================
st.title("Geração de Ofícios - Saldo Dívida RFB")
st.markdown("---")

# --- Lógica de Seleção do Arquivo Word ---
usar_modelo_local = False
doc_template = None

# Verifica se o arquivo existe na pasta indicada
if os.path.exists(CAMINHO_MODELO_LOCAL):
    st.success(f"✅ Modelo encontrado em: `{CAMINHO_MODELO_LOCAL}`")
    usar_modelo_local = True
else:
    st.warning(f"⚠️ Modelo não encontrado na pasta padrão. Por favor, faça o upload.")

# Colunas de Upload
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. Planilha Excel (Dados)", type=["xlsx"])

with col2:
    if not usar_modelo_local:
        uploaded_template = st.file_uploader("2. Modelo Word (.docx)", type=["docx"])
    else:
        st.info("Usando modelo local. (Para trocar, renomeie ou remova o arquivo da pasta).")
        uploaded_template = None

# Sidebar Config
st.sidebar.header("Parâmetros")
num_inicial = st.sidebar.number_input("Nº Inicial", value=1, step=1)
ano_doc = st.sidebar.number_input("Ano", value=2026)

# Instruções
with st.expander("📝 Instruções para o Modelo Word"):
    st.markdown("""
    Certifique-se que o arquivo `.docx` tenha estes códigos onde os dados devem entrar:
    - `{{MUNICIPIO}}`
    - `{{UF}}`
    - `{{PREFEITO}}`
    - `{{NUM_OFICIO}}` (Ex: 001/2026)
    - `{{DATA_EXTENSO}}` (Ex: Goiânia, 27 de...)
    - `{{TABELA_DEBITOS}}` **(Coloque isso numa linha sozinha entre o texto e a assinatura)**
    """)

# ================= 3. PROCESSAMENTO =================
if st.button("Gerar Arquivos (ZIP)"):
    # Validação
    if not uploaded_excel:
        st.error("Faltou a planilha Excel!")
        st.stop()
    
    if not usar_modelo_local and not uploaded_template:
        st.error("Faltou o Modelo Word!")
        st.stop()

    try:
        # Carrega Excel
        df = pd.read_excel(uploaded_excel, engine='openpyxl')
        df = df.dropna(subset=['Processo'])
        db_prefeitos = carregar_prefeitos()
        
        col_municipio = 'Município' if 'Município' in df.columns else df.columns[0]
        municipios = sorted(df[col_municipio].astype(str).unique())
        
        # Buffer ZIP
        zip_buffer = io.BytesIO()
        contador = num_inicial
        
        # Data
        hoje = datetime.now()
        meses = {1:"janeiro", 2:"fevereiro", 3:"março", 4:"abril", 5:"maio", 6:"junho",
                 7:"julho", 8:"agosto", 9:"setembro", 10:"outubro", 11:"novembro", 12:"dezembro"}
        data_extenso = f"Goiânia, {hoje.day} de {meses[hoje.month]} de {hoje.year}."

        progress = st.progress(0)
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, muni in enumerate(municipios):
                # Carrega o Word (Do disco ou do Upload)
                if usar_modelo_local:
                    doc = Document(CAMINHO_MODELO_LOCAL)
                else:
                    uploaded_template.seek(0)
                    doc = Document(uploaded_template)

                # Filtra dados
                df_muni = df[df[col_municipio] == muni]
                
                # UF
                uf = "GO"
                if 'Arquivo' in df_muni.columns:
                    try: 
                        parts = str(df_muni.iloc[0]['Arquivo']).split('-')
                        if len(parts) > 0 and len(parts[0].strip()) == 2: uf = parts[0].strip()
                    except: pass
                
                # Prefeito
                nome_pref = db_prefeitos.get(muni.upper(), "PREFEITO(A) MUNICIPAL")
                
                # Variáveis
                num_fmt = f"{contador:03d}/{ano_doc}" # 001/2026
                
                # Substituições
                replaces = {
                    "{{MUNICIPIO}}": muni.upper(),
                    "{{UF}}": uf,
                    "{{PREFEITO}}": nome_pref.upper(),
                    "{{NUM_OFICIO}}": num_fmt,
                    "{{DATA_EXTENSO}}": data_extenso
                }
                
                for k, v in replaces.items():
                    replace_everywhere(doc, k, v)
                
                # Tabela
                # 1. Cria a tabela (ela vai pro fim do doc temporariamente)
                tabela = criar_tabela_divida(doc, df_muni)
                
                # 2. Tenta mover para o placeholder {{TABELA_DEBITOS}}
                sucesso_mover = mover_tabela_para_placeholder(doc, tabela, "{{TABELA_DEBITOS}}")
                
                if not sucesso_mover:
                    # Se não achou o placeholder, avisa no console (opcional)
                    print(f"Aviso: Placeholder de tabela não encontrado para {muni}")

                # Salva no ZIP
                doc_io = io.BytesIO()
                doc.save(doc_io)
                zf.writestr(f"{contador:03d}-{ano_doc} - {muni}.docx", doc_io.getvalue())
                
                contador += 1
                progress.progress((i+1)/len(municipios))
                
        st.success(f"Sucesso! {len(municipios)} ofícios gerados.")
        st.download_button("⬇️ Baixar Ofícios (ZIP)", zip_buffer.getvalue(), 
                           file_name=f"Oficios_SaldoDivida_{datetime.now().strftime('%H%M')}.zip", 
                           mime="application/zip")

    except Exception as e:
        st.error(f"Erro: {e}")
