import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import zipfile
import os
from datetime import datetime
import unicodedata

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Ofícios - ConPrev", layout="wide")

# ================= 1. FUNÇÕES DE LIMPEZA E NORMALIZAÇÃO =================

def remove_accents(input_str):
    """Remove acentos (Á -> A, ç -> c) para facilitar a busca."""
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def normalize_key(text):
    """
    Normaliza o nome da cidade para busca:
    - Maiúsculo
    - Sem acentos
    - Sem prefixos como 'MUNICIPIO DE'
    """
    text = remove_accents(str(text)).upper().strip()
    
    # Prefixos comuns que atrapalham a busca direta
    prefixes = [
        "MUNICIPIO DE ", 
        "PREFEITURA MUNICIPAL DE ", 
        "PREFEITURA DE ", 
        "CAMARA MUNICIPAL DE " # Opcional: remover se quiser diferenciar Câmara
    ]
    
    for p in prefixes:
        if text.startswith(p):
            text = text[len(p):].strip()
            
    return text

def carregar_dicionario_responsaveis(arquivo_upload):
    """
    Lê a lista de responsáveis e cria um dicionário normalizado.
    """
    try:
        # Tenta ler CSV com diferentes separadores e encodings
        if arquivo_upload.name.endswith('.csv'):
            try:
                df = pd.read_csv(arquivo_upload, sep=';', encoding='utf-8-sig')
            except:
                arquivo_upload.seek(0)
                try:
                    df = pd.read_csv(arquivo_upload, sep=';', encoding='latin-1')
                except:
                    arquivo_upload.seek(0)
                    df = pd.read_csv(arquivo_upload, sep=',', encoding='utf-8')
        else:
            df = pd.read_excel(arquivo_upload)

        # Normaliza colunas para achar 'Município' e 'Responsável'
        df.columns = [remove_accents(c).strip().lower() for c in df.columns]
        
        col_muni = next((c for c in df.columns if 'municipio' in c or 'cidade' in c), None)
        col_resp = next((c for c in df.columns if 'responsavel' in c or 'nome' in c or 'prefeito' in c), None)

        if not col_muni or not col_resp:
            st.error(f"Erro: Não encontrei as colunas 'Município' e 'Responsável' no arquivo. Colunas encontradas: {list(df.columns)}")
            return {}

        dic_resp = {}
        for _, row in df.iterrows():
            raw_muni = str(row[col_muni])
            clean_muni = normalize_key(raw_muni) # Ex: "MUNICIPIO DE ALMAS" -> "ALMAS"
            
            # Prioriza entradas que originalmente começavam com MUNICIPIO ou PREFEITURA
            # (Para evitar pegar o Presidente da Câmara se houver duplicidade)
            raw_upper = remove_accents(raw_muni).upper()
            is_priority = "MUNICIPIO" in raw_upper or "PREFEITURA" in raw_upper
            
            if clean_muni not in dic_resp or is_priority:
                dic_resp[clean_muni] = str(row[col_resp]).strip()
            
        return dic_resp

    except Exception as e:
        st.error(f"Erro ao ler arquivo de responsáveis: {e}")
        return {}

def buscar_responsavel(municipio_divida, db_responsaveis):
    """
    Busca o responsável tentando match exato ou aproximado.
    """
    muni_norm = normalize_key(municipio_divida) # Ex: "SÃO VALÉRIO" -> "SAO VALERIO"
    
    # 1. Tentativa Exata
    if muni_norm in db_responsaveis:
        return db_responsaveis[muni_norm]
    
    # 2. Tentativa "Contém" (Ex: 'BANDEIRANTES' busca 'BANDEIRANTES DO TOCANTINS')
    # Verifica se a chave do dicionário começa com o nome da cidade da dívida
    for key in db_responsaveis:
        if key.startswith(muni_norm): 
            return db_responsaveis[key]
            
    return "PREFEITO(A) MUNICIPAL"

# ================= 2. MANIPULAÇÃO WORD =================

def replace_everywhere(doc: Document, old: str, new: str) -> None:
    def repl(par):
        if old in par.text:
            for run in par.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)
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

def formatar_valor(val):
    if isinstance(val, (int, float)):
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(val)

def adicionar_linha_tabela(table, orgao, processo, valor, is_placeholder=False):
    row_cells = table.add_row().cells
    row_cells[0].text = orgao
    row_cells[1].text = processo
    row_cells[2].text = valor
    
    for i, cell in enumerate(row_cells):
        cell.vertical_alignment = 1
        for p in cell.paragraphs:
            if is_placeholder:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                if i == 2: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif i == 1: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else: p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if p.runs: p.runs[0].font.size = Pt(10)
            else: p.add_run().font.size = Pt(10)

def preencher_tabela(table, df_municipio):
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False 

    hdr_cells = table.rows[0].cells
    titulos = ['Órgão', 'Processo / Documento', 'Saldo em 31/12/2025']
    for i, titulo in enumerate(titulos):
        hdr_cells[i].text = titulo
        for p in hdr_cells[i].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(10)
            if not p.runs: p.add_run(titulo).font.bold = True

    df_work = df_municipio.copy()
    df_work['Sistema'] = df_work['Sistema'].fillna('').astype(str)
    
    df_pgfn = df_work[df_work['Sistema'].str.contains("PGFN", case=False)]
    df_rfb = df_work[~df_work['Sistema'].str.contains("PGFN", case=False)]

    if not df_rfb.empty:
        for _, row in df_rfb.iterrows():
            adicionar_linha_tabela(table, "Receita Federal do Brasil", str(row['Processo']), formatar_valor(row['Valor Original']))
    else:
        adicionar_linha_tabela(table, "Receita Federal do Brasil", "-", "-", is_placeholder=True)

    if not df_pgfn.empty:
        for _, row in df_pgfn.iterrows():
            adicionar_linha_tabela(table, "Procuradoria da Fazenda Nacional", str(row['Processo']), formatar_valor(row['Valor Original']))
    else:
        adicionar_linha_tabela(table, "Procuradoria da Fazenda Nacional", "-", "-", is_placeholder=True)

def inserir_tabela_no_placeholder(doc, df_municipio, placeholder="{{TABELA}}"):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = ""
            table = doc.add_table(rows=1, cols=3)
            paragraph._p.addnext(table._tbl)
            preencher_tabela(table, df_municipio)
            return True
    return False

# ================= 3. INTERFACE =================
st.title("Gerador de Ofícios - Saldo Dívida RFB")
st.markdown("---")

col1, col2, col3 = st.columns(3)
with col1:
    uploaded_excel = st.file_uploader("1. Planilha de Dívidas (Excel)", type=["xlsx"])
with col2:
    uploaded_template = st.file_uploader("2. Modelo do Ofício (Word)", type=["docx"])
with col3:
    uploaded_responsaveis = st.file_uploader("3. Lista de Responsáveis (CSV)", type=["csv"])

st.sidebar.header("Configuração")
num_inicial = st.sidebar.number_input("Número Inicial", value=46, step=1)
ano_doc = st.sidebar.number_input("Ano", value=2026)

# ================= 4. PROCESSAMENTO =================
if st.button("🚀 Gerar Arquivos (ZIP)"):
    if not uploaded_excel or not uploaded_template:
        st.error("Faltam arquivos (Planilha ou Modelo).")
        st.stop()
    
    db_responsaveis = {}
    if uploaded_responsaveis:
        db_responsaveis = carregar_dicionario_responsaveis(uploaded_responsaveis)
        st.success(f"Lista de Responsáveis carregada: {len(db_responsaveis)} registros encontrados.")

    try:
        # Carrega Dívidas
        df = pd.read_excel(uploaded_excel, engine='openpyxl')
        df = df.dropna(subset=['Processo'])
        col_municipio = 'Município' if 'Município' in df.columns else df.columns[0]
        df[col_municipio] = df[col_municipio].astype(str).str.strip()
        municipios = sorted(df[col_municipio].unique())

        zip_buffer = io.BytesIO()
        contador = num_inicial
        hoje = datetime.now()
        meses = {1:"janeiro", 2:"fevereiro", 3:"março", 4:"abril", 5:"maio", 6:"junho",
                 7:"julho", 8:"agosto", 9:"setembro", 10:"outubro", 11:"novembro", 12:"dezembro"}
        data_extenso = f"Goiânia, {hoje.day} de {meses[hoje.month]} de {hoje.year}."

        progress = st.progress(0)
        logs = []
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, muni in enumerate(municipios):
                uploaded_template.seek(0)
                doc = Document(uploaded_template)

                df_muni = df[df[col_municipio] == muni]
                
                # --- Busca UF ---
                uf = "GO"
                if 'Arquivo' in df_muni.columns:
                    try: 
                        parts = str(df_muni.iloc[0]['Arquivo']).split('-')
                        if len(parts) > 0 and len(parts[0].strip()) == 2: uf = parts[0].strip()
                    except: pass
                
                # --- Busca Responsável Inteligente ---
                nome_pref = "PREFEITO(A) MUNICIPAL"
                if db_responsaveis:
                    nome_pref = buscar_responsavel(muni, db_responsaveis)
                    if nome_pref == "PREFEITO(A) MUNICIPAL":
                        logs.append(f"⚠️ {muni}: Responsável não encontrado (Tentado buscar como: '{normalize_key(muni)}')")

                num_fmt = f"{contador:03d}/{ano_doc}"
                
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
                sucesso = inserir_tabela_no_placeholder(doc, df_muni, "{{TABELA}}")
                if not sucesso:
                    sucesso = inserir_tabela_no_placeholder(doc, df_muni, "{{TABELA_DEBITOS}}")
                
                if not sucesso:
                    logs.append(f"❌ {muni}: Placeholder {{TABELA}} não encontrado no Word.")
                    table_fallback = doc.add_table(rows=1, cols=3)
                    preencher_tabela(table_fallback, df_muni)

                doc_io = io.BytesIO()
                doc.save(doc_io)
                
                nome_zip = f"{contador:03d}-{ano_doc} - {uf} - {muni} - Saldo Divida RFB-PGFN.docx"
                zf.writestr(nome_zip, doc_io.getvalue())
                
                contador += 1
                progress.progress((i+1)/len(municipios))
        
        st.success(f"✅ Concluído! {len(municipios)} ofícios gerados.")
        
        if logs:
            with st.expander("⚠️ Relatório de Alertas"):
                st.write(f"Total de alertas: {len(logs)}")
                for log in logs: st.write(log)

        st.download_button("⬇️ Baixar ZIP Completo", zip_buffer.getvalue(), 
                           file_name=f"Oficios_SaldoDivida_{datetime.now().strftime('%H%M')}.zip", 
                           mime="application/zip")

    except Exception as e:
        st.error(f"Erro Crítico: {e}")
