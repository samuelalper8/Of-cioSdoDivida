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

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Ofícios - ConPrev", layout="wide")

# ================= 1. FUNÇÕES DE SUPORTE =================

def gerar_modelo_csv():
    """Gera um CSV modelo para o usuário baixar."""
    data = {
        'Município': ['Goiânia', 'Anápolis', 'Aparecida de Goiânia'],
        'Responsável': ['Nome do Prefeito 1', 'Nome do Prefeito 2', 'Nome do Prefeito 3']
    }
    df = pd.read_json(pd.DataFrame(data).to_json()) # Truque para garantir tipos
    # Gera CSV separado por ponto e vírgula (Padrão Excel Brasil)
    return df.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

def carregar_dicionario_responsaveis(arquivo_upload):
    """
    Lê o arquivo de responsáveis (CSV ou Excel) e retorna um dicionário:
    { 'MUNICÍPIO': 'NOME DO RESPONSÁVEL' }
    """
    try:
        if arquivo_upload.name.endswith('.csv'):
            # Tenta ler com separador ; (comum no Brasil) e depois ,
            try:
                df = pd.read_csv(arquivo_upload, sep=';', encoding='utf-8-sig')
            except:
                arquivo_upload.seek(0)
                df = pd.read_csv(arquivo_upload, sep=',', encoding='latin-1')
        else:
            df = pd.read_excel(arquivo_upload)

        # Padronização de colunas (Remove acentos e espaços para achar as colunas)
        df.columns = df.columns.str.strip().str.lower()
        
        # Procura colunas chaves
        col_muni = next((c for c in df.columns if 'munic' in c or 'cidade' in c), None)
        col_resp = next((c for c in df.columns if 'respons' in c or 'nome' in c or 'prefeito' in c), None)

        if not col_muni or not col_resp:
            st.error("Erro na Planilha de Responsáveis: Não encontrei as colunas 'Município' e 'Responsável'. Baixe o modelo para ver o formato correto.")
            return {}

        # Cria o dicionário { CIDADE: NOME }
        dic_resp = {}
        for _, row in df.iterrows():
            cidade = str(row[col_muni]).strip().upper()
            nome = str(row[col_resp]).strip()
            dic_resp[cidade] = nome
            
        return dic_resp

    except Exception as e:
        st.error(f"Erro ao ler arquivo de responsáveis: {e}")
        return {}

# ================= 2. MANIPULAÇÃO WORD =================

def replace_everywhere(doc: Document, old: str, new: str) -> None:
    """Substitui texto em todo o documento."""
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

    for index, row in df_municipio.iterrows():
        row_cells = table.add_row().cells
        
        orgao = "Receita Federal do Brasil"
        if "PGFN" in str(row.get('Sistema', '')): orgao = "Procuradoria da Fazenda Nacional"
            
        processo = str(row['Processo'])
        val = row['Valor Original']
        
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
                if p.runs: p.runs[0].font.size = Pt(10)
                else: p.add_run().font.size = Pt(10)

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
st.markdown("Preencha os dados abaixo para gerar os documentos.")

# Container de Download do Modelo
with st.container():
    st.info("💡 **Dica:** Baixe o modelo da lista de responsáveis para preencher corretamente.")
    csv_modelo = gerar_modelo_csv()
    st.download_button(
        label="📥 Baixar Modelo de Lista de Responsáveis (CSV)",
        data=csv_modelo,
        file_name="Modelo_Responsaveis.csv",
        mime="text/csv",
    )

st.markdown("---")

col1, col2, col3 = st.columns(3)
with col1:
    uploaded_excel = st.file_uploader("1. Planilha de Dívidas (Excel)", type=["xlsx"])
with col2:
    uploaded_template = st.file_uploader("2. Modelo do Ofício (Word)", type=["docx"])
with col3:
    uploaded_responsaveis = st.file_uploader("3. Lista de Responsáveis (CSV/Excel)", type=["csv", "xlsx"])

st.sidebar.header("Configuração")
num_inicial = st.sidebar.number_input("Número Inicial", value=46, step=1)
ano_doc = st.sidebar.number_input("Ano", value=2026)

# ================= 4. PROCESSAMENTO =================
if st.button("🚀 Gerar Arquivos (ZIP)"):
    # Validações
    if not uploaded_excel:
        st.error("Faltou a Planilha de Dívidas!")
        st.stop()
    if not uploaded_template:
        st.error("Faltou o Modelo Word!")
        st.stop()
    if not uploaded_responsaveis:
        st.error("Faltou a Lista de Responsáveis!")
        st.stop()

    try:
        # 1. Carrega Dados da Dívida
        df = pd.read_excel(uploaded_excel, engine='openpyxl')
        df = df.dropna(subset=['Processo'])
        col_municipio = 'Município' if 'Município' in df.columns else df.columns[0]
        df[col_municipio] = df[col_municipio].astype(str).str.strip()
        municipios = sorted(df[col_municipio].unique())

        # 2. Carrega Dados dos Responsáveis (Dinâmico)
        db_responsaveis = carregar_dicionario_responsaveis(uploaded_responsaveis)
        
        if not db_responsaveis:
            st.stop() # Parar se houve erro na leitura dos responsáveis

        # 3. Preparação
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
                
                # --- Busca Responsável no Arquivo Uploaded ---
                nome_pref = db_responsaveis.get(muni.upper(), "RESPONSÁVEL NÃO ENCONTRADO")
                if nome_pref == "RESPONSÁVEL NÃO ENCONTRADO":
                    logs.append(f"⚠️ {muni}: Responsável não encontrado na lista enviada.")

                num_fmt = f"{contador:03d}/{ano_doc}"
                
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
                sucesso = inserir_tabela_no_placeholder(doc, df_muni, "{{TABELA}}")
                if not sucesso:
                    sucesso = inserir_tabela_no_placeholder(doc, df_muni, "{{TABELA_DEBITOS}}")
                
                if not sucesso:
                    logs.append(f"⚠️ {muni}: Placeholder {{TABELA}} não encontrado no Word.")
                    table_fallback = doc.add_table(rows=1, cols=3)
                    preencher_tabela(table_fallback, df_muni)

                # Salva
                doc_io = io.BytesIO()
                doc.save(doc_io)
                
                # Nome do arquivo com UF + Saldo Divida RFB-PGFN
                nome_zip = f"{contador:03d}-{ano_doc} - {uf} - {muni} - Saldo Divida RFB-PGFN.docx"
                zf.writestr(nome_zip, doc_io.getvalue())
                
                contador += 1
                progress.progress((i+1)/len(municipios))
        
        st.success(f"✅ Processamento concluído! {len(municipios)} ofícios gerados.")
        
        if logs:
            with st.expander("⚠️ Alertas de Processamento"):
                for log in logs: st.write(log)

        st.download_button("⬇️ Baixar ZIP Completo", zip_buffer.getvalue(), 
                           file_name=f"Oficios_SaldoDivida_{datetime.now().strftime('%H%M')}.zip", 
                           mime="application/zip")

    except Exception as e:
        st.error(f"Erro Crítico: {e}")
