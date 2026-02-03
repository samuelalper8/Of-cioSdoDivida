import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import zipfile
import os
from datetime import datetime
import unicodedata
import re
import pdfplumber # Nova biblioteca para ler os extratos

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Ofícios - ConPrev", layout="wide")

# ================= 1. FUNÇÕES DE LIMPEZA E NORMALIZAÇÃO =================

def remove_accents(input_str):
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def normalize_key(text):
    """Normaliza nomes de cidades para busca."""
    if pd.isna(text): return ""
    text = remove_accents(str(text)).upper().strip()
    prefixes = ["MUNICIPIO DE ", "PREFEITURA MUNICIPAL DE ", "PREFEITURA DE ", "CAMARA MUNICIPAL DE ", "FUNDO MUNICIPAL DE "]
    for p in prefixes:
        if text.startswith(p):
            text = text[len(p):].strip()
    return text

def parse_currency(value_str):
    """Converte strings como '20.782,71' para float."""
    try:
        clean = str(value_str).replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
        return float(clean)
    except:
        return 0.0

def formatar_valor(val):
    if isinstance(val, (int, float)):
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(val)

# ================= 2. EXTRAÇÃO DE PDFS (NOVA FUNÇÃO) =================

def extrair_dados_pgfn(uploaded_pdfs):
    """
    Lê os PDFs enviados e extrai: Município, Inscrição, Modalidade e Valor.
    Retorna um dicionário: { 'MUNICIPIO': [ {dados_divida}, ... ] }
    """
    dados_extraidos = {}

    for pdf_file in uploaded_pdfs:
        try:
            # Tenta descobrir o município pelo nome do arquivo (ex: GO_Itaberai_PGFN...)
            filename = pdf_file.name
            parts = filename.split('_')
            municipio_nome = "DESCONHECIDO"
            
            # Heurística simples para pegar o nome da cidade no arquivo
            for part in parts:
                if part.upper() not in ["GO", "TO", "PGFN", "RPPS", "PASEP", "FMS", "SMS"]:
                    # Assume que partes que não são siglas comuns podem ser a cidade
                    # Remove acentos para garantir
                    municipio_nome = normalize_key(part)
                    break
            
            with pdfplumber.open(pdf_file) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() + "\n"
                
                # --- Lógica de Extração (Regex) ---
                
                # 1. Inscrição / Processo
                inscricao = "Não identificado"
                # Padrão comum PGFN: 11 7 11 002247-69 ou Numeração única
                match_insc = re.search(r'N[º°].*?inscrição:?\s*([\d\s\.\/-]+)', text, re.IGNORECASE)
                if match_insc:
                    inscricao = match_insc.group(1).strip()
                
                # 2. Modalidade
                modalidade = "Dívida Ativa" # Padrão
                # Procura termos comuns nos extratos
                if "Transação Excepcional" in text:
                    modalidade = "Transação Excepcional"
                elif "Parcelamento Simplificado" in text:
                    modalidade = "Parcelamento Simplificado"
                elif "Sispar" in text or "SISPAR" in text:
                    modalidade = "Parcelamento SISPAR"
                else:
                    # Tenta capturar campo "Modalidade:"
                    match_mod = re.search(r'Modalidade:?\s*(.*?)\n', text, re.IGNORECASE)
                    if match_mod:
                        modalidade = match_mod.group(1).strip()

                # 3. Valor (Saldo Devedor com Juros ou Valor Consolidado)
                valor = 0.0
                # Prioridade 1: Saldo Devedor Total / Atualizado
                match_val = re.search(r'(Saldo Devedor|Valor Consolidado|Valor Total).*?R\$\s*([\d\.,]+)', text, re.IGNORECASE)
                
                if match_val:
                    valor = parse_currency(match_val.group(2))
                else:
                    # Tenta achar valores isolados no fim do documento (comum em extratos simples)
                    # Busca o último valor monetário grande na página
                    valores_encontrados = re.findall(r'([\d\.,]{5,})', text) # Pega numeros com formato de dinheiro
                    if valores_encontrados:
                         # Assume o último como o total (arriscado, mas fallback)
                         valor = parse_currency(valores_encontrados[-1])

                # Adiciona ao dicionário
                if municipio_nome not in dados_extraidos:
                    dados_extraidos[municipio_nome] = []
                
                dados_extraidos[municipio_nome].append({
                    'Processo': inscricao,
                    'Modalidade': modalidade,
                    'Valor Original': valor,
                    'Sistema': 'PGFN (PDF)',
                    'Fonte': 'PDF'
                })

        except Exception as e:
            print(f"Erro ao ler PDF {pdf_file.name}: {e}")
            
    return dados_extraidos

# ================= 3. CARGA DE DADOS =================

def carregar_dicionario_responsaveis(arquivo_upload):
    try:
        if arquivo_upload.name.endswith('.csv'):
            try: df = pd.read_csv(arquivo_upload, sep=';', encoding='utf-8-sig')
            except: 
                arquivo_upload.seek(0)
                try: df = pd.read_csv(arquivo_upload, sep=';', encoding='latin-1')
                except: 
                    arquivo_upload.seek(0)
                    df = pd.read_csv(arquivo_upload, sep=',', encoding='utf-8')
        else:
            df = pd.read_excel(arquivo_upload)

        df.columns = [remove_accents(c).strip().lower() for c in df.columns]
        col_muni = next((c for c in df.columns if any(x in c for x in ['municipio', 'cidade'])), None)
        col_resp = next((c for c in df.columns if any(x in c for x in ['responsavel', 'nome', 'prefeito'])), None)

        if not col_muni or not col_resp: return {}

        dic_resp = {}
        for _, row in df.iterrows():
            raw_muni = str(row[col_muni])
            clean_muni = normalize_key(raw_muni)
            raw_upper = remove_accents(raw_muni).upper()
            is_priority = "MUNICIPIO" in raw_upper or "PREFEITURA" in raw_upper
            if clean_muni and (clean_muni not in dic_resp or is_priority):
                dic_resp[clean_muni] = str(row[col_resp]).strip()
        return dic_resp
    except: return {}

def buscar_responsavel(municipio_divida, db_responsaveis):
    muni_target = normalize_key(municipio_divida) 
    if muni_target in db_responsaveis: return db_responsaveis[muni_target]
    for key_db in db_responsaveis:
        if key_db.startswith(muni_target) or muni_target.startswith(key_db):
            return db_responsaveis[key_db]
    return "PREFEITO(A) MUNICIPAL"

# ================= 4. MANIPULAÇÃO WORD =================

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

def adicionar_linha_tabela(table, orgao, modalidade, processo, valor, is_placeholder=False):
    row_cells = table.add_row().cells
    
    # Coluna 1: Órgão + Modalidade
    p1 = row_cells[0].paragraphs[0]
    r1_orgao = p1.add_run(orgao)
    r1_orgao.font.bold = False
    if not is_placeholder and modalidade:
        p1.add_run(f"\n({modalidade})").font.size = Pt(9)
    
    # Coluna 2: Processo
    row_cells[1].text = processo
    
    # Coluna 3: Valor
    row_cells[2].text = valor
    
    # --- FORMATAÇÃO PEDIDA ---
    # Coluna 1 (Órgão/Mod): Alinhado à Esquerda
    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Coluna 2 (Processo): Centralizado
    for p in row_cells[1].paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    # Coluna 3 (Valor): Centralizado (conforme pedido)
    for p in row_cells[2].paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER # Alterado de RIGHT para CENTER

    # Ajuste de Fonte Geral
    for cell in row_cells:
        cell.vertical_alignment = 1
        for p in cell.paragraphs:
            if p.runs: 
                p.runs[0].font.size = Pt(10)

def preencher_tabela(table, df_excel_muni, lista_pdfs_muni):
    """
    Combina dados do Excel (RFB) e da lista de PDFs extraídos (PGFN).
    """
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False 

    # Cabeçalho
    hdr_cells = table.rows[0].cells
    titulos = ['Órgão / Modalidade', 'Processo / Documento', 'Saldo em 31/12/2025'] # Atualizei título
    for i, titulo in enumerate(titulos):
        hdr_cells[i].text = titulo
        for p in hdr_cells[i].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(10)
            if not p.runs: p.add_run(titulo).font.bold = True

    # --- 1. DADOS DA RFB (Vindos do Excel) ---
    # Filtra apenas o que NÃO é PGFN no Excel (pois PGFN vamos preferir do PDF se tiver)
    df_rfb = df_excel_muni[~df_excel_muni['Sistema'].astype(str).str.contains("PGFN", case=False, na=False)]
    
    if not df_rfb.empty:
        for _, row in df_rfb.iterrows():
            modalidade = str(row['Modalidade']) if pd.notna(row['Modalidade']) else ""
            adicionar_linha_tabela(
                table, 
                "Receita Federal do Brasil", 
                modalidade,
                str(row['Processo']), 
                formatar_valor(row['Valor Original'])
            )
    else:
        adicionar_linha_tabela(table, "Receita Federal do Brasil", "", "-", "-", is_placeholder=True)

    # --- 2. DADOS DA PGFN (Vindos dos PDFs - Prioridade) ---
    if lista_pdfs_muni:
        # Se temos PDFs, usamos os dados deles
        for item in lista_pdfs_muni:
            adicionar_linha_tabela(
                table,
                "Procuradoria Geral da Fazenda Nacional", # Nome Atualizado
                item['Modalidade'],
                item['Processo'],
                formatar_valor(item['Valor Original'])
            )
    else:
        # Se NÃO temos PDFs, olhamos se sobrou algo de PGFN no Excel
        df_pgfn_excel = df_excel_muni[df_excel_muni['Sistema'].astype(str).str.contains("PGFN", case=False, na=False)]
        
        if not df_pgfn_excel.empty:
            for _, row in df_pgfn_excel.iterrows():
                modalidade = str(row['Modalidade']) if pd.notna(row['Modalidade']) else ""
                adicionar_linha_tabela(
                    table,
                    "Procuradoria Geral da Fazenda Nacional",
                    modalidade,
                    str(row['Processo']),
                    formatar_valor(row['Valor Original'])
                )
        else:
            # Se não tem nem no PDF nem no Excel
            adicionar_linha_tabela(table, "Procuradoria Geral da Fazenda Nacional", "", "-", "-", is_placeholder=True)

def inserir_tabela_no_placeholder(doc, df_municipio, dados_pgfn_pdf, placeholder="{{TABELA}}"):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = ""
            table = doc.add_table(rows=1, cols=3)
            paragraph._p.addnext(table._tbl)
            preencher_tabela(table, df_municipio, dados_pgfn_pdf)
            return True
    return False

# ================= 5. INTERFACE =================
st.title("Gerador de Ofícios Inteligente 2.0")
st.markdown("Agora com suporte a leitura de **PDFs da PGFN** e formatação personalizada.")

col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. Planilha Excel (Dados RFB)", type=["xlsx"])
    uploaded_pdfs = st.file_uploader("4. PDFs PGFN (Arraste todos aqui)", type=["pdf"], accept_multiple_files=True)
with col2:
    uploaded_template = st.file_uploader("2. Modelo do Ofício (Word)", type=["docx"])
    uploaded_responsaveis = st.file_uploader("3. Lista de Responsáveis (CSV)", type=["csv"])

st.sidebar.header("Configuração")
num_inicial = st.sidebar.number_input("Número Inicial", value=46, step=1)
ano_doc = st.sidebar.number_input("Ano", value=2026)

# ================= 6. PROCESSAMENTO =================
if st.button("🚀 Gerar Arquivos (ZIP)"):
    if not uploaded_excel or not uploaded_template:
        st.error("Arquivos obrigatórios faltando (Excel ou Word)!")
        st.stop()
    
    # Carga de Responsáveis
    db_responsaveis = {}
    if uploaded_responsaveis:
        db_responsaveis = carregar_dicionario_responsaveis(uploaded_responsaveis)

    # Carga de PDFs (Processamento Pesado)
    dados_pgfn_extraidos = {}
    if uploaded_pdfs:
        with st.spinner("Lendo PDFs da PGFN... isso pode levar alguns segundos."):
            dados_pgfn_extraidos = extrair_dados_pgfn(uploaded_pdfs)
        st.success(f"{len(uploaded_pdfs)} PDFs processados com sucesso!")

    try:
        # Carga Excel
        df = pd.read_excel(uploaded_excel, engine='openpyxl')
        df = df.dropna(subset=['Processo'])
        col_municipio = 'Município' if 'Município' in df.columns else df.columns[0]
        df[col_municipio] = df[col_municipio].astype(str).str.strip()
        
        # Lista unificada de municípios (Excel + PDFs)
        municipios_excel = set(df[col_municipio].unique())
        municipios_pdfs = set(dados_pgfn_extraidos.keys())
        municipios_totais = sorted(list(municipios_excel.union(municipios_pdfs)))

        zip_buffer = io.BytesIO()
        contador = num_inicial
        hoje = datetime.now()
        meses = {1:"janeiro", 2:"fevereiro", 3:"março", 4:"abril", 5:"maio", 6:"junho",
                 7:"julho", 8:"agosto", 9:"setembro", 10:"outubro", 11:"novembro", 12:"dezembro"}
        data_extenso = f"Goiânia, {hoje.day} de {meses[hoje.month]} de {hoje.year}."

        progress = st.progress(0)
        logs = []
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, muni_raw in enumerate(municipios_totais):
                # Normaliza para busca
                muni_clean = normalize_key(muni_raw)
                
                # Prepara Doc
                uploaded_template.seek(0)
                doc = Document(uploaded_template)

                # Dados Excel para este município
                df_muni = df[df[col_municipio] == muni_raw]
                
                # Dados PDF para este município
                # Tenta achar a chave no dicionario de PDF que bate com a cidade atual
                lista_pgfn = []
                # Busca direta
                if muni_clean in dados_pgfn_extraidos:
                    lista_pgfn = dados_pgfn_extraidos[muni_clean]
                else:
                    # Busca aproximada
                    for k_pdf in dados_pgfn_extraidos.keys():
                        if k_pdf.startswith(muni_clean) or muni_clean.startswith(k_pdf):
                            lista_pgfn = dados_pgfn_extraidos[k_pdf]
                            break

                # UF
                uf = "GO"
                if not df_muni.empty and 'Arquivo' in df_muni.columns:
                    try: 
                        parts = str(df_muni.iloc[0]['Arquivo']).split('-')
                        if len(parts) > 0 and len(parts[0].strip()) == 2: uf = parts[0].strip()
                    except: pass
                
                # Responsável
                nome_pref = "PREFEITO(A) MUNICIPAL"
                if db_responsaveis:
                    nome_pref = buscar_responsavel(muni_raw, db_responsaveis)
                    if nome_pref == "PREFEITO(A) MUNICIPAL":
                        logs.append(f"⚠️ {muni_raw}: Responsável não encontrado.")

                num_fmt = f"{contador:03d}/{ano_doc}"
                
                # Text Replaces
                replaces = {
                    "{{MUNICIPIO}}": muni_raw.upper(),
                    "{{UF}}": uf,
                    "{{PREFEITO}}": nome_pref.upper(),
                    "{{NUM_OFICIO}}": num_fmt,
                    "{{DATA_EXTENSO}}": data_extenso
                }
                
                for k, v in replaces.items():
                    replace_everywhere(doc, k, v)
                
                # Tabela Híbrida (Excel + PDF)
                sucesso = inserir_tabela_no_placeholder(doc, df_muni, lista_pgfn, "{{TABELA}}")
                if not sucesso:
                    sucesso = inserir_tabela_no_placeholder(doc, df_muni, lista_pgfn, "{{TABELA_DEBITOS}}")
                
                if not sucesso:
                    logs.append(f"❌ {muni_raw}: Placeholder {{TABELA}} ausente.")

                doc_io = io.BytesIO()
                doc.save(doc_io)
                
                nome_zip = f"{contador:03d}-{ano_doc} - {uf} - {muni_raw} - Saldo Divida RFB-PGFN.docx"
                zf.writestr(nome_zip, doc_io.getvalue())
                
                contador += 1
                progress.progress((i+1)/len(municipios_totais))
        
        st.success(f"✅ Concluído! {len(municipios_totais)} ofícios gerados.")
        
        if logs:
            with st.expander("⚠️ Alertas"):
                for log in logs: st.write(log)

        st.download_button("⬇️ Baixar ZIP", zip_buffer.getvalue(), 
                           file_name=f"Oficios_PGFN_Atualizados_{datetime.now().strftime('%H%M')}.zip", 
                           mime="application/zip")

    except Exception as e:
        st.error(f"Erro Crítico: {e}")
