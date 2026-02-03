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
    """Remove acentos (Á -> A, ç -> c)."""
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def normalize_key_nospace(text):
    """
    Normaliza removendo espaços para dar match em nomes de arquivo.
    Ex: 'BarroAlto' == 'BARRO ALTO' -> Ambos viram 'BARROALTO'
    """
    if pd.isna(text): return ""
    text = remove_accents(str(text)).upper().strip()
    # Remove prefixos comuns antes de tirar os espaços
    prefixes = ["MUNICIPIO DE ", "PREFEITURA DE ", "PREFEITURA MUNICIPAL DE "]
    for p in prefixes:
        if text.startswith(p):
            text = text[len(p):]
    # Remove espaços
    return text.replace(" ", "").replace("_", "")

def normalize_key_standard(text):
    """Normalização padrão para busca de responsáveis (com espaços)."""
    if pd.isna(text): return ""
    text = remove_accents(str(text)).upper().strip()
    prefixes = ["MUNICIPIO DE ", "PREFEITURA DE ", "PREFEITURA MUNICIPAL DE "]
    for p in prefixes:
        if text.startswith(p):
            text = text[len(p):].strip()
    return text

# ================= 2. CARREGAMENTO DE DADOS =================

def gerar_modelo_responsaveis():
    data = {'Município': ['Goiânia', 'Anápolis'], 'Responsável': ['Prefeito A', 'Prefeito B']}
    return pd.DataFrame(data).to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

def gerar_modelo_pgfn():
    """Gera modelo do Extrato PGFN."""
    data = {
        'Arquivo': ['GO_CidadeExemplo_PGFN.pdf', 'TO_OutraCidade_PGFN.pdf'],
        'Identificador': ['123456789', '987654321'],
        'Modalidade': ['Parcelamento Lei 13.485', 'Transação Excepcional'],
        'Saldo (R$)': ['10000.50', '5000.00']
    }
    return pd.DataFrame(data).to_csv(index=False, sep=',', encoding='utf-8-sig').encode('utf-8-sig')

def carregar_responsaveis(arquivo):
    try:
        if arquivo.name.endswith('.csv'):
            try: df = pd.read_csv(arquivo, sep=';', encoding='utf-8-sig')
            except: 
                arquivo.seek(0)
                df = pd.read_csv(arquivo, sep=',', encoding='utf-8')
        else:
            df = pd.read_excel(arquivo)
        
        df.columns = [remove_accents(c).strip().lower() for c in df.columns]
        col_muni = next((c for c in df.columns if 'municipio' in c or 'cidade' in c), None)
        col_resp = next((c for c in df.columns if 'responsavel' in c or 'nome' in c), None)
        
        dic = {}
        if col_muni and col_resp:
            for _, row in df.iterrows():
                key = normalize_key_standard(row[col_muni])
                dic[key] = str(row[col_resp]).strip()
        return dic
    except: return {}

def carregar_pgfn_csv(arquivo):
    """
    Lê o CSV de Extrato PGFN e agrupa por município.
    Retorna: { 'BARROALTO': [ {row_data}, ... ] }
    """
    try:
        if arquivo.name.endswith('.csv'):
            df = pd.read_csv(arquivo)
        else:
            df = pd.read_excel(arquivo)
        
        # Identifica colunas (flexível)
        cols = {c.lower(): c for c in df.columns}
        col_arq = cols.get('arquivo', df.columns[0])
        col_id = next((c for c in df.columns if 'identificador' in c.lower() or 'processo' in c.lower()), None)
        col_mod = next((c for c in df.columns if 'modalidade' in c.lower()), None)
        col_val = next((c for c in df.columns if 'saldo' in c.lower() or 'valor' in c.lower()), None)

        dados_por_muni = {}

        for _, row in df.iterrows():
            # Extrai cidade do nome do arquivo (ex: GO_BarroAlto_...)
            nome_arq = str(row[col_arq])
            parts = nome_arq.split('_')
            
            if len(parts) >= 2:
                # Assume que a cidade é a segunda parte (parts[1])
                # Ex: GO (0), BarroAlto (1), ...
                cidade_raw = parts[1]
                key = normalize_key_nospace(cidade_raw)
                
                if key not in dados_por_muni:
                    dados_por_muni[key] = []
                
                valor = row[col_val]
                # Limpeza básica de valor se vier como string R$
                if isinstance(valor, str):
                    valor = valor.replace('R$', '').replace('.', '').replace(',', '.')
                
                try:
                    valor_float = float(valor)
                except:
                    valor_float = 0.0

                dados_por_muni[key].append({
                    'Processo': str(row[col_id]),
                    'Modalidade': str(row[col_mod]),
                    'Valor Original': valor_float,
                    'Fonte': 'PGFN CSV'
                })
        
        return dados_por_muni

    except Exception as e:
        st.error(f"Erro ao ler PGFN: {e}")
        return {}

def buscar_responsavel(muni_divida, db_resp):
    target = normalize_key_standard(muni_divida)
    if target in db_resp: return db_resp[target]
    # Busca aproximada
    for k in db_resp:
        if k.startswith(target) or target.startswith(k):
            return db_resp[k]
    return "PREFEITO(A) MUNICIPAL"

# ================= 3. MANIPULAÇÃO WORD =================

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

def adicionar_linha_tabela(table, orgao, modalidade, processo, valor, is_placeholder=False):
    row_cells = table.add_row().cells
    
    # --- Coluna 1: Órgão e Modalidade ---
    p1 = row_cells[0].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT # Alinhado à ESQUERDA
    r_org = p1.add_run(orgao)
    
    if modalidade and not is_placeholder:
        r_mod = p1.add_run(f"\n({modalidade})")
        r_mod.font.size = Pt(8) # Fonte menor para modalidade

    # --- Coluna 2: Processo ---
    p2 = row_cells[1].paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER # CENTRALIZADO
    p2.add_run(processo)

    # --- Coluna 3: Valor ---
    p3 = row_cells[2].paragraphs[0]
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER # CENTRALIZADO (Solicitado)
    p3.add_run(valor)

    # Ajuste geral de fonte
    for cell in row_cells:
        cell.vertical_alignment = 1
        for p in cell.paragraphs:
            if not p.runs: continue
            # Garante tamanho 10 para o texto principal, 8 já foi setado para mod
            if p.runs[0].font.size != Pt(8):
                p.runs[0].font.size = Pt(10)

def preencher_tabela(table, df_rfb_muni, lista_pgfn_muni):
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False 

    # Cabeçalho
    hdr_cells = table.rows[0].cells
    titulos = ['Órgão / Modalidade', 'Processo / Documento', 'Saldo em 31/12/2025']
    for i, titulo in enumerate(titulos):
        hdr_cells[i].text = titulo
        for p in hdr_cells[i].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(10)
            if not p.runs: p.add_run(titulo).font.bold = True

    # 1. Dados RFB (Excel)
    # Filtra o que NÃO é PGFN no Excel
    df_rfb = df_rfb_muni[~df_rfb_muni['Sistema'].astype(str).str.contains("PGFN", case=False, na=False)]
    
    if not df_rfb.empty:
        for _, row in df_rfb.iterrows():
            mod = str(row['Modalidade']) if pd.notna(row['Modalidade']) else ""
            adicionar_linha_tabela(table, "Receita Federal do Brasil", mod, str(row['Processo']), formatar_valor(row['Valor Original']))
    else:
        adicionar_linha_tabela(table, "Receita Federal do Brasil", "", "-", "-", is_placeholder=True)

    # 2. Dados PGFN (CSV Extrato - Prioridade)
    if lista_pgfn_muni:
        for item in lista_pgfn_muni:
            adicionar_linha_tabela(
                table, 
                "Procuradoria Geral da Fazenda Nacional", 
                item['Modalidade'], 
                item['Processo'], 
                formatar_valor(item['Valor Original'])
            )
    else:
        # Fallback: Tenta achar PGFN no Excel original se não tiver no CSV
        df_pgfn_excel = df_rfb_muni[df_rfb_muni['Sistema'].astype(str).str.contains("PGFN", case=False, na=False)]
        if not df_pgfn_excel.empty:
            for _, row in df_pgfn_excel.iterrows():
                mod = str(row['Modalidade']) if pd.notna(row['Modalidade']) else ""
                adicionar_linha_tabela(table, "Procuradoria Geral da Fazenda Nacional", mod, str(row['Processo']), formatar_valor(row['Valor Original']))
        else:
            # Vazio
            adicionar_linha_tabela(table, "Procuradoria Geral da Fazenda Nacional", "", "-", "-", is_placeholder=True)

def inserir_tabela_no_placeholder(doc, df_rfb, lista_pgfn, placeholder="{{TABELA}}"):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = ""
            table = doc.add_table(rows=1, cols=3)
            paragraph._p.addnext(table._tbl)
            preencher_tabela(table, df_rfb, lista_pgfn)
            return True
    return False

# ================= 4. INTERFACE =================
st.title("Gerador de Ofícios 3.0 (Integrado)")

# Área de Downloads de Modelos
with st.expander("📂 Baixar Modelos de Planilhas"):
    c1, c2 = st.columns(2)
    c1.download_button("📥 Modelo Responsáveis (CSV)", gerar_modelo_responsaveis(), "Modelo_Responsaveis.csv", "text/csv")
    c2.download_button("📥 Modelo Extrato PGFN (CSV)", gerar_modelo_pgfn(), "Modelo_Extrato_PGFN.csv", "text/csv")

col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. Dívidas RFB (Excel)", type=["xlsx"])
    uploaded_pgfn = st.file_uploader("4. Extrato PGFN (CSV/Excel)", type=["csv", "xlsx"])
with col2:
    uploaded_template = st.file_uploader("2. Modelo Word (.docx)", type=["docx"])
    uploaded_resp = st.file_uploader("3. Lista Responsáveis (CSV)", type=["csv"])

st.sidebar.header("Parâmetros")
num_inicial = st.sidebar.number_input("Nº Inicial", value=46)
ano_doc = st.sidebar.number_input("Ano", value=2026)

# ================= 5. PROCESSAMENTO =================
if st.button("🚀 Gerar Arquivos (ZIP)"):
    if not uploaded_excel or not uploaded_template:
        st.error("Arquivos 1 e 2 são obrigatórios.")
        st.stop()
    
    # Cargas
    db_resp = carregar_responsaveis(uploaded_resp) if uploaded_resp else {}
    db_pgfn = carregar_pgfn_csv(uploaded_pgfn) if uploaded_pgfn else {}
    
    if uploaded_pgfn:
        st.success(f"PGFN: Dados carregados para {len(db_pgfn)} municípios.")

    try:
        df = pd.read_excel(uploaded_excel, engine='openpyxl')
        df = df.dropna(subset=['Processo'])
        col_muni = 'Município' if 'Município' in df.columns else df.columns[0]
        df[col_muni] = df[col_muni].astype(str).str.strip()
        
        # Lista unificada de municípios
        munis_excel = set(df[col_muni].unique())
        # Mapeia chaves PGFN de volta para nomes legíveis é difícil, vamos iterar pelo Excel e tentar achar no PGFN
        # (Idealmente o Excel mestre tem todos os municipios)
        municipios = sorted(list(munis_excel))

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

                # Dados RFB
                df_rfb_muni = df[df[col_muni] == muni]
                
                # Dados PGFN (Busca por chave sem espaço)
                key_pgfn = normalize_key_nospace(muni)
                lista_pgfn = db_pgfn.get(key_pgfn, [])
                
                # Se não achou exato, tenta busca parcial (ex: BarroAlto vs BARROALTO)
                if not lista_pgfn:
                     for k in db_pgfn:
                         if key_pgfn in k or k in key_pgfn:
                             lista_pgfn = db_pgfn[k]
                             break

                # UF
                uf = "GO"
                if not df_rfb_muni.empty and 'Arquivo' in df_rfb_muni.columns:
                    try: 
                        parts = str(df_rfb_muni.iloc[0]['Arquivo']).split('-')
                        if len(parts) > 0 and len(parts[0].strip()) == 2: uf = parts[0].strip()
                    except: pass
                
                # Prefeito
                nome_pref = buscar_responsavel(muni, db_resp)
                if nome_pref == "PREFEITO(A) MUNICIPAL" and db_resp:
                    logs.append(f"⚠️ {muni}: Responsável não encontrado.")

                # Replaces
                replaces = {
                    "{{MUNICIPIO}}": muni.upper(),
                    "{{UF}}": uf,
                    "{{PREFEITO}}": nome_pref.upper(),
                    "{{NUM_OFICIO}}": f"{contador:03d}/{ano_doc}",
                    "{{DATA_EXTENSO}}": data_extenso
                }
                for k, v in replaces.items(): replace_everywhere(doc, k, v)
                
                # Tabela
                sucesso = inserir_tabela_no_placeholder(doc, df_rfb_muni, lista_pgfn, "{{TABELA}}")
                if not sucesso:
                    inserir_tabela_no_placeholder(doc, df_rfb_muni, lista_pgfn, "{{TABELA_DEBITOS}}")

                # Salva
                doc_io = io.BytesIO()
                doc.save(doc_io)
                fname = f"{contador:03d}-{ano_doc} - {uf} - {muni} - Saldo Divida RFB-PGFN.docx"
                zf.writestr(fname, doc_io.getvalue())
                
                contador += 1
                progress.progress((i+1)/len(municipios))
        
        st.success(f"✅ Sucesso! {len(municipios)} documentos gerados.")
        if logs:
            with st.expander("Alertas"):
                for l in logs: st.write(l)
        
        st.download_button("⬇️ Baixar ZIP", zip_buffer.getvalue(), f"Oficios_{datetime.now().strftime('%H%M')}.zip", "application/zip")

    except Exception as e:
        st.error(f"Erro: {e}")
