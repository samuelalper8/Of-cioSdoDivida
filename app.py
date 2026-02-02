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

# ================= 1. BASE DE DADOS DE PREFEITOS (EMBUTIDA) =================
# O sistema busca pela chave em MAIÚSCULO.
DB_PREFEITOS = {
    "AMARALINA": "Dásio Marques",
    "BALIZA": "Fernanda Nolasco",
    "BARRO ALTO": "Prof. Álvaro",
    "BELA VISTA DE GOIAS": "Nárcia Kelly",
    "BRAZABRANTES": "Jânio",
    "BURITI ALEGRE": "André de Sousa",
    "CAIAPÔNIA": "Argemiro Rodrigues",
    "CAIAPONIA": "Argemiro Rodrigues",
    "CAMPINAÇU": "Dr. Douglas",
    "CAMPINACU": "Dr. Douglas",
    "CERES": "Inês Brito",
    "CÓRREGO DO OURO": "Lúcia Lolly",
    "CORREGO DO OURO": "Lúcia Lolly",
    "CORUMBÁ GOIÁS": "Chico Vaca",
    "CORUMBA GOIAS": "Chico Vaca",
    "CRISTALINA": "Daniel do Sindicato",
    "CRIXÁS": "Dr. Carlos",
    "CRIXAS": "Dr. Carlos",
    "GOIÁS": "Prof. Anderson",
    "GOIAS": "Prof. Anderson",
    "GOIATUBA": "Zezinho Vieira",
    "HIDROLINA": "Zica",
    "ITABERAÍ": "Wilian",
    "ITABERAI": "Wilian",
    "ITAPACI": "Mário Macaco",
    "JARAGUÁ": "Paulo Vitor",
    "JARAGUA": "Paulo Vitor",
    "MONTES CLAROS GOIÁS": "Dr. Romer",
    "MONTES CLAROS GOIAS": "Dr. Romer",
    "NOVO GAMA": "Carlinhos do Mangão",
    "NERÓPOLIS": "Luiz Alberto Franco Araujo",
    "NEROPOLIS": "Luiz Alberto Franco Araujo",
    "PARANAIGUARA": "Adalberto Amorim",
    "PEROLÂNDIA": "Grete",
    "PEROLANDIA": "Grete",
    "PILAR DE GOIÁS": "Tiagão",
    "PILAR DE GOIAS": "Tiagão",
    "PIRANHAS": "Marco Rogério",
    "RIANÁPOLIS": "Zé Carlos",
    "RIANAPOLIS": "Zé Carlos",
    "RIO QUENTE": "Ana Paula",
    "SÃO FRANCISCO GOIÁS": "Cleuton",
    "SAO FRANCISCO GOIAS": "Cleuton",
    "SÃO LUÍS MONTES BELOS": "Major Eldecírio",
    "SAO LUIS MONTES BELOS": "Major Eldecírio",
    "SERRANÓPOLIS": "Tio Dé",
    "SERRANOPOLIS": "Tio Dé",
    "TERESINA GOIÁS": "Baiano",
    "TERESINA GOIAS": "Baiano",
    "TRINDADE": "Marden Júnior",
    "UIRAPURU": "Elivan Carreiro",
    "ALCINÓPOLIS": "Dalmy Crisóstomo",
    "ALCINOPOLIS": "Dalmy Crisóstomo",
    "ANASTÁCIO": "Nildo Alves",
    "ANASTACIO": "Nildo Alves",
    "AQUIDAUANA": "Mauro Luiz Batista",
    "CHAPADÃO DO SUL": "João Carlos Krug",
    "CHAPADAO DO SUL": "João Carlos Krug",
    "COXIM": "Edilson Magro",
    "IGUATEMI": "Dr. Lídio",
    "JAPORÃ": "Paulo César",
    "JAPORA": "Paulo César",
    "JARAGUARI": "Edson Rodrigues",
    "SETE QUEDAS": "Chico Biasi",
    "SONORA": "Enelto Ramos",
    "TACURU": "Rogério Torquetti",
    "ALMAS": "Vagner",
    "BANDEIRANTES DO TOCANTINS": "Saulo Gonçalves Borges",
    "BARRA DO OURO": "Nélio",
    "BREJINHO DE NAZARÉ": "Miyuki",
    "BREJINHO DE NAZARE": "Miyuki",
    "CRISTALÂNDIA": "Wilson Junior Carvalho De Oliveira",
    "CRISTALANDIA": "Wilson Junior Carvalho De Oliveira",
    "GUARAÍ": "Fátima Coelho",
    "GUARAI": "Fátima Coelho",
    "JAÚ DO TOCANTINS": "Luciene Lourenco De Araujo",
    "JAU DO TOCANTINS": "Luciene Lourenco De Araujo",
    "LAGOA DA CONFUSÃO": "Thiago Soares Carlos",
    "LAGOA DA CONFUSAO": "Thiago Soares Carlos",
    "LAJEADO": "Júnior",
    "MAURILÂNDIA DO TOCANTINS": "Rafael",
    "MAURILANDIA DO TOCANTINS": "Rafael",
    "NATIVIDADE": "Dr. Thiago",
    "PALMEIRAS DO TOCANTINS": "Nalva",
    "PALMEIRÓPOLIS": "Bartolomeu",
    "PALMEIROPOLIS": "Bartolomeu",
    "PARAÍSO DO TOCANTINS": "Celso Morais",
    "PARAISO DO TOCANTINS": "Celso Morais",
    "PARANÃ": "Fabrício Viana",
    "PARANA": "Fabrício Viana",
    "PEDRO AFONSO": "Joaquim Pinheiro",
    "PEIXE": "Zé Augusto",
    "SANTA MARIA DO TOCANTINS": "Itamar",
    "SANTA RITA DO TOCANTINS": "Neila",
    "SÃO VALÉRIO DA NATIVIDADE": "Prof. Olímpio",
    "SAO VALERIO DA NATIVIDADE": "Prof. Olímpio",
    "SILVANÓPOLIS": "Gernivon",
    "SILVANOPOLIS": "Gernivon"
}

# ================= 2. FUNÇÕES DE MANIPULAÇÃO WORD =================

def replace_everywhere(doc: Document, old: str, new: str) -> None:
    """Substitui texto preservando a formatação o máximo possível."""
    for p in doc.paragraphs:
        if old in p.text:
            p.text = p.text.replace(old, new)
            
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old in p.text:
                        p.text = p.text.replace(old, new)
                        
    for s in doc.sections:
        for h in [s.header, s.first_page_header, s.footer, s.first_page_footer]:
            if h:
                for p in h.paragraphs:
                    if old in p.text:
                        p.text = p.text.replace(old, new)

def mover_tabela_para_placeholder(doc, table, placeholder_text):
    """Move a tabela para o local do placeholder."""
    target_p = None
    for p in doc.paragraphs:
        if placeholder_text in p.text:
            target_p = p
            break
    
    if target_p:
        target_p._p.addnext(table._tbl)
        target_p.text = "" 
        return True
    return False

def criar_tabela_divida(doc, df_municipio):
    """Cria a tabela de débitos."""
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False 

    # Cabeçalho
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Órgão'
    hdr_cells[1].text = 'Processo / Documento'
    hdr_cells[2].text = 'Saldo em 31/12/2025'
    
    for cell in hdr_cells:
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(10)

    # Dados
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
                
    return table

# ================= 3. INTERFACE =================
st.title("Gerador de Ofícios em Lote")
st.markdown("Faça upload da **Planilha Excel** e do **Modelo Word**.")

col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. Planilha Excel (Dados)", type=["xlsx"])

with col2:
    uploaded_template = st.file_uploader("2. Modelo Word (.docx)", type=["docx"])

st.sidebar.header("Configuração")
num_inicial = st.sidebar.number_input("Número Inicial", value=1, step=1)
ano_doc = st.sidebar.number_input("Ano", value=2026)

with st.expander("ℹ️ Placeholders Obrigatórios no Word"):
    st.markdown("""
    * `{{MUNICIPIO}}`
    * `{{PREFEITO}}` (Será preenchido automaticamente pela lista)
    * `{{UF}}`
    * `{{NUM_OFICIO}}`
    * `{{DATA_EXTENSO}}`
    * **`{{TABELA}}`** (Em uma linha vazia, onde entra a tabela)
    """)

# ================= 4. PROCESSAMENTO =================
if st.button("🚀 Gerar Arquivos (ZIP)"):
    if not uploaded_excel or not uploaded_template:
        st.error("Por favor, faça upload dos dois arquivos.")
        st.stop()

    try:
        df = pd.read_excel(uploaded_excel, engine='openpyxl')
        df = df.dropna(subset=['Processo'])
        
        col_municipio = 'Município' if 'Município' in df.columns else df.columns[0]
        # Garante que é string e remove espaços extras
        df[col_municipio] = df[col_municipio].astype(str).str.strip()
        municipios = sorted(df[col_municipio].unique())
        
        zip_buffer = io.BytesIO()
        contador = num_inicial
        
        hoje = datetime.now()
        meses = {1:"janeiro", 2:"fevereiro", 3:"março", 4:"abril", 5:"maio", 6:"junho",
                 7:"julho", 8:"agosto", 9:"setembro", 10:"outubro", 11:"novembro", 12:"dezembro"}
        data_extenso = f"Goiânia, {hoje.day} de {meses[hoje.month]} de {hoje.year}."

        progress = st.progress(0)
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, muni in enumerate(municipios):
                # Carrega modelo limpo
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
                
                # --- BUSCA PREFEITO NA LISTA EMBUTIDA ---
                # Converte o nome da cidade para maiúsculo para buscar no dicionário
                nome_pref = DB_PREFEITOS.get(muni.upper(), "PREFEITO(A) MUNICIPAL")
                
                num_fmt = f"{contador:03d}/{ano_doc}"
                
                # Substituições
                replaces = {
                    "{{MUNICIPIO}}": muni.upper(),
                    "{{UF}}": uf,
                    "{{PREFEITO}}": nome_pref.upper(), # Coloca o nome em maiúsculo
                    "{{NUM_OFICIO}}": num_fmt,
                    "{{DATA_EXTENSO}}": data_extenso
                }
                
                for k, v in replaces.items():
                    replace_everywhere(doc, k, v)
                
                # Tabela
                tabela = criar_tabela_divida(doc, df_muni)
                sucesso = mover_tabela_para_placeholder(doc, tabela, "{{TABELA}}")
                if not sucesso:
                    mover_tabela_para_placeholder(doc, tabela, "{{TABELA_DEBITOS}}")

                # Salva
                doc_io = io.BytesIO()
                doc.save(doc_io)
                
                nome_zip = f"{contador:03d}-{ano_doc} - {muni}.docx"
                zf.writestr(nome_zip, doc_io.getvalue())
                
                contador += 1
                progress.progress((i+1)/len(municipios))
                
        st.success(f"✅ Sucesso! {len(municipios)} ofícios gerados.")
        st.download_button("⬇️ Baixar ZIP", zip_buffer.getvalue(), 
                           file_name=f"Oficios_Com_Prefeitos_{datetime.now().strftime('%H%M')}.zip", 
                           mime="application/zip")

    except Exception as e:
        st.error(f"Erro: {e}")
