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
    """Substitui texto em parágrafos, tabelas e cabeçalhos."""
    def repl(par):
        if old in par.text:
            # Tenta substituir preservando formatação (runs)
            for run in par.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)
            # Fallback
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
    """Preenche a tabela com dados."""
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False 

    # Cabeçalho
    hdr_cells = table.rows[0].cells
    titulos = ['Órgão', 'Processo / Documento', 'Saldo em 31/12/2025']
    for i, titulo in enumerate(titulos):
        hdr_cells[i].text = titulo
        for p in hdr_cells[i].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(10)
            if not p.runs: 
                 p.add_run(titulo).font.bold = True

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

def inserir_tabela_no_placeholder(doc, df_municipio, placeholder="{{TABELA}}"):
    """Substitui o placeholder pela tabela no local exato."""
    found = False
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = ""
            table = doc.add_table(rows=1, cols=3)
            paragraph._p.addnext(table._tbl)
            preencher_tabela(table, df_municipio)
            found = True
            break
    return found

# ================= 3. INTERFACE =================
st.title("Gerador de Ofícios - Saldo Dívida RFB")
st.markdown("Faça upload da **Planilha** e do **Modelo Word**.")

col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. Planilha Excel (Dados)", type=["xlsx"])

with col2:
    uploaded_template = st.file_uploader("2. Modelo Word (.docx)", type=["docx"])

st.sidebar.header("Configuração")
num_inicial = st.sidebar.number_input("Número Inicial", value=46, step=1)
ano_doc = st.sidebar.number_input("Ano", value=2026)

with st.expander("ℹ️ Placeholders e Dicas"):
    st.markdown("""
    * **`{{TABELA}}`**: Use este placeholder em uma linha limpa para inserir a tabela.
    * O sistema preenche automaticamente: `{{MUNICIPIO}}`, `{{UF}}`, `{{PREFEITO}}`, `{{NUM_OFICIO}}`.
    """)

# ================= 4. PROCESSAMENTO =================
if st.button("🚀 Gerar Arquivos (ZIP)"):
    if not uploaded_excel or not uploaded_template:
        st.error("Uploads obrigatórios faltando!")
        st.stop()

    try:
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
                
                nome_pref = DB_PREFEITOS.get(muni.upper(), "PREFEITO(A) MUNICIPAL")
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
                    logs.append(f"⚠️ {muni}: Placeholder {{TABELA}} não encontrado. Tabela foi pro final.")
                    table_fallback = doc.add_table(rows=1, cols=3)
                    preencher_tabela(table_fallback, df_muni)

                # Salva no ZIP
                doc_io = io.BytesIO()
                doc.save(doc_io)
                
                # --- NOME DO ARQUIVO ATUALIZADO ---
                nome_zip = f"{contador:03d}-{ano_doc} - {uf} - {muni} - Saldo Divida RFB-PGFN.docx"
                zf.writestr(nome_zip, doc_io.getvalue())
                
                contador += 1
                progress.progress((i+1)/len(municipios))
        
        st.success(f"✅ Processamento concluído! {len(municipios)} ofícios gerados.")
        
        if logs:
            with st.expander("⚠️ Alertas de Formatação"):
                for log in logs: st.write(log)

        st.download_button("⬇️ Baixar ZIP (Nomes Atualizados)", zip_buffer.getvalue(), 
                           file_name=f"Oficios_SaldoDivida_{datetime.now().strftime('%H%M')}.zip", 
                           mime="application/zip")

    except Exception as e:
        st.error(f"Erro: {e}")
