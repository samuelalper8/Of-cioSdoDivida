import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import os

# --- Configuração da Página ---
st.set_page_config(page_title="Gerador de Ofícios - ConPrev", layout="wide")

# --- 1. Base de Dados de Prefeitos (Manual) ---
# Dicionário formatado: Chave em MAIÚSCULO -> Nome do Prefeito
DB_PREFEITOS = {
    "AMARALINA": "Dásio Marques",
    "BALIZA": "Fernanda Nolasco",
    "BARRO ALTO": "Prof. Álvaro",
    "BELA VISTA DE GOIAS": "Nárcia Kelly",
    "BRAZABRANTES": "Jânio",
    "BURITI ALEGRE": "André de Sousa",
    "CAIAPONIA": "Argemiro Rodrigues",
    "CAIAPÔNIA": "Argemiro Rodrigues", # Variação com acento
    "CAMPINACU": "Dr. Douglas",
    "CAMPINAÇU": "Dr. Douglas",
    "CERES": "Inês Brito",
    "CORREGO DO OURO": "Lúcia Lolly",
    "CÓRREGO DO OURO": "Lúcia Lolly",
    "CORUMBA GOIAS": "Chico Vaca",
    "CORUMBÁ GOIÁS": "Chico Vaca",
    "CRISTALINA": "Daniel do Sindicato",
    "CRIXAS": "Dr. Carlos",
    "CRIXÁS": "Dr. Carlos",
    "GOIAS": "Prof. Anderson",
    "GOIÁS": "Prof. Anderson",
    "GOIATUBA": "Zezinho Vieira",
    "HIDROLINA": "Zica",
    "ITABERAI": "Wilian",
    "ITABERAÍ": "Wilian",
    "ITAPACI": "Mário Macaco",
    "JARAGUA": "Paulo Vitor",
    "JARAGUÁ": "Paulo Vitor",
    "MONTES CLAROS GOIAS": "Dr. Romer",
    "MONTES CLAROS GOIÁS": "Dr. Romer",
    "NOVO GAMA": "Carlinhos do Mangão",
    "NEROPOLIS": "Luiz Alberto Franco Araujo",
    "NERÓPOLIS": "Luiz Alberto Franco Araujo",
    "PARANAIGUARA": "Adalberto Amorim",
    "PEROLANDIA": "Grete",
    "PEROLÂNDIA": "Grete",
    "PILAR DE GOIAS": "Tiagão",
    "PILAR DE GOIÁS": "Tiagão",
    "PIRANHAS": "Marco Rogério",
    "RIANAPOLIS": "Zé Carlos",
    "RIANÁPOLIS": "Zé Carlos",
    "RIO QUENTE": "Ana Paula",
    "SAO FRANCISCO GOIAS": "Cleuton",
    "SÃO FRANCISCO GOIÁS": "Cleuton",
    "SAO LUIS MONTES BELOS": "Major Eldecírio",
    "SÃO LUÍS MONTES BELOS": "Major Eldecírio",
    "SERRANOPOLIS": "Tio Dé",
    "SERRANÓPOLIS": "Tio Dé",
    "TERESINA GOIAS": "Baiano",
    "TERESINA GOIÁS": "Baiano",
    "TRINDADE": "Marden Júnior",
    "UIRAPURU": "Elivan Carreiro",
    "ALCINOPOLIS": "Dalmy Crisóstomo",
    "ALCINÓPOLIS": "Dalmy Crisóstomo",
    "ANASTACIO": "Nildo Alves",
    "ANASTÁCIO": "Nildo Alves",
    "AQUIDAUANA": "Mauro Luiz Batista",
    "CHAPADAO DO SUL": "João Carlos Krug",
    "CHAPADÃO DO SUL": "João Carlos Krug",
    "COXIM": "Edilson Magro",
    "IGUATEMI": "Dr. Lídio",
    "JAPORA": "Paulo César",
    "JAPORÃ": "Paulo César",
    "JARAGUARI": "Edson Rodrigues",
    "SETE QUEDAS": "Chico Biasi",
    "SONORA": "Enelto Ramos",
    "TACURU": "Rogério Torquetti",
    "ALMAS": "Vagner",
    "BANDEIRANTES DO TOCANTINS": "Saulo Gonçalves Borges",
    "BARRA DO OURO": "Nélio",
    "BREJINHO DE NAZARE": "Miyuki",
    "BREJINHO DE NAZARÉ": "Miyuki",
    "CRISTALANDIA": "Wilson Junior Carvalho De Oliveira",
    "CRISTALÂNDIA": "Wilson Junior Carvalho De Oliveira",
    "GUARAI": "Fátima Coelho",
    "GUARAÍ": "Fátima Coelho",
    "JAU DO TOCANTINS": "Luciene Lourenco De Araujo",
    "JAÚ DO TOCANTINS": "Luciene Lourenco De Araujo",
    "LAGOA DA CONFUSAO": "Thiago Soares Carlos",
    "LAGOA DA CONFUSÃO": "Thiago Soares Carlos",
    "LAJEADO": "Júnior",
    "MAURILANDIA DO TOCANTINS": "Rafael",
    "MAURILÂNDIA DO TOCANTINS": "Rafael",
    "NATIVIDADE": "Dr. Thiago",
    "PALMEIRAS DO TOCANTINS": "Nalva",
    "PALMEIROPOLIS": "Bartolomeu",
    "PALMEIRÓPOLIS": "Bartolomeu",
    "PARAISO DO TOCANTINS": "Celso Morais",
    "PARAÍSO DO TOCANTINS": "Celso Morais",
    "PARANA": "Fabrício Viana",
    "PARANÃ": "Fabrício Viana",
    "PEDRO AFONSO": "Joaquim Pinheiro",
    "PEIXE": "Zé Augusto",
    "SANTA MARIA DO TOCANTINS": "Itamar",
    "SANTA RITA DO TOCANTINS": "Neila",
    "SAO VALERIO DA NATIVIDADE": "Prof. Olímpio",
    "SÃO VALÉRIO DA NATIVIDADE": "Prof. Olímpio",
    "SILVANOPOLIS": "Gernivon",
    "SILVANÓPOLIS": "Gernivon"
}

# --- 2. Função para Limpar Texto (CORREÇÃO DE ERRO) ---
def limpar_texto(texto):
    """
    Remove caracteres que o FPDF padrão não suporta (Latin-1).
    Substitui travessões, aspas curvas, etc.
    """
    if pd.isna(texto):
        return ""
    
    texto = str(texto)
    
    # Substituições manuais de caracteres problemáticos comuns
    substituicoes = {
        '–': '-',  # Travessão médio (En-dash) -> Hífen
        '—': '-',  # Travessão longo (Em-dash) -> Hífen
        '“': '"',  # Aspas curvas esquerda -> Aspas retas
        '”': '"',  # Aspas curvas direita -> Aspas retas
        "’": "'",  # Apóstrofo curvo -> Apóstrofo reto
        "‘": "'",
        '\u200b': '', # Espaço largura zero
        '\xa0': ' '   # Non-breaking space
    }
    
    for original, novo in substituicoes.items():
        texto = texto.replace(original, novo)
    
    # Garante codificação Latin-1 (substitui o que não conseguir por ?)
    return texto.encode('latin-1', 'replace').decode('latin-1')

# --- 3. Carregamento de Dados ---
@st.cache_data
def load_data():
    excel_path = "Relatorio_Dividas_RFB_Completo.xlsx"
    try:
        df = pd.read_excel(excel_path, engine='openpyxl')
    except Exception as e:
        st.error(f"Erro ao ler Excel: {e}")
        st.stop()
    
    df = df.dropna(subset=['Processo'])
    
    def extract_uf(arquivo_str):
        try:
            parts = str(arquivo_str).split('-')
            if len(parts) > 0 and len(parts[0].strip()) == 2:
                return parts[0].strip()
            return "GO"
        except:
            return "GO"

    if 'Arquivo' in df.columns:
        df['UF_EXTRAIDA'] = df['Arquivo'].apply(extract_uf)
    else:
        df['UF_EXTRAIDA'] = "GO"
        
    return df

df = load_data()

# --- 4. Classe PDF ---
class PDF(FPDF):
    def header(self):
        if os.path.exists("PapelTimbrado_2026.jpg"):
            self.image("PapelTimbrado_2026.jpg", x=0, y=0, w=210)
        self.ln(50)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)

# --- 5. Interface Streamlit ---
st.title("📄 Gerador de Ofícios - ConPrev")

# Sidebar
st.sidebar.header("Selecione o Cliente")

coluna_municipio = 'Município'
if coluna_municipio in df.columns:
    lista_municipios = sorted(df[coluna_municipio].astype(str).unique())
    municipio_selecionado = st.sidebar.selectbox("Município", lista_municipios)
    
    df_filtered = df[df[coluna_municipio] == municipio_selecionado]
else:
    st.error("Coluna Município não encontrada.")
    st.stop()

uf_atual = df_filtered['UF_EXTRAIDA'].iloc[0] if not df_filtered.empty else "GO"

st.sidebar.markdown("---")
st.sidebar.header("Dados do Ofício")
num_oficio = st.sidebar.text_input("Número do Ofício", "00023")

# Lógica Automática de Prefeitos
nome_padrao = DB_PREFEITOS.get(municipio_selecionado.upper(), "")
nome_prefeito = st.sidebar.text_input("Nome do Prefeito", value=nome_padrao)

if not nome_padrao:
    st.sidebar.warning(f"Prefeito de '{municipio_selecionado}' não encontrado na base. Digite manualmente.")

# Data Automática
data_raw = datetime.now()
meses = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}
data_formatada = f"{data_raw.day} de {meses[data_raw.month]} de {data_raw.year}"

# Visualização
st.subheader(f"Débitos de {municipio_selecionado} ({uf_atual})")
st.dataframe(df_filtered[['Processo', 'Modalidade', 'Sistema', 'Valor Original']], use_container_width=True)

# --- 6. Geração do PDF ---
if st.button("Gerar PDF e Baixar"):
    if df_filtered.empty:
        st.warning("Sem dados.")
    else:
        pdf = PDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=25)
        pdf.set_font("Arial", size=11)
        
        # --- APLICAÇÃO DA LIMPEZA DE TEXTO (limpar_texto) ---
        
        # Data
        pdf.set_xy(120, 55)
        pdf.cell(0, 10, limpar_texto(f"Goiânia, {data_formatada}."), ln=True)
        
        # Ofício
        pdf.set_xy(20, 70)
        pdf.set_font("Arial", 'B', size=11)
        pdf.cell(0, 5, limpar_texto(f"Ofício DFF nº {num_oficio}/2026"), ln=True)
        pdf.ln(5)
        
        # Destinatário
        pdf.set_font("Arial", size=11)
        destinatario = f"EXCELENTÍSSIMO SENHOR\n{nome_prefeito.upper()}\nPREFEITO MUNICIPAL DE {municipio_selecionado} – {uf_atual}"
        pdf.multi_cell(0, 5, limpar_texto(destinatario))
        pdf.ln(10)
        
        # Texto Intro
        texto_intro = (
            "Assunto: Ficam apresentados os valores e a documentação comprobatória dos saldos de débitos "
            "existentes em 31 de dezembro de 2025, destinados à composição do Balanço Patrimonial.\n\n"
            "Senhor Prefeito,\n\n"
            "Ao tempo em que lhe cumprimento, na qualidade de assessoria do Município para assuntos relacionados "
            "a atos de pessoal e ao fisco federal, no âmbito das ações de conformidade administrativa, venho, por meio "
            "do presente, apresentar os valores e a documentação requisitados por esta assessoria especializada, "
            "referentes aos saldos de débitos destinados à composição do Balanço Patrimonial.\n\n"
            "Nesse contexto, discriminam-se abaixo o órgão de origem, o número do processo e os respectivos "
            "valores apurados em 31/12/2025:"
        )
        pdf.multi_cell(0, 6, limpar_texto(texto_intro))
        pdf.ln(5)
        
        # Tabela
        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(240, 240, 240)
        pdf.cell(90, 8, limpar_texto("Processo / Documento / Sistema"), 1, 0, 'C', fill=True)
        pdf.cell(50, 8, limpar_texto("Saldo Devedor (R$)"), 1, 1, 'C', fill=True)
        
        pdf.set_font("Arial", size=10)
        
        for index, row in df_filtered.iterrows():
            processo = str(row['Processo'])
            sistema = str(row['Sistema']) if pd.notna(row['Sistema']) else ""
            
            val = row['Valor Original']
            if isinstance(val, (int, float)):
                valor_str = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                valor_str = str(val)

            conteudo_processo = f"{processo}\n({sistema})"
            
            # Altura dinâmica
            x_start = pdf.get_x()
            y_start = pdf.get_y()
            
            pdf.multi_cell(90, 6, limpar_texto(conteudo_processo), border=1)
            
            y_end = pdf.get_y()
            row_height = y_end - y_start
            
            pdf.set_xy(x_start + 90, y_start)
            pdf.cell(50, row_height, limpar_texto(valor_str), border=1, ln=1, align='R')
            
        pdf.ln(5)
        
        # Encerramento
        texto_final = (
            "Solicita-se, por oportuno, que a referida documentação seja encaminhada ao setor contábil, "
            "a fim de que sejam adotadas as providências e registros contábeis cabíveis.\n\n"
            "Esta consultoria agradece a confiança depositada e permanece à disposição para quaisquer "
            "esclarecimentos adicionais.\n\n"
            "Atenciosamente,"
        )
        pdf.multi_cell(0, 6, limpar_texto(texto_final))
        pdf.ln(15)
        
        # Assinaturas
        y_assinaturas = pdf.get_y()
        
        pdf.set_xy(20, y_assinaturas)
        pdf.cell(80, 5, limpar_texto("Rubens Pires Malaquias"), 0, 1, 'C')
        pdf.cell(80, 5, limpar_texto("Diretor Técnico e Consultor"), 0, 1, 'C')
        pdf.cell(80, 5, limpar_texto("CRA/GO 6-007-48"), 0, 0, 'C')
        
        pdf.set_xy(110, y_assinaturas)
        pdf.cell(80, 5, limpar_texto("Glayzer Antônio Gomes da Silva"), 0, 1, 'C')
        pdf.cell(80, 5, limpar_texto("Advogado Especialista"), 0, 1, 'C')
        
        nome_arquivo_pdf = f"Oficio_{municipio_selecionado}_{num_oficio}.pdf"
        pdf.output(nome_arquivo_pdf)
        
        with open(nome_arquivo_pdf, "rb") as f:
            st.download_button(
                label="⬇️ Baixar Ofício em PDF",
                data=f,
                file_name=nome_arquivo_pdf,
                mime="application/pdf"
            )
