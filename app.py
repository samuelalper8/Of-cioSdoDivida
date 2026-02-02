import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import os

# --- Configuração da Página ---
st.set_page_config(page_title="Gerador de Ofícios - ConPrev", layout="wide")

# --- 1. Carregamento e Tratamento de Dados (VERSÃO EXCEL) ---
@st.cache_data
def load_data():
    # Nome do arquivo Excel
    excel_path = "Relatorio_Dividas_RFB_Completo.xlsx"
    
    try:
        # Lê o Excel diretamente. 
        # sheet_name=0 pega a primeira aba. Se tiver outro nome, ajuste aqui.
        df = pd.read_excel(excel_path, engine='openpyxl')
    except FileNotFoundError:
        st.error(f"Arquivo '{excel_path}' não encontrado. Verifique se o nome está correto e se o arquivo está na pasta.")
        st.stop()
    except Exception as e:
        st.error(f"Erro ao ler o Excel: {e}")
        st.stop()
    
    # Remove linhas onde o 'Processo' está vazio (linhas em branco ou totais)
    df = df.dropna(subset=['Processo'])
    
    # Extrai UF da coluna 'Arquivo' (se existir) para garantir o cabeçalho correto
    def extract_uf(arquivo_str):
        try:
            parts = str(arquivo_str).split('-')
            if len(parts) > 0 and len(parts[0].strip()) == 2:
                return parts[0].strip()
            return "GO"
        except:
            return "GO"

    # Verifica se a coluna Arquivo existe antes de tentar extrair
    if 'Arquivo' in df.columns:
        df['UF_EXTRAIDA'] = df['Arquivo'].apply(extract_uf)
    else:
        df['UF_EXTRAIDA'] = "GO" # Valor padrão se não tiver a coluna
    
    return df

df = load_data()

# --- 2. Classe PDF Personalizada ---
class PDF(FPDF):
    def header(self):
        if os.path.exists("PapelTimbrado_2026.jpg"):
            self.image("PapelTimbrado_2026.jpg", x=0, y=0, w=210)
        self.ln(50)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)

# --- 3. Interface do Streamlit ---
st.title("📄 Gerador de Ofícios - Leitura Excel (.xlsx)")

# Filtros Laterais
st.sidebar.header("Selecione o Cliente")

# Verifica se a coluna Município existe (baseado no seu arquivo 'Organizado')
coluna_municipio = 'Município' # Ajuste aqui se o nome no Excel for diferente (ex: 'MUNICIPIO')

if coluna_municipio in df.columns:
    lista_municipios = sorted(df[coluna_municipio].astype(str).unique())
    municipio_selecionado = st.sidebar.selectbox("Município", lista_municipios)
    
    # Filtra o DataFrame
    df_filtered = df[df[coluna_municipio] == municipio_selecionado]
else:
    st.error(f"A coluna '{coluna_municipio}' não foi encontrada no Excel.")
    st.stop()

# Pega UF
uf_atual = df_filtered['UF_EXTRAIDA'].iloc[0] if not df_filtered.empty else "GO"

# Inputs Manuais
st.sidebar.markdown("---")
st.sidebar.header("Dados do Ofício")
num_oficio = st.sidebar.text_input("Número do Ofício", "00023")
nome_prefeito = st.sidebar.text_input("Nome do Prefeito", "NOME DO PREFEITO AQUI")

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

# --- 4. Geração do PDF ---
if st.button("Gerar PDF e Baixar"):
    if df_filtered.empty:
        st.warning("Sem dados para gerar.")
    else:
        pdf = PDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=25)
        pdf.set_font("Arial", size=11)
        
        # Data
        pdf.set_xy(120, 55)
        pdf.cell(0, 10, f"Goiânia, {data_formatada}.", ln=True)
        
        # Ofício
        pdf.set_xy(20, 70)
        pdf.set_font("Arial", 'B', size=11)
        pdf.cell(0, 5, f"Ofício DFF nº {num_oficio}/2026", ln=True)
        pdf.ln(5)
        
        # Destinatário
        pdf.set_font("Arial", size=11)
        pdf.multi_cell(0, 5, f"EXCELENTÍSSIMO SENHOR\n{nome_prefeito.upper()}\nPREFEITO MUNICIPAL DE {municipio_selecionado} – {uf_atual}")
        pdf.ln(10)
        
        # Texto Padrão
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
        pdf.multi_cell(0, 6, texto_intro)
        pdf.ln(5)
        
        # Tabela
        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(240, 240, 240)
        pdf.cell(90, 8, "Processo / Documento / Sistema", 1, 0, 'C', fill=True)
        pdf.cell(50, 8, "Saldo Devedor (R$)", 1, 1, 'C', fill=True)
        
        pdf.set_font("Arial", size=10)
        
        for index, row in df_filtered.iterrows():
            processo = str(row['Processo'])
            sistema = str(row['Sistema']) if pd.notna(row['Sistema']) else ""
            
            # Formatação do Valor (Assume que no Excel já está texto ou número)
            val = row['Valor Original']
            if isinstance(val, (int, float)):
                valor_str = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                valor_str = str(val)

            conteudo_processo = f"{processo}\n({sistema})"
            
            x_start = pdf.get_x()
            y_start = pdf.get_y()
            
            pdf.multi_cell(90, 6, conteudo_processo, border=1)
            
            y_end = pdf.get_y()
            row_height = y_end - y_start
            
            pdf.set_xy(x_start + 90, y_start)
            pdf.cell(50, row_height, valor_str, border=1, ln=1, align='R')
            
        pdf.ln(5)
        
        # Encerramento
        texto_final = (
            "Solicita-se, por oportuno, que a referida documentação seja encaminhada ao setor contábil, "
            "a fim de que sejam adotadas as providências e registros contábeis cabíveis.\n\n"
            "Esta consultoria agradece a confiança depositada e permanece à disposição para quaisquer "
            "esclarecimentos adicionais.\n\n"
            "Atenciosamente,"
        )
        pdf.multi_cell(0, 6, texto_final)
        pdf.ln(15)
        
        # Assinaturas
        y_assinaturas = pdf.get_y()
        
        pdf.set_xy(20, y_assinaturas)
        pdf.cell(80, 5, "Rubens Pires Malaquias", 0, 1, 'C')
        pdf.cell(80, 5, "Diretor Técnico e Consultor", 0, 1, 'C')
        pdf.cell(80, 5, "CRA/GO 6-007-48", 0, 0, 'C')
        
        pdf.set_xy(110, y_assinaturas)
        pdf.cell(80, 5, "Glayzer Antônio Gomes da Silva", 0, 1, 'C')
        pdf.cell(80, 5, "Advogado Especialista", 0, 1, 'C')
        
        nome_arquivo_pdf = f"Oficio_{municipio_selecionado}_{num_oficio}.pdf"
        pdf.output(nome_arquivo_pdf)
        
        with open(nome_arquivo_pdf, "rb") as f:
            st.download_button(
                label="⬇️ Baixar Ofício em PDF",
                data=f,
                file_name=nome_arquivo_pdf,
                mime="application/pdf"
            )
