import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import os

# --- Configuração da Página ---
st.set_page_config(page_title="Gerador de Ofícios - ConPrev", layout="wide")

# --- 1. Carregamento e Tratamento de Dados ---
@st.cache_data
def load_data():
    # Nome exato do novo arquivo enviado
    csv_path = "Relatorio_Dividas_RFB_Completo.xlsx - Organizado.csv"
    
    try:
        df = pd.read_csv(csv_path)
    except FileNotFoundError:
        st.error(f"Arquivo '{csv_path}' não encontrado. Verifique se ele está na mesma pasta do script.")
        st.stop()
    
    # Remove linhas onde o 'Processo' está vazio (geralmente linhas de total ou espaçamento)
    df = df.dropna(subset=['Processo'])
    
    # Extrai UF da coluna 'Arquivo' (ex: "GO - BARRO ALTO...") para usar no cabeçalho
    def extract_uf(arquivo_str):
        try:
            # Pega os 2 primeiros caracteres antes do hífen
            parts = str(arquivo_str).split('-')
            if len(parts) > 0 and len(parts[0].strip()) == 2:
                return parts[0].strip()
            return "GO" # Padrão
        except:
            return "GO"

    df['UF_EXTRAIDA'] = df['Arquivo'].apply(extract_uf)
    
    return df

df = load_data()

# --- 2. Classe PDF Personalizada ---
class PDF(FPDF):
    def header(self):
        # Insere o Papel Timbrado como fundo
        if os.path.exists("PapelTimbrado_2026.jpg"):
            # Ajuste x, y, w conforme necessário (A4 = 210mm largura)
            self.image("PapelTimbrado_2026.jpg", x=0, y=0, w=210)
        self.ln(50) # Espaço para o cabeçalho do timbrado

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        # self.cell(0, 10, 'Página ' + str(self.page_no()), 0, 0, 'C')

# --- 3. Interface do Streamlit ---
st.title("📄 Gerador de Ofícios - Base Atualizada")

# Filtros Laterais
st.sidebar.header("Selecione o Cliente")

# Obtém lista única de municípios ordenados
lista_municipios = sorted(df['Município'].dropna().unique())
municipio_selecionado = st.sidebar.selectbox("Município", lista_municipios)

# Filtra o DataFrame
df_filtered = df[df['Município'] == municipio_selecionado]

# Tenta pegar a UF da primeira linha filtrada
uf_atual = df_filtered['UF_EXTRAIDA'].iloc[0] if not df_filtered.empty else "GO"

# Inputs Manuais
st.sidebar.markdown("---")
st.sidebar.header("Dados do Ofício")
num_oficio = st.sidebar.text_input("Número do Ofício (Ex: 00023)", "00023")
nome_prefeito = st.sidebar.text_input("Nome do Prefeito", "NOME DO PREFEITO AQUI")

# Data Automática em Português
data_raw = datetime.now()
meses = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}
data_formatada = f"{data_raw.day} de {meses[data_raw.month]} de {data_raw.year}"

# Visualização da Tabela no Streamlit
st.subheader(f"Débitos de {municipio_selecionado} ({uf_atual})")
st.info(f"Total de registros encontrados: {len(df_filtered)}")
st.dataframe(df_filtered[['Processo', 'Modalidade', 'Sistema', 'Valor Original']], use_container_width=True)

# --- 4. Geração do PDF ---
if st.button("Gerar PDF e Baixar"):
    if df_filtered.empty:
        st.warning("Não há dados para gerar o PDF deste município.")
    else:
        pdf = PDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=25)
        
        # Configurações de Fonte
        pdf.set_font("Arial", size=11)
        
        # --- Texto do Ofício ---
        # Data
        pdf.set_xy(120, 55)
        pdf.cell(0, 10, f"Goiânia, {data_formatada}.", ln=True)
        
        # Número do Ofício
        pdf.set_xy(20, 70)
        pdf.set_font("Arial", 'B', size=11)
        pdf.cell(0, 5, f"Ofício DFF nº {num_oficio}/2026", ln=True)
        pdf.ln(5)
        
        # Destinatário
        pdf.set_font("Arial", size=11)
        pdf.multi_cell(0, 5, f"EXCELENTÍSSIMO SENHOR\n{nome_prefeito.upper()}\nPREFEITO MUNICIPAL DE {municipio_selecionado} – {uf_atual}")
        pdf.ln(10)
        
        # Corpo do Texto (Padrão DOCX)
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
        
        # --- Tabela ---
        # Cabeçalho
        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(240, 240, 240) # Cinza claro
        pdf.cell(90, 8, "Processo / Documento / Sistema", 1, 0, 'C', fill=True)
        pdf.cell(50, 8, "Saldo Devedor (R$)", 1, 1, 'C', fill=True)
        
        # Linhas
        pdf.set_font("Arial", size=10)
        
        for index, row in df_filtered.iterrows():
            # Dados das colunas do novo Excel
            processo = str(row['Processo'])
            sistema = str(row['Sistema']) if pd.notna(row['Sistema']) else ""
            valor_str = str(row['Valor Original']) # Usa o valor já formatado "1.000,00"
            
            conteudo_processo = f"{processo}\n({sistema})"
            
            # Controle de altura da célula (multiline)
            x_start = pdf.get_x()
            y_start = pdf.get_y()
            
            pdf.multi_cell(90, 6, conteudo_processo, border=1)
            
            y_end = pdf.get_y()
            row_height = y_end - y_start
            
            # Célula do Valor (alinhada à direita)
            pdf.set_xy(x_start + 90, y_start)
            pdf.cell(50, row_height, valor_str, border=1, ln=1, align='R')
            
        pdf.ln(5)
        
        # --- Encerramento ---
        texto_final = (
            "Solicita-se, por oportuno, que a referida documentação seja encaminhada ao setor contábil, "
            "a fim de que sejam adotadas as providências e registros contábeis cabíveis.\n\n"
            "Esta consultoria agradece a confiança depositada e permanece à disposição para quaisquer "
            "esclarecimentos adicionais.\n\n"
            "Atenciosamente,"
        )
        pdf.multi_cell(0, 6, texto_final)
        pdf.ln(15)
        
        # --- Assinaturas ---
        y_assinaturas = pdf.get_y()
        
        # Coluna 1: Rubens
        pdf.set_xy(20, y_assinaturas)
        # pdf.image("assinatura_rubens.png", x=35, y=y_assinaturas-15, w=30) # Descomente se tiver a imagem
        pdf.cell(80, 5, "Rubens Pires Malaquias", 0, 1, 'C')
        pdf.cell(80, 5, "Diretor Técnico e Consultor", 0, 1, 'C')
        pdf.cell(80, 5, "CRA/GO 6-007-48", 0, 0, 'C')
        
        # Coluna 2: Glayzer
        pdf.set_xy(110, y_assinaturas)
        # pdf.image("assinatura_glayzer.png", x=125, y=y_assinaturas-15, w=30) # Descomente se tiver a imagem
        pdf.cell(80, 5, "Glayzer Antônio Gomes da Silva", 0, 1, 'C')
        pdf.cell(80, 5, "Advogado Especialista", 0, 1, 'C')
        
        # Output
        nome_arquivo_pdf = f"Oficio_{municipio_selecionado}_{num_oficio}.pdf"
        pdf.output(nome_arquivo_pdf)
        
        with open(nome_arquivo_pdf, "rb") as f:
            st.download_button(
                label="⬇️ Baixar Ofício em PDF",
                data=f,
                file_name=nome_arquivo_pdf,
                mime="application/pdf"
            )
