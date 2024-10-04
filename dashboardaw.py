import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import plotly.express as px
import re
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Configurações de Página
st.set_page_config(
    page_title="Dashboard de Projetos",
    page_icon="📊",
    layout="wide"
)

# Funções de conversão de dados
def parse_duracao(duracao_str):
    match = re.search(r'(\d+)', str(duracao_str))
    return int(match.group(1)) if match else None

def parse_date(date_str):
    try:
        return pd.to_datetime(date_str[4:], format='%d/%m/%y', errors='coerce')
    except ValueError:
        return pd.NaT

# Função para gerar PDF das atividades atrasadas
def gerar_pdf_atividades_atrasadas(atividades_atrasadas):
    pdf_buffer = io.BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=letter)
    c.drawString(100, 750, "Relatório de Atividades Atrasadas")
    
    # Adicionar conteúdo das atividades atrasadas
    y_position = 730
    for i, (_, row) in enumerate(atividades_atrasadas.iterrows()):
        c.drawString(50, y_position, f"{i + 1}. {row['Nome da tarefa']} - {row['Término'].strftime('%d/%m/%Y')}")
        y_position -= 20
        if y_position < 50:
            c.showPage()
            y_position = 750

    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer

# Sidebar
st.sidebar.title("Painel de Controle")
st.sidebar.info("Navegue pelas seções para acessar diferentes dados do projeto.")

# Upload do Cronograma
uploaded_file = st.file_uploader("Carregar Cronograma", type=["xlsx"])

if uploaded_file is not None:
    # Leitura e processamento do arquivo carregado
    df_raw = pd.read_excel(uploaded_file)

    # Ajustes para converter duração e datas
    df_raw['Duração'] = df_raw['Duração'].apply(parse_duracao)
    df_raw['Início'] = df_raw['Início'].apply(parse_date)
    df_raw['Término'] = df_raw['Término'].apply(parse_date)

    # Verificar colunas necessárias
    if 'Duração' not in df_raw.columns or 'Início' not in df_raw.columns or 'Término' not in df_raw.columns:
        st.error("O arquivo deve conter as colunas 'Duração', 'Início', e 'Término'.")
    else:
        # Cálculo do prazo total com base nas datas
        data_inicio_mais_cedo = df_raw['Início'].min()
        data_termino_mais_tarde = df_raw['Término'].max()
        prazo_total = (data_termino_mais_tarde - data_inicio_mais_cedo).days

        # Filtros de atividades
        proximos_15_dias = pd.Timestamp.today() + pd.Timedelta(days=15)
        atividades_proximos_15_dias = df_raw[(df_raw['Início'] <= proximos_15_dias) & (df_raw['Término'] >= pd.Timestamp.today())]
        atividades_sem_predecessora = df_raw[df_raw['Predecessoras'].isna()]
        caminho_critico = df_raw[df_raw['Duração'] > 15]
        atividades_atrasadas = df_raw[(df_raw['Término'] < pd.Timestamp.today()) & (df_raw['% concluída'] != 1)]
        proxima_semana = pd.Timestamp.today() + pd.Timedelta(days=7)
        atividades_proxima_semana = df_raw[(df_raw['Início'] <= proxima_semana) & (df_raw['Término'] >= pd.Timestamp.today())]

        # Indicadores
        atividades_concluidas = len(df_raw[df_raw['% concluída'] == 1])
        st.title("Dashboard do Projeto")
        col1, col2, col3 = st.columns(3)
        col1.metric("Atividades Concluídas", atividades_concluidas)
        
        with col2:
            st.metric("Atividades Atrasadas", len(atividades_atrasadas))
            if st.button("Gerar PDF de Atividades Atrasadas"):
                pdf_atividades = gerar_pdf_atividades_atrasadas(atividades_atrasadas)
                st.download_button(
                    label="📥 Baixar PDF de Atividades Atrasadas",
                    data=pdf_atividades,
                    file_name="atividades_atrasadas.pdf",
                    mime="application/pdf"
                )

        col3.metric("Prazo Total do Projeto", f"{prazo_total} dias")

        # Cálculo da Curva S
        df_raw['Semana'] = df_raw['Início'].dt.isocalendar().week
        progresso_semanal = df_raw.groupby('Semana')['% concluída'].sum().cumsum()
        curva_s_df = pd.DataFrame({'Semana': progresso_semanal.index, 'Progresso Acumulado': progresso_semanal})

        # Gráfico da Curva S
        fig_curva_s = px.line(curva_s_df, x='Semana', y='Progresso Acumulado', title="Curva S - Progresso Acumulado do Projeto")
        st.plotly_chart(fig_curva_s, use_container_width=True)

        # Expansíveis
        with st.expander("Dados do Cronograma"):
            st.dataframe(df_raw)
        with st.expander("Atividades sem Predecessoras"):
            st.dataframe(atividades_sem_predecessora)
        with st.expander("Caminho Crítico"):
            st.dataframe(caminho_critico)
        with st.expander("Atividades Atrasadas"):
            st.dataframe(atividades_atrasadas)
        with st.expander("Atividades para Próxima Semana"):
            st.dataframe(atividades_proxima_semana)
        with st.expander("Atividades para os Próximos 15 Dias"):
            st.dataframe(atividades_proximos_15_dias)

        # Exportações
        pdf_data = io.BytesIO()
        st.download_button(label="📥 Baixar Relatório em PDF", data=pdf_data.getvalue(), file_name="relatorio_projeto.pdf", mime="application/pdf")

        excel_output = io.BytesIO()
        wb = Workbook()
        ws_curva_s = wb.active
        ws_curva_s.title = 'Curva S'
        for r in dataframe_to_rows(curva_s_df, index=False, header=True):
            ws_curva_s.append(r)
        wb.save(excel_output)
        excel_output.seek(0)
        st.download_button(label="📥 Baixar Relatório em Excel", data=excel_output.getvalue(), file_name="relatorio_projeto.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.warning("Por favor, carregue o cronograma para visualizar o painel.")
