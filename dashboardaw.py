import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import plotly.express as px

# Configurações de Página
st.set_page_config(
    page_title="Dashboard de Projetos",
    page_icon="📊",
    layout="wide"
)

# Estilos customizados
st.markdown(
    """
    <style>
    .reportview-container .main .block-container {
        padding: 1rem;
        background-color: #f5f5f5;
    }
    .sidebar .sidebar-content {
        background-color: #263238;
        color: white;
    }
    .sidebar .sidebar-content .btn {
        background-color: #f9aa33;
        color: black;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Sidebar
st.sidebar.title("Painel de Controle")
st.sidebar.info("Navegue pelas seções para acessar diferentes dados do projeto.")

# Upload do Cronograma
uploaded_file = st.file_uploader("Carregar Cronograma", type=["xlsx"])

if uploaded_file is not None:
    # Leitura do arquivo carregado
    df_raw = pd.read_excel(uploaded_file)
    
    # Verificação de colunas necessárias
    required_columns = ['Início', 'Término', 'Predecessoras', 'Duracao', 'Status']
    missing_columns = [col for col in required_columns if col not in df_raw.columns]
    
    if missing_columns:
        st.error(f"O arquivo está faltando as seguintes colunas necessárias: {', '.join(missing_columns)}")
    else:
        # Conversão de colunas para datetime com tratamento de erros
        df_raw['Início'] = pd.to_datetime(df_raw['Início'], errors='coerce')
        df_raw['Término'] = pd.to_datetime(df_raw['Término'], errors='coerce')

        # Verificar valores inválidos nas colunas de data
        if df_raw['Início'].isna().sum() > 0 or df_raw['Término'].isna().sum() > 0:
            st.warning("Algumas datas nos campos 'Início' ou 'Término' não puderam ser interpretadas e foram ignoradas.")

        # Filtros de atividades para próximos 15 dias e sem predecessoras
        proximos_15_dias = pd.Timestamp.today() + pd.Timedelta(days=15)
        atividades_proximos_15_dias = df_raw[(df_raw['Início'] <= proximos_15_dias) & (df_raw['Término'] >= pd.Timestamp.today())]
        atividades_sem_predecessora = df_raw[df_raw['Predecessoras'].isna()]

        # Cálculo do caminho crítico (atividades com duração superior a 15 dias)
        caminho_critico = df_raw[df_raw['Duracao'] > 15]

        # Filtro de atividades atrasadas (com término antes da data atual)
        atividades_atrasadas = df_raw[df_raw['Término'] < pd.Timestamp.today()]

        # Atividades para a próxima semana
        proxima_semana = pd.Timestamp.today() + pd.Timedelta(days=7)
        atividades_proxima_semana = df_raw[(df_raw['Início'] <= proxima_semana) & (df_raw['Término'] >= pd.Timestamp.today())]

        # Indicadores principais
        st.title("Dashboard do Projeto")
        col1, col2, col3 = st.columns(3)
        col1.metric("Atividades Concluídas", len(df_raw[df_raw['Status'] == 'Concluído']))
        col2.metric("Atividades Atrasadas", len(atividades_atrasadas))
        col3.metric("Prazo Total do Projeto", f"{df_raw['Duracao'].sum()} dias")

        # Gráfico Curva S (simulação)
        curva_s_df = pd.DataFrame({
            'Semana': list(range(1, 21)),
            'Progresso': [5, 10, 20, 30, 40, 50, 60, 65, 70, 80, 85, 90, 92, 93, 95, 96, 97, 98, 99, 100]
        })
        fig_curva_s = px.line(curva_s_df, x='Semana', y='Progresso', title="Curva S - Progresso do Projeto")
        st.plotly_chart(fig_curva_s, use_container_width=True)

        # Seções Expansíveis para Visualização de Dados
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

        # Exportar Relatório em PDF (simulação)
        pdf_data = io.BytesIO()  # Placeholder para o PDF
        st.download_button(
            label="📥 Baixar Relatório em PDF",
            data=pdf_data.getvalue(),
            file_name="relatorio_projeto.pdf",
            mime="application/pdf"
        )

        # Exportar Excel
        excel_output = io.BytesIO()
        wb = Workbook()
        ws_curva_s = wb.active
        ws_curva_s.title = 'Curva S'
        for r in dataframe_to_rows(curva_s_df, index=False, header=True):
            ws_curva_s.append(r)

        ws_atividades_proxima_semana = wb.create_sheet(title="Atividades Próxima Semana")
        for r in dataframe_to_rows(atividades_proxima_semana, index=False, header=True):
            ws_atividades_proxima_semana.append(r)

        ws_atividades_proximos_15_dias = wb.create_sheet(title="Atividades Próximos 15 Dias")
        for r in dataframe_to_rows(atividades_proximos_15_dias, index=False, header=True):
            ws_atividades_proximos_15_dias.append(r)

        wb.save(excel_output)
        excel_output.seek(0)
        st.download_button(
            label="📥 Baixar Relatório em Excel",
            data=excel_output.getvalue(),
            file_name="relatorio_projeto.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Por favor, carregue o cronograma para visualizar o painel.")
