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

# Dados simulados (exemplo, substitua pelo seu DataFrame real)
df_raw = pd.DataFrame({
    'Início': pd.date_range(start='2024-01-01', periods=100, freq='D'),
    'Término': pd.date_range(start='2024-01-15', periods=100, freq='D'),
    'Predecessoras': [None, 1, 2, None, 4] * 20,
    'Duracao': [15, 10, 5, 20, 10] * 20,
    'Status': ['Concluído', 'Pendente', 'Concluído', 'Atrasado', 'Em andamento'] * 20
})

# Filtragens e métricas
proximos_15_dias = pd.Timestamp.today() + pd.Timedelta(days=15)
atividades_proximos_15_dias = df_raw[(df_raw['Início'] <= proximos_15_dias) & (df_raw['Término'] >= pd.Timestamp.today())]
atividades_sem_predecessora = df_raw[df_raw['Predecessoras'].isna()]
caminho_critico = df_raw[df_raw['Duracao'] > 15]
atividades_atrasadas = df_raw[df_raw['Término'] < pd.Timestamp.today()]

# Indicadores
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

# Expanders
with st.expander("Dados do Cronograma"):
    st.dataframe(df_raw)

with st.expander("Atividades sem Predecessoras"):
    st.dataframe(atividades_sem_predecessora)

with st.expander("Caminho Crítico"):
    st.dataframe(caminho_critico)

with st.expander("Atividades Atrasadas"):
    st.dataframe(atividades_atrasadas)

with st.expander("Atividades para Próximos 15 Dias"):
    st.dataframe(atividades_proximos_15_dias)

# Exportar Relatório em PDF (simulação)
# pdf_data = gerar_relatorio_pdf(df_raw, caminho_critico, atividades_sem_predecessora, atividades_atrasadas, curva_s_img) # Substituir pela função real
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
wb.save(excel_output)
excel_output.seek(0)
st.download_button(
    label="📥 Baixar Relatório em Excel",
    data=excel_output.getvalue(),
    file_name="relatorio_projeto.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
