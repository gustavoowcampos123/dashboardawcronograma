import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import re
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime

# Configuração da página
st.set_page_config(page_title="Dashboard de Projetos", page_icon="📊", layout="wide")

# Funções auxiliares
def parse_duracao(duracao_str):
    match = re.search(r'(\d+)', str(duracao_str))
    return int(match.group(1)) if match else None

def parse_date(date_str):
    try:
        return pd.to_datetime(date_str[4:], format='%d/%m/%y', errors='coerce')
    except:
        return pd.NaT

def gerar_curva_s(df_raw, start_date_str='16/09/2024'):
    start_date = pd.to_datetime(start_date_str)
    end_date = df_raw['Término'].max()
    weeks = pd.date_range(start=start_date, end=end_date, freq='W-MON')
    progress_by_week = pd.DataFrame(weeks, columns=['Data'])
    progress_by_week['% Executado'] = 0.0
    for _, row in df_raw.iterrows():
        if pd.notna(row['Início']) and pd.notna(row['Término']):
            task_weeks = pd.date_range(start=row['Início'], end=row['Término'], freq='W-MON')
            weekly_progress = 1 / len(task_weeks) if len(task_weeks) > 0 else 1
            for week in task_weeks:
                progress_by_week.loc[progress_by_week['Data'] == week, '% Executado'] += weekly_progress
    progress_by_week['% Executado Acumulado'] = progress_by_week['% Executado'].cumsum() * 100
    max_progress = progress_by_week['% Executado Acumulado'].max()
    if max_progress > 0:
        progress_by_week['% Executado Acumulado'] = (progress_by_week['% Executado Acumulado'] / max_progress) * 100
    return progress_by_week

def salvar_curva_s_imagem(curva_s_df):
    fig, ax = plt.subplots()
    ax.plot(curva_s_df['Data'], curva_s_df['% Executado Acumulado'], marker='o')
    ax.set_title('Curva S - Progresso Acumulado')
    ax.set_xlabel('Data')
    ax.set_ylabel('% Executado Acumulado')
    fig.autofmt_xdate()
    caminho = '/tmp/curva_s.png'
    plt.savefig(caminho)
    plt.close()
    return caminho

def gerar_pdf_relatorio_completo(curva_s_path, atividades_atrasadas, atividades_semana, atividades_15dias):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    c.setFont("Helvetica-Bold", 16)
    c.drawString(180, height - 40, "RELATÓRIO DE PROJETO")

    # Imagem da curva S
    c.drawImage(ImageReader(curva_s_path), 70, height - 300, width=460, preserveAspectRatio=True)

    # Atividades Atrasadas
    y = height - 320
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, y, f"Atividades Atrasadas ({len(atividades_atrasadas)}):")
    c.setFont("Helvetica", 10)
    y -= 20
    for i, (_, row) in enumerate(atividades_atrasadas.iterrows()):
        text = f"{i+1}. {row['Nome da tarefa']} - Término: {row['Término'].strftime('%d/%m/%Y')}"
        c.drawString(50, y, text)
        y -= 15
        if y < 60:
            c.showPage()
            y = height - 40

    # Atividades Semana
    y -= 20
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, y, "Atividades para a Semana:")
    c.setFont("Helvetica", 10)
    y -= 20
    for i, (_, row) in enumerate(atividades_semana.iterrows()):
        text = f"{i+1}. {row['Nome da tarefa']} - Início: {row['Início'].strftime('%d/%m/%Y')}"
        c.drawString(50, y, text)
        y -= 15
        if y < 60:
            c.showPage()
            y = height - 40

    # Atividades 15 dias
    y -= 20
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, y, "Atividades para os Próximos 15 Dias:")
    c.setFont("Helvetica", 10)
    y -= 20
    for i, (_, row) in enumerate(atividades_15dias.iterrows()):
        text = f"{i+1}. {row['Nome da tarefa']} - Início: {row['Início'].strftime('%d/%m/%Y')}"
        c.drawString(50, y, text)
        y -= 15
        if y < 60:
            c.showPage()
            y = height - 40

    c.save()
    buffer.seek(0)
    return buffer

# Upload do arquivo
uploaded_file = st.file_uploader("Carregar Cronograma", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file)
    df_raw['Duração'] = df_raw['Duração'].apply(parse_duracao)
    df_raw['Início'] = df_raw['Início'].apply(parse_date)
    df_raw['Término'] = df_raw['Término'].apply(parse_date)

    st.title("📊 Dashboard do Projeto")

    # Filtros e indicadores
    hoje = pd.Timestamp.today()
    data_inicio = df_raw['Início'].min()
    data_fim = df_raw['Término'].max()
    prazo_total = (data_fim - data_inicio).days
    dias_restantes = (data_fim - hoje).days
    concluidas = len(df_raw[df_raw['% concluída'] == 1])
    atrasadas = df_raw[(df_raw['Término'] < hoje) & (df_raw['% concluída'] != 1)]
    semana = hoje + pd.Timedelta(days=7)
    atividades_semana = df_raw[(df_raw['Início'] <= semana) & (df_raw['Término'] >= hoje)]
    quinze_dias = hoje + pd.Timedelta(days=15)
    atividades_15dias = df_raw[(df_raw['Início'] <= quinze_dias) & (df_raw['Término'] >= hoje)]

    col1, col2, col3 = st.columns(3)
    col1.metric("Atividades Concluídas", concluidas)
    col2.metric("Prazo Total", f"{prazo_total} dias")
    col3.metric("Dias para Finalizar", f"{dias_restantes} dias")

    # Curva S
    curva_s_df = gerar_curva_s(df_raw, start_date_str=data_inicio.strftime('%d/%m/%Y'))
    fig_curva_s = px.line(curva_s_df, x='Data', y='% Executado Acumulado', title="Curva S - Progresso do Projeto")
    st.plotly_chart(fig_curva_s, use_container_width=True)

    # Gerar imagem da curva e PDF
    curva_path = salvar_curva_s_imagem(curva_s_df)
    pdf_final = gerar_pdf_relatorio_completo(curva_path, atrasadas, atividades_semana, atividades_15dias)
    st.download_button("📥 Baixar Relatório em PDF", data=pdf_final, file_name="relatorio_projeto.pdf", mime="application/pdf")

    # Exibir tabelas
    with st.expander("📋 Dados do Cronograma"):
        st.dataframe(df_raw)
    with st.expander("⏱️ Atividades Atrasadas"):
        st.dataframe(atrasadas)
    with st.expander("📆 Atividades para a Semana"):
        st.dataframe(atividades_semana)
    with st.expander("🗓️ Atividades para os Próximos 15 Dias"):
        st.dataframe(atividades_15dias)
