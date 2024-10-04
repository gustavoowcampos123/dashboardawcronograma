import streamlit as st
import pandas as pd
import io
import openai
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

# Configuração da chave da API OpenAI a partir dos segredos do Streamlit
openai.api_key = st.secrets["openai"]["api_key"]

# Funções de conversão de dados
def parse_duracao(duracao_str):
    match = re.search(r'(\d+)', str(duracao_str))
    return int(match.group(1)) if match else None

def parse_date(date_str):
    try:
        return pd.to_datetime(date_str[4:], format='%d/%m/%y', errors='coerce')
    except ValueError:
        return pd.NaT

def clean_weekday_abbreviation(date_str):
    return date_str[4:] if isinstance(date_str, str) else date_str

def gerar_curva_s(df_raw, start_date_str='16/09/2024'):
    # Função já existente para gerar Curva S

# Funções para ChatGPT
def analisar_atrasos(df):
    prompt = f"""
    Abaixo estão as tarefas de um projeto de construção com datas de início e término:
    {df.to_string(index=False)}
    
    Analise esses dados e identifique as tarefas que têm maior probabilidade de atraso. 
    Considere a duração das tarefas e suas datas de início e término. Quais estratégias podem ser adotadas para minimizar o risco de atrasos? 
    """
    
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        max_tokens=200
    )
    return response.choices[0].text.strip()

def chatbot_assistente(pergunta, dados_do_cronograma):
    prompt = f"""
    Aqui estão alguns dados de cronograma de um projeto:
    {dados_do_cronograma.to_string(index=False)}
    
    Abaixo está uma pergunta do usuário sobre o cronograma:
    {pergunta}
    
    Responda a essa pergunta com detalhes, considerando as informações fornecidas.
    """
    
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        max_tokens=150
    )
    return response.choices[0].text.strip()

# Interface Streamlit
st.sidebar.title("Painel de Controle")
uploaded_file = st.file_uploader("Carregar Cronograma", type=["xlsx"])

if uploaded_file is not None:
    # Leitura do arquivo carregado
    df_raw = pd.read_excel(uploaded_file)
    df_raw['Duração'] = df_raw['Duração'].apply(parse_duracao)
    df_raw['Início'] = df_raw['Início'].apply(parse_date)
    df_raw['Término'] = df_raw['Término'].apply(parse_date)
    
    # Análise de Atrasos
    if st.button("Análise Previsiva de Atrasos"):
        analise_atrasos_texto = analisar_atrasos(df_raw)
        st.write("**Análise de Atrasos:**")
        st.write(analise_atrasos_texto)
    
    # Chatbot de Assistência
    pergunta = st.text_input("Pergunte ao assistente sobre o cronograma:")
    if pergunta:
        resposta_chatbot = chatbot_assistente(pergunta, df_raw)
        st.write("**Resposta do Assistente:**")
        st.write(resposta_chatbot)

    # Demais elementos do dashboard já existentes...
else:
    st.warning("Por favor, carregue o cronograma para visualizar o painel.")
