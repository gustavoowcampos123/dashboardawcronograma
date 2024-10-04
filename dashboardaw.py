import streamlit as st
import pandas as pd
import openai

# Configuração da chave da API OpenAI diretamente dos segredos do Streamlit
openai.api_key = "org-KKZTeMEKFocZz7O14lGkSfJE"


# Função para analisar atrasos usando a nova API de completions
def analisar_atrasos(df):
    prompt = f"""
    Abaixo estão as tarefas de um projeto de construção com datas de início e término:
    {df.to_string(index=False)}
    
    Analise esses dados e identifique as tarefas que têm maior probabilidade de atraso. 
    Considere a duração das tarefas e suas datas de início e término. Quais estratégias podem ser adotadas para minimizar o risco de atrasos? 
    """
    
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=200
    )
    return response['choices'][0]['message']['content'].strip()

# Exemplo de uso no Streamlit
uploaded_file = st.file_uploader("Carregar Cronograma", type=["xlsx"])
if uploaded_file is not None:
    df_raw = pd.read_excel(uploaded_file)
    if st.button("Análise Previsiva de Atrasos"):
        try:
            analise_atrasos_texto = analisar_atrasos(df_raw)
            st.write("**Análise de Atrasos:**")
            st.write(analise_atrasos_texto)
        except openai.error.AuthenticationError:
            st.error("A chave de API não é válida ou expirou.")
        except Exception as e:
            st.error(f"Ocorreu um erro: {e}")
else:
    st.warning("Por favor, carregue o cronograma para visualizar o painel.")
