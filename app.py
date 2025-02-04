import streamlit as st
import pandas as pd
import urllib.parse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time

def gerar_mensagem(grupo):
    razao_social = grupo['Razão social'].iloc[0]
    lista_notas = [
        f"- **NF {nf}** no valor de R${valor:.2f} com **{atraso} dias de atraso**"
        for nf, valor, atraso in zip(grupo['NF'], grupo['Valor'], grupo['Atraso'])
    ]
    mensagem = (
        f"Bom dia, falo na empresa *{razao_social}*?\n\n"
        f"Falo em nome do financeiro da empresa *Broker Norte, representação da Nestlé.*\n\n"
        f"Estamos entrando em contato pois identificamos os seguintes pagamentos pendentes:\n\n"
        + "\n".join(lista_notas) + "\n\n"
        f"Caso já tenha sido efetuado o pagamento, por gentileza, envie-nos o comprovante.\n\n"
        f"Se o pagamento ainda não foi feito e houver qualquer dúvida ou se precisar de assistência para efetuar o pagamento, "
        f"estamos à disposição para ajudar.\n\n"
        f"Agradecemos desde já pela sua atenção.\n\n"
        f"*Financeiro*\n"
        f"*Broker Norte - Nestlé*"
    )
    return mensagem

st.title("Automação de Cobrança via WhatsApp")

uploaded_file = st.file_uploader("Faça upload da planilha de cobrança", type=["xlsx"])

if uploaded_file:
    contatos_df = pd.read_excel(uploaded_file)
    mensagens = []
    for numero, grupo in contatos_df.groupby('Numero'):
        mensagem = gerar_mensagem(grupo)
        mensagens.append({'Numero': numero, 'Razão social': grupo['Razão social'].iloc[0], 'Mensagem': mensagem})
    mensagens_df = pd.DataFrame(mensagens)
    st.write("Pré-visualização das mensagens:")
    st.dataframe(mensagens_df)
    
    if st.button("Iniciar envio pelo WhatsApp Web"):
        navegador = webdriver.Chrome()
        try:
            navegador.get("https://web.whatsapp.com/")
            WebDriverWait(navegador, 60).until(
                EC.presence_of_element_located((By.ID, "side"))
            )
            for _, row in mensagens_df.iterrows():
                numero = row["Numero"]
                mensagem = row["Mensagem"]
                texto = urllib.parse.quote(mensagem)
                link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
                navegador.get(link)
                WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div/p/span'))
                )
                element = navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div/p/span')
                element.send_keys(Keys.ENTER)
                time.sleep(20)
        finally:
            navegador.quit()
        st.success("Mensagens enviadas com sucesso!")
