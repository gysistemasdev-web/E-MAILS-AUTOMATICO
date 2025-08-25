import pandas as pd
import smtplib, time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st

# Configurações do servidor Skymail
SMTP_HOST = "smtp.skymail.net.br"
SMTP_PORT = 465   # porta de envio da Skymail (SSL/TLS)

def enviar_email(email_user, email_pass, para, assunto, corpo_html):
    msg = MIMEMultipart("alternative")
    msg["From"] = email_user
    msg["To"] = para
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo_html, "html"))

    # 🔑 Conexão com SSL direto (porta 465)
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(email_user, email_pass)
        server.sendmail(email_user, [para], msg.as_string())

# Interface web
st.title("📧 Envio Automático de E-mails")

email_user = st.text_input("Seu e-mail (Skymail)", "gabryell@acelnet.com.br")
email_pass = st.text_input("Sua senha", type="password")

uploaded_file = st.file_uploader("Carregue a planilha (.xlsx ou .csv)", type=["xlsx","csv"])
assunto = st.text_input("Assunto do e-mail", "Comunicado importante - {{responsavel}}")
corpo = st.text_area("Corpo do e-mail (HTML permitido)", 
"""
<p>Olá {{responsavel}},</p>
<p>Segue comunicado referente aos nossos serviços.</p>
<p>Atenciosamente,<br>Equipe ACEL</p>
""", height=200)

pausa = st.slider("Intervalo entre e-mails (segundos)", 0.5, 5.0, 1.0)

modo_teste = st.checkbox("🔒 Modo Teste (enviar só para mim)", value=True)

if uploaded_file is not None:
    # Carrega planilha
    if uploaded_file.name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file, header=0)
    else:
        df = pd.read_csv(uploaded_file, header=0)

    # 🔑 Normaliza os nomes das colunas
    df.columns = df.columns.str.strip().str.upper()

    st.subheader("Prévia dos dados carregados")
    st.write(df.head())

    # preview dos primeiros e-mails
    if "E-MAIL" in df.columns and "RESPONSAVEL" in df.columns:
        st.subheader("📌 Preview dos primeiros 5 e-mails")
        preview = []
        for _, row in df.head(5).iterrows():
            responsavel = str(row["RESPONSAVEL"])
            emails = str(row["E-MAIL"]).replace(";", "-").split("-")
            for e in emails:
                e = e.strip()
                if "@" in e:
                    assunto_p = assunto.replace("{{responsavel}}", responsavel)
                    corpo_p = corpo.replace("{{responsavel}}", responsavel)
                    destino = email_user if modo_teste else e
                    preview.append({"Para": destino, "Assunto": assunto_p, "Corpo": corpo_p})
        st.table(preview)

        # Botão para enviar
        if st.button("🚀 Enviar todos os e-mails"):
            enviados = 0
            for _, row in df.iterrows():
                responsavel = str(row["RESPONSAVEL"])
                emails = str(row["E-MAIL"]).replace(";", "-").split("-")
                for e in emails:
                    e = e.strip()
                    if "@" not in e:
                        continue
                    assunto_p = assunto.replace("{{responsavel}}", responsavel)
                    corpo_p = corpo.replace("{{responsavel}}", responsavel)
                    destino = email_user if modo_teste else e
                    try:
                        enviar_email(email_user, email_pass, destino, assunto_p, corpo_p)
                        enviados += 1
                        st.write(f"✅ Enviado: {destino}")
                        time.sleep(pausa)
                    except Exception as ex:
                        st.write(f"⚠️ Erro com {e}: {ex}")
            st.success(f"Finalizado. Total enviados: {enviados}")
    else:
        st.error("A planilha precisa ter as colunas: E-MAIL e RESPONSAVEL")
