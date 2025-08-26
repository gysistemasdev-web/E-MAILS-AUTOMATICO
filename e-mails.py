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

# ================================
# Configurações de Layout
# ================================
st.set_page_config(page_title="Envio Automático de E-mails", layout="wide")

# Logo da ACEL centralizada
st.markdown(
    """
    <div style="text-align:center;">
        <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/LOGO ACEL.jpg" width="200">
    </div>
    """,
    unsafe_allow_html=True
)

# Título estilizado
st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>📧 Envio Automático de E-mails</h2>", unsafe_allow_html=True)

# ================================
# Entrada de dados
# ================================
email_user = st.text_input("Seu e-mail (Skymail)", "gabryell@acelnet.com.br")
email_pass = st.text_input("Sua senha", type="password")

uploaded_file = st.file_uploader("Carregue a planilha (.xlsx ou .csv)", type=["xlsx","csv"])
assunto = st.text_input("Assunto do e-mail", "Comunicado importante - {{responsavel}}")

# Texto simples convertido em HTML
st.subheader("Corpo do e-mail (digite em texto simples)")
texto_puro = st.text_area("Digite aqui:", 
"""
Bom dia, Prezados(as),

A Acel tem como prioridade estar sempre ao lado de seus clientes, adotando as melhores práticas
para simplificar a rotina e assegurar segurança em todas as etapas dos processos contábeis e fiscais.

Atenciosamente,
Equipe ACEL
""", height=250)

def converter_para_html(texto):
    linhas = texto.split("\n")
    html = ""
    for linha in linhas:
        linha = linha.strip()
        if linha:
            html += f"<p>{linha}</p>\n"
    return html

corpo = converter_para_html(texto_puro)

st.subheader("📌 Pré-visualização em HTML")
st.code(corpo, language="html")

pausa = st.slider("Intervalo entre e-mails (segundos)", 0.5, 5.0, 1.0)

modo_teste = st.checkbox("🔒 Modo Teste (enviar só para mim)", value=True)

# ================================
# Processamento da planilha
# ================================
if uploaded_file is not None:
    if uploaded_file.name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file, header=0)
    else:
        df = pd.read_csv(uploaded_file, header=0)

    df.columns = df.columns.str.strip().str.upper()

    st.subheader("Prévia dos dados carregados")
    st.write(df.head())

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

# ================================
# Rodapé
# ================================
st.markdown(
    """
    <hr>
    <p style="text-align:center; font-size:11px; color:#888;">
    Sistema elaborado por <b>GY SISTEMAS</b> - GABRYELL FELIX, YAGO SILVA
    </p>
    """,
    unsafe_allow_html=True
)
