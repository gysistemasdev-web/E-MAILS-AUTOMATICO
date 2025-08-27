import pandas as pd
import smtplib, time, re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st

# ================================
# CONFIG SMTP (Skymail)
# ================================
SMTP_HOST = "smtp.skymail.net.br"
SMTP_PORT = 465   # SSL/TLS

def enviar_email(email_user, email_pass, para, cc_list, assunto, corpo_html):
    msg = MIMEMultipart("alternative")
    msg["From"] = email_user
    msg["To"] = para
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo_html, "html"))

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(email_user, email_pass)
        server.sendmail(email_user, [para] + cc_list, msg.as_string())

# ================================
# LAYOUT / TEMA
# ================================
st.set_page_config(page_title="Envio Automático de E-mails", layout="wide")
st.markdown(
    """
    <style>
    .stApp { background:#ffffff; color:#1E3A8A; font-family:Arial, sans-serif; }
    .stButton>button{ background:#1E3A8A; color:#fff; border-radius:8px; padding:8px 14px; font-weight:700;}
    .stButton>button:hover{ background:#3B82F6; }
    .stSlider > div > div > div { background:#1E3A8A; }
    h2, h3, .stMarkdown { color:#1E3A8A; }
    .box { background:#F3F6FF; border:1px solid #D9E2FF; border-radius:12px; padding:16px; }
    </style>
    """,
    unsafe_allow_html=True
)

# Logo ACEL
st.markdown(
    """
    <div style="text-align:center; margin-bottom:20px;">
      <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/LOGO%20ACEL.jpg" width="250">
    </div>
    """,
    unsafe_allow_html=True
)
st.markdown("<h2 style='text-align:center;'>📧 Envio Automático de E-mails</h2>", unsafe_allow_html=True)

# ================================
# AJUDA
# ================================
with st.expander("ℹ️ Como usar o sistema"):
    st.markdown("""
    ### 📝 Formatação no texto (conversor automático):
    - **negrito** → escreva entre duas estrelas: `**assim**`
    - vermelho → escreva entre duas hashtags duplas: `##assim##`

    ### 📂 Planilha
    - Colunas exigidas: **E-MAIL** e **RESPONSAVEL**
    - Vários e-mails na mesma célula: separe por `;` ou `-`
    - O primeiro será o "Para", os outros irão em "Cc"

    ### ✉️ Envio
    - **Modo Teste**: envia só para seu e-mail (seguro para validar)
    - Desative para enviar aos destinatários reais
    - Preview mostra os 5 primeiros antes do disparo
    """)

# ================================
# MENSAGEM / ENVIO
# ================================
email_user = st.text_input("Seu e-mail (Skymail)")
email_pass = st.text_input("Sua senha", type="password")

uploaded_file = st.file_uploader("Carregue a planilha (.xlsx ou .csv)", type=["xlsx","csv"])
assunto = st.text_input("Assunto do e-mail", "Comunicado importante - {{responsavel}}")

st.subheader("Corpo do e-mail (digite em texto simples)")
texto_puro = st.text_area("Digite aqui:",
"""
Bom dia, **Prezados(as)**,

A Acel tem como prioridade ##estar sempre ao lado de seus clientes##, adotando as melhores práticas
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
            linha = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", linha)   # negrito
            linha = re.sub(r"##(.*?)##", r"<span style='color:red;'>\1</span>", linha)  # vermelho
            html += f"<p>{linha}</p>\n"
    return html

corpo_base = converter_para_html(texto_puro)

st.subheader("📌 HTML gerado")
st.code(corpo_base, language="html")

st.subheader("🔎 Prévia formatada")
st.markdown(f"<div class='box'>{corpo_base}</div>", unsafe_allow_html=True)

pausa = st.slider("Intervalo entre e-mails (segundos)", 0.5, 5.0, 1.0)
modo_teste = st.checkbox("🔒 Modo Teste (enviar só para mim)", value=True)

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
            emails = [e.strip() for e in emails if "@" in e]
            if emails:
                destino = emails[0]
                copias = emails[1:]
                assunto_p = assunto.replace("{{responsavel}}", responsavel)
                corpo_p = corpo_base.replace("{{responsavel}}", responsavel)
                destino_final = email_user if modo_teste else destino
                preview.append({"Para": destino_final, "Cc": ", ".join(copias), "Assunto": assunto_p, "Corpo": corpo_p})
        st.table(preview)

        if st.button("🚀 Enviar todos os e-mails"):
            enviados = 0
            for _, row in df.iterrows():
                responsavel = str(row["RESPONSAVEL"])
                emails = str(row["E-MAIL"]).replace(";", "-").split("-")
                emails = [e.strip() for e in emails if "@" in e]
                if not emails:
                    continue
                destino = emails[0]
                copias = emails[1:]
                assunto_p = assunto.replace("{{responsavel}}", responsavel)
                corpo_p = corpo_base.replace("{{responsavel}}", responsavel)
                destino_final = email_user if modo_teste else destino
                copias_final = [] if modo_teste else copias
                try:
                    enviar_email(email_user, email_pass, destino_final, copias_final, assunto_p, corpo_p)
                    enviados += 1
                    st.write(f"✅ Enviado: {destino_final} (Cc: {', '.join(copias_final)})")
                    time.sleep(pausa)
                except Exception as ex:
                    st.write(f"⚠️ Erro com {destino_final}: {ex}")
            st.success(f"Finalizado. Total enviados: {enviados}")
    else:
        st.error("A planilha precisa ter as colunas: E-MAIL e RESPONSAVEL")
