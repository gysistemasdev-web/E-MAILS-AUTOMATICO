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

def enviar_email(email_user, email_pass, para, assunto, corpo_html):
    msg = MIMEMultipart("alternative")
    msg["From"] = email_user
    msg["To"] = para
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo_html, "html"))
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(email_user, email_pass)
        server.sendmail(email_user, [para], msg.as_string())

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
      <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/7d65ef74bd8aa33b461e015093402792c923d81d/LOGO%20ACEL.jpg" width="250">
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
    - Colunas exigidas: **E-MAIL** e **RESPONSAVEL** (maiúsculas; o app normaliza).
    - Vários e-mails na mesma célula: separe por `;` ou `-`.

    ### ✉️ Envio
    - **Modo Teste**: envia só para seu e-mail (seguro para validar).
    - Desative para enviar aos destinatários da planilha.
    - Preview mostra os 5 primeiros antes do disparo.

    ### 🔏 Assinaturas
    - Vá na aba **Assinaturas**, escolha o nome, URL personalizada ou cole HTML.
    - A assinatura selecionada é adicionada ao final do e-mail automaticamente.
    """)

# ================================
# ASSINATURAS (catálogo)
# ================================
ASSINATURAS = {
    "Leonardo": """
        <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/b3d8ac4a483b53c9f7c61c982fc74e98afce1826/LEONARDO.png"
             alt="Assinatura - Leonardo" width="450">
    """,
    "Erika": """
        <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/b3d8ac4a483b53c9f7c61c982fc74e98afce1826/ERIKA%201.png"
             alt="Assinatura - Erika" width="450">
    """,
    "Caroline": """
        <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/b3d8ac4a483b53c9f7c61c982fc74e98afce1826/CAROLINE.png"
             alt="Assinatura - Caroline" width="450">
    """,
    "Halline": """
        <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/b3d8ac4a483b53c9f7c61c982fc74e98afce1826/HALLINE.jpg"
             alt="Assinatura - Halline" width="450">
    """,
    "Marcio": """
        <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/b3d8ac4a483b53c9f7c61c982fc74e98afce1826/MARCIO.png"
             alt="Assinatura - Marcio" width="450">
    """,
    "Maria Alice": """
        <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/b3d8ac4a483b53c9f7c61c982fc74e98afce1826/MARIA%20ALICE%20-%20DP.jpg"
             alt="Assinatura - Maria Alice" width="450">
    """,
    "Ronaldo": """
        <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/b3d8ac4a483b53c9f7c61c982fc74e98afce1826/RONALDO.png"
             alt="Assinatura - Ronaldo" width="450">
    """,
    "Victor": """
        <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/b3d8ac4a483b53c9f7c61c982fc74e98afce1826/VICTOR.jpg"
             alt="Assinatura - Victor" width="450">
    """,
    "Robson": """
        <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/b3d8ac4a483b53c9f7c61c982fc74e98afce1826/acel%20assinaturas.png"
             alt="Assinatura - Robson" width="450">
    """,
}

# Tabs: Mensagem / Assinaturas
tab_msg, tab_ass = st.tabs(["📝 Mensagem", "🖋️ Assinaturas"])

with tab_ass:
    st.markdown("#### Selecione a assinatura a incluir no fim do e-mail")
    modo_ass = st.radio("Modo de assinatura", ["Sem assinatura", "Catálogo (por nome)", "URL personalizada (imagem)", "HTML manual"], index=1)
    assinatura_html = ""

    if modo_ass == "Catálogo (por nome)":
        nome = st.selectbox("Escolha o nome", list(ASSINATURAS.keys()))
        assinatura_html = ASSINATURAS.get(nome, "")
    elif modo_ass == "URL personalizada (imagem)":
        url_img = st.text_input("URL da imagem (https)", "https://...")
        if url_img and url_img.startswith("http"):
            assinatura_html = f'<img src="{url_img}" alt="Assinatura" width="450">'
    elif modo_ass == "HTML manual":
        assinatura_html = st.text_area("Cole seu HTML de assinatura", "<p>Minha assinatura</p>", height=160)

    st.markdown("**Prévia da assinatura**")
    st.markdown(f"<div class='box'>{assinatura_html or '<i>(sem assinatura)</i>'}</div>", unsafe_allow_html=True)

with tab_msg:
    email_user = st.text_input("Seu e-mail (Skymail)", "gabryell@acelnet.com.br")
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
                # **negrito**
                linha = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", linha)
                # ##vermelho##
                linha = re.sub(r"##(.*?)##", r"<span style='color:red;'>\1</span>", linha)
                html += f"<p>{linha}</p>\n"
        return html

    corpo_base = converter_para_html(texto_puro)

    # Assinatura no preview (se houver)
    corpo_preview = corpo_base
    if assinatura_html:
        corpo_preview += "<hr>\n" + assinatura_html

    st.subheader("📌 HTML gerado")
    st.code(corpo_preview, language="html")

    st.subheader("🔎 Prévia formatada")
    st.markdown(f"<div class='box'>{corpo_preview}</div>", unsafe_allow_html=True)

    pausa = st.slider("Intervalo entre e-mails (segundos)", 0.5, 5.0, 1.0)
    modo_teste = st.checkbox("🔒 Modo Teste (enviar só para mim)", value=True)

    # ================================
    # Processamento da planilha + envio
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
                        corpo_p = corpo_base.replace("{{responsavel}}", responsavel)
                        # anexa assinatura escolhida
                        corpo_p = (corpo_p + "<hr>\n" + assinatura_html) if assinatura_html else corpo_p
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
                        corpo_p = corpo_base.replace("{{responsavel}}", responsavel)
                        if assinatura_html:
                            corpo_p = corpo_p + "<hr>\n" + assinatura_html
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

