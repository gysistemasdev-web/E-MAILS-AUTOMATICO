import pandas as pd
import smtplib, time, re, json, os
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
# LOGIN / REGISTRO
# ================================
USERS_FILE = "usuarios.json"

def carregar_usuarios():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r") as f:
            return json.load(f)
    return {}

def salvar_usuarios(users):
    with open(USERS_FILE, "w") as f:
        json.dump(users, f)

usuarios = carregar_usuarios()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "usuario" not in st.session_state:
    st.session_state.usuario = None

if not st.session_state.logged_in:
    st.title("🔒 Login - Sistema de Envio ACEL")

    escolha = st.radio("Selecione:", ["Login", "Registrar"])
    email = st.text_input("E-mail corporativo")
    senha = st.text_input("Senha", type="password")

    if escolha == "Registrar":
        if st.button("Criar conta"):
            if not email.endswith("@acelnet.com.br"):
                st.error("❌ Apenas e-mails @acelnet.com.br podem se registrar.")
            elif email in usuarios:
                st.error("❌ Usuário já existe.")
            else:
                usuarios[email] = senha
                salvar_usuarios(usuarios)
                st.success("✅ Conta criada com sucesso! Agora faça login.")

    elif escolha == "Login":
        if st.button("Entrar"):
            if email in usuarios and usuarios[email] == senha:
                st.session_state.logged_in = True
                st.session_state.usuario = email
                st.success(f"✅ Bem-vindo, {email}")
                st.rerun()
            else:
                st.error("❌ E-mail ou senha inválidos.")

    st.stop()

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
    "Leonardo": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/LEONARDO.png" width="450">""",
    "Erika": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/ERIKA%201.png" width="450">""",
    "Caroline": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/CAROLINE.png" width="450">""",
    "Halline": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/HALLINE.jpg" width="450">""",
    "Marcio": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/MARCIO.png" width="450">""",
    "Maria Alice": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/MARIA%20ALICE%20-%20DP.jpg" width="450">""",
    "Ronaldo": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/RONALDO.png" width="450">""",
    "Victor": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/VICTOR.jpg" width="450">""",
    "Robson": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/acel%20assinaturas.png" width="450">"""
}

# ================================
# TABS: MENSAGEM / ASSINATURAS
# ================================
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
                linha = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", linha)   # negrito
                linha = re.sub(r"##(.*?)##", r"<span style='color:red;'>\1</span>", linha)  # vermelho
                html += f"<p>{linha}</p>\n"
        return html

    corpo_base = converter_para_html(texto_puro)
    corpo_preview = corpo_base + ("<hr>" + assinatura_html if assinatura_html else "")

    st.subheader("📌 HTML gerado")
    st.code(corpo_preview, language="html")

    st.subheader("🔎 Prévia formatada")
    st.markdown(f"<div class='box'>{corpo_preview}</div>", unsafe_allow_html=True)

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
                for e in emails:
                    e = e.strip()
                    if "@" in e:
                        assunto_p = assunto.replace("{{responsavel}}", responsavel)
                        corpo_p = corpo_base.replace("{{responsavel}}", responsavel)
                        corpo_p = corpo_p + ("<hr>" + assinatura_html if assinatura_html else "")
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
                        corpo_p = corpo_p + ("<hr>" + assinatura_html if assinatura_html else "")
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
# PAINEL ADMINISTRATIVO (somente admins)
# ================================
ADMINS = [
    "gabryell@acelnet.com.br",
    "marcio@acelnet.com.br",
    "leonardo@acelnet.com.br",
    "victor@acelnet.com.br"
]

if st.session_state.usuario in ADMINS:
    st.markdown("## ⚙️ Painel Administrativo")

    usuarios = carregar_usuarios()

    # Mostrar tabela com e-mail e senha em texto puro
    st.write("📋 Usuários cadastrados (com senha):")
    data = [{"E-mail": u, "Senha": p} for u, p in usuarios.items()]
    st.table(data)

    excluir = st.selectbox("Selecione um usuário para excluir", [""] + list(usuarios.keys()))

    if excluir and excluir != "":
        if st.button("❌ Excluir usuário"):
            usuarios.pop(excluir, None)
            salvar_usuarios(usuarios)
            st.success(f"Usuário {excluir} removido com sucesso!")
            st.rerun()

# ================================
# Rodapé
# ================================
st.markdown(
    f"""
    <hr>
    <p style="text-align:center; font-size:11px; color:#888;">
      Sistema elaborado por <b>GY SISTEMAS</b> - GABRYELL FELIX, YAGO SILVA<br>
      🔓 Usuário logado: {st.session_state.usuario}
    </p>
    """,
    unsafe_allow_html=True
)
