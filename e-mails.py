# emails.py — ACEL | Envio Automático de E-mails + Aba de Assinaturas
# ---------------------------------------------------------------
# Execução: streamlit run emails.py

import os, re, json, time, smtplib, base64, hashlib
import pandas as pd
from io import BytesIO
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st

# ================================
# CONFIG SMTP (Skymail)
# ================================
SMTP_HOST = "smtp.skymail.net.br"
SMTP_PORT = 465  # SSL/TLS

def enviar_email(email_user, email_pass, para, cc_list, bcc_list, assunto, corpo_html):
    msg = MIMEMultipart("alternative")
    msg["From"] = email_user
    msg["To"] = para
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo_html, "html"))

    dest_all = [para] + cc_list + bcc_list
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(email_user, email_pass)
        server.sendmail(email_user, dest_all, msg.as_string())

# ================================
# LOGIN / REGISTRO
# ================================
USERS_FILE = "usuarios.json"

def carregar_usuarios():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def salvar_usuarios(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

usuarios = carregar_usuarios()
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "usuario" not in st.session_state:
    st.session_state.usuario = None

# ================================
# PÁGINA / ESTILO
# ================================
st.set_page_config(page_title="Envio Automático de E-mails • ACEL", layout="wide")
st.markdown("""
<style>
.stApp { background:#ffffff; color:#0f172a; font-family:Arial, sans-serif; }
.stButton>button{ background:#1E3A8A; color:#fff; border-radius:8px; padding:8px 14px; font-weight:700;}
.stButton>button:hover{ background:#3B82F6; }
.stSlider > div > div > div { background:#1E3A8A; }
h2,h3,.stMarkdown { color:#1E3A8A; }
.box { background:#F3F6FF; border:1px solid #D9E2FF; border-radius:12px; padding:16px; }
</style>
""", unsafe_allow_html=True)

# ================================
# LOGO
# ================================
st.markdown("""
<div style="text-align:center; margin-bottom:20px;">
  <img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/LOGO%20ACEL.jpg" width="220">
</div>
""", unsafe_allow_html=True)

# ================================
# TABS
# ================================
aba_envio, aba_assinaturas = st.tabs(["📧 Envio de E-mails", "🖊️ Assinaturas"])

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
ASSINATURA_USUARIO = {
    "gabryell@acelnet.com.br": ASSINATURAS["Leonardo"],
    "marcio@acelnet.com.br": ASSINATURAS["Marcio"],
    "leonardo@acelnet.com.br": ASSINATURAS["Leonardo"],
    "victor@acelnet.com.br": ASSINATURAS["Victor"],
    "erika@acelnet.com.br": ASSINATURAS["Erika"],
    "caroline@acelnet.com.br": ASSINATURAS["Caroline"],
    "halline@acelnet.com.br": ASSINATURAS["Halline"],
    "mariaalice@acelnet.com.br": ASSINATURAS["Maria Alice"],
    "ronaldo@acelnet.com.br": ASSINATURAS["Ronaldo"],
    "robson@acelnet.com.br": ASSINATURAS["Robson"]
}

# ================================
# FUNÇÕES AUXILIARES
# ================================
def converter_para_html(texto):
    linhas = texto.split("\n")
    html = ""
    for linha in linhas:
        linha = linha.strip()
        if not linha:
            continue
        linha = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", linha)
        linha = re.sub(r"##(.+?)##", r"<span style='color:red;'>\1</span>", linha)
        html += f"<p>{linha}</p>\n"
    return html

def parse_multiplos_emails(celula):
    bruto = str(celula or "")
    partes = [e.strip() for e in re.split(r"[;,\s]+", bruto) if e and "@" in e]
    if not partes:
        return None, []
    return partes[0], partes[1:]

def img_to_data_uri(file):
    b = file.read()
    mime = file.type if hasattr(file, "type") and file.type else "image/png"
    b64 = base64.b64encode(b).decode("utf-8")
    return f"<img src='data:{mime};base64,{b64}' width='450'>"

# ================================
# LOGIN (aba de Envio)
# ================================
with aba_envio:
    if not st.session_state.logged_in:
        st.title("🔒 Login - Sistema de Envio ACEL")
        escolha = st.radio("Selecione:", ["Login", "Registrar"], horizontal=True)
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
        else:
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
# CONTEÚDO ABA ENVIO
# ================================
with aba_envio:
    st.markdown("<h2>📧 Envio Automático de E-mails</h2>", unsafe_allow_html=True)

    assinatura_padrao = ASSINATURA_USUARIO.get(st.session_state.usuario, "")
    assinatura_html = st.session_state.get("assinatura_html", assinatura_padrao)

    colA, colB = st.columns([1,1])
    with colA:
        email_user = st.text_input("Seu e-mail (Skymail)", st.session_state.usuario or "")
        email_pass = st.text_input("Sua senha", type="password")
        assunto = st.text_input("Assunto do e-mail", "Comunicado importante - {{responsavel}}")
        pausa = st.slider("Intervalo entre e-mails (segundos)", 0.5, 5.0, 1.0)
        modo_teste = st.checkbox("🔒 Modo Teste (enviar só para mim)", value=True)

        bcc_raw = st.text_input("BCC (opcional) — separe por ; ou ,", "")
        bcc_list = [e.strip() for e in re.split(r"[;,\s]+", bcc_raw) if "@" in e]

    with colB:
        uploaded_file = st.file_uploader("Carregue a planilha (.xlsx ou .csv)", type=["xlsx","csv"])
        st.subheader("Corpo do e-mail (digite em texto simples)")
        texto_puro = st.text_area("Digite aqui:", """Bom dia,

Prezados(as),

A Acel tem como prioridade ##estar sempre ao lado de seus clientes##, adotando as melhores práticas para simplificar a rotina e assegurar segurança em todas as etapas dos processos contábeis e fiscais.

Atenciosamente,
Equipe ACEL
""", height=260)

    corpo_base = converter_para_html(texto_puro)
    corpo_preview = corpo_base + ("<hr>" + assinatura_html if assinatura_html else "")
    st.subheader("📌 HTML gerado")
    st.code(corpo_preview, language="html")
    st.subheader("🔎 Prévia formatada")
    st.markdown(f"<div class='box'>{corpo_preview}</div>", unsafe_allow_html=True)

    if uploaded_file is not None:
        if uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=0)
        else:
            df = pd.read_csv(uploaded_file, header=0)

        df.columns = df.columns.str.strip().str.upper()
        st.subheader("Prévia dos dados carregados")
        st.write(df.head())

        if "E-MAIL" in df.columns and "RESPONSAVEL" in df.columns:
            st.subheader("📌 Preview dos primeiros 5 envios")
            preview_rows = []
            for _, row in df.head(5).iterrows():
                responsavel = str(row["RESPONSAVEL"])
                to, cc = parse_multiplos_emails(row["E-MAIL"])
                if not to:
                    continue
                assunto_p = assunto.replace("{{responsavel}}", responsavel)
                corpo_p = corpo_base.replace("{{responsavel}}", responsavel)
                corpo_p = corpo_p + ("<hr>" + assinatura_html if assinatura_html else "")
                destino_para = email_user if modo_teste else to
                destino_cc = [] if modo_teste else cc
                preview_rows.append({
                    "Para": destino_para,
                    "Cc": ", ".join(destino_cc) if destino_cc else "-",
                    "Assunto": assunto_p
                })
            st.table(preview_rows)

            if st.button("🚀 Enviar todos os e-mails"):
                if not email_user or not email_pass:
                    st.error("Informe seu e-mail e senha do Skymail.")
                else:
                    enviados, falhas = 0, 0
                    for _, row in df.iterrows():
                        responsavel = str(row["RESPONSAVEL"])
                        to, cc = parse_multiplos_emails(row["E-MAIL"])
                        if not to:
                            continue
                        assunto_p = assunto.replace("{{responsavel}}", responsavel)
                        corpo_p = corpo_base.replace("{{responsavel}}", responsavel)
                        corpo_p = corpo_p + ("<hr>" + assinatura_html if assinatura_html else "")

                        destino_para = email_user if modo_teste else to
                        destino_cc = [] if modo_teste else cc
                        destino_bcc = [] if modo_teste else bcc_list

                        try:
                            enviar_email(email_user, email_pass, destino_para, destino_cc, destino_bcc, assunto_p, corpo_p)
                            enviados += 1
                            st.write(f"✅ Enviado: {destino_para}")
                            time.sleep(pausa)
                        except Exception as ex:
                            falhas += 1
                            st.write(f"⚠️ Erro com {to}: {ex}")

                    st.success(f"Finalizado. Total enviados: {enviados} | Falhas: {falhas}")
        else:
            st.error("A planilha precisa ter as colunas: E-MAIL e RESPONSAVEL")

# ================================
# CONTEÚDO ABA ASSINATURAS
# ================================
with aba_assinaturas:
    st.markdown("<h2>🖊️ Assinaturas</h2>", unsafe_allow_html=True)

    assinatura_padrao = ASSINATURA_USUARIO.get(st.session_state.usuario, "")

    tab_catalogo, tab_upload, tab_url = st.tabs(["📚 Catálogo", "📤 Upload", "🔗 URL"])

    with tab_catalogo:
        nomes = list(ASSINATURAS.keys())
        escolha = st.selectbox("Escolha do catálogo", nomes, index=0)
        html_escolha = ASSINATURAS[escolha]
        st.markdown(html_escolha, unsafe_allow_html=True)
        if st.button("Usar esta assinatura"):
            st.session_state.assinatura_html = html_escolha
            st.success("Assinatura selecionada!")

    with tab_upload:
        up = st.file_uploader("Envie uma imagem (PNG/JPG)", type=["png","jpg","jpeg"])
        if up:
            data_uri = img_to_data_uri(up)
            st.markdown(data_uri, unsafe_allow_html=True)
            if st.button("Usar a imagem enviada"):
                st.session_state.assinatura_html = data_uri
                st.success("Assinatura personalizada definida!")

    with tab_url:
        url = st.text_input("Cole o URL da imagem (https://...)", "")
        if url:
            prev = f"<img src='{url}' width='450'>"
            st.markdown(prev, unsafe_allow_html=True)
        if st.button("Usar a imagem do URL"):
            if url:
                st.session_state.assinatura_html = f"<img src='{url}' width='450'>"
                st.success("Assinatura via URL definida!")

    st.divider()
    atual = st.session_state.get("assinatura_html", assinatura_padrao)
    st.markdown("**Assinatura em uso (pré-visualização):**")
    if atual:
        st.markdown(atual, unsafe_allow_html=True)
    else:
        st.info("Nenhuma assinatura definida.")

    if st.button("Reverter para assinatura padrão do meu usuário"):
        st.session_state.assinatura_html = assinatura_padrao
        st.success("Assinatura padrão aplicada!")

# ================================
# PAINEL ADMIN (opcional)
# ================================
ADMINS = ["gabryell@acelnet.com.br", "marcio@acelnet.com.br", "leonardo@acelnet.com.br", "victor@acelnet.com.br"]
if st.session_state.usuario in ADMINS:
    with st.expander("🛠️ Painel Administrativo"):
        st.write("Recursos futuros (logs, limites, etc.)")
