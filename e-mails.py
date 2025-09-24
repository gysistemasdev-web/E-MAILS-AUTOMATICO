# emails.py — ACEL | Envio Automático de E-mails (com prévias + CC Global + Anexos)
# Execução: streamlit run emails.py

import os, re, json, time, smtplib, base64, shutil
import pandas as pd
from json import JSONDecodeError
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import streamlit as st

# ================================
# CONFIG SMTP (Skymail)
# ================================
SMTP_HOST = "smtp.skymail.net.br"
SMTP_PORT = 465  # SSL/TLS

def enviar_email(email_user, email_pass, para, cc_list, bcc_list, assunto, corpo_html, anexos=None):
    msg = MIMEMultipart("mixed")  # mixed permite anexos
    msg["From"] = email_user
    msg["To"] = para
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = assunto

    # corpo do e-mail em HTML
    corpo = MIMEMultipart("alternative")
    corpo.attach(MIMEText(corpo_html, "html"))
    msg.attach(corpo)

    # anexos
    if anexos:
        for file in anexos:
            try:
                file.seek(0)  # garante leitura desde o início
                data = file.read()
                part = MIMEApplication(data, Name=file.name)
                part.add_header('Content-Disposition', 'attachment', filename=file.name)
                msg.attach(part)
            except Exception as e:
                print(f"Erro ao anexar {file.name}: {e}")

    dest_all = [para] + cc_list + bcc_list
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(email_user, email_pass)
        server.sendmail(email_user, dest_all, msg.as_string())

# ================================
# LOGIN / REGISTRO (robusto)
# ================================
USERS_FILE = "usuarios.json"

def carregar_usuarios():
    if not os.path.exists(USERS_FILE):
        return {}
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                return {}
            return json.loads(content)
    except JSONDecodeError:
        try:
            shutil.copyfile(USERS_FILE, USERS_FILE + ".bad")
        except Exception:
            pass
        return {}
    except Exception:
        return {}

def salvar_usuarios(users):
    tmp = USERS_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, USERS_FILE)

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
aba_envio, aba_assinaturas, aba_ajuda = st.tabs(["📧 Envio de E-mails", "🖊️ Assinaturas", "ℹ️ Como usar o sistema"])

# ================================
# ASSINATURAS (catálogo)
# ================================
ASSINATURAS = {
    "Leonardo": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/LEONARDO.png" width="450">""",
    "Erika": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/ERIKA%201.png" width="450">""",
    "Caroline": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/main/CAROLINE.png" width="450">""",
    "Halline": """<img src="https://raw.githubusercontent.com/gysistemasdev-web/E-MAILS-AUTOMATICO/f7cbeee80b479c47cd3d09f8e4377c4b560dad2a/HALLINE.png" width="450">""",
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
# ABA ENVIO — com PRÉVIAS, CC Global e ANEXOS
# ================================
with aba_envio:
    # Login
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

        cc_global_raw = st.text_input("CC Global (vai em todos os envios)", "")
        cc_global = [e.strip() for e in re.split(r"[;,\s]+", cc_global_raw) if "@" in e]

        bcc_raw = st.text_input("BCC (oculto, opcional)", "")
        bcc_list = [e.strip() for e in re.split(r"[;,\s]+", bcc_raw) if "@" in e]

    with colB:
        uploaded_file = st.file_uploader("Carregue a planilha (.xlsx ou .csv)", type=["xlsx","csv"])
        anexos = st.file_uploader("Adicionar anexos (opcional)", type=None, accept_multiple_files=True)
        st.subheader("Corpo do e-mail")
        texto_puro = st.text_area("Digite aqui:", "Bom dia,\n\nPrezados(as),\n\n...", height=260)

    corpo_base = converter_para_html(texto_puro)
    corpo_preview = corpo_base + ("<hr>" + assinatura_html if assinatura_html else "")
    st.subheader("📌 HTML gerado")
    st.code(corpo_preview, language="html")
    st.subheader("🔎 Prévia formatada")
    st.markdown(f"<div class='box'>{corpo_preview}</div>", unsafe_allow_html=True)

    df = None
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, header=0) if uploaded_file.name.endswith(".xlsx") else pd.read_csv(uploaded_file, header=0)
        df.columns = df.columns.str.strip().str.upper()

        st.subheader("📂 Prévia dos dados carregados")
        st.write(df.head())

        if {"E-MAIL", "RESPONSAVEL"}.issubset(df.columns):
            st.subheader("📌 Preview dos primeiros 5 envios")
            preview_rows = []
            for _, row in df.head(5).iterrows():
                responsavel = str(row["RESPONSAVEL"])
                to, cc = parse_multiplos_emails(row["E-MAIL"])
                if not to:
                    continue
                assunto_p = assunto.replace("{{responsavel}}", responsavel)
                destino_para = email_user if modo_teste else to
                destino_cc = [] if modo_teste else (cc + cc_global)
                preview_rows.append({
                    "Para": destino_para,
                    "Cc": ", ".join(destino_cc) if destino_cc else "-",
                    "Assunto": assunto_p,
                    "Anexos": ", ".join([f.name for f in anexos]) if anexos else "-"
                })
            st.table(preview_rows)
        else:
            st.error("A planilha precisa ter as colunas: E-MAIL e RESPONSAVEL")

    enviar_click = st.button("🚀 Enviar todos os e-mails", use_container_width=True)
    if enviar_click:
        if not email_user or not email_pass:
            st.error("Informe seu e-mail e senha do Skymail.")
        elif uploaded_file is None:
            st.error("Carregue a planilha (.xlsx ou .csv).")
        elif df is None or not {"E-MAIL", "RESPONSAVEL"}.issubset(df.columns):
            st.error("A planilha precisa ter as colunas: E-MAIL e RESPONSAVEL.")
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
                destino_cc = [] if modo_teste else (cc + cc_global)
                destino_bcc = [] if modo_teste else bcc_list

                try:
                    enviar_email(email_user, email_pass, destino_para, destino_cc, destino_bcc, assunto_p, corpo_p, anexos=anexos)
                    enviados += 1
                    st.write(f"✅ Enviado: {destino_para}" + (f" | Cc: {', '.join(destino_cc)}" if destino_cc else "") + (f" | Anexos: {', '.join([f.name for f in anexos])}" if anexos else ""))
                    time.sleep(pausa)
                except Exception as ex:
                    falhas += 1
                    st.write(f"⚠️ Erro com {destino_para}: {ex}")
            st.success(f"Finalizado. Total enviados: {enviados} | Falhas: {falhas}")

# ================================
# ABA ASSINATURAS
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
        up = st.file_uploader("Envie uma imagem (PNG/JPG)", type=["png","jpg","jpeg"], key="upl_assin")
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

    atual = st.session_state.get("assinatura_html", assinatura_padrao)
    st.markdown("**Assinatura em uso:**")
    if atual:
        st.markdown(atual, unsafe_allow_html=True)
    if st.button("Reverter para assinatura padrão"):
        st.session_state.assinatura_html = assinatura_padrao
        st.success("Assinatura padrão aplicada!")

# ================================
# ABA AJUDA
# ================================
with aba_ajuda:
    st.markdown("<h2>ℹ️ Como usar o sistema</h2>", unsafe_allow_html=True)
    st.markdown("""
### 📝 Formatação do texto
- **negrito** → escreva entre duas estrelas: **assim**
- ##vermelho## → escreva entre hashtags duplas: ##assim##

### 📂 Planilha
- Colunas obrigatórias: **E-MAIL** e **RESPONSAVEL**
- Vários e-mails na mesma célula → separe por ; , ou espaço
- O primeiro e-mail vira **Para**, os demais vão em **Cc**

### ✉️ Envio
- **Modo Teste** envia só para o seu e-mail
- Prévia mostra as 5 primeiras linhas da planilha e os 5 primeiros envios
- Intervalo ajustável entre disparos
- **CC Global**: todos os envios incluem esse e-mail em cópia
- **BCC**: cópia oculta opcional em todos os envios
- **Anexos**: todos os anexos carregados vão para todos os e-mails

### 🖊️ Assinaturas
- Cada usuário tem assinatura padrão vinculada
- Pode trocar no catálogo, enviar uma imagem ou usar uma URL
- Assinatura aparece no final do e-mail automaticamente
""")
