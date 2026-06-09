import streamlit as st
import pandas as pd
import os
import re
import zipfile
from io import BytesIO
from datetime import datetime

from formatador import (
    limpar_nome,
    normalizar_nome,
    encontrar_coluna_nome,
    encontrar_coluna_email,
    ler_arquivo,
    aplicar_filtro,
    gerar_textos,
    exportar_excel,
    gerar_pdf,
    gerar_pdf_bytes,
    salvar_rascunho,
    enviar_email_smtp
)

# =========================
# CONFIG
# =========================
st.set_page_config(
    page_title="Formatador LACOM Jr.",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================
# ESTADO DO TEMA
# =========================
if "theme" not in st.session_state:
    st.session_state.theme = "dark"

# =========================
# ESTILO — tema escuro (padrão)
# =========================
st.markdown(
    f"""
    <style>

    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    :root {{
        --bg-page: #0B1A1A;
        --bg-sidebar: #081414;
        --bg-card: #112828;
        --bg-card-alt: #1E2A2A;
        --bg-hover: #163E3E;
        --bg-hover-alt: #2A3E3E;
        --bg-alert: #162828;
        --bg-active: #0E3030;

        --text-primary: #D1FAE5;
        --text-secondary: #A7F3D0;
        --text-heading: #2DD4BF;
        --text-accent: #2DD4BF;
        --text-tab: #5EEAD4;
        --text-input: #FFFFFF;

        --border: #2A4040;
        --border-muted: #1A3535;
        --border-accent: #2DD4BF;
        --border-alert: #2A4848;
        --border-dashed: #2A4848;
    }}

    html, body, .stApp, .stMarkdown {{
        font-family: 'Inter', sans-serif;
        color: var(--text-primary);
    }}
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {{
        color: var(--text-heading);
    }}

    .stApp {{
        background-color: var(--bg-page);
    }}

    .block-container {{
        padding-top: 0rem !important;
    }}

    section[data-testid="stSidebar"] {{
        background-color: var(--bg-sidebar);
        border-right: 1px solid var(--border-muted);
    }}

    section[data-testid="stSidebar"] * {{
        color: var(--text-secondary) !important;
    }}

    .stTabs [data-baseweb="tab-list"] {{
        background-color: var(--bg-page);
        border-bottom: 2px solid var(--border-muted);
        gap: 6px;
        padding-top: 12px;
    }}

    .stTabs [data-baseweb="tab"] {{
        background-color: var(--bg-card);
        color: var(--text-tab) !important;
        border-radius: 10px 10px 0 0;
        padding: 10px 30px;
        font-weight: 600;
        font-size: 14px;
        border: 1px solid var(--border-muted);
        border-bottom: none;
        transition: all 0.2s ease;
    }}

    .stTabs [aria-selected="true"] {{
        background-color: var(--bg-active) !important;
        color: var(--text-accent) !important;
        border-color: var(--border-accent) !important;
    }}

    .stTabs [data-baseweb="tab"]:hover {{
        background-color: var(--bg-hover) !important;
        color: var(--text-accent) !important;
    }}

    .stButton button {{
        background: linear-gradient(135deg, #0D9488, #06B6D4);
        color: white !important;
        border-radius: 10px;
        border: none;
        height: 48px;
        font-size: 15px;
        font-weight: 700;
        width: 100%;
        letter-spacing: 0.4px;
        transition: opacity 0.2s ease;
    }}

    .stButton button:hover {{
        opacity: 0.85;
        color: white !important;
    }}

    h1, h2, h3 {{
        color: var(--text-heading);
    }}

    .stTextInput input,
    .stTextArea textarea,
    .stSelectbox div[data-baseweb="select"],
    .stSelectbox div[data-baseweb="select"] > div,
    .stMultiSelect div[data-baseweb="select"] {{
        background-color: var(--bg-card-alt) !important;
        color: var(--text-input) !important;
        border-color: var(--border) !important;
        border-radius: 8px;
    }}

    [data-baseweb="popover"] li,
    [data-baseweb="menu"] li {{
        background-color: var(--bg-card-alt) !important;
        color: var(--text-input) !important;
    }}

    [data-baseweb="popover"] li:hover,
    [data-baseweb="menu"] li:hover {{
        background-color: var(--bg-hover-alt) !important;
    }}

    [data-testid="stFileUploader"] {{
        background-color: var(--bg-card-alt);
        border: 2px dashed var(--border-dashed);
        border-radius: 12px;
        padding: 16px;
    }}
    [data-testid="stFileUploader"] * {{
        color: var(--text-primary) !important;
    }}
    [data-testid="stFileUploader"] button {{
        background: var(--bg-card) !important;
        border: 1px solid var(--border) !important;
        color: var(--text-accent) !important;
        border-radius: 8px;
        height: 36px;
        font-weight: 600;
    }}

    .stDataFrame {{
        border-radius: 10px;
        overflow: hidden;
    }}

    hr {{
        border-color: var(--border-muted);
    }}

    .stCheckbox label,
    .stCheckbox label *,
    .stCheckbox .st-cb-label,
    .stCheckbox .st-cb-label *,
    .stCheckbox p,
    .stCheckbox span {{
        color: var(--text-primary) !important;
    }}
    .stCheckbox svg,
    .stCheckbox path {{
        fill: var(--text-primary) !important;
    }}

    div[data-testid="stAlert"] {{
        background-color: var(--bg-alert) !important;
        border: 1px solid var(--border-alert) !important;
        border-left: 4px solid var(--border-alert) !important;
        border-radius: 10px !important;
    }}

    div[data-testid="stAlert"] p,
    div[data-testid="stAlert"] span,
    div[data-testid="stAlert"] div,
    div[data-testid="stAlert"] svg {{
        color: var(--text-secondary) !important;
        fill: var(--text-secondary) !important;
    }}

    [data-testid="stDownloadButton"] button {{
        background: linear-gradient(135deg, #065F46, #0D9488) !important;
    }}

    label[data-testid="stWidgetLabel"] p,
    .stTextInput label,
    .stTextArea label,
    .stSelectbox label {{
        color: var(--text-secondary) !important;
        font-size: 13px;
    }}

    /* Botão de alternância de tema — sem fundo gradiente */
    [data-testid="stHorizontalBlock"]:first-of-type [data-testid="column"]:last-child .stButton button {{
        background: transparent !important;
        border: 1px solid var(--border-muted) !important;
        color: var(--text-secondary) !important;
        height: 36px !important;
        padding: 2px 14px !important;
        font-size: 13px !important;
        font-weight: 600 !important;
        width: auto !important;
        letter-spacing: 0.3px !important;
    }}
    [data-testid="stHorizontalBlock"]:first-of-type [data-testid="column"]:last-child .stButton button:hover {{
        background: var(--bg-hover) !important;
        opacity: 1 !important;
    }}

    </style>
    """,
    unsafe_allow_html=True
)

# =========================
# ESTILO — tema claro (sobrescreve variáveis)
# =========================
if st.session_state.theme == "light":
    st.markdown(
        """
        <style>
        :root {
            --bg-page: #D6E8E4;
            --bg-sidebar: #CCE0DC;
            --bg-card: #EAF5F2;
            --bg-card-alt: #EAF5F2;
            --bg-hover: #CCE0DC;
            --bg-hover-alt: #C2D8D4;
            --bg-alert: #E0F0EC;
            --bg-active: #EAF5F2;

            --text-primary: #0B1A1A;
            --text-secondary: #0B1A1A;
            --text-heading: #0B1A1A;
            --text-accent: #0B1A1A;
            --text-tab: #0B1A1A;
            --text-input: #0B1A1A;

            --border: #B0C8C4;
            --border-muted: #C0D6D2;
            --border-accent: #0B1A1A;
            --border-alert: #A0CCC4;
            --border-dashed: #B0C8C4;
        }
        .stCheckbox label,
        .stCheckbox label *,
        .stCheckbox .st-cb-label { color: #0B1A1A !important; }
        .stCheckbox svg,
        .stCheckbox path { fill: #0B1A1A !important; }
        </style>
        """,
        unsafe_allow_html=True
    )

# =========================
# PASTAS
# =========================
os.makedirs("saidas", exist_ok=True)
os.makedirs("historico", exist_ok=True)
os.makedirs("pdfs", exist_ok=True)

# =========================
# COVER (banner retangular)
# =========================
if os.path.exists("cover.png"):
    st.image("cover.png", width="stretch")
elif os.path.exists("banner.png"):
    st.image("banner.png", width="stretch")

# =========================
# HEADER: logo + titulo
# =========================
col_logo, col_title, col_theme = st.columns([1, 8, 1])

with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=72)

with col_title:
    st.markdown(
        """
        <div style='padding-top: 8px;'>
            <h1 style='
                margin: 0;
                font-size: 28px;
                font-weight: 700;
                line-height: 1.1;
            '>GESTÃO</h1>
            <span style='
                font-size: 11px;
                font-weight: 600;
                letter-spacing: 3px;
                text-transform: uppercase;
            '>LACOM Jr.</span>
        </div>
        """,
        unsafe_allow_html=True
    )

with col_theme:
    theme_label = "Claro" if st.session_state.theme == "light" else "Escuro"
    if st.button(theme_label, key="theme_toggle", help="Alternar tema claro/escuro"):
        st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"
        st.rerun()

st.divider()

# =========================
# SESSION STATE
# =========================
if "texto_certificado" not in st.session_state:
    st.session_state.texto_certificado = (
        "Certifico que {nome} participou da palestra."
    )

# =========================
# ABAS PRINCIPAIS
# =========================
aba_cert, aba_email, aba_hist = st.tabs([
    "Formatador de Certificados",
    "Formatador de E-mails",
    "Histórico"
])

# ======================================
# ABA 1 — FORMATADOR DE CERTIFICADOS
# ======================================
with aba_cert:

    st.subheader("📂 Upload da Planilha")

    arquivo = st.file_uploader(
        "Selecione uma planilha (.xlsx ou .csv)",
        type=["xlsx", "csv"],
        key="upload_cert"
    )

    if arquivo is not None:

        df = ler_arquivo(arquivo)
        coluna_nome = encontrar_coluna_nome(df.columns)

        if coluna_nome is None:
            st.error("Nenhuma coluna de nome encontrada.")
            st.stop()

        # Detecta coluna de email ANTES de modificar dados
        coluna_email = encontrar_coluna_email(df.columns)

        df[coluna_nome] = df[coluna_nome].apply(limpar_nome)
        df[coluna_nome] = df[coluna_nome].apply(normalizar_nome)

        st.info("Arquivo carregado com sucesso!")
        st.info(f"Coluna de nome encontrada: **{coluna_nome}**")

        if coluna_email is None:
            st.warning("Nenhuma coluna de email encontrada.")
        else:
            st.info(f"Coluna de email encontrada: **{coluna_email}**")

        # Filtros
        st.subheader("Filtros")

        usar_filtro = st.checkbox("Aplicar filtro", key="filtro_cert")

        if usar_filtro:
            coluna_filtro = st.selectbox(
                "Escolha a coluna para filtrar",
                df.columns,
                key="col_filtro_cert"
            )
            valores = (
                df[coluna_filtro]
                .dropna()
                .astype(str)
                .unique()
            )
            valor_filtro = st.selectbox(
                "Escolha o valor",
                valores,
                key="val_filtro_cert"
            )
            df = aplicar_filtro(df, coluna_filtro, valor_filtro)

        st.info(f"**{len(df)}** registros encontrados")

        # Preview
        st.subheader("Preview da Planilha")
        st.dataframe(df, width="stretch")

        # Texto certificado
        st.subheader("Texto do Certificado")

        texto_padrao = st.text_area(
            "Use {nome} para inserir o nome automaticamente",
            value=st.session_state.texto_certificado,
            height=160,
            key="texto_cert"
        )
        st.session_state.texto_certificado = texto_padrao

        gerar_pdfs = st.checkbox(
            "Gerar PDFs individuais",
            key="check_pdf"
        )

        signature_file = st.file_uploader(
            "Imagem da assinatura",
            type=["png", "jpg", "jpeg"],
            key="sig_upload"
        )

        sig_path = None
        if signature_file is not None:
            import tempfile
            sig_path = os.path.join(tempfile.gettempdir(), "_sig_temp.png")
            with open(sig_path, "wb") as f:
                f.write(signature_file.getbuffer())

        st.divider()

        # =========================
        # FLUXO DE GERAÇÃO
        # =========================
        if "gerou" not in st.session_state:
            st.session_state.gerou = False

        if st.button("Gerar Certificados", key="btn_cert"):
            with st.spinner("Gerando certificados..."):
                try:
                    nome_base = os.path.splitext(arquivo.name)[0]
                    nome_saida = f"saida_{nome_base}.xlsx"
                    caminho_saida = os.path.join("saidas", nome_saida)

                    textos = gerar_textos(df, coluna_nome, texto_padrao)
                    nomes_originais = df[coluna_nome].tolist()
                    emails_col = df[coluna_email].tolist() if coluna_email is not None else None
                    exportar_excel(textos, caminho_saida, nomes=nomes_originais, emails=emails_col)

                    zip_path = None
                    if gerar_pdfs:
                        zip_path = os.path.join("saidas", f"certificados_{nome_base}.zip")
                        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                            for texto, nome in zip(textos, df[coluna_nome]):
                                pdf_buf = gerar_pdf_bytes(
                                    nome, texto, signature_path=sig_path
                                )
                                nome_arquivo = re.sub(r'[\\/*?:"<>|]', "_", str(nome)) + ".pdf"
                                zf.writestr(nome_arquivo, pdf_buf.getvalue())
                                # Também salva individualmente para usar nos e-mails
                                pdf_indiv_path = os.path.join("pdfs", nome_arquivo)
                                os.makedirs("pdfs", exist_ok=True)
                                with open(pdf_indiv_path, "wb") as f:
                                    f.write(pdf_buf.getvalue())

                    # Historico
                    historico_path = os.path.join("historico", "historico.xlsx")
                    novo_historico = pd.DataFrame({
                        "Arquivo":    [arquivo.name],
                        "Data":       [datetime.now().strftime("%d/%m/%Y %H:%M")],
                        "Quantidade": [len(df)]
                    })
                    if os.path.exists(historico_path):
                        historico_antigo = pd.read_excel(historico_path)
                        novo_historico = pd.concat(
                            [historico_antigo, novo_historico],
                            ignore_index=True
                        )
                    novo_historico.to_excel(historico_path, index=False)

                    st.session_state.gerou = True
                    st.session_state.caminho_saida = caminho_saida
                    st.session_state.nome_saida = nome_saida
                    st.session_state.zip_path = zip_path
                    st.session_state.nome_base = nome_base
                except Exception as e:
                    import traceback
                    st.error(f"Erro ao gerar certificados: {e}")
                    st.code(traceback.format_exc())

        if st.session_state.get("gerou", False):
            st.success("Certificados gerados com sucesso!")

            col_d1, col_d2 = st.columns(2)

            with col_d1:
                if os.path.exists(st.session_state.caminho_saida):
                    with open(st.session_state.caminho_saida, "rb") as file:
                        st.download_button(
                            label="⬇ Baixar Planilha (.xlsx)",
                            data=file,
                            file_name=st.session_state.nome_saida,
                            mime=(
                                "application/vnd.openxmlformats-"
                                "officedocument.spreadsheetml.sheet"
                            ),
                            key="dl_excel"
                        )

            with col_d2:
                zip_path = st.session_state.get("zip_path")
                if zip_path and os.path.exists(zip_path):
                    with open(zip_path, "rb") as file:
                        st.download_button(
                            label="Baixar Certificados (ZIP)",
                            data=file,
                            file_name=f"certificados_{st.session_state.nome_base}.zip",
                            mime="application/zip",
                            key="dl_zip"
                        )



# ======================================
# ABA 2 — FORMATADOR DE E-MAILS
# ======================================
with aba_email:

    st.subheader("📂 Upload da Planilha")

    arquivo_email = st.file_uploader(
        "Selecione uma planilha (.xlsx ou .csv)",
        type=["xlsx", "csv"],
        key="upload_email"
    )

    if arquivo_email is not None:

        df_email = ler_arquivo(arquivo_email)
        coluna_nome_e = encontrar_coluna_nome(df_email.columns)

        if coluna_nome_e is None:
            st.error("Nenhuma coluna de nome encontrada.")
            st.stop()

        coluna_email_e = encontrar_coluna_email(df_email.columns)

        df_email[coluna_nome_e] = (
            df_email[coluna_nome_e].apply(limpar_nome)
        )
        df_email[coluna_nome_e] = (
            df_email[coluna_nome_e].apply(normalizar_nome)
        )

        st.success("Arquivo carregado com sucesso!")
        st.success(f"Coluna de nome encontrada: **{coluna_nome_e}**")

        if coluna_email_e is None:
            st.warning("Nenhuma coluna de email encontrada.")
        else:
            st.success(f"Coluna de email encontrada: **{coluna_email_e}**")

        # Filtros
        st.subheader("Filtros")

        usar_filtro_e = st.checkbox("Aplicar filtro", key="filtro_email")

        if usar_filtro_e:
            coluna_filtro_e = st.selectbox(
                "Escolha a coluna para filtrar",
                df_email.columns,
                key="col_filtro_email"
            )
            valores_e = (
                df_email[coluna_filtro_e]
                .dropna()
                .astype(str)
                .unique()
            )
            valor_filtro_e = st.selectbox(
                "Escolha o valor",
                valores_e,
                key="val_filtro_email"
            )
            df_email = aplicar_filtro(
                df_email, coluna_filtro_e, valor_filtro_e
            )

        st.success(f"**{len(df_email)}** registros encontrados")

        # Preview
        st.subheader("Preview da Planilha")
        st.dataframe(df_email, width="stretch")

        # Config email
        st.subheader("Configuracao do E-mail")

        assunto_email = st.text_input(
            "Assunto do email",
            value="Certificado de Atividade Complementar",
            key="assunto"
        )

        mensagem_email = st.text_area(
            "Mensagem do email — use {nome} para personalizar",
            value=(
                "Parabens, {nome}!\n\n"
                "Segue em anexo seu certificado de participacão.\n\n"
                "Atenciosamente,\n"
                "Seu nome ou equipe"
            ),
            height=200,
            key="mensagem"
        )

        col_e1, col_e2 = st.columns(2)

        with col_e1:
            email_remetente = st.text_input(
                "Seu Gmail",
                key="gmail"
            )

        with col_e2:
            senha_email = st.text_input(
                "Senha de aplicativo",
                type="password",
                key="senha"
            )


        anexar_pdf = st.checkbox(
            "Anexar PDF do certificado",
            value=True,
            key="check_anexar"
        )

        st.divider()

        # --- Salvar rascunhos + enviar ---
        if coluna_email_e is None:
            st.error("Impossivel continuar: coluna de email nao encontrada.")
        elif not email_remetente or not senha_email:
            st.info("Preencha o Gmail e a senha de aplicativo.")
        else:
            textos_email = gerar_textos(
                df_email,
                coluna_nome_e,
                st.session_state.get(
                    "texto_certificado",
                    "Certifico que {nome} participou da palestra, com duração de 2h, ministrada por Rafael Sant’anna, promovida pela LACOM Jr, Empresa Júnior de Comunicação Digital da Fundação Getúlio Vargas, no dia 06 de maio de 2026."
                )
            )

            # Indexar todos os PDFs da pasta (case-insensitive)
            pdf_index = {}
            if anexar_pdf and os.path.isdir("pdfs"):
                for f in os.listdir("pdfs"):
                    if f.lower().endswith(".pdf"):
                        stem = os.path.splitext(f)[0]
                        chave = re.sub(r'[\\/*?:"<>|]', "_", stem).lower()
                        pdf_index[chave] = os.path.join("pdfs", f)

            # Prepara lista de rascunhos
            rascunhos = []
            tem_noname = False
            for texto, nome, email_dest in zip(
                textos_email, df_email[coluna_nome_e], df_email[coluna_email_e]
            ):
                nome_str = str(nome).strip() if pd.notna(nome) else ""
                if not nome_str:
                    tem_noname = True
                    continue

                chave = re.sub(r'[\\/*?:"<>|]', "_", nome_str).lower()
                if not chave:
                    chave = "certificado"

                # Busca no indice (case-insensitive)
                pdf_a_anexar = pdf_index.get(chave)

                rascunhos.append({
                    "nome": nome_str,
                    "email": email_dest,
                    "texto": texto,
                    "pdf": pdf_a_anexar,
                    "tem_pdf": pdf_a_anexar is not None
                })

            if tem_noname:
                st.warning("Alguns registros estao sem nome e foram ignorados.")

            if not rascunhos:
                st.warning("Nenhum rascunho valido para processar.")
                st.stop()

            st.success(f"**{len(rascunhos)}** e-mail(s) prontos.")

            # Estado para controlar se rascunhos ja foram salvos
            if "rascunhos_salvos" not in st.session_state:
                st.session_state.rascunhos_salvos = False
            if "envio_concluido" not in st.session_state:
                st.session_state.envio_concluido = False

            # --- Botao 1: Salvar rascunhos ---
            if st.button(
                "Salvar E-mails como rascunho",
                key="btn_salvar",
                disabled=st.session_state.rascunhos_salvos
            ):
                erros = []
                salvos = 0

                with st.spinner("Salvando rascunhos no Gmail..."):
                    for r in rascunhos:
                        try:
                            salvar_rascunho(
                                email_destino=r["email"],
                                nome=r["nome"],
                                assunto=assunto_email,
                                mensagem=mensagem_email,
                                arquivo_pdf=r["pdf"] if r["tem_pdf"] else None,
                                email_remetente=email_remetente,
                                senha=senha_email
                            )
                            salvos += 1
                        except Exception as e:
                            erros.append(f"{r['nome']}: {e}")

                if erros:
                    st.warning(
                        f"{len(erros)} erro(s). Salvos: {salvos}. "
                        + " | ".join(erros[:3])
                    )
                if salvos > 0:
                    st.session_state.rascunhos_salvos = True
                    st.success("Rascunhos salvos com sucesso!")

            # --- Botao 2: Enviar (so aparece apos salvar rascunhos) ---
            if st.session_state.rascunhos_salvos and not st.session_state.envio_concluido:

                if st.button(
                    "Enviar todos os e-mails",
                    key="btn_enviar"
                ):
                    erros = []
                    enviados = 0

                    with st.spinner("Enviando e-mails..."):
                        for r in rascunhos:
                            try:
                                enviar_email_smtp(
                                    email_destino=r["email"],
                                    nome=r["nome"],
                                    assunto=assunto_email,
                                    mensagem=mensagem_email,
                                    arquivo_pdf=r["pdf"] if r["tem_pdf"] else None,
                                    email_remetente=email_remetente,
                                    senha=senha_email
                                )
                                enviados += 1
                            except Exception as e:
                                erros.append(f"{r['nome']}: {e}")

                    if erros:
                        st.warning(
                            f"{len(erros)} erro(s). Enviados: {enviados}. "
                            + " | ".join(erros[:3])
                        )
                    if enviados > 0:
                        st.session_state.envio_concluido = True
                        st.success(
                            f"{enviados} e-mail(s) enviado(s) com sucesso!"
                        )

            if st.session_state.envio_concluido:
                st.success("Todos os e-mails foram enviados!")

# ======================================
# ABA 3 — HISTORICO
# ======================================
with aba_hist:

    st.subheader("Histórico de Geração de Certificados")

    historico_path = os.path.join("historico", "historico.xlsx")

    if os.path.exists(historico_path):
        historico_df = pd.read_excel(historico_path)
        st.dataframe(historico_df, width="stretch")
    else:
        st.info(
            "Nenhum historico encontrado ainda. "
            "Gere certificados para registrar aqui."
        )
