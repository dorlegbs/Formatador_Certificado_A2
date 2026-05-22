import streamlit as st
import pandas as pd
import os
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
# ESTILO
# =========================
st.markdown(
    """
    <style>

    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        color: #D1FAE5;
    }

    .stApp {
        background-color: #0B1A1A;
    }

    /* Remover padding topo */
    .block-container {
        padding-top: 0rem !important;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #081414;
        border-right: 1px solid #1A3A3A;
    }

    section[data-testid="stSidebar"] * {
        color: #A7F3D0 !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #0B1A1A;
        border-bottom: 2px solid #1A3535;
        gap: 6px;
        padding-top: 12px;
    }

    .stTabs [data-baseweb="tab"] {
        background-color: #112828;
        color: #5EEAD4 !important;
        border-radius: 10px 10px 0 0;
        padding: 10px 30px;
        font-weight: 600;
        font-size: 14px;
        border: 1px solid #1A3535;
        border-bottom: none;
        transition: all 0.2s ease;
    }

    .stTabs [aria-selected="true"] {
        background-color: #0E3030 !important;
        color: #2DD4BF !important;
        border-color: #2DD4BF !important;
    }

    .stTabs [data-baseweb="tab"]:hover {
        background-color: #163E3E !important;
        color: #2DD4BF !important;
    }

    /* Botao principal */
    .stButton button {
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
    }

    .stButton button:hover {
        opacity: 0.85;
        color: white !important;
    }

    /* Headings */
    h1, h2, h3 {
        color: #2DD4BF;
    }

    /* =====================
       INPUTS — cinza escuro
       texto branco
    ===================== */
    .stTextInput input,
    .stTextArea textarea,
    .stSelectbox div[data-baseweb="select"],
    .stSelectbox div[data-baseweb="select"] > div,
    .stMultiSelect div[data-baseweb="select"] {
        background-color: #1E2A2A !important;
        color: #FFFFFF !important;
        border-color: #2A4040 !important;
        border-radius: 8px;
    }

    /* Dropdown options */
    [data-baseweb="popover"] li,
    [data-baseweb="menu"] li {
        background-color: #1E2A2A !important;
        color: #FFFFFF !important;
    }

    [data-baseweb="popover"] li:hover,
    [data-baseweb="menu"] li:hover {
        background-color: #2A3E3E !important;
    }

    /* File uploader */
    [data-testid="stFileUploader"] {
        background-color: #1E2A2A;
        border: 2px dashed #2A4848;
        border-radius: 12px;
        padding: 16px;
    }

    /* Dataframe */
    .stDataFrame {
        border-radius: 10px;
        overflow: hidden;
    }

    /* Divider */
    hr {
        border-color: #1A3535;
    }

    /* Checkbox */
    .stCheckbox label {
        color: #A7F3D0 !important;
        font-size: 14px;
    }

    /* =====================
       ALERTAS UNIFICADOS
       cinza-esverdeado escuro
    ===================== */
    div[data-testid="stAlert"] {
        background-color: #162828 !important;
        border: 1px solid #2A4848 !important;
        border-left: 4px solid #2A4848 !important;
        border-radius: 10px !important;
    }

    div[data-testid="stAlert"] p,
    div[data-testid="stAlert"] span,
    div[data-testid="stAlert"] div,
    div[data-testid="stAlert"] svg {
        color: #A7F3D0 !important;
        fill: #A7F3D0 !important;
    }

    /* Download button */
    [data-testid="stDownloadButton"] button {
        background: linear-gradient(135deg, #065F46, #0D9488) !important;
    }

    /* Labels dos inputs */
    label[data-testid="stWidgetLabel"] p,
    .stTextInput label,
    .stTextArea label,
    .stSelectbox label {
        color: #A7F3D0 !important;
        font-size: 13px;
    }

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
    st.image("cover.png", use_container_width=True)
elif os.path.exists("banner.png"):
    st.image("banner.png", use_container_width=True)

# =========================
# HEADER: logo + titulo
# =========================
col_logo, col_title = st.columns([1, 9])

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
                color: #2DD4BF;
                line-height: 1.1;
            '>GESTÃO</h1>
            <span style='
                font-size: 11px;
                font-weight: 600;
                letter-spacing: 3px;
                color: #5EEAD4;
                text-transform: uppercase;
            '>LACOM Jr.</span>
        </div>
        """,
        unsafe_allow_html=True
    )

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

        st.success("Arquivo carregado com sucesso!")
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
        st.dataframe(df, use_container_width=True)

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
            "Imagem da assinatura (opcional)",
            type=["png", "jpg", "jpeg"],
            key="sig_upload"
        )

        if signature_file is not None:
            import tempfile
            sig_path = os.path.join(tempfile.gettempdir(), "_sig_temp.png")
            with open(sig_path, "wb") as f:
                f.write(signature_file.getbuffer())
        else:
            sig_path = os.path.join(
                os.path.dirname(__file__), "assinatura_vitoria.png"
            )
            if not os.path.exists(sig_path):
                sig_path = None

        st.divider()

        # =========================
        # FLUXO DE GERAÇÃO
        # =========================
        if "mostrar_pastas" not in st.session_state:
            st.session_state.mostrar_pastas = False
        if "gerou" not in st.session_state:
            st.session_state.gerou = False

        if st.button("Gerar Certificados", key="btn_cert"):
            st.session_state.mostrar_pastas = True
            st.session_state.gerou = False

        if st.session_state.mostrar_pastas:
            st.markdown("### 📁 Escolher pastas de destino")

            pasta_excel = st.text_input(
                "Pasta para salvar a planilha Excel (deixe em branco para usar 'saidas/')",
                key="pasta_excel"
            )

            if gerar_pdfs:
                pasta_pdf = st.text_input(
                    "Pasta para salvar os PDFs (deixe em branco para usar 'pdfs/')",
                    key="pasta_pdf"
                )

            col_ok, col_no = st.columns(2)
            with col_ok:
                if st.button("✅ Confirmar e Gerar", key="btn_confirm"):
                    try:
                        nome_base = os.path.splitext(arquivo.name)[0]
                        nome_saida = f"saida_{nome_base}.xlsx"

                        if pasta_excel.strip():
                            caminho_saida = os.path.join(pasta_excel, nome_saida)
                        else:
                            caminho_saida = os.path.join("saidas", nome_saida)

                        textos = gerar_textos(df, coluna_nome, texto_padrao)
                        exportar_excel(textos, caminho_saida)

                        if gerar_pdfs:
                            pdf_dir = pasta_pdf.strip() if pasta_pdf.strip() else None
                            for texto, nome in zip(textos, df[coluna_nome]):
                                gerar_pdf(nome, texto, signature_path=sig_path, output_dir=pdf_dir)

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
                        st.session_state.mostrar_pastas = False
                    except Exception as e:
                        import traceback
                        st.error(f"Erro ao gerar certificados: {e}")
                        st.code(traceback.format_exc())

            with col_no:
                if st.button("❌ Cancelar", key="btn_cancel"):
                    st.session_state.mostrar_pastas = False

        if st.session_state.gerou:
            st.success("Certificados gerados com sucesso!")
            with open(st.session_state.caminho_saida, "rb") as file:
                st.download_button(
                    label="⬇️ Baixar planilha gerada",
                    data=file,
                    file_name=st.session_state.nome_saida,
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    )
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
        st.dataframe(df_email, use_container_width=True)

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
                "Segue em anexo seu certificado de participacao.\n\n"
                "Atenciosamente,\n"
                "Equipe LACOM Jr."
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
            import re as _re
            pdf_index = {}
            if anexar_pdf and os.path.isdir("pdfs"):
                for f in os.listdir("pdfs"):
                    if f.lower().endswith(".pdf"):
                        stem = os.path.splitext(f)[0]
                        chave = _re.sub(r'[\\/*?:"<>|]', "_", stem).lower()
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

                chave = _re.sub(r'[\\/*?:"<>|]', "_", nome_str).lower()
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
        st.dataframe(historico_df, use_container_width=True)
    else:
        st.info(
            "Nenhum historico encontrado ainda. "
            "Gere certificados para registrar aqui."
        )
