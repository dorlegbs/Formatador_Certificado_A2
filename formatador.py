import pandas as pd
import re
import os
import time
import imaplib
import smtplib
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header


# =========================
# LIMPAR NOMES
# =========================
def limpar_nome(nome):

    if pd.isna(nome):
        return ""

    nome = str(nome)

    nome = nome.strip()

    nome = re.sub(r"\s+", " ", nome)

    return nome


# =========================
# NORMALIZAR NOMES
# (primeira maiúscula, preposições minúsculas)
# =========================
_PREPOSICOES = {"da", "das", "de", "do", "dos", "e"}

def normalizar_nome(nome):

    if pd.isna(nome) or not str(nome).strip():
        return ""

    nome = str(nome).strip().lower()

    palavras = nome.split()
    resultado = []

    for i, palavra in enumerate(palavras):

        # Preposições sempre minúsculas
        if palavra in _PREPOSICOES:
            resultado.append(palavra)
            continue

        # Palavras com hífen: capitaliza cada parte
        if "-" in palavra:
            partes = palavra.split("-")
            partes = [p.capitalize() for p in partes]
            resultado.append("-".join(partes))
        else:
            resultado.append(palavra.capitalize())

    return " ".join(resultado)


# =========================
# ENCONTRAR COLUNA NOME
# =========================
def encontrar_coluna_nome(colunas):

    for coluna in colunas:

        coluna_formatada = (
            str(coluna)
            .lower()
            .strip()
        )

        if "nome" in coluna_formatada:

            return coluna

    return None


# =========================
# ENCONTRAR COLUNA EMAIL
# Aceita qualquer variação:
# "email", "e-mail", "e-mails",
# "E-mail dos participantes", etc.
# =========================
def encontrar_coluna_email(colunas):

    for coluna in colunas:

        normalizado = (
            str(coluna)
            .lower()
            .strip()
            .replace("-", "")
            .replace(" ", "")
        )

        # Aceita: email, e-mail, mail, correio
        if any(k in normalizado for k in ("email", "mail", "correio")):

            return coluna

    return None


# =========================
# LER ARQUIVO
# =========================
def ler_arquivo(arquivo):

    if arquivo.name.endswith(".csv"):

        df = pd.read_csv(arquivo)

    else:

        df = pd.read_excel(arquivo)

    return df


# =========================
# APLICAR FILTRO
# =========================
def aplicar_filtro(
    df,
    coluna,
    valor
):

    df = df[
        df[coluna]
        .astype(str)
        .str.contains(
            valor,
            case=False,
            na=False,
            regex=False
        )
    ]

    return df


# =========================
# GERAR TEXTOS
# =========================
def gerar_textos(
    df,
    coluna_nome,
    texto_padrao
):

    novos_textos = []

    for nome in df[coluna_nome]:

        texto_final = texto_padrao.replace(
            "{nome}",
            nome
        )

        novos_textos.append(texto_final)

    return novos_textos


# =========================
# EXPORTAR EXCEL
# =========================
def exportar_excel(
    textos,
    nome_saida
):

    novo_df = pd.DataFrame({

        "Nomes": textos

    })

    novo_df.to_excel(
        nome_saida,
        index=False
    )





# =========================
# GERAR PDF — MODELO LACOM
# =========================
def _registrar_fontes():
    """Registra Garet e Nunito se disponiveis, retorna (body, bold)."""
    font_dir = os.path.dirname(os.path.abspath(__file__))

    body = "Helvetica"
    bold = "Helvetica-Bold"

    # --- Garet Bold (título + assinatura) ---
    heavy_path = os.path.join(font_dir, "Garet-Heavy.ttf")
    if os.path.exists(heavy_path):
        try:
            pdfmetrics.registerFont(TTFont("Garet-Heavy", heavy_path))
            bold = "Garet-Heavy"
        except Exception:
            pass

    # --- Nunito Regular (corpo do texto) ---
    nunito_path = os.path.join(font_dir, "nunito.regular.ttf")
    if not os.path.exists(nunito_path):
        nunito_path = os.path.join(font_dir, "Nunito-Regular.ttf")
    if os.path.exists(nunito_path):
        try:
            pdfmetrics.registerFont(TTFont("Nunito-Regular", nunito_path))
            body = "Nunito-Regular"
        except Exception:
            pass

    # Fallback: se Nunito nao existe, usa Garet-Book
    if body == "Helvetica":
        book_path = os.path.join(font_dir, "Garet-Book.ttf")
        if os.path.exists(book_path):
            try:
                pdfmetrics.registerFont(TTFont("Garet-Book", book_path))
                body = "Garet-Book"
            except Exception:
                pass

    return body, bold


def gerar_pdf(
    nome,
    texto,
    signature_path=None,
    bg_path=None
):

    os.makedirs("pdfs", exist_ok=True)

    nome = nome.strip() if nome else ""
    if not nome:
        nome = "SemNome"
    nome_arquivo = re.sub(r'[\\/*?:"<>|]', "_", nome)
    if not nome_arquivo.strip("._ "):
        nome_arquivo = "SemNome"

    caminho_pdf = os.path.join("pdfs", f"{nome_arquivo}.pdf")

    # Registra fontes
    font_body, font_bold = _registrar_fontes()

    W, H = landscape(A4)  # 841,89 x 595,28 pt  (29,7 x 21 cm)

    c = canvas.Canvas(caminho_pdf, pagesize=landscape(A4))

    # ============================================================
    #  FUNDO — imagem do template
    # ============================================================
    if bg_path is None:
        bg_path = "Fundo_Lacom_certificado.png"

    if os.path.exists(bg_path):
        c.drawImage(bg_path, 0, 0, width=W, height=H, preserveAspectRatio=False)
    else:
        c.setFillColor(colors.white)
        c.rect(0, 0, W, H, fill=1, stroke=0)

    # ============================================================
    #  TÍTULO "CERTIFICADO"
    #  Modelo: x=7,97cm  y=5,45cm  w=13,75cm  h=2,29cm
    #  Fonte: Garet Bold, 54,4pt  Cor: #1f2b5b
    #  Em pontos (origin reportlab = bottom-left):
    #    centro_x = 7,97 + 13,75/2 = 14,845 cm
    #    box_bottom = 21 - 5,45 - 2,29 = 13,26 cm
    # ============================================================
    titulo_center_x = 148.45 * mm
    titulo_box_bottom = 132.6 * mm
    titulo_baseline = titulo_box_bottom + 54.4 * 0.33 * 0.3528 * mm

    c.setFont(font_bold, 54.4)
    c.setFillColor(colors.HexColor("#1f2b5b"))
    c.drawCentredString(titulo_center_x, titulo_baseline, "CERTIFICADO")

    # ============================================================
    #  CORPO DO TEXTO
    #  Modelo: Nunito Regular, 21pt, tracking 12, line-height 1,392
    #  Caixa: x=2,1cm  y=9,12cm  w=25,5cm  h=3,99cm
    #    centro_x = 2,1 + 25,5/2 = 14,85 cm
    #    box_center (vertical) = 21 - 9,12 - 3,99/2 = 9,885 cm
    # ============================================================
    c.setFont(font_body, 21)
    c.setFillColor(colors.black)
    c._charSpace = 0.25  # tracking 12 (12/1000 * 21pt)

    palavras = texto.split()
    linhas = []
    linha = ""
    max_chars = 65

    for palavra in palavras:
        teste = linha + palavra + " "
        if len(teste) <= max_chars:
            linha = teste
        else:
            linhas.append(linha.strip())
            linha = palavra + " "

    if linha.strip():
        linhas.append(linha.strip())

    body_center_x = 148.5 * mm
    line_h = round(21 * 1.392)           # 29 pt (entrelinha canva 1,392)
    total_h = len(linhas) * line_h
    body_box_center = 98.85 * mm
    start_y = body_box_center + total_h / 2 - 21 * 0.4

    for i, l in enumerate(linhas):
        c.drawCentredString(body_center_x, start_y - i * line_h, l)

    c._charSpace = 0  # reset tracking

    # ============================================================
    #  ASSINATURA — IMAGEM centralizada sobre a linha do fundo
    #  Linha no fundo: x=22,04 a 28,23cm, y=18,43cm (do topo)
    #  Centro: x=25,135cm, y=18,43cm
    #  Tamanho: 9,3 cm × 2,925 cm
    # ============================================================
    if signature_path and os.path.exists(signature_path):
        img_w = 9.3 * cm
        img_h = 2.925 * cm
        img_x = 25.135 * cm - img_w / 2
        img_y = H - 17.97 * cm - img_h / 2

        c.drawImage(
            signature_path,
            img_x, img_y,
            width=img_w,
            height=img_h,
            preserveAspectRatio=True,
            mask='auto'
        )

    c.save()


# =========================
# ENVIAR EMAIL
# =========================
def _montar_msg(email_destino, nome, assunto, mensagem, arquivo_pdf, email_remetente):

    nome = str(nome).strip() if nome else "Participante"
    mensagem = mensagem.replace("{nome}", nome)

    msg = MIMEMultipart()
    msg["From"]    = email_remetente
    msg["To"]      = email_destino
    msg["Subject"] = assunto

    msg.attach(MIMEText(mensagem, "plain"))

    if arquivo_pdf:
        pdf_path = str(arquivo_pdf)
        if os.path.isfile(pdf_path):
            nome_anexo = os.path.basename(pdf_path)
            if not nome_anexo or nome_anexo.startswith(".") or nome_anexo.lower() in ("noname", ".pdf"):
                nome_anexo = f"{nome}.pdf"
            with open(pdf_path, "rb") as anexo:
                parte = MIMEBase("application", "octet-stream")
                parte.set_payload(anexo.read())
            encoders.encode_base64(parte)
            nome_anexo_encoded = Header(nome_anexo, "utf-8").encode()
            parte.add_header("Content-Disposition", f"attachment; filename={nome_anexo_encoded}")
            msg.attach(parte)

    return msg


def salvar_rascunho(
    email_destino,
    nome,
    assunto,
    mensagem,
    arquivo_pdf,
    email_remetente,
    senha
):

    msg = _montar_msg(email_destino, nome, assunto, mensagem, arquivo_pdf, email_remetente)

    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(email_remetente, senha)

    pasta_draft = None
    for nome_pasta in ["[Gmail]/Drafts", "[Gmail]/Rascunhos", "Drafts", "Rascunhos"]:
        try:
            status, _ = imap.select(nome_pasta)
            if status == "OK":
                pasta_draft = nome_pasta
                break
        except Exception:
            continue

    if not pasta_draft:
        imap.logout()
        raise Exception("Pasta de rascunhos nao encontrada no Gmail.")

    imap.append(pasta_draft, "\\Draft", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
    imap.logout()


def enviar_email_smtp(
    email_destino,
    nome,
    assunto,
    mensagem,
    arquivo_pdf,
    email_remetente,
    senha
):

    msg = _montar_msg(email_destino, nome, assunto, mensagem, arquivo_pdf, email_remetente)

    servidor = smtplib.SMTP("smtp.gmail.com", 587)
    servidor.starttls()
    servidor.login(email_remetente, senha)
    servidor.send_message(msg)
    servidor.quit()