#!/usr/bin/env python3
"""
RataoXML - localiza arquivos de notas fiscais (XML/PDF/etc.) pela chave de 44
digitos no nome do arquivo, move (ou copia) para uma pasta de backup, gera um
ZIP e envia por e-mail.

Interface grafica (sem argumentos):
    python procurar.py

Console:
    python procurar.py lista.json -r                      -> procura e move
    python procurar.py lista.txt -d BKP_ABC --copiar      -> copia em vez de mover
    python procurar.py lista.txt --zip                    -> tambem gera o ZIP
    python procurar.py lista.txt --email alguem@x.com     -> gera o ZIP e envia

Configuracoes de e-mail: tela do programa (botao "Salvar configuracoes"),
arquivo config.json ao lado do programa ou variaveis de ambiente
(RESEND_API_KEY, SMTP_USER, SMTP_PASSWORD).
"""
import argparse
import base64
import datetime
import json
import os
import queue
import re
import shutil
import smtplib
import ssl
import subprocess
import sys
import tempfile
import threading
import urllib.error
import urllib.request
import zipfile
from email.message import EmailMessage

try:
    import tkinter as tk
    from tkinter import filedialog, ttk
    from tkinter import font as tkfont
except ImportError:  # ambiente sem interface grafica: so o modo console funciona
    tk = None

try:  # icone na bandeja do sistema (opcional)
    import pystray
    from PIL import Image
except ImportError:
    pystray = None
    Image = None

VERSAO = "2.1"
PASTA_DESTINO_PADRAO = "BKP_XML"
REGEX_CHAVE = re.compile(r"(?<!\d)\d{44}(?!\d)")
RESEND_FROM_PADRAO = "onboarding@resend.dev"
MENSAGEM_PADRAO = "Segue em anexo os XMLs."
SMTP_PRESETS = {"gmail": ("smtp.gmail.com", 587), "outlook": ("smtp-mail.outlook.com", 587)}
# limites aproximados do tamanho do ZIP (o anexo cresce ~33% ao ser codificado)
LIMITE_MB = {"resend": 30, "gmail": 18, "outlook": 18}
MODELOS = {"55": "NF-e", "65": "NFC-e", "57": "CT-e", "58": "MDF-e", "67": "CT-e OS"}


# =============================================================================
#  NUCLEO (sem interface): configuracao, busca, backup, zip e e-mail
# =============================================================================
def caminho_recurso(nome):
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, nome)


def pasta_app():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


ARQUIVO_CONFIG = os.path.join(pasta_app(), "config.json")


def carregar_config():
    try:
        with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as f:
            dados = json.load(f)
            return dados if isinstance(dados, dict) else {}
    except (OSError, ValueError):
        return {}


def salvar_config(cfg):
    with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def config_padrao():
    cfg = {
        "provider": os.getenv("EMAIL_PROVIDER", "resend"),
        "destino": PASTA_DESTINO_PADRAO,
        "destinatario": os.getenv("EMAIL_DESTINO", ""),
        "assunto": "RataoXML - Backup XML",
        "mensagem": MENSAGEM_PADRAO,
        "resend_api_key": os.getenv("RESEND_API_KEY", ""),
        "resend_from": os.getenv("RESEND_FROM_EMAIL", RESEND_FROM_PADRAO),
        "smtp_user": os.getenv("SMTP_USER", ""),
        "smtp_password": os.getenv("SMTP_PASSWORD", ""),
        "somente_xml": True,
    }
    cfg.update(carregar_config())
    return cfg


def ler_chaves_do_texto(texto):
    """Aceita o JSON da validacao, texto livre ou so as chaves (uma por linha)."""
    if not texto:
        return []
    return list(dict.fromkeys(REGEX_CHAVE.findall(texto)))


def ler_arquivo_texto(caminho):
    with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()


def validar_email(email):
    return bool(email and re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email.strip()))


def parse_destinatarios(texto):
    """Separa e valida e-mails escritos com vírgula, ponto e vírgula, espaço ou quebra de linha."""
    partes = [p for p in re.split(r"[,;\s]+", texto or "") if p]
    validos = list(dict.fromkeys(p for p in partes if validar_email(p)))
    invalidos = [p for p in partes if not validar_email(p)]
    return validos, invalidos


def ler_data(texto):
    """Converte 'dd/mm/aaaa' (ou aaaa-mm-dd) em date; retorna None se inválida."""
    texto = (texto or "").strip()
    for formato in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(texto, formato).date()
        except ValueError:
            pass
    return None


def fmt_data(d):
    return d.strftime("%d/%m/%Y") if d else ""


def info_chave(chave):
    """Modelo e numero da nota extraidos da chave de acesso (44 digitos)."""
    if len(chave) != 44 or not chave.isdigit():
        return "-", "-"
    modelo = MODELOS.get(chave[20:22], chave[20:22])
    try:
        numero = str(int(chave[25:34]))
    except ValueError:
        numero = "-"
    return modelo, numero


def abrir_pasta(caminho):
    try:
        if sys.platform.startswith("win"):
            os.startfile(caminho)  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", caminho])
        else:
            subprocess.Popen(["xdg-open", caminho])
    except Exception:
        pass


# ------------------------------------------------------- procurar e mover
def listar_arquivos(pasta_base, pasta_destino, recursivo):
    destino_abs = os.path.abspath(pasta_destino)
    for raiz, dirs, arquivos in os.walk(pasta_base):
        dirs[:] = [d for d in dirs if os.path.abspath(os.path.join(raiz, d)) != destino_abs]
        for nome in arquivos:
            yield os.path.join(raiz, nome)
        if not recursivo:
            break


def localizar_arquivos(pasta_base, chaves, pasta_destino, recursivo=True):
    """Retorna {chave: [caminhos]} para as chaves que aparecem no nome de algum arquivo."""
    conjunto = set(chaves)
    encontradas = {}
    for caminho in listar_arquivos(pasta_base, pasta_destino, recursivo):
        for chave in REGEX_CHAVE.findall(os.path.basename(caminho)):
            if chave in conjunto:
                encontradas.setdefault(chave, []).append(caminho)
                break
    return encontradas


def mover_arquivos(encontradas, pasta_destino, copiar=False, progresso=None):
    itens = [(c, p) for c, caminhos in encontradas.items() for p in caminhos]
    processados, ignorados, erros = [], [], []
    if itens:
        os.makedirs(pasta_destino, exist_ok=True)
    for i, (_chave, caminho) in enumerate(itens, 1):
        nome = os.path.basename(caminho)
        destino_final = os.path.join(pasta_destino, nome)
        if os.path.exists(destino_final):
            ignorados.append(nome)
        else:
            try:
                (shutil.copy2 if copiar else shutil.move)(caminho, destino_final)
                processados.append(nome)
            except Exception as exc:
                erros.append(f"{nome}: {exc}")
        if progresso:
            progresso(i, len(itens))
    return {"processados": processados, "ignorados": ignorados, "erros": erros}


def localizar_e_mover(pasta_base, texto, pasta_destino, recursivo=True, copiar=False):
    chaves = ler_chaves_do_texto(texto)
    if not chaves:
        return {"sucesso": False, "mensagem": "Nenhuma chave de 44 dígitos foi encontrada na lista."}
    encontradas = localizar_arquivos(pasta_base, chaves, pasta_destino, recursivo)
    res = mover_arquivos(encontradas, pasta_destino, copiar)
    res.update(
        sucesso=True,
        mensagem="Busca concluída.",
        total=len(chaves),
        faltantes=[c for c in chaves if c not in encontradas],
    )
    return res


def resumo_mover(res, copiar=False):
    verbo = "copiados" if copiar else "movidos"
    linhas = [f"Chaves na lista: {res['total']}", f"Arquivos {verbo}: {len(res['processados'])}"]
    linhas += [f"  - {n}" for n in res["processados"]]
    if res["ignorados"]:
        linhas.append(f"Ignorados (já existem na pasta de destino): {len(res['ignorados'])}")
        linhas += [f"  - {n}" for n in res["ignorados"]]
    if res["erros"]:
        linhas.append("Erros:")
        linhas += [f"  - {e}" for e in res["erros"]]
    if res["faltantes"]:
        linhas.append(f"Chaves sem arquivo correspondente: {len(res['faltantes'])}")
        linhas += [f"  - {c}" for c in res["faltantes"]]
    return "\n".join(linhas)


# ------------------------------------------------------ procurar por periodo
RE_DATA_EMISSAO = re.compile(rb"<(?:\w+:)?(?:dhEmi|dEmi|dhEvento)>\s*(\d{4})-(\d{2})-(\d{2})")
RE_DATA_PROTOCOLO = re.compile(rb"<(?:\w+:)?dhRecbto>\s*(\d{4})-(\d{2})-(\d{2})")


def data_do_xml(caminho, criterio="emissao"):
    """Retorna (data, origem). origem = 'emissao' (lida no XML) ou 'arquivo' (data de modificação)."""
    if criterio == "emissao":
        try:
            with open(caminho, "rb") as f:
                cabecalho = f.read(65536)
        except OSError:
            cabecalho = b""
        for padrao in (RE_DATA_EMISSAO, RE_DATA_PROTOCOLO):
            m = padrao.search(cabecalho)
            if m:
                try:
                    return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))), "emissao"
                except ValueError:
                    pass
    try:
        return datetime.date.fromtimestamp(os.path.getmtime(caminho)), "arquivo"
    except OSError:
        return None, ""


def localizar_xml_por_periodo(pasta_base, pasta_destino, inicio, fim, recursivo=True, criterio="emissao",
                              progresso=None):
    """Procura XMLs cuja data (emissão ou arquivo) está entre inicio e fim (inclusive)."""
    caminhos = [p for p in listar_arquivos(pasta_base, pasta_destino, recursivo) if p.lower().endswith(".xml")]
    encontradas, datas, fallback = {}, {}, 0
    for i, caminho in enumerate(caminhos, 1):
        data, origem = data_do_xml(caminho, criterio)
        if criterio == "emissao" and origem == "arquivo":
            fallback += 1
        if data and inicio <= data <= fim:
            nome = os.path.basename(caminho)
            m = REGEX_CHAVE.search(nome)
            chave = m.group(0) if m else nome
            encontradas.setdefault(chave, []).append(caminho)
            datas.setdefault(chave, data)
        if progresso and (i % 40 == 0 or i == len(caminhos)):
            progresso(i, len(caminhos))
    ordem = sorted(encontradas, key=lambda c: (datas[c], c))
    return {"encontradas": {c: encontradas[c] for c in ordem}, "datas": datas,
            "total_xml": len(caminhos), "fallback": fallback}


# --------------------------------------------------------------------- zip
def zipar_pasta(pasta_destino, somente_xml=True, somente_nomes=None, nome_zip=None):
    """Zipa a pasta. somente_nomes: se informado, inclui só esses arquivos (os da busca atual)."""
    if not os.path.isdir(pasta_destino):
        return {"sucesso": False, "mensagem": f"A pasta '{pasta_destino}' não existe. Faça o backup primeiro."}

    permitidos = {n.lower() for n in somente_nomes} if somente_nomes else None
    arquivos = []
    for nome in sorted(os.listdir(pasta_destino)):
        caminho = os.path.join(pasta_destino, nome)
        if not os.path.isfile(caminho) or nome.lower().endswith(".zip"):
            continue
        if somente_xml and not nome.lower().endswith(".xml"):
            continue
        if permitidos is not None and nome.lower() not in permitidos:
            continue
        arquivos.append(caminho)

    if not arquivos:
        tipo = "XML" if somente_xml else "arquivo"
        return {"sucesso": False, "mensagem": f"Nenhum {tipo} encontrado em '{pasta_destino}' para zipar."}

    nome_zip = nome_zip or f"backup_xml_{datetime.datetime.now():%Y%m%d_%H%M%S}.zip"
    zip_path = os.path.join(pasta_destino, nome_zip)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as pacote:
        for arquivo in arquivos:
            pacote.write(arquivo, arcname=os.path.basename(arquivo))

    tamanho = os.path.getsize(zip_path) / (1024 * 1024)
    return {
        "sucesso": True,
        "zip": zip_path,
        "mensagem": f"ZIP criado: {nome_zip} ({len(arquivos)} arquivo(s), {tamanho:.2f} MB).",
    }


# ------------------------------------------------------------------ e-mail
def _msg_erro_resend(codigo, corpo):
    """Traduz a resposta de erro do Resend para uma mensagem curta e clara."""
    try:
        msg = json.loads(corpo).get("message") or corpo
    except (ValueError, AttributeError):
        msg = corpo
    if codigo == 403 and "own email address" in msg:
        m = re.search(r"\(([^)\s]+@[^)\s]+)\)", msg)
        conta = m.group(1) if m else "da sua conta Resend"
        return (f"O Resend só envia para o e-mail da própria conta ({conta}) enquanto não houver um domínio "
                "verificado. Envie para esse e-mail ou use Gmail.")
    if codigo == 401 or "api key" in msg.lower():
        return "API key do Resend inválida ou sem permissão de envio. Confira a key informada."
    return f"Resend recusou o envio (HTTP {codigo}): {msg}"


def _enviar_resend(cfg, destinatarios, zip_path, assunto, corpo):
    api_key = cfg.get("resend_api_key") or os.getenv("RESEND_API_KEY")
    if not api_key:
        return {"sucesso": False, "mensagem": "Informe a API key do Resend (Ctrl+T)."}

    with open(zip_path, "rb") as f:
        conteudo = base64.b64encode(f.read()).decode("ascii")
    payload = {
        "from": cfg.get("resend_from") or RESEND_FROM_PADRAO,
        "to": destinatarios,
        "subject": assunto,
        "text": corpo,
        "attachments": [{"filename": os.path.basename(zip_path), "content": conteudo}],
    }
    req = urllib.request.Request(
        "https://api.resend.com/emails",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "RataoXML/1.0",  # sem isso a Cloudflare do Resend pode bloquear (erro 1010)
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=90):
            pass
        return {"sucesso": True, "mensagem": f"E-mail enviado para {', '.join(destinatarios)} via Resend."}
    except urllib.error.HTTPError as exc:
        detalhe = exc.read().decode("utf-8", errors="ignore")
        return {"sucesso": False, "mensagem": _msg_erro_resend(exc.code, detalhe)}
    except urllib.error.URLError as exc:
        return {"sucesso": False, "mensagem": f"Não foi possível conectar à API do Resend: {exc.reason}"}
    except TimeoutError:
        return {"sucesso": False, "mensagem": "A API do Resend não respondeu dentro do tempo limite."}


def _enviar_smtp(cfg, provider, destinatarios, zip_path, assunto, corpo):
    usuario = cfg.get("smtp_user") or os.getenv("SMTP_USER") or ""
    senha = cfg.get("smtp_password") or os.getenv("SMTP_PASSWORD") or ""
    if not usuario or not senha:
        return {"sucesso": False, "mensagem": "Informe o usuário e a senha de app do e-mail (Ctrl+T)."}

    host, porta = SMTP_PRESETS[provider]
    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = usuario
    msg["To"] = ", ".join(destinatarios)
    msg.set_content(corpo)
    with open(zip_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="zip",
                           filename=os.path.basename(zip_path))
    try:
        with smtplib.SMTP(host, porta, timeout=60) as servidor:
            servidor.starttls(context=ssl.create_default_context())
            servidor.login(usuario, senha.replace(" ", ""))
            servidor.send_message(msg)
        return {"sucesso": True, "mensagem": f"E-mail enviado para {', '.join(destinatarios)} via {provider}."}
    except smtplib.SMTPAuthenticationError:
        return {"sucesso": False, "mensagem": "Login recusado. Use uma senha de app (não a senha normal da conta)."}
    except Exception as exc:
        return {"sucesso": False, "mensagem": f"Erro de rede ao enviar o e-mail: {exc}"}


def enviar_email(cfg, destinatario, zip_path):
    """destinatario: um ou vários e-mails separados por vírgula/espaço."""
    if not zip_path or not os.path.exists(zip_path):
        return {"sucesso": False, "mensagem": "Nenhum ZIP disponível para envio."}
    validos, invalidos = parse_destinatarios(destinatario)
    if not validos or invalidos:
        ruim = ", ".join(invalidos) if invalidos else "nenhum"
        return {"sucesso": False, "mensagem": f"E-mail inválido ou ausente: {ruim}."}

    provider = (cfg.get("provider") or "resend").lower()
    if provider not in LIMITE_MB:
        return {"sucesso": False, "mensagem": f"Provedor desconhecido: {provider}."}

    mb = os.path.getsize(zip_path) / (1024 * 1024)
    if mb > LIMITE_MB[provider]:
        return {"sucesso": False,
                "mensagem": f"O ZIP tem {mb:.1f} MB e passa do limite de cerca de {LIMITE_MB[provider]} MB do {provider}."}

    assunto = cfg.get("assunto") or "RataoXML - Backup XML"
    corpo = f"{cfg.get('mensagem') or MENSAGEM_PADRAO}\n\nData do envio: {datetime.datetime.now():%d/%m/%Y %H:%M:%S}"
    try:
        if provider == "resend":
            return _enviar_resend(cfg, validos, zip_path, assunto, corpo)
        return _enviar_smtp(cfg, provider, validos, zip_path, assunto, corpo)
    except Exception as exc:
        return {"sucesso": False, "mensagem": f"Falha no envio do e-mail: {exc}"}


def zipar_e_enviar(cfg, pasta_destino, destinatario, somente_xml=True, somente_nomes=None, nome_zip=None):
    res_zip = zipar_pasta(pasta_destino, somente_xml, somente_nomes, nome_zip)
    if not res_zip["sucesso"]:
        return res_zip
    res_mail = enviar_email(cfg, destinatario, res_zip["zip"])
    return {
        "sucesso": res_mail["sucesso"],
        "zip": res_zip["zip"],
        "mensagem": res_zip["mensagem"] + "\n" + res_mail["mensagem"],
    }


def testar_envio(cfg, destinatario):
    """Envia um e-mail real e mínimo para confirmar provedor, credenciais e destinatário."""
    with tempfile.TemporaryDirectory(prefix="rataoxml_teste_") as pasta:
        zip_teste = os.path.join(pasta, "teste_rataoxml.zip")
        with zipfile.ZipFile(zip_teste, "w", zipfile.ZIP_DEFLATED) as pacote:
            pacote.writestr("teste.txt", "Teste de envio do RataoXML.")
        cfg_teste = dict(cfg, assunto="RataoXML - Teste de envio",
                         mensagem="Este é um e-mail de teste do RataoXML. Se você recebeu esta mensagem, "
                                  "as configurações de envio estão corretas.")
        return enviar_email(cfg_teste, destinatario, zip_teste)


# =============================================================================
#  INTERFACE GRAFICA
# =============================================================================
_ESC = [1.0]


def px(n):
    """Escala um tamanho em pixels de acordo com o DPI da tela."""
    return int(round(n * _ESC[0]))


class T:
    """Paleta de cores (tema escuro)."""
    BG = "#0B1020"
    SIDEBAR = "#0E1430"
    CARD = "#131B36"
    SURFACE = "#1A2345"
    BORDER = "#27325C"
    TEXT = "#EAEEFB"
    MUTED = "#8E9AC0"
    DIM = "#5E6A94"
    ACCENT = "#6366F1"
    ACCENT_H = "#7B7EF6"
    ACCENT_D = "#4B4FD6"
    OK, OK_BG = "#34D399", "#10302A"
    WARN, WARN_BG = "#FBBF24", "#35290E"
    ERR, ERR_BG = "#F87171", "#3B1A1D"
    INFO, INFO_BG = "#A5B4FC", "#1F2452"


class F:
    """Fontes (criadas depois que a janela existe)."""


def criar_fontes(root):
    familias = set(tkfont.families(root))

    def escolher(*nomes):
        for n in nomes:
            if n in familias:
                return n
        return "TkDefaultFont"

    base = escolher("Segoe UI Variable Text", "Segoe UI", "Inter", "SF Pro Text", "Helvetica Neue",
                    "Noto Sans", "DejaVu Sans")
    mono = escolher("Cascadia Mono", "Consolas", "SF Mono", "Menlo", "DejaVu Sans Mono")
    F.titulo = tkfont.Font(root, family=base, size=20, weight="bold")
    F.marca = tkfont.Font(root, family=base, size=15, weight="bold")
    F.h2 = tkfont.Font(root, family=base, size=11, weight="bold")
    F.corpo = tkfont.Font(root, family=base, size=10)
    F.peq = tkfont.Font(root, family=base, size=9)
    F.peq_b = tkfont.Font(root, family=base, size=9, weight="bold")
    F.btn = tkfont.Font(root, family=base, size=10, weight="bold")
    F.mono = tkfont.Font(root, family=mono, size=9)


def criar_estilos(root):
    st = ttk.Style(root)
    try:
        st.theme_use("clam")
    except tk.TclError:
        pass
    st.configure("Treeview", background=T.CARD, fieldbackground=T.CARD, foreground=T.TEXT,
                 borderwidth=0, relief="flat", rowheight=px(30), font=F.corpo)
    st.configure("Treeview.Heading", background=T.SURFACE, foreground=T.MUTED, relief="flat",
                 borderwidth=0, font=F.peq_b, padding=(px(8), px(7)))
    st.map("Treeview", background=[("selected", "#2A3470")], foreground=[("selected", T.TEXT)])
    st.map("Treeview.Heading", background=[("active", T.SURFACE)])
    try:
        st.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        st.layout("Vertical.TScrollbar", [("Vertical.Scrollbar.trough", {
            "sticky": "ns",
            "children": [("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
    except tk.TclError:
        pass
    st.configure("Vertical.TScrollbar", background=T.BORDER, troughcolor=T.CARD, bordercolor=T.CARD,
                 lightcolor=T.BORDER, darkcolor=T.BORDER, relief="flat", width=px(9))
    st.map("Vertical.TScrollbar", background=[("active", T.DIM)])
    st.configure("Accent.Horizontal.TProgressbar", troughcolor=T.SURFACE, background=T.ACCENT,
                 bordercolor=T.SURFACE, lightcolor=T.ACCENT, darkcolor=T.ACCENT, thickness=px(8))


def ativar_dpi():
    if os.name != "nt":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def titulo_escuro(root):
    """Barra de titulo escura no Windows 10/11."""
    if os.name != "nt":
        return
    try:
        import ctypes
        root.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        valor = ctypes.c_int(1)
        for atributo in (20, 19):
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, atributo, ctypes.byref(valor), ctypes.sizeof(valor))
    except Exception:
        pass


def _rr(c, x1, y1, x2, y2, r, **kw):
    """Retangulo com cantos arredondados em um Canvas."""
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    pts = [x1 + r, y1, x1 + r, y1, x2 - r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y1 + r,
           x2, y2 - r, x2, y2 - r, x2, y2, x2 - r, y2, x2 - r, y2, x1 + r, y2, x1 + r, y2,
           x1, y2, x1, y2 - r, x1, y2 - r, x1, y1 + r, x1, y1 + r, x1, y1]
    return c.create_polygon(pts, smooth=True, **kw)


# ----------------------------------------------------------------- widgets
class Logo(tk.Canvas):
    def __init__(self, pai, tam):
        super().__init__(pai, width=tam, height=tam, bg=pai.cget("bg"), highlightthickness=0)
        s = tam
        _rr(self, 0, 0, s, s, s * 0.28, fill=T.ACCENT, outline="")
        x1, y1, x2, y2, d = s * 0.27, s * 0.19, s * 0.65, s * 0.79, s * 0.14
        self.create_polygon(x1, y1, x2 - d, y1, x2, y1 + d, x2, y2, x1, y2, fill="#FFFFFF", outline="")
        self.create_polygon(x2 - d, y1, x2, y1 + d, x2 - d, y1 + d, fill="#C7CBFB", outline="")
        for fy, fw in ((0.38, 0.7), (0.50, 0.7), (0.62, 0.45)):
            yy = s * fy
            ini = x1 + s * 0.06
            self.create_line(ini, yy, ini + (x2 - x1 - s * 0.12) * fw, yy, fill=T.ACCENT,
                             width=max(2, int(s * 0.045)), capstyle="round")
        cx, cy, r = s * 0.69, s * 0.71, s * 0.17
        self.create_oval(cx - r, cy - r, cx + r, cy + r, fill=T.OK, outline=T.ACCENT, width=max(2, int(s * 0.06)))
        self.create_line(cx - r * 0.45, cy, cx - r * 0.1, cy + r * 0.38, cx + r * 0.5, cy - r * 0.35,
                         fill="#FFFFFF", width=max(2, int(s * 0.055)), capstyle="round", joinstyle="round")


class Botao(tk.Canvas):
    """Botao arredondado com hover. estilos: primario, secundario, fantasma."""

    def __init__(self, pai, texto, comando, estilo="primario", largura=None, altura=42):
        super().__init__(pai, bg=pai.cget("bg"), highlightthickness=0, cursor="hand2",
                         width=largura or (F.btn.measure(texto) + px(44)), height=px(altura))
        self.texto, self.comando = texto, comando
        self.cores = {
            "primario": dict(bg=T.ACCENT, hv=T.ACCENT_H, fg="#FFFFFF", borda=T.ACCENT),
            "secundario": dict(bg=T.SURFACE, hv="#232E58", fg=T.TEXT, borda=T.BORDER),
            "fantasma": dict(bg=None, hv=T.SURFACE, fg=T.MUTED, borda=None),
        }[estilo]
        self.ativo, self.hover = True, False
        self.bind("<Configure>", lambda _e: self._desenhar())
        self.bind("<Enter>", lambda _e: self._set_hover(True))
        self.bind("<Leave>", lambda _e: self._set_hover(False))
        self.bind("<ButtonRelease-1>", self._soltar)

    def _set_hover(self, valor):
        self.hover = valor
        self._desenhar()

    def _soltar(self, evento):
        if self.ativo and 0 <= evento.x <= self.winfo_width() and 0 <= evento.y <= self.winfo_height():
            self.comando()

    def habilitar(self, flag):
        self.ativo = bool(flag)
        self.configure(cursor="hand2" if self.ativo else "arrow")
        self._desenhar()

    def definir_texto(self, texto):
        self.texto = texto
        self._desenhar()

    def _desenhar(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 5:
            w, h = int(self.cget("width")), int(self.cget("height"))
        c = self.cores
        if not self.ativo:
            fundo, borda, fg = "#151D3B", T.BORDER, T.DIM
        else:
            fundo = c["hv"] if self.hover else c["bg"]
            borda, fg = c["borda"], c["fg"]
            if c["bg"] is None and not self.hover:
                fundo = ""
        _rr(self, 1, 1, w - 1, h - 1, px(11), fill=fundo or "", outline=borda or "")
        self.create_text(w / 2, h / 2, text=self.texto, fill=fg, font=F.btn)


class Chip(tk.Label):
    CORES = {"ok": (T.OK_BG, T.OK), "aviso": (T.WARN_BG, T.WARN), "erro": (T.ERR_BG, T.ERR),
             "info": (T.INFO_BG, T.INFO)}

    def __init__(self, pai, texto="", tipo="info"):
        super().__init__(pai, font=F.peq_b, padx=px(12), pady=px(5))
        self.definir(texto, tipo)

    def definir(self, texto, tipo="info"):
        bg, fg = self.CORES[tipo]
        self.configure(text=texto, bg=bg, fg=fg)


class Banner(tk.Frame):
    """Aviso colorido que aparece e some (sucesso, erro, informacao)."""

    def __init__(self, pai):
        super().__init__(pai, bg=T.BG)
        self.barra = tk.Frame(self, width=px(4))
        self.barra.pack(side="left", fill="y")
        self.rotulo = tk.Label(self, font=F.corpo, justify="left", anchor="w", padx=px(14), pady=px(10))
        self.rotulo.pack(side="left", fill="x", expand=True)
        self.bind("<Configure>", lambda e: self.rotulo.configure(wraplength=max(100, e.width - px(40))))

    def mostrar(self, texto, tipo="info"):
        bg, fg = Chip.CORES[{"ok": "ok", "erro": "erro", "aviso": "aviso"}.get(tipo, "info")]
        self.configure(bg=bg)
        self.barra.configure(bg=fg)
        self.rotulo.configure(text=texto, bg=bg, fg=fg)
        self.pack(fill="x", pady=(px(12), 0))

    def ocultar(self):
        self.pack_forget()


class Cartao(tk.Frame):
    def __init__(self, pai, pad=18):
        super().__init__(pai, bg=T.CARD, highlightthickness=1, highlightbackground=T.BORDER)
        self.interno = tk.Frame(self, bg=T.CARD)
        self.interno.pack(fill="both", expand=True, padx=px(pad), pady=px(pad))


class Campo(tk.Frame):
    """Campo de texto com rotulo e borda que acende no foco (senha tem botão Mostrar/Ocultar)."""

    def __init__(self, pai, variavel, rotulo=None, senha=False, somente_leitura=False):
        super().__init__(pai, bg=pai.cget("bg"))
        if rotulo:
            tk.Label(self, text=rotulo, font=F.peq_b, fg=T.MUTED, bg=self.cget("bg")).pack(anchor="w")
        self.caixa = tk.Frame(self, bg=T.SURFACE, highlightthickness=1, highlightbackground=T.BORDER)
        self.caixa.pack(fill="x", pady=(px(5), 0))
        self.entrada = tk.Entry(
            self.caixa, textvariable=variavel, bg=T.SURFACE, fg=T.TEXT, insertbackground=T.TEXT,
            relief="flat", bd=0, font=F.corpo, show="•" if senha else "", highlightthickness=0,
            readonlybackground=T.SURFACE, selectbackground=T.ACCENT_D, selectforeground="#FFFFFF",
            state="readonly" if somente_leitura else "normal")
        if senha:
            self.olho = tk.Label(self.caixa, text="Mostrar", font=F.peq_b, fg=T.MUTED, bg=T.SURFACE,
                                 cursor="hand2", padx=px(12))
            self.olho.pack(side="right")
            self.olho.bind("<Button-1>", self._alternar)
        self.entrada.pack(side="left", fill="x", expand=True, padx=(px(12), px(4) if senha else px(12)),
                          pady=px(9))
        self.entrada.bind("<FocusIn>", lambda _e: self.caixa.configure(highlightbackground=T.ACCENT))
        self.entrada.bind("<FocusOut>", lambda _e: self.caixa.configure(highlightbackground=T.BORDER))

    def _alternar(self, _evento=None):
        oculto = self.entrada.cget("show") != ""
        self.entrada.configure(show="" if oculto else "•")
        self.olho.configure(text="Ocultar" if oculto else "Mostrar")


class CaixaTexto(tk.Frame):
    def __init__(self, pai, altura=5, mono=False, somente_leitura=False):
        super().__init__(pai, bg=T.SURFACE, highlightthickness=1, highlightbackground=T.BORDER)
        self.somente_leitura = somente_leitura
        self.texto = tk.Text(
            self, height=altura, bg=T.SURFACE, fg=T.TEXT, insertbackground=T.TEXT, relief="flat", bd=0,
            wrap="word", font=F.mono if mono else F.corpo, padx=px(12), pady=px(9), highlightthickness=0,
            selectbackground=T.ACCENT_D, selectforeground="#FFFFFF")
        barra = ttk.Scrollbar(self, orient="vertical", command=self.texto.yview)
        self.texto.configure(yscrollcommand=barra.set)
        barra.pack(side="right", fill="y")
        self.texto.pack(side="left", fill="both", expand=True)
        self.texto.bind("<FocusIn>", lambda _e: self.configure(highlightbackground=T.ACCENT))
        self.texto.bind("<FocusOut>", lambda _e: self.configure(highlightbackground=T.BORDER))
        if somente_leitura:
            self.texto.configure(state="disabled")

    def _editar(self, funcao):
        self.texto.configure(state="normal")
        funcao()
        if self.somente_leitura:
            self.texto.configure(state="disabled")

    def obter(self):
        return self.texto.get("1.0", "end")

    def definir(self, conteudo):
        self._editar(lambda: (self.texto.delete("1.0", "end"), self.texto.insert("end", conteudo)))

    def anexar(self, conteudo):
        self._editar(lambda: (self.texto.insert("end", conteudo), self.texto.see("end")))

    def limpar(self):
        self._editar(lambda: self.texto.delete("1.0", "end"))


class Segmentado(tk.Frame):
    """Seletor em segmentos (tipo abas) ligado a uma StringVar."""

    def __init__(self, pai, opcoes, variavel, comando=None):
        super().__init__(pai, bg=T.SURFACE, highlightthickness=1, highlightbackground=T.BORDER)
        self.var, self.comando, self.itens = variavel, comando, {}
        for valor, rotulo in opcoes:
            lb = tk.Label(self, text=rotulo, font=F.peq_b, padx=px(16), pady=px(7), cursor="hand2")
            lb.pack(side="left", padx=px(2), pady=px(2))
            lb.bind("<Button-1>", lambda _e, v=valor: self._escolher(v))
            self.itens[valor] = lb
        variavel.trace_add("write", lambda *_a: self._pintar())
        self._pintar()

    def _escolher(self, valor):
        self.var.set(valor)
        if self.comando:
            self.comando()

    def _pintar(self):
        for valor, lb in self.itens.items():
            ativo = valor == self.var.get()
            lb.configure(bg=T.ACCENT if ativo else T.SURFACE, fg="#FFFFFF" if ativo else T.MUTED)


class Alternador(tk.Frame):
    """Interruptor (liga/desliga) ligado a uma BooleanVar."""

    def __init__(self, pai, texto, variavel):
        super().__init__(pai, bg=pai.cget("bg"), cursor="hand2")
        self.var = variavel
        self.c = tk.Canvas(self, width=px(40), height=px(22), bg=self.cget("bg"), highlightthickness=0,
                           cursor="hand2")
        self.c.pack(side="left")
        self.lb = tk.Label(self, text=texto, font=F.corpo, fg=T.TEXT, bg=self.cget("bg"), cursor="hand2")
        self.lb.pack(side="left", padx=(px(10), 0))
        for w in (self.c, self.lb, self):
            w.bind("<Button-1>", lambda _e: self.var.set(not self.var.get()))
        variavel.trace_add("write", lambda *_a: self._desenhar())
        self._desenhar()

    def _desenhar(self):
        self.c.delete("all")
        w, h, on = px(40), px(22), self.var.get()
        _rr(self.c, 1, 1, w - 1, h - 1, h / 2, fill=T.ACCENT if on else T.SURFACE,
            outline=T.ACCENT if on else T.BORDER)
        r = h / 2 - px(4)
        cx = w - h / 2 if on else h / 2
        self.c.create_oval(cx - r, h / 2 - r, cx + r, h / 2 + r, fill="#FFFFFF" if on else T.MUTED, outline="")


class ZonaArquivo(tk.Canvas):
    """Area grande e clicavel para escolher o arquivo da lista."""

    def __init__(self, pai, comando):
        super().__init__(pai, height=px(150), bg=pai.cget("bg"), highlightthickness=0, cursor="hand2")
        self.comando, self.hover, self.ok = comando, False, False
        self.titulo = "Clique para selecionar a lista de notas"
        self.sub = "JSON ou TXT  ·  as chaves de 44 dígitos são reconhecidas automaticamente"
        self.bind("<Configure>", lambda _e: self._desenhar())
        self.bind("<Enter>", lambda _e: self._set_hover(True))
        self.bind("<Leave>", lambda _e: self._set_hover(False))
        self.bind("<Button-1>", lambda _e: self.comando())

    def _set_hover(self, valor):
        self.hover = valor
        self._desenhar()

    def definir(self, titulo, sub, ok):
        self.titulo, self.sub, self.ok = titulo, sub, ok
        self._desenhar()

    def _desenhar(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 5:
            return
        cor = T.ACCENT if self.hover else (T.OK if self.ok else T.BORDER)
        _rr(self, 2, 2, w - 2, h - 2, px(16), fill="#171F3F" if self.hover else T.CARD, outline=cor,
            width=2, dash=() if self.ok else (6, 5))
        cx, cy, r = w / 2, h / 2 - px(26), px(24)
        self.create_oval(cx - r, cy - r, cx + r, cy + r, fill=T.OK_BG if self.ok else T.INFO_BG, outline="")
        k = px(1)
        if self.ok:
            self.create_line(cx - 9 * k, cy, cx - 3 * k, cy + 7 * k, cx + 10 * k, cy - 7 * k, fill=T.OK,
                             width=px(3), capstyle="round", joinstyle="round")
        else:
            self.create_line(cx, cy + 8 * k, cx, cy - 9 * k, fill=T.INFO, width=px(3), capstyle="round")
            self.create_line(cx - 7 * k, cy - 2 * k, cx, cy - 9 * k, cx + 7 * k, cy - 2 * k, fill=T.INFO,
                             width=px(3), capstyle="round", joinstyle="round")
            self.create_line(cx - 12 * k, cy + 6 * k, cx - 12 * k, cy + 12 * k, cx + 12 * k, cy + 12 * k,
                             cx + 12 * k, cy + 6 * k, fill=T.INFO, width=px(2), capstyle="round",
                             joinstyle="round")
        self.create_text(cx, h / 2 + px(20), text=self.titulo, fill=T.TEXT, font=F.h2)
        self.create_text(cx, h / 2 + px(42), text=self.sub, fill=T.MUTED, font=F.peq)


MESES = ("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro",
         "outubro", "novembro", "dezembro")
DIAS_SEMANA = ("D", "S", "T", "Q", "Q", "S", "S")


class Calendario(tk.Canvas):
    """Calendário com seleção de período: clique na data inicial e depois na final."""

    def __init__(self, pai, comando=None):
        self.larg, self.alt = px(312), px(306)
        super().__init__(pai, width=self.larg, height=self.alt, bg=T.CARD, highlightthickness=0,
                         cursor="hand2")
        hoje = datetime.date.today()
        self.ano, self.mes = hoje.year, hoje.month
        self.ini = self.fim = None
        self.hover = None
        self.comando = comando
        self.topo, self.sem = px(46), px(26)
        self.cw = self.larg / 7
        self.ch = (self.alt - self.topo - self.sem - px(6)) / 6
        self.bind("<Button-1>", self._clique)
        self.bind("<Motion>", lambda e: self._set_hover(self._alvo(e.x, e.y)))
        self.bind("<Leave>", lambda _e: self._set_hover(None))
        self.bind("<MouseWheel>", lambda e: self._mudar(-1 if e.delta > 0 else 1))
        self.bind("<Button-4>", lambda _e: self._mudar(-1))
        self.bind("<Button-5>", lambda _e: self._mudar(1))
        self._desenhar()

    # -- dados
    def _datas(self):
        primeiro = datetime.date(self.ano, self.mes, 1)
        inicio = primeiro - datetime.timedelta(days=(primeiro.weekday() + 1) % 7)  # semana começa no domingo
        return [inicio + datetime.timedelta(days=i) for i in range(42)]

    def _alvo(self, x, y):
        if y < self.topo:
            if x < px(32):
                return "ano-"
            if x < px(64):
                return "mes-"
            if x > self.larg - px(32):
                return "ano+"
            if x > self.larg - px(64):
                return "mes+"
            return None
        y0 = self.topo + self.sem
        if y < y0 or y > y0 + 6 * self.ch:
            return None
        coluna = min(6, max(0, int(x // self.cw)))
        linha = min(5, int((y - y0) // self.ch))
        return self._datas()[linha * 7 + coluna]

    def _mudar(self, meses):
        total = self.ano * 12 + (self.mes - 1) + meses
        self.ano, self.mes = total // 12, total % 12 + 1
        self._desenhar()

    def _set_hover(self, alvo):
        if alvo != self.hover:
            self.hover = alvo
            self._desenhar()

    def _clique(self, evento):
        alvo = self._alvo(evento.x, evento.y)
        if alvo == "ano-":
            self._mudar(-12)
        elif alvo == "ano+":
            self._mudar(12)
        elif alvo == "mes-":
            self._mudar(-1)
        elif alvo == "mes+":
            self._mudar(1)
        elif isinstance(alvo, datetime.date):
            self._selecionar(alvo)

    def _selecionar(self, d):
        if self.ini is None or self.fim is not None:
            self.ini, self.fim = d, None
        elif d < self.ini:
            self.ini, self.fim = d, self.ini
        else:
            self.fim = d
        if d.month != self.mes or d.year != self.ano:
            self.ano, self.mes = d.year, d.month
        self._desenhar()
        if self.comando:
            self.comando(self.ini, self.fim)

    def definir(self, ini, fim):
        """Define o período por código (atalhos ou datas digitadas), sem disparar o callback."""
        self.ini, self.fim = ini, fim
        ref = ini or fim
        if ref:
            self.ano, self.mes = ref.year, ref.month
        self._desenhar()

    # -- desenho
    def _desenhar(self):
        self.delete("all")
        w, hoje = self.larg, datetime.date.today()
        for chave, x, texto in (("ano-", px(18), "«"), ("mes-", px(48), "‹"),
                                ("mes+", w - px(48), "›"), ("ano+", w - px(18), "»")):
            ativo = self.hover == chave
            if ativo:
                r = px(13)
                self.create_oval(x - r, self.topo / 2 - r, x + r, self.topo / 2 + r, fill=T.SURFACE, outline="")
            self.create_text(x, self.topo / 2, text=texto, fill=T.TEXT if ativo else T.MUTED, font=F.h2)
        self.create_text(w / 2, self.topo / 2, text=f"{MESES[self.mes - 1].capitalize()} de {self.ano}",
                         fill=T.TEXT, font=F.h2)
        for i, letra in enumerate(DIAS_SEMANA):
            self.create_text(i * self.cw + self.cw / 2, self.topo + self.sem / 2, text=letra, fill=T.DIM,
                             font=F.peq_b)

        y0 = self.topo + self.sem
        ini, fim = self.ini, self.fim
        previa = None
        if ini and not fim and isinstance(self.hover, datetime.date) and self.hover > ini:
            previa = self.hover  # mostra o intervalo enquanto o mouse passa
        fim_visivel = fim or previa
        raio = min(self.cw, self.ch) / 2 - px(3)
        for idx, d in enumerate(self._datas()):
            linha, coluna = divmod(idx, 7)
            x0, y = coluna * self.cw, y0 + linha * self.ch
            cx, cy = x0 + self.cw / 2, y + self.ch / 2
            extremo = d == ini or d == fim
            no_intervalo = bool(ini and fim_visivel and ini <= d <= fim_visivel)
            if no_intervalo and ini != fim_visivel:
                esq = cx if d == ini else x0
                dir_ = cx if d == fim_visivel else x0 + self.cw
                self.create_rectangle(esq, cy - raio, dir_, cy + raio, fill="#252C66", outline="")
            if extremo:
                self.create_oval(cx - raio, cy - raio, cx + raio, cy + raio, fill=T.ACCENT, outline="")
            elif d == previa:
                self.create_oval(cx - raio, cy - raio, cx + raio, cy + raio, outline=T.ACCENT, width=2)
            elif self.hover == d:
                self.create_oval(cx - raio, cy - raio, cx + raio, cy + raio, fill=T.SURFACE, outline="")
            if d == hoje and not extremo:
                self.create_oval(cx - raio, cy - raio, cx + raio, cy + raio, outline=T.OK, width=1)
            if extremo:
                cor, fonte = "#FFFFFF", F.btn
            elif d.month == self.mes:
                cor, fonte = T.TEXT, F.corpo
            else:
                cor, fonte = T.DIM, F.corpo
            self.create_text(cx, cy, text=str(d.day), fill=cor, font=fonte)


class Lateral(tk.Frame):
    PASSOS = [("Lista", "Escolha o arquivo"), ("Chaves", "Revise o que foi achado"),
              ("Backup", "Mover para a pasta"), ("Envio", "ZIP e e-mail")]

    def __init__(self, pai, comando_config=None):
        super().__init__(pai, bg=T.SIDEBAR, width=px(260))
        self.pack_propagate(False)
        topo = tk.Frame(self, bg=T.SIDEBAR)
        topo.pack(fill="x", padx=px(24), pady=(px(30), px(30)))
        Logo(topo, px(48)).pack(side="left")
        textos = tk.Frame(topo, bg=T.SIDEBAR)
        textos.pack(side="left", padx=(px(12), 0))
        tk.Label(textos, text="RataoXML", font=F.marca, fg=T.TEXT, bg=T.SIDEBAR).pack(anchor="w")
        tk.Label(textos, text="Backup de notas fiscais", font=F.peq, fg=T.MUTED, bg=T.SIDEBAR).pack(anchor="w")

        self.canvas = tk.Canvas(self, bg=T.SIDEBAR, highlightthickness=0, height=px(290))
        self.canvas.pack(fill="x", padx=px(24))
        tk.Label(self, text=f"v{VERSAO}  ·  2026", font=F.peq, fg=T.DIM, bg=T.SIDEBAR).pack(
            side="bottom", pady=px(20))
        if comando_config:
            Botao(self, "Configurações  (Ctrl+T)", comando_config, "secundario", largura=px(212),
                  altura=38).pack(side="bottom", pady=(0, px(4)))

    def atualizar(self, atual, concluidos):
        c = self.canvas
        c.delete("all")
        passo, r = px(68), px(15)
        x = r + px(2)
        for i, (titulo, sub) in enumerate(self.PASSOS):
            y = px(22) + i * passo
            feito = i in concluidos and i != atual
            if i < len(self.PASSOS) - 1:
                c.create_line(x, y + r + px(3), x, y + passo - r - px(3),
                              fill=T.OK if i in concluidos else T.BORDER, width=2)
            if i == atual:
                c.create_oval(x - r, y - r, x + r, y + r, fill=T.ACCENT, outline="")
                c.create_text(x, y, text=str(i + 1), fill="#FFFFFF", font=F.btn)
            elif feito:
                c.create_oval(x - r, y - r, x + r, y + r, fill=T.OK_BG, outline=T.OK, width=2)
                k = px(1)
                c.create_line(x - 5 * k, y, x - 1 * k, y + 4 * k, x + 6 * k, y - 4 * k, fill=T.OK,
                              width=2, capstyle="round", joinstyle="round")
            else:
                c.create_oval(x - r, y - r, x + r, y + r, fill=T.SIDEBAR, outline=T.BORDER, width=2)
                c.create_text(x, y, text=str(i + 1), fill=T.DIM, font=F.btn)
            ativo = i == atual or feito
            c.create_text(x + r + px(14), y - px(8), text=titulo, anchor="w", font=F.h2,
                          fill=T.TEXT if ativo else T.MUTED)
            c.create_text(x + r + px(14), y + px(10), text=sub, anchor="w", font=F.peq, fill=T.DIM)


# ------------------------------------------------------------------ paginas
class Pagina(tk.Frame):
    TITULO = ""
    SUBTITULO = ""

    def __init__(self, app):
        super().__init__(app.conteudo, bg=T.BG)
        self.app = app
        cab = tk.Frame(self, bg=T.BG)
        cab.pack(fill="x", padx=px(36), pady=(px(30), 0))
        self.lbl_titulo = tk.Label(cab, text=self.TITULO, font=F.titulo, fg=T.TEXT, bg=T.BG)
        self.lbl_titulo.pack(anchor="w")
        self.lbl_sub = tk.Label(cab, text=self.SUBTITULO, font=F.corpo, fg=T.MUTED, bg=T.BG,
                                justify="left", wraplength=px(700))
        self.lbl_sub.pack(anchor="w", pady=(px(4), 0))
        self.avisos = tk.Frame(self, bg=T.BG)
        self.avisos.pack(fill="x", padx=px(36))
        self.banner = Banner(self.avisos)
        self.rodape = tk.Frame(self, bg=T.BG)
        self.rodape.pack(side="bottom", fill="x", padx=px(36), pady=(0, px(24)))
        self.corpo = tk.Frame(self, bg=T.BG)
        self.corpo.pack(fill="both", expand=True, padx=px(36), pady=px(16))
        self.montar()

    def montar(self):
        pass

    def ao_entrar(self):
        pass

    def ao_erro(self):
        pass


class PaginaLista(Pagina):
    TEXTOS = {
        "lista": ("Selecione a lista de notas",
                  "Aceita o JSON da validação ou um TXT com as chaves. Qualquer sequência de 44 dígitos é reconhecida."),
        "periodo": ("Buscar XMLs por período",
                    "Clique na data inicial e depois na final no calendário (ou digite as datas). "
                    "O programa procura na pasta os XMLs desse período."),
    }
    TITULO, SUBTITULO = TEXTOS["lista"]

    def montar(self):
        app = self.app
        Segmentado(self.corpo, [("lista", "Por lista de chaves"), ("periodo", "Por período")],
                   app.var_busca, comando=self.trocar_modo).pack(anchor="w", pady=(0, px(14)))
        self.quadro_lista = tk.Frame(self.corpo, bg=T.BG)
        self.quadro_periodo = tk.Frame(self.corpo, bg=T.BG)
        self._montar_lista(self.quadro_lista)
        self._montar_periodo(self.quadro_periodo)

        self.chip = Chip(self.rodape)
        self.chip.pack(side="left")
        self.btn = Botao(self.rodape, "Buscar chaves  →", self.avancar, "primario", largura=px(250))
        self.btn.pack(side="right")
        self.trocar_modo()

    # -- blocos
    def _bloco_pasta(self, pai):
        app = self.app
        bloco = tk.Frame(pai, bg=T.BG)
        linha = tk.Frame(bloco, bg=T.BG)
        linha.pack(fill="x")
        linha.columnconfigure(0, weight=1)
        Campo(linha, app.var_pasta, "Pasta de busca", somente_leitura=True).grid(row=0, column=0, sticky="ew")
        Botao(linha, "Alterar", self.escolher_pasta, "secundario", largura=px(96)).grid(
            row=0, column=1, sticky="s", padx=(px(10), 0))
        Alternador(bloco, "Incluir subpastas na busca", app.var_recursivo).pack(anchor="w", pady=(px(12), 0))
        return bloco

    def _montar_lista(self, q):
        self.zona = ZonaArquivo(q, self.escolher_arquivo)
        self.zona.pack(fill="x")
        tk.Label(q, text="Ou cole o conteúdo da lista", font=F.peq_b, fg=T.MUTED, bg=T.BG).pack(
            anchor="w", pady=(px(16), px(6)))
        self.caixa = CaixaTexto(q, altura=4)
        self.caixa.pack(fill="x")
        self.caixa.texto.bind("<KeyRelease>", lambda _e: self.atualizar_contagem())
        self.caixa.texto.bind("<<Paste>>", lambda _e: self.after(60, self.atualizar_contagem))
        self._bloco_pasta(q).pack(fill="x", pady=(px(16), 0))

    def _montar_periodo(self, q):
        app = self.app
        cartao = Cartao(q, pad=10)
        cartao.grid(row=0, column=0, sticky="n")
        self.calendario = Calendario(cartao.interno, comando=self.ao_escolher_periodo)
        self.calendario.pack()

        dire = tk.Frame(q, bg=T.BG)
        dire.grid(row=0, column=1, sticky="new", padx=(px(20), 0))
        q.columnconfigure(1, weight=1)
        dire.columnconfigure(0, weight=1, uniform="d")
        dire.columnconfigure(1, weight=1, uniform="d")
        self.campo_ini = Campo(dire, app.var_ini, "Data inicial")
        self.campo_ini.grid(row=0, column=0, sticky="ew", padx=(0, px(6)))
        self.campo_fim = Campo(dire, app.var_fim, "Data final")
        self.campo_fim.grid(row=0, column=1, sticky="ew", padx=(px(6), 0))
        for campo in (self.campo_ini, self.campo_fim):
            campo.entrada.bind("<FocusOut>", lambda _e: self.digitou())
            campo.entrada.bind("<Return>", lambda _e: self.digitou())

        tk.Label(dire, text="Atalhos", font=F.peq_b, fg=T.MUTED, bg=T.BG).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(px(12), px(5)))
        atalhos = tk.Frame(dire, bg=T.BG)
        atalhos.grid(row=2, column=0, columnspan=2, sticky="ew")
        atalhos.columnconfigure(0, weight=1, uniform="a")
        atalhos.columnconfigure(1, weight=1, uniform="a")
        for i, (rotulo, nome) in enumerate((("Hoje", "hoje"), ("Últimos 7 dias", "7"),
                                            ("Este mês", "mes"), ("Mês passado", "passado"))):
            Botao(atalhos, rotulo, lambda n=nome: self.atalho(n), "secundario", altura=34,
                  largura=px(120)).grid(row=i // 2, column=i % 2, sticky="ew", pady=(0, px(6)),
                                        padx=(0, px(6)) if i % 2 == 0 else (px(6), 0))

        tk.Label(dire, text="Considerar a data de", font=F.peq_b, fg=T.MUTED, bg=T.BG).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(px(6), px(5)))
        Segmentado(dire, [("emissao", "Emissão (no XML)"), ("arquivo", "Data do arquivo")],
                   app.var_criterio).grid(row=4, column=0, columnspan=2, sticky="w")
        self._bloco_pasta(dire).grid(row=5, column=0, columnspan=2, sticky="ew", pady=(px(14), 0))

    # -- modo
    def trocar_modo(self):
        modo = self.app.var_busca.get()
        self.quadro_lista.pack_forget()
        self.quadro_periodo.pack_forget()
        (self.quadro_lista if modo == "lista" else self.quadro_periodo).pack(fill="x")
        titulo, sub = self.TEXTOS[modo]
        self.lbl_titulo.configure(text=titulo)
        self.lbl_sub.configure(text=sub)
        self.btn.definir_texto("Buscar chaves  →" if modo == "lista" else "Buscar XMLs do período  →")
        self.banner.ocultar()
        self.atualizar_contagem()

    def periodo(self):
        ini, fim = ler_data(self.app.var_ini.get()), ler_data(self.app.var_fim.get())
        if ini and not self.app.var_fim.get().strip():
            fim = ini  # só a data inicial = apenas esse dia
        if ini and fim and fim < ini:
            ini, fim = fim, ini
        return (ini, fim) if ini and fim else (None, None)

    def atualizar_contagem(self):
        if self.app.var_busca.get() == "lista":
            n = len(ler_chaves_do_texto(self.caixa.obter()))
            self.chip.definir(f"{n} chave(s) detectada(s)" if n else "Nenhuma chave detectada",
                              "ok" if n else "info")
            self.btn.habilitar(n > 0)
            return
        ini, fim = self.periodo()
        if ini:
            dias = (fim - ini).days + 1
            self.chip.definir(f"{fmt_data(ini)} → {fmt_data(fim)}  ·  {dias} dia(s)", "ok")
        else:
            algo = self.app.var_ini.get().strip() or self.app.var_fim.get().strip()
            self.chip.definir("Data inválida (use dd/mm/aaaa)" if algo else "Selecione o período",
                              "aviso" if algo else "info")
        self.btn.habilitar(bool(ini))

    # -- periodo
    def ao_escolher_periodo(self, ini, fim):
        self.app.var_ini.set(fmt_data(ini))
        self.app.var_fim.set(fmt_data(fim))
        self.atualizar_contagem()

    def digitou(self):
        ini, fim = ler_data(self.app.var_ini.get()), ler_data(self.app.var_fim.get())
        if ini:
            self.calendario.definir(ini, fim if fim and fim >= ini else None)
        self.atualizar_contagem()

    def atalho(self, nome):
        hoje = datetime.date.today()
        if nome == "hoje":
            ini = fim = hoje
        elif nome == "7":
            ini, fim = hoje - datetime.timedelta(days=6), hoje
        elif nome == "mes":
            ini, fim = hoje.replace(day=1), hoje
        else:
            fim = hoje.replace(day=1) - datetime.timedelta(days=1)
            ini = fim.replace(day=1)
        self.calendario.definir(ini, fim)
        self.ao_escolher_periodo(ini, fim)

    # -- lista
    def escolher_pasta(self):
        pasta = filedialog.askdirectory(title="Pasta onde estão os arquivos", initialdir=self.app.var_pasta.get())
        if pasta:
            self.app.var_pasta.set(pasta)

    def escolher_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a lista de notas (JSON ou TXT)", initialdir=self.app.var_pasta.get(),
            filetypes=[("Lista de notas", "*.json *.txt"), ("Todos os arquivos", "*.*")])
        if not caminho:
            return
        try:
            conteudo = ler_arquivo_texto(caminho)
        except Exception as exc:
            self.banner.mostrar(f"Não foi possível abrir o arquivo: {exc}", "erro")
            return
        self.caixa.definir(conteudo)
        n = len(ler_chaves_do_texto(conteudo))
        nome = os.path.basename(caminho)
        if n:
            self.banner.ocultar()
            self.zona.definir(nome, f"{n} chave(s) detectada(s)  ·  clique para trocar o arquivo", True)
        else:
            self.zona.definir(nome, "Nenhuma chave encontrada  ·  clique para escolher outro arquivo", False)
            self.banner.mostrar("Nenhuma chave de 44 dígitos foi encontrada nesse arquivo.", "erro")
        self.atualizar_contagem()

    def avancar(self):
        app = self.app
        if app.var_busca.get() == "lista":
            chaves = ler_chaves_do_texto(self.caixa.obter())
            if not chaves:
                self.banner.mostrar("Selecione um arquivo ou cole uma lista com chaves de 44 dígitos.", "erro")
                return
            app.estado.update(modo="lista", chaves=chaves, datas={})
        else:
            ini, fim = self.periodo()
            if not ini:
                self.banner.mostrar("Selecione o período no calendário ou digite datas válidas (dd/mm/aaaa).", "erro")
                return
            app.estado.update(modo="periodo", ini=ini, fim=fim, criterio=app.var_criterio.get())
        app.concluir(0)
        app.ir(1)


class PaginaChaves(Pagina):
    TITULO = "Chaves encontradas"
    COLS_LISTA = ("n", "modelo", "num", "chave", "arq", "status")
    COLS_PERIODO = ("n", "modelo", "num", "chave", "data")

    def montar(self):
        self.var_filtro = tk.StringVar(value="todas")
        topo = tk.Frame(self.corpo, bg=T.BG)
        topo.pack(fill="x")
        self.chips = tk.Frame(topo, bg=T.BG)
        self.chips.pack(side="left")
        self.seg_filtro = Segmentado(topo, [("todas", "Todas"), ("ok", "Com arquivo"), ("falta", "Sem arquivo")],
                                     self.var_filtro, comando=self.preencher)
        self.seg_filtro.pack(side="right")

        self.barra = ttk.Progressbar(self.corpo, style="Accent.Horizontal.TProgressbar", mode="indeterminate")
        self.lbl_busca = tk.Label(self.corpo, text="", font=F.peq, fg=T.MUTED, bg=T.BG)

        cartao = self.cartao = Cartao(self.corpo, pad=2)
        cartao.pack(fill="both", expand=True, pady=(px(12), 0))
        colunas = (("n", "#", 40, "center"), ("modelo", "Modelo", 68, "w"), ("num", "Nº", 64, "w"),
                   ("chave", "Chave de acesso", px(390), "w"), ("data", "Emissão", 96, "w"),
                   ("arq", "Arquivos", 76, "center"), ("status", "Status", 100, "w"))
        self.tree = ttk.Treeview(cartao.interno, columns=[c[0] for c in colunas], show="headings",
                                 selectmode="browse", displaycolumns=self.COLS_LISTA)
        for ident, titulo, largura, ancora in colunas:
            self.tree.heading(ident, text=titulo, anchor="w" if ancora == "w" else "center")
            self.tree.column(ident, width=px(largura) if ident != "chave" else largura, anchor=ancora,
                             stretch=ident in ("status", "data"))
        self.tree.tag_configure("ok", foreground=T.TEXT)
        self.tree.tag_configure("falta", foreground=T.WARN)
        rolagem = ttk.Scrollbar(cartao.interno, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=rolagem.set)
        rolagem.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Configure>", lambda _e: self._ajustar())

        Botao(self.rodape, "←  Voltar", lambda: self.app.ir(0), "fantasma", largura=px(110)).pack(side="left")
        self.btn = Botao(self.rodape, "Mover para backup  →", self.avancar, "primario", largura=px(230))
        self.btn.pack(side="right")
        self.btn.habilitar(False)

    def ao_entrar(self):
        app = self.app
        est = app.estado
        base, destino = app.pastas()
        periodo = est.get("modo") == "periodo"
        recursivo = app.var_recursivo.get()
        sufixo = "  (com subpastas)" if recursivo else ""
        self.banner.ocultar()
        self.tree.delete(*self.tree.get_children())
        self.btn.habilitar(False)
        self.btn.definir_texto(f"Mover para {app.var_destino.get() or PASTA_DESTINO_PADRAO}  →")
        self.tree.configure(displaycolumns=self.COLS_PERIODO if periodo else self.COLS_LISTA)
        self.after(80, self._ajustar)
        if periodo:
            self.seg_filtro.pack_forget()
            self.var_filtro.set("todas")
            self.lbl_titulo.configure(text="XMLs do período")
            self.lbl_sub.configure(text=f"{fmt_data(est['ini'])} a {fmt_data(est['fim'])}  ·  procurando em: {base}{sufixo}")
        else:
            self.seg_filtro.pack(side="right")
            self.lbl_titulo.configure(text="Chaves encontradas")
            self.lbl_sub.configure(text=f"Procurando em: {base}{sufixo}")
        self.barra.pack(fill="x", pady=(px(12), 0), before=self.cartao)
        self.barra.start(12)
        self.lbl_busca.configure(text="Lendo arquivos da pasta..." if periodo else "Procurando arquivos na pasta...")
        self.lbl_busca.pack(anchor="w", pady=(px(6), 0), before=self.cartao)

        if periodo:
            ini, fim, criterio = est["ini"], est["fim"], est["criterio"]

            def tarefa():
                return localizar_xml_por_periodo(
                    base, destino, ini, fim, recursivo, criterio,
                    lambda i, n: app.postar(lambda: self.lbl_busca.configure(text=f"Lendo XMLs...  {i} de {n}")))

            app.rodar(tarefa, self.fim_busca_periodo)
        else:
            chaves = est["chaves"]
            self.tree.column("chave", width=max((F.corpo.measure(c) for c in chaves[:300]), default=0) + px(10))
            app.rodar(lambda: localizar_arquivos(base, chaves, destino, recursivo), self.fim_busca)

    def _ajustar(self):
        """A última coluna visível ocupa o espaço que sobra (sem vão antes da barra de rolagem)."""
        cols = list(self.tree.cget("displaycolumns"))
        if not cols or cols == ["#all"]:
            return
        fixo = sum(int(self.tree.column(c, "width")) for c in cols[:-1])
        minimo = {"data": 96, "status": 100}.get(cols[-1], 90)
        self.tree.column(cols[-1], width=max(self.tree.winfo_width() - fixo, px(minimo)))

    def _parar_busca(self):
        self.barra.stop()
        self.barra.pack_forget()
        self.lbl_busca.pack_forget()
        for w in self.chips.winfo_children():
            w.destroy()

    def fim_busca(self, encontradas):
        self._parar_busca()
        self.app.estado["encontradas"] = encontradas
        chaves = self.app.estado["chaves"]
        com = sum(1 for c in chaves if c in encontradas)
        arquivos = sum(len(v) for v in encontradas.values())
        Chip(self.chips, f"{len(chaves)} chave(s)", "info").pack(side="left")
        Chip(self.chips, f"{com} com arquivo  ·  {arquivos} arquivo(s)", "ok").pack(side="left", padx=px(8))
        if len(chaves) - com:
            Chip(self.chips, f"{len(chaves) - com} sem arquivo", "aviso").pack(side="left")
        self.preencher()
        if arquivos:
            self.btn.habilitar(True)
        else:
            self.banner.mostrar("Nenhum arquivo com essas chaves foi encontrado na pasta de busca. "
                                "Volte e confira a pasta.", "aviso")

    def fim_busca_periodo(self, res):
        self._parar_busca()
        est = self.app.estado
        encontradas = res["encontradas"]
        est["encontradas"], est["datas"] = encontradas, res["datas"]
        est["chaves"] = list(encontradas)
        arquivos = sum(len(v) for v in encontradas.values())
        Chip(self.chips, f"{len(encontradas)} XML(s) no período", "ok" if encontradas else "aviso").pack(side="left")
        Chip(self.chips, f"{res['total_xml']} XML(s) lidos na pasta", "info").pack(side="left", padx=px(8))
        if res["fallback"]:
            Chip(self.chips, f"{res['fallback']} sem data de emissão (usou a data do arquivo)",
                 "aviso").pack(side="left")
        self.preencher()
        if arquivos:
            self.btn.habilitar(True)
        else:
            self.banner.mostrar(f"Nenhum XML com data entre {fmt_data(est['ini'])} e {fmt_data(est['fim'])} "
                                "foi encontrado. Volte e confira o período, a pasta e o critério de data.",
                                "aviso")

    def preencher(self):
        self.tree.delete(*self.tree.get_children())
        est = self.app.estado
        encontradas, datas = est.get("encontradas", {}), est.get("datas", {})
        filtro = self.var_filtro.get()
        for i, chave in enumerate(est["chaves"], 1):
            arquivos = encontradas.get(chave, [])
            if (filtro == "ok" and not arquivos) or (filtro == "falta" and arquivos):
                continue
            modelo, numero = info_chave(chave)
            self.tree.insert("", "end", values=(i, modelo, numero, chave, fmt_data(datas.get(chave)),
                                                len(arquivos), "Encontrada" if arquivos else "Sem arquivo"),
                             tags=("ok" if arquivos else "falta",))

    def ao_erro(self):
        self._parar_busca()

    def avancar(self):
        self.app.concluir(1)
        self.app.ir(2)


class PaginaBackup(Pagina):
    TITULO = "Mover para a pasta de backup"
    SUBTITULO = "Os arquivos encontrados serão levados para a pasta de destino."

    def montar(self):
        app = self.app
        cartao = Cartao(self.corpo)
        cartao.pack(fill="x")
        linha = tk.Frame(cartao.interno, bg=T.CARD)
        linha.pack(fill="x")
        linha.columnconfigure(0, weight=1)
        Campo(linha, app.var_destino, "Nome da pasta de destino").grid(row=0, column=0, sticky="ew",
                                                                       padx=(0, px(16)))
        col = tk.Frame(linha, bg=T.CARD)
        col.grid(row=0, column=1, sticky="s")
        tk.Label(col, text="Ação", font=F.peq_b, fg=T.MUTED, bg=T.CARD).pack(anchor="w")
        Segmentado(col, [("mover", "Mover"), ("copiar", "Copiar")], app.var_modo,
                   comando=self.atualizar_resumo).pack(pady=(px(5), 0))
        app.var_destino.trace_add("write", lambda *_a: self.atualizar_resumo())
        self.resumo = tk.Label(cartao.interno, text="", font=F.corpo, fg=T.MUTED, bg=T.CARD, justify="left",
                               anchor="w", wraplength=px(640))
        self.resumo.pack(anchor="w", fill="x", pady=(px(14), 0))

        self.barra = ttk.Progressbar(self.corpo, style="Accent.Horizontal.TProgressbar", mode="determinate")
        self.barra.pack(fill="x", pady=(px(18), 0))
        self.lbl_prog = tk.Label(self.corpo, text="Pronto para começar.", font=F.peq, fg=T.MUTED, bg=T.BG)
        self.lbl_prog.pack(anchor="w", pady=(px(6), 0))
        self.chips = tk.Frame(self.corpo, bg=T.BG)
        self.chips.pack(anchor="w", pady=(px(8), 0))
        self.log = CaixaTexto(self.corpo, altura=7, mono=True, somente_leitura=True)
        self.log.pack(fill="both", expand=True, pady=(px(10), 0))

        self.btn_voltar = Botao(self.rodape, "←  Voltar", lambda: self.app.ir(1), "fantasma", largura=px(110))
        self.btn_voltar.pack(side="left")
        self.btn_mover = Botao(self.rodape, "Mover agora", self.mover, "primario", largura=px(170))
        self.btn_abrir = Botao(self.rodape, "Abrir pasta", self.abrir, "secundario", largura=px(140))
        self.btn_seguir = Botao(self.rodape, "ZIP e e-mail  →", lambda: self.app.ir(3), "primario",
                                largura=px(180))
        self._rodape(False)

    def _rodape(self, pronto):
        for b in (self.btn_mover, self.btn_abrir, self.btn_seguir):
            b.pack_forget()
        if pronto:
            self.btn_seguir.pack(side="right")
            self.btn_abrir.pack(side="right", padx=(0, px(10)))
        else:
            self.btn_mover.pack(side="right")

    def atualizar_resumo(self):
        app = self.app
        encontradas = app.estado.get("encontradas", {})
        n = sum(len(v) for v in encontradas.values())
        verbo = "copiados" if app.var_modo.get() == "copiar" else "movidos"
        _base, destino = app.pastas()
        self.resumo.configure(text=f"{n} arquivo(s) de {len(encontradas)} nota(s) serão {verbo} para:\n{destino}")
        self.btn_mover.definir_texto("Copiar agora" if app.var_modo.get() == "copiar" else "Mover agora")

    def ao_entrar(self):
        self.banner.ocultar()
        self.log.limpar()
        for w in self.chips.winfo_children():
            w.destroy()
        self.barra.configure(value=0)
        self.lbl_prog.configure(text="Pronto para começar.")
        self.btn_voltar.habilitar(True)
        self.btn_mover.habilitar(True)
        self._rodape(False)
        self.atualizar_resumo()

    def ao_erro(self):
        self.btn_mover.habilitar(True)
        self.btn_voltar.habilitar(True)

    def progresso(self, i, n):
        self.barra.configure(maximum=max(n, 1), value=i)
        self.lbl_prog.configure(text=f"{i} de {n} arquivo(s)")

    def mover(self):
        app = self.app
        encontradas = app.estado["encontradas"]
        _base, destino = app.pastas()
        copiar = app.var_modo.get() == "copiar"
        self.banner.ocultar()
        self.log.limpar()
        self.btn_mover.habilitar(False)
        self.btn_voltar.habilitar(False)
        self.lbl_prog.configure(text="Processando...")

        def tarefa():
            return mover_arquivos(encontradas, destino, copiar, lambda i, n: app.postar(lambda: self.progresso(i, n)))

        app.rodar(tarefa, lambda res: self.fim(res, copiar, destino))

    def fim(self, res, copiar, destino):
        app = self.app
        verbo = "copiados" if copiar else "movidos"
        for w in self.chips.winfo_children():
            w.destroy()
        Chip(self.chips, f"{len(res['processados'])} {verbo}", "ok").pack(side="left")
        if res["ignorados"]:
            Chip(self.chips, f"{len(res['ignorados'])} já existiam", "aviso").pack(side="left", padx=px(8))
        if res["erros"]:
            Chip(self.chips, f"{len(res['erros'])} com erro", "erro").pack(side="left")
        linhas = [f"OK       {n}" for n in res["processados"]]
        linhas += [f"JÁ EXISTE {n}" for n in res["ignorados"]]
        linhas += [f"ERRO     {e}" for e in res["erros"]]
        faltantes = [c for c in app.estado["chaves"] if c not in app.estado["encontradas"]]
        linhas += [f"SEM ARQUIVO {c}" for c in faltantes]
        self.log.definir("\n".join(linhas) + "\n")
        self.lbl_prog.configure(text="Concluído.")
        app.estado["lote"] = set(res["processados"]) | set(res["ignorados"])
        self.resumo.configure(text=f"{len(res['processados'])} arquivo(s) {verbo} para:\n{destino}")
        if res["erros"]:
            self.banner.mostrar("O backup terminou, mas alguns arquivos deram erro. Veja o detalhe abaixo.", "aviso")
        else:
            self.banner.mostrar(f"Backup concluído em {destino}", "ok")
        app.concluir(2)
        self.btn_voltar.habilitar(True)
        self._rodape(True)

    def abrir(self):
        abrir_pasta(self.app.pastas()[1])


class PaginaEnvio(Pagina):
    TITULO = "ZIP e envio por e-mail"
    SUBTITULO = "Compacte os arquivos do backup e envie o pacote para quem precisar."

    def montar(self):
        app = self.app
        v = app.v
        self._auto = {}
        opcoes = tk.Frame(self.corpo, bg=T.BG)
        opcoes.pack(fill="x")
        Alternador(opcoes, "Zipar somente arquivos .xml", app.var_somente_xml).pack(side="left")
        Alternador(opcoes, "Só os arquivos desta busca", app.var_so_lote).pack(side="left", padx=(px(28), 0))

        c1 = Cartao(self.corpo)
        c1.pack(fill="x", pady=(px(14), 0))
        g = tk.Frame(c1.interno, bg=T.CARD)
        g.pack(fill="x")
        g.columnconfigure(0, weight=1, uniform="a")
        g.columnconfigure(1, weight=1, uniform="a")
        tk.Label(g, text="Mensagem", font=F.h2, fg=T.TEXT, bg=T.CARD).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, px(8)))
        Campo(g, v["destinatario"], "Destinatário(s)  ·  separe vários por vírgula").grid(
            row=1, column=0, columnspan=2, sticky="ew")
        Campo(g, v["assunto"], "Assunto").grid(row=2, column=0, sticky="ew", padx=(0, px(8)), pady=(px(10), 0))
        Campo(g, v["mensagem"], "Texto do e-mail").grid(row=2, column=1, sticky="ew", padx=(px(8), 0),
                                                         pady=(px(10), 0))

        c2 = Cartao(self.corpo)
        c2.pack(fill="x", pady=(px(14), 0))
        topo = tk.Frame(c2.interno, bg=T.CARD)
        topo.pack(fill="x")
        tk.Label(topo, text="Envio", font=F.h2, fg=T.TEXT, bg=T.CARD).pack(side="left")
        Segmentado(topo, [("resend", "Resend"), ("gmail", "Gmail"), ("outlook", "Outlook")],
                   v["provider"], comando=self.atualizar_status).pack(side="right")
        linha = tk.Frame(c2.interno, bg=T.CARD)
        linha.pack(fill="x", pady=(px(12), 0))
        self.status = Chip(linha)
        self.status.pack(side="left")
        Botao(linha, "Configurar e-mail…", app.abrir_config, "secundario", altura=36,
              largura=px(190)).pack(side="right")

        Botao(self.rodape, "←  Voltar", lambda: self.app.ir(2), "fantasma", largura=px(110)).pack(side="left")
        self.btn_enviar = Botao(self.rodape, "Criar ZIP e enviar", self.enviar, "primario", largura=px(190))
        self.btn_enviar.pack(side="right")
        self.btn_zip = Botao(self.rodape, "Somente criar ZIP", self.criar_zip, "secundario", largura=px(180))
        self.btn_zip.pack(side="right", padx=(0, px(10)))

    def atualizar_status(self):
        cfg = self.app.cfg_email()
        p = cfg["provider"]
        if p == "resend":
            if cfg["resend_api_key"]:
                self.status.definir(f"Resend  ·  API key configurada  ·  remetente {cfg['resend_from'] or RESEND_FROM_PADRAO}", "ok")
            else:
                self.status.definir("Resend  ·  API key não configurada (Ctrl+T)", "aviso")
        elif cfg["smtp_user"] and cfg["smtp_password"]:
            self.status.definir(f"{p.capitalize()}  ·  conta {cfg['smtp_user']}", "ok")
        else:
            self.status.definir(f"{p.capitalize()}  ·  e-mail e senha de app não configurados (Ctrl+T)", "aviso")

    def _sugerir_textos(self):
        """No modo período, sugere assunto/texto com as datas (sem apagar o que o usuário escreveu)."""
        est, v = self.app.estado, self.app.v
        if est.get("modo") == "periodo":
            periodo = f"{fmt_data(est['ini'])} a {fmt_data(est['fim'])}"
            sugestoes = {"assunto": (f"RataoXML - XMLs de {periodo}", "RataoXML - Backup XML"),
                         "mensagem": (f"Segue em anexo os XMLs do período de {periodo}.", MENSAGEM_PADRAO)}
        else:
            sugestoes = {"assunto": ("RataoXML - Backup XML", ""), "mensagem": (MENSAGEM_PADRAO, "")}
        for campo, (novo, padrao) in sugestoes.items():
            atual = v[campo].get().strip()
            if atual in (padrao, self._auto.get(campo, "")) or (not atual):
                v[campo].set(novo)
                self._auto[campo] = novo

    def _params_zip(self):
        app = self.app
        est = app.estado
        lote = est.get("lote") if app.var_so_lote.get() else None
        nome = None
        if est.get("modo") == "periodo":
            nome = f"xml_{est['ini']:%Y%m%d}_a_{est['fim']:%Y%m%d}.zip"
        return (lote or None), nome

    def ao_entrar(self):
        self.banner.ocultar()
        self.lbl_sub.configure(text=f"Pasta: {self.app.pastas()[1]}")
        self._sugerir_textos()
        self.atualizar_status()
        self.travar(False)

    def travar(self, flag):
        self.btn_enviar.habilitar(not flag)
        self.btn_zip.habilitar(not flag)

    def ao_erro(self):
        self.travar(False)

    def criar_zip(self):
        _base, destino = self.app.pastas()
        somente = self.app.var_somente_xml.get()
        nomes, nome = self._params_zip()
        self.banner.ocultar()
        self.travar(True)
        self.app.rodar(lambda: zipar_pasta(destino, somente, nomes, nome), self.fim)

    def enviar(self):
        app = self.app
        cfg = app.cfg_email()
        validos, invalidos = parse_destinatarios(cfg["destinatario"])
        if not validos or invalidos:
            ruim = f" ({', '.join(invalidos)})" if invalidos else ""
            self.banner.mostrar(f"Informe pelo menos um e-mail de destinatário válido{ruim}.", "erro")
            return
        _base, destino = app.pastas()
        somente = app.var_somente_xml.get()
        nomes, nome = self._params_zip()
        self.banner.mostrar("Criando o ZIP e enviando o e-mail...", "info")
        self.travar(True)
        app.rodar(lambda: zipar_e_enviar(cfg, destino, cfg["destinatario"], somente, nomes, nome), self.fim)

    def fim(self, res):
        self.travar(False)
        self.banner.mostrar(res["mensagem"], "ok" if res["sucesso"] else "erro")
        if res["sucesso"]:
            self.app.concluir(3)


# ------------------------------------------------------------------ aplicativo
class JanelaConfig(tk.Toplevel):
    """Tela de configurações (Ctrl+T): provedor, contas de e-mail, chaves de API e padrões."""

    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("Configurações — RataoXML")
        self.configure(bg=T.BG)
        self.resizable(False, False)
        self.transient(app.root)
        try:
            icone = caminho_recurso("RataoXML.ico")
            if os.path.isfile(icone):
                self.iconbitmap(icone)
        except tk.TclError:
            pass
        # cópias editáveis: só valem quando clicar em Salvar
        self.d = {k: tk.StringVar(value=var.get()) for k, var in app.v.items()}
        self.d_destino = tk.StringVar(value=app.var_destino.get())
        self.d_somente_xml = tk.BooleanVar(value=app.var_somente_xml.get())
        self.aba = tk.StringVar(value="envio")
        self._montar()
        w, h = px(680), px(640)
        x = app.root.winfo_rootx() + (app.root.winfo_width() - w) // 2
        y = app.root.winfo_rooty() + max(10, (app.root.winfo_height() - h) // 2)
        self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")
        self.bind("<Escape>", lambda _e: self.destroy())
        self.bind("<Control-s>", lambda _e: self.salvar())
        self.update_idletasks()
        titulo_escuro(self)
        try:
            self.grab_set()
        except tk.TclError:
            pass

    def _montar(self):
        d = self.d
        cab = tk.Frame(self, bg=T.BG)
        cab.pack(fill="x", padx=px(28), pady=(px(24), 0))
        tk.Label(cab, text="Configurações", font=F.titulo, fg=T.TEXT, bg=T.BG).pack(anchor="w")
        tk.Label(cab, text="Contas de e-mail, chaves de API e padrões de envio.", font=F.corpo, fg=T.MUTED,
                 bg=T.BG).pack(anchor="w", pady=(px(4), 0))

        rod = tk.Frame(self, bg=T.BG)
        rod.pack(side="bottom", fill="x", padx=px(28), pady=(0, px(22)))
        Botao(rod, "Cancelar", self.destroy, "fantasma", largura=px(110)).pack(side="left")
        Botao(rod, "Salvar", self.salvar, "primario", largura=px(130)).pack(side="right")
        self.btn_testar = Botao(rod, "Testar envio", self.testar, "secundario", largura=px(150))
        self.btn_testar.pack(side="right", padx=(0, px(10)))

        Segmentado(self, [("envio", "Envio"), ("resend", "Resend"), ("smtp", "Gmail / Outlook"),
                          ("padroes", "Padrões")], self.aba, comando=self._mostrar_aba).pack(
            anchor="w", padx=px(28), pady=(px(16), px(12)))
        self.area = tk.Frame(self, bg=T.BG, height=px(300))
        self.area.pack(fill="x", padx=px(28))
        self.area.pack_propagate(False)
        self.avisos = tk.Frame(self, bg=T.BG)
        self.avisos.pack(fill="x", padx=px(28))
        self.banner = Banner(self.avisos)

        def nota(pai, texto):
            return tk.Label(pai, text=texto, font=F.peq, fg=T.MUTED, bg=T.CARD, justify="left", anchor="w",
                            wraplength=px(580))

        # aba Envio
        c = Cartao(self.area)
        i = c.interno
        tk.Label(i, text="Provedor padrão", font=F.peq_b, fg=T.MUTED, bg=T.CARD).pack(anchor="w")
        Segmentado(i, [("resend", "Resend"), ("gmail", "Gmail"), ("outlook", "Outlook")],
                   d["provider"]).pack(anchor="w", pady=(px(5), px(12)))
        Campo(i, d["destinatario"], "Destinatários  ·  separe vários por vírgula").pack(fill="x")
        linha = tk.Frame(i, bg=T.CARD)
        linha.pack(fill="x", pady=(px(10), 0))
        linha.columnconfigure(0, weight=1, uniform="x")
        linha.columnconfigure(1, weight=1, uniform="x")
        Campo(linha, d["assunto"], "Assunto padrão").grid(row=0, column=0, sticky="ew", padx=(0, px(8)))
        Campo(linha, d["mensagem"], "Texto do e-mail").grid(row=0, column=1, sticky="ew", padx=(px(8), 0))
        # aba Resend
        r = Cartao(self.area)
        Campo(r.interno, d["resend_api_key"], "API key do Resend", senha=True).pack(fill="x")
        Campo(r.interno, d["resend_from"], "Remetente").pack(fill="x", pady=(px(10), 0))
        nota(r.interno, "Crie a chave em resend.com/api-keys. Sem domínio verificado, o Resend só entrega para o "
                        "e-mail da própria conta (remetente onboarding@resend.dev). Para enviar a outras pessoas, "
                        "verifique um domínio em resend.com/domains e use um remetente desse domínio, "
                        "por exemplo nfe@seudominio.com.br.").pack(fill="x", pady=(px(14), 0))
        # aba Gmail / Outlook
        g = Cartao(self.area)
        Campo(g.interno, d["smtp_user"], "Usuário (e-mail da conta)").pack(fill="x")
        Campo(g.interno, d["smtp_password"], "Senha de app", senha=True).pack(fill="x", pady=(px(10), 0))
        nota(g.interno, "Gmail: gere a senha de app em myaccount.google.com/apppasswords (exige verificação em duas "
                        "etapas). Outlook: use uma senha de app da conta Microsoft. Não use a senha normal. "
                        "Escolha o provedor na aba Envio.").pack(fill="x", pady=(px(14), 0))
        # aba Padroes
        p = Cartao(self.area)
        Campo(p.interno, self.d_destino, "Nome da pasta de backup").pack(fill="x")
        Alternador(p.interno, "Zipar somente arquivos .xml", self.d_somente_xml).pack(anchor="w", pady=(px(14), 0))
        nota(p.interno, f"As configurações ficam salvas em:\n{ARQUIVO_CONFIG}\n\nA API key e a senha ficam em texto "
                        "simples nesse arquivo: não o compartilhe.").pack(fill="x", pady=(px(14), 0))
        self.abas = {"envio": c, "resend": r, "smtp": g, "padroes": p}
        self._mostrar_aba()

    def _mostrar_aba(self):
        for cartao in self.abas.values():
            cartao.pack_forget()
        self.abas[self.aba.get()].pack(fill="both", expand=True)

    def _cfg(self):
        return {k: v.get().strip() for k, v in self.d.items()}

    def testar(self):
        cfg = self._cfg()
        validos, invalidos = parse_destinatarios(cfg["destinatario"])
        if not validos or invalidos:
            self.aba.set("envio")
            self._mostrar_aba()
            self.banner.mostrar("Informe um destinatário válido na aba Envio para enviar o teste.", "erro")
            return
        self.banner.mostrar("Enviando e-mail de teste...", "info")
        self.btn_testar.habilitar(False)
        self.app.rodar(lambda: testar_envio(cfg, cfg["destinatario"]), self._fim_teste)

    def _fim_teste(self, res):
        if not self.winfo_exists():
            return
        self.btn_testar.habilitar(True)
        self.banner.mostrar(res["mensagem"], "ok" if res["sucesso"] else "erro")

    def salvar(self):
        cfg = self._cfg()
        if cfg["destinatario"]:
            validos, invalidos = parse_destinatarios(cfg["destinatario"])
            if invalidos or not validos:
                self.aba.set("envio")
                self._mostrar_aba()
                self.banner.mostrar(f"E-mail inválido: {', '.join(invalidos) or cfg['destinatario']}", "erro")
                return
        app = self.app
        for chave, valor in cfg.items():
            app.v[chave].set(valor)
        app.var_destino.set(self.d_destino.get().strip() or PASTA_DESTINO_PADRAO)
        app.var_somente_xml.set(self.d_somente_xml.get())
        try:
            salvar_config(app.valores_config())
        except Exception as exc:
            self.banner.mostrar(f"Não foi possível salvar: {exc}", "erro")
            return
        app.paginas[3].atualizar_status()
        self.banner.mostrar("Configurações salvas.", "ok")
        self.after(700, self.destroy)


class App:
    def __init__(self, root):
        self.root = root
        cfg = config_padrao()
        self.var_pasta = tk.StringVar(value=os.getcwd())
        self.var_recursivo = tk.BooleanVar(value=True)
        self.var_destino = tk.StringVar(value=cfg["destino"])
        self.var_modo = tk.StringVar(value="mover")
        self.var_somente_xml = tk.BooleanVar(value=bool(cfg["somente_xml"]))
        self.var_busca = tk.StringVar(value="lista")
        self.var_criterio = tk.StringVar(value="emissao")
        self.var_ini = tk.StringVar(value="")
        self.var_fim = tk.StringVar(value="")
        self.var_so_lote = tk.BooleanVar(value=True)
        self.v = {k: tk.StringVar(value=cfg[k]) for k in
                  ("provider", "destinatario", "assunto", "mensagem", "resend_api_key", "resend_from",
                   "smtp_user", "smtp_password")}
        self.estado = {"modo": "lista", "chaves": [], "encontradas": {}, "datas": {}, "lote": set()}
        self.fila = queue.Queue()
        self.atual, self.concluidos = 0, set()
        self.janela_config = None

        self.lateral = Lateral(root, self.abrir_config)
        self.lateral.pack(side="left", fill="y")
        tk.Frame(root, bg=T.BORDER, width=1).pack(side="left", fill="y")
        self.conteudo = tk.Frame(root, bg=T.BG)
        self.conteudo.pack(side="left", fill="both", expand=True)
        self.paginas = [PaginaLista(self), PaginaChaves(self), PaginaBackup(self), PaginaEnvio(self)]
        for p in self.paginas:
            p.place(x=0, y=0, relwidth=1, relheight=1)
        self.ir(0)
        root.after(60, self._processar_fila)

    # navegacao
    def ir(self, i):
        self.atual = i
        self.lateral.atualizar(i, self.concluidos)
        pagina = self.paginas[i]
        pagina.tkraise()
        pagina.ao_entrar()

    def concluir(self, i):
        self.concluidos.add(i)
        self.lateral.atualizar(self.atual, self.concluidos)

    def abrir_config(self):
        janela = self.janela_config
        try:
            if janela is not None and janela.winfo_exists():
                janela.lift()
                janela.focus_force()
                return
        except tk.TclError:
            pass
        self.janela_config = JanelaConfig(self)

    # configuracao
    def pastas(self):
        base = self.var_pasta.get().strip() or os.getcwd()
        return base, os.path.join(base, self.var_destino.get().strip() or PASTA_DESTINO_PADRAO)

    def cfg_email(self):
        return {k: v.get().strip() for k, v in self.v.items()}

    def valores_config(self):
        cfg = self.cfg_email()
        cfg["destino"] = self.var_destino.get().strip() or PASTA_DESTINO_PADRAO
        cfg["somente_xml"] = self.var_somente_xml.get()
        return cfg

    # tarefas em segundo plano (a janela nunca trava)
    def postar(self, funcao):
        self.fila.put(funcao)

    def rodar(self, tarefa, ao_terminar):
        def alvo():
            try:
                res = tarefa()
            except Exception as exc:
                self.fila.put(lambda: self._erro(exc))
                return
            self.fila.put(lambda: ao_terminar(res))

        threading.Thread(target=alvo, daemon=True).start()

    def _erro(self, exc):
        pagina = self.paginas[self.atual]
        pagina.banner.mostrar(f"Erro: {exc}", "erro")
        pagina.ao_erro()

    def _processar_fila(self):
        try:
            while True:
                self.fila.get_nowait()()
        except queue.Empty:
            pass
        self.root.after(60, self._processar_fila)


# ----------------------------------------------------------- bandeja do sistema
_ICONE_BANDEJA = None


def abrir_da_bandeja(root):
    root.after(0, lambda: (root.deiconify(), root.lift(), root.focus_force()))


def sair_da_bandeja(root):
    if _ICONE_BANDEJA is not None:
        _ICONE_BANDEJA.stop()
    root.after(0, root.destroy)


def fechar_janela(root):
    """Com ícone na bandeja, fechar apenas esconde a janela; sem bandeja, encerra o programa."""
    if _ICONE_BANDEJA is not None:
        root.withdraw()
    else:
        root.destroy()


def iniciar_bandeja(root):
    global _ICONE_BANDEJA
    if pystray is None or Image is None:
        return None
    caminho_icone = caminho_recurso("RataoXML.ico")
    if not os.path.isfile(caminho_icone):
        return None
    try:
        _ICONE_BANDEJA = pystray.Icon(
            "RataoXML", Image.open(caminho_icone), "RataoXML",
            menu=pystray.Menu(
                pystray.MenuItem("Abrir RataoXML", lambda icon, item: abrir_da_bandeja(root), default=True),
                pystray.MenuItem("Sair", lambda icon, item: sair_da_bandeja(root)),
            ),
        )
        threading.Thread(target=_ICONE_BANDEJA.run, daemon=True).start()
        return _ICONE_BANDEJA
    except Exception:
        _ICONE_BANDEJA = None
        return None


def iniciar_interface():
    ativar_dpi()
    root = tk.Tk()
    _ESC[0] = max(1.0, root.winfo_fpixels("1i") / 96.0)
    criar_fontes(root)
    criar_estilos(root)
    root.title("RataoXML")
    root.configure(bg=T.BG)
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    w, h = min(px(1080), sw - 40), min(px(780), sh - 80)
    root.geometry(f"{w}x{h}+{(sw - w) // 2}+{max(0, (sh - h) // 2 - 20)}")
    root.minsize(min(px(980), w), min(px(700), h))
    try:
        icone = caminho_recurso("RataoXML.ico")
        if os.path.isfile(icone):
            root.iconbitmap(icone)
    except tk.TclError:
        pass
    root.protocol("WM_DELETE_WINDOW", lambda: fechar_janela(root))
    iniciar_bandeja(root)
    app = App(root)
    root.bind("<Control-t>", lambda _e: app.abrir_config())
    root.bind("<Control-T>", lambda _e: app.abrir_config())
    titulo_escuro(root)
    root.mainloop()


# ----------------------------------------------------------------- console
def parse_args():
    ap = argparse.ArgumentParser(description="RataoXML - move arquivos de notas e envia o ZIP por e-mail.")
    ap.add_argument("lista", nargs="?", help="arquivo JSON/TXT com as chaves (padrão: lista.txt da pasta)")
    ap.add_argument("--de", help="busca por período: data inicial (dd/mm/aaaa)")
    ap.add_argument("--ate", help="busca por período: data final (padrão: hoje)")
    ap.add_argument("--data-arquivo", action="store_true", help="no período, usar a data do arquivo e não a emissão")
    ap.add_argument("-d", "--destino", default=PASTA_DESTINO_PADRAO, help="nome da pasta de destino")
    ap.add_argument("-r", "--recursivo", action="store_true", help="procurar também nas subpastas")
    ap.add_argument("--copiar", action="store_true", help="copiar em vez de mover")
    ap.add_argument("--zip", action="store_true", help="gerar o ZIP da pasta de destino")
    ap.add_argument("--todos", action="store_true", help="no ZIP, incluir todos os arquivos (não só .xml)")
    ap.add_argument("--pasta-toda", action="store_true", help="no ZIP, incluir a pasta toda (não só esta busca)")
    ap.add_argument("--email", help="destinatário(s): gera o ZIP e envia por e-mail")
    ap.add_argument("--provider", choices=["resend", "gmail", "outlook"], help="provedor de e-mail")
    return ap.parse_args()


def main_console(args):
    pasta_base = os.getcwd()
    pasta_destino = os.path.join(pasta_base, args.destino)
    nome_zip = None

    if args.de:
        ini = ler_data(args.de)
        fim = ler_data(args.ate) if args.ate else datetime.date.today()
        if not ini or not fim:
            print("Datas inválidas. Use o formato dd/mm/aaaa.")
            return
        if fim < ini:
            ini, fim = fim, ini
        res = localizar_xml_por_periodo(pasta_base, pasta_destino, ini, fim, args.recursivo,
                                        "arquivo" if args.data_arquivo else "emissao")
        print(f"{len(res['encontradas'])} XML(s) entre {fmt_data(ini)} e {fmt_data(fim)} "
              f"({res['total_xml']} lidos na pasta).")
        mv = mover_arquivos(res["encontradas"], pasta_destino, args.copiar)
        verbo = "copiados" if args.copiar else "movidos"
        print(f"Arquivos {verbo}: {len(mv['processados'])}")
        for n in mv["processados"]:
            print("  -", n)
        if mv["ignorados"]:
            print(f"Ignorados (já existem na pasta de destino): {len(mv['ignorados'])}")
        for e in mv["erros"]:
            print("  ERRO:", e)
        lote = set(mv["processados"]) | set(mv["ignorados"])
        nome_zip = f"xml_{ini:%Y%m%d}_a_{fim:%Y%m%d}.zip"
    else:
        caminho_lista = args.lista or os.path.join(pasta_base, "lista.txt")
        if not os.path.isfile(caminho_lista):
            print(f"Lista não encontrada: {caminho_lista}")
            return
        res = localizar_e_mover(pasta_base, ler_arquivo_texto(caminho_lista), pasta_destino, args.recursivo,
                                args.copiar)
        if not res["sucesso"]:
            print(res["mensagem"])
            return
        print(resumo_mover(res, args.copiar))
        lote = set(res["processados"]) | set(res["ignorados"])

    if args.email or args.zip:
        cfg = config_padrao()
        if args.provider:
            cfg["provider"] = args.provider
        somente_xml = not args.todos
        nomes = None if args.pasta_toda else (lote or None)
        if args.email:
            r = zipar_e_enviar(cfg, pasta_destino, args.email, somente_xml, nomes, nome_zip)
        else:
            r = zipar_pasta(pasta_destino, somente_xml, nomes, nome_zip)
        print(r["mensagem"])


def main():
    if len(sys.argv) > 1:
        main_console(parse_args())
        return
    if tk is None:
        print("Interface gráfica indisponível (tkinter não instalado). Use: python procurar.py lista.txt")
        return
    iniciar_interface()


if __name__ == "__main__":
    main()