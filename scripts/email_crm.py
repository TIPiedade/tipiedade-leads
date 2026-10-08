"""
Emails de nurturing com a base aprovada no CRM.

O HTML vem da função gerarHTMLEmail do docs/index.html (a mesma da pré-visualização
no CRM), e é ajustado para os programas de email:
  - logotipo embutido (data:) → imagem alojada no GitHub Pages (o Gmail/Outlook bloqueiam data:)
  - ícones SVG → PNG alojados (o Gmail/Outlook não mostram SVG)
  - texto corrido justificado
"""
import hashlib, html as htmlmod, os, re, subprocess

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INDEX = os.path.join(RAIZ, "docs", "index.html")
RENDER = os.path.join(RAIZ, "scripts", "render_email_crm.js")
BASE_URL = os.environ.get("EMAIL_ASSETS_URL", "https://tipiedade.github.io/tipiedade-leads/email/")
ICONES = os.path.join(RAIZ, "docs", "email", "i")


def _svg_para_img(m):
    svg = m.group(0)
    nome = hashlib.sha1(svg.encode()).hexdigest()[:10] + ".png"
    if not os.path.exists(os.path.join(ICONES, nome)):
        raise RuntimeError(f"Ícone sem PNG alojado ({nome}) — gerar com scripts/gerar_icones_email.py")
    w = re.search(r'width="(\d+)"', svg)
    h = re.search(r'height="(\d+)"', svg)
    estilo = re.search(r'style="([^"]*)"', svg[:svg.find(">")])
    w, h = (w.group(1) if w else "24"), (h.group(1) if h else "24")
    st = "border:0;outline:none;text-decoration:none;" + (estilo.group(1) if estilo else "display:inline-block")
    return f'<img src="{BASE_URL}i/{nome}" width="{w}" height="{h}" alt="" style="{st}">'


def _justificar(m):
    tag = m.group(0)
    if "text-align" in tag or "line-height" not in tag:
        return tag
    return tag.replace('style="', 'style="text-align:justify;', 1)


def html_email(grupo, num):
    html = subprocess.run(["node", RENDER, INDEX, grupo, str(num)],
                          capture_output=True, text=True, check=True).stdout
    html = re.sub(r'src="data:image/png;base64,[^"]*"', f'src="{BASE_URL}logo-white.png"', html, count=1)
    html = re.sub(r"<svg\b.*?</svg>", _svg_para_img, html, flags=re.S)
    html = re.sub(r'<p style="[^"]*">', _justificar, html)
    if "data:image" in html or "<svg" in html:
        raise RuntimeError("Ficou conteúdo embutido que os programas de email bloqueiam")
    return html


def texto_email(html):
    t = re.sub(r"<br\s*/?>", "\n", html)
    t = re.sub(r"</(p|tr|div)>", "\n", t)
    t = re.sub(r"<[^>]+>", "", t)
    t = htmlmod.unescape(t)
    return re.sub(r"\n\s*\n+", "\n\n", t).strip()
