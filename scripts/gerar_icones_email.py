"""
Gera os PNG alojados dos ícones SVG usados nos emails do CRM (docs/email/i/<hash>.png)
e o logotipo branco (docs/email/logo-white.png). Correr depois de mudar ícones ou
logotipo no gerarHTMLEmail do docs/index.html. Requer: pip install playwright pillow;
playwright install chromium.
"""
import base64, hashlib, io, os, re, subprocess
from PIL import Image
from playwright.sync_api import sync_playwright

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INDEX = os.path.join(RAIZ, "docs", "index.html")
DEST = os.path.join(RAIZ, "docs", "email")
GRUPOS = ["restaurante", "hotel", "pastelaria", "catering", "distribuidor", "gourmet", "cafe", "cervejaria"]

src = open(INDEX, encoding="utf8").read()
b64 = re.search(r"var LOGO_WHITE_B64 = 'data:image/png;base64,([^']*)';", src).group(1)
im = Image.open(io.BytesIO(base64.b64decode(b64)))
im.resize((round(im.width * 156 / im.height), 156), Image.LANCZOS).save(os.path.join(DEST, "logo-white.png"), optimize=True)

svgs = set()
for g in GRUPOS:
    for n in (1, 2, 3, 4):
        h = subprocess.run(["node", os.path.join(RAIZ, "scripts", "render_email_crm.js"), INDEX, g, str(n)],
                           capture_output=True, text=True, check=True).stdout
        svgs.update(re.findall(r"<svg\b.*?</svg>", h, re.S))
os.makedirs(os.path.join(DEST, "i"), exist_ok=True)
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(device_scale_factor=3)
    for s in svgs:
        nome = hashlib.sha1(s.encode()).hexdigest()[:10] + ".png"
        pg.set_content('<html><body style="margin:0">' + s.replace("<svg", '<svg id="x"', 1) + "</body></html>")
        pg.locator("#x").screenshot(path=os.path.join(DEST, "i", nome), omit_background=True)
    b.close()
print(f"{len(svgs)} ícones + logotipo em {DEST}")
