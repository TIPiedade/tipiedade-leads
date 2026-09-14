#!/usr/bin/env python3
"""
Actualiza o HIST_EMBEBIDO no docs/index.html com os dados do historico.json.
Correr semanalmente após novos dados serem carregados.
"""
import json, re, base64, urllib.request, os
from datetime import date

TOKEN = os.environ.get("GH_PAT","")
REPO  = os.environ.get("GITHUB_REPOSITORY","TIPiedade/tipiedade-leads")

def gh_get(path):
    url = f"https://api.github.com/repos/{REPO}/contents/{path}"
    req = urllib.request.Request(url, headers={"Authorization":f"token {TOKEN}","Accept":"application/vnd.github+json"})
    with urllib.request.urlopen(req) as r: return json.loads(r.read())

def gh_put(path, content, sha, msg):
    url = f"https://api.github.com/repos/{REPO}/contents/{path}"
    body = json.dumps({"message":msg,"content":base64.b64encode(content if isinstance(content,bytes) else content.encode()).decode(),"sha":sha}).encode()
    req = urllib.request.Request(url, data=body, method="PUT", headers={"Authorization":f"token {TOKEN}","Accept":"application/vnd.github+json","Content-Type":"application/json"})
    with urllib.request.urlopen(req) as r: return json.loads(r.read())

# Carregar historico.json
hist_data = gh_get("historico.json")
hist = json.loads(base64.b64decode(hist_data["content"]))
n = len(hist["leads"])
print(f"historico.json: {n} leads")

# Carregar index.html
html_data = gh_get("docs/index.html")
html = base64.b64decode(html_data["content"]).decode()
sha  = html_data["sha"]

# Substituir HIST_EMBEBIDO
hist_json = json.dumps(hist, ensure_ascii=False, separators=(',',':'))
new_emb = f"// Dados embutidos — {date.today()} — {n} leads\nvar HIST_EMBEBIDO = {hist_json};\n"
html_new = re.sub(r"// Dados embutidos[^\n]*\nvar HIST_EMBEBIDO = \{[\s\S]*?\};\n", new_emb, html, count=1)

if html_new == html:
    # Tentar padrão alternativo
    idx = html.find("var HIST_EMBEBIDO =")
    if idx < 0:
        print("ERRO: HIST_EMBEBIDO não encontrado")
        exit(1)
    line_start = html.rfind("\n", 0, idx) + 1
    line_end = html.find(";\n", idx) + 2
    html_new = html[:line_start] + new_emb + html[line_end:]

print(f"HTML antigo: {len(html):,} chars")
print(f"HTML novo:   {len(html_new):,} chars")

res = gh_put("docs/index.html", html_new, sha, f"auto: {n} leads embutidas ({date.today()})")
print(f"✓ HTML actualizado: {res['commit']['sha'][:12]}")
