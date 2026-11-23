#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Resumo de segunda-feira — Ti'Piedade
Envia a cada comercial, às 07h00 de segunda, o que tem para fazer na semana:
alvos de prospeção, fila de chamadas e o que fechou na semana anterior.

Lê o historico.json do repositório e envia pela API da Brevo — o mesmo
mecanismo e os mesmos segredos que o generate_leads.py já usa:
  BREVO_API_KEY, EMAIL_NUNO, EMAIL_JOAO, EMAIL_OSCAR, EMAIL_RUI
Não é preciso configurar nada de novo.
"""
import json, os, sys, urllib.request, urllib.error
from datetime import date, timedelta

EMAIL_FROM = "sales@tipiedade.com"
EMAIL_CC   = "sales@tipiedade.com"
HUB        = "https://tipiedade.github.io/tipiedade-leads/ops/"
LOGO       = "https://tipiedade.github.io/tipiedade-leads/email/logo-white.png"
META_CONTACTOS = int(os.environ.get("META_CONTACTOS", "10"))

COMERCIAIS = {
    "nuno":  os.environ.get("EMAIL_NUNO", ""),
    "joao":  os.environ.get("EMAIL_JOAO", ""),
    "oscar": os.environ.get("EMAIL_OSCAR", ""),
}

NOMES  = {"nuno": "Nuno", "joao": "João", "oscar": "Óscar", "rui": "Rui", "ana": "Ana Filipa"}
DIAS   = {1: "segunda", 2: "terça", 3: "quarta", 4: "quinta", 5: "sexta"}
MESES  = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"]
ARQUIVADOS = {"Desistiu", "N/A"}


def carregar():
    caminho = os.path.join(os.path.dirname(__file__), "..", "historico.json")
    with open(caminho, encoding="utf-8") as f:
        return json.load(f).get("leads", {})


def do_user(v, u):
    return v.get("canal_id") == u or (v.get("comercial") or "").lower().startswith(u)


def ativo(v):
    return v.get("estado", "Em aberto") not in ARQUIVADOS


def trabalhada(v):
    # contacto humano: visita, chamada ou mudança de estado. Emails automáticos não contam.
    return bool(v.get("visitas")) or bool((v.get("cc") or {}).get("chamadas")) \
        or v.get("estado", "Em aberto") != "Em aberto"


def recolher(leads, u, hoje):
    sem_passada = (hoje - timedelta(days=7)).isoformat()
    meus = [(k, v) for k, v in leads.items() if do_user(v, u) and ativo(v)]

    alvos = [v for k, v in meus
             if v.get("estado") != "Cliente" and not (v.get("cc") or {}).get("ativo")]
    # por zona, as que ainda não foram tocadas
    zonas = {}
    for v in alvos:
        if not trabalhada(v):
            zonas[v.get("zona") or "—"] = zonas.get(v.get("zona") or "—", 0) + 1
    top_zonas = sorted(zonas.items(), key=lambda x: -x[1])[:4]

    prontas = [v for k, v in meus
               if len(v.get("emails_enviados") or []) >= 4 and v.get("estado") != "Cliente"][:6]

    chamadas = {}
    for k, v in meus:
        cc = v.get("cc") or {}
        if cc.get("ativo") and cc.get("dia"):
            chamadas.setdefault(int(cc["dia"]), []).append(v)

    alerta = []
    for k, v in meus:
        ch = (v.get("cc") or {}).get("chamadas") or []
        if len(ch) >= 2 and all(c.get("res") == "sr" for c in ch[-2:]):
            alerta.append(v)

    feitas = sum(
        len([x for x in (v.get("visitas") or []) if x.get("data", "") >= sem_passada]) +
        len([x for x in ((v.get("cc") or {}).get("chamadas") or []) if x.get("data", "") >= sem_passada])
        for k, v in meus)
    encomendas = sum(
        len([x for x in ((v.get("cc") or {}).get("chamadas") or [])
             if x.get("data", "") >= sem_passada and x.get("res") == "enc"])
        for k, v in meus)
    novos = [v for k, v in meus if v.get("estado") == "Cliente" and
             any(x.get("data", "") >= sem_passada for x in (v.get("visitas") or []))]

    return dict(total=len(meus), alvos=len(alvos), top_zonas=top_zonas, prontas=prontas,
                chamadas=chamadas, alerta=alerta, feitas=feitas,
                encomendas=encomendas, novos=len(novos))


def bloco(titulo, corpo):
    return f"""
    <tr><td style="padding:0 28px 18px">
      <div style="font:600 11px/1 Inter,Helvetica,Arial,sans-serif;letter-spacing:.11em;text-transform:uppercase;color:#A6937F;padding-bottom:11px">{titulo}</div>
      {corpo}
    </td></tr>"""


def html(u, d, hoje):
    nome = NOMES.get(u, u.title())
    data_txt = f"{hoje.day} de {MESES[hoje.month - 1]}"
    pct = round(d["feitas"] / META_CONTACTOS * 100) if META_CONTACTOS else 0
    cor_meta = "#3C8A58" if d["feitas"] >= META_CONTACTOS else "#E8951B"

    zonas = "".join(
        f"""<tr><td style="padding:9px 0;border-bottom:1px solid #EFE7DD">
            <span style="font:600 14px Inter,Helvetica,Arial,sans-serif;color:#1E120B">{z}</span></td>
        <td align="right" style="padding:9px 0;border-bottom:1px solid #EFE7DD">
            <span style="font:600 14px Inter,Helvetica,Arial,sans-serif;color:#9A5C06">{n} por tocar</span></td></tr>"""
        for z, n in d["top_zonas"]) or \
        """<tr><td style="font:14px Inter,Helvetica,Arial,sans-serif;color:#6F5C4B;padding:9px 0">Sem zonas por abrir — boa altura para recuperar clientes antigos.</td></tr>"""

    prontas = "".join(
        f"""<tr><td style="padding:9px 0;border-bottom:1px solid #EFE7DD">
            <span style="font:600 14px Inter,Helvetica,Arial,sans-serif;color:#1E120B">{v.get('nome','—')}</span><br>
            <span style="font:13px Inter,Helvetica,Arial,sans-serif;color:#A6937F">{v.get('zona','')} · {v.get('tipo','')}</span></td></tr>"""
        for v in d["prontas"]) or \
        """<tr><td style="font:14px Inter,Helvetica,Arial,sans-serif;color:#6F5C4B;padding:9px 0">Nenhuma a aguardar visita.</td></tr>"""

    chamadas = "".join(
        f"""<tr><td style="padding:9px 0;border-bottom:1px solid #EFE7DD">
            <span style="font:600 14px Inter,Helvetica,Arial,sans-serif;color:#1E120B;text-transform:capitalize">{DIAS.get(dia, dia)}</span></td>
        <td align="right" style="padding:9px 0;border-bottom:1px solid #EFE7DD">
            <span style="font:600 14px Inter,Helvetica,Arial,sans-serif;color:#4A95A0">{len(vs)} chamadas</span></td></tr>"""
        for dia, vs in sorted(d["chamadas"].items())) or \
        """<tr><td style="font:14px Inter,Helvetica,Arial,sans-serif;color:#6F5C4B;padding:9px 0">Ainda não há clientes na ronda semanal.</td></tr>"""

    alerta_html = ""
    if d["alerta"]:
        linhas = "".join(f"<div style='font:14px Inter,Helvetica,Arial,sans-serif;color:#BE4238;padding:3px 0'>• <b>{v.get('nome','—')}</b> — {v.get('zona','')}</div>"
                         for v in d["alerta"][:5])
        alerta_html = f"""
        <tr><td style="padding:0 28px 18px">
          <div style="background:#FBEBE9;border:1px solid #EFC9C4;border-radius:12px;padding:16px 18px">
            <div style="font:700 13px Inter,Helvetica,Arial,sans-serif;color:#BE4238;padding-bottom:7px">⚠ Passar lá esta semana</div>
            <div style="font:13px Inter,Helvetica,Arial,sans-serif;color:#6F5C4B;padding-bottom:9px">Duas chamadas seguidas sem resposta:</div>
            {linhas}
          </div></td></tr>"""

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#FAF7F3">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#FAF7F3;padding:26px 14px">
<tr><td align="center">
<table width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#FFFFFF;border-radius:18px;overflow:hidden;box-shadow:0 2px 12px rgba(43,23,16,.07)">

  <tr><td style="background:#2B1710;padding:26px 28px">
    <img src="{LOGO}" alt="Ti'Piedade" width="42" style="display:block;border:0">
    <div style="font:600 11px/1 Inter,Helvetica,Arial,sans-serif;letter-spacing:.19em;text-transform:uppercase;color:#E8951B;padding:16px 0 7px">Segunda-feira, {data_txt}</div>
    <div style="font:600 25px/1.2 Georgia,serif;color:#FFFFFF">Bom dia, {nome}</div>
    <div style="font:14px/1.5 Inter,Helvetica,Arial,sans-serif;color:rgba(255,255,255,.6);padding-top:7px">
      {d['alvos']} alvos por abrir · {sum(len(v) for v in d['chamadas'].values())} chamadas marcadas esta semana</div>
  </td></tr>

  <tr><td style="padding:24px 28px 18px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td width="33%" style="padding-right:8px">
        <div style="background:#FAF7F3;border-radius:12px;padding:14px">
          <div style="font:600 26px/1 Georgia,serif;color:{cor_meta}">{d['feitas']}</div>
          <div style="font:12px Inter,Helvetica,Arial,sans-serif;color:#A6937F;padding-top:4px">contactos na semana passada</div></div></td>
      <td width="33%" style="padding:0 4px">
        <div style="background:#FAF7F3;border-radius:12px;padding:14px">
          <div style="font:600 26px/1 Georgia,serif;color:#1E120B">{d['encomendas']}</div>
          <div style="font:12px Inter,Helvetica,Arial,sans-serif;color:#A6937F;padding-top:4px">encomendas registadas</div></div></td>
      <td width="33%" style="padding-left:8px">
        <div style="background:#FAF7F3;border-radius:12px;padding:14px">
          <div style="font:600 26px/1 Georgia,serif;color:#3C8A58">{d['novos']}</div>
          <div style="font:12px Inter,Helvetica,Arial,sans-serif;color:#A6937F;padding-top:4px">clientes novos</div></div></td>
    </tr></table>
    <div style="font:13px/1.5 Inter,Helvetica,Arial,sans-serif;color:#6F5C4B;padding-top:12px">
      Meta da semana: <b style="color:#1E120B">{META_CONTACTOS} contactos</b>. Semana passada ficaste em {pct}%.</div>
  </td></tr>

  {alerta_html}

  {bloco("Onde abrir portas hoje", f'<table width="100%" cellpadding="0" cellspacing="0">{zonas}</table>')}
  {bloco("Prontas para visita", f'<table width="100%" cellpadding="0" cellspacing="0">{prontas}</table>')}
  {bloco("A fábrica liga a estes clientes", f'<table width="100%" cellpadding="0" cellspacing="0">{chamadas}</table>')}

  <tr><td style="padding:6px 28px 30px" align="center">
    <a href="{HUB}" style="display:inline-block;background:#E8951B;color:#2B1710;font:600 15px Inter,Helvetica,Arial,sans-serif;text-decoration:none;padding:14px 30px;border-radius:99px">Abrir a lista de hoje</a>
    <div style="font:12px Inter,Helvetica,Arial,sans-serif;color:#A6937F;padding-top:13px">Carteira ativa: {d['total']} leads</div>
  </td></tr>

  <tr><td style="background:#FAF7F3;padding:18px 28px;border-top:1px solid #EFE7DD" align="center">
    <div style="font:12px/1.6 Inter,Helvetica,Arial,sans-serif;color:#A6937F">
      Pão de Ló Ti'Piedade · Departamento Comercial<br>
      <a href="mailto:sales@tipiedade.com" style="color:#9A5C06;text-decoration:none">sales@tipiedade.com</a></div>
  </td></tr>
</table></td></tr></table></body></html>"""


def texto(u, d):
    nome = NOMES.get(u, u.title())
    linhas = [f"Bom dia, {nome}.", "",
              f"{d['alvos']} alvos por abrir · {sum(len(v) for v in d['chamadas'].values())} chamadas marcadas.",
              f"Semana passada: {d['feitas']} contactos, {d['encomendas']} encomendas, {d['novos']} clientes novos.", ""]
    if d["top_zonas"]:
        linhas.append("Onde abrir portas:")
        linhas += [f"  - {z}: {n} por tocar" for z, n in d["top_zonas"]]
        linhas.append("")
    if d["alerta"]:
        linhas.append("Passar lá (2 chamadas sem resposta):")
        linhas += [f"  - {v.get('nome','—')} ({v.get('zona','')})" for v in d["alerta"][:5]]
        linhas.append("")
    linhas.append(f"Lista de hoje: {HUB}")
    return "\n".join(linhas)


def brevo_send(to_email, to_name, subject, html_content, text_content=None):
    """Envia pela API HTTP da Brevo. Mesmo padrão do generate_leads.py."""
    api_key = os.environ.get("BREVO_API_KEY", "")
    if not api_key:
        print("[ERRO] BREVO_API_KEY não configurada.", flush=True)
        raise SystemExit(1)

    payload = {
        "sender": {"name": "Pão de Ló Ti'Piedade", "email": EMAIL_FROM},
        "to": [{"email": to_email, "name": to_name}],
        "bcc": [{"email": EMAIL_CC}],
        "subject": subject,
        "htmlContent": html_content,
    }
    if text_content:
        payload["textContent"] = text_content

    req = urllib.request.Request(
        "https://api.brevo.com/v3/smtp/email",
        data=json.dumps(payload).encode("utf-8"),
        headers={"api-key": api_key, "Content-Type": "application/json", "Accept": "application/json"},
        method="POST")
    try:
        with urllib.request.urlopen(req) as r:
            return r.status in (200, 201)
    except urllib.error.HTTPError as e:
        print(f"[Brevo] Erro HTTP {e.code}: {e.read().decode()[:300]}", flush=True)
        if e.code == 401:
            raise SystemExit(1)
        return False


def main():
    seco = os.environ.get("ENVIAR", "sim").strip().lower() not in ("sim", "yes", "true", "1")
    leads = carregar()
    hoje = date.today()
    enviados = 0

    for u, destino in COMERCIAIS.items():
        d = recolher(leads, u, hoje)
        chamadas = sum(len(v) for v in d["chamadas"].values())
        assunto = f"A tua semana · {d['alvos']} alvos e {chamadas} chamadas"
        if not destino:
            print(f"[aviso] sem email para {u} (segredo EMAIL_{u.upper()}) — ignorado")
            continue
        if seco:
            print(f"[simulação] {u} <{destino}>: {assunto} | {d['feitas']} contactos na semana passada")
            continue
        if brevo_send(destino, NOMES.get(u, u), assunto, html(u, d, hoje), texto(u, d)):
            enviados += 1
            print(f"enviado a {u} <{destino}>: {d['alvos']} alvos, {d['feitas']} contactos")
        else:
            print(f"[falhou] {u} <{destino}>")

    print(f"{enviados} email(s) enviado(s).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
