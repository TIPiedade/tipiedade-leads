"""
Envio avulso de nurturing a leads de uma campanha registada no historico.json.

O envio semanal (generate_leads.py, quarta-feira) só trata as leads que o próprio
sistema gera. Este script envia o email N da sequência às leads do CRM marcadas
com um campo "campanha" (ex.: "casamentos-2026"), com a base de email aprovada no
CRM (scripts/email_crm.py).

Variáveis de ambiente:
  CAMPANHA       nome da campanha (obrigatório)
  EMAIL_NUM      número do email da sequência (1–4, por omissão 1)
  ENVIAR         "sim" envia às leads; "teste" envia um único exemplo só para EMAIL_RUI
                 (sem registar nada); qualquer outro valor só lista (ensaio)
  BREVO_API_KEY, GITHUB_TOKEN, GITHUB_REPOSITORY, EMAIL_RUI — como no generate_leads.py
"""
import os, sys, time

sys.path.insert(0, os.path.dirname(__file__))
import generate_leads as g
import email_crm

CAMPANHA  = os.environ.get("CAMPANHA", "").strip()
EMAIL_NUM = int(os.environ.get("EMAIL_NUM", "1") or 1)
MODO_ENVIO = os.environ.get("ENVIAR", "nao").strip().lower()
ENVIAR    = MODO_ENVIO == "sim"
TESTE     = MODO_ENVIO == "teste"


def emails_validos(campo):
    return [e.strip() for e in (campo or "").replace(";", ",").split(",") if "@" in e.strip()]


def ja_enviado(lead, num):
    return any(e.get("num") == num for e in lead.get("emails_enviados", []))


def selecionar(historico):
    sel = []
    for k, v in historico["leads"].items():
        if v.get("campanha") != CAMPANHA or v.get("nurturing_excluir"):
            continue
        if v.get("estado") == "Desistiu" or ja_enviado(v, EMAIL_NUM):
            continue
        if EMAIL_NUM > 1 and not ja_enviado(v, EMAIL_NUM - 1):
            continue
        if emails_validos(v.get("email_lead")):
            sel.append((k, v))
    return sorted(sel, key=lambda kv: (kv[1].get("comercial", ""), kv[1].get("nome", "")))


def main():
    if not CAMPANHA:
        print("[ERRO] Indica a CAMPANHA.", flush=True)
        raise SystemExit(1)
    historico, _ = g.ler_historico()
    sel = selecionar(historico)
    grupo_assunto = g.ASSUNTOS["catering"].get(EMAIL_NUM, "")
    print(f"▶ Campanha {CAMPANHA} — Email {EMAIL_NUM} — {len(sel)} destinatários — {'ENVIO' if ENVIAR else ('TESTE (só para ' + g.EMAIL_RUI + ')' if TESTE else 'ENSAIO (nada é enviado)')}", flush=True)
    for k, v in sel:
        print(f"  {v.get('comercial',''):6} | {v['nome']} → {v['email_lead']}", flush=True)
    # Mesmo no ensaio, gerar o email para confirmar que a base do CRM está em ordem
    for grupo in sorted({g.tipologia_grupo(v.get("tipo", "") + " " + v.get("tipologia_cliente", "")) for _, v in sel}):
        email_crm.html_email(grupo, EMAIL_NUM)
        print(f"  ✓ email {EMAIL_NUM} ({grupo}) gerado com a base do CRM", flush=True)
    if TESTE and sel:
        k, v = sel[0]
        grupo = g.tipologia_grupo(v.get("tipo", "") + " " + v.get("tipologia_cliente", ""))
        assunto = g.ASSUNTOS.get(grupo, g.ASSUNTOS["restaurante"]).get(EMAIL_NUM, "")
        html = email_crm.html_email(grupo, EMAIL_NUM)
        ok = g.brevo_send(g.EMAIL_RUI, "Teste campanha", f"[TESTE] {assunto} | Ti'Piedade", html, email_crm.texto_email(html))
        print(f"  {'✓' if ok else '✗'} teste enviado para {g.EMAIL_RUI} (nada registado no CRM)", flush=True)
        if not ok:
            raise SystemExit(1)
        return
    if not ENVIAR or not sel:
        return

    enviados, falhados = [], []
    for k, v in sel:
        grupo = g.tipologia_grupo(v.get("tipo", "") + " " + v.get("tipologia_cliente", ""))
        assunto = g.ASSUNTOS.get(grupo, g.ASSUNTOS["restaurante"]).get(EMAIL_NUM, "")
        # Base aprovada no CRM (mesma da pré-visualização), adaptada aos programas de email
        html = email_crm.html_email(grupo, EMAIL_NUM)
        texto = email_crm.texto_email(html)
        ok = False
        for dest in emails_validos(v["email_lead"]):
            if g.brevo_send(dest, v["nome"], f"{assunto} | Ti'Piedade", html, texto):
                ok = True
        (enviados if ok else falhados).append(k)
        print(f"  {'✓' if ok else '✗'} {v['nome']}", flush=True)
        time.sleep(1)

    # Reler o histórico imediatamente antes de gravar, para não apagar edições feitas no CRM durante o envio
    historico, sha = g.ler_historico()
    d = g.hoje()
    hora = time.strftime("%H:%M")
    for k in enviados:
        lead = historico["leads"].get(k)
        if not lead:
            continue
        lead.setdefault("emails_enviados", []).append({"num": EMAIL_NUM, "data": d, "assunto": f"Nurturing Email {EMAIL_NUM}"})
        lead.setdefault("actividade", []).append({"tipo": "email", "label": f"Email {EMAIL_NUM} de nurturing enviado",
                                                   "nota": f"Campanha {CAMPANHA}", "data": d, "hora": hora})
        if lead.get("estado") == "Em aberto":
            lead["estado"] = "Contactado"
    g.escrever_historico(historico, sha)

    resumo = (f"Campanha {CAMPANHA} — Email {EMAIL_NUM} (\"{grupo_assunto}\")\n\n"
              f"Enviados: {len(enviados)}\nFalhados: {len(falhados)}\n\n"
              + "\n".join(f"✓ {historico['leads'][k]['comercial']} — {historico['leads'][k]['nome']}" for k in enviados if k in historico["leads"])
              + ("\n\nFalhados:\n" + "\n".join(falhados) if falhados else ""))
    g.brevo_send_resumo(g.EMAIL_RUI, f"[Campanha {CAMPANHA}] Email {EMAIL_NUM} enviado a {len(enviados)} leads | Ti'Piedade", resumo)
    print(f"\n📧 {len(enviados)} enviados, {len(falhados)} falhados", flush=True)
    if falhados and not enviados:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
