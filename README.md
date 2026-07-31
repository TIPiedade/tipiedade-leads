# Ti'Piedade — Leads HORECA Automáticos

Envia automaticamente 20 leads por comercial **todas as segundas-feiras às 07h30** (Lisboa), com o Excel em anexo por email.

---

## Configuração (uma única vez)

### 1. Criar repositório no GitHub

1. Vai a [github.com](https://github.com) e cria conta gratuita (se não tiveres)
2. Clica **New repository** → nome: `tipiedade-leads` → **Private** → Create
3. Faz upload de todos estes ficheiros para o repositório (arrasta e solta)

### 2. Adicionar os Secrets (credenciais)

No repositório GitHub:
**Settings → Secrets and variables → Actions → New repository secret**

Adiciona estes 7 secrets:

| Secret | O que preencher |
|--------|----------------|
| `SMTP_HOST` | Servidor SMTP do teu email (ex: `mail.tipiedade.com`) |
| `SMTP_PORT` | Porta SMTP — normalmente `587` |
| `SMTP_USER` | Email de envio (ex: `comercial@tipiedade.com`) |
| `SMTP_PASS` | Password do email de envio |
| `EMAIL_NUNO` | Email do Nuno |
| `EMAIL_JOAO` | Email do João |
| `EMAIL_OSCAR` | Email do Óscar |

> **Nota:** O endereço `sales@tipiedade.com` recebe automaticamente cópia de todos os envios — está definido diretamente no script e não precisa de ser configurado aqui.

> **Não sabes o SMTP?** Pergunta ao teu provedor de email ou IT. Se usares Gmail pessoal, usa `smtp.gmail.com`, porta `587`, e cria uma "App Password" nas definições de segurança Google.

### 3. Ativar o workflow

1. No repositório, vai ao separador **Actions**
2. Se aparecer aviso de ativação, clica **"I understand my workflows, go ahead and enable them"**
3. Pronto — corre sozinho todas as segundas às 07h30

---

## Testar manualmente

Para testar sem esperar pela segunda-feira:

1. **Actions** → **Leads HORECA Semanais — Ti'Piedade**
2. **Run workflow** → **Run workflow**
3. Os emails chegam em segundos

O Excel gerado fica também guardado em **Actions → [execução] → Artifacts** por 30 dias.

---

## Atualizar a base de leads

Edita o ficheiro `scripts/generate_leads.py` diretamente no GitHub:
- Cada zona é um bloco `"Nome da Zona": [ ... ]`
- Cada lead é um dicionário com: `n` (nome), `t` (tipo), `m` (morada), `tel`, `email`, `p` (prioridade: Alta/Média/Baixa), `tCliente`, `gancho`
- Guarda o ficheiro → o próximo envio já usa os dados atualizados

---

## Estrutura do repositório

```
tipiedade-leads/
├── .github/
│   └── workflows/
│       └── leads_semanais.yml   ← agenda e configuração do envio
├── scripts/
│   └── generate_leads.py        ← lógica de geração e envio
└── README.md                    ← este ficheiro
```
