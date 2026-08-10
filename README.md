# Ti'Piedade — Sistema Completo de Prospeção HORECA

## O que é este sistema

Um sistema integrado de geração de leads, nurturing automático e CRM, construído sobre GitHub (gratuito, sem servidor).

## Fluxo semanal automático

| Dia | Hora | O que acontece |
|-----|------|----------------|
| **Segunda** | 07h30 | Excel completo enviado para `sales@tipiedade.com` para revisão |
| **Quarta** | 07h30 | Leads + email de nurturing enviados a cada comercial. Histórico atualizado. |

## Leads gerados por semana

| Quem recebe | Quantidade | Canal |
|-------------|-----------|-------|
| Nuno | 20 leads | HORECA — Lisboa, Santarém, Linha Sintra-Cascais |
| João | 20 leads | HORECA — Lisboa, Margem Sul, Costa Oeste, Leiria |
| Óscar | 20 leads | HORECA — Ericeira-Caldas, Coimbra, Porto, Braga, Guimarães |
| Rui | 5 leads | Catering & Eventos (Portugal inteiro) |
| Rui | 5 leads | Distribuidores de congelados (zonas não cobertas) |
| **Total** | **70 leads/semana** | |

## CRM (GitHub Pages)

Disponível em: `https://TIPiedade.github.io/tipiedade-leads/`

Funcionalidades:
- Dashboard com KPIs e funil de vendas
- Vistas por comercial, canal e estado
- Registo de visitas com resultado e notas
- Registo de desistências com motivo
- Histórico de nurturing por lead
- Pesquisa e filtros
- Sincronização automática com o `historico.json`

---

## Configuração

### 1. Secrets (Settings → Secrets → Actions)

| Secret | Descrição |
|--------|-----------|
| `SMTP_HOST` | Servidor SMTP |
| `SMTP_PORT` | Porta (587) |
| `SMTP_USER` | Email de envio |
| `SMTP_PASS` | Password |
| `EMAIL_NUNO` | Email do Nuno |
| `EMAIL_JOAO` | Email do João |
| `EMAIL_OSCAR` | Email do Óscar |
| `EMAIL_RUI` | Email do Rui (sales@tipiedade.com) |
| `GH_PAT` | GitHub Personal Access Token (scope: repo) |

### 2. Ativar GitHub Pages

Settings → Pages → Source: **Deploy from branch** → Branch: `main` → Folder: `/docs`

### 3. Criar GH_PAT

GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic) → Generate new token → scope: `repo` → copiar e adicionar como secret `GH_PAT`

### 4. Ativar workflow

Actions → confirmar ativação → pronto.

---

## Estrutura

```
tipiedade-leads/
├── .github/workflows/leads_semanais.yml   ← automação (2ª e 4ª)
├── scripts/
│   ├── generate_leads.py                  ← lógica principal
│   └── leads_extra.py                     ← base catering + distribuidores
├── docs/
│   └── index.html                         ← CRM (GitHub Pages)
├── historico.json                         ← base de dados (auto-atualizado)
└── README.md
```
