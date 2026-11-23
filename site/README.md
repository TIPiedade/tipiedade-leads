# site/ — o que vai para o cPanel

Esta pasta é o que está publicado em **ops.tipiedade.com** (alojamento WebHS, cPanel).
O GitHub deixa de servir a aplicação: aqui só se guarda o código.

## Estrutura no servidor

```
public_html/ops/
├── index.html        ← o hub (esta pasta)
├── catalogo.html
├── manifest.json, sw.js, icon*.png, .htaccess
├── img/              ← fotos do catálogo
└── crm/              ← cópia de docs/ (CRM antigo e restantes ferramentas)
```

A pasta `crm/` **não está aqui** para não duplicar ficheiros: é a pasta `docs/`
deste repositório, com duas diferenças aplicadas no momento de publicar —
os vídeos de `ficheiros/` e o `HIST_EMBEBIDO` removido do `index.html`.

## Publicar uma alteração

1. Editar os ficheiros em `site/`
2. Zipar o conteúdo e extrair em `public_html/ops/` pelo Gestor de Ficheiros
3. No telemóvel, fechar e reabrir a app (o service worker vai buscar a versão nova)

## Dados

A aplicação lê e escreve o `historico.json` deste repositório através do
Cloudflare Worker, com a sessão do utilizador. Funciona com o repositório
privado — o Worker autentica-se com o seu próprio token.
