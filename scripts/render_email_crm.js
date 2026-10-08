// Gera o HTML de um email de nurturing com a mesma função do CRM (docs/index.html → gerarHTMLEmail),
// para o email enviado ser igual à pré-visualização aprovada no CRM.
// Uso: node render_email_crm.js <index.html> <grupo> <num>
const fs = require('fs');
const [, , file, grupo, num] = process.argv;
const src = fs.readFileSync(file, 'utf8');
const logo = src.match(/var LOGO_WHITE_B64 = '[^']*';/);
const i = src.indexOf('function gerarHTMLEmail(grupo, num) {');
const j = src.indexOf('\nfunction ', i + 10);
if (!logo || i < 0 || j < 0) { console.error('gerarHTMLEmail não encontrado em ' + file); process.exit(1); }
const gerar = new Function(logo[0] + '\n' + src.slice(i, j) + '\nreturn gerarHTMLEmail;')();
process.stdout.write(gerar(grupo, Number(num)));
