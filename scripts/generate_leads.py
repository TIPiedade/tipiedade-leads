"""
Pão de Ló Ti'Piedade — Sistema Completo de Prospeção HORECA
- 20 leads/semana por comercial (HORECA por zonas)
- 5 leads/semana catering & eventos (para Rui)
- 5 leads/semana distribuidores congelados (para Rui)
- Nurturing automático (4 emails por lead)
- Histórico guardado no GitHub
MODO=coordenador → segunda-feira → resumo para Rui
MODO=comerciais  → quarta-feira → leads + nurturing para equipa
"""

import math, smtplib, os, datetime, json, base64, urllib.request, urllib.error
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import sys, importlib.util

EMAIL_FROM = "sales@tipiedade.com"          # remetente
EMAIL_CC   = "sales@tipiedade.com"          # BCC 1
EMAIL_BCC2 = "geral@tipiedade.com"           # BCC 2
EMAIL_RUI  = os.environ.get("EMAIL_RUI", EMAIL_CC)
REPO_OWNER = os.environ.get("GITHUB_REPOSITORY","TIPiedade/tipiedade-leads").split("/")[0]
REPO_NAME  = os.environ.get("GITHUB_REPOSITORY","TIPiedade/tipiedade-leads").split("/")[1]
HIST_FILE  = "historico.json"

# ── Importar base real (leads dos ficheiros Excel) ───────────
spec2 = importlib.util.spec_from_file_location("db_horeca_generated", os.path.join(os.path.dirname(__file__), "db_horeca_generated.py"))
db_real = importlib.util.module_from_spec(spec2)
spec2.loader.exec_module(db_real)
DB_HORECA_REAL = db_real.get_db_horeca()

# ── Importar base extra (catering + distribuidores) ──────────
spec = importlib.util.spec_from_file_location("leads_extra", os.path.join(os.path.dirname(__file__), "leads_extra.py"))
extra = importlib.util.module_from_spec(spec)
spec.loader.exec_module(extra)
DB_CATERING      = extra.DB_CATERING
DB_DISTRIBUIDORES = extra.DB_DISTRIBUIDORES

# ════════════════════════════════════════════════════════════════
# BASE DE LEADS HORECA (comerciais)
# ════════════════════════════════════════════════════════════════

DB_HORECA = {
"Lisboa": [
  {"n":"Tasca do Chico","t":"Restaurante","m":"R. do Diário de Notícias 39, Lisboa","tel":"965 059 670","email":"info@tascadochico.pt","p":"Alta","tCliente":"Tasca contemporânea","gancho":"Rotatividade alta de turistas — sobremesa pronta a servir com zero desperdício."},
  {"n":"Solar dos Presuntos","t":"Restaurante","m":"R. das Portas de Santo Antão 150, Lisboa","tel":"213 424 253","email":"geral@solardospresuntos.com","p":"Alta","tCliente":"Restaurante tradicional / turismo","gancho":"Volume de turistas — pão de ló é produto de eleição para terminar uma refeição típica."},
  {"n":"Martinho da Arcada","t":"Restaurante","m":"Pr. do Comércio 3, Lisboa","tel":"218 879 259","email":"geral@martinhodaarcada.pt","p":"Alta","tCliente":"Histórico / turismo","gancho":"O produto mais português para o restaurante mais antigo de Lisboa."},
  {"n":"Pastéis de Belém","t":"Pastelaria","m":"R. de Belém 84, Lisboa","tel":"213 637 423","email":"geral@pasteisdebelem.pt","p":"Alta","tCliente":"Pastelaria icónica / turismo","gancho":"Duas referências da doçaria artesanal portuguesa — produto complementar."},
  {"n":"Landeau Chocolate","t":"Café","m":"R. das Flores 70, Lisboa","tel":"214 792 178","email":"hello@landeau.pt","p":"Alta","tCliente":"Café de nicho / produto","gancho":"Valoriza produtos com história — dose individual em congelado resolve logística."},
  {"n":"Taberna Rua das Flores","t":"Restaurante","m":"R. das Flores 103, Lisboa","tel":"213 479 418","email":"info@tabernaruadasflores.pt","p":"Média","tCliente":"Tasca contemporânea","gancho":"Rotatividade alta — produto pronto a servir elimina desperdício."},
  {"n":"Pharmácia","t":"Restaurante","m":"R. Marechal Saldanha 1, Lisboa","tel":"213 465 146","email":"geral@museudafarmacia.pt","p":"Média","tCliente":"Restaurante temático / cultura","gancho":"Clientela atenta à origem — pão de ló como sobremesa de autor."},
  {"n":"Decadente","t":"Restaurante","m":"R. de São Pedro de Alcântara 45, Lisboa","tel":"213 957 936","email":"info@odecadente.pt","p":"Média","tCliente":"Bistrô / residentes","gancho":"Carta rotativa — dose individual encaixa na filosofia anti-desperdício."},
  {"n":"Copenhagen Coffee Lab","t":"Café","m":"R. Nova da Piedade 10, Lisboa","tel":"—","email":"hello@copenhagencoffeelab.com","p":"Média","tCliente":"Café de especialidade","gancho":"Pairing pão de ló + café de especialidade."},
  {"n":"Mercearia do Bairro","t":"Mercearia Gourmet","m":"R. do Açúcar 83, Lisboa","tel":"—","email":"geral@merceariadobairro.pt","p":"Média","tCliente":"Mercearia gourmet","gancho":"Produto artesanal com 40 anos — diferenciador face ao industrial."},
  {"n":"Clube de Jornalistas","t":"Restaurante","m":"R. das Trinas 129, Lisboa","tel":"213 977 138","email":"geral@clubedejornalistas.com","p":"Média","tCliente":"Restaurante clássico","gancho":"Clientela fiel — sobremesa clássica como âncora de carta."},
  {"n":"Tasca do Lagarto","t":"Restaurante","m":"R. dos Bacalhoeiros 34, Lisboa","tel":"—","email":"info@tascadolagarto.pt","p":"Média","tCliente":"Tasca / turismo","gancho":"Alfama — turistas com apetência por sobremesas genuinamente portuguesas."},
  {"n":"Café de São Bento","t":"Café","m":"R. de São Bento 212, Lisboa","tel":"213 952 911","email":"geral@cafesaobento.pt","p":"Média","tCliente":"Café clássico","gancho":"Clientela estabelecida — pão de ló como sobremesa de balcão premium."},
  {"n":"O Pitéu da Graça","t":"Restaurante","m":"Pr. da Graça 96, Lisboa","tel":"218 870 565","email":"—","p":"Baixa","tCliente":"Restaurante de bairro","gancho":"Dose individual congelada controla custo e elimina desperdício."},
  {"n":"Mini Bar Teatro","t":"Restaurante","m":"R. António Maria Cardoso 58, Lisboa","tel":"211 305 393","email":"info@minibar.pt","p":"Média","tCliente":"Restaurante criativo","gancho":"Público jovem — pão de ló chocolate ou canela como sobremesa diferenciada."},
],
"Santarém": [
  {"n":"Restaurante O Salazares","t":"Restaurante","m":"R. de São Martinho 2, Santarém","tel":"243 322 384","email":"info@osalazares.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Capital gastronómica regional — receita secular tem aceitação natural."},
  {"n":"Restaurante Portas do Sol","t":"Restaurante","m":"Jardim das Portas do Sol, Santarém","tel":"243 309 520","email":"info@portasdosol.pt","p":"Alta","tCliente":"Restaurante com vista / turismo","gancho":"Turismo de alto valor — sobremesa em volume sem perda de qualidade."},
  {"n":"Hotel Cristal Santarém","t":"Hotel","m":"R. Francisco Moreira 7, Santarém","tel":"243 377 575","email":"reservas@hotelcristal.pt","p":"Alta","tCliente":"Hotel 3* / negócios","gancho":"Carta sólida sem pasteleiro próprio."},
  {"n":"Pastelaria Bijou","t":"Pastelaria","m":"Av. Bernardo Santareno, Santarém","tel":"243 322 507","email":"—","p":"Alta","tCliente":"Pastelaria de referência local","gancho":"Dose individual — impulso ao balcão."},
  {"n":"Pastelaria Paraíso","t":"Pastelaria","m":"Av. Marquês de Sá da Bandeira, Santarém","tel":"243 322 441","email":"—","p":"Média","tCliente":"Pastelaria de bairro","gancho":"Diferenciação com produto artesanal."},
  {"n":"Tasca do Escondidinho","t":"Restaurante","m":"R. Capelo e Ivens, Santarém","tel":"243 323 991","email":"—","p":"Média","tCliente":"Tasca tradicional","gancho":"Pão de ló como sobremesa caseira."},
  {"n":"Mercearia Tradicional do Ribatejo","t":"Mercearia Gourmet","m":"Lg. da Feira, Santarém","tel":"—","email":"—","p":"Média","tCliente":"Mercearia gourmet","gancho":"Identidade ribatejana — produto regional artesanal."},
],
"Linha Sintra–Cascais": [
  {"n":"Bar do Fundo","t":"Restaurante","m":"Av. Alfredo Coelho, Praia Grande, Colares","tel":"219 282 092","email":"info@bardofundo.pt","p":"Alta","tCliente":"Restaurante premium / vista mar","gancho":"Clientela de poder de compra elevado — sobremesa com narrativa."},
  {"n":"Restaurante Azenhas do Mar","t":"Restaurante","m":"Lugar das Piscinas, 2705-098 Colares","tel":"219 280 739","email":"info@azenhasdomar.com","p":"Alta","tCliente":"Restaurante icónico / turismo","gancho":"Mais fotografado de Portugal — sobremesa artesanal reforça o posicionamento."},
  {"n":"Taberna Clandestina","t":"Restaurante","m":"R. Afonso Sanches 36, Cascais","tel":"916 229 630","email":"info@tabernaclandestina.pt","p":"Alta","tCliente":"Gastropub / residentes premium","gancho":"Carta criativa portuguesa — Ti'Piedade como sobremesa de referência."},
  {"n":"Hífen","t":"Restaurante","m":"Av. Dom Carlos I 48, Cascais","tel":"915 546 537","email":"info@hifenrestaurant.com","p":"Alta","tCliente":"Restaurante de referência / Cascais","gancho":"Produto com história diferencia a experiência."},
  {"n":"Hotel Palácio Estoril","t":"Hotel","m":"R. Particular, Estoril","tel":"214 648 000","email":"info@palacioestoril.com","p":"Alta","tCliente":"Hotel 5* / luxo","gancho":"Qualidade constante em grande volume sem pasteleiro."},
  {"n":"Lawrence's Hotel (Sintra)","t":"Hotel","m":"R. Consiglieri Pedroso 38, Sintra","tel":"219 105 500","email":"info@lawrenceshotel.com","p":"Alta","tCliente":"Boutique hotel / turismo cultural","gancho":"Hotel histórico — produto artesanal complementa a narrativa de autenticidade."},
  {"n":"Gourmet Italiano (Cascais)","t":"Mercearia Gourmet","m":"Av. Infante Dom Henrique 1027 D, Cascais","tel":"214 842 127","email":"info@gourmetitaliano.pt","p":"Alta","tCliente":"Deli gourmet / expatriados","gancho":"Clientela internacional com elevado poder de compra."},
  {"n":"Emporium Gourmet","t":"Mercearia Gourmet","m":"Av. Nossa Senhora do Cabo 101, Cascais","tel":"211 541 588","email":"info@emporiumgourmet.pt","p":"Alta","tCliente":"Mercearia gourmet","gancho":"História de 40 anos — receita da D. Piedade vende-se sozinha."},
  {"n":"Taberna Económica de Cascais","t":"Restaurante","m":"R. Sebastião José de Carvalho e Melo 35, Cascais","tel":"214 832 214","email":"info@tabernaeconomicadecascais.com","p":"Alta","tCliente":"Taberna / turismo","gancho":"Volume de turistas — produto congelado garante consistência."},
  {"n":"Angra Gatti","t":"Restaurante","m":"Av. Alfredo Coelho 57, Praia Grande, Colares","tel":"965 770 247","email":"info@angragatti.com","p":"Alta","tCliente":"Restaurante italiano / destino","gancho":"Sobremesa portuguesa como proposta de fecho de refeição."},
  {"n":"Mana","t":"Restaurante","m":"Tv. Navegantes 13, Cascais","tel":"915 669 206","email":"info@manacascais.pt","p":"Média","tCliente":"Restaurante & bar / trendy","gancho":"Chocolate ou canela na carta."},
  {"n":"Café Paris (Sintra)","t":"Café","m":"Pr. da República 32, Sintra","tel":"219 232 375","email":"—","p":"Média","tCliente":"Café turístico","gancho":"Produto icónico para turistas internacionais em Sintra."},
  {"n":"Casa da Galé","t":"Restaurante","m":"Av. Alfredo Coelho 61, Praia Grande, Colares","tel":"219 291 218","email":"—","p":"Média","tCliente":"Restaurante peixe / local","gancho":"Final natural de uma refeição de peixe."},
],
"SuperIndep_Nuno": [
  {"n":"Mercado Municipal de Campo de Ourique","t":"Supermercado Independente","m":"R. Coelho da Rocha, Lisboa","tel":"213 954 628","email":"mercado@cm-lisboa.pt","p":"Alta","tCliente":"Mercado alimentar / produto fresco","gancho":"Clientela premium — dose individual congelada com margem interessante."},
  {"n":"Honest Greens Market (Cascais)","t":"Supermercado Independente","m":"Av. Marginal, São João do Estoril","tel":"—","email":"—","p":"Média","tCliente":"Supermercado independente / saudável","gancho":"Classe média-alta — produto artesanal diferencia o linear."},
],
"Margem Sul": [
  {"n":"O Farol Design Hotel","t":"Hotel","m":"R. do Farol 1, Cacilhas, Almada","tel":"210 407 040","email":"info@farolhotel.com","p":"Alta","tCliente":"Boutique hotel / design","gancho":"Hotel de autor — produto artesanal de alto valor percebido."},
  {"n":"Hotel Sana Sesimbra","t":"Hotel","m":"Av. 25 de Abril, Sesimbra","tel":"212 289 000","email":"info@sesimbra.sanahotels.com","p":"Alta","tCliente":"Resort / lazer","gancho":"F&B de resort — qualidade constante sem pasteleiro."},
  {"n":"Restaurante Ribamar (Sesimbra)","t":"Restaurante","m":"Av. dos Náufragos 29, Sesimbra","tel":"212 233 853","email":"info@restauranteribamar.com","p":"Alta","tCliente":"Restaurante peixe / turismo","gancho":"Destino de verão — produto pronto a servir em período de pico."},
  {"n":"Tasca D'Avenida","t":"Restaurante","m":"Av. Dom Afonso Henriques 10C, Almada","tel":"968 348 036","email":"—","p":"Alta","tCliente":"Tasca contemporânea","gancho":"Almoços de negócios — sobremesa premium com margem confortável."},
  {"n":"The Baptist","t":"Restaurante","m":"R. Afonso Galo 56, Almada","tel":"212 750 996","email":"—","p":"Média","tCliente":"Restaurante casual","gancho":"Volume e rotatividade — sobremesa pronta a servir."},
  {"n":"Restaurante Paladar (Palmela)","t":"Restaurante","m":"R. João de Deus, Palmela","tel":"—","email":"—","p":"Média","tCliente":"Restaurante regional","gancho":"Turismo de interior com apetência por produto artesanal."},
  {"n":"Pastelaria A Floresta (Barreiro)","t":"Pastelaria","m":"Av. Bento Gonçalves, Barreiro","tel":"—","email":"—","p":"Média","tCliente":"Pastelaria de bairro","gancho":"Dose individual ao balcão com margem interessante."},
],
"SuperIndep_Joao": [
  {"n":"Supermercado Apolónia (Leiria)","t":"Supermercado Independente","m":"Av. Heróis de Angola, Leiria","tel":"244 859 900","email":"leiria@apolonia.pt","p":"Alta","tCliente":"Supermercado premium independente","gancho":"Foco em produto nacional — Ti'Piedade é referência perfeita."},
  {"n":"Mini Mercado Tradicional da Caparica","t":"Supermercado Independente","m":"R. da Liberdade, Costa da Caparica","tel":"—","email":"—","p":"Média","tCliente":"Mini mercado / bairro de praia","gancho":"Zona balnear — produto congelado no linear de sobremesas."},
],
"Costa Oeste (S. Martinho–Vieira)": [
  {"n":"Hotel Columbano (S. Martinho do Porto)","t":"Hotel","m":"Av. Marginal, S. Martinho do Porto","tel":"262 989 220","email":"info@hotelcolumbano.com","p":"Alta","tCliente":"Hotel 4* / praia","gancho":"Procura de verão intensa — produto congelado mantém qualidade."},
  {"n":"Tasca do Zé (Nazaré)","t":"Restaurante","m":"R. Mouzinho de Albuquerque 22, Nazaré","tel":"262 551 945","email":"—","p":"Alta","tCliente":"Tasca turística / Nazaré","gancho":"Turismo internacional — pão de ló como sobremesa típica portuguesa."},
  {"n":"Restaurante O Casalinho (Nazaré)","t":"Restaurante","m":"R. do Elevador 24, Nazaré","tel":"262 552 608","email":"—","p":"Alta","tCliente":"Restaurante peixe / turismo","gancho":"Turistas internacionais — produto de identidade nacional."},
  {"n":"Hotel Maré (Nazaré)","t":"Hotel","m":"R. Mouzinho de Albuquerque 8, Nazaré","tel":"262 550 000","email":"info@hotelmare.com","p":"Alta","tCliente":"Hotel / turismo surf","gancho":"Público de surf que valoriza produto artesanal local."},
  {"n":"Solar de Alcobaça","t":"Restaurante","m":"Pr. 25 de Abril, Alcobaça","tel":"262 598 312","email":"—","p":"Média","tCliente":"Restaurante turístico / mosteiro","gancho":"Destino patrimonial — produto com raízes medievais."},
  {"n":"Restaurante A Tasquinha (Alcobaça)","t":"Restaurante","m":"R. Frei António Brandão 2, Alcobaça","tel":"262 582 397","email":"—","p":"Média","tCliente":"Restaurante local","gancho":"Produto artesanal como âncora da carta de sobremesas."},
],
"Leiria": [
  {"n":"Tromba Rija","t":"Restaurante","m":"R. Professores Portelas, Marrazes, Leiria","tel":"244 855 072","email":"geral@trombarja.com","p":"Alta","tCliente":"Restaurante regional / referência","gancho":"Referência gastronómica — produto artesanal encaixa na valorização do território."},
  {"n":"Hotel Eurosol Leiria","t":"Hotel","m":"R. Comissão da Iniciativa, Leiria","tel":"244 838 201","email":"info@eurosolleiria.pt","p":"Alta","tCliente":"Hotel 4* / negócios","gancho":"Carta consistente — produto congelado resolve sobremesas sem pastelaria."},
  {"n":"Pastelaria Garrett","t":"Pastelaria","m":"Pr. Rodrigues Lobo, Leiria","tel":"244 812 370","email":"—","p":"Alta","tCliente":"Pastelaria de referência","gancho":"Complementa o catálogo e diferencia da concorrência."},
  {"n":"Tasca da Barrosinha","t":"Restaurante","m":"Pr. Rodrigues Lobo 2, Leiria","tel":"244 823 703","email":"—","p":"Alta","tCliente":"Tasca tradicional","gancho":"Clientela local — sobremesa clássica com margem confortável."},
  {"n":"O Funil (Leiria)","t":"Restaurante","m":"Av. Heróis de Angola 66, Leiria","tel":"244 832 522","email":"—","p":"Média","tCliente":"Restaurante casual","gancho":"Volume de almoços — produto pronto acelera rotatividade."},
  {"n":"Patisserie Almonda (Tomar)","t":"Pastelaria","m":"Av. Marquês de Tomar, Tomar","tel":"249 312 252","email":"—","p":"Média","tCliente":"Pastelaria de destino","gancho":"Destino turístico — produto português de referência."},
],
"Ericeira–Caldas da Rainha": [
  {"n":"Marginal (Peniche)","t":"Restaurante","m":"Estr. Marginal Norte, Peniche","tel":"968 907 248","email":"marginalrestaurante@gmail.com","p":"Alta","tCliente":"Restaurante premium / costa","gancho":"Vista mar — sobremesa artesanal fecha a experiência."},
  {"n":"Restaurante Sueste (Ericeira)","t":"Restaurante","m":"R. Eduardo Burnay 22, Ericeira","tel":"261 862 108","email":"info@sueste.pt","p":"Alta","tCliente":"Restaurante peixe / turismo surf","gancho":"Comunidade surf internacional — produto artesanal autêntico."},
  {"n":"Hotel Termas das Caldas","t":"Hotel","m":"Pr. 25 de Abril, Caldas da Rainha","tel":"262 830 200","email":"info@termascaldas.pt","p":"Alta","tCliente":"Hotel termal / saúde","gancho":"Produto sem conservantes, ingredientes simples e receita secular."},
  {"n":"Adega do Caseiro (Caldas)","t":"Restaurante","m":"R. Eng. Duarte Pacheco, Caldas da Rainha","tel":"262 831 291","email":"—","p":"Alta","tCliente":"Restaurante regional","gancho":"Referência local — identidade regional reforça o posicionamento."},
  {"n":"Chico Neto (Ribamar)","t":"Restaurante","m":"R. das Armaçõe 26, Ribamar","tel":"261 422 106","email":"—","p":"Alta","tCliente":"Restaurante peixe / local","gancho":"Clientela de fim-de-semana — sobremesa clássica para famílias."},
  {"n":"O Viveiro (Ribamar)","t":"Restaurante","m":"R. das Armaçõe 7, Ribamar","tel":"261 422 197","email":"—","p":"Alta","tCliente":"Restaurante peixe / vista mar","gancho":"Vista mar — produto artesanal eleva a carta sem complexidade."},
  {"n":"Cafetaria Puro Cake Lab","t":"Pastelaria","m":"Pr. Jacob Rodrigues Pereira 18, Peniche","tel":"916 950 480","email":"info@purocakelab.pt","p":"Alta","tCliente":"Pastelaria artesanal","gancho":"Valoriza produto artesanal — Ti'Piedade como oferta complementar."},
  {"n":"Pastelaria Princesa do Mar","t":"Pastelaria","m":"R. António Maria Oliveira 34, Peniche","tel":"262 782 929","email":"—","p":"Alta","tCliente":"Pastelaria local","gancho":"Dose individual com margem de revenda interessante."},
  {"n":"Restaurante Ó Baleal","t":"Restaurante","m":"Baleal, Peniche","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante surf / natureza","gancho":"Baleal ícone do surf — produto artesanal autêntico."},
],
"SuperIndep_Oscar": [
  {"n":"Supermercado Apolónia (Porto)","t":"Supermercado Independente","m":"R. de Júlio Dinis 826, Porto","tel":"226 066 730","email":"porto@apolonia.pt","p":"Alta","tCliente":"Supermercado premium independente","gancho":"Foco em produto nacional — Ti'Piedade é referência natural."},
  {"n":"Mercado Bom Sucesso (Porto)","t":"Supermercado Independente","m":"Pr. do Bom Sucesso 74, Porto","tel":"226 088 800","email":"info@mercadobomsucesso.com","p":"Alta","tCliente":"Mercado gourmet / turismo","gancho":"Mercado de referência no Porto — produto artesanal com 40 anos encaixa naturalmente."},
],
"Coimbra": [
  {"n":"Fangas Mercearia Bar","t":"Mercearia Gourmet","m":"R. Fernandes Tomás 45, Coimbra","tel":"239 115 540","email":"info@fangas.pt","p":"Alta","tCliente":"Mercearia gourmet / bar","gancho":"Curadoria nacional — pão de ló com 40 anos é produto natural."},
  {"n":"Hotel Quinta das Lágrimas","t":"Hotel","m":"Santa Clara, Coimbra","tel":"239 802 380","email":"reservas@quintadaslagrimas.pt","p":"Alta","tCliente":"Hotel 5* / romance","gancho":"Narrativa histórica — produto artesanal reforça experiência portuguesa."},
  {"n":"Restaurante O Trovador","t":"Restaurante","m":"Lg. da Sé Velha 15, Coimbra","tel":"239 825 475","email":"info@otrovador.pt","p":"Alta","tCliente":"Restaurante histórico / fado","gancho":"Fado ao vivo — pão de ló como sobremesa de fim de noite."},
  {"n":"Café Santa Cruz","t":"Café","m":"Pr. 8 de Maio, Coimbra","tel":"239 833 617","email":"cafesantacruz@sapo.pt","p":"Alta","tCliente":"Café histórico / turismo","gancho":"Mais emblemático de Coimbra — produto artesanal com história."},
  {"n":"Pastelaria Briosa","t":"Pastelaria","m":"R. Direita, Coimbra","tel":"239 824 764","email":"—","p":"Alta","tCliente":"Pastelaria de referência","gancho":"Complementa o catálogo sem competir diretamente."},
  {"n":"Adega Paço do Conde","t":"Restaurante","m":"R. Paço do Conde 1, Coimbra","tel":"239 825 605","email":"—","p":"Média","tCliente":"Restaurante clássico","gancho":"Clientela académica — sobremesa clássica universitária."},
  {"n":"Tasca da Rua Nova","t":"Restaurante","m":"R. Nova 44, Coimbra","tel":"239 826 669","email":"—","p":"Média","tCliente":"Tasca contemporânea","gancho":"Carta curta — produto artesanal como âncora."},
],
"Porto": [
  {"n":"O Gaveto","t":"Restaurante","m":"R. Roberto Ivens 826, Matosinhos","tel":"229 381 879","email":"info@restauranteogaveto.com","p":"Alta","tCliente":"Restaurante peixe / referência","gancho":"Clientela exigente — pão de ló é a sobremesa natural de uma refeição portuguesa."},
  {"n":"Mercearia das Flores","t":"Mercearia Gourmet","m":"R. das Flores 110, Porto","tel":"222 013 290","email":"info@merceariadas flores.pt","p":"Alta","tCliente":"Mercearia gourmet / design","gancho":"Curadoria nacional — história de 40 anos e receita intacta."},
  {"n":"Hotel Infante de Sagres","t":"Hotel","m":"Pr. Filipa de Lencastre 62, Porto","tel":"223 398 500","email":"info@hotelinfantesagres.pt","p":"Alta","tCliente":"Hotel 5* histórico","gancho":"Hotel histórico — produto artesanal complementa experiência premium."},
  {"n":"Café Majestic","t":"Café","m":"R. de Santa Catarina 112, Porto","tel":"222 003 887","email":"geral@cafemajestic.com","p":"Alta","tCliente":"Café histórico / turismo","gancho":"Dos cafés mais visitados da Europa — doçaria nacional premium."},
  {"n":"Taberninha do Manel","t":"Restaurante","m":"Av. Gustavo Eiffel 274, Porto","tel":"222 086 389","email":"—","p":"Alta","tCliente":"Tasca histórica / turismo","gancho":"Pão de ló como âncora de cozinha portuguesa."},
  {"n":"Pastelaria Luca","t":"Pastelaria","m":"R. de Sá da Bandeira 118, Porto","tel":"222 084 010","email":"—","p":"Alta","tCliente":"Pastelaria de referência","gancho":"Produto artesanal de autor complementa o catálogo."},
  {"n":"Casa de Pasto da Palmeira","t":"Restaurante","m":"R. de Palmeira 2, Porto","tel":"222 005 753","email":"—","p":"Alta","tCliente":"Casa de pasto / turismo","gancho":"Cozinha portuguesa simples — pão de ló é a sobremesa perfeita."},
  {"n":"Aduela","t":"Restaurante","m":"R. do Oliveiras 38, Porto","tel":"222 008 757","email":"—","p":"Média","tCliente":"Restaurante casual / wine bar","gancho":"Jovem e urbano — chocolate ou canela como sobremesa diferenciada."},
],
"Braga": [
  {"n":"Bem Me Quer (Braga)","t":"Restaurante","m":"Pr. do Município, Braga","tel":"253 278 916","email":"geral@restaurantebemmequeer.pt","p":"Alta","tCliente":"Restaurante de referência","gancho":"Âncora premium da carta de sobremesas."},
  {"n":"Hotel Meliá Braga","t":"Hotel","m":"Av. General Carrilho da Silva Pinto 8, Braga","tel":"253 144 000","email":"melia.braga@melia.com","p":"Alta","tCliente":"Hotel 5* / congressos","gancho":"Volume de eventos — dose individual para grande escala com qualidade consistente."},
  {"n":"Restaurante Inácio","t":"Restaurante","m":"Campo das Hortas 4, Braga","tel":"253 613 235","email":"—","p":"Alta","tCliente":"Restaurante clássico","gancho":"Referência local — produto artesanal reforça qualidade."},
  {"n":"Pastelaria Riquexó","t":"Pastelaria","m":"Av. Central 69, Braga","tel":"253 215 055","email":"—","p":"Alta","tCliente":"Pastelaria clássica","gancho":"Complementa o catálogo sem concorrer diretamente."},
  {"n":"Pastelaria Oliveira","t":"Pastelaria","m":"R. do Souto 128, Braga","tel":"253 215 990","email":"—","p":"Alta","tCliente":"Pastelaria de referência","gancho":"Muito frequentada — diferenciação com margem elevada."},
  {"n":"Taberna Belga","t":"Restaurante","m":"R. de Maximinos 121, Braga","tel":"253 204 786","email":"—","p":"Média","tCliente":"Gastropub / cerveja artesanal","gancho":"Sobremesa artesanal portuguesa como fecho diferenciado."},
],
"Guimarães": [
  {"n":"Solar do Arco","t":"Restaurante","m":"R. de Santa Maria 48, Guimarães","tel":"253 513 072","email":"info@solardoarco.pt","p":"Alta","tCliente":"Restaurante histórico","gancho":"Centro histórico Património Mundial — receita secular encaixa perfeitamente."},
  {"n":"Pousada de Guimarães","t":"Hotel","m":"R. Conde de Margaride 153, Guimarães","tel":"253 511 249","email":"pousadaguimaraes@pousadas.pt","p":"Alta","tCliente":"Pousada histórica / turismo","gancho":"Mosteiro medieval — receita levada ao Japão no séc. XVI."},
  {"n":"El Rei","t":"Restaurante","m":"Pr. de São Tiago 20, Guimarães","tel":"253 419 096","email":"—","p":"Alta","tCliente":"Restaurante centro histórico","gancho":"Clientela turística internacional com apetência por produto típico."},
  {"n":"Pastelaria Clarinha","t":"Pastelaria","m":"R. de Santo António, Guimarães","tel":"253 512 552","email":"—","p":"Alta","tCliente":"Pastelaria de referência","gancho":"Complementa o catálogo com margem interessante."},
  {"n":"Mercearia Vimaranes","t":"Mercearia Gourmet","m":"Lg. do Toural, Guimarães","tel":"—","email":"—","p":"Média","tCliente":"Mercearia gourmet","gancho":"Identidade nacional forte para mercearia gourmet."},
  {"n":"Sabores do Minho","t":"Restaurante","m":"R. de Couros 24, Guimarães","tel":"—","email":"—","p":"Média","tCliente":"Restaurante regional","gancho":"Pão de ló como sobremesa de eleição em carta regional."},
],
}

COMERCIAIS = {
    "nuno":  {"nome":"Nuno",  "email":os.environ.get("EMAIL_NUNO",""),  "canal":"horeca", "zonas":["Lisboa","Santarém","Linha Sintra–Cascais","SuperIndep_Nuno"]},
    "joao":  {"nome":"João",  "email":os.environ.get("EMAIL_JOAO",""),  "canal":"horeca", "zonas":["Lisboa","Margem Sul","Costa Oeste (S. Martinho–Vieira)","Leiria","SuperIndep_Joao"]},
    "oscar": {"nome":"Óscar", "email":os.environ.get("EMAIL_OSCAR",""), "canal":"horeca", "zonas":["Ericeira–Caldas da Rainha","Coimbra","Porto","Braga","Guimarães","SuperIndep_Oscar"]},
    "rui_catering":     {"nome":"Rui", "email":EMAIL_RUI, "canal":"catering",     "zonas":[]},
    "rui_distribuidores":{"nome":"Rui","email":EMAIL_RUI, "canal":"distribuidores","zonas":[]},
}

TIPOS_HORECA = ["Restaurante","Pastelaria","Hotel","Mercearia Gourmet","Café","Supermercado Independente"]

# ════════════════════════════════════════════════════════════════
# UTILITÁRIOS
# ════════════════════════════════════════════════════════════════

def semana_num():
    return datetime.date.today().isocalendar()[1]

def semana_datas():
    d = datetime.date.today()
    seg = d - datetime.timedelta(days=d.weekday())
    sex = seg + datetime.timedelta(days=4)
    meses = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    return f"{seg.day} {meses[seg.month-1]} – {sex.day} {meses[sex.month-1]} {sex.year}"

def hoje():
    return datetime.date.today().strftime("%Y-%m-%d")

def seeded_shuffle(lst, seed):
    result = list(lst)
    for i in range(len(result)-1, 0, -1):
        j = int(abs(math.sin(seed*(i+1)*9301+49297)*233280)) % (i+1)
        result[i], result[j] = result[j], result[i]
    return result

def gerar_leads_horeca(com_id, sem):
    """Gera 20 leads para o comercial, usando a base de dados real dos ficheiros Excel."""
    # Usar base real se disponível, senão fallback para DB_HORECA
    pool_real = DB_HORECA_REAL.get(com_id, [])
    pool_fallback = []
    if not pool_real:
        com = COMERCIAIS[com_id]
        for zona in com["zonas"]:
            for lead in DB_HORECA.get(zona, []):
                if lead["t"] in TIPOS_HORECA:
                    pool_fallback.append({**lead, "zona": zona, "canal":"HORECA"})
        pool = pool_fallback
    else:
        pool = [{**l, "canal":"HORECA"} for l in pool_real]
    
    shuffled = seeded_shuffle(pool, sem*1000 + list(COMERCIAIS.keys()).index(com_id)+1)
    alta  = [l for l in shuffled if l["p"]=="Alta"]
    resto = [l for l in shuffled if l["p"]!="Alta"]
    return (alta+resto)[:20]

def gerar_leads_canal(db, canal_id, sem, n=5):
    shuffled = seeded_shuffle(db, sem*2000 + hash(canal_id)%1000)
    return [{**l, "canal": canal_id} for l in shuffled[:n]]

def lead_key(canal_id, lead):
    return f"{canal_id}::{lead['n']}::{lead.get('zona','PT')}"

# ════════════════════════════════════════════════════════════════
# HISTÓRICO — GitHub API
# ════════════════════════════════════════════════════════════════

def github_api(method, path, data=None):
    token = os.environ.get("GH_PAT","")
    if not token:
        print("[AVISO] GH_PAT não configurado — histórico não será guardado.")
        return None
    url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/{path}"
    headers = {"Authorization":f"Bearer {token}","Accept":"application/vnd.github+json","X-GitHub-Api-Version":"2022-11-28","Content-Type":"application/json"}
    body = json.dumps(data).encode() if data else None
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read())
    except urllib.error.HTTPError as e:
        print(f"[GitHub API] {method} {path} → {e.code}: {e.read().decode()}")
        return None

def ler_historico():
    resp = github_api("GET", f"contents/{HIST_FILE}")
    if not resp:
        return {"leads":{},"ultima_atualizacao":None,"versao":"2.0"}, None
    content = base64.b64decode(resp["content"]).decode("utf-8")
    return json.loads(content), resp["sha"]

def escrever_historico(historico, sha):
    content_b64 = base64.b64encode(json.dumps(historico, ensure_ascii=False, indent=2).encode()).decode()
    data = {"message":f"histórico: {hoje()}","content":content_b64}
    if sha: data["sha"] = sha
    result = github_api("PUT", f"contents/{HIST_FILE}", data)
    print(f"{'✓' if result else '✗'} Histórico {'guardado' if result else 'ERRO ao guardar'}")

def registar_envios(historico, canal_id, leads, email_num, comercial_nome):
    d = hoje()
    for lead in leads:
        k = lead_key(canal_id, lead)
        if k not in historico["leads"]:
            historico["leads"][k] = {
                "nome": lead["n"], "tipo": lead["t"],
                "zona": lead.get("zona","Portugal"), "morada": lead.get("m",""),
                "email_lead": lead.get("email",""), "tel": lead.get("tel",""),
                "comercial": comercial_nome, "canal_id": canal_id,
                "canal": lead.get("canal",""), "prioridade": lead.get("p",""),
                "tipologia_cliente": lead.get("tCliente",""),
                "gancho": lead.get("gancho",""),
                "semana_entrada": semana_num(),
                "data_entrada": d,
                "estado": "Em aberto",
                "emails_enviados": [],
                "visitas": [], "notas": ""
            }
        entry = historico["leads"][k]
        if not any(e["num"]==email_num and e["data"]==d for e in entry["emails_enviados"]):
            entry["emails_enviados"].append({"num":email_num,"data":d,"assunto":f"Nurturing Email {email_num}"})
    historico["ultima_atualizacao"] = d
    return historico

def calcular_email_num(historico, canal_id, lead):
    k = lead_key(canal_id, lead)
    if k not in historico["leads"]: return 1
    enviados = [e["num"] for e in historico["leads"][k].get("emails_enviados",[])]
    for n in [1,2,3,4]:
        if n not in enviados: return n
    return None

# ════════════════════════════════════════════════════════════════
# EXCEL
# ════════════════════════════════════════════════════════════════

C={"hdr":"5C2D0E","ouro":"C49A3C","zebra":"FDF6EF","borda":"E8D5B8","br":"FFFFFF",
   "e1":"D4EDDA","e2":"CCE5FF","e3":"FFF3CD","e4":"F8D7DA","done":"E2E3E5"}

def fill(c): return PatternFill("solid",fgColor=c)
def brd():
    s=Side(style="thin",color=C["borda"])
    return Border(left=s,right=s,top=s,bottom=s)
def ctr(): return Alignment(horizontal="center",vertical="center",wrap_text=True)
def esq(): return Alignment(horizontal="left",vertical="top",wrap_text=True)

def hdr_cell(ws, row, col, val):
    c=ws.cell(row=row,column=col,value=val)
    c.font=Font(name="Arial",bold=True,size=9,color=C["ouro"])
    c.fill=fill(C["hdr"]); c.alignment=ctr(); c.border=brd()
    return c

def aba_leads(wb, nome_aba, leads, canal_id, historico):
    ws=wb.create_sheet(nome_aba)
    ws.merge_cells("A1:Q1")
    ws["A1"]=f"PAO DE LÓ TI'PIEDADE — {nome_aba} — Semana {semana_num()}"
    ws["A1"].font=Font(name="Arial",bold=True,size=13,color=C["br"])
    ws["A1"].fill=fill(C["hdr"]); ws["A1"].alignment=ctr(); ws.row_dimensions[1].height=30
    ws.merge_cells("A2:Q2")
    ws["A2"]=f"{semana_datas()} · Pão de Ló Ti'Piedade Unidose 85g"
    ws["A2"].font=Font(name="Arial",size=9,italic=True,color="6B5744")
    ws["A2"].alignment=ctr(); ws.row_dimensions[2].height=16

    hdrs=["#","Espaço / Empresa","Tipo","Tipologia","Zona","Morada","Tel","Email","Gancho","Prio","E1","E2","E3","E4","Próx.Email","Estado","Observações"]
    for col,h in enumerate(hdrs,1): hdr_cell(ws,3,col,h)
    ws.row_dimensions[3].height=26

    for i,lead in enumerate(leads):
        row=4+i; k=lead_key(canal_id,lead)
        hist=historico["leads"].get(k,{})
        nums={e["num"]:e["data"] for e in hist.get("emails_enviados",[])}
        proximo=calcular_email_num(historico,canal_id,lead)
        cor=C["zebra"] if i%2==0 else C["br"]
        base=[i+1,lead["n"],lead["t"],lead.get("tCliente",""),lead.get("zona",""),lead.get("m",""),lead.get("tel",""),lead.get("email",""),lead.get("gancho",""),lead.get("p","")]
        for col,v in enumerate(base,1):
            c=ws.cell(row=row,column=col,value=v)
            c.fill=fill(cor); c.border=brd()
            c.font=Font(name="Arial",size=9,bold=(col==2))
            c.alignment=ctr() if col in [1,3,7,10] else esq()
        cores_e={1:C["e1"],2:C["e2"],3:C["e3"],4:C["e4"]}
        for n in [1,2,3,4]:
            col=10+n; data_e=nums.get(n,"")
            c=ws.cell(row=row,column=col,value=data_e if data_e else "—")
            c.fill=fill(cores_e[n] if data_e else cor); c.border=brd()
            c.font=Font(name="Arial",size=9,bold=bool(data_e)); c.alignment=ctr()
        estado=hist.get("estado","Em aberto")
        c=ws.cell(row=row,column=15,value=f"Email {proximo}" if proximo else "✓ Pronto p/ visita")
        c.fill=fill(C["e3"] if proximo else C["e1"]); c.border=brd()
        c.font=Font(name="Arial",size=9,bold=True); c.alignment=ctr()
        c=ws.cell(row=row,column=16,value=estado)
        c.fill=fill(cor); c.border=brd(); c.font=Font(name="Arial",size=9); c.alignment=ctr()
        c=ws.cell(row=row,column=17,value=hist.get("notas",""))
        c.fill=fill(cor); c.border=brd(); c.font=Font(name="Arial",size=9); c.alignment=esq()
        ws.row_dimensions[row].height=34

    for col,w in enumerate([4,28,16,22,18,32,13,26,38,8,12,12,12,12,16,16,28],1):
        ws.column_dimensions[get_column_letter(col)].width=w
    ws.freeze_panes="K4"; ws.auto_filter.ref=f"A3:Q{3+len(leads)}"

def aba_historico(wb, historico):
    ws=wb.create_sheet("📊 Histórico",0)
    ws.merge_cells("A1:K1")
    ws["A1"]=f"TI'PIEDADE — Histórico Completo de Prospeção (atualizado {hoje()})"
    ws["A1"].font=Font(name="Arial",bold=True,size=13,color=C["br"])
    ws["A1"].fill=fill(C["hdr"]); ws["A1"].alignment=ctr(); ws.row_dimensions[1].height=28
    hdrs=["Comercial","Canal","Lead","Tipo","Zona","Email 1","Email 2","Email 3","Email 4","Estado","Notas"]
    for col,h in enumerate(hdrs,1): hdr_cell(ws,2,col,h)
    ws.row_dimensions[2].height=24
    row_h=3
    for k,v in sorted(historico["leads"].items(),key=lambda x:(x[1].get("comercial",""),x[1].get("canal",""))):
        nums={e["num"]:e["data"] for e in v.get("emails_enviados",[])}
        completo=len(nums)>=4
        cor=C["e1"] if completo else C["zebra"] if row_h%2==0 else C["br"]
        vals=[v.get("comercial",""),v.get("canal",""),v.get("nome",""),v.get("tipo",""),v.get("zona",""),
              nums.get(1,"—"),nums.get(2,"—"),nums.get(3,"—"),nums.get(4,"—"),
              v.get("estado","Em aberto"),v.get("notas","")]
        for col,val in enumerate(vals,1):
            c=ws.cell(row=row_h,column=col,value=val)
            c.fill=fill(cor); c.border=brd()
            c.font=Font(name="Arial",size=9,bold=(col==3)); c.alignment=ctr() if col in [6,7,8,9,10] else esq()
        ws.row_dimensions[row_h].height=20; row_h+=1
    for col,w in enumerate([14,16,28,16,18,13,13,13,13,16,30],1):
        ws.column_dimensions[get_column_letter(col)].width=w
    ws.freeze_panes="A3"; ws.auto_filter.ref=f"A2:K{row_h-1}"

def criar_excel(sem, historico):
    wb=openpyxl.Workbook(); wb.remove(wb.active)
    aba_historico(wb,historico)

    # Abas HORECA por comercial
    for com_id in ["nuno","joao","oscar"]:
        com=COMERCIAIS[com_id]
        leads=gerar_leads_horeca(com_id,sem)
        aba_leads(wb,f"HORECA — {com['nome']}",leads,com_id,historico)

    # Aba Catering & Eventos (Rui)
    leads_cat=gerar_leads_canal(DB_CATERING,"rui_catering",sem,5)
    aba_leads(wb,"Catering & Eventos — Rui",leads_cat,"rui_catering",historico)

    # Aba Distribuidores (Rui)
    leads_dist=gerar_leads_canal(DB_DISTRIBUIDORES,"rui_distribuidores",sem,5)
    aba_leads(wb,"Distribuidores — Rui",leads_dist,"rui_distribuidores",historico)

    fname=f"TiPiedade_Leads_S{sem}_{datetime.date.today().year}.xlsx"
    wb.save(fname); return fname

# ════════════════════════════════════════════════════════════════
# NURTURING — CONTEÚDO PERSONALIZADO POR TIPOLOGIA
# ════════════════════════════════════════════════════════════════

def tipologia_grupo(tipo):
    """Mapeia tipologia do lead para grupo de email."""
    t = (tipo or "").lower()
    if any(x in t for x in ["hotel","resort","pousada"]):            return "hotel"
    if any(x in t for x in ["pastelaria","padaria","confeitaria"]):  return "pastelaria"
    if any(x in t for x in ["catering","evento","banquete"]):        return "catering"
    if any(x in t for x in ["distribuidor","congelado","frigori"]):  return "distribuidor"
    if any(x in t for x in ["mercearia","gourmet","deli","garrafeira"]): return "gourmet"
    if any(x in t for x in ["café","coffee","brunch","chá"]):        return "cafe"
    if any(x in t for x in ["supermercado","mercado"]):              return "super"
    if any(x in t for x in ["cervejaria"]):                          return "cervejaria"
    return "restaurante"  # default

# ── Assuntos por tipologia e número de email ──────────────────
ASSUNTOS = {
    "restaurante": {
        1: "Uma sobremesa portuguesa que se vende sozinha — e sem pasteleiro",
        2: "Como um restaurante da vossa zona eliminou as quebras em sobremesas",
        3: "A tendência que está a mudar as cartas de sobremesa em Portugal",
        4: "Queremos passar com uma amostra — leva 10 minutos",
    },
    "hotel": {
        1: "Sobremesa artesanal portuguesa para o vosso F&B — sem complexidade",
        2: "Como um hotel da região passou a oferecer sobremesas de qualidade sem pasteleiro",
        3: "O que os hóspedes estão a pedir mais nos hotéis portugueses",
        4: "Podemos passar com amostras para a equipa de F&B experimentar",
    },
    "pastelaria": {
        1: "Pão de Ló Ti'Piedade — um artesanal com 40 anos para complementar a vossa vitrina",
        2: "Como outras pastelarias estão a diferenciar-se com o pão de ló Ti'Piedade",
        3: "Produto artesanal, dose individual, margem elevada — faz sentido para vocês?",
        4: "Queremos deixar amostras — sem compromisso, só para provar",
    },
    "catering": {
        1: "Sobremesa individual artesanal para os vossos eventos — sem logística complexa",
        2: "Como empresas de catering estão a usar o pão de ló Ti'Piedade em eventos premium",
        3: "A sobremesa que mais impressiona em casamentos e eventos corporativos",
        4: "Podemos reunir para apresentar o produto e condições para catering",
    },
    "distribuidor": {
        1: "Pão de Ló Ti'Piedade em congelado — referência premium para a vossa carteira",
        2: "Uma marca com 40 anos e procura crescente no canal HORECA",
        3: "Zonas sem cobertura, margem interessante — vale a pena conversar",
        4: "Podemos apresentar as condições comerciais para distribuição regional",
    },
    "gourmet": {
        1: "Pão de Ló Ti'Piedade — 40 anos de receita artesanal para a vossa prateleira",
        2: "Como mercearias gourmet estão a destacar-se com o pão de ló Ti'Piedade",
        3: "Produto com narrativa, origem e história — exatamente o que os vossos clientes procuram",
        4: "Queremos passar para mostrar o produto e proposta comercial",
    },
    "cafe": {
        1: "O pairing perfeito para o vosso café — pão de ló artesanal em dose individual",
        2: "Como cafés de especialidade estão a aumentar o ticket médio com o Ti'Piedade",
        3: "Produto artesanal português com história — ideal para o público que frequenta o vosso espaço",
        4: "Queremos passar com amostras para a vossa equipa provar",
    },
    "super": {
        1: "Pão de Ló Ti'Piedade — sobremesa artesanal congelada para o vosso linear",
        2: "Como supermercados independentes estão a diferenciar-se com produto artesanal",
        3: "Sem grandes grupos, sem intermediários — direto do produtor para o vosso linear",
        4: "Podemos reunir para apresentar condições comerciais e produto",
    },
    "cervejaria": {
        1: "A sobremesa portuguesa que fecha qualquer refeição — Pão de Ló Ti'Piedade",
        2: "Grandes volumes, zero desperdício — como o Ti'Piedade resolve a sobremesa em cervejarias",
        3: "Dose individual congelada: qualidade constante, custo previsível, margem alta",
        4: "Queremos passar para deixar amostras e apresentar condições",
    },
}

# ── Corpos de email por tipologia ─────────────────────────────
def corpo_email(num, grupo, nome_lead, zona, nome_comercial, tel_comercial=""):
    """Gera o corpo do email de nurturing personalizado por tipologia e número."""

    assinatura = f"""Com os melhores cumprimentos,
{nome_comercial}
Equipa Comercial Ti'Piedade
{tel_comercial}
comercial@tipiedade.com | www.tipiedade.com"""

    corpos = {
        "restaurante": {
            1: f"""Exmo(a). Sr(a),

O meu nome é {nome_comercial} e represento o Pão de Ló Ti'Piedade — empresa portuguesa em terceira geração familiar, com mais de 40 anos de história na doçaria artesanal.

Estamos a trabalhar com restaurantes em {zona} e identificámos {nome_lead} como um espaço com o perfil certo para o nosso produto.

O Pão de Ló Ti'Piedade em unidose de 85g (congelado) resolve um problema que muitos restaurantes conhecem bem: ter uma sobremesa de qualidade na carta, sem desperdício e sem precisar de pasteleiro.

✦ Descongela rapidamente — pronto a servir em minutos
✦ Quatro sabores: Original, Chocolate, Canela e Café
✦ Ingredientes simples, sem conservantes, receita com 40 anos
✦ Margem confortável para o operador

Nas próximas semanas partilharemos alguns exemplos de como está a funcionar noutros restaurantes da vossa zona.

{assinatura}""",

            2: f"""Exmo(a). Sr(a),

Na semana passada apresentámos o Pão de Ló Ti'Piedade. Hoje queremos partilhar um caso concreto.

Um restaurante de peixe da vossa zona — com clientela local e de fim-de-semana — introduziu o Ti'Piedade há cerca de dois meses. O que nos disseram:

→ A sobremesa passou a ser a mais pedida, acima do pudim e da mousse
→ Desperdício zero — serve apenas o que precisa, quando precisa
→ O responsável: "É a coisa mais fácil que temos na cozinha. Sai do congelador, vai ao prato."

Em restaurantes de peixe e marisco, o pão de ló é o final natural de uma refeição. Os clientes reconhecem-no e pedem-no.

Se quiser saber mais sobre como funciona na prática, estamos disponíveis para uma conversa rápida.

{assinatura}""",

            3: f"""Exmo(a). Sr(a),

Uma tendência que estamos a observar no setor: os clientes pedem cada vez mais sobremesas com história e origem, mas os operadores não querem complexidade.

O pão de ló é a sobremesa portuguesa mais reconhecida fora de Portugal — e a receita Ti'Piedade foi levada pelos portugueses ao Japão no século XVI, onde existe até hoje. É uma história que se conta em segundos e que os clientes valorizam.

Para {nome_lead}, isso significa uma sobremesa diferenciada na carta, sem investimento em pastelaria, com custo fixo por dose e margem previsível.

Na semana que vem a nossa equipa passa pela vossa zona. Gostaríamos de deixar amostras para a equipa de cozinha experimentar.

{assinatura}""",

            4: f"""Exmo(a). Sr(a),

Ao longo das últimas semanas partilhámos a história do Ti'Piedade e alguns exemplos de como está a funcionar noutros restaurantes.

Esta semana {nome_comercial} vai passar por {zona} e gostaria de parar 10 minutos em {nome_lead} para deixar amostras dos quatro sabores — Original, Chocolate, Canela e Café.

Sem reunião, sem apresentação. Só o produto, para provarem.

Se preferir agendar: {tel_comercial if tel_comercial else "responda a este email"}.

{assinatura}""",
        },

        "hotel": {
            1: f"""Exmo(a). Sr(a),

O meu nome é {nome_comercial} e represento o Pão de Ló Ti'Piedade — empresa portuguesa com mais de 40 anos de história na doçaria artesanal.

Contacto-o(a) porque trabalhamos com vários hotéis na região e identificámos um desafio comum: manter uma carta de sobremesas de qualidade constante, especialmente em períodos de maior ocupação, sem depender de pasteleiro especializado.

O Pão de Ló Ti'Piedade em unidose de 85g (congelado) resolve exatamente isso:

✦ Qualidade artesanal consistente, independentemente do volume
✦ Zero preparação — descongela e serve
✦ Pode ser usado no restaurante, pequeno-almoço, room service ou eventos
✦ Quatro sabores: Original, Chocolate, Canela e Café

Nas próximas semanas partilhamos exemplos de como está a funcionar em unidades hoteleiras semelhantes à vossa.

{assinatura}""",

            2: f"""Exmo(a). Sr(a),

Partilhamos hoje um exemplo real de como o Pão de Ló Ti'Piedade está a ser usado no setor hoteleiro.

Um hotel de 4* na região centro — com restaurante próprio e room service — introduziu o produto há três meses. O responsável de F&B partilhou connosco:

→ A sobremesa passou a fazer parte do menu de room service com grande aceitação
→ A consistência de qualidade eliminou reclamações sobre sobremesas
→ O custo por dose é previsível e a margem é superior à pastelaria fresca

Para hotéis com eventos e grupos, a dose individual em congelado é especialmente eficiente: serve exatamente o que precisa, sem perdas.

{assinatura}""",

            3: f"""Exmo(a). Sr(a),

Os hóspedes internacionais estão cada vez mais atentos à autenticidade dos produtos que consomem — especialmente em contexto de viagem.

O Pão de Ló Ti'Piedade tem uma história que funciona: receita levada pelos portugueses ao Japão no século XVI, mantida intacta há 40 anos, produzida em Portugal por uma família em terceira geração. É o tipo de produto que um hóspede leva na memória.

Para a equipa de F&B de {nome_lead}, isso significa uma sobremesa com valor percebido elevado, sem complexidade operacional.

Na próxima semana podemos passar para uma conversa rápida ou deixar amostras.

{assinatura}""",

            4: f"""Exmo(a). Sr(a),

Nas últimas semanas apresentámos o Ti'Piedade e partilhámos casos reais de utilização em hotéis.

Esta semana gostaríamos de passar por {nome_lead} para deixar amostras dos quatro sabores com a equipa de F&B. Demora menos de 15 minutos e não requer compromisso.

Para agendar: {tel_comercial if tel_comercial else "responda a este email"}.

{assinatura}""",
        },

        "pastelaria": {
            1: f"""Exmo(a). Sr(a),

O meu nome é {nome_comercial} e represento o Pão de Ló Ti'Piedade — empresa portuguesa em terceira geração, com mais de 40 anos a produzir pão de ló artesanal com a receita original da D.ª Piedade.

Contactamos {nome_lead} porque acreditamos que o nosso produto pode ser um complemento interessante ao vosso catálogo — não para competir com o que já têm, mas para adicionar uma referência nacional em formato individual.

O Pão de Ló Ti'Piedade em unidose de 85g (congelado):

✦ Dose individual pronta a vender ao balcão após descongelamento
✦ Quatro sabores: Original, Chocolate, Canela e Café
✦ Receita secular, ingredientes simples, sem conservantes
✦ Margem de revenda competitiva

Nas próximas semanas partilhamos exemplos de como outras pastelarias estão a trabalhar o produto.

{assinatura}""",

            2: f"""Exmo(a). Sr(a),

Uma pastelaria de referência em Lisboa começou a trabalhar o Ti'Piedade há quatro meses. O que nos disseram:

→ O produto tornou-se uma das referências de venda por impulso ao balcão
→ A dose individual permite controlo de stock sem perdas
→ "Os clientes pedem porque reconhecem. Não precisamos de explicar o que é."

O pão de ló Ti'Piedade não concorre com a produção própria — complementa-a. É uma referência nacional que os clientes procuram.

{assinatura}""",

            3: f"""Exmo(a). Sr(a),

O consumidor de pastelaria está a valorizar cada vez mais os produtos com origem, história e receita verificável.

O Ti'Piedade tem exatamente isso: 40 anos de receita da D.ª Piedade, produzido em Portugal, ingredientes simples e sem aditivos industriais. É um produto que se explica em duas frases e que os clientes entendem imediatamente.

Para {nome_lead}, adicionar o Ti'Piedade ao balcão é adicionar uma referência com valor percebido elevado e sem esforço de produção.

Podemos passar na próxima semana com amostras?

{assinatura}""",

            4: f"""Exmo(a). Sr(a),

Esta semana {nome_comercial} passa por {zona} e gostaria de parar em {nome_lead} para deixar amostras dos quatro sabores — para a equipa provar e avaliar se faz sentido para o vosso balcão.

Sem compromisso. Só o produto.

Para agendar: {tel_comercial if tel_comercial else "responda a este email"}.

{assinatura}""",
        },

        "catering": {
            1: f"""Exmo(a). Sr(a),

O meu nome é {nome_comercial} e represento o Pão de Ló Ti'Piedade — empresa portuguesa com mais de 40 anos de história na doçaria artesanal.

Para empresas de catering e organização de eventos, o desafio das sobremesas é sempre o mesmo: qualidade constante, facilidade logística e custo controlado.

O Pão de Ló Ti'Piedade em unidose de 85g (congelado) responde exatamente a isso:

✦ Dose individual — sem corte, sem preparação, sem desperdício
✦ Qualidade artesanal consistente em qualquer volume
✦ Quatro sabores para variar por evento: Original, Chocolate, Canela e Café
✦ Embalagem individual elegante — adequada para serviço em mesa

{assinatura}""",

            2: f"""Exmo(a). Sr(a),

Uma empresa de catering de casamentos em Lisboa começou a usar o Ti'Piedade há seis meses. O feedback dos clientes foi imediato:

→ O pão de ló é reconhecido como "sobremesa portuguesa de sempre" — funciona como momento de identidade no menu
→ Em eventos com 200+ convidados, a dose individual eliminou toda a logística de corte e emplatamento
→ "É o produto mais fácil de gerir num evento. Zero perdas, zero improvisação."

Para casamentos e eventos corporativos, o Ti'Piedade posiciona-se como a sobremesa com narrativa — algo que fica na memória dos convidados.

{assinatura}""",

            3: f"""Exmo(a). Sr(a),

A tendência em eventos premium é clara: os convidados querem autenticidade e os organizadores querem simplicidade operacional.

O Pão de Ló Ti'Piedade é a interseção perfeita: produto artesanal com 40 anos de história, dose individual pronta a servir, e uma narrativa que o responsável de sala consegue contar em segundos.

Temos condições comerciais específicas para catering e eventos, com preços por volume. Podemos apresentar?

{assinatura}""",

            4: f"""Exmo(a). Sr(a),

Esta semana {nome_comercial} está disponível para uma reunião com {nome_lead} para apresentar o produto, condições para catering e deixar amostras dos quatro sabores.

Para agendar: {tel_comercial if tel_comercial else "responda a este email"}.

{assinatura}""",
        },

        "distribuidor": {
            1: f"""Exmo(a). Sr(a),

O meu nome é {nome_comercial} e represento o Pão de Ló Ti'Piedade — empresa portuguesa em terceira geração, com mais de 40 anos a produzir doçaria artesanal regional.

Estamos a alargar a nossa rede de distribuição de congelados e identificámos {nome_lead} como um parceiro com cobertura na zona que nos interessa.

O que propomos:

✦ Pão de Ló Ti'Piedade unidose 85g (congelado) — produto de alto valor percebido
✦ Margem de distribuição competitiva
✦ Procura crescente no canal HORECA e retalho gourmet
✦ Marca com 40 anos e reconhecimento nacional

Nas próximas semanas partilhamos mais detalhes sobre o produto e condições.

{assinatura}""",

            2: f"""Exmo(a). Sr(a),

O Pão de Ló Ti'Piedade tem hoje distribuição em mais de 500 pontos de venda em Portugal — mas existem regiões com potencial por desenvolver.

A procura no canal HORECA tem crescido consistentemente: restaurantes, hotéis e pastelarias procuram um produto de doçaria artesanal congelado com qualidade constante e que os clientes reconheçam.

Para um distribuidor com a vossa cobertura regional, o Ti'Piedade é uma referência com diferenciação clara face aos produtos industriais existentes no mercado.

{assinatura}""",

            3: f"""Exmo(a). Sr(a),

As condições que praticamos para distribuidores regionais incluem:

✦ Preços por volume com margens ajustadas à distribuição
✦ Apoio em materiais de comunicação para os pontos de venda
✦ Flexibilidade de encomenda (MOQ negociável por zona)
✦ Produto com validade longa em congelado — sem pressão de rotação

Podemos reunir para apresentar a proposta comercial completa?

{assinatura}""",

            4: f"""Exmo(a). Sr(a),

Esta semana {nome_comercial} está disponível para uma reunião com {nome_lead} para apresentar as condições de distribuição e deixar amostras do produto.

Para agendar: {tel_comercial if tel_comercial else "responda a este email"}.

{assinatura}""",
        },

        "gourmet": {
            1: f"""Exmo(a). Sr(a),

O meu nome é {nome_comercial} e represento o Pão de Ló Ti'Piedade — empresa portuguesa em terceira geração, com mais de 40 anos de história.

A nossa receita — criada pela D.ª Piedade e mantida intacta — é uma das referências da doçaria artesanal portuguesa. Hoje chegamos a mais de 500 pontos de venda em todo o país.

Para espaços como {nome_lead}, o Ti'Piedade é um produto com narrativa forte e valor percebido elevado:

✦ Receita secular levada pelos portugueses ao Japão no século XVI
✦ Produção artesanal, ingredientes simples, sem conservantes
✦ Unidose 85g em congelado — fácil de expor e vender
✦ Quatro sabores: Original, Chocolate, Canela e Café

{assinatura}""",

            2: f"""Exmo(a). Sr(a),

Uma mercearia gourmet em Cascais começou a trabalhar o Ti'Piedade há três meses. O que nos disseram:

→ "Os clientes compram porque reconhecem a marca. Não precisamos de explicar."
→ O produto tornou-se uma das referências de doçaria nacional no espaço
→ A dose individual em congelado facilita a gestão de stock sem perdas

Para espaços gourmet, o Ti'Piedade é exatamente o tipo de produto que os clientes procuram: artesanal, com história, de origem verificável.

{assinatura}""",

            3: f"""Exmo(a). Sr(a),

O consumidor gourmet valoriza três coisas: origem, história e autenticidade. O Ti'Piedade tem as três.

40 anos de receita intacta. Produção familiar em terceira geração. Ingredientes simples que qualquer cliente consegue ler e entender.

Para {nome_lead}, adicionar o Ti'Piedade ao linear é adicionar uma referência com identidade forte e fácil de comunicar.

Podemos passar na próxima semana?

{assinatura}""",

            4: f"""Exmo(a). Sr(a),

Esta semana {nome_comercial} passa por {zona} e gostaria de visitar {nome_lead} para apresentar o produto e proposta comercial.

Para agendar: {tel_comercial if tel_comercial else "responda a este email"}.

{assinatura}""",
        },

        "cafe": {
            1: f"""Exmo(a). Sr(a),

O meu nome é {nome_comercial} e represento o Pão de Ló Ti'Piedade — empresa portuguesa com mais de 40 anos de história na doçaria artesanal.

Para cafés e espaços de brunch, o Ti'Piedade resolve um problema que conhecemos bem: ter uma opção de doçaria de qualidade, fácil de servir, com margem interessante.

O Pão de Ló Ti'Piedade em unidose de 85g (congelado):

✦ Descongela rapidamente — pronto ao balcão em pouco tempo
✦ Pairing natural com café de especialidade
✦ Produto reconhecido — os clientes já conhecem e pedem
✦ Quatro sabores: Original, Chocolate, Canela e Café

{assinatura}""",

            2: f"""Exmo(a). Sr(a),

Um café de especialidade em Lisboa introduziu o Ti'Piedade há dois meses. O responsável partilhou:

→ "O pão de ló com café tornou-se o nosso produto de tarde mais vendido."
→ A dose individual eliminou desperdício — antes tinham bolos inteiros que nem sempre vendiam
→ O ticket médio da tarde subiu com a combinação café + fatia

Para espaços como {nome_lead}, o Ti'Piedade é a opção de doçaria artesanal que encaixa no posicionamento sem exigir produção própria.

{assinatura}""",

            3: f"""Exmo(a). Sr(a),

O público de café de especialidade e brunch é exatamente o público que mais valoriza produto artesanal com história.

O Ti'Piedade tem 40 anos de receita intacta, ingredientes simples e produção familiar — é o tipo de produto que fica bem numa ardósia e que o barista consegue explicar em duas frases.

Podemos passar na próxima semana com amostras?

{assinatura}""",

            4: f"""Exmo(a). Sr(a),

Esta semana {nome_comercial} passa por {zona} e gostaria de parar em {nome_lead} para deixar amostras dos quatro sabores.

Para agendar: {tel_comercial if tel_comercial else "responda a este email"}.

{assinatura}""",
        },

        "super": {
            1: f"""Exmo(a). Sr(a),

O meu nome é {nome_comercial} e represento o Pão de Ló Ti'Piedade — empresa portuguesa com mais de 40 anos de história.

Estamos a alargar a presença em supermercados e mercearias independentes e identificámos {nome_lead} como um parceiro com o perfil certo.

O que propomos:

✦ Pão de Ló Ti'Piedade unidose 85g (congelado) — linear de congelados ou balcão
✦ Produto com reconhecimento nacional e procura crescente
✦ Quatro sabores: Original, Chocolate, Canela e Café
✦ Condições comerciais para retalhistas independentes

{assinatura}""",

            2: f"""Exmo(a). Sr(a),

Supermercados independentes que trabalham o Ti'Piedade reportam consistentemente o mesmo resultado: o produto vende sem esforço de comunicação, porque os clientes já o conhecem.

A dose individual em congelado é especialmente adequada para o retalho independente — sem risco de validade, sem perdas, rotação previsível.

Para {nome_lead}, é uma referência de doçaria artesanal que diferencia o linear face aos produtos industriais da grande distribuição.

{assinatura}""",

            3: f"""Exmo(a). Sr(a),

Trabalhar com fornecedores fora dos grandes grupos tem vantagens claras: preços mais competitivos, relação direta e produto diferenciado.

O Ti'Piedade é exatamente isso — produto artesanal, familiar, com 40 anos, que os consumidores reconhecem e que não encontram nos lineares da grande distribuição.

Podemos reunir para apresentar condições?

{assinatura}""",

            4: f"""Exmo(a). Sr(a),

Esta semana {nome_comercial} passa por {zona} e gostaria de visitar {nome_lead} com amostras e proposta comercial.

Para agendar: {tel_comercial if tel_comercial else "responda a este email"}.

{assinatura}""",
        },

        "cervejaria": {
            1: f"""Exmo(a). Sr(a),

O meu nome é {nome_comercial} e represento o Pão de Ló Ti'Piedade — empresa portuguesa com mais de 40 anos de história na doçaria artesanal.

Para cervejarias com grande volume de refeições, a sobremesa é frequentemente o elemento mais difícil de gerir: custo variável, desperdício, e falta de consistência.

O Pão de Ló Ti'Piedade em unidose de 85g (congelado) resolve tudo isso:

✦ Dose individual — serve só o que precisa, zero perdas
✦ Custo fixo e previsível por sobremesa
✦ Zero preparação — direto do congelador ao prato
✦ Clientes reconhecem e pedem — sem esforço de venda

{assinatura}""",

            2: f"""Exmo(a). Sr(a),

Uma cervejaria de referência em Lisboa introduziu o Ti'Piedade há quatro meses. O responsável de sala:

→ "Acabámos com o problema das sobremesas. O pão de ló sai sempre bem, em qualquer quantidade."
→ O desperdício em sobremesas caiu para zero
→ A margem por sobremesa aumentou face à mousse e pudim que tinham antes

Em cervejarias com volume alto e rotatividade rápida, a dose individual congelada é a solução mais eficiente.

{assinatura}""",

            3: f"""Exmo(a). Sr(a),

A sobremesa numa cervejaria precisa de ser rápida, consistente e com boa margem. O Ti'Piedade é tudo isso.

40 anos de receita portuguesa intacta. Produto que os clientes reconhecem e pedem sem precisar de ler a carta. Dose individual que se serve em segundos.

Para {nome_lead}, isso significa mais receita por mesa com menos complexidade operacional.

Podemos passar na próxima semana?

{assinatura}""",

            4: f"""Exmo(a). Sr(a),

Esta semana {nome_comercial} passa por {zona} e gostaria de parar em {nome_lead} para deixar amostras e mostrar como funciona na prática.

Para agendar: {tel_comercial if tel_comercial else "responda a este email"}.

{assinatura}""",
        },
    }

    grupo_corpos = corpos.get(grupo, corpos["restaurante"])
    return grupo_corpos.get(num, grupo_corpos[1])


def corpo_html(email_num, grupo, nome_lead, zona, nome_comercial, texto_plain):
    """Gera o email HTML com visual Ti'Piedade."""

    # Cor de acento por número de email
    cor_num = {1:"#5C2D0E", 2:"#D4682A", 3:"#C49A3C", 4:"#5C2D0E"}
    etiqueta_num = {1:"Apresentação", 2:"Caso de sucesso", 3:"Tendência", 4:"Convite"}
    icone_grupo = {
        "restaurante":"🍽", "hotel":"🏨", "pastelaria":"🥐",
        "catering":"🎪", "distribuidor":"🚛", "gourmet":"🫙",
        "cafe":"☕", "super":"🛒", "cervejaria":"🍺"
    }

    # Converter texto plain em parágrafos HTML
    paragrafos = ""
    for linha in texto_plain.strip().split("\n"):
        l = linha.strip()
        if not l:
            paragrafos += "<br>"
        elif l.startswith("✦"):
            paragrafos += f'<p style="margin:4px 0;padding-left:16px;color:#5C2D0E"><span style="color:#C49A3C;font-weight:700">✦</span> {l[1:].strip()}</p>'
        elif l.startswith("→"):
            paragrafos += f'<p style="margin:4px 0;padding-left:16px;color:#3D1C07"><span style="color:#D4682A;font-weight:700">→</span> {l[1:].strip()}</p>'
        elif l.startswith("Com os melhores") or l.startswith("Pão de Ló"):
            paragrafos += f'<p style="margin:2px 0;color:#6B5744;font-size:13px">{l}</p>'
        else:
            paragrafos += f'<p style="margin:8px 0;color:#2C1A0A;line-height:1.6">{l}</p>'

    num_badge = cor_num.get(email_num, "#5C2D0E")
    etiq = etiqueta_num.get(email_num, "")
    ico = icone_grupo.get(grupo, "🍽")

    return f"""<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Pão de Ló Ti'Piedade</title>
</head>
<body style="margin:0;padding:0;background:#F8F0E3;font-family:Georgia,serif;">

<table width="100%" cellpadding="0" cellspacing="0" style="background:#F8F0E3;padding:32px 0">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%">

  <!-- HEADER -->
  <tr>
    <td style="background:#5C2D0E;border-radius:14px 14px 0 0;padding:0">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="padding:28px 36px 20px">
            <p style="margin:0;font-family:Georgia,serif;font-size:26px;font-weight:700;color:#ffffff;letter-spacing:.02em">
              Pão de Ló <span style="color:#C49A3C">Ti'Piedade</span>
            </p>
            <p style="margin:6px 0 0;font-size:11px;color:#F2A46A;letter-spacing:.12em;text-transform:uppercase">
              Doçaria Artesanal Portuguesa · Desde 1984
            </p>
          </td>
          <td align="right" style="padding:28px 36px 20px;vertical-align:top">
            <span style="display:inline-block;background:rgba(255,255,255,.12);border:1px solid rgba(196,154,60,.4);border-radius:99px;padding:4px 14px;font-size:11px;font-weight:700;color:#C49A3C;letter-spacing:.08em;text-transform:uppercase">
              {ico} {etiq}
            </span>
            <p style="margin:6px 0 0;font-size:10px;color:rgba(255,255,255,.4);text-align:right">
              Email {email_num} de 4
            </p>
          </td>
        </tr>
        <tr>
          <td colspan="2" style="padding:0 36px 0">
            <div style="height:3px;background:linear-gradient(90deg,#C49A3C,#D4682A,#C49A3C);border-radius:2px"></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- CORPO -->
  <tr>
    <td style="background:#ffffff;padding:36px 40px 28px;border-left:1px solid #E8D5B8;border-right:1px solid #E8D5B8">
      {paragrafos}
    </td>
  </tr>

  <!-- PRODUTO DESTAQUE -->
  <tr>
    <td style="background:#FDF6EF;border:1px solid #E8D5B8;border-top:none;padding:20px 40px">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="border-left:3px solid #C49A3C;padding-left:14px">
            <p style="margin:0;font-size:11px;font-weight:700;color:#C49A3C;letter-spacing:.1em;text-transform:uppercase">O produto</p>
            <p style="margin:6px 0 0;font-size:14px;font-weight:700;color:#5C2D0E">Pão de Ló Ti'Piedade — Unidose 85g (congelado)</p>
            <p style="margin:4px 0 0;font-size:12px;color:#6B5744">Original · Chocolate · Canela · Café &nbsp;|&nbsp; Receita artesanal desde 1984 &nbsp;|&nbsp; Sem conservantes</p>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- FOOTER -->
  <tr>
    <td style="background:#5C2D0E;border-radius:0 0 14px 14px;padding:20px 36px">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td>
            <p style="margin:0;font-size:12px;color:rgba(255,255,255,.7)">
              <strong style="color:#C49A3C">{nome_comercial}</strong> · Equipa Comercial Ti'Piedade
            </p>
            <p style="margin:4px 0 0;font-size:11px;color:rgba(255,255,255,.45)">
              sales@tipiedade.com &nbsp;|&nbsp; www.tipiedade.com
            </p>
          </td>
          <td align="right">
            <a href="https://www.tipiedade.com" style="display:inline-block;background:#C49A3C;color:#fff;font-size:11px;font-weight:700;padding:7px 16px;border-radius:6px;text-decoration:none;letter-spacing:.04em">
              Ver produto ↗
            </a>
          </td>
        </tr>
        <tr>
          <td colspan="2" style="padding-top:12px;border-top:1px solid rgba(255,255,255,.1);margin-top:12px">
            <p style="margin:0;font-size:10px;color:rgba(255,255,255,.3);line-height:1.5">
              Recebeu este email porque o vosso espaço foi identificado como potencial parceiro do Pão de Ló Ti'Piedade.
              Para não receber mais comunicações, responda com o assunto "Remover".
            </p>
          </td>
        </tr>
      </table>
    </td>
  </tr>

</table>
</td></tr>
</table>
</body>
</html>"""


def enviar_nurturing_lead(server, smtp_user, lead, email_num, nome_comercial):
    """Envia o email de nurturing em HTML directamente ao email da lead."""
    email_dest = lead.get("email","").strip()
    if not email_dest or email_dest == "—" or "@" not in email_dest:
        return False

    emails = [e.strip() for e in email_dest.replace(";",",").split(",") if "@" in e.strip()]
    if not emails:
        return False

    grupo   = tipologia_grupo(lead.get("t",""))
    assunto = ASSUNTOS.get(grupo, ASSUNTOS["restaurante"]).get(email_num, "")
    texto   = corpo_email(email_num, grupo, lead.get("n",""), lead.get("zona",""), nome_comercial)
    html    = corpo_html(email_num, grupo, lead.get("n",""), lead.get("zona",""), nome_comercial, texto)

    for dest_email in emails:
        msg = MIMEMultipart("alternative")
        msg["From"]    = f"Pão de Ló Ti'Piedade <{EMAIL_FROM}>"
        msg["To"]      = dest_email
        msg["Subject"] = f"{assunto} | Ti'Piedade"
        # Plain text fallback
        msg.attach(MIMEText(texto, "plain", "utf-8"))
        # HTML principal
        msg.attach(MIMEText(html, "html", "utf-8"))
        # Enviar para destinatário + BCC duplo
        server.sendmail(EMAIL_FROM, [dest_email, EMAIL_CC, EMAIL_BCC2], msg.as_string())

    return True


def enviar(server, de, para, cc, assunto, corpo, ficheiro=None):
    """Envia email com ou sem anexo."""
    msg = MIMEMultipart()
    msg["From"] = f"Ti'Piedade HORECA <{de}>"
    msg["To"]   = para
    msg["CC"]   = cc
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "plain", "utf-8"))
    if ficheiro:
        with open(ficheiro,"rb") as f:
            part = MIMEBase("application","octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(ficheiro)}"')
        msg.attach(part)
    server.sendmail(de, [para, cc], msg.as_string())


def enviar_emails(ficheiro, sem, modo):
    smtp_host = os.environ["SMTP_HOST"]
    smtp_port = int(os.environ.get("SMTP_PORT","587"))
    smtp_user = os.environ["SMTP_USER"]
    smtp_pass = os.environ["SMTP_PASS"]

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.ehlo(); server.starttls(); server.login(smtp_user, smtp_pass)

        if modo == "coordenador":
            total_horeca = sum(len(gerar_leads_horeca(c,sem)) for c in ["nuno","joao","oscar"])
            corpo = f"""Bom dia,

Resumo de leads gerados para revisão — Semana {sem} ({semana_datas()}):

HORECA por comercial:
  Nuno  → {len(gerar_leads_horeca('nuno',sem))} leads
  João  → {len(gerar_leads_horeca('joao',sem))} leads
  Óscar → {len(gerar_leads_horeca('oscar',sem))} leads

Para ti:
  Catering & Eventos → {len(gerar_leads_canal(DB_CATERING,'rui_catering',sem,5))} leads
  Distribuidores     → {len(gerar_leads_canal(DB_DISTRIBUIDORES,'rui_distribuidores',sem,5))} leads

Total: {total_horeca + 10} leads esta semana

Na quarta-feira:
→ Os comerciais recebem as suas leads
→ O Email 1 de nurturing é enviado automaticamente a cada lead com email disponível
→ O histórico é atualizado no CRM

O Excel em anexo tem o detalhe completo.

Ti'Piedade — Sistema de Prospeção HORECA
"""
            enviar(server, smtp_user, EMAIL_RUI, EMAIL_CC,
                   f"[REVISÃO] Leads Semana {sem} — {total_horeca+10} contactos | Ti'Piedade",
                   corpo, ficheiro)
            print(f"✓ Resumo enviado para {EMAIL_RUI}")

        elif modo == "comerciais":
            historico, sha = ler_historico()
            enviados_leads = 0
            sem_email = 0

            for com_id in ["nuno","joao","oscar"]:
                com = COMERCIAIS[com_id]
                leads = gerar_leads_horeca(com_id, sem)
                dest = com["email"]

                # Determinar email_num para esta semana (leads novas = 1, continuação = próximo)
                email_num = next(
                    (calcular_email_num(historico, com_id, l) for l in leads
                     if calcular_email_num(historico, com_id, l)),
                    1
                )

                # ── Enviar nurturing diretamente às leads ──────────
                leads_com_email = 0
                for lead in leads:
                    n = calcular_email_num(historico, com_id, lead)
                    if n is None:
                        continue  # sequência completa
                    ok = enviar_nurturing_lead(server, smtp_user, lead, n, com["nome"])
                    if ok:
                        leads_com_email += 1
                        enviados_leads += 1
                        historico = registar_envios(historico, com_id, [lead], n, com["nome"])
                    else:
                        sem_email += 1

                # ── Email resumo ao comercial ───────────────────────
                if dest:
                    grupo_exemplo = tipologia_grupo(leads[0]["t"]) if leads else "restaurante"
                    assunto_ex = ASSUNTOS.get(grupo_exemplo, ASSUNTOS["restaurante"]).get(email_num,"")
                    corpo_com = f"""Olá {com['nome']},

Aqui estão os teus {len(leads)} leads HORECA para a semana {sem} ({semana_datas()}).

Esta semana enviámos automaticamente o Email {email_num} da sequência de nurturing a {leads_com_email} leads com email disponível ({sem_email} sem email — contacto telefónico direto).

O assunto usado foi: "{assunto_ex}"

No Excel em anexo (coluna "Próx. Email") vês o estado de cada lead na sequência.
Quando aparecer "✓ Pronto p/ visita" — o contacto recebeu os 4 emails. É altura de visitar.

Bom trabalho,
Equipa Comercial Ti'Piedade
"""
                    enviar(server, smtp_user, dest, EMAIL_CC,
                           f"Leads Semana {sem} — {len(leads)} contactos | Ti'Piedade",
                           corpo_com, ficheiro)
                    print(f"✓ {com['nome']} — {leads_com_email} emails nurturing enviados às leads")

            # ── Rui — catering ────────────────────────────────────
            leads_cat = gerar_leads_canal(DB_CATERING, "rui_catering", sem, 5)
            for lead in leads_cat:
                n = calcular_email_num(historico, "rui_catering", lead)
                if n:
                    ok = enviar_nurturing_lead(server, smtp_user, lead, n, "Rui")
                    if ok:
                        enviados_leads += 1
                        historico = registar_envios(historico, "rui_catering", [lead], n, "Rui")

            corpo_cat = f"""Olá,

Aqui estão os 5 leads de Catering & Eventos para a semana {sem} ({semana_datas()}).

O Email 1 de nurturing foi enviado automaticamente às leads com email disponível.
Para as restantes, o contacto é telefónico direto.

Ti'Piedade — Sistema de Prospeção
"""
            enviar(server, smtp_user, EMAIL_RUI, EMAIL_CC,
                   f"Leads Catering & Eventos — Semana {sem} | Ti'Piedade",
                   corpo_cat, ficheiro)
            print(f"✓ Catering → Rui")

            # ── Rui — distribuidores ──────────────────────────────
            leads_dist = gerar_leads_canal(DB_DISTRIBUIDORES, "rui_distribuidores", sem, 5)
            for lead in leads_dist:
                n = calcular_email_num(historico, "rui_distribuidores", lead)
                if n:
                    ok = enviar_nurturing_lead(server, smtp_user, lead, n, "Rui")
                    if ok:
                        enviados_leads += 1
                        historico = registar_envios(historico, "rui_distribuidores", [lead], n, "Rui")

            corpo_dist = f"""Olá,

Aqui estão os 5 leads de Distribuidores para a semana {sem} ({semana_datas()}).

O Email 1 de nurturing foi enviado automaticamente às leads com email disponível.

Ti'Piedade — Sistema de Prospeção
"""
            enviar(server, smtp_user, EMAIL_RUI, EMAIL_CC,
                   f"Leads Distribuidores — Semana {sem} | Ti'Piedade",
                   corpo_dist, ficheiro)
            print(f"✓ Distribuidores → Rui")

            print(f"\n📧 Total nurturing enviado: {enviados_leads} emails às leads | {sem_email} sem email")
            escrever_historico(historico, sha)

# ════════════════════════════════════════════════════════════════

if __name__=="__main__":
    sem=semana_num(); modo=os.environ.get("MODO","coordenador")
    print(f"▶ Semana {sem} — {semana_datas()} — Modo: {modo}")
    historico,sha=ler_historico()
    ficheiro=criar_excel(sem,historico)
    print(f"✓ Excel: {ficheiro}")
    enviar_emails(ficheiro,sem,modo)
    print("✓ Concluído.")
