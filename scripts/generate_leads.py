"""
Pão de Ló Ti'Piedade — Gerador de Leads HORECA
Corre toda a segunda-feira via GitHub Actions.
Gera um Excel por comercial + resumo e envia por email SMTP.
"""

import math
import smtplib
import os
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ════════════════════════════════════════════════════════════════
# BASE DE LEADS
# ════════════════════════════════════════════════════════════════

DB = {
"Lisboa": [
  {"n":"Tasca do Chico","t":"Restaurante","m":"R. do Diário de Notícias 39, Lisboa","tel":"965 059 670","email":"info@tascadochico.pt","p":"Alta","tCliente":"Tasca contemporânea","gancho":"Rotatividade alta de turistas — a sobremesa pronta a servir resolve o fim de refeição com zero desperdício."},
  {"n":"Solar dos Presuntos","t":"Restaurante","m":"R. das Portas de Santo Antão 150, Lisboa","tel":"213 424 253","email":"geral@solardospresuntos.com","p":"Alta","tCliente":"Restaurante tradicional / turismo","gancho":"Volume de turistas alto — pão de ló português é produto de eleição para terminar uma refeição típica."},
  {"n":"Martinho da Arcada","t":"Restaurante","m":"Pr. do Comércio 3, Lisboa","tel":"218 879 259","email":"geral@martinhodaarcada.pt","p":"Alta","tCliente":"Histórico / turismo","gancho":"O produto mais português possível para o restaurante mais antigo de Lisboa — narrativa fácil de vender ao garção."},
  {"n":"Pastéis de Belém","t":"Pastelaria","m":"R. de Belém 84, Lisboa","tel":"213 637 423","email":"geral@pasteisdebelem.pt","p":"Alta","tCliente":"Pastelaria icónica / turismo","gancho":"Parceria entre duas referências da doçaria artesanal portuguesa — produto complementar, não concorrente."},
  {"n":"Landeau Chocolate","t":"Café","m":"R. das Flores 70, Lisboa","tel":"214 792 178","email":"hello@landeau.pt","p":"Alta","tCliente":"Café de nicho / produto","gancho":"Espaço que valoriza produtos com história e origem — dose individual em congelado resolve logística sem perder qualidade."},
  {"n":"Taberna Rua das Flores","t":"Restaurante","m":"R. das Flores 103, Lisboa","tel":"213 479 418","email":"info@tabernaruadasflores.pt","p":"Média","tCliente":"Tasca contemporânea","gancho":"Rotatividade alta de turistas — produto pronto a servir elimina desperdício."},
  {"n":"Pharmácia","t":"Restaurante","m":"R. Marechal Saldanha 1, Lisboa","tel":"213 465 146","email":"geral@museudafarmacia.pt","p":"Média","tCliente":"Restaurante temático / cultura","gancho":"Espaço criativo com clientela atenta à origem dos produtos — pão de ló como sobremesa de autor."},
  {"n":"Decadente","t":"Restaurante","m":"R. de São Pedro de Alcântara 45, Lisboa","tel":"213 957 936","email":"info@odecadente.pt","p":"Média","tCliente":"Bistrô / residentes","gancho":"Carta curta e rotativa — produto congelado em dose individual encaixa na filosofia de evitar desperdício."},
  {"n":"Mini Bar Teatro","t":"Restaurante","m":"R. António Maria Cardoso 58, Lisboa","tel":"211 305 393","email":"info@minibar.pt","p":"Média","tCliente":"Restaurante criativo","gancho":"Público jovem e urbano — pão de ló chocolate ou canela como sobremesa diferenciada."},
  {"n":"Copenhagen Coffee Lab","t":"Café","m":"R. Nova da Piedade 10, Lisboa","tel":"—","email":"hello@copenhagencoffeelab.com","p":"Média","tCliente":"Café de especialidade","gancho":"Audiência de nicho que valoriza produto artesanal — pairing pão de ló + café de especialidade."},
  {"n":"Mercearia do Bairro","t":"Mercearia Gourmet","m":"R. do Açúcar 83, Lisboa","tel":"—","email":"geral@merceariadobairro.pt","p":"Média","tCliente":"Mercearia de bairro gourmet","gancho":"Produto português artesanal com 40 anos de história — diferenciador face a produto industrial."},
  {"n":"Clube de Jornalistas","t":"Restaurante","m":"R. das Trinas 129, Lisboa","tel":"213 977 138","email":"geral@clubedejornalistas.com","p":"Média","tCliente":"Restaurante clássico / residentes","gancho":"Clientela fiel e recorrente — sobremesa clássica como âncora de carta."},
  {"n":"Tasca do Lagarto","t":"Restaurante","m":"R. dos Bacalhoeiros 34, Lisboa","tel":"—","email":"info@tascadolagarto.pt","p":"Média","tCliente":"Tasca / turismo","gancho":"Localização Alfama — turistas com apetência por sobremesas genuinamente portuguesas."},
  {"n":"Café de São Bento","t":"Café","m":"R. de São Bento 212, Lisboa","tel":"213 952 911","email":"geral@cafesaobento.pt","p":"Média","tCliente":"Café clássico","gancho":"Espaço histórico com clientela estabelecida — pão de ló como sobremesa de balcão premium."},
  {"n":"O Pitéu da Graça","t":"Restaurante","m":"Pr. da Graça 96, Lisboa","tel":"218 870 565","email":"—","p":"Baixa","tCliente":"Restaurante de bairro","gancho":"Dose individual congelada controla custo e elimina desperdício."},
],
"Santarém": [
  {"n":"Restaurante O Salazares","t":"Restaurante","m":"R. de São Martinho 2, Santarém","tel":"243 322 384","email":"info@osalazares.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Capital gastronómica regional — produto com raízes locais e receita secular tem aceitação natural."},
  {"n":"Restaurante Portas do Sol","t":"Restaurante","m":"Jardim das Portas do Sol, Santarém","tel":"243 309 520","email":"info@portasdosol.pt","p":"Alta","tCliente":"Restaurante com vista / turismo","gancho":"Turismo de alto valor e grupos — sobremesa pronta a servir em volume sem perda de qualidade."},
  {"n":"Hotel Cristal Santarém","t":"Hotel","m":"R. Francisco Moreira 7, Santarém","tel":"243 377 575","email":"reservas@hotelcristal.pt","p":"Alta","tCliente":"Hotel 3* / negócios","gancho":"Restaurante de hotel com necessidade de carta de sobremesas sólida sem pasteleiro próprio."},
  {"n":"Pastelaria Bijou","t":"Pastelaria","m":"Av. Bernardo Santareno, Santarém","tel":"243 322 507","email":"—","p":"Alta","tCliente":"Pastelaria de referência local","gancho":"Produto complementar ao catálogo existente — dose individual como oferta de impulso ao balcão."},
  {"n":"Pastelaria Paraíso","t":"Pastelaria","m":"Av. Marquês de Sá da Bandeira, Santarém","tel":"243 322 441","email":"—","p":"Média","tCliente":"Pastelaria de bairro","gancho":"Diferenciação face à concorrência com um produto artesanal."},
  {"n":"Tasca do Escondidinho","t":"Restaurante","m":"R. Capelo e Ivens, Santarém","tel":"243 323 991","email":"—","p":"Média","tCliente":"Tasca tradicional","gancho":"Clientela local com apreço por doçaria tradicional — pão de ló como proposta de sobremesa caseira."},
  {"n":"Mercearia Tradicional do Ribatejo","t":"Mercearia Gourmet","m":"Lg. da Feira, Santarém","tel":"—","email":"—","p":"Média","tCliente":"Mercearia gourmet","gancho":"Produto regional artesanal — argumento de proximidade e identidade ribatejana."},
],
"Linha Sintra–Cascais": [
  {"n":"Bar do Fundo","t":"Restaurante","m":"Av. Alfredo Coelho, Praia Grande, Colares","tel":"219 282 092","email":"info@bardofundo.pt","p":"Alta","tCliente":"Restaurante premium / vista mar","gancho":"Clientela turística de poder de compra elevado — pão de ló como sobremesa de autor com vista para o Atlântico."},
  {"n":"Restaurante Azenhas do Mar","t":"Restaurante","m":"Lugar das Piscinas, 2705-098 Colares","tel":"219 280 739","email":"info@azenhasdomar.com","p":"Alta","tCliente":"Restaurante icónico / turismo","gancho":"Um dos restaurantes mais fotografados de Portugal — sobremesa artesanal portuguesa reforça o posicionamento."},
  {"n":"Taberna Clandestina","t":"Restaurante","m":"R. Afonso Sanches 36, Cascais","tel":"916 229 630","email":"info@tabernaclandestina.pt","p":"Alta","tCliente":"Gastropub / residentes premium","gancho":"Carta criativa com enfoque em produto português — pão de ló Ti'Piedade como sobremesa de referência."},
  {"n":"Hífen","t":"Restaurante","m":"Av. Dom Carlos I 48, Cascais","tel":"915 546 537","email":"info@hifenrestaurant.com","p":"Alta","tCliente":"Restaurante de referência / Cascais","gancho":"Restaurante com listas de espera — produto com história diferencia a experiência."},
  {"n":"Hotel Palácio Estoril","t":"Hotel","m":"R. Particular, Estoril","tel":"214 648 000","email":"info@palacioestoril.com","p":"Alta","tCliente":"Hotel 5* / luxo","gancho":"F&B de hotel de luxo — unidose congelada mantém qualidade constante em grande volume."},
  {"n":"Lawrence's Hotel (Sintra)","t":"Hotel","m":"R. Consiglieri Pedroso 38, Sintra","tel":"219 105 500","email":"info@lawrenceshotel.com","p":"Alta","tCliente":"Boutique hotel / turismo cultural","gancho":"Hotel histórico no coração de Sintra — produto artesanal português complementa a narrativa de autenticidade."},
  {"n":"Gourmet Italiano (Cascais)","t":"Mercearia Gourmet","m":"Av. Infante Dom Henrique 1027 D, Cascais","tel":"214 842 127","email":"info@gourmetitaliano.pt","p":"Alta","tCliente":"Deli gourmet / expatriados","gancho":"Clientela internacional com elevado poder de compra — produto português artesanal como oferta local de qualidade."},
  {"n":"Emporium Gourmet","t":"Mercearia Gourmet","m":"Av. Nossa Senhora do Cabo 101, Cascais","tel":"211 541 588","email":"info@emporiumgourmet.pt","p":"Alta","tCliente":"Mercearia gourmet","gancho":"Espaço que cuida da narrativa de cada produto — história de 40 anos e receita da D. Piedade vende-se sozinha."},
  {"n":"Taberna Económica de Cascais","t":"Restaurante","m":"R. Sebastião José de Carvalho e Melo 35, Cascais","tel":"214 832 214","email":"info@tabernaeconomicadecascais.com","p":"Alta","tCliente":"Taberna / turismo","gancho":"Volume de turistas e rotatividade alta — produto congelado garante consistência e elimina desperdício."},
  {"n":"Angra Gatti","t":"Restaurante","m":"Av. Alfredo Coelho 57, Praia Grande, Colares","tel":"965 770 247","email":"info@angragatti.com","p":"Alta","tCliente":"Restaurante italiano / destino","gancho":"Clientela de destino — sobremesa portuguesa como proposta de fecho de refeição."},
  {"n":"Mana","t":"Restaurante","m":"Tv. Navegantes 13, Cascais","tel":"915 669 206","email":"info@manacascais.pt","p":"Média","tCliente":"Restaurante & bar / trendy","gancho":"Público jovem e urbano — pão de ló chocolate ou canela na carta."},
  {"n":"Café Paris (Sintra)","t":"Café","m":"Pr. da República 32, Sintra","tel":"219 232 375","email":"—","p":"Média","tCliente":"Café turístico","gancho":"Ponto de passagem obrigatório em Sintra — produto português icónico para turistas internacionais."},
  {"n":"Casa da Galé","t":"Restaurante","m":"Av. Alfredo Coelho 61, Praia Grande, Colares","tel":"219 291 218","email":"—","p":"Média","tCliente":"Restaurante peixe / local","gancho":"Refeições de peixe pedem sobremesa leve — pão de ló é o final natural."},
],
"SuperIndep_Nuno": [
  {"n":"Mercado Municipal de Campo de Ourique","t":"Supermercado Independente","m":"R. Coelho da Rocha, Lisboa","tel":"213 954 628","email":"mercado@cm-lisboa.pt","p":"Alta","tCliente":"Mercado alimentar / produto fresco","gancho":"Espaço com lojistas independentes e clientela premium — pão de ló em congelado é produto de rotação com margens interessantes."},
  {"n":"Honest Greens Market (Cascais)","t":"Supermercado Independente","m":"Av. Marginal, São João do Estoril","tel":"—","email":"—","p":"Média","tCliente":"Supermercado independente / saudável","gancho":"Clientela de classe média-alta com procura de produto artesanal — dose individual em congelado é formato conveniente."},
],
"Margem Sul": [
  {"n":"O Farol Design Hotel","t":"Hotel","m":"R. do Farol 1, Cacilhas, Almada","tel":"210 407 040","email":"info@farolhotel.com","p":"Alta","tCliente":"Boutique hotel / design","gancho":"Hotel de autor com público sofisticado — produto artesanal português de alto valor percebido."},
  {"n":"Hotel Sana Sesimbra","t":"Hotel","m":"Av. 25 de Abril, Sesimbra","tel":"212 289 000","email":"info@sesimbra.sanahotels.com","p":"Alta","tCliente":"Resort / lazer","gancho":"F&B de resort com volume elevado — unidose congelada garante qualidade constante."},
  {"n":"Restaurante Ribamar (Sesimbra)","t":"Restaurante","m":"Av. dos Náufragos 29, Sesimbra","tel":"212 233 853","email":"info@restauranteribamar.com","p":"Alta","tCliente":"Restaurante peixe / turismo","gancho":"Destino de verão — produto pronto a servir resolve o fim de refeição em período de pico."},
  {"n":"Tasca D'Avenida","t":"Restaurante","m":"Av. Dom Afonso Henriques 10C, Almada","tel":"968 348 036","email":"—","p":"Alta","tCliente":"Tasca contemporânea","gancho":"Clientela de almoço de negócios — sobremesa premium com margem confortável."},
  {"n":"The Baptist","t":"Restaurante","m":"R. Afonso Galo 56, Almada","tel":"212 750 996","email":"—","p":"Média","tCliente":"Restaurante casual","gancho":"Volume de clientes e rotatividade — produto pronto a servir em dose individual resolve a sobremesa."},
  {"n":"Restaurante Paladar (Palmela)","t":"Restaurante","m":"R. João de Deus, Palmela","tel":"—","email":"—","p":"Média","tCliente":"Restaurante regional","gancho":"Zona de turismo de interior com apetência por produto artesanal regional."},
  {"n":"Pastelaria A Floresta (Barreiro)","t":"Pastelaria","m":"Av. Bento Gonçalves, Barreiro","tel":"—","email":"—","p":"Média","tCliente":"Pastelaria de bairro","gancho":"Produto de diferenciação — dose individual ao balcão com margem interessante."},
],
"SuperIndep_Joao": [
  {"n":"Supermercado Apolónia (Leiria)","t":"Supermercado Independente","m":"Av. Heróis de Angola, Leiria","tel":"244 859 900","email":"leiria@apolonia.pt","p":"Alta","tCliente":"Supermercado premium independente","gancho":"Apolónia posiciona-se em produto de qualidade e nacional — pão de ló congelado em unidose é referência perfeita para o segmento gourmet."},
  {"n":"Mini Mercado Tradicional da Caparica","t":"Supermercado Independente","m":"R. da Liberdade, Costa da Caparica","tel":"—","email":"—","p":"Média","tCliente":"Mini mercado / bairro de praia","gancho":"Zona balnear com consumo de verão elevado — produto congelado com exposição no linear de sobremesas."},
],
"Costa Oeste (S. Martinho–Vieira)": [
  {"n":"Hotel Columbano (S. Martinho do Porto)","t":"Hotel","m":"Av. Marginal, S. Martinho do Porto","tel":"262 989 220","email":"info@hotelcolumbano.com","p":"Alta","tCliente":"Hotel 4* / praia","gancho":"F&B de hotel de praia com procura de verão intensa — produto congelado mantém qualidade sem complexidade logística."},
  {"n":"Tasca do Zé (Nazaré)","t":"Restaurante","m":"R. Mouzinho de Albuquerque 22, Nazaré","tel":"262 551 945","email":"—","p":"Alta","tCliente":"Tasca turística / Nazaré","gancho":"A Nazaré atrai turismo internacional massivo — pão de ló como sobremesa típica portuguesa."},
  {"n":"Restaurante O Casalinho (Nazaré)","t":"Restaurante","m":"R. do Elevador 24, Nazaré","tel":"262 552 608","email":"—","p":"Alta","tCliente":"Restaurante peixe / turismo","gancho":"Refeições de peixe com turistas internacionais — produto de identidade nacional para fechar a refeição."},
  {"n":"Hotel Maré (Nazaré)","t":"Hotel","m":"R. Mouzinho de Albuquerque 8, Nazaré","tel":"262 550 000","email":"info@hotelmare.com","p":"Alta","tCliente":"Hotel / turismo surf","gancho":"Público de surf e natureza que valoriza produto artesanal com raízes locais."},
  {"n":"Solar de Alcobaça","t":"Restaurante","m":"Pr. 25 de Abril, Alcobaça","tel":"262 598 312","email":"—","p":"Média","tCliente":"Restaurante turístico / mosteiro","gancho":"Alcobaça é destino patrimonial — produto com raízes medievais encaixa perfeitamente."},
  {"n":"Restaurante A Tasquinha (Alcobaça)","t":"Restaurante","m":"R. Frei António Brandão 2, Alcobaça","tel":"262 582 397","email":"—","p":"Média","tCliente":"Restaurante local","gancho":"Clientela mista de residentes e turistas — produto artesanal como âncora da carta de sobremesas."},
],
"Leiria": [
  {"n":"Tromba Rija","t":"Restaurante","m":"R. Professores Portelas, Marrazes, Leiria","tel":"244 855 072","email":"geral@trombarja.com","p":"Alta","tCliente":"Restaurante regional / referência","gancho":"Referência gastronómica da região — produto artesanal com receita secular encaixa no posicionamento de valorização do território."},
  {"n":"Hotel Eurosol Leiria","t":"Hotel","m":"R. Comissão da Iniciativa, Leiria","tel":"244 838 201","email":"info@eurosolleiria.pt","p":"Alta","tCliente":"Hotel 4* / negócios","gancho":"Restaurante de hotel com necessidade de carta consistente — produto congelado resolve sobremesas sem chefe de pastelaria."},
  {"n":"Pastelaria Garrett","t":"Pastelaria","m":"Pr. Rodrigues Lobo, Leiria","tel":"244 812 370","email":"—","p":"Alta","tCliente":"Pastelaria de referência","gancho":"Espaço de referência local — produto artesanal complementa o catálogo e diferencia da concorrência."},
  {"n":"Tasca da Barrosinha","t":"Restaurante","m":"Pr. Rodrigues Lobo 2, Leiria","tel":"244 823 703","email":"—","p":"Alta","tCliente":"Tasca tradicional","gancho":"Clientela local fidelizada — pão de ló como proposta de sobremesa clássica com margem confortável."},
  {"n":"O Funil (Leiria)","t":"Restaurante","m":"Av. Heróis de Angola 66, Leiria","tel":"244 832 522","email":"—","p":"Média","tCliente":"Restaurante casual","gancho":"Volume de almoços de negócios — produto pronto a servir acelera rotatividade."},
  {"n":"Patisserie Almonda (Tomar)","t":"Pastelaria","m":"Av. Marquês de Tomar, Tomar","tel":"249 312 252","email":"—","p":"Média","tCliente":"Pastelaria de destino","gancho":"Tomar é destino turístico — produto português de referência para turistas."},
],
"Ericeira–Caldas da Rainha": [
  {"n":"Marginal (Peniche)","t":"Restaurante","m":"Estr. Marginal Norte, Peniche","tel":"968 907 248","email":"marginalrestaurante@gmail.com","p":"Alta","tCliente":"Restaurante premium / costa","gancho":"Refeição premium com vista mar — sobremesa artesanal portuguesa fecha a experiência de forma coerente."},
  {"n":"Restaurante Sueste (Ericeira)","t":"Restaurante","m":"R. Eduardo Burnay 22, Ericeira","tel":"261 862 108","email":"info@sueste.pt","p":"Alta","tCliente":"Restaurante peixe / turismo surf","gancho":"Ericeira tem comunidade surf internacional — produto artesanal português como proposta autêntica."},
  {"n":"Hotel Termas das Caldas","t":"Hotel","m":"Pr. 25 de Abril, Caldas da Rainha","tel":"262 830 200","email":"info@termascaldas.pt","p":"Alta","tCliente":"Hotel termal / saúde","gancho":"Clientela de bem-estar — produto artesanal sem conservantes, ingredientes simples e receita secular."},
  {"n":"Adega do Caseiro (Caldas)","t":"Restaurante","m":"R. Eng. Duarte Pacheco, Caldas da Rainha","tel":"262 831 291","email":"—","p":"Alta","tCliente":"Restaurante regional","gancho":"Referência local — produto com identidade regional reforça o posicionamento do espaço."},
  {"n":"Chico Neto (Ribamar)","t":"Restaurante","m":"R. das Armaçõe 26, Ribamar","tel":"261 422 106","email":"—","p":"Alta","tCliente":"Restaurante peixe / local","gancho":"Clientela de fim-de-semana — sobremesa clássica para famílias."},
  {"n":"O Viveiro (Ribamar)","t":"Restaurante","m":"R. das Armaçõe 7, Ribamar","tel":"261 422 197","email":"—","p":"Alta","tCliente":"Restaurante peixe / vista mar","gancho":"Vista mar e clientela de qualidade — produto artesanal português eleva a carta sem complexidade."},
  {"n":"Cafetaria Puro Cake Lab","t":"Pastelaria","m":"Pr. Jacob Rodrigues Pereira 18, Peniche","tel":"916 950 480","email":"info@purocakelab.pt","p":"Alta","tCliente":"Pastelaria artesanal","gancho":"Espaço que valoriza produto artesanal — pão de ló Ti'Piedade como oferta complementar de referência."},
  {"n":"Pastelaria Princesa do Mar","t":"Pastelaria","m":"R. António Maria Oliveira 34, Peniche","tel":"262 782 929","email":"—","p":"Alta","tCliente":"Pastelaria local","gancho":"Produto complementar ao catálogo existente — dose individual com margem de revenda interessante."},
  {"n":"Restaurante Ó Baleal","t":"Restaurante","m":"Baleal, Peniche","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante surf / natureza","gancho":"Baleal é ícone do surf — produto artesanal português como proposta autêntica para cliente internacional."},
],
"SuperIndep_Oscar": [
  {"n":"Supermercado Apolónia (Porto)","t":"Supermercado Independente","m":"R. de Júlio Dinis 826, Porto","tel":"226 066 730","email":"porto@apolonia.pt","p":"Alta","tCliente":"Supermercado premium independente","gancho":"Rede independente com foco em produto nacional de qualidade — pão de ló Ti'Piedade é referência natural para o segmento."},
  {"n":"Mercado Bom Sucesso (Porto)","t":"Supermercado Independente","m":"Pr. do Bom Sucesso 74, Porto","tel":"226 088 800","email":"info@mercadobomsucesso.com","p":"Alta","tCliente":"Mercado gourmet / turismo","gancho":"Mercado de referência no Porto com lojistas independentes e turismo massivo — produto artesanal português com 40 anos encaixa naturalmente."},
],
"Coimbra": [
  {"n":"Fangas Mercearia Bar","t":"Mercearia Gourmet","m":"R. Fernandes Tomás 45, Coimbra","tel":"239 115 540","email":"info@fangas.pt","p":"Alta","tCliente":"Mercearia gourmet / bar","gancho":"Curadoria de produto nacional — pão de ló com 40 anos de história é produto de prateleira natural."},
  {"n":"Hotel Quinta das Lágrimas","t":"Hotel","m":"Santa Clara, Coimbra","tel":"239 802 380","email":"reservas@quintadaslagrimas.pt","p":"Alta","tCliente":"Hotel 5* / romance / natureza","gancho":"Hotel de luxo com narrativa histórica — produto artesanal com receita secular reforça a experiência portuguesa autêntica."},
  {"n":"Restaurante O Trovador","t":"Restaurante","m":"Lg. da Sé Velha 15, Coimbra","tel":"239 825 475","email":"info@otrovador.pt","p":"Alta","tCliente":"Restaurante histórico / fado","gancho":"Restaurante histórico com fado ao vivo — pão de ló como sobremesa de fim de noite com narrativa de Portugal."},
  {"n":"Café Santa Cruz","t":"Café","m":"Pr. 8 de Maio, Coimbra","tel":"239 833 617","email":"cafesantacruz@sapo.pt","p":"Alta","tCliente":"Café histórico / turismo","gancho":"Café mais emblemático de Coimbra — produto português artesanal com história."},
  {"n":"Pastelaria Briosa","t":"Pastelaria","m":"R. Direita, Coimbra","tel":"239 824 764","email":"—","p":"Alta","tCliente":"Pastelaria de referência","gancho":"Referência em Coimbra — produto de autor que complementa o catálogo sem competir diretamente."},
  {"n":"Adega Paço do Conde","t":"Restaurante","m":"R. Paço do Conde 1, Coimbra","tel":"239 825 605","email":"—","p":"Média","tCliente":"Restaurante clássico","gancho":"Clientela académica e cultural — pão de ló como sobremesa clássica universitária."},
  {"n":"Tasca da Rua Nova","t":"Restaurante","m":"R. Nova 44, Coimbra","tel":"239 826 669","email":"—","p":"Média","tCliente":"Tasca contemporânea","gancho":"Carta curta e bem curada — produto artesanal como âncora da sobremesa."},
],
"Porto": [
  {"n":"O Gaveto","t":"Restaurante","m":"R. Roberto Ivens 826, Matosinhos","tel":"229 381 879","email":"info@restauranteogaveto.com","p":"Alta","tCliente":"Restaurante peixe / referência","gancho":"Refeições de peixe com clientela exigente — pão de ló é a sobremesa natural para uma refeição típica portuguesa."},
  {"n":"Mercearia das Flores","t":"Mercearia Gourmet","m":"R. das Flores 110, Porto","tel":"222 013 290","email":"info@merceariadas flores.pt","p":"Alta","tCliente":"Mercearia gourmet / design","gancho":"Curadoria de produto nacional de qualidade — história de 40 anos e receita intacta é produto editorial."},
  {"n":"Hotel Infante de Sagres","t":"Hotel","m":"Pr. Filipa de Lencastre 62, Porto","tel":"223 398 500","email":"info@hotelinfantesagres.pt","p":"Alta","tCliente":"Hotel 5* histórico","gancho":"Hotel histórico no coração do Porto — produto artesanal português complementa a experiência premium."},
  {"n":"Café Majestic","t":"Café","m":"R. de Santa Catarina 112, Porto","tel":"222 003 887","email":"geral@cafemajestic.com","p":"Alta","tCliente":"Café histórico / turismo","gancho":"Um dos cafés mais visitados da Europa — pão de ló artesanal como proposta de doçaria nacional premium."},
  {"n":"Taberninha do Manel","t":"Restaurante","m":"Av. Gustavo Eiffel 274, Porto","tel":"222 086 389","email":"—","p":"Alta","tCliente":"Tasca histórica / turismo","gancho":"Clientela de turismo e residentes — pão de ló como sobremesa âncora num espaço de cozinha portuguesa."},
  {"n":"Pastelaria Luca","t":"Pastelaria","m":"R. de Sá da Bandeira 118, Porto","tel":"222 084 010","email":"—","p":"Alta","tCliente":"Pastelaria de referência","gancho":"Referência no Porto — produto artesanal nacional de autor complementa o catálogo existente."},
  {"n":"Casa de Pasto da Palmeira","t":"Restaurante","m":"R. de Palmeira 2, Porto","tel":"222 005 753","email":"—","p":"Alta","tCliente":"Casa de pasto / turismo","gancho":"Conceito de cozinha portuguesa simples e de qualidade — pão de ló é a sobremesa perfeita."},
  {"n":"Aduela","t":"Restaurante","m":"R. do Oliveiras 38, Porto","tel":"222 008 757","email":"—","p":"Média","tCliente":"Restaurante casual / wine bar","gancho":"Clientela jovem e urbana — pão de ló chocolate ou canela como sobremesa diferenciada."},
],
"Braga": [
  {"n":"Bem Me Quer (Braga)","t":"Restaurante","m":"Pr. do Município, Braga","tel":"253 278 916","email":"geral@restaurantebemmequeer.pt","p":"Alta","tCliente":"Restaurante de referência","gancho":"Referência em Braga — produto artesanal português como âncora premium da carta de sobremesas."},
  {"n":"Hotel Meliá Braga","t":"Hotel","m":"Av. General Carrilho da Silva Pinto 8, Braga","tel":"253 144 000","email":"melia.braga@melia.com","p":"Alta","tCliente":"Hotel 5* / congressos","gancho":"Volume de eventos e congressos — produto congelado em dose individual para servir em grande escala com qualidade consistente."},
  {"n":"Restaurante Inácio","t":"Restaurante","m":"Campo das Hortas 4, Braga","tel":"253 613 235","email":"—","p":"Alta","tCliente":"Restaurante clássico","gancho":"Referência gastronómica local — produto de pastelaria artesanal que reforça o posicionamento de qualidade."},
  {"n":"Pastelaria Riquexó","t":"Pastelaria","m":"Av. Central 69, Braga","tel":"253 215 055","email":"—","p":"Alta","tCliente":"Pastelaria clássica","gancho":"Espaço de referência — produto nacional artesanal que complementa o catálogo sem concorrer diretamente."},
  {"n":"Pastelaria Oliveira","t":"Pastelaria","m":"R. do Souto 128, Braga","tel":"253 215 990","email":"—","p":"Alta","tCliente":"Pastelaria de referência","gancho":"Central e muito frequentada — produto de diferenciação com margem elevada ao balcão."},
  {"n":"Taberna Belga","t":"Restaurante","m":"R. de Maximinos 121, Braga","tel":"253 204 786","email":"—","p":"Média","tCliente":"Gastropub / cerveja artesanal","gancho":"Público jovem — sobremesa artesanal portuguesa como proposta de fecho de refeição diferenciada."},
],
"Guimarães": [
  {"n":"Solar do Arco","t":"Restaurante","m":"R. de Santa Maria 48, Guimarães","tel":"253 513 072","email":"info@solardoarco.pt","p":"Alta","tCliente":"Restaurante histórico","gancho":"Centro histórico Património Mundial — produto artesanal português com receita secular encaixa perfeitamente."},
  {"n":"Pousada de Guimarães","t":"Hotel","m":"R. Conde de Margaride 153, Guimarães","tel":"253 511 249","email":"pousadaguimaraes@pousadas.pt","p":"Alta","tCliente":"Pousada histórica / turismo","gancho":"Pousada em mosteiro medieval — produto com receita levada pelos portugueses ao Japão no séc. XVI."},
  {"n":"El Rei","t":"Restaurante","m":"Pr. de São Tiago 20, Guimarães","tel":"253 419 096","email":"—","p":"Alta","tCliente":"Restaurante centro histórico","gancho":"Zona mais visitada de Guimarães — clientela turística internacional com apetência por produto típico."},
  {"n":"Pastelaria Clarinha","t":"Pastelaria","m":"R. de Santo António, Guimarães","tel":"253 512 552","email":"—","p":"Alta","tCliente":"Pastelaria de referência","gancho":"Referência local — produto nacional artesanal que complementa o catálogo com margem interessante."},
  {"n":"Mercearia Vimaranes","t":"Mercearia Gourmet","m":"Lg. do Toural, Guimarães","tel":"—","email":"—","p":"Média","tCliente":"Mercearia gourmet","gancho":"Produto com identidade nacional forte — proposta de qualidade e história para mercearia gourmet."},
  {"n":"Sabores do Minho","t":"Restaurante","m":"R. de Couros 24, Guimarães","tel":"—","email":"—","p":"Média","tCliente":"Restaurante regional","gancho":"Cozinha do Minho e turismo local — pão de ló como sobremesa de eleição numa carta de produtos regionais."},
],
}

COMERCIAIS = {
    "nuno":  {"nome": "Nuno",  "email": os.environ.get("EMAIL_NUNO",""),  "zonas": ["Lisboa","Santarém","Linha Sintra–Cascais","SuperIndep_Nuno"]},
    "joao":  {"nome": "João",  "email": os.environ.get("EMAIL_JOAO",""),  "zonas": ["Lisboa","Margem Sul","Costa Oeste (S. Martinho–Vieira)","Leiria","SuperIndep_Joao"]},
    "oscar": {"nome": "Óscar", "email": os.environ.get("EMAIL_OSCAR",""), "zonas": ["Ericeira–Caldas da Rainha","Coimbra","Porto","Braga","Guimarães","SuperIndep_Oscar"]},
}

TIPOS_ATIVOS = ["Restaurante","Pastelaria","Hotel","Mercearia Gourmet","Café","Supermercado Independente"]

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

def seeded_shuffle(lst, seed):
    result = list(lst)
    for i in range(len(result)-1, 0, -1):
        j = int(abs(math.sin(seed*(i+1)*9301+49297)*233280)) % (i+1)
        result[i], result[j] = result[j], result[i]
    return result

def gerar_leads(com_id, sem):
    com = COMERCIAIS[com_id]
    pool = []
    for zona in com["zonas"]:
        for lead in DB.get(zona, []):
            if lead["t"] in TIPOS_ATIVOS:
                zona_label = zona.replace("SuperIndep_Nuno","Lisboa").replace("SuperIndep_Joao","Margem Sul").replace("SuperIndep_Oscar","Porto")
                pool.append({**lead, "zona": zona_label})
    shuffled = seeded_shuffle(pool, sem * 1000 + list(COMERCIAIS.keys()).index(com_id) + 1)
    alta  = [l for l in shuffled if l["p"] == "Alta"]
    resto = [l for l in shuffled if l["p"] != "Alta"]
    return (alta + resto)[:20]

# ════════════════════════════════════════════════════════════════
# EXCEL
# ════════════════════════════════════════════════════════════════

COR_HDR = "5C2D0E"; COR_OURO = "C49A3C"; COR_ZEBRA = "FDF6EF"; COR_BORDA = "E8D5B8"; BRANCO = "FFFFFF"

def fill(c): return PatternFill("solid", fgColor=c)
def borda():
    s = Side(style="thin", color=COR_BORDA)
    return Border(left=s, right=s, top=s, bottom=s)
def centro(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def esq():    return Alignment(horizontal="left", vertical="top", wrap_text=True)

def criar_excel(sem):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for com_id, com in COMERCIAIS.items():
        leads = gerar_leads(com_id, sem)
        ws = wb.create_sheet(com["nome"])

        ws.merge_cells("A1:N1")
        ws["A1"] = f"PAO DE LÓ TI'PIEDADE — Leads Semana {sem} — {com['nome']}"
        ws["A1"].font = Font(name="Arial", bold=True, size=13, color=BRANCO)
        ws["A1"].fill = fill(COR_HDR); ws["A1"].alignment = centro()
        ws.row_dimensions[1].height = 30

        ws.merge_cells("A2:N2")
        ws["A2"] = f"{semana_datas()} · Unidose 85g congelado · Canal HORECA"
        ws["A2"].font = Font(name="Arial", size=9, italic=True, color="6B5744")
        ws["A2"].alignment = centro(); ws.row_dimensions[2].height = 16

        hdrs = ["#","Espaço","Tipo","Tipologia cliente","Zona","Morada","Contacto","Email",
                "Gancho comercial","Prioridade","Estado","Data contacto","Resultado","Observações"]
        for col, h in enumerate(hdrs, 1):
            c = ws.cell(row=3, column=col, value=h)
            c.font = Font(name="Arial", bold=True, size=9, color=COR_OURO)
            c.fill = fill(COR_HDR); c.alignment = centro(); c.border = borda()
        ws.row_dimensions[3].height = 26

        for i, lead in enumerate(leads):
            row = 4 + i
            cor = COR_ZEBRA if i % 2 == 0 else BRANCO
            vals = [i+1, lead["n"], lead["t"], lead["tCliente"], lead["zona"],
                    lead["m"], lead["tel"], lead["email"], lead["gancho"],
                    lead["p"], "Em aberto", "", "", ""]
            for col, v in enumerate(vals, 1):
                c = ws.cell(row=row, column=col, value=v)
                c.fill = fill(cor); c.border = borda()
                c.font = Font(name="Arial", size=9, bold=(col==2))
                c.alignment = centro() if col in [1,3,7,10,11,12,13] else esq()
            ws.row_dimensions[row].height = 36

        for col, w in enumerate([4,28,16,24,20,34,14,26,44,10,18,13,20,28], 1):
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.freeze_panes = "K4"
        ws.auto_filter.ref = f"A3:N{3+len(leads)}"

    # Resumo
    ws_r = wb.create_sheet("Resumo", 0)
    ws_r.merge_cells("A1:D1")
    ws_r["A1"] = f"TI'PIEDADE — Resumo Leads Semana {sem}"
    ws_r["A1"].font = Font(name="Arial", bold=True, size=13, color=BRANCO)
    ws_r["A1"].fill = fill(COR_HDR); ws_r["A1"].alignment = centro()
    ws_r.row_dimensions[1].height = 28

    for col, h in enumerate(["Comercial","Leads","Zonas","Email"], 1):
        c = ws_r.cell(row=2, column=col, value=h)
        c.font = Font(name="Arial", bold=True, size=9, color=COR_OURO)
        c.fill = fill(COR_HDR); c.alignment = centro(); c.border = borda()

    for i, (com_id, com) in enumerate(COMERCIAIS.items()):
        leads = gerar_leads(com_id, sem)
        row = 3 + i
        cor = COR_ZEBRA if i % 2 == 0 else BRANCO
        zonas_label = ", ".join(z for z in com["zonas"] if not z.startswith("SuperIndep"))
        for col, v in enumerate([com["nome"], len(leads), zonas_label, com["email"]], 1):
            c = ws_r.cell(row=row, column=col, value=v)
            c.fill = fill(cor); c.border = borda()
            c.font = Font(name="Arial", size=10, bold=(col==1))
            c.alignment = esq()
        ws_r.row_dimensions[row].height = 22

    ws_r.column_dimensions["A"].width = 14; ws_r.column_dimensions["B"].width = 8
    ws_r.column_dimensions["C"].width = 52; ws_r.column_dimensions["D"].width = 30

    fname = f"TiPiedade_Leads_S{sem}_{datetime.date.today().year}.xlsx"
    wb.save(fname)
    return fname

# ════════════════════════════════════════════════════════════════
# EMAIL
# ════════════════════════════════════════════════════════════════

EMAIL_CC = "sales@tipiedade.com"  # Recebe cópia de todos os envios

def enviar_emails(ficheiro, sem):
    smtp_host = os.environ["SMTP_HOST"]
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))
    smtp_user = os.environ["SMTP_USER"]
    smtp_pass = os.environ["SMTP_PASS"]

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.ehlo(); server.starttls(); server.login(smtp_user, smtp_pass)

        for com_id, com in COMERCIAIS.items():
            leads = gerar_leads(com_id, sem)
            dest = com["email"]
            if not dest:
                print(f"[AVISO] Email do {com['nome']} não configurado — a saltar.")
                continue

            msg = MIMEMultipart()
            msg["From"]    = f"Ti'Piedade HORECA <{smtp_user}>"
            msg["To"]      = dest
            msg["CC"]      = EMAIL_CC
            msg["Subject"] = f"Os teus {len(leads)} leads HORECA — Semana {sem} | Ti'Piedade"

            zonas_label = "\n".join(f"  • {z}" for z in com["zonas"] if not z.startswith("SuperIndep"))
            corpo = f"""Olá {com['nome']},

Seguem os teus {len(leads)} leads para a semana {sem} ({semana_datas()}).

Zonas desta semana:
{zonas_label}

O Excel em anexo inclui tipologia, gancho comercial e email de cada espaço.
Bons negócios,

Equipa Comercial Ti'Piedade
"""
            msg.attach(MIMEText(corpo, "plain", "utf-8"))
            with open(ficheiro, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{ficheiro}"')
            msg.attach(part)

            server.sendmail(smtp_user, [dest, EMAIL_CC], msg.as_string())
            print(f"✓ Email enviado: {com['nome']} ({dest}) + cópia {EMAIL_CC}")

# ════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    sem = semana_num()
    print(f"▶ Semana {sem} — {semana_datas()}")
    ficheiro = criar_excel(sem)
    print(f"✓ Excel: {ficheiro}")
    enviar_emails(ficheiro, sem)
    print("✓ Concluído.")
