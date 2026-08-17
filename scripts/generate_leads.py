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

# ── Base de leads reais (integrada directamente) ─────────────
# BASE DE LEADS REAIS — gerado automaticamente
# 239 leads de Base_Geral_Leads_Horeca + Alvos_TiPiedade

DB_NUNO = [
  {"n":"Copenhagen Coffee Lab (Cais do Sodré)","t":"Café de Especialidade & Brunch","m":"Praça de São Paulo 4, 1200-428 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Café de Especialidade & Brunch","gancho":"Público internacional de alta rotação. Focar no preço por dose que permite margem muito alta.","zona":"Lisboa (Cais Sodré)"},
  {"n":"Gleba (Amoreiras)","t":"Padaria Artesanal & Gourmet","m":"Avenida Engenheiro Duarte Pacheco, Amoreiras Shopping Center, Loja 1010, 1070-103 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Padaria Artesanal & Gourmet","gancho":"Compradores de elevado poder de compra. Venda por impulso para o lanche das famílias.","zona":"Lisboa (Amoreiras)"},
  {"n":"Choupana Caffe","t":"Pastelaria Premium & Brunch","m":"Avenida da República 25A, 1050-186 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Pastelaria Premium & Brunch","gancho":"Local de imensa rotação. Ter o Pão de Ló Ti Piedade em exposição rústica atrai lanches de grupo.","zona":"Lisboa (Avenidas Novas)"},
  {"n":"Tartine","t":"Pastelaria & Padaria Fina","m":"Rua Serpa Pinto 15A, 1200-426 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Pastelaria & Padaria Fina","gancho":"Junto ao Teatro de São Carlos. Foco na altíssima qualidade técnica do pão de ló Ti Piedade.","zona":"Lisboa (Chiado)"},
  {"n":"Fabrica Coffee Roasters (Rua da Flores)","t":"Café de Especialidade","m":"Rua das Flores 63, 1200-193 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Café de Especialidade","gancho":"Harmonização com cafés finos. Foco em turistas e nómadas digitais com ticket de compra alto.","zona":"Lisboa (Baixa)"},
  {"n":"Leitaria da Quinta do Paço (Graça)","t":"Pastelaria Premium","m":"Rua da Graça 90, 1170-170 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Pastelaria Premium","gancho":"Venda no balcão de lanches de fim de semana. Pão de ló fofo como o acompanhamento ideal.","zona":"Lisboa (Graça)"},
  {"n":"Gleba (Mercado da Vila)","t":"Padaria Artesanal & Gourmet","m":"Rua Padre Moisés da Silva, 2750-437 Cascais","tel":"—","email":"—","p":"Alta","tCliente":"Padaria Artesanal & Gourmet","gancho":"Mercado movimentado. Pão de ló premium para a sobremesa de domingo de clientes exigentes.","zona":"Cascais"},
  {"n":"The Millstone Sourdough","t":"Padaria Artesanal Premium","m":"Rua Nova da Alfarrobeira 11, 2750-449 Cascais","tel":"—","email":"—","p":"Alta","tCliente":"Padaria Artesanal Premium","gancho":"Processo rústico. Poupa-lhes fabrico mantendo o padrão artesanal limpo e sem aditivos.","zona":"Cascais"},
  {"n":"Lulu - Specialty Coffee & Brunch","t":"Café de Especialidade & Brunch","m":"Avenida Valbom 12, 2750-508 Cascais","tel":"—","email":"—","p":"Alta","tCliente":"Café de Especialidade & Brunch","gancho":"Brunch sofisticado. Apresentar o conceito de fatia de pão de ló com manteiga salgada artesanal.","zona":"Cascais"},
  {"n":"Local Healthy Kitchen (Cascais)","t":"Brunch & Comida Saudável","m":"Rua Padre Moisés da Silva, Mercado da Vila, 2750-437 Cascais","tel":"—","email":"—","p":"Alta","tCliente":"Brunch & Comida Saudável","gancho":"Propor pão de ló artesanal como doce rústico e natural, isento de corantes industriais.","zona":"Cascais"},
  {"n":"Garrafeira Imperial do Estoril","t":"Garrafeira & Gourmet Deli","m":"Avenida Saboia 320, 2765-277 Estoril","tel":"—","email":"—","p":"Alta","tCliente":"Garrafeira & Gourmet Deli","gancho":"Combinação ideal com vinhos generosos (Porto, Madeira, Carcavelos). Venda do bolo inteiro premium.","zona":"Estoril"},
  {"n":"Café Saudade","t":"Casa de Chá & Café de Charme","m":"Avenida Miguel Bombarda 6, 2710-590 Sintra","tel":"—","email":"—","p":"Alta","tCliente":"Casa de Chá & Café de Charme","gancho":"Lanches de charme. Reforça o menu de chás ingleses e infusões com pão de ló tradicional.","zona":"Sintra"},
  {"n":"Garrafeira de Sintra","t":"Garrafeira & Gourmet Deli","m":"Rua Consiglieri Pedroso 11, 2710-550 Sintra","tel":"—","email":"—","p":"Alta","tCliente":"Garrafeira & Gourmet Deli","gancho":"Venda casada com vinhos generosos regionais. Excelente opção de presente regional para turismo.","zona":"Sintra"},
  {"n":"Sintra In Love","t":"Cafetaria Gourmet & Regional","m":"Rua das Padarias 12, 2710-603 Sintra","tel":"—","email":"—","p":"Alta","tCliente":"Cafetaria Gourmet & Regional","gancho":"Centro histórico. Turistas à procura de doçaria portuguesa real. Vender a história do Ti Piedade.","zona":"Sintra"},
  {"n":"Aromas de Sintra","t":"Casa de Chá & Pastelaria Fina","m":"Largo Dr. Virgilio Horta 5, 2710-501 Sintra","tel":"—","email":"—","p":"Alta","tCliente":"Casa de Chá & Pastelaria Fina","gancho":"Excelente para o lanche da tarde. O pão de ló de qualidade mantém-se incrivelmente fresco na vitrine.","zona":"Sintra"},
  {"n":"Frade dos Mares","t":"Restaurante (Médio/Alto Padrão)","m":"Av. Dom Carlos I 55, 1200-109 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Cozinha tradicional refinada. Sobremesa de altíssimo nível, sempre disponível (congelada) com quebra zero.","zona":"Lisboa"},
  {"n":"A Casa do Bacalhau","t":"Restaurante (Médio/Alto Padrão)","m":"Rua do Grilo 54, 1900-706 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Parceiro tradicional perfeito. Descongelar no dia conforme as reservas garante controle total de custos.","zona":"Lisboa"},
  {"n":"Restaurante Solar dos Presuntos","t":"Restaurante (Médio/Alto Padrão)","m":"Rua das Portas de Santo Antão 150, 1150-269 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Instituição tradicional. O Ti Piedade oferece uma consistência e sabor conventual que os clientes exigem.","zona":"Lisboa"},
  {"n":"O Javali","t":"Restaurante Tradicional de Caça & Grelhados","m":"Rua de São Bento 344, 1200-822 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional de Caça & Grelhados","gancho":"Menu rústico. Sugerir o pão de ló com uma redução de frutos silvestres ou um licor local.","zona":"Lisboa"},
  {"n":"Faz Figura","t":"Restaurante (Médio/Alto Padrão / Vista Rio)","m":"Rua do Paraíso 15B, 1100-396 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão / Vista Rio)","gancho":"Jantares executivos e eventos. O formato congelado permite ter uma excelente sobremesa sem desperdícios diários.","zona":"Lisboa"},
  {"n":"Páteo do Guincho","t":"Restaurante (Médio/Alto Padrão)","m":"Estrada do Guincho, 2750-642 Cascais","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Sobremesa rápida de alta margem: fatia de pão de ló morna com bola de gelado artesanal de nata.","zona":"Cascais"},
  {"n":"Polvo Vadio","t":"Restaurante de Peixe & Marisco","m":"Rua Afonso Sanches 26, 2750-282 Cascais","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante de Peixe & Marisco","gancho":"Foco no porcionamento rápido. O pão de ló fatiado e pronto a servir após o marisco é um sucesso de conforto.","zona":"Cascais"},
  {"n":"Tacho Real","t":"Restaurante (Médio/Alto Padrão)","m":"Rua do Ferraria 4, 2710-590 Sintra","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Em pleno centro histórico. Custo fixo por dose e rapidez de serviço para as mesas de grupos turísticos.","zona":"Sintra"},
  {"n":"Restaurante Curral dos Caprinos","t":"Restaurante Regional Premium","m":"Rua da Capela 13, Cabriz, 2710-539 Sintra","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Regional Premium","gancho":"Pratos pesados de forno. O Pão de Ló Ti Piedade como a sobremesa leve, fofa e tradicional perfeita para fechar.","zona":"Sintra"},
  {"n":"Restaurante Lawrence's","t":"Restaurante de Hotel Histórico","m":"Rua Consiglieri Pedroso 38, 2710-550 Sintra","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante de Hotel Histórico","gancho":"Hotel mais antigo da Península Ibérica. O apelo histórico do Ti Piedade encaixa na narrativa de requinte histórico.","zona":"Sintra"},
  {"n":"Augusto Lisboa","t":"Café & Brunch Premium","m":"Rua de Santa Marinha 26, 1100-491 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Café & Brunch Premium","gancho":"Brunch muito concorrido na zona histórica. Perfeito para introduzir uma opção de pão de ló tostado com mel artesanal.","zona":"Lisboa"},
  {"n":"Tacho do Pescador","t":"Restaurante Tradicional Premium","m":"Rua do Pimenta 15, 1990-254 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Junto à FIL. Público executivo e feiras. Sobremesa tradicional portuguesa rápida de porcionar com margem alta.","zona":"Lisboa (Parque das Nações)"},
  {"n":"Gleba (Parque das Nações)","t":"Padaria Artesanal & Gourmet","m":"Alameda dos Oceanos, Lote 1.06.1.1 - Loja 2, 1990-207 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Padaria Artesanal & Gourmet","gancho":"Zona executiva e residencial premium. Foco na qualidade limpa do bolo Ti Piedade para as lancheiras familiares.","zona":"Lisboa (Pq. Nações)"},
  {"n":"Copenhagen Coffee Lab (Parque das Nações)","t":"Café de Especialidade & Brunch","m":"Rua do Bojador 47, 1990-048 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Café de Especialidade & Brunch","gancho":"Executivos e trabalhadores remotos. Lanche de qualidade a acompanhar o café com margem de lucro muito alta.","zona":"Lisboa (Pq. Nações)"},
  {"n":"Angra Gatti","t":"Italiano","m":"—","tel":"—","email":"geral@angragatti.pt","p":"Alta","tCliente":"Italiano","gancho":"Lead identificado pela equipa comercial — Italiano em Praia Grande.","zona":"Praia Grande"},
  {"n":"Arribas Terrace","t":"Hotel/restaurante","m":"—","tel":"—","email":"reservas@arribashotel.com","p":"Alta","tCliente":"Hotel/restaurante","gancho":"Lead identificado pela equipa comercial — Hotel/restaurante em Hotel Arribas, Praia Grande.","zona":"Hotel Arribas, Praia Grande"},
  {"n":"Baía do Peixe Cascais","t":"Restaurante","m":"—","tel":"—","email":"reservas@baiadopeixe.com","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Baía de Cascais.","zona":"Baía de Cascais"},
  {"n":"Bar do Fundo","t":"Restaurante premium","m":"—","tel":"—","email":"info@bardofundo.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Praia Grande.","zona":"Praia Grande"},
  {"n":"Bar do Guincho","t":"Beach restaurant","m":"—","tel":"—","email":"geral@bardoguincho.pt","p":"Alta","tCliente":"Beach restaurant","gancho":"Lead identificado pela equipa comercial — Beach restaurant em Praia do Guincho.","zona":"Praia do Guincho"},
  {"n":"Beira Mar","t":"Restaurante costeiro","m":"—","tel":"—","email":"geral@restaurantebeiramar.pt","p":"Alta","tCliente":"Restaurante costeiro","gancho":"Lead identificado pela equipa comercial — Restaurante costeiro em Cascais.","zona":"Cascais"},
  {"n":"Capricciosa Cascais","t":"Restaurante italiano","m":"—","tel":"—","email":"TheFork","p":"Alta","tCliente":"Restaurante italiano","gancho":"Lead identificado pela equipa comercial — Restaurante italiano em Cascais.","zona":"Cascais"},
  {"n":"Fortaleza do Guincho","t":"Hotel/restaurante Michelin","m":"—","tel":"—","email":"info@fortalezadoguincho.com","p":"Alta","tCliente":"Hotel/restaurante Michelin","gancho":"Lead identificado pela equipa comercial — Hotel/restaurante Michelin em Estrada do Guincho.","zona":"Estrada do Guincho"},
  {"n":"Fuji Sushi & Steak","t":"Restaurante internacional","m":"—","tel":"—","email":"TheFork","p":"Alta","tCliente":"Restaurante internacional","gancho":"Lead identificado pela equipa comercial — Restaurante internacional em Cascais.","zona":"Cascais"},
  {"n":"Furnas do Guincho","t":"Restaurante premium","m":"—","tel":"—","email":"geral@furnasdoguincho.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Cascais.","zona":"Cascais"},
  {"n":"Hífen","t":"Restaurante contemporâneo","m":"—","tel":"—","email":"reservashifen@gmail.com","p":"Alta","tCliente":"Restaurante contemporâneo","gancho":"Lead identificado pela equipa comercial — Restaurante contemporâneo em Marina de Cascais.","zona":"Marina de Cascais"},
  {"n":"House of Wonders","t":"Restaurante / brunch","m":"—","tel":"—","email":"info@houseofwonders.pt","p":"Alta","tCliente":"Restaurante / brunch","gancho":"Lead identificado pela equipa comercial — Restaurante / brunch em Rua da Misericórdia, Cascais.","zona":"Rua da Misericórdia, Cascais"},
  {"n":"Marisco na Praça","t":"Restaurante","m":"—","tel":"—","email":"geral@marisconapraca.pt","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Mercado da Vila Cascais.","zona":"Mercado da Vila Cascais"},
  {"n":"Monte Mar","t":"Restaurante premium","m":"—","tel":"—","email":"montemar@montemar.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Cascais.","zona":"Cascais"},
  {"n":"Monte Mar Cascais","t":"Restaurante premium","m":"—","tel":"—","email":"geral@montemar.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Boca do Inferno, Cascais.","zona":"Boca do Inferno, Cascais"},
  {"n":"Moinho Dom Quixote","t":"Restaurante panorâmico","m":"—","tel":"—","email":"info@moinhodomquixote.pt","p":"Alta","tCliente":"Restaurante panorâmico","gancho":"Lead identificado pela equipa comercial — Restaurante panorâmico em Colares.","zona":"Colares"},
  {"n":"A Toca do Júlio","t":"Restaurante tradicional","m":"—","tel":"—","email":"geral@atocadojulio.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Colares.","zona":"Colares"},
  {"n":"Restaurante Azenhas do Mar","t":"Restaurante premium","m":"—","tel":"—","email":"reservas@azenhasdomar.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Azenhas do Mar.","zona":"Azenhas do Mar"},
  {"n":"Água e Sal","t":"Restaurante contemporâneo","m":"—","tel":"—","email":"geral@aguaesal.pt","p":"Alta","tCliente":"Restaurante contemporâneo","gancho":"Lead identificado pela equipa comercial — Restaurante contemporâneo em Azenhas do Mar.","zona":"Azenhas do Mar"},
  {"n":"Adega das Azenhas","t":"Restaurante tradicional","m":"—","tel":"—","email":"geral@adegadasazenhas.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Azenhas do Mar.","zona":"Azenhas do Mar"},
  {"n":"Terrace Restaurant","t":"Restaurante/hotel","m":"—","tel":"—","email":"reservas@azenhasdohotel.pt","p":"Alta","tCliente":"Restaurante/hotel","gancho":"Lead identificado pela equipa comercial — Restaurante/hotel em Azenhas do Mar.","zona":"Azenhas do Mar"},
  {"n":"Kailua Fonte da Telha","t":"Beach club","m":"—","tel":"—","email":"geral@kailua.pt","p":"Alta","tCliente":"Beach club","gancho":"Lead identificado pela equipa comercial — Beach club em Fonte da Telha.","zona":"Fonte da Telha"},
  {"n":"Oh! Vargas","t":"Restaurante contemporâneo","m":"—","tel":"—","email":"geral@ohvargas.pt","p":"Alta","tCliente":"Restaurante contemporâneo","gancho":"Lead identificado pela equipa comercial — Restaurante contemporâneo em Santarém Centro.","zona":"Santarém Centro"},
  {"n":"Taberna Ó Balcão","t":"Restaurante premium/tradicional","m":"—","tel":"—","email":"reservas@obalcao.pt","p":"Alta","tCliente":"Restaurante premium/tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante premium/tradicional em Santarém.","zona":"Santarém"},
  {"n":"Tascá","t":"Restaurante moderno","m":"—","tel":"—","email":"geral@tasca.pt","p":"Alta","tCliente":"Restaurante moderno","gancho":"Lead identificado pela equipa comercial — Restaurante moderno em Santarém.","zona":"Santarém"},
  {"n":"O Quinzena","t":"Restaurante tradicional","m":"—","tel":"—","email":"geral@oquinzena.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Santarém.","zona":"Santarém"},
  {"n":"Restaurante São Domingos","t":"Restaurante eventos","m":"—","tel":"—","email":"geral@saodomingos.pt","p":"Alta","tCliente":"Restaurante eventos","gancho":"Lead identificado pela equipa comercial — Restaurante eventos em Santarém.","zona":"Santarém"},
  {"n":"Dom Papinhas","t":"Restaurante familiar","m":"—","tel":"—","email":"geral@dompapinhas.pt","p":"Alta","tCliente":"Restaurante familiar","gancho":"Lead identificado pela equipa comercial — Restaurante familiar em Santarém.","zona":"Santarém"},
  {"n":"Adega do Avô","t":"Restaurante regional","m":"—","tel":"—","email":"geral@adegadoavo.pt","p":"Alta","tCliente":"Restaurante regional","gancho":"Lead identificado pela equipa comercial — Restaurante regional em Santarém.","zona":"Santarém"},
  {"n":"Pátio da Graça","t":"Restaurante contemporâneo","m":"—","tel":"—","email":"reservas@patiodagraca.pt","p":"Alta","tCliente":"Restaurante contemporâneo","gancho":"Lead identificado pela equipa comercial — Restaurante contemporâneo em Santarém.","zona":"Santarém"},
  {"n":"Cantinho do Avillez Porto","t":"Restaurante premium","m":"—","tel":"—","email":"porto@cantinhodoavillez.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Rua Mouzinho da Silveira.","zona":"Rua Mouzinho da Silveira"},
  {"n":"Taberna dos Mercadores","t":"Restaurante turístico","m":"—","tel":"—","email":"reservas@tabernadosmercadores.pt","p":"Alta","tCliente":"Restaurante turístico","gancho":"Lead identificado pela equipa comercial — Restaurante turístico em Ribeira.","zona":"Ribeira"},
  {"n":"Adega São Nicolau","t":"Restaurante tradicional","m":"—","tel":"—","email":"geral@adegasaonicolau.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Ribeira.","zona":"Ribeira"},
  {"n":"Caféína","t":"Restaurante contemporâneo","m":"—","tel":"—","email":"geral@cafeina.pt","p":"Alta","tCliente":"Restaurante contemporâneo","gancho":"Lead identificado pela equipa comercial — Restaurante contemporâneo em Foz do Douro.","zona":"Foz do Douro"},
  {"n":"Wish Restaurante","t":"Restaurante premium","m":"—","tel":"—","email":"reservas@wish.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Boavista.","zona":"Boavista"},
  {"n":"Vinum","t":"Restaurante premium vínico","m":"—","tel":"—","email":"reservas@vinumatgrahams.com","p":"Alta","tCliente":"Restaurante premium vínico","gancho":"Lead identificado pela equipa comercial — Restaurante premium vínico em Caves Graham’s.","zona":"Caves Graham’s"},
  {"n":"Stramuntana","t":"Restaurante premium","m":"—","tel":"—","email":"reservas@stramuntana.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Canidelo.","zona":"Canidelo"},
  {"n":"Muchaxo Restaurant","t":"Hotel/restaurante","m":"—","tel":"—","email":"recepcao@muchaxo.com","p":"Alta","tCliente":"Hotel/restaurante","gancho":"Lead identificado pela equipa comercial — Hotel/restaurante em Praia do Guincho.","zona":"Praia do Guincho"},
  {"n":"O Faroleiro","t":"Restaurante português costeiro","m":"—","tel":"—","email":"info@faroleiro.com","p":"Alta","tCliente":"Restaurante português costeiro","gancho":"Lead identificado pela equipa comercial — Restaurante português costeiro em Guincho.","zona":"Guincho"},
  {"n":"Palm Tree Pub & Restaurant","t":"Restaurante/bar","m":"—","tel":"—","email":"info@palmtree.pt","p":"Alta","tCliente":"Restaurante/bar","gancho":"Lead identificado pela equipa comercial — Restaurante/bar em Cascais.","zona":"Cascais"},
  {"n":"Panorama Beach Club","t":"Beach Club / Restaurante","m":"—","tel":"—","email":"panoramaguincho@villacolletion.pt","p":"Alta","tCliente":"Beach Club / Restaurante","gancho":"Lead identificado pela equipa comercial — Beach Club / Restaurante em Guincho.","zona":"Guincho"},
  {"n":"Pigalle Smash Bistro","t":"Bistro americano","m":"—","tel":"—","email":"TheFork","p":"Alta","tCliente":"Bistro americano","gancho":"Lead identificado pela equipa comercial — Bistro americano em Cascais.","zona":"Cascais"},
  {"n":"Porto Santa Maria","t":"Restaurante peixe/marisco premium","m":"—","tel":"—","email":"reservas@portosantamaria.com · 214 879 450","p":"Alta","tCliente":"Restaurante peixe/marisco premium","gancho":"Lead identificado pela equipa comercial — Restaurante peixe/marisco premium em Guincho.","zona":"Guincho"},
  {"n":"Restaurante Nortada","t":"Restaurante","m":"—","tel":"—","email":"geral@restaurantenortada.com","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Praia Grande, Sintra.","zona":"Praia Grande, Sintra"},
  {"n":"Santini Cascais","t":"Gelataria premium","m":"—","tel":"—","email":"cascais@santini.pt","p":"Alta","tCliente":"Gelataria premium","gancho":"Lead identificado pela equipa comercial — Gelataria premium em Baía de Cascais.","zona":"Baía de Cascais"},
  {"n":"Spot by Fortaleza do Guincho","t":"Restaurante (Michelin)","m":"—","tel":"—","email":"restaurante@guinchotel.pt","p":"Alta","tCliente":"Restaurante (Michelin)","gancho":"Lead identificado pela equipa comercial — Restaurante (Michelin) em Guincho.","zona":"Guincho"},
  {"n":"Visconde da Luz","t":"Restaurante tradicional","m":"—","tel":"—","email":"geral@viscondedaluz.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Cascais Centro.","zona":"Cascais Centro"},
  {"n":"Restaurante Azenhas do Mar","t":"Restaurante marisco/peixe com vista mar","m":"—","tel":"—","email":"Google Maps / site próprio","p":"Alta","tCliente":"Restaurante marisco/peixe com vista mar","gancho":"Lead identificado pela equipa comercial — Restaurante marisco/peixe com vista mar em Azenhas do Mar.","zona":"Azenhas do Mar"},
  {"n":"Opíparo","t":"Restaurante cozinha portuguesa reinventada","m":"—","tel":"—","email":"Google Maps","p":"Alta","tCliente":"Restaurante cozinha portuguesa reinventada","gancho":"Lead identificado pela equipa comercial — Restaurante cozinha portuguesa reinventada em Azenhas do Mar.","zona":"Azenhas do Mar"},
  {"n":"Hamburgueria do Maçãs","t":"Restaurante informal","m":"—","tel":"—","email":"Google Maps","p":"Alta","tCliente":"Restaurante informal","gancho":"Lead identificado pela equipa comercial — Restaurante informal em Azenhas do Mar.","zona":"Azenhas do Mar"},
  {"n":"Cervejaria Boa Esperança","t":"Cervejaria","m":"Avenida Gomes Pereira 3 A Loja","tel":"—","email":"boaesperanca3a@gmail.com","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
  {"n":"Browers Beato","t":"Cervejaria","m":"Tv. Grilo 1","tel":"—","email":"beato@thebrowerscompany.com","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
  {"n":"Dote - Cervejaria Moderna","t":"Cervejaria","m":"Alvalade, Av. Republica, Parque Nações, Barata Salgueiro, Odivelas","tel":"—","email":"dote@dote.pt","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
  {"n":"Edmundo","t":"Cervejaria","m":"Avenida Gomes Pereira, 1 - Estrada de Benfica","tel":"—","email":"restaurante.edmundo@gmail.com","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
  {"n":"Farol","t":"Cervejaria","m":"Saldanha","tel":"—","email":"tcardoso1990@gmail.com","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
  {"n":"Gambrinus","t":"Cervejaria","m":"Rua das Portas de Santo Antão, nº 23","tel":"—","email":"info@gambrinuslisboa.com","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
  {"n":"Cervejaria Liberdade","t":"Cervejaria","m":"Rua Castilho, 14","tel":"—","email":"info@avliberdade.com; cervejaria.avenidaliberdade@tivoli-hotels.com","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
  {"n":"Pinóquio","t":"Cervejaria","m":"Praça dos Restauradores 79 80","tel":"—","email":"geral@restaurantepinoquio.pt","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
  {"n":"Portugália","t":"Cervejaria","m":"Almirante Reis","tel":"—","email":"pareis@portugalia.pt","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
  {"n":"Cervejaria Ramiro","t":"Cervejaria","m":"Av. Alm. Reis 1 H","tel":"—","email":"ramiro@cervejariaramiro.pt","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
  {"n":"Cervejaria Sem Vergonha","t":"Cervejaria","m":"Tv. de Santa Quitéria 38 D","tel":"—","email":"cervejariasemvergonha@gmail.com","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
  {"n":"Cervejaria Trindade","t":"Cervejaria","m":"Rua Nova da Trindade, nº 20 C","tel":"—","email":"ct.chiado@cervejariatrindade.pt","p":"Alta","tCliente":"Cervejaria Lisboa","gancho":"Grande volume de refeições — pão de ló como sobremesa clássica portuguesa de fecho, sem complexidade operacional.","zona":"Lisboa"},
]

DB_JOAO = [
  {"n":"Lupis Coffee Shop","t":"Café de Especialidade & Brunch","m":"Rua de Marvila 42, 1950-197 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Café de Especialidade & Brunch","gancho":"Público artístico e de design de Marvila. O aspeto rústico e autêntico do pão de ló Ti Piedade tem enorme apelo visual.","zona":"Lisboa (Marvila)"},
  {"n":"Restaurante Dinastia Tang","t":"Restaurante (Médio/Alto Padrão)","m":"Rua do Açúcar 107, 1950-006 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Encontro cultural. Embora especializado, procuram sobremesas portuguesas tradicionais de altíssima qualidade para os clientes locais.","zona":"Lisboa (Marvila)"},
  {"n":"Pão do Beco","t":"Padaria Artesanal","m":"Rua Capitão Leitão 72A, 2800-135 Almada","tel":"—","email":"—","p":"Alta","tCliente":"Padaria Artesanal","gancho":"Foco na autenticidade da fermentação natural. O Ti Piedade como o pão de ló sem químicos industriais.","zona":"Almada"},
  {"n":"Under The Cover / O Pão de Ló","t":"Café & Brunch Moderno","m":"Rua Capitão Leitão 54, 2800-135 Almada","tel":"—","email":"—","p":"Alta","tCliente":"Café & Brunch Moderno","gancho":"Cafetaria moderna com foco em lanches e pastelaria cuidada. Excelente para fatias e meias-fatias.","zona":"Almada"},
  {"n":"Mundet Factory","t":"Restaurante & Brunch Premium","m":"Avenida Metalúrgica Augusto de Castro 1, 2840-515 Seixal","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante & Brunch Premium","gancho":"Restaurante moderno, muito movimentado. Foco no controlo estrito de quebras e custos de stock via formato congelado.","zona":"Seixal"},
  {"n":"Let's Coffe","t":"Café de Especialidade","m":"Praça dos Mártires da Pátria 12, 2840-496 Seixal","tel":"—","email":"—","p":"Alta","tCliente":"Café de Especialidade","gancho":"Café gourmet junto à baía do Seixal. Sobremesa tradicional portuguesa em harmonia com café premium.","zona":"Seixal"},
  {"n":"Pão com Alma","t":"Padaria Artesanal","m":"Rua Miguel Bombarda 124, 2830-355 Barreiro","tel":"—","email":"—","p":"Alta","tCliente":"Padaria Artesanal","gancho":"Identidade e tradição no fabrico. O Pão de Ló Ti Piedade partilha dos mesmos princípios artesanais.","zona":"Barreiro"},
  {"n":"Restaurante Taberna do Manel","t":"Restaurante Tradicional Premium","m":"Rua Conselheiro Joaquim António d'Aguiar 32, 2830-334 Barreiro","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Clássico da gastronomia do Barreiro. Sobremesa de altíssima consistência, sem preocupações com validade curta de pastelaria diária.","zona":"Barreiro"},
  {"n":"Dr. Bernard","t":"Restaurante de Praia & Brunch","m":"Avenida General Humberto Delgado, Praia do Tarquínio-Paraíso, 2825-366 Costa da Caparica","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante de Praia & Brunch","gancho":"Público moderno e internacional. Fatia de pão de ló servida com coberturas de frutos ou gelado de limão no brunch.","zona":"Costa da Caparica"},
  {"n":"Palms","t":"Restaurante de Praia & Brunch Premium","m":"Praia do CDS, Apoio 9, Avenida General Humberto Delgado, 2825-366 Costa da Caparica","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante de Praia & Brunch Premium","gancho":"Público cosmopolita. Apresentar o Pão de Ló como o autêntico bolo tradicional português para brunch ou lanche na praia.","zona":"Costa da Caparica"},
  {"n":"Restaurante O Sentido do Mar","t":"Restaurante de Peixe (Médio/Alto Padrão)","m":"Praia do Tarquínio, Apoio de Praia 4, 2825-366 Costa da Caparica","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante de Peixe (Médio/Alto Padrão)","gancho":"Restaurante sofisticado com vista para o mar. Sobremesa tradicional premium pronta a fatiar, de custo controlado e quebra zero.","zona":"Costa da Caparica"},
  {"n":"Restaurante Carolina do Aires","t":"Restaurante Tradicional Premium","m":"Avenida General Humberto Delgado 10, 2825-366 Costa da Caparica","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Restaurante clássico com enorme tradição de peixe. Sobremesa portuguesa tradicional que assegura fidelidade ao sabor.","zona":"Costa da Caparica"},
  {"n":"Ponto Final","t":"Restaurante (Médio/Alto Padrão)","m":"Rua do Ginjal 72, 2800-285 Almada","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Vistas incríveis, alta rotação de estrangeiros e locais. Sobremesa ultra-rápida na mesa sem desperdício de stock.","zona":"Almada (Cacilhas)"},
  {"n":"Atira-te ao Rio","t":"Restaurante (Médio/Alto Padrão)","m":"Rua do Ginjal 69, 2800-285 Almada","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Cozinha requintada portuguesa e contemporânea. Foco no sabor rústico e excelente aspeto da fatia no prato.","zona":"Almada (Cacilhas)"},
  {"n":"Restaurante Cabrinha","t":"Restaurante Tradicional Premium","m":"Rua de Cacilhas 33, 2800-135 Almada","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Especialistas em peixe grelhado e marisco. Fecho de refeição com pão de ló húmido tradicional, com custo fixo por dose.","zona":"Almada"},
  {"n":"Amarra ao Cais","t":"Restaurante (Médio/Alto Padrão)","m":"Jardim do Rio, Cais do Ginjal, 2800-285 Almada","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Localização premium. Sobremesas rústicas rápidas e fáceis de marginar mais de 65% na carta.","zona":"Almada"},
  {"n":"O Pescador","t":"Restaurante Tradicional Premium","m":"Avenida D. Carlos I 12, 2840-515 Seixal","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Foco na clientela local de domingo. Pão de ló fatiado na hora com um fio de doce de ovos (receita fácil de assinar).","zona":"Seixal"},
  {"n":"O Farribas","t":"Restaurante (Médio/Alto Padrão)","m":"Avenida Luísa Todi 244, 2900-452 Setúbal","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Na principal avenida. Doçaria portuguesa autêntica fornecida congelada para controlo total de quebras pós-refeição.","zona":"Setúbal"},
  {"n":"Restaurante Novo 10","t":"Restaurante Tradicional Premium","m":"Avenida Luísa Todi 220, 2900-452 Setúbal","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Especialistas em peixe assado. O Pão de Ló Ti Piedade como a sobremesa ideal de alto valor percebido pelos clientes.","zona":"Setúbal"},
  {"n":"Casa do Peixe (Setúbal)","t":"Restaurante (Médio/Alto Padrão)","m":"Rua Praia da Saúde 10, 2900-572 Setúbal","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Restaurante à beira d'água. Oferecer sobremesa de pão de ló com consistência garantida e zero perdas semanais.","zona":"Setúbal"},
  {"n":"Casa das Tortas de Azeitão","t":"Pastelaria Gourmet & Chá","m":"Praça da República 1, 2925-520 Azeitão","tel":"—","email":"—","p":"Alta","tCliente":"Pastelaria Gourmet & Chá","gancho":"Ponto de passagem turístico. O Pão de Ló Ti Piedade junta-se à oferta de doçaria regional como a opção fofa sem cremes.","zona":"Azeitão"},
  {"n":"Garrafeira de Azeitão","t":"Garrafeira & Gourmet Deli","m":"Avenida 25 de Abril 12, 2925-501 Azeitão","tel":"—","email":"—","p":"Alta","tCliente":"Garrafeira & Gourmet Deli","gancho":"Venda casada com Moscatel de Setúbal. O Pão de Ló é o casamento perfeito para degustações premium na garrafeira.","zona":"Azeitão"},
  {"n":"O Quintal","t":"Restaurante (Médio/Alto Padrão)","m":"Rua José Augusto Coelho 82, 2925-538 Azeitão","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Cozinha tradicional refinada. Foco na margem absoluta que o produto congelado traz, com 0% de quebra alimentar.","zona":"Azeitão"},
  {"n":"Restaurante Ribeirinha do Sado","t":"Restaurante Tradicional Premium","m":"Avenida Luísa Todi 34, 2900-450 Setúbal","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Paragem clássica para peixe. Sobremesa tradicional pronta a fatiar, de custo controlado e quebra zero.","zona":"Setúbal"},
  {"n":"O Alfeite","t":"Restaurante Tradicional Premium","m":"Rua de Serpa Pinto 4, 2800-205 Almada","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Restaurante típico de almoços executivos. Solução congelada garante quebras zero em dias de menor afluência.","zona":"Almada"},
  {"n":"Lisboa à Vista","t":"Restaurante (Médio/Alto Padrão)","m":"Av. Metalúrgica Augusto de Castro, Baía do Seixal, 2840-515 Seixal","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Restaurante num barco histórico. Doçaria portuguesa tradicional com aspeto rústico fantástico para acompanhar vinhos generosos.","zona":"Seixal"},
  {"n":"Heim Cafe","t":"Brunch & Lanches","m":"Rua de Santos-O-Velho 4, 1200-109 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Brunch & Lanches","gancho":"O queridinho do brunch. Inserir fatia de Pão de Ló tostada com toppings no menu diário.","zona":"Lisboa (Campo de Ourique)"},
  {"n":"Dear Breakfast (Saldanha)","t":"Brunch Club & Café","m":"Rua Eng. Vieira da Silva 8, 1050-105 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Brunch Club & Café","gancho":"Brunch executivo e sofisticado. Pão de ló Ti Piedade como opção portuguesa no menu.","zona":"Lisboa (Saldanha)"},
  {"n":"Cajú","t":"Brunch & Healthy Food","m":"Rua da Alegria 73, 1250-182 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Brunch & Healthy Food","gancho":"Público internacional. Oferecer a experiência do autêntico pão de ló tradicional fofo.","zona":"Lisboa (Príncipe Real)"},
  {"n":"Isco (Alvalade)","t":"Padaria Artesanal & Bistrô","m":"Rua do Arco do Carvalhão 244, 1350-026 Lisboa","tel":"—","email":"—","p":"Alta","tCliente":"Padaria Artesanal & Bistrô","gancho":"Artesanal puro. Pão de ló premium poupa-lhes tempo de pastelaria mantendo a excelência.","zona":"Lisboa (Alvalade)"},
  {"n":"Borda D`Agua","t":"Restaurante Praia","m":"—","tel":"—","email":"geral@bordadagua.com.pt","p":"Alta","tCliente":"Restaurante Praia","gancho":"Lead identificado pela equipa comercial — Restaurante Praia em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Gordos Caparica","t":"Restaurante","m":"—","tel":"—","email":"gordos.caparica@gmail.com","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Restaurante Leblon","t":"Restaurante","m":"—","tel":"—","email":"leblonsaojoao@gmail.com","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Coco Beach","t":"Restaurante","m":"—","tel":"—","email":"Cocobeach.caparica@gmail.com","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Praia - Sea, Salt & Pepper","t":"Restaurante","m":"—","tel":"—","email":"reservas@apraia.pt","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Fortuna Tacos & Burger","t":"Restaurante","m":"—","tel":"—","email":"fortuna.tacos@gmail.com","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Zama Beach Club","t":"Restaurante","m":"—","tel":"—","email":"zama.costadacaparica@gmail.com","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"O Xéxéxé","t":"Restaurante","m":"—","tel":"—","email":"oxexexedacosta@gmail.com","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"AHOY Surf & Snacks","t":"Restaurante","m":"—","tel":"—","email":"Ahoycafe@hotmail.com","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Clássico Beach Bar by Olivier","t":"Restaurante","m":"—","tel":"—","email":"classicobeachbar@olivier.pt","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"O Barbas","t":"Restaurante","m":"—","tel":"—","email":"restauranteobarbas@gmail.com","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Muse Brunch Café & Wine Bar","t":"Café & Wine Bar","m":"—","tel":"—","email":"info@musewinebar.pt","p":"Alta","tCliente":"Café & Wine Bar","gancho":"Lead identificado pela equipa comercial — Café & Wine Bar em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"O Mercado","t":"Restaurante tradicional","m":"—","tel":"—","email":"omercadocc@gmail.com","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Apeixonado","t":"Restaurante tradicional","m":"—","tel":"—","email":"apeixonado@ponteslda.com","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Ponto Final","t":"Restaurante turístico","m":"—","tel":"—","email":"reservas@pontofinal.pt","p":"Alta","tCliente":"Restaurante turístico","gancho":"Lead identificado pela equipa comercial — Restaurante turístico em Cacilhas.","zona":"Cacilhas"},
  {"n":"Atira-te ao Rio","t":"Restaurante premium","m":"—","tel":"—","email":"geral@atirateaorio.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Cacilhas.","zona":"Cacilhas"},
  {"n":"Solar Beirão","t":"Restaurante tradicional","m":"—","tel":"—","email":"geral@solarbeirao.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Almada.","zona":"Almada"},
  {"n":"O Martinho","t":"Restaurante","m":"—","tel":"—","email":"geral@omartinho.pt","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Praia da Princesa","t":"Beach restaurant","m":"—","tel":"—","email":"reservas@praiadaprincesa.pt","p":"Alta","tCliente":"Beach restaurant","gancho":"Lead identificado pela equipa comercial — Beach restaurant em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Dr. Bernard","t":"Restaurante/bar","m":"—","tel":"—","email":"geral@drbernard.pt","p":"Alta","tCliente":"Restaurante/bar","gancho":"Lead identificado pela equipa comercial — Restaurante/bar em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Sentido do Mar","t":"Restaurante","m":"—","tel":"—","email":"reservas@sentidodomar.pt","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Costa da Caparica.","zona":"Costa da Caparica"},
  {"n":"Atlas Leiria","t":"Restaurante/bar moderno","m":"—","tel":"—","email":"geral@atlasleiria.pt","p":"Alta","tCliente":"Restaurante/bar moderno","gancho":"Lead identificado pela equipa comercial — Restaurante/bar moderno em Leiria Centro.","zona":"Leiria Centro"},
  {"n":"Casinha Velha","t":"Restaurante premium","m":"—","tel":"—","email":"reservas@casinhavelha.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Leiria.","zona":"Leiria"},
  {"n":"Mata Bicho","t":"Restaurante contemporâneo","m":"—","tel":"—","email":"geral@matabicho.pt","p":"Alta","tCliente":"Restaurante contemporâneo","gancho":"Lead identificado pela equipa comercial — Restaurante contemporâneo em Leiria.","zona":"Leiria"},
  {"n":"Cervejaria João Gordo","t":"Restaurante tradicional","m":"—","tel":"—","email":"geral@joaogordo.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Leiria.","zona":"Leiria"},
  {"n":"Tromba Rija","t":"Restaurante grande volume","m":"—","tel":"—","email":"reservas@trombarija.pt","p":"Alta","tCliente":"Restaurante grande volume","gancho":"Lead identificado pela equipa comercial — Restaurante grande volume em Leiria.","zona":"Leiria"},
  {"n":"Malagueta Afrobar","t":"Restaurante moderno","m":"—","tel":"—","email":"geral@malagueta.pt","p":"Alta","tCliente":"Restaurante moderno","gancho":"Lead identificado pela equipa comercial — Restaurante moderno em Leiria.","zona":"Leiria"},
  {"n":"Taberna do Quinzena","t":"Restaurante português","m":"—","tel":"—","email":"geral@tabernadequinzena.pt","p":"Alta","tCliente":"Restaurante português","gancho":"Lead identificado pela equipa comercial — Restaurante português em Leiria.","zona":"Leiria"},
  {"n":"Tokio Sushi","t":"Restaurante moderno","m":"—","tel":"—","email":"reservas@tokiosushi.pt","p":"Alta","tCliente":"Restaurante moderno","gancho":"Lead identificado pela equipa comercial — Restaurante moderno em Leiria.","zona":"Leiria"},
  {"n":"Sauvage Gourmet","t":"Brunch / Gourmet (★ 4.8)","m":"—","tel":"—","email":"geral@sauvage-gourmet.pt","p":"Alta","tCliente":"Brunch / Gourmet (★ 4.8)","gancho":"Lead identificado pela equipa comercial — Brunch / Gourmet (★ 4.8) em Almada.","zona":"Almada"},
  {"n":"De Raiz no Museu (Chef Luís Calei)","t":"Fine dining / cozinha autoral","m":"—","tel":"—","email":"geral@de-raiz.pt","p":"Alta","tCliente":"Fine dining / cozinha autoral","gancho":"Lead identificado pela equipa comercial — Fine dining / cozinha autoral em Almada.","zona":"Almada"},
  {"n":"Soul Sushi – Japanese Fusion","t":"Restaurante fusão (★ 9.8 TheFork)","m":"—","tel":"—","email":"soulsushibar@gmail.com","p":"Alta","tCliente":"Restaurante fusão (★ 9.8 TheFork)","gancho":"Lead identificado pela equipa comercial — Restaurante fusão (★ 9.8 TheFork) em Almada.","zona":"Almada"},
  {"n":"Contrabando Mexican Food","t":"Restaurante temático","m":"—","tel":"—","email":"geral.contrabando@gmail.com","p":"Alta","tCliente":"Restaurante temático","gancho":"Lead identificado pela equipa comercial — Restaurante temático em Almada.","zona":"Almada"},
  {"n":"Pastelaria Condestável","t":"Pastelaria","m":"—","tel":"—","email":"pastcondestavel@gmail.com","p":"Alta","tCliente":"Pastelaria","gancho":"Lead identificado pela equipa comercial — Pastelaria em Almada.","zona":"Almada"},
]

DB_OSCAR = [
  {"n":"Oitoo","t":"Padaria Artesanal & Café","m":"Rua de Cedofeita 443, 4050-181 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Padaria Artesanal & Café","gancho":"Foco no público moderno e artesanal. Pão de ló fofo para lanche premium.","zona":"Porto"},
  {"n":"Masseira","t":"Padaria Artesanal Sourdough","m":"Rua de D. Manuel II 296, 4050-343 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Padaria Artesanal Sourdough","gancho":"Identidade rústica. Combina com a integridade do pão de ló artesanal.","zona":"Porto"},
  {"n":"Chá das Cinco","t":"Pastelaria Artesanal & Casa de Chá","m":"Praça da Alegria 50, 4000-027 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Pastelaria Artesanal & Casa de Chá","gancho":"O spot de eleição para bolos à fatia no Porto. Propor alternativa sem creme e tradicional.","zona":"Porto"},
  {"n":"Lazy Breakfast Club","t":"Brunch Club & Café","m":"Rua das Oliveiras 110, 4050-449 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Brunch Club & Café","gancho":"Inovação: sugerir fatia de pão de ló tostada com manteiga artesanal ou toppings doces.","zona":"Porto"},
  {"n":"Duquesa","t":"Brunch & Specialty Coffee","m":"Rua de Cedofeita 300, 4050-174 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Brunch & Specialty Coffee","gancho":"Público jovem e turistas. Perfeito para introduzir uma fatia com café de especialidade.","zona":"Porto"},
  {"n":"Nola Kitchen","t":"Cafetaria Saudável & Brunch","m":"Praça de D. Filipa de Lencastre 25, 4050-259 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Cafetaria Saudável & Brunch","gancho":"Argumentar a ausência de corantes e conservantes (ingredientes 100% limpos e tradicionais).","zona":"Porto"},
  {"n":"My Coffee Porto","t":"Café de Especialidade","m":"Escadas do Caminho Novo 11, 4050-554 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Café de Especialidade","gancho":"Vista deslumbrante, turistas. O Pão de Ló é a 'fatia portuguesa' ideal para acompanhar o café.","zona":"Porto"},
  {"n":"Do Norte Cafe by Hungry Biker","t":"Brunch & Cozy Cafe","m":"Rua do Almada 516, 4050-039 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Brunch & Cozy Cafe","gancho":"Ambiente rústico e acolhedor. Foco na porção individual do pão de ló para brunch.","zona":"Porto"},
  {"n":"Epoca Café","t":"Café de Especialidade & Lanches","m":"Rua do Rosário 22, 4050-522 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Café de Especialidade & Lanches","gancho":"Foco no público local premium que valoriza simplicidade com imensa qualidade e frescura.","zona":"Porto"},
  {"n":"Leitaria da Quinta do Paço (Matosinhos)","t":"Pastelaria Premium & Lanches","m":"Rua Brito Capelo 1237, 4450-072 Matosinhos","tel":"—","email":"—","p":"Alta","tCliente":"Pastelaria Premium & Lanches","gancho":"Venda por impulso. Pão de ló fofo para lanches em família ao fim de semana.","zona":"Matosinhos"},
  {"n":"Adega de São Nicolau","t":"Restaurante (Médio/Alto Padrão)","m":"Rua de São Nicolau 1, 4050-561 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Clássico ribeirinho. Sobremesa tradicional fantástica para acompanhar Vinho do Porto. Quebra zero via congelado.","zona":"Porto"},
  {"n":"Brasão Cervejaria (Aliados)","t":"Restaurante (Médio/Alto Padrão)","m":"Rua de Ramalho Ortigão 28, 4000-407 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Grande volume. Controle total de custos de dose e facilidade na gestão de stocks congelados.","zona":"Porto"},
  {"n":"O Paparico","t":"Restaurante (Alta Cozinha Portuguesa)","m":"Rua de Costa Cabral 2343, 4200-232 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Alta Cozinha Portuguesa)","gancho":"História e tradição. O pão de ló Ti Piedade como exemplo da doçaria conventual autêntica.","zona":"Porto"},
  {"n":"Taberna dos Mercadores","t":"Restaurante (Médio/Alto Padrão)","m":"Rua dos Mercadores 36, 4050-373 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Espaço pequeno. Solução congelada economiza espaço de cozinha e evita quebras.","zona":"Porto"},
  {"n":"Cantinho do Avillez (Porto)","t":"Restaurante (Médio/Alto Padrão)","m":"Rua de Mouzinho da Silveira 166, 4050-416 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Cozinha portuguesa com assinatura. Apresentar consistência do nosso padrão premium.","zona":"Porto"},
  {"n":"Muu Steakhouse","t":"Restaurante (Médio/Alto Padrão)","m":"Rua do Almada 149, 4050-037 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Foco na sobremesa de conforto após as carnes premium. Servir quente com sorvete ácido.","zona":"Porto"},
  {"n":"Abadia do Porto","t":"Restaurante Tradicional Premium","m":"Rua do Ateneu Comercial do Porto 22, 4000-380 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Público conservador que exige doçaria tradicional portuguesa à séria. Garantia de qualidade constante.","zona":"Porto"},
  {"n":"A Regaleira","t":"Restaurante Tradicional Premium","m":"Rua do Bonjardim 87, 4000-124 Porto","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Berço da francesinha, mas forte em pratos tradicionais. Sobremesa tradicional rápida e lucrativa.","zona":"Porto"},
  {"n":"Semente - Padaria Artesanal","t":"Padaria Artesanal & Orgânica","m":"Rua de S. Vicente 112, 4710-062 Braga","tel":"—","email":"—","p":"Alta","tCliente":"Padaria Artesanal & Orgânica","gancho":"Público biológico. Foco na receita tradicional pura, sem conservantes ou aditivos industriais.","zona":"Braga"},
  {"n":"Ferreira Capa","t":"Pastelaria de Referência / Tradicional","m":"Rua do Souto 131, 4700-329 Braga","tel":"—","email":"—","p":"Alta","tCliente":"Pastelaria de Referência / Tradicional","gancho":"Ponto histórico na cidade. Oferecer como opção premium de pão de ló húmido embalado para famílias.","zona":"Braga"},
  {"n":"Koyo Specialty Coffee","t":"Café de Especialidade","m":"Rua Dom Diogo de Sousa 37, 4700-424 Braga","tel":"—","email":"—","p":"Alta","tCliente":"Café de Especialidade","gancho":"Harmonia perfeita com café de especialidade ácido. Uma fatia simples servida elegantemente.","zona":"Braga"},
  {"n":"Cozinha da Terra","t":"Restaurante (Médio/Alto Padrão)","m":"Rua de São Lourenço, 4705-551 Braga","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Restaurante rústico minhoto. Pão de ló fofo fatiado na dose certa. Solução congelada com zero perdas.","zona":"Braga"},
  {"n":"Retroaria Bracarense","t":"Café de Charme / Tradicional","m":"Rua de São Marcos 17, 4700-328 Braga","tel":"—","email":"—","p":"Alta","tCliente":"Café de Charme / Tradicional","gancho":"Atmosfera vintage portuguesa. Encaixa perfeitamente no menu de lanche e chás tradicionais.","zona":"Braga"},
  {"n":"Restaurante Cozinha da Sé","t":"Restaurante (Médio/Alto Padrão)","m":"Rua Dom Diogo de Sousa 3, 4700-424 Braga","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Junto à Sé. Público local e turistas exigentes. Doçaria tradicional pronta a servir com quebra zero.","zona":"Braga"},
  {"n":"Taberna Belga","t":"Restaurante Tradicional / Elevada Rotação","m":"Rua Cónego Rafael Alvares da Costa 19, 4715-176 Braga","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante Tradicional / Elevada Rotação","gancho":"Rotação altíssima. Necessitam de sobremesas rápidas, deliciosas e de custo controlado.","zona":"Braga"},
  {"n":"Restaurante Augusto","t":"Restaurante (Médio/Alto Padrão)","m":"Rua de São Miguel 20, 4700-305 Braga","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Excelente garrafeira. Propor o Pão de Ló em fatia para acompanhar a carta de vinhos doces.","zona":"Braga"},
  {"n":"Padaria da Esquina","t":"Padaria Artesanal & Cafetaria","m":"Largo do Toural 104, 4810-427 Guimarães","tel":"—","email":"—","p":"Alta","tCliente":"Padaria Artesanal & Cafetaria","gancho":"Ponto de altíssima visibilidade. Clientes com apetite por pão artesanal e pastelaria fina.","zona":"Guimarães"},
  {"n":"A Cozinha por António Loureiro","t":"Restaurante (Médio/Alto Padrão / Autor)","m":"Largo do Serralho 4, 4800-414 Guimarães","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão / Autor)","gancho":"Restaurante sustentável. Foco estratégico no desperdício zero que o formato congelado garante.","zona":"Guimarães"},
  {"n":"Histórico by Papaboa","t":"Restaurante (Médio/Alto Padrão)","m":"Rua de Val Donas 4, 4810-230 Guimarães","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Palacete histórico. Pão de ló Ti Piedade como a sobremesa ideal que honra a história de Portugal.","zona":"Guimarães"},
  {"n":"Buxa","t":"Restaurante (Médio/Alto Padrão)","m":"Praça de São Tiago 18, 4810-244 Guimarães","tel":"—","email":"—","p":"Alta","tCliente":"Restaurante (Médio/Alto Padrão)","gancho":"Cozinha tradicional na praça mais movimentada. Rapidez de serviço sem risco de sobremesas estragadas.","zona":"Guimarães"},
  {"n":"Restaurante Nelson","t":"Restaurante peixe/marisco","m":"—","tel":"—","email":"geral@restaurantenelson.pt","p":"Alta","tCliente":"Restaurante peixe/marisco","gancho":"Lead identificado pela equipa comercial — Restaurante peixe/marisco em Ribamar.","zona":"Ribamar"},
  {"n":"A Sardinha","t":"Restaurante peixe","m":"—","tel":"—","email":"geral@asardinha.pt; restauranteasardinha@gmail.com","p":"Alta","tCliente":"Restaurante peixe","gancho":"Lead identificado pela equipa comercial — Restaurante peixe em Peniche.","zona":"Peniche"},
  {"n":"Golfinho Azul","t":"Restaurante","m":"—","tel":"—","email":"geral@golfinhoazul.pt","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Ribamar.","zona":"Ribamar"},
  {"n":"Foz Restaurante","t":"Restaurante contemporâneo","m":"—","tel":"—","email":"reservas@fozrestaurante.pt","p":"Alta","tCliente":"Restaurante contemporâneo","gancho":"Lead identificado pela equipa comercial — Restaurante contemporâneo em Ribamar.","zona":"Ribamar"},
  {"n":"Porto das Barcas","t":"Restaurante","m":"—","tel":"—","email":"geral@portodasbarcas.pt","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Ribamar.","zona":"Ribamar"},
  {"n":"Café Central Ribamar","t":"Restaurante/café","m":"—","tel":"—","email":"geral@cafecentralribamar.pt","p":"Alta","tCliente":"Restaurante/café","gancho":"Lead identificado pela equipa comercial — Restaurante/café em Ribamar.","zona":"Ribamar"},
  {"n":"Cantinho da Fonte","t":"Restaurante tradicional","m":"—","tel":"—","email":"geral@cantinhodafonte.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Peniche.","zona":"Peniche"},
  {"n":"Estelas","t":"Restaurante / Bar","m":"—","tel":"—","email":"restaurante.estelas@sapo.pt","p":"Alta","tCliente":"Restaurante / Bar","gancho":"Lead identificado pela equipa comercial — Restaurante / Bar em Peniche.","zona":"Peniche"},
  {"n":"Pateo da Saudade","t":"Restaurante tradicional","m":"—","tel":"—","email":"pateodasaudade@gmail.com","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Peniche.","zona":"Peniche"},
  {"n":"Profresco","t":"Restaurante","m":"—","tel":"—","email":"profresco@profresco.pt; geral@profresco.pt","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Peniche.","zona":"Peniche"},
  {"n":"Mar d’Areia","t":"Restaurante","m":"—","tel":"—","email":"geral@mardareia.pt","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Ericeira.","zona":"Ericeira"},
  {"n":"Tik Tapas","t":"Tapas/wine bar","m":"—","tel":"—","email":"reservas@tiktapas.pt","p":"Alta","tCliente":"Tapas/wine bar","gancho":"Lead identificado pela equipa comercial — Tapas/wine bar em Ericeira.","zona":"Ericeira"},
  {"n":"Adega da Vila","t":"Restaurante tradicional","m":"—","tel":"—","email":"geral@adegadavila.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Mafra.","zona":"Mafra"},
  {"n":"Uni Sushi","t":"Restaurante moderno","m":"—","tel":"—","email":"geral@unisushi.pt","p":"Alta","tCliente":"Restaurante moderno","gancho":"Lead identificado pela equipa comercial — Restaurante moderno em Ericeira.","zona":"Ericeira"},
  {"n":"Furnas Ericeira","t":"Restaurante premium","m":"—","tel":"—","email":"reservas@furnasericeira.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Ericeira.","zona":"Ericeira"},
  {"n":"Pedra Dura","t":"Restaurante grande volume","m":"—","tel":"—","email":"geral@pedradura.pt","p":"Alta","tCliente":"Restaurante grande volume","gancho":"Lead identificado pela equipa comercial — Restaurante grande volume em Ericeira.","zona":"Ericeira"},
  {"n":"Restaurante Prim","t":"Restaurante contemporâneo","m":"—","tel":"—","email":"geral@prim.pt","p":"Alta","tCliente":"Restaurante contemporâneo","gancho":"Lead identificado pela equipa comercial — Restaurante contemporâneo em Ericeira.","zona":"Ericeira"},
  {"n":"Jangada","t":"Restaurante marisco","m":"—","tel":"—","email":"reservas@jangada.pt","p":"Alta","tCliente":"Restaurante marisco","gancho":"Lead identificado pela equipa comercial — Restaurante marisco em Ericeira.","zona":"Ericeira"},
  {"n":"Cozinha da Sé","t":"Restaurante premium","m":"—","tel":"—","email":"reservas@cozinhadase.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Braga Centro.","zona":"Braga Centro"},
  {"n":"Dona Júlia","t":"Restaurante tradicional","m":"—","tel":"—","email":"geral@donajulia.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Braga.","zona":"Braga"},
  {"n":"Taberna Belga","t":"Restaurante grande volume","m":"—","tel":"—","email":"geral@tabernabelga.pt","p":"Alta","tCliente":"Restaurante grande volume","gancho":"Lead identificado pela equipa comercial — Restaurante grande volume em Braga.","zona":"Braga"},
  {"n":"Retrokitchen","t":"Restaurante contemporâneo","m":"—","tel":"—","email":"geral@retrokitchen.pt","p":"Alta","tCliente":"Restaurante contemporâneo","gancho":"Lead identificado pela equipa comercial — Restaurante contemporâneo em Braga.","zona":"Braga"},
  {"n":"Restaurante Tia Isabel","t":"Restaurante tradicional","m":"—","tel":"—","email":"reservas@tiaisabel.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Braga.","zona":"Braga"},
  {"n":"Bacalhau na Vila","t":"Restaurante português","m":"—","tel":"—","email":"geral@bacalhaunavila.pt","p":"Alta","tCliente":"Restaurante português","gancho":"Lead identificado pela equipa comercial — Restaurante português em Braga.","zona":"Braga"},
  {"n":"Michizaki","t":"Restaurante moderno","m":"—","tel":"—","email":"geral@michizaki.pt","p":"Alta","tCliente":"Restaurante moderno","gancho":"Lead identificado pela equipa comercial — Restaurante moderno em Braga.","zona":"Braga"},
  {"n":"Restaurante Arcoense","t":"Restaurante premium","m":"—","tel":"—","email":"geral@arcoense.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Braga.","zona":"Braga"},
  {"n":"Brasão Coliseu","t":"Restaurante grande volume","m":"—","tel":"—","email":"geral@brasao.pt","p":"Alta","tCliente":"Restaurante grande volume","gancho":"Lead identificado pela equipa comercial — Restaurante grande volume em Porto Centro.","zona":"Porto Centro"},
  {"n":"TerraPlana Café","t":"Restaurante/brunch","m":"—","tel":"—","email":"geral@terraplana.pt","p":"Alta","tCliente":"Restaurante/brunch","gancho":"Lead identificado pela equipa comercial — Restaurante/brunch em Porto.","zona":"Porto"},
  {"n":"Flow Restaurant & Bar","t":"Restaurante premium","m":"—","tel":"—","email":"reservas@flow.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Porto.","zona":"Porto"},
  {"n":"DOP","t":"Restaurante fine dining","m":"—","tel":"—","email":"geral@doprestaurante.pt","p":"Alta","tCliente":"Restaurante fine dining","gancho":"Lead identificado pela equipa comercial — Restaurante fine dining em Porto.","zona":"Porto"},
  {"n":"Tapabento","t":"Restaurante turístico","m":"—","tel":"—","email":"geral@tapabento.pt","p":"Alta","tCliente":"Restaurante turístico","gancho":"Lead identificado pela equipa comercial — Restaurante turístico em Porto Centro.","zona":"Porto Centro"},
  {"n":"Terminal 4450","t":"Restaurante premium","m":"—","tel":"—","email":"reservas@terminal4450.pt","p":"Alta","tCliente":"Restaurante premium","gancho":"Lead identificado pela equipa comercial — Restaurante premium em Matosinhos.","zona":"Matosinhos"},
  {"n":"Gaveto","t":"Restaurante referência","m":"—","tel":"—","email":"geral@gaveto.pt","p":"Alta","tCliente":"Restaurante referência","gancho":"Lead identificado pela equipa comercial — Restaurante referência em Matosinhos.","zona":"Matosinhos"},
  {"n":"O Gaveto Marisqueira","t":"Marisqueira","m":"—","tel":"—","email":"reservas@gaveto.pt","p":"Alta","tCliente":"Marisqueira","gancho":"Lead identificado pela equipa comercial — Marisqueira em Matosinhos.","zona":"Matosinhos"},
  {"n":"Salta o Muro","t":"Restaurante contemporâneo","m":"—","tel":"—","email":"geral@saltaomuro.pt","p":"Alta","tCliente":"Restaurante contemporâneo","gancho":"Lead identificado pela equipa comercial — Restaurante contemporâneo em Matosinhos.","zona":"Matosinhos"},
  {"n":"Meia-Nau","t":"Restaurante moderno","m":"—","tel":"—","email":"geral@meianau.pt","p":"Alta","tCliente":"Restaurante moderno","gancho":"Lead identificado pela equipa comercial — Restaurante moderno em Matosinhos.","zona":"Matosinhos"},
  {"n":"Casa Serrão","t":"Restaurante tradicional","m":"—","tel":"—","email":"reservas@casaserrao.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Matosinhos.","zona":"Matosinhos"},
  {"n":"Tito II","t":"Restaurante peixe/marisco","m":"—","tel":"—","email":"geral@tito.pt","p":"Alta","tCliente":"Restaurante peixe/marisco","gancho":"Lead identificado pela equipa comercial — Restaurante peixe/marisco em Matosinhos.","zona":"Matosinhos"},
  {"n":"Restaurante Mauritânia","t":"Restaurante grande volume","m":"—","tel":"—","email":"geral@mauritania.pt","p":"Alta","tCliente":"Restaurante grande volume","gancho":"Lead identificado pela equipa comercial — Restaurante grande volume em Matosinhos.","zona":"Matosinhos"},
  {"n":"DeCastro Gaia","t":"Restaurante contemporâneo","m":"—","tel":"—","email":"geral@decastro.pt","p":"Alta","tCliente":"Restaurante contemporâneo","gancho":"Lead identificado pela equipa comercial — Restaurante contemporâneo em Gaia.","zona":"Gaia"},
  {"n":"Ar de Rio","t":"Restaurante turístico","m":"—","tel":"—","email":"reservas@arderio.pt","p":"Alta","tCliente":"Restaurante turístico","gancho":"Lead identificado pela equipa comercial — Restaurante turístico em Cais de Gaia.","zona":"Cais de Gaia"},
  {"n":"Casa Adão","t":"Restaurante tradicional","m":"—","tel":"—","email":"geral@casaadao.pt","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Gaia.","zona":"Gaia"},
  {"n":"7 Grottos","t":"Restaurante moderno","m":"—","tel":"—","email":"geral@7grottos.pt","p":"Alta","tCliente":"Restaurante moderno","gancho":"Lead identificado pela equipa comercial — Restaurante moderno em Gaia.","zona":"Gaia"},
  {"n":"Esplanada do Teleférico","t":"Restaurante turístico","m":"—","tel":"—","email":"geral@esplanadadoteleferico.pt","p":"Alta","tCliente":"Restaurante turístico","gancho":"Lead identificado pela equipa comercial — Restaurante turístico em Gaia.","zona":"Gaia"},
  {"n":"Mar na Brasa","t":"Restaurante peixe","m":"—","tel":"—","email":"geral@marnabrasa.pt","p":"Alta","tCliente":"Restaurante peixe","gancho":"Lead identificado pela equipa comercial — Restaurante peixe em Gaia.","zona":"Gaia"},
  {"n":"Marisqueira de Ribamar","t":"Marisqueira (desde 1972)","m":"—","tel":"—","email":"marisqueiraribamar.com","p":"Alta","tCliente":"Marisqueira (desde 1972)","gancho":"Lead identificado pela equipa comercial — Marisqueira (desde 1972) em Ribamar.","zona":"Ribamar"},
  {"n":"Cervejaria O Pescador","t":"Cervejaria/Marisqueira","m":"—","tel":"—","email":"opescadorribamar@hotmail.com","p":"Alta","tCliente":"Cervejaria/Marisqueira","gancho":"Lead identificado pela equipa comercial — Cervejaria/Marisqueira em Ribamar.","zona":"Ribamar"},
  {"n":"Casa Rodrigues","t":"Restaurante familiar","m":"—","tel":"—","email":"casarodriguesribamar@gmail.com","p":"Alta","tCliente":"Restaurante familiar","gancho":"Lead identificado pela equipa comercial — Restaurante familiar em Ribamar.","zona":"Ribamar"},
  {"n":"Do Mar À Mesa","t":"Restaurante peixe/marisco","m":"—","tel":"—","email":"domaramesa.geral@gmail.com","p":"Alta","tCliente":"Restaurante peixe/marisco","gancho":"Lead identificado pela equipa comercial — Restaurante peixe/marisco em Ribamar.","zona":"Ribamar"},
  {"n":"Restaurante do Parque","t":"Restaurante tradicional","m":"—","tel":"—","email":"restaurantedoparque.peniche@gmail.com","p":"Alta","tCliente":"Restaurante tradicional","gancho":"Lead identificado pela equipa comercial — Restaurante tradicional em Peniche.","zona":"Peniche"},
  {"n":"Sushi Fish","t":"Restaurante moderno","m":"—","tel":"—","email":"geral@sushifish.pt","p":"Alta","tCliente":"Restaurante moderno","gancho":"Lead identificado pela equipa comercial — Restaurante moderno em Peniche.","zona":"Peniche"},
  {"n":"Taberna do Ganhão","t":"Restaurante","m":"—","tel":"—","email":"geral@tabernadoganhao.pt","p":"Alta","tCliente":"Restaurante","gancho":"Lead identificado pela equipa comercial — Restaurante em Peniche.","zona":"Peniche"},
  {"n":"Tasca do Joel","t":"Restaurante referência","m":"—","tel":"—","email":"reservas@tascadojoel.pt; reservas.tascadojoel@gmail.com","p":"Alta","tCliente":"Restaurante referência","gancho":"Lead identificado pela equipa comercial — Restaurante referência em Peniche.","zona":"Peniche"},
  {"n":"Tribeca Restaurante-Brasserie","t":"Brasserie","m":"—","tel":"—","email":"​tribeca-peniche@hotmail.com; tribeca@tribeca-restaurante.com","p":"Alta","tCliente":"Brasserie","gancho":"Lead identificado pela equipa comercial — Brasserie em Peniche.","zona":"Peniche"},
]


# ── Função de acesso ──────────────────────────────────────────
def get_db_horeca():
    """Retorna a base de leads reais por comercial."""
    return {
        "nuno":  DB_NUNO,
        "joao":  DB_JOAO,
        "oscar": DB_OSCAR,
    }

DB_HORECA_REAL = get_db_horeca()

# ── Base extra (catering + distribuidores) — integrada directamente ──
"""
Base de dados de leads para catering/eventos e distribuidores de congelados.
Importado pelo generate_leads.py
"""

# ════════════════════════════════════════════════════════════════
# CATERING & EVENTOS (Portugal inteiro)
# ════════════════════════════════════════════════════════════════
DB_CATERING = [
  {"n":"Doce Evento","t":"Catering & Eventos","m":"R. Tomás Ribeiro 34, Lisboa","tel":"213 540 210","email":"geral@doceevento.pt","p":"Alta","tCliente":"Catering premium casamentos e eventos corporativos","gancho":"Volume elevado por evento — unidose 85g é sobremesa elegante sem logística complexa.","zona":"Lisboa"},
  {"n":"Saveurs Catering","t":"Catering & Eventos","m":"Av. da República 50, Lisboa","tel":"217 960 400","email":"info@saveurs.pt","p":"Alta","tCliente":"Catering corporativo e eventos internacionais","gancho":"Clientela exigente — produto artesanal português com 40 anos diferencia a proposta.","zona":"Lisboa"},
  {"n":"Essência de Sabor","t":"Catering & Eventos","m":"R. Actor Vale 8, Lisboa","tel":"214 105 020","email":"geral@essenciadesabor.pt","p":"Alta","tCliente":"Casamentos e eventos sociais","gancho":"Sobremesa individual elegante — sem necessidade de pasteleiro no evento.","zona":"Lisboa"},
  {"n":"Eurest Portugal","t":"Catering & Eventos","m":"Av. Fontes Pereira de Melo 16, Lisboa","tel":"213 186 000","email":"geral@eurest.pt","p":"Alta","tCliente":"Catering colectivo / grandes volumes","gancho":"Volume de refeições diário — produto congelado em dose individual garante qualidade constante.","zona":"Lisboa"},
  {"n":"Gertal Companhia Geral","t":"Catering & Eventos","m":"R. Actor Tasso 12, Lisboa","tel":"213 619 200","email":"info@gertal.pt","p":"Alta","tCliente":"Catering colectivo / empresas e hospitais","gancho":"Grande volume com necessidade de sobremesa diferenciada e de fácil execução.","zona":"Lisboa"},
  {"n":"Quinta de Sant'Ana Eventos","t":"Catering & Eventos","m":"Gradil, Mafra","tel":"261 963 480","email":"eventos@quintadesantana.pt","p":"Alta","tCliente":"Eventos de luxo / casamentos premium","gancho":"Casamentos em quinta histórica — pão de ló artesanal é a sobremesa com mais narrativa.","zona":"Oeste"},
  {"n":"Solar de Mil Reis Eventos","t":"Catering & Eventos","m":"Av. Eng. Duarte Pacheco, Leiria","tel":"244 820 000","email":"eventos@solarmilreis.pt","p":"Alta","tCliente":"Eventos e banquetes","gancho":"Região com crescimento de eventos — produto artesanal diferencia a oferta de sobremesas.","zona":"Centro"},
  {"n":"Monte da Ravasqueira","t":"Catering & Eventos","m":"Estrada Monte da Ravasqueira, Arraiolos","tel":"266 498 280","email":"eventos@ravasqueira.com","p":"Alta","tCliente":"Eventos premium / enoturismo","gancho":"Enoturismo de luxo com sobremesas de autor — pão de ló Ti'Piedade como proposta artesanal.","zona":"Alentejo"},
  {"n":"Herdade do Esporão Eventos","t":"Catering & Eventos","m":"Herdade do Esporão, Reguengos de Monsaraz","tel":"266 509 280","email":"turismo@esporao.com","p":"Alta","tCliente":"Enoturismo / eventos internacionais","gancho":"Eventos com clientela internacional premium — produto português com receita secular.","zona":"Alentejo"},
  {"n":"Catering Delícias do Norte","t":"Catering & Eventos","m":"R. do Bonjardim 312, Porto","tel":"222 054 780","email":"info@deliciasdoporto.pt","p":"Alta","tCliente":"Catering casamentos / norte","gancho":"Norte com grande tradição de casamentos — dose individual em congelado facilita logística.","zona":"Porto"},
  {"n":"Quinta da Aveleda Eventos","t":"Catering & Eventos","m":"Quinta da Aveleda, Penafiel","tel":"255 718 200","email":"enoturismo@aveleda.pt","p":"Alta","tCliente":"Enoturismo e eventos premium","gancho":"Quinta histórica com eventos de alto valor — produto artesanal português complementa a experiência.","zona":"Norte"},
  {"n":"Vatel Portugal","t":"Catering & Eventos","m":"R. Gonçalo Cristóvão 195, Porto","tel":"222 073 900","email":"porto@vatel.pt","p":"Média","tCliente":"Escola de hotelaria / catering","gancho":"Formação e eventos — produto de referência para demonstrações e menus de escola.","zona":"Porto"},
  {"n":"CateringLab","t":"Catering & Eventos","m":"Av. de Braga 210, Guimarães","tel":"253 400 100","email":"info@cateringlab.pt","p":"Média","tCliente":"Catering eventos empresariais","gancho":"Região com indústria têxtil — eventos empresariais com necessidade de produto premium acessível.","zona":"Norte"},
  {"n":"Nobre Catering","t":"Catering & Eventos","m":"R. Alexandre Herculano 12, Coimbra","tel":"239 701 200","email":"geral@nobrecatering.pt","p":"Média","tCliente":"Catering académico e social","gancho":"Coimbra com muitos eventos académicos — sobremesa artesanal diferencia o menu.","zona":"Centro"},
  {"n":"Sabor a Festa Catering","t":"Catering & Eventos","m":"R. dos Correeiros 4, Faro","tel":"289 890 120","email":"geral@saborafesta.pt","p":"Média","tCliente":"Catering Algarve / turismo","gancho":"Algarve com turismo de alto valor — produto artesanal português autêntico.","zona":"Algarve"},
  {"n":"Quinta dos Vales Eventos","t":"Catering & Eventos","m":"Sítio dos Vales, Lagoa, Algarve","tel":"282 431 036","email":"eventos@quintadosvales.pt","p":"Alta","tCliente":"Enoturismo / casamentos Algarve","gancho":"Destino de casamentos internacionais — produto artesanal português é proposta de valor forte.","zona":"Algarve"},
  {"n":"Alma Catering","t":"Catering & Eventos","m":"R. Prior do Crato 30, Évora","tel":"266 705 360","email":"info@almacatering.pt","p":"Média","tCliente":"Catering Alentejo / eventos culturais","gancho":"Évora Património Mundial — eventos com clientela que valoriza produto nacional autêntico.","zona":"Alentejo"},
  {"n":"Banquetes Royal","t":"Catering & Eventos","m":"R. do Campo Alegre 1070, Porto","tel":"226 074 500","email":"geral@banquetesroyal.pt","p":"Média","tCliente":"Casamentos e banquetes","gancho":"Volume por evento — dose individual elimina desperdício em refeições de grande grupo.","zona":"Porto"},
  {"n":"Sabores com História","t":"Catering & Eventos","m":"Av. Infante Santo 42, Setúbal","tel":"265 522 800","email":"geral@saborescomhistoria.pt","p":"Média","tCliente":"Catering eventos sul","gancho":"Margem sul em crescimento — produto artesanal para eventos de nível médio-alto.","zona":"Sul"},
  {"n":"Quinta do Rol Eventos","t":"Catering & Eventos","m":"Torres Vedras","tel":"261 967 040","email":"eventos@quintadorol.com","p":"Alta","tCliente":"Enoturismo / casamentos Oeste","gancho":"Casamentos em adega — pão de ló artesanal marida perfeitamente com vinho e tradição.","zona":"Oeste"},
]

# ════════════════════════════════════════════════════════════════
# DISTRIBUIDORES DE CONGELADOS (zonas não cobertas pela equipa)
# ════════════════════════════════════════════════════════════════
DB_DISTRIBUIDORES = [
  {"n":"Frigoríficos do Algarve","t":"Distribuidor Congelados","m":"Zona Industrial, Loulé","tel":"289 416 200","email":"geral@frigorificos-algarve.pt","p":"Alta","tCliente":"Distribuidor regional congelados Algarve","gancho":"Algarve sem cobertura — canal HORECA forte com turismo de alto valor durante todo o ano.","zona":"Algarve"},
  {"n":"Distrinor Alimentar","t":"Distribuidor Congelados","m":"Zona Industrial de Braga","tel":"253 607 500","email":"comercial@distrinor.pt","p":"Alta","tCliente":"Distribuidor alimentar Norte","gancho":"Cobertura norte — complementa carteira de congelados com produto artesanal de alto valor percebido.","zona":"Norte"},
  {"n":"Irmãos Antunes Distribuição","t":"Distribuidor Congelados","m":"Zona Industrial, Évora","tel":"266 748 300","email":"geral@irmaosantunes.pt","p":"Alta","tCliente":"Distribuidor Alentejo","gancho":"Alentejo sem representação — região em crescimento turístico e gastronómico.","zona":"Alentejo"},
  {"n":"Frioeste","t":"Distribuidor Congelados","m":"R. Industrial, Viseu","tel":"232 420 800","email":"comercial@frioeste.pt","p":"Alta","tCliente":"Distribuidor congelados Viseu / Dão-Lafões","gancho":"Região interior sem cobertura — base de restauração e hotelaria a crescer.","zona":"Interior Norte"},
  {"n":"Caloricedos","t":"Distribuidor Congelados","m":"Zona Industrial, Castelo Branco","tel":"272 330 500","email":"geral@caloricedos.pt","p":"Alta","tCliente":"Distribuidor Beira Interior","gancho":"Beira Interior sem cobertura — território com hotelaria rural e turismo de natureza.","zona":"Interior"},
  {"n":"Setasul Distribuição","t":"Distribuidor Congelados","m":"Zona Industrial, Setúbal","tel":"265 700 400","email":"comercial@setasul.pt","p":"Alta","tCliente":"Distribuidor Setúbal / Alentejo Litoral","gancho":"Zona costeira com restauração forte — produto diferenciador na carteira de congelados.","zona":"Sul"},
  {"n":"Transmontana Alimentar","t":"Distribuidor Congelados","m":"Zona Industrial, Bragança","tel":"273 331 200","email":"geral@transmontanaalimentar.pt","p":"Alta","tCliente":"Distribuidor Trás-os-Montes","gancho":"Região com gastronomia forte e turismo rural crescente — produto com identidade nacional.","zona":"Nordeste"},
  {"n":"Frioraia","t":"Distribuidor Congelados","m":"Zona Industrial, Santarém","tel":"243 300 200","email":"comercial@frioraia.pt","p":"Alta","tCliente":"Distribuidor Ribatejo / Médio Tejo","gancho":"Zona limítrofe às áreas cobertas — pode complementar a distribuição com maior penetração.","zona":"Ribatejo"},
  {"n":"Atlântico Frio","t":"Distribuidor Congelados","m":"Zona Industrial, Viana do Castelo","tel":"258 800 300","email":"geral@atlanticofrio.pt","p":"Alta","tCliente":"Distribuidor Minho / Alto Lima","gancho":"Norte com tradição de festas e eventos — produto artesanal com grande aceitação local.","zona":"Norte"},
  {"n":"Giroalimentar","t":"Distribuidor Congelados","m":"Zona Industrial, Figueira da Foz","tel":"233 400 500","email":"comercial@giroalimentar.pt","p":"Alta","tCliente":"Distribuidor Pinhal Litoral / Figueira","gancho":"Costa com restauração turística — Ti'Piedade como sobremesa de referência no linear.","zona":"Centro Litoral"},
  {"n":"Friponente","t":"Distribuidor Congelados","m":"Zona Industrial, Portalegre","tel":"245 200 400","email":"geral@friponente.pt","p":"Média","tCliente":"Distribuidor Alto Alentejo","gancho":"Região com potencial não explorado — pousadas, hotéis rurais e restauração de qualidade.","zona":"Alentejo"},
  {"n":"Norte Frio Distribuição","t":"Distribuidor Congelados","m":"Zona Industrial, Vila Real","tel":"259 370 300","email":"comercial@nortefrio.pt","p":"Média","tCliente":"Distribuidor Douro / Trás-os-Montes Sul","gancho":"Enoturismo do Douro em crescimento — produto artesanal português premium.","zona":"Douro"},
  {"n":"Mediterrânico Alimentar","t":"Distribuidor Congelados","m":"Zona Industrial, Tavira","tel":"281 325 600","email":"geral@mediterranico-alimentar.pt","p":"Alta","tCliente":"Distribuidor Sotavento Algarvio","gancho":"Algarve Oriental com turismo internacional — produto artesanal de alto valor percebido.","zona":"Algarve"},
  {"n":"Expofrio","t":"Distribuidor Congelados","m":"Zona Industrial, Aveiro","tel":"234 380 700","email":"comercial@expofrio.pt","p":"Alta","tCliente":"Distribuidor Aveiro / Baixo Vouga","gancho":"Região industrial com restauração em crescimento — Ti'Piedade na carteira de congelados premium.","zona":"Centro Norte"},
  {"n":"Frioguarda","t":"Distribuidor Congelados","m":"Zona Industrial, Guarda","tel":"271 210 500","email":"geral@frioguarda.pt","p":"Média","tCliente":"Distribuidor Serra da Estrela / Beira Alta","gancho":"Turismo de montanha e neve — produto de doçaria artesanal premium em contexto de acolhimento.","zona":"Interior Norte"},
]


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

    assinatura = """Com os melhores cumprimentos,

Rui Bernardes
Departamento Comercial | Pão de Ló Ti'Piedade
sales@tipiedade.com | www.tipiedade.com"""

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
    <td style="background:#4a2417;border-radius:14px 14px 0 0;padding:0">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="padding:24px 36px 16px">
            <img src="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAkACQAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAGXAvkDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDzSiiiszqCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACjBGMjr0orsPDVnba14bv7CePdLbt5sLKPmXI6D2yp49658TiFh4e0ktLq/zNKVP2kuVbnH1p2+hXtzo0+qRqvkQnByeWHcj6U+PRZLW+totYSWygmYqJDjj3+mSM/WvR7PTbO28NSaal2rQFHjafI435/DPzVxY/MlQjH2et2vNW669zfD4Z1G+boeRUAE9Bmtq/0WP+1msNHkkvnjXMhAHXvj1xxW7/AGfFofgWe4ltzHqFzmJy4wwBYjaPQbRn3+mK6KmOpxjDl1cmkls9er8jKNCTbvsjiKKKK7TEKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKVW2n2PBHqKSrBtXay+1oVZA22QDqh7ZHoex9iKUmluNJvYicIFAVtxHfGBj/GmU9JCuFLEJuycDP6VYvbM2/lzIMwSjKEdAe4/X8uuDkBcyTUX1C11cqUUUVQgooooAKKKKACuo8B3XkeIfJJ4niZce4+b+hrl6sWE0dvqNtNKpaOOV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b6mvbWlvLavLC0gRZPLB+yCUscZ3HJ+UewB/GoLyxVIblvlE9rKEkCfddTnDAduRyPccDBp0NzNbQqthcIF80S5MoRgcY2sDjP8AI/pUd1cRJbzxxqqyXEiu6I25Y1GflB75Jz3xgc9a44qpz6P+ut/lf/gaG75eUzqKKK7TAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKK2bPwprN8iyR2TJG3RpSE/Q8/pWdStTpK9SSS8yowlJ2irmNRV7UNH1DSyPttq8QJwGOCpPpkcVRqoTjNc0HdeQnFxdmFFW9P0y81W4aCyh82RV3ldwXjIGeSPUVpf8Ibr/8Az4f+Ro//AIqs6mJoU3yzmk/NpFRpTkrxTZhUVrXvhnV9PtHurq08uFMbm8xDjJwOAc9TWTV06tOquanJNeTuKUZRdpKwUVdsNIv9USZrK3MwhAL4YAjOccE89D0qlTU4tuKeq38hOLSuwoqzFp91NYT3scWbeAgSPuHyknA4znvVampJ3s9gaa3CirNhp91qdyLa0i8yUgkLuA4HuTVahSTbinqgs7XCip7SyuL6UxW0e9wpcjIHA5J5qCjmTdr6is7XCipIIZLm4jgiXdJIwRFzjJJwBWlrU4V47AabDZG2G1wu13dvVnA5+lRKpaaguvnt+pSjo2ZNFFXrrR7+zsoLye2ZbecAxyAgggjI6Hjj1qpTjFpN2b2Eot6oo0UVZutPurKO3e4i2LcRiSI7gdynoeDx+NNySaTerCzepWoopyI8rhI1Z3Y4CqMk0xDaK2l8Ja60ZkGnvgdi6g/lnNZVzaz2c5huYXikHVXXBrKFelUdoSTfk0VKnOOslYiooorUkKKKkghkubiOCJd0kjBEXOMknAFDaSuwI6K3f+EN1/8A58P/ACNH/wDFVXvP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8enIqjRRSjCMb8qtfUHJvcASCCDgjvXRQeM9UjgWGdba7Ven2iPcfrwRzXO0VFWhSrJKpG9ioVJQ+F2NbVfEeo6xEsNxIqwqciKNdq+31xWTRRVU6UKUeWCshSlKTvJ3Jbad7W6iuIsCSJw65GRkHIp15dy315LdTkGWVtzYGBmoKKrkjzc1tRXdrBVu91K4v4raOcqVtohFHgY+UetVKKHCLak1qgTaVgBIIIOCO9XtR1a71XyDduHaFNitjkj39TVGik4RclJrVbApNKwVNbXU9ncJPbStFKhyGU4NQ0VTSasxJ21R0X/CZXxVt9np7u2SZGg+bJ79cVk6jql5qtx515O0jD7o6Ko9AO1U6Kwp4WjSlzQikzSVWclaTCiiitzMKltp3tbqK4iwJInDrkZGQcioqKGk1ZgnYluZ3urqW4lwZJXLtgYGScmoqKKEklZA9Qq3p+pXGmSyyWxUNLEYm3DPynGf5VUopShGa5ZK6Gm07oKt22pXFrZXNmhUwXIAkRhnkdCPQ1UoolCMlaSuCbWqCrb6lcPpcenZUW6OZMAYLN6k96qUUShGVrrYE2tgq9Y6tdafBcwQlTFcrskR1yCOf8ao0UpwjNcsldApOLugroovG2tQwpGJYiEUKC0eScetc7RUVcPSrWVSKdu5UKkofC7HS/8ACda3/wA9If8Av0Kwby7lvryW6nIMsrbmwMDNQUVNLC0aLvTik/IJ1ZzVpO5b03UrjSrwXVqVEoUr8wyMGtr/AITrW/8AnpD/AN+hXNUUVcJQqy5qkE35ocas4K0XY6C58Z6vdWstvK8JjlQo2IwDgjBrn6KKqlQpUVanFL0JnUlP4nct2upXFnZ3drEVEV0oWTIycDOMfnVSiitFCKbaW4m29GWxqVwNKbTSVa2MnmgMMlW9Qe3/ANc1UooojCMb2W4Nt7hRRRTEFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFAH//2Q==" alt="Ti'Piedade" style="height:52px;display:block;border:0">
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
    <td style="background:#4a2417;border-radius:0 0 14px 14px;padding:20px 36px">
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
        msg["From"]    = "Pão de Ló Ti'Piedade <sales@tipiedade.com>"
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
                    ok = enviar_nurturing_lead(server, smtp_user, lead, n, "Rui Bernardes")
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

Aqui estão os teus leads HORECA para a semana {sem} ({semana_datas()}).

O Rui Bernardes (Departamento Comercial) enviou automaticamente o Email {email_num} da sequência de nurturing a {leads_com_email} leads com email disponível ({sem_email} sem email — para estes contacta por telefone).

A sequência tem 4 emails semanais. Quando aparecer "✓ Pronto p/ visita" no Excel — o contacto já recebeu todos os emails e está preparado para a tua visita.

O teu papel entra aqui: contacto presencial, apresentação do produto e fecho da venda.

Bom trabalho,
Rui Bernardes
Departamento Comercial | Ti'Piedade
sales@tipiedade.com
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
                    ok = enviar_nurturing_lead(server, smtp_user, lead, n, "Rui Bernardes")
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
                    ok = enviar_nurturing_lead(server, smtp_user, lead, n, "Rui Bernardes")
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
