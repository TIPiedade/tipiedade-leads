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

import math, os, datetime, json, base64, urllib.request, urllib.error
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Normalização de tipologias ────────────────────────────────────────────────
MAPA_TIPOLOGIA = {
    # Café & Brunch
    "Café de Especialidade":"Café & Brunch","Café de Especialidade & Brunch":"Café & Brunch",
    "Café de Especialidade & Lanches":"Café & Brunch","Café & Brunch Premium":"Café & Brunch",
    "Café & Brunch Moderno":"Café & Brunch","Cafetaria Gourmet & Regional":"Café & Brunch",
    "Cafetaria Saudável & Brunch":"Café & Brunch","Café Histórico":"Café & Brunch",
    "Café / Restaurante Histórico":"Café & Brunch","Café Histórico / Turismo":"Café & Brunch",
    "Café Tradicional":"Café & Brunch","Café de Charme / Tradicional":"Café & Brunch",
    "Café & Wine Bar":"Café & Brunch","Brunch & Comida Saudável":"Café & Brunch",
    "Brunch & Cozy Cafe":"Café & Brunch","Brunch & Healthy Food":"Café & Brunch",
    "Brunch & Lanches":"Café & Brunch","Brunch & Specialty Coffee":"Café & Brunch",
    "Brunch / Gourmet (★ 4.8)":"Café & Brunch","Brunch Bar":"Café & Brunch",
    "Brunch Club & Café":"Café & Brunch","Loja de Chocolate / Café":"Café & Brunch",
    "Restaurante / brunch":"Café & Brunch","Restaurante & Brunch Premium":"Café & Brunch",
    "Restaurante de Praia & Brunch":"Café & Brunch","Restaurante de Praia & Brunch Premium":"Café & Brunch",
    "Restaurante/brunch":"Café & Brunch",
    # Pastelaria & Padaria
    "Pastelaria Premium":"Pastelaria & Padaria","Pastelaria Premium & Brunch":"Pastelaria & Padaria",
    "Pastelaria Premium & Lanches":"Pastelaria & Padaria","Pastelaria & Padaria Fina":"Pastelaria & Padaria",
    "Pastelaria":"Pastelaria & Padaria","Pastelaria Artesanal":"Pastelaria & Padaria",
    "Pastelaria Artesanal & Casa de Chá":"Pastelaria & Padaria","Pastelaria Gourmet & Chá":"Pastelaria & Padaria",
    "Pastelaria Histórica":"Pastelaria & Padaria","Pastelaria Regional":"Pastelaria & Padaria",
    "Pastelaria de Referência / Tradicional":"Pastelaria & Padaria","Pastelaria / Café Clássico":"Pastelaria & Padaria",
    "Pastelaria/Distribuidor PT":"Pastelaria & Padaria","Padaria Artesanal & Gourmet":"Pastelaria & Padaria",
    "Padaria Artesanal Premium":"Pastelaria & Padaria","Padaria Artesanal":"Pastelaria & Padaria",
    "Padaria Artesanal & Bistrô":"Pastelaria & Padaria","Padaria Artesanal & Cafetaria":"Pastelaria & Padaria",
    "Padaria Artesanal & Café":"Pastelaria & Padaria","Padaria Artesanal & Orgânica":"Pastelaria & Padaria",
    "Padaria Artesanal Sourdough":"Pastelaria & Padaria","Gelataria premium":"Pastelaria & Padaria",
    "Grupo Restauração / Pastelaria":"Pastelaria & Padaria","Grupo Cafetaria / Pastelaria":"Pastelaria & Padaria",
    # Casa de Chá & Pastelaria
    "Casa de Chá & Café de Charme":"Casa de Chá & Pastelaria","Casa de Chá & Pastelaria Fina":"Casa de Chá & Pastelaria",
    # Restaurante Premium
    "Restaurante Premium":"Restaurante Premium","Restaurante premium":"Restaurante Premium",
    "Restaurante Gourmet":"Restaurante Premium","Restaurante Moderno":"Restaurante Premium",
    "Restaurante moderno":"Restaurante Premium","Restaurante Contemporâneo":"Restaurante Premium",
    "Restaurante contemporâneo":"Restaurante Premium","Restaurante Português Moderno":"Restaurante Premium",
    "Restaurante cozinha portuguesa reinventada":"Restaurante Premium","Restaurante premium vínico":"Restaurante Premium",
    "Restaurante premium/tradicional":"Restaurante Premium","Restaurante Autor":"Restaurante Premium",
    "Restaurante (Alta Cozinha Portuguesa)":"Restaurante Premium","Restaurante (Médio/Alto Padrão / Autor)":"Restaurante Premium",
    "Fine dining / cozinha autoral":"Restaurante Premium","Restaurante Fine Dining":"Restaurante Premium",
    "Restaurante fine dining":"Restaurante Premium","Restaurante fusão (★ 9.8 TheFork)":"Restaurante Premium",
    "Restaurante referência":"Restaurante Premium","Restaurante português":"Restaurante Premium","Restaurante/café":"Restaurante Premium",
    # Restaurante Michelin / Hotel
    "Restaurante (Michelin)":"Restaurante Michelin / Hotel","Hotel/restaurante Michelin":"Restaurante Michelin / Hotel",
    "Hotel/restaurante":"Restaurante Michelin / Hotel","Restaurante de Hotel Histórico":"Restaurante Michelin / Hotel",
    "Restaurante/hotel":"Restaurante Michelin / Hotel","Cadeia Hoteleira":"Restaurante Michelin / Hotel",
    "Cadeia Hoteleira Nacional":"Restaurante Michelin / Hotel",
    # Restaurante Tradicional
    "Restaurante Tradicional":"Restaurante Tradicional","Restaurante tradicional":"Restaurante Tradicional",
    "Restaurante Tradicional Premium":"Restaurante Tradicional","Restaurante Tradicional de Caça & Grelhados":"Restaurante Tradicional",
    "Restaurante Regional Premium":"Restaurante Tradicional","Restaurante regional":"Restaurante Tradicional",
    "Restaurante Regional":"Restaurante Tradicional","Restaurante Temático":"Restaurante Tradicional",
    "Restaurante temático":"Restaurante Tradicional","Restaurante Temático / Turismo":"Restaurante Tradicional",
    "Restaurante turístico":"Restaurante Tradicional","Restaurante português costeiro":"Restaurante Tradicional",
    "Restaurante (Médio/Alto Padrão / Vista Rio)":"Restaurante Tradicional","Restaurante (Médio/Alto Padrão)":"Restaurante Tradicional",
    "Restaurante & Taberna":"Restaurante Tradicional","Taberna / Restaurante Tradicional":"Restaurante Tradicional",
    "Taberna Contemporânea":"Restaurante Tradicional","Adega / Restaurante":"Restaurante Tradicional",
    "Adega / Taberna":"Restaurante Tradicional","Restaurante Panorâmico":"Restaurante Tradicional",
    "Restaurante panorâmico":"Restaurante Tradicional","Restaurante Praia":"Restaurante Tradicional",
    "Restaurante grande volume":"Restaurante Tradicional","Restauração Colectiva":"Restaurante Tradicional",
    "Restauração Colectiva / Facility":"Restaurante Tradicional","Restauração em Aeroportos":"Restaurante Tradicional",
    # Restaurante Peixe & Marisco
    "Restaurante de Peixe & Marisco":"Restaurante Peixe & Marisco","Restaurante peixe/marisco premium":"Restaurante Peixe & Marisco",
    "Restaurante marisco/peixe com vista mar":"Restaurante Peixe & Marisco","Restaurante costeiro":"Restaurante Peixe & Marisco",
    "Restaurante de Peixe":"Restaurante Peixe & Marisco","Restaurante de Peixe (Médio/Alto Padrão)":"Restaurante Peixe & Marisco",
    "Restaurante de Peixe Moderno":"Restaurante Peixe & Marisco","Restaurante marisco":"Restaurante Peixe & Marisco",
    "Restaurante peixe/marisco":"Restaurante Peixe & Marisco","Restaurante peixe":"Restaurante Peixe & Marisco",
    "Marisqueira":"Restaurante Peixe & Marisco","Marisqueira (desde 1972)":"Restaurante Peixe & Marisco",
    "Cervejaria / Marisco":"Restaurante Peixe & Marisco","Cervejaria/Marisqueira":"Restaurante Peixe & Marisco",
    # Restaurante Casual
    "Restaurante":"Restaurante Casual","Restaurante familiar":"Restaurante Casual","Restaurante informal":"Restaurante Casual",
    "Restaurante internacional":"Restaurante Casual","Restaurante italiano":"Restaurante Casual","Italiano":"Restaurante Casual",
    "Restaurante/bar":"Restaurante Casual","Restaurante / Bar":"Restaurante Casual","Restaurante/bar moderno":"Restaurante Casual",
    "Restaurante/Distribuidor":"Restaurante Casual","Pizzaria Gourmet":"Restaurante Casual","Bistro americano":"Restaurante Casual",
    "Cervejaria":"Restaurante Casual","Cervejaria / Restaurante":"Restaurante Casual","Cervejaria Tradicional":"Restaurante Casual",
    "Cervejaria Lisboa":"Restaurante Casual","Tapas/wine bar":"Restaurante Casual","Tasca / Petiscos":"Restaurante Casual","Brasserie":"Restaurante Casual",
    # Tasca & Fado
    "Tasca Contemporânea":"Tasca & Fado","Tasca Moderna":"Tasca & Fado","Tasca Moderna / Ribeirinha":"Tasca & Fado",
    "Tasca Tradicional":"Tasca & Fado","Casa de Fado Premium":"Tasca & Fado","Restaurante / Casa de Fado":"Tasca & Fado",
    "Restaurante / Fado":"Tasca & Fado","Fado & Restaurante":"Tasca & Fado",
    # Beach Club & Bar
    "Beach Club / Restaurante":"Beach Club & Bar","Beach club":"Beach Club & Bar","Beach restaurant":"Beach Club & Bar",
    # Catering & Eventos
    "Catering & Eventos":"Catering & Eventos","Restaurante & Espaço de Eventos":"Catering & Eventos",
    "Restaurante eventos":"Catering & Eventos","Restaurante & Eventos":"Catering & Eventos",
    "Grupo Restauração / Catering":"Catering & Eventos","Grupo Restauração / Entretenimento":"Catering & Eventos",
    # Grupo Restauração
    "Grupo Restauração":"Grupo Restauração","Grupo Restauração Premium":"Grupo Restauração",
    # Garrafeira & Gourmet
    "Garrafeira & Gourmet Deli":"Garrafeira & Gourmet","Mercearia Gourmet PT":"Garrafeira & Gourmet",
    "Mercearia Gourmet Europeia":"Garrafeira & Gourmet","Mercearia Portuguesa NL":"Garrafeira & Gourmet",
    "Mercearia Saudade":"Garrafeira & Gourmet","Mercearia Saudade IT":"Garrafeira & Gourmet",
    "Mercearia Saudade Online":"Garrafeira & Gourmet","Mercearia Saudade WA":"Garrafeira & Gourmet",
    "Mercearia Saudade Wallonia":"Garrafeira & Gourmet","Mercearia/Distribuidor PT":"Garrafeira & Gourmet",
    "Mercearia/Gastronomia PT":"Garrafeira & Gourmet",
    # Distribuidor
    "Distribuidor Congelados":"Distribuidor","Distribuidor Congelados HORECA":"Distribuidor",
    "Distribuidor HORECA Ibérico":"Distribuidor","Distribuidor HORECA Ibérico CH":"Distribuidor",
    "Distribuidor HORECA Étnico":"Distribuidor","Distribuidor Ibérico AU":"Distribuidor",
    "Distribuidor Ibérico BE":"Distribuidor","Distribuidor Ibérico França":"Distribuidor",
    "Distribuidor Ibérico HORECA":"Distribuidor","Distribuidor Ibérico IT":"Distribuidor",
    "Distribuidor Ibérico NL":"Distribuidor","Distribuidor PT Angola":"Distribuidor",
    "Distribuidor PT País Basco":"Distribuidor","Distribuidor Produtos PT":"Distribuidor",
    "Distribuidor Produtos PT Suíça":"Distribuidor","Distribuidora PT (Grupo Delta)":"Distribuidor",
    "Exportador/Distribuidor Alimentar":"Distribuidor","Gastronomia/Distribuidor PT":"Distribuidor",
    "Importador Congelados PT":"Distribuidor","Importador Congelados PT CA":"Distribuidor",
    "Importador Especialidades PT":"Distribuidor","Importador Gourmet PT":"Distribuidor",
    "Importador Histórico PT Suíça":"Distribuidor","Importador Ibérico HORECA":"Distribuidor",
    "Importador Retalho PT AO":"Distribuidor","Importador Retalho Saudade":"Distribuidor",
    "Importador/Distribuidor PT":"Distribuidor","Importador/Distribuidor PT AU":"Distribuidor",
    "Importador/Distribuidor PT BE":"Distribuidor","Importador/Distribuidor PT BR":"Distribuidor",
    "Importador/Distribuidor PT Gourmet":"Distribuidor","Importador/Distribuidor PT HORECA":"Distribuidor",
    "Importador/Distribuidor PT IT":"Distribuidor","Importador/Distribuidor PT JP":"Distribuidor",
    "Importador/Distribuidor PT+ES":"Distribuidor","Importador/Mercearia PT":"Distribuidor",
    "Promotor/Distribuidor PT":"Distribuidor","Retalho Gourmet Premium UK":"Distribuidor",
    "Grande Distribuição AU":"Distribuidor","Grande Distribuição BE":"Distribuidor",
    "Grande Distribuição BR":"Distribuidor","Grande Distribuição CA":"Distribuidor",
    "Grande Distribuição CH":"Distribuidor","Grande Distribuição DE":"Distribuidor",
    "Grande Distribuição ES":"Distribuidor","Grande Distribuição FR":"Distribuidor",
    "Grande Distribuição JP":"Distribuidor","Grande Distribuição LU":"Distribuidor",
    "Grande Distribuição NL":"Distribuidor","Grande Distribuição UK":"Distribuidor",
    # E-commerce Saudade
    "E-commerce Saudade Brasil":"E-commerce Saudade","E-commerce Saudade Europa":"E-commerce Saudade",
    "E-commerce Saudade UK":"E-commerce Saudade",
}

def normalizar_tipologia(t):
    return MAPA_TIPOLOGIA.get(t, t)

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
  # ── Nuno — Batch 2 (40 leads) ──────────────────────────────────────────
  {"n":"Eleven","t":"Restaurante Fine Dining","m":"Rua Marquês da Fronteira, Jardim Amália Rodrigues, Lisboa","tel":"213 862 211","email":"eleven@eleven.pt","p":"Alta","tCliente":"Restaurante Fine Dining","gancho":"Sobremesa de autor com receita conventual secular — encaixa na narrativa de produto premium português.","zona":"Lisboa (Marquês de Pombal)"},
  {"n":"Bairro do Avillez","t":"Restaurante & Taberna","m":"Rua Nova da Trindade 18, 1200-303 Lisboa","tel":"215 830 290","email":"geral@bairrodoavillez.pt","p":"Alta","tCliente":"Restaurante Premium","gancho":"José Avillez valoriza produto artesanal nacional — pão de ló Ti'Piedade como sobremesa âncora.","zona":"Lisboa (Chiado)"},
  {"n":"Solar dos Nunes","t":"Restaurante Tradicional","m":"Rua dos Lusíadas 68, 1300-372 Lisboa","tel":"213 647 359","email":"restaurantesolardosnunes@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Referência gastronómica em Lisboa Ocidental — clientela fiel que valoriza receitas clássicas.","zona":"Lisboa (Belém)"},
  {"n":"A Cevicheria","t":"Restaurante Moderno","m":"Rua Dom Pedro V 129, 1250-096 Lisboa","tel":"218 038 815","email":"info@acevicheria.pt","p":"Alta","tCliente":"Restaurante Moderno","gancho":"Chef Kiko tem abordagem criativa — propor o pão de ló desconstruído como sobremesa de fusão.","zona":"Lisboa (Príncipe Real)"},
  {"n":"Palácio Chiado","t":"Restaurante & Eventos","m":"Rua do Alecrim 70, 1200-018 Lisboa","tel":"213 460 491","email":"eventos@palaciochiado.pt","p":"Alta","tCliente":"Restaurante & Espaço de Eventos","gancho":"Eventos corporativos e casamentos — dose individual congelada elimina desperdício em banquetes.","zona":"Lisboa (Chiado)"},
  {"n":"Pharmácia","t":"Restaurante Temático","m":"Rua Marechal Saldanha 1, 1249-069 Lisboa","tel":"213 462 146","email":"pharmacy@museudafarmacia.pt","p":"Alta","tCliente":"Restaurante Temático / Turismo","gancho":"Turistas internacionais — narrativa do produto centenário encaixa no conceito do restaurante.","zona":"Lisboa (Bairro Alto)"},
  {"n":"Tasca do Chico","t":"Fado & Restaurante","m":"Rua do Diário de Notícias 39, 1200-143 Lisboa","tel":"961 339 696","email":"tascadochico@gmail.com","p":"Alta","tCliente":"Restaurante / Casa de Fado","gancho":"Experiência cultural portuguesa completa — pão de ló como sobremesa de fecho da noite de fado.","zona":"Lisboa (Bairro Alto)"},
  {"n":"Tasca da Esquina","t":"Restaurante Moderno Português","m":"Rua Domingos Sequeira 41C, 1350-119 Lisboa","tel":"210 993 939","email":"tacsadasesquina@gmail.com","p":"Alta","tCliente":"Restaurante Moderno","gancho":"Chef Victor Sobral usa produto nacional de excelência — pão de ló como sobremesa âncora do menu.","zona":"Lisboa (Campo de Ourique)"},
  {"n":"ZeroZero","t":"Pizzaria Gourmet","m":"Rua da Escola Politécnica 32, 1250-099 Lisboa","tel":"213 475 313","email":"reservas@zerozero.pt","p":"Alta","tCliente":"Restaurante Gourmet","gancho":"Menu cuidado — propor pão de ló como única sobremesa portuguesa num menu internacional.","zona":"Lisboa (Rato)"},
  {"n":"Café de São Bento","t":"Restaurante Clássico","m":"Rua de São Bento 212, 1200-822 Lisboa","tel":"213 952 911","email":"geral@cafesaobento.com","p":"Alta","tCliente":"Restaurante Clássico","gancho":"Clientela política e empresarial — sobremesa tradicional de qualidade para fechar reuniões à mesa.","zona":"Lisboa (São Bento)"},
  {"n":"O Corvo","t":"Tasca Moderna","m":"Rua João Ortigão Ramos 6B, 1600-394 Lisboa","tel":"217 590 496","email":"restaurante.ocorvo@gmail.com","p":"Alta","tCliente":"Tasca Moderna","gancho":"Referência gastronómica de bairro — pão de ló como sobremesa rotativa da semana.","zona":"Lisboa (Campolide)"},
  {"n":"Feitoria","t":"Restaurante Fine Dining","m":"Altis Belém Hotel, Doca do Bom Sucesso, 1400-038 Lisboa","tel":"210 400 208","email":"feitoria@altishotels.com","p":"Alta","tCliente":"Fine Dining / Hotel 5 estrelas","gancho":"Uma estrela Michelin — produto artesanal com receita secular encaixa na proposta de alta gastronomia.","zona":"Lisboa (Belém)"},
  {"n":"Loco","t":"Restaurante Fine Dining","m":"Rua dos Navegantes 53B, 1200-731 Lisboa","tel":"213 951 861","email":"info@loco.pt","p":"Alta","tCliente":"Fine Dining","gancho":"Uma estrela Michelin — propor pão de ló como base para sobremesa de autor com técnica moderna.","zona":"Lisboa (Santos)"},
  {"n":"Prado","t":"Restaurante Contemporâneo","m":"Travessa das Pedras Negras 2, 1100-404 Lisboa","tel":"910 458 978","email":"geral@restauranteprado.pt","p":"Alta","tCliente":"Restaurante Farm-to-Table","gancho":"Foco em produto local e artesanal — receita conventual de Alcobaça é exatamente o que procuram.","zona":"Lisboa (Alfama)"},
  {"n":"Taberna Albricoque","t":"Tasca Tradicional","m":"Rua do Terreiro do Trigo 10, 1100-522 Lisboa","tel":"218 870 098","email":"albricoque.taberna@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Foco em produto português genuíno — sobremesa conventual de autor.","zona":"Lisboa (Alfama)"},
  {"n":"O Cortiço & Netos","t":"Restaurante Tradicional","m":"Calçada da Graça 4, 1170-167 Lisboa","tel":"218 860 479","email":"info@ocorticopresunto.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Embutidos e produto artesanal nacional — pão de ló como fecho natural de um menu 100% português.","zona":"Lisboa (Graça)"},
  {"n":"Tasca do Vigário","t":"Tasca Tradicional","m":"Rua do Vigário 70, 1100-612 Lisboa","tel":"218 860 357","email":"tascavigario@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Bairro histórico com turistas e locais — sobremesa clássica com história diferencia o menu.","zona":"Lisboa (Alfama)"},
  {"n":"Zé da Mouraria","t":"Tasca Tradicional","m":"Rua dos Lagares 25, 1100-313 Lisboa","tel":"218 874 208","email":"zedamouraria@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Mouraria em ascensão turística — autenticidade do produto Ti'Piedade como parte da experiência.","zona":"Lisboa (Mouraria)"},
  {"n":"Restaurante Aqui Há Peixe","t":"Restaurante de Peixe","m":"Rua Trindade Coelho 15, 1200-470 Lisboa","tel":"213 432 154","email":"aquihapeixe@gmail.com","p":"Alta","tCliente":"Restaurante de Peixe & Marisco","gancho":"Peixe fresco e produto nacional — sobremesa doce e leve após o marisco é o fecho ideal.","zona":"Lisboa (Príncipe Real)"},
  {"n":"Tasca do Jaime","t":"Tasca Tradicional","m":"Rua das Portas de Santo Antão 37, 1150-268 Lisboa","tel":"213 424 765","email":"tascadojaime@gmail.com","p":"Média","tCliente":"Tasca Tradicional","gancho":"Zona turística intensa — produto português genuíno para turistas que procuram autenticidade.","zona":"Lisboa (Rossio)"},
  {"n":"Cervejaria Solmar","t":"Cervejaria Clássica","m":"Rua das Portas de Santo Antão 108, 1150-269 Lisboa","tel":"213 463 647","email":"geral@restaurantesolmar.pt","p":"Alta","tCliente":"Cervejaria Clássica","gancho":"Instituição de Lisboa — sobremesa de fecho clássica para a clientela fiel.","zona":"Lisboa (Rossio)"},
  {"n":"Café Nicola","t":"Café Histórico","m":"Praça Dom Pedro IV 24, 1100-200 Lisboa","tel":"213 460 579","email":"cafeni@cafeni.pt","p":"Alta","tCliente":"Café Histórico / Turismo","gancho":"Ícone cultural de Lisboa — pão de ló artesanal com 40 anos de receita encaixa na narrativa histórica.","zona":"Lisboa (Rossio)"},
  {"n":"Gambrinus","t":"Restaurante Clássico","m":"Rua das Portas de Santo Antão 23, 1150-264 Lisboa","tel":"213 421 466","email":"geral@restaurantegambrinus.com","p":"Alta","tCliente":"Restaurante Clássico Premium","gancho":"Referência de Lisboa há décadas — sobremesa de qualidade para clientela exigente.","zona":"Lisboa (Rossio)"},
  {"n":"Restaurante 1300 Taberna","t":"Tasca Moderna","m":"Rua de Alcântara 1300, 1300-023 Lisboa","tel":"213 649 170","email":"reservas@1300taberna.com","p":"Alta","tCliente":"Tasca Moderna / Ribeirinha","gancho":"Zona ribeirinha em crescimento — produto artesanal nacional num espaço com muita identidade.","zona":"Lisboa (Alcântara)"},
  {"n":"Rio Maravilha","t":"Restaurante & Rooftop","m":"Rua de Alcântara (LX Factory), Lisboa","tel":"213 932 930","email":"info@riomaravilha.pt","p":"Alta","tCliente":"Restaurante / Rooftop Bar","gancho":"LX Factory — público jovem e criativo, turistas. Sobremesa artesanal como produto de storytelling.","zona":"Lisboa (LX Factory)"},
  {"n":"A Taberna da Rua das Flores","t":"Tasca Tradicional","m":"Rua das Flores 103, 1200-193 Lisboa","tel":"213 479 418","email":"ataberndaruadasflores@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Referência no Chiado — menu curto de qualidade, sobremesa convencional de autor.","zona":"Lisboa (Chiado)"},
  {"n":"Taberna da Rua do Açúcar","t":"Tasca Moderna","m":"Rua do Açúcar 83A, 1950-006 Lisboa","tel":"213 910 651","email":"reservas@tabernacasadilnot.com","p":"Alta","tCliente":"Tasca Moderna","gancho":"Marvila em crescimento — clientela criativa e gastronómica que valoriza produto artesanal.","zona":"Lisboa (Marvila)"},
  {"n":"Café de Campo","t":"Café & Restaurante","m":"Praça do Campo Grande 28, 1700-093 Lisboa","tel":"217 932 932","email":"geral@cafedecampo.pt","p":"Média","tCliente":"Café & Almoço Executivo","gancho":"Zona universitária e empresarial — lanche de qualidade e sobremesa de almoço de alto valor.","zona":"Lisboa (Campo Grande)"},
  {"n":"Café Garrett","t":"Café Histórico","m":"Largo do Chiado 1-2, 1200-108 Lisboa","tel":"213 460 540","email":"geral@cafegarrett.pt","p":"Alta","tCliente":"Café Histórico / Turismo","gancho":"Ícone do Chiado — pão de ló artesanal centenário encaixa na narrativa histórica do espaço.","zona":"Lisboa (Chiado)"},
  {"n":"Pátio do Petisco","t":"Tasca de Petiscos","m":"Calçada do Combro 45, 1200-113 Lisboa","tel":"213 460 121","email":"patiodopetisco@gmail.com","p":"Alta","tCliente":"Tasca de Petiscos","gancho":"Tasca de vizinhança com clientela fiel — sobremesa simples, boa e bem portuguesa.","zona":"Lisboa (Bica)"},
  {"n":"Taberna do Largo","t":"Tasca Contemporânea","m":"Rua de Santana à Lapa 1, 1200-432 Lisboa","tel":"213 932 622","email":"tabernadolargo@gmail.com","p":"Alta","tCliente":"Tasca Contemporânea","gancho":"Lapa residencial e turística — menu de qualidade com produto nacional em destaque.","zona":"Lisboa (Lapa)"},
  {"n":"Enoteca Chafariz do Vinho","t":"Enoteca & Restaurante","m":"Rua da Mãe de Água 1, 1250-161 Lisboa","tel":"213 422 079","email":"chafariz@chafarizdevinho.com","p":"Alta","tCliente":"Enoteca Premium","gancho":"Maridagem com vinhos generosos — pão de ló com Moscatel de Setúbal é proposta vencedora.","zona":"Lisboa (Amoreiras)"},
  {"n":"O Pitéu da Graça","t":"Tasca Tradicional","m":"Praça da Graça 96, 1170-165 Lisboa","tel":"218 870 150","email":"opiteudagraca@gmail.com","p":"Média","tCliente":"Tasca Tradicional","gancho":"Vizinhança de bairro em ascensão — sobremesa portuguesa clássica de fecho.","zona":"Lisboa (Graça)"},
  {"n":"Café Mexicana","t":"Pastelaria Clássica","m":"Avenida de Roma 79A, 1700-345 Lisboa","tel":"217 965 256","email":"cafemexicana@cafemexicana.pt","p":"Alta","tCliente":"Pastelaria Clássica","gancho":"Pastelaria de referência de Lisboa há décadas — pão de ló artesanal diferencia a oferta.","zona":"Lisboa (Roma-Areeiro)"},
  {"n":"Pastelataria","t":"Pastelaria Moderna","m":"Rua Actor Taborda 61A, 1000-001 Lisboa","tel":"—","email":"geral@pastelataria.pt","p":"Alta","tCliente":"Pastelaria Moderna","gancho":"Conceito moderno com produto de qualidade — pão de ló Ti'Piedade como âncora da oferta tradicional.","zona":"Lisboa (Intendente)"},
  {"n":"Dois Palitos","t":"Restaurante Português Moderno","m":"Rua do Grilo 45, 1900-706 Lisboa","tel":"218 684 034","email":"doispalitos@gmail.com","p":"Alta","tCliente":"Restaurante Português Moderno","gancho":"Junto ao Museu do Azulejo — turistas e residentes exigentes. Sobremesa artesanal com narrativa.","zona":"Lisboa (Xabregas)"},
  {"n":"Restaurante Os Tibetanos","t":"Restaurante Vegetariano","m":"Rua do Salitre 117, 1269-063 Lisboa","tel":"213 142 038","email":"tibetanos@tibetanos.com","p":"Média","tCliente":"Restaurante Vegetariano","gancho":"Produto sem conservantes — pão de ló artesanal enquadra-se na filosofia de produto natural.","zona":"Lisboa (Avenidas Novas)"},
  {"n":"Zé da Venda","t":"Mercearia & Petiscos","m":"Rua da Madalena 187, 1100-321 Lisboa","tel":"218 867 727","email":"zedavenda@gmail.com","p":"Alta","tCliente":"Mercearia Gourmet","gancho":"Venda a fatia e por unidade — margem elevada em produto artesanal de qualidade.","zona":"Lisboa (Baixa)"},
  {"n":"Comida de Santo","t":"Restaurante Brasileiro","m":"Calçada Engenheiro Miguel Pais 39, 1200-352 Lisboa","tel":"213 963 339","email":"info@comidadesanto.com","p":"Alta","tCliente":"Restaurante Internacional","gancho":"Público brasileiro valoriza bolo de textura fofa — pão de ló é muito próximo do bolo formigueiro.","zona":"Lisboa (Rato)"},
  {"n":"A Baiuca","t":"Casa de Fado","m":"Rua de São Miguel 20, 1100-543 Lisboa","tel":"218 867 284","email":"abaiuca.fado@gmail.com","p":"Alta","tCliente":"Casa de Fado / Turismo","gancho":"Experiência cultural portuguesa completa — sobremesa artesanal com história fecha a noite de fado.","zona":"Lisboa (Alfama)"},
  {"n":"Mesa de Frades","t":"Casa de Fado Premium","m":"Rua dos Remédios 139A, 1100-461 Lisboa","tel":"917 029 436","email":"mesadefrades@gmail.com","p":"Alta","tCliente":"Casa de Fado Premium","gancho":"Espaço premium em antiga ermida — sobremesa artesanal convencional encaixa na experiência única.","zona":"Lisboa (Alfama)"},
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
  # ── João — Batch 2 (40 leads) ──────────────────────────────────────────
  {"n":"Tasca do Chico","t":"Casa de Fado / Tasca","m":"Rua do Diário de Notícias 39, 1200-145 Lisboa","tel":"213 424 759","email":"tascadochico@gmail.com","p":"Alta","tCliente":"Casa de Fado / Tasca","gancho":"Fado e gastronomia tradicional. Público internacional exige sobremesa portuguesa autêntica.","zona":"Lisboa (Bairro Alto)"},
  {"n":"Solar dos Presuntos","t":"Restaurante Tradicional Premium","m":"Rua das Portas de Santo Antão 150, 1150-269 Lisboa","tel":"213 424 253","email":"geral@solardospresuntos.com","p":"Alta","tCliente":"Restaurante Tradicional Premium","gancho":"Referência gastronómica em Lisboa. Sobremesa artesanal completa uma experiência premium.","zona":"Lisboa (Rossio)"},
  {"n":"O Corvo","t":"Taberna Contemporânea","m":"Calçada do Duque 5, 1200-159 Lisboa","tel":"213 427 106","email":"ocorvo@taberna.pt","p":"Alta","tCliente":"Taberna Contemporânea","gancho":"Taberna moderna com produto de qualidade. Pão de ló artesanal como sobremesa âncora.","zona":"Lisboa (Chiado)"},
  {"n":"Terra","t":"Restaurante Vegetariano / Biológico","m":"Rua da Palmeira 15, 1200-313 Lisboa","tel":"213 421 407","email":"info@terra.com.pt","p":"Média","tCliente":"Restaurante Vegetariano","gancho":"Ingredientes sem conservantes. Pão de ló artesanal alinha com filosofia de produto limpo.","zona":"Lisboa (Príncipe Real)"},
  {"n":"Tascardoso","t":"Tasca Moderna","m":"Rua de São Paulo 226, 1200-427 Lisboa","tel":"213 421 831","email":"tascardoso@gmail.com","p":"Alta","tCliente":"Tasca Moderna","gancho":"Clientela local fiel. Sobremesa portuguesa artesanal para fechar refeição com qualidade.","zona":"Lisboa (Cais do Sodré)"},
  {"n":"Bettina & Niccolò Corallo","t":"Loja de Chocolate / Café","m":"Rua do Tomás Ribeiro 63, 1050-228 Lisboa","tel":"213 860 536","email":"info@corallo.pt","p":"Alta","tCliente":"Loja Gourmet","gancho":"Produto premium e artesanal. Pão de ló Ti'Piedade complementa a oferta de doçaria fina.","zona":"Lisboa (Avenidas Novas)"},
  {"n":"Restaurante 100 Maneiras","t":"Restaurante Autor","m":"Largo da Academia Nacional de Belas Artes 1, 1200-005 Lisboa","tel":"910 307 575","email":"reservas@100maneiras.com","p":"Alta","tCliente":"Restaurante Autor","gancho":"Cozinha de autor. Propor pão de ló como elemento de sobremesa em interpretação contemporânea.","zona":"Lisboa (Bairro Alto)"},
  {"n":"Pharmácia","t":"Restaurante Temático","m":"Rua Marechal Saldanha 1, 1200-396 Lisboa","tel":"213 462 146","email":"pharmacia@pharmacia.pt","p":"Alta","tCliente":"Restaurante Temático","gancho":"Temática histórica. Sobremesa artesanal com narrativa centenária convence turistas.","zona":"Lisboa (Santa Catarina)"},
  {"n":"Sea Me","t":"Restaurante de Peixe Moderno","m":"Rua do Loreto 21, 1200-240 Lisboa","tel":"213 461 564","email":"info@seame.pt","p":"Alta","tCliente":"Restaurante de Peixe Moderno","gancho":"Peixe de qualidade. Sobremesa tradicional como contraponto ao menu de peixe sofisticado.","zona":"Lisboa (Chiado)"},
  {"n":"A Cevicheria","t":"Restaurante Latino-Moderno","m":"Rua Dom Pedro V 129, 1269-103 Lisboa","tel":"218 038 815","email":"info@acevicheria.pt","p":"Alta","tCliente":"Restaurante Latino-Moderno","gancho":"Conceito moderno. Oferecer pão de ló artesanal como surpresa de sobremesa portuguesa.","zona":"Lisboa (Príncipe Real)"},
  {"n":"Flores do Bairro","t":"Restaurante","m":"Rua das Flores 64, 1200-194 Lisboa","tel":"213 420 190","email":"floresdobairro@gmail.com","p":"Alta","tCliente":"Restaurante","gancho":"Zona histórica premium. Sobremesa artesanal enriquece experiência gastronómica local.","zona":"Lisboa (Chiado)"},
  {"n":"Sal Grosso","t":"Tasca Contemporânea","m":"Rua da Moeda 1, 1200-307 Lisboa","tel":"213 475 462","email":"salgrosso@gmail.com","p":"Alta","tCliente":"Tasca Contemporânea","gancho":"Perto do mercado da Ribeira. Produto artesanal local com história é apreciado pelo público.","zona":"Lisboa (Cais do Sodré)"},
  {"n":"Taberna da Rua das Flores","t":"Taberna Tradicional","m":"Rua das Flores 103, 1200-195 Lisboa","tel":"213 479 418","email":"tabernaderua@gmail.com","p":"Alta","tCliente":"Taberna Tradicional","gancho":"Petiscos e vinhos com clientela fiel. Sobremesa artesanal para fechar a conta.","zona":"Lisboa (Chiado)"},
  {"n":"Sacramento do Chiado","t":"Restaurante","m":"Calçada Sacramento 40, 1200-394 Lisboa","tel":"213 420 572","email":"sacramento@sacramento.pt","p":"Alta","tCliente":"Restaurante","gancho":"Zona premium. Pão de ló artesanal Ti'Piedade como sobremesa destaque na carta.","zona":"Lisboa (Chiado)"},
  {"n":"Tasca do Miraldo","t":"Tasca Tradicional","m":"Rua das Flores 22, 1200-192 Lisboa","tel":"213 420 890","email":"tascamiraldo@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Clientela local fiel que valoriza produto nacional autêntico.","zona":"Lisboa (Chiado)"},
  {"n":"Cervejaria Liberdade","t":"Cervejaria / Marisco","m":"Avenida da Liberdade 185, 1269-053 Lisboa","tel":"213 512 620","email":"info@cervejarialiberdade.pt","p":"Alta","tCliente":"Cervejaria / Marisco","gancho":"Grande volume de clientes. Sobremesa congelada com zero quebra para fim de refeição.","zona":"Lisboa (Avenida da Liberdade)"},
  {"n":"Eleven","t":"Restaurante Fine Dining","m":"Rua Marquês de Fronteira, Jardim Amália Rodrigues, 1070-296 Lisboa","tel":"213 862 211","email":"eleven@eleven.pt","p":"Alta","tCliente":"Restaurante Fine Dining","gancho":"Restaurante de topo. Propor mini-dose artesanal de pão de ló como mignardise premium.","zona":"Lisboa (Parque Eduardo VII)"},
  {"n":"Loco","t":"Restaurante Contemporâneo","m":"Rua dos Navegantes 53B, 1200-727 Lisboa","tel":"213 951 861","email":"loco@loco.pt","p":"Alta","tCliente":"Restaurante Contemporâneo","gancho":"Conceito contemporâneo. Pão de ló Ti'Piedade reinterpretado em menu de degustação.","zona":"Lisboa (Alcântara)"},
  {"n":"Cantinho do Avillez","t":"Restaurante Contemporâneo","m":"Rua dos Duques de Bragança 7, 1200-162 Lisboa","tel":"211 992 369","email":"cantinho@joseavillez.pt","p":"Alta","tCliente":"Restaurante Contemporâneo","gancho":"Assinatura José Avillez. Apresentar consistência e qualidade do produto artesanal.","zona":"Lisboa (Chiado)"},
  {"n":"Restaurante ZéZé","t":"Restaurante Tradicional","m":"Rua Cândido dos Reis 8, 2800-121 Almada","tel":"212 760 432","email":"restaurantezeze@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Almada clássico. Clientela local que aprecia produto português de qualidade.","zona":"Almada"},
  {"n":"A Floresta","t":"Pastelaria Artesanal","m":"Rua Alexandre Herculano 12, 2900-153 Setúbal","tel":"265 523 140","email":"pastafloresta@gmail.com","p":"Alta","tCliente":"Pastelaria Artesanal","gancho":"Pastelaria de referência local. Pão de ló Ti'Piedade complementa oferta de pastelaria artesanal.","zona":"Setúbal"},
  {"n":"Café Âncora d'Ouro","t":"Café Tradicional","m":"Praça de Bocage 35, 2900-301 Setúbal","tel":"265 523 730","email":"ancoraouro@gmail.com","p":"Alta","tCliente":"Café Tradicional","gancho":"Praça principal de Setúbal. Ponto de passagem com grande volume de clientes diários.","zona":"Setúbal"},
  {"n":"Tasca do Camilo","t":"Tasca Tradicional","m":"Largo do Rio Sado 5, 2900-441 Setúbal","tel":"265 239 130","email":"tascacamilo@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Peixe e marisco fresco. Sobremesa portuguesa artesanal como fecho natural da refeição.","zona":"Setúbal"},
  {"n":"Restaurante Bocage","t":"Restaurante Tradicional","m":"Rua de São Sebastião 51, 2900-399 Setúbal","tel":"265 235 760","email":"rbocage@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Restaurante histórico de Setúbal. Doçaria artesanal para fechar menu de peixe fresco.","zona":"Setúbal"},
  {"n":"Casa da Baía","t":"Restaurante Panorâmico","m":"Estrada Nacional 252, 2925-506 Azeitão","tel":"212 198 562","email":"casadabaia@gmail.com","p":"Alta","tCliente":"Restaurante Panorâmico","gancho":"Vista para a Arrábida. Experiência premium que justifica sobremesa artesanal de qualidade.","zona":"Azeitão"},
  {"n":"Quinta da Bacalhôa","t":"Restaurante & Enoteca","m":"Azeitão, 2925-001 Azeitão","tel":"212 198 018","email":"enoteca@bacalhoa.pt","p":"Alta","tCliente":"Restaurante & Enoteca","gancho":"Enoturismo e gastronomia. Pão de ló com Moscatel de Setúbal é proposta ganhadora.","zona":"Azeitão"},
  {"n":"Tasca do Zé Bettencourt","t":"Tasca Tradicional","m":"Rua das Flores 10, 2830-302 Barreiro","tel":"212 078 430","email":"tascazebettencourt@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Tasca clássica do Barreiro. Clientela local fidelizada aprecia produto nacional autêntico.","zona":"Barreiro"},
  {"n":"Snack Bar Aquário","t":"Restaurante Tradicional","m":"Avenida 25 de Abril 45, 2830-152 Barreiro","tel":"212 076 520","email":"aquario.barreiro@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Refeições rápidas e menu diário. Sobremesa congelada com zero quebra ideal para rotação alta.","zona":"Barreiro"},
  {"n":"O Cortiço","t":"Restaurante Tradicional","m":"Rua Capitão Leitão 15, 2800-149 Almada","tel":"212 760 890","email":"ocortico.almada@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Restaurante de almoços executivos. Sobremesa portuguesa artesanal que fideliza clientela.","zona":"Almada"},
  {"n":"Casa do Pão de Ló de Alfeizerão","t":"Pastelaria Regional","m":"Rua 25 de Abril 12, 2460-503 Alfeizerão","tel":"262 998 120","email":"paolodeloalfeizerao@gmail.com","p":"Alta","tCliente":"Pastelaria Regional","gancho":"Especialidade regional. Apresentar diferenciação Ti'Piedade como versão artesanal premium.","zona":"Caldas da Rainha"},
  {"n":"Restaurante Sabores do Ribatejo","t":"Restaurante Regional","m":"Rua da Estação 23, 2000-100 Santarém","tel":"243 323 540","email":"saboresribatejo@gmail.com","p":"Alta","tCliente":"Restaurante Regional","gancho":"Gastronomia ribatejana autêntica. Pão de ló Ti'Piedade como sobremesa histórica da região.","zona":"Santarém"},
  {"n":"Cervejaria Bairro Alto","t":"Cervejaria Tradicional","m":"Rua do Norte 86, 1200-281 Lisboa","tel":"213 420 145","email":"cervejariabairo@gmail.com","p":"Alta","tCliente":"Cervejaria Tradicional","gancho":"Grande rotação de clientes no coração de Lisboa. Sobremesa simples e lucrativa.","zona":"Lisboa (Bairro Alto)"},
  {"n":"Café Buenos Aires","t":"Café Histórico","m":"Calçada do Combro 57, 1200-113 Lisboa","tel":"213 420 739","email":"cafebuenosaires@gmail.com","p":"Alta","tCliente":"Café Histórico","gancho":"Café tradicional de Lisboa. Pão de ló artesanal na vitrine atrai clientela local e turistas.","zona":"Lisboa (Bica)"},
  {"n":"Tasca do Cherne","t":"Restaurante de Peixe","m":"Travessa da Ermida 28, 2800-163 Almada","tel":"212 763 401","email":"tascadocherne@gmail.com","p":"Alta","tCliente":"Restaurante de Peixe","gancho":"Especialidade em peixe. Sobremesa artesanal portuguesa fecha o menu com valor acrescentado.","zona":"Almada"},
  {"n":"Café Versailles","t":"Pastelaria / Café Clássico","m":"Avenida da República 15A, 1050-185 Lisboa","tel":"213 546 340","email":"cafeversailles@gmail.com","p":"Alta","tCliente":"Pastelaria Clássica","gancho":"Ícone de Lisboa. Pão de ló artesanal como peça de destaque na vitrine clássica.","zona":"Lisboa (Saldanha)"},
  {"n":"Tasca do Pombal","t":"Tasca Tradicional","m":"Rua Marquês de Pombal 8, 2830-323 Barreiro","tel":"212 074 891","email":"tascapombal@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Tasca de bairro com clientela fiel. Produto português artesanal sempre bem recebido.","zona":"Barreiro"},
  {"n":"Restaurante Horizonte","t":"Restaurante Panorâmico","m":"Avenida dos Náufragos 120, 2825-421 Costa da Caparica","tel":"212 902 456","email":"rhorizonte@gmail.com","p":"Alta","tCliente":"Restaurante Panorâmico","gancho":"Vista para o oceano. Experiência premium a pedir produto artesanal de qualidade.","zona":"Costa da Caparica"},
  {"n":"Marisqueira Rui","t":"Marisqueira / Cervejaria","m":"Rua Capitão Salgueiro Maia 12, 2840-501 Seixal","tel":"212 232 890","email":"marisqueirarui@gmail.com","p":"Alta","tCliente":"Marisqueira / Cervejaria","gancho":"Marisco e peixe fresco. Doçaria artesanal portuguesa de fecho de refeição com custo controlado.","zona":"Seixal"},
  {"n":"Tasca da Margem","t":"Tasca Tradicional","m":"Rua da Bela Vista 44, 2840-390 Corroios","tel":"212 569 341","email":"tascadamargem@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Bairro residencial. Clientela regular que valoriza produto nacional artesanal.","zona":"Seixal (Corroios)"},
  {"n":"Restaurante Azul do Mar","t":"Restaurante de Peixe","m":"Rua da Praia 56, 2825-484 Trafaria","tel":"212 951 230","email":"azuldomar@gmail.com","p":"Alta","tCliente":"Restaurante de Peixe","gancho":"Peixe fresco de Trafaria. Sobremesa artesanal complementa menu de mar com valor acrescentado.","zona":"Trafaria"},
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
  # ── Óscar — Batch 2 (40 leads) ──────────────────────────────────────────
  {"n":"Café Guarany","t":"Café / Restaurante Histórico","m":"Avenida dos Aliados 85-89, 4000-067 Porto","tel":"222 010 272","email":"cafeguarany@cafeguarany.com","p":"Alta","tCliente":"Café Histórico","gancho":"Ícone do Porto. Pão de ló artesanal Ti'Piedade como sobremesa histórica portuguesa.","zona":"Porto (Aliados)"},
  {"n":"Zenith Brunch & Cocktails","t":"Brunch Bar","m":"Rua do Almada 390, 4050-038 Porto","tel":"220 171 180","email":"porto@zenithbrunch.com","p":"Alta","tCliente":"Brunch Bar","gancho":"Brunch premium com cocktails. Fatia de pão de ló como opção doce portuguesa no menu.","zona":"Porto"},
  {"n":"Café Santiago","t":"Cervejaria / Restaurante","m":"Rua de Passos Manuel 226, 4000-382 Porto","tel":"222 055 797","email":"cafesantiago@cafesantiago.pt","p":"Alta","tCliente":"Cervejaria","gancho":"Casa da francesinha mais famosa. Grande volume — sobremesa congelada com zero quebra.","zona":"Porto"},
  {"n":"Taberna dos Frades","t":"Taberna / Restaurante Tradicional","m":"Rua de Belomonte 89, 4050-093 Porto","tel":"222 088 380","email":"tabernadosfrades@gmail.com","p":"Alta","tCliente":"Taberna Tradicional","gancho":"Espaço com história e charme. Sobremesa portuguesa artesanal para fechar refeição.","zona":"Porto (Ribeira)"},
  {"n":"Pedro Lemos","t":"Restaurante Fine Dining","m":"Rua do Padre Luís Cabral 974, 4150-459 Porto","tel":"220 115 986","email":"restaurante@pedrolemos.net","p":"Alta","tCliente":"Restaurante Fine Dining","gancho":"Único Michelin no Porto. Propor pão de ló artesanal como elemento de sobremesa autoral.","zona":"Porto (Foz)"},
  {"n":"Restaurante DOP","t":"Restaurante Premium","m":"Largo de São Domingos 18, 4050-545 Porto","tel":"222 014 313","email":"dop@ruipaula.com","p":"Alta","tCliente":"Restaurante Premium","gancho":"Rui Paula no Porto. Consistência e qualidade do produto artesanal Ti'Piedade.","zona":"Porto"},
  {"n":"Wine Quay Bar","t":"Wine Bar","m":"Rua de Monchique 111, 4050-393 Porto","tel":"222 000 333","email":"winequaybar@gmail.com","p":"Alta","tCliente":"Wine Bar","gancho":"Vinhos do Porto e Douro. Pão de ló com vinho generoso é maridagem vencedora.","zona":"Porto (Ribeira)"},
  {"n":"Flor dos Congregados","t":"Tasca Tradicional","m":"Travessa dos Congregados 11, 4000-096 Porto","tel":"222 002 822","email":"flordoscongregados@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Tasca clássica com grande tradição. Sobremesa portuguesa artesanal de fecho natural.","zona":"Porto"},
  {"n":"Museu d'Arte e Pastelaria","t":"Café / Pastelaria Artística","m":"Rua do Miguel Bombarda 94, 4050-379 Porto","tel":"222 010 561","email":"museudarte@gmail.com","p":"Alta","tCliente":"Café / Pastelaria Artística","gancho":"Zona de galerias de arte. Pão de ló artesanal Ti'Piedade como objeto de culinária local.","zona":"Porto (Miguel Bombarda)"},
  {"n":"Restaurante Tripeiro","t":"Restaurante Tradicional","m":"Rua Passos Manuel 195, 4000-382 Porto","tel":"222 005 886","email":"tripeiro@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Especialidade portuense. Sobremesa portuguesa artesanal como fecho de menu tradicional.","zona":"Porto"},
  {"n":"Casa Branca","t":"Restaurante Contemporâneo","m":"Rua de Cedofeita 560, 4050-181 Porto","tel":"222 074 891","email":"casabranca.porto@gmail.com","p":"Alta","tCliente":"Restaurante Contemporâneo","gancho":"Menu contemporâneo criativo. Pão de ló artesanal como contraste clássico na carta.","zona":"Porto"},
  {"n":"Casanova","t":"Restaurante Italiano / Internacional","m":"Avenida da Boavista 961, 4100-128 Porto","tel":"226 177 176","email":"casanova.porto@gmail.com","p":"Alta","tCliente":"Restaurante Internacional","gancho":"Grande volume. Sobremesa artesanal portuguesa em alternativa ao tiramisù clássico.","zona":"Porto (Boavista)"},
  {"n":"Café Majestic","t":"Café Histórico","m":"Rua de Santa Catarina 112, 4000-442 Porto","tel":"222 003 887","email":"cafemajestic@cafemajestic.com","p":"Alta","tCliente":"Café Histórico","gancho":"Ícone do Porto. Pão de ló artesanal encaixa na narrativa histórica centenária do espaço.","zona":"Porto (Santa Catarina)"},
  {"n":"A Grade","t":"Taberna / Restaurante","m":"Rua do Bonjardim 572, 4000-126 Porto","tel":"222 008 133","email":"agrade.porto@gmail.com","p":"Alta","tCliente":"Taberna Tradicional","gancho":"Almoços executivos. Sobremesa artesanal rápida e lucrativa de fecho de refeição.","zona":"Porto"},
  {"n":"Tasca do Ze","t":"Tasca Tradicional","m":"Rua da Boa Hora 72, 4050-093 Porto","tel":"222 003 456","email":"tascadoze@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Tasca de bairro histórico. Clientela fiel que valoriza produto português artesanal.","zona":"Porto (Ribeira)"},
  {"n":"Adega Velha","t":"Adega / Restaurante","m":"Rua de São João 67, 4050-530 Porto","tel":"222 056 781","email":"adegavelha.porto@gmail.com","p":"Alta","tCliente":"Adega Tradicional","gancho":"Vinhos a copo e refeição. Pão de ló com vinho do Porto é proposta premium de fecho.","zona":"Porto"},
  {"n":"Confeitaria do Bolhão","t":"Confeitaria / Pastelaria","m":"Rua Formosa 339, 4000-253 Porto","tel":"222 325 800","email":"confeitariabolhao@gmail.com","p":"Alta","tCliente":"Pastelaria Histórica","gancho":"Pastelaria tradicional do Porto. Pão de ló artesanal Ti'Piedade como opção de qualidade.","zona":"Porto (Bolhão)"},
  {"n":"Portucale","t":"Restaurante Panorâmico","m":"Rua da Alegria 598, 4000-040 Porto","tel":"225 370 717","email":"portucale@gmail.com","p":"Alta","tCliente":"Restaurante Panorâmico","gancho":"Vista panorâmica do Porto. Sobremesa premium que eleva a experiência gastronómica.","zona":"Porto"},
  {"n":"Casa de Pasto","t":"Restaurante Tradicional","m":"Rua de Cândido dos Reis 89, 4050-152 Porto","tel":"222 006 543","email":"casadepasto.porto@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Cozinha portuguesa tradicional. Pão de ló artesanal como sobremesa âncora da carta.","zona":"Porto"},
  {"n":"Antiqvm","t":"Restaurante Premium","m":"Rua de Entre-Quintas 220, 4050-240 Porto","tel":"226 009 190","email":"antiqvm@gmail.com","p":"Alta","tCliente":"Restaurante Premium","gancho":"Jardim histórico. Sobremesa artesanal de referência para público exigente.","zona":"Porto (Foz)"},
  {"n":"Leça da Palmeira - O Chanquinhas","t":"Restaurante Tradicional","m":"Rua da Caçadores 86, 4450-243 Leça da Palmeira","tel":"229 952 229","email":"chanquinhas@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Clássico da Foz. Peixe e marisco com doçaria artesanal portuguesa de fecho.","zona":"Leça da Palmeira"},
  {"n":"O Pescador Matosinhos","t":"Restaurante de Peixe","m":"Rua Heróis de França 230, 4450-162 Matosinhos","tel":"229 392 027","email":"opescador.matosinhos@gmail.com","p":"Alta","tCliente":"Restaurante de Peixe","gancho":"Especialistas em peixe fresco. Sobremesa portuguesa de fecho com custo controlado.","zona":"Matosinhos"},
  {"n":"Laviña","t":"Restaurante Tradicional","m":"Rua Herois de França 244, 4450-162 Matosinhos","tel":"229 382 720","email":"lavina@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Restaurante de referência em Matosinhos. Pão de ló artesanal complementa carta de peixe.","zona":"Matosinhos"},
  {"n":"Sr. Peixe","t":"Restaurante de Peixe","m":"Rua de Augusto Gomes 83, 4450-070 Matosinhos","tel":"229 375 240","email":"srpeixe@gmail.com","p":"Alta","tCliente":"Restaurante de Peixe","gancho":"Marisco e peixe fresco. Sobremesa artesanal portuguesa como alternativa diferenciadora.","zona":"Matosinhos"},
  {"n":"Taberna da Esquina (Coimbra)","t":"Taberna Contemporânea","m":"Rua das Azeiteiras 65, 3000-050 Coimbra","tel":"239 824 561","email":"tabernadadaesquina@gmail.com","p":"Alta","tCliente":"Taberna Contemporânea","gancho":"Cozinha contemporânea em Coimbra. Pão de ló artesanal como elemento da tradição regional.","zona":"Coimbra"},
  {"n":"Ze Manel dos Ossos","t":"Tasca Tradicional","m":"Beco do Forno 12, 3000-078 Coimbra","tel":"239 823 790","email":"zemaneldososos@gmail.com","p":"Alta","tCliente":"Tasca Tradicional","gancho":"Tasca histórica de Coimbra. Sobremesa artesanal para fechar refeição com produto local.","zona":"Coimbra"},
  {"n":"Restaurante Democrática","t":"Restaurante Tradicional","m":"Travessa da Rua Nova 7, 3000-329 Coimbra","tel":"239 823 784","email":"rdemocrática@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Cozinha tradicional da região de Coimbra. Pão de ló artesanal com história local.","zona":"Coimbra"},
  {"n":"O Trovador","t":"Restaurante / Fado","m":"Largo da Sé Velha 15, 3000-383 Coimbra","tel":"239 825 475","email":"otrovador@gmail.com","p":"Alta","tCliente":"Restaurante / Fado","gancho":"Fado de Coimbra com jantar. Sobremesa portuguesa artesanal fecha experiência cultural.","zona":"Coimbra"},
  {"n":"Restaurante Jardim da Manga","t":"Restaurante Tradicional","m":"Rua Olímpio Nicolau Fernandes 9, 3000-295 Coimbra","tel":"239 820 156","email":"jardimdamanga@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Zona histórica de Coimbra. Produto artesanal regional para clientela local e turistas.","zona":"Coimbra"},
  {"n":"Café Santa Cruz","t":"Café Histórico","m":"Praça 8 de Maio 1, 3000-300 Coimbra","tel":"239 833 617","email":"cafesantacruz@gmail.com","p":"Alta","tCliente":"Café Histórico","gancho":"Café histórico no ex-convento. Pão de ló artesanal como produto histórico regional.","zona":"Coimbra"},
  {"n":"Pastelaria Briosa","t":"Pastelaria Histórica","m":"Rua Padre António Vieira 7, 3000-339 Coimbra","tel":"239 826 834","email":"briosa@pastelariabriosa.pt","p":"Alta","tCliente":"Pastelaria Histórica","gancho":"Pastelaria de referência em Coimbra. Complementar a oferta com pão de ló artesanal premium.","zona":"Coimbra"},
  {"n":"Tasca do Sete","t":"Tasca / Petiscos","m":"Rua do Sete 7, 3000-390 Coimbra","tel":"239 840 120","email":"tascadosete@gmail.com","p":"Alta","tCliente":"Tasca / Petiscos","gancho":"Petiscos e vinho. Pão de ló artesanal Ti'Piedade como toque final da refeição.","zona":"Coimbra"},
  {"n":"Restaurante O Açude","t":"Restaurante Tradicional","m":"Mata Nacional do Buçaco, 3050-261 Luso","tel":"231 937 937","email":"oacude.bussaco@gmail.com","p":"Alta","tCliente":"Restaurante Turístico","gancho":"Mata do Buçaco. Produto artesanal histórico complementa experiência de enoturismo.","zona":"Buçaco / Luso"},
  {"n":"A Toca do Senhor Vinho","t":"Adega / Taberna","m":"Rua de Ferreira de Castro 12, 3800-125 Aveiro","tel":"234 424 890","email":"toca.senhirvinho@gmail.com","p":"Alta","tCliente":"Adega Tradicional","gancho":"Vinhos a copo e petiscos. Pão de ló com vinho generoso é proposta de fecho premium.","zona":"Aveiro"},
  {"n":"Restaurante O Mercado (Aveiro)","t":"Restaurante Contemporâneo","m":"Rua de Coimbra 19, 3800-078 Aveiro","tel":"234 383 512","email":"restauranteomercado@gmail.com","p":"Alta","tCliente":"Restaurante Contemporâneo","gancho":"Restaurante contemporâneo em Aveiro. Pão de ló artesanal de Alcobaça com história regional.","zona":"Aveiro"},
  {"n":"Salpoente","t":"Restaurante Fine Dining","m":"Canal de São Roque 83, 3800-256 Aveiro","tel":"234 382 674","email":"salpoente@salpoente.com","p":"Alta","tCliente":"Restaurante Fine Dining","gancho":"Fine dining junto à Ria de Aveiro. Sobremesa artesanal como mignardise premium.","zona":"Aveiro"},
  {"n":"A Barraca","t":"Restaurante Tradicional","m":"Rua José Estêvão 54, 3800-168 Aveiro","tel":"234 422 116","email":"abarraca.aveiro@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Especialidades da Ria. Doçaria artesanal portuguesa para fechar menu de peixe e marisco.","zona":"Aveiro"},
  {"n":"Cervejaria Restaurante Jerónimos","t":"Cervejaria / Restaurante","m":"Rua de Belém 84, 1300-085 Lisboa","tel":"213 637 423","email":"jeronimos@cervejariaj.pt","p":"Alta","tCliente":"Cervejaria Turística","gancho":"Zona Belém com turistas. Sobremesa artesanal portuguesa como experiência autentica.","zona":"Lisboa (Belém)"},
  {"n":"Pastéis de Belém","t":"Pastelaria Histórica","m":"Rua de Belém 84-92, 1300-085 Lisboa","tel":"213 637 423","email":"comercial@pasteisdebelém.pt","p":"Alta","tCliente":"Pastelaria Histórica","gancho":"Ícone mundial. Propor pão de ló artesanal Ti'Piedade como produto histórico complementar.","zona":"Lisboa (Belém)"},
  {"n":"O Frade (Coimbra)","t":"Restaurante Tradicional","m":"Rua do Corvo 34, 3000-000 Coimbra","tel":"239 839 260","email":"ofrade.coimbra@gmail.com","p":"Alta","tCliente":"Restaurante Tradicional","gancho":"Zona histórica de Coimbra. Produto artesanal regional para fechar menu de almoço.","zona":"Coimbra"},
  {"n":"Restaurante Arcadas (Coimbra)","t":"Restaurante Fine Dining","m":"Rua da Sota 1, 3000-395 Coimbra","tel":"239 802 380","email":"arcadas@quintadaslagrimas.pt","p":"Alta","tCliente":"Restaurante Fine Dining","gancho":"Hotel 5 estrelas em Coimbra. Sobremesa artesanal premium para clientela exigente.","zona":"Coimbra"},
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
    override = os.environ.get("SEMANA_OVERRIDE", "")
    if override.strip().isdigit():
        return int(override.strip())
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
    # GITHUB_TOKEN é gerado automaticamente em cada run e nunca expira.
    # GH_PAT é fallback para compatibilidade com runs locais.
    token = os.environ.get("GITHUB_TOKEN","") or os.environ.get("GH_PAT","")
    if not token:
        print("[AVISO] GITHUB_TOKEN/GH_PAT não configurado — histórico não será guardado.")
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
    if result:
        print(f"✓ Histórico guardado", flush=True)
    else:
        print(f"✗ ERRO CRÍTICO: Histórico não guardado — verifica permissões do GITHUB_TOKEN", flush=True)
        raise SystemExit(1)

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
                "tipologia_cliente": normalizar_tipologia(lead.get("tCliente","")),
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
Rui Bernardes
Responsável HORECA · Pão de Ló Ti'Piedade
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
          <td style="padding:24px 36px 16px">
            <p style="margin:0;font-size:22px;font-weight:700;color:#C49A3C;letter-spacing:.04em;font-family:Georgia,serif">Pão de Ló</p>
            <p style="margin:2px 0 0;font-size:13px;font-weight:400;color:rgba(255,255,255,.75);letter-spacing:.12em;text-transform:uppercase;font-family:Georgia,serif">Ti&#8217;Piedade</p>
          </td>
          <td align="right" style="padding:28px 36px 20px;vertical-align:top">
            <span style="display:inline-block;background:rgba(255,255,255,.12);border:1px solid rgba(196,154,60,.4);border-radius:99px;padding:4px 14px;font-size:11px;font-weight:700;color:#C49A3C;letter-spacing:.08em;text-transform:uppercase">
              {ico} {etiq}
            </span>
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
              <strong style="color:#C49A3C">Rui Bernardes</strong> · Responsável HORECA Ti'Piedade
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


def brevo_send(to_email, to_name, subject, html_content, text_content=None, attachment_path=None):
    """Envia email via Brevo HTTP API. Retorna True se OK."""
    api_key = os.environ.get("BREVO_API_KEY","")
    if not api_key:
        print("[ERRO CRÍTICO] BREVO_API_KEY não configurada — nenhum email será enviado.", flush=True)
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

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            attachment_b64 = base64.b64encode(f.read()).decode()
        payload["attachment"] = [{
            "content": attachment_b64,
            "name": os.path.basename(attachment_path)
        }]

    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        "https://api.brevo.com/v3/smtp/email",
        data=body,
        headers={
            "api-key": api_key,
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST"
    )
    try:
        with urllib.request.urlopen(req) as r:
            return r.status in (200, 201)
    except urllib.error.HTTPError as e:
        body = e.read().decode()[:400]
        print(f"[Brevo] Erro HTTP {e.code}: {body}", flush=True)
        if e.code == 401:
            raise SystemExit(1)  # chave inválida — falha imediata
        return False


def enviar_nurturing_lead(lead, email_num, nome_comercial):
    """Envia o email de nurturing via Brevo directamente ao email da lead."""
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

    ok_total = False
    for dest_email in emails:
        ok = brevo_send(dest_email, lead.get("n",""), f"{assunto} | Ti'Piedade", html, texto)
        if ok:
            ok_total = True
    return ok_total


def brevo_send_resumo(para, assunto, corpo_texto, ficheiro=None):
    """Envia email de resumo/comercial via Brevo, com anexo opcional."""
    html_body = f"<pre style='font-family:Arial,sans-serif;font-size:13px'>{corpo_texto}</pre>"
    return brevo_send(para, para, assunto, html_body, corpo_texto, ficheiro)


def enviar_emails(ficheiro, sem, modo):
    if modo == "coordenador":
        historico, _ = ler_historico()
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
        ok = brevo_send_resumo(EMAIL_RUI,
                               f"[REVISÃO] Leads Semana {sem} — {total_horeca+10} contactos | Ti'Piedade",
                               corpo, ficheiro)
        print(f"{'✓' if ok else '✗'} Resumo enviado para {EMAIL_RUI}")

    elif modo == "comerciais":
        historico, sha = ler_historico()
        enviados_leads = 0
        sem_email = 0

        for com_id in ["nuno","joao","oscar"]:
            com = COMERCIAIS[com_id]
            leads = gerar_leads_horeca(com_id, sem)
            dest = com["email"]

            # Determinar email_num para esta semana
            email_num = next(
                (calcular_email_num(historico, com_id, l) for l in leads
                 if calcular_email_num(historico, com_id, l)),
                1
            )

            # ── Enviar nurturing directamente às leads ─────────────
            leads_com_email = 0
            for lead in leads:
                n = calcular_email_num(historico, com_id, lead)
                if n is None:
                    continue  # sequência completa
                ok = enviar_nurturing_lead(lead, n, com["nome"])
                if ok:
                    leads_com_email += 1
                    enviados_leads += 1
                    historico = registar_envios(historico, com_id, [lead], n, com["nome"])
                else:
                    sem_email += 1

            # ── Email resumo ao comercial ──────────────────────────
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
                ok = brevo_send_resumo(dest,
                                       f"Leads Semana {sem} — {len(leads)} contactos | Ti'Piedade",
                                       corpo_com, ficheiro)
                print(f"{'✓' if ok else '✗'} {com['nome']} — {leads_com_email} nurturing + resumo enviado")

        # ── Rui — catering ─────────────────────────────────────────
        leads_cat = gerar_leads_canal(DB_CATERING, "rui_catering", sem, 5)
        for lead in leads_cat:
            n = calcular_email_num(historico, "rui_catering", lead)
            if n:
                ok = enviar_nurturing_lead(lead, n, "Rui")
                if ok:
                    enviados_leads += 1
                    historico = registar_envios(historico, "rui_catering", [lead], n, "Rui")

        corpo_cat = f"""Olá,

Aqui estão os 5 leads de Catering & Eventos para a semana {sem} ({semana_datas()}).

O Email 1 de nurturing foi enviado automaticamente às leads com email disponível.
Para as restantes, o contacto é telefónico direto.

Ti'Piedade — Sistema de Prospeção
"""
        ok = brevo_send_resumo(EMAIL_RUI,
                               f"Leads Catering & Eventos — Semana {sem} | Ti'Piedade",
                               corpo_cat, ficheiro)
        print(f"{'✓' if ok else '✗'} Catering → Rui")

        # ── Rui — distribuidores ───────────────────────────────────
        leads_dist = gerar_leads_canal(DB_DISTRIBUIDORES, "rui_distribuidores", sem, 5)
        for lead in leads_dist:
            n = calcular_email_num(historico, "rui_distribuidores", lead)
            if n:
                ok = enviar_nurturing_lead(lead, n, "Rui")
                if ok:
                    enviados_leads += 1
                    historico = registar_envios(historico, "rui_distribuidores", [lead], n, "Rui")

        corpo_dist = f"""Olá,

Aqui estão os 5 leads de Distribuidores para a semana {sem} ({semana_datas()}).

O Email 1 de nurturing foi enviado automaticamente às leads com email disponível.

Ti'Piedade — Sistema de Prospeção
"""
        ok = brevo_send_resumo(EMAIL_RUI,
                               f"Leads Distribuidores — Semana {sem} | Ti'Piedade",
                               corpo_dist, ficheiro)
        print(f"{'✓' if ok else '✗'} Distribuidores → Rui")

        print(f"\n📧 Total nurturing enviado: {enviados_leads} emails às leads | {sem_email} sem email", flush=True)
        escrever_historico(historico, sha)

        # Falhar explicitamente se havia leads com email mas nenhum foi enviado
        if sem_email > 0 and enviados_leads == 0:
            print(f"[ERRO] {sem_email} leads com email disponível mas nenhum envio concluído.", flush=True)
            raise SystemExit(1)

# ════════════════════════════════════════════════════════════════

if __name__=="__main__":
    sem=semana_num(); modo=os.environ.get("MODO","coordenador")
    print(f"▶ Semana {sem} — {semana_datas()} — Modo: {modo}")
    historico,sha=ler_historico()
    ficheiro=criar_excel(sem,historico)
    print(f"✓ Excel: {ficheiro}")
    enviar_emails(ficheiro,sem,modo)
    print("✓ Concluído.")
