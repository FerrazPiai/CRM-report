#!/usr/bin/env python3
"""
Clusterização e higienização da base de leads — Camiseta Express.

Lê todas as planilhas exportadas da pasta do Drive (colocadas em ./dados_brutos),
padroniza telefones e e-mails, remove duplicados entre fontes, classifica cada
contato por segmento de empresa e por temperatura (data de entrada + engajamento)
e gera uma planilha única, separada por abas, pronta para importar no RD Station.

Uso:
    pip install pandas openpyxl xlsxwriter
    python3 processar_base.py            # lê ./dados_brutos e grava em ./saida

Veja README.md para a metodologia completa.
"""
from __future__ import annotations

import csv
import io
import hashlib
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
RAW = BASE_DIR / "dados_brutos"
OUT = BASE_DIR / "saida"
HOJE = pd.Timestamp("2026-09-30")          # data de referência do processamento
DDD_PADRAO = "11"                          # operação em São Paulo

# --------------------------------------------------------------------------------------
# Tabelas de apoio
# --------------------------------------------------------------------------------------
DDD_UF = {
    **{d: "SP" for d in "11 12 13 14 15 16 17 18 19".split()},
    **{d: "RJ" for d in "21 22 24".split()}, **{d: "ES" for d in "27 28".split()},
    **{d: "MG" for d in "31 32 33 34 35 37 38".split()},
    **{d: "PR" for d in "41 42 43 44 45 46".split()}, **{d: "SC" for d in "47 48 49".split()},
    **{d: "RS" for d in "51 53 54 55".split()}, "61": "DF", "62": "GO", "64": "GO",
    "63": "TO", "65": "MT", "66": "MT", "67": "MS", "68": "AC", "69": "RO",
    **{d: "BA" for d in "71 73 74 75 77".split()}, "79": "SE", "81": "PE", "87": "PE",
    "82": "AL", "83": "PB", "84": "RN", "85": "CE", "88": "CE", "86": "PI", "89": "PI",
    "91": "PA", "93": "PA", "94": "PA", "92": "AM", "97": "AM", "95": "RR", "96": "AP",
    "98": "MA", "99": "MA",
}

# Provedores de e-mail pessoal (não identificam empresa)
FREE_DOMAINS = {
    "gmail.com", "googlemail.com", "hotmail.com", "hotmail.com.br", "outlook.com",
    "outlook.com.br", "live.com", "live.com.br", "msn.com", "yahoo.com", "yahoo.com.br",
    "ymail.com", "rocketmail.com", "uol.com.br", "bol.com.br", "terra.com.br", "ig.com.br",
    "icloud.com", "me.com", "mac.com", "globo.com", "globomail.com", "r7.com", "zipmail.com.br",
    "oi.com.br", "aol.com", "protonmail.com", "proton.me", "gmx.com", "gmx.net", "mail.com",
    "email.com", "superig.com.br", "pop.com.br", "click21.com.br", "brturbo.com.br",
    "veloxmail.com.br", "yahoo.es", "hotmail.es", "hotmail.fr", "live.it", "yandex.com",
    "yandex.ru", "mail.ru", "qq.com", "163.com", "zoho.com", "tutanota.com", "bol.com",
    "estadao.com.br", "folha.com.br", "ibest.com.br", "itelefonica.com.br", "sercomtel.com.br",
}

FREE_RE = re.compile(r"^(gmail|googlemail|hotmail|yahoo|ymail|outlook|live|msn|gmx|aol|icloud|yandex|protonmail)\.[a-z.]{2,8}$")


def is_free(dom: str) -> bool:
    return dom in FREE_DOMAINS or bool(FREE_RE.match(dom))


# Correções diretas de domínios digitados errado
DOMAIN_FIX = {
    "gmail.com.br": "gmail.com", "gmail.co": "gmail.com", "gmail.con": "gmail.com",
    "gmail.cm": "gmail.com", "gmail.om": "gmail.com", "gmail.comm": "gmail.com",
    "gmail.c": "gmail.com", "gmail.coml": "gmail.com", "gmail.cim": "gmail.com",
    "gmail.vom": "gmail.com", "gmail.xom": "gmail.com", "gmail.cpm": "gmail.com",
    "gmail.col": "gmail.com", "gmail.coom": "gmail.com", "gmail.ccom": "gmail.com",
    "gmail.br": "gmail.com", "gmail": "gmail.com", "gmailcom": "gmail.com",
    "hotmail.co": "hotmail.com", "hotmail.con": "hotmail.com", "hotmail.cm": "hotmail.com",
    "hotmail.om": "hotmail.com", "hotmail.comm": "hotmail.com", "hotmail.c": "hotmail.com",
    "hotmail.com.b": "hotmail.com.br", "hotmail.com.brr": "hotmail.com.br",
    "hotmail.br": "hotmail.com.br", "hotmail": "hotmail.com", "hotmailcom": "hotmail.com",
    "yahoo.com.b": "yahoo.com.br", "yahoo.co": "yahoo.com", "yahoo.con": "yahoo.com",
    "yahoo.br": "yahoo.com.br", "yahoo.combr": "yahoo.com.br",
    "outlook.co": "outlook.com", "outlook.con": "outlook.com", "outlook.br": "outlook.com.br",
    "uol.com": "uol.com.br", "uol.com.b": "uol.com.br", "uol.br": "uol.com.br", "uol": "uol.com.br",
    "bol.com.b": "bol.com.br", "bol.br": "bol.com.br",
    "terra.com": "terra.com.br", "terra.br": "terra.com.br", "ig.com": "ig.com.br",
    "icloud.co": "icloud.com", "icloud.con": "icloud.com", "icloud.com.br": "icloud.com",
    "live.co": "live.com", "live.con": "live.com", "gamil.com": "gmail.com", "gmial.com": "gmail.com",
    "gmal.com": "gmail.com", "gmil.com": "gmail.com", "hotmal.com": "hotmail.com", "hotmial.com": "hotmail.com",
    "hotamil.com": "hotmail.com", "gotmail.com": "hotmail.com", "40gmail.com": "gmail.com",
    "gmail.fom": "gmail.com", "gmaill.com": "gmail.com", "gmail.cok": "gmail.com", "homail.com": "hotmail.com", "yaho.com.br": "yahoo.com.br",
    "yahho.com.br": "yahoo.com.br", "yaoo.com.br": "yahoo.com.br", "outlok.com": "outlook.com",
}
# raiz do provedor -> domínio canônico (para correção por similaridade)
PROVIDER_ROOTS = {
    "gmail": "gmail.com", "hotmail": None, "outlook": None, "yahoo": None,
    "icloud": "icloud.com",
}
# domínios reais parecidos com provedores (não corrigir)
PROVEDORES_LEGITIMOS = {"hotmart", "gmx", "globo", "gmailx", "hotel", "hotels", "outlet", "yahii"}
# e-mails automáticos / de sistema (não são leads)
SISTEMA_RE = re.compile(r"^((no[-_.]?reply|do[-_.]?not[-_.]?reply|mailer[-_.]?daemon|postmaster|catchall|bounces?|"
                        r"notificacao|notifications?|nfe|nf-e|newsletter|alerts?)(?=[@._\-\d])|register-)", re.I)
TLD_FIX = [
    (r"\.com\.br\.+$", ".com.br"), (r"\.(combr|com\.b|com\.brr|com\.bt|com\.bf|com\.nr|con\.br|cpm\.br|comb\.br|com\.rb)$", ".com.br"),
    (r"\.(con|cmo|comm|coom|ccom|cm|om|vom|xom|cpm|cim|coml|col|comn|fom|com[a-z])$", ".com"),
    (r"\.br\.com$", ".com.br"), (r"\.com\.r$", ".com.br"), (r"\.(org\.b|org\.brr|ogr\.br)$", ".org.br"),
    (r"\.(gov\.b|gov\.brr)$", ".gov.br"),
]
JUNK_LOCALS = {"teste", "test", "naotem", "nao", "naopossui", "sem", "semEmail", "email",
               "x", "xx", "xxx", "asdf", "a", "aa", "aaa", "abc", "exemplo", "example",
               "nenhum", "nt", "n", "000", "0"}
EMAIL_RE = re.compile(r"^[a-z0-9](?:[a-z0-9._%+\-]*[a-z0-9_\-])?@[a-z0-9](?:[a-z0-9\-]*[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9\-]*[a-z0-9])?)*\.[a-z]{2,}$")
SPAM_TLDS = (".top", ".ru", ".xyz", ".icu", ".cn", ".su", ".click", ".buzz", ".monster", ".tw", ".pl", ".ua", ".kz", ".by")

DOMINIO_INTERNO = ("camisetaexpress.com.br", "camisetaexpress.com")
# domínios de envio automático (não são leads)
SISTEMA_DOM_RE = re.compile(r"(docusign|prod\.outlook\.com|email\.amazon|skype\.com|trackmail|onlyappwebmail|"
                            r"mail\.instagram|facebookmail|linkedin\.com|mercadopago|pagseguro\.uol|e\.mail\.)")
# e-mails da equipe usados em testes do site. Padrões extras (1 regex por linha) podem ser colocados em
# dados_brutos/emails_internos.txt — fica fora do git para não expor e-mails pessoais.
_extra = RAW / "emails_internos.txt"
_padroes = [r"testegoogle\d*@gmail\.com"] + (
    [l.strip() for l in _extra.read_text().splitlines() if l.strip() and not l.startswith("#")] if _extra.exists() else [])
EMAILS_INTERNOS_RE = re.compile(r"^(" + "|".join(_padroes) + r")$")

# Segmentos (nome da aba -> regras). A ordem define a prioridade.
SEGMENTOS = [
    ("Governo e Setor Público", {
        "tld": (".gov.br", ".jus.br", ".mp.br", ".leg.br", ".mil.br", ".gov", ".def.br"),
        "kw": ("prefeitura", "camara municipal", "tribunal", "secretaria", "governo", "ministerio",
               "detran", "policia", "defensoria", "sabesp", "metrosp", "cptm", "correios", "poupatempo")}),
    ("Educação e Escolas", {
        "listas": ("escola_de_ingles",), "tld": (".edu.br", ".edu"),
        "kw": ("escola", "colegio", "school", "educa", "ensino", "english", "idioma", "ingles",
               "faculdade", "universidade", "unive", "creche", "pedagog", "fisk", "wizard", "ccaa",
               "cultura inglesa", "kumon", "senac", "senai", "sesi", "etec", "fatec", "usp", "metodista", "academy",
               "ufv", "ufrj", "ufmg", "unifesp", "espm", "fiap", "ibmec", "belasartes", "senna", "lemann",
               "unicamp", "unesp", "fgv", "insper", "mackenzie", "puc", "anhanguera", "estacio",
               "objetivo", "anglo", "positivo", "curso", "uniforme escolar")}),
    ("Agências Mkt e Publicidade", {
        "listas": ("agencias_de_mkt_e_publicidade", "abap"),
        "kw": ("agencia", "propaganda", "publicidade", "publicitari", "marketing", "comunicacao",
               "comunica", "criativ", "branding", "midia", "media", "advertising", "ads", "mkt",
               "producoes", "produtora", "live marketing", "trade", "ativacao", "endomarketing",
               "design", "studio", "estudio", "agency", "ogilvy", "wmccann", "publicis", "dentsu",
               "havas", "=africa", "fcb", "grey", "leo burnett", "wunderman", "vml", "globo", "4allcom", "voeideias",
               "link-worldwide", "l-founders", "reddoor", "llacuna", "valiant", "sherpa42", "ideias", "conteudo",
               "content", "editora", "jornal", "revista", "radio", "tv", "=record", "=band", "sbt", "influenc")}),
    ("Brindes e Promocionais", {
        "listas": ("brasil_promotions",),
        "kw": ("brinde", "promo", "gift", "merchand", "personaliz", "presente", "souvenir",
               "kit", "incentivo", "brind", "creata", "danny")}),
    ("Gráficas e Impressão", {
        "listas": ("impressao_digital",),
        "kw": ("grafica", "graphic", "print", "impress", "silk", "serigraf", "estamp", "sublima",
               "bordado", "copiadora", "plotter", "offset", "etiqueta", "rotulo", "dtf", "graf", "papelaria")}),
    ("Eventos e Buffet", {
        "listas": ("buffet_infantil",),
        "kw": ("buffet", "evento", "event", "festa", "cerimonial", "feira", "=expo", "exposic", "reedexpo", "congress",
               "turismo", "viage", "travel", "hotel", "show", "espetaculo", "teatro", "mci-group", "rxglobal",
               "cenografia", "camarim", "flytour", "zenatur", "=tour", "disney", "cinema", "music", "=festival")}),
    ("Franquias", {
        "listas": ("franquias",),
        "kw": ("franquia", "franchis", "franqueado")}),
    ("Alimentação e Bebidas", {
        "listas": ("empresas_alimenticias",),
        "kw": ("aliment", "food", "restaurante", "cafe", "padaria", "bebida", "cerveja", "chocolate",
               "doce", "pizz", "burger", "frigorif", "nestle", "ambev", "heineken", "coca-cola", "cocacola",
               "pepsico", "bauducco", "mondelez", "kof.com", "sorvete", "gelato", "acai", "confeitaria",
               "hamburg", "churrasc", "cervejaria", "vinho", "wine", "carnes", "laticin", "danone",
               "unilever", "brf", "seara", "jbs", "marilan", "cacau", "sucos", "emporio", "gastronom", "nestle", "mdias",
               "callebaut", "vigor", "tirolez", "panco", "redbull", "garoto", "coracoes", "flavors", "castrolanda",
               "=kof", "sodexo", "ifood", "outback", "mcdonalds", "arcosdorados", "bk", "habibs", "cacaushow",
               "kopenhagen", "lacta", "pepsi", "ambev", "diageo", "pernod", "camil", "josapar", "piracanjuba")}),
    ("Condomínios e Imobiliárias", {
        "listas": ("condominios",),
        "kw": ("condominio", "imobili", "residencial", "edificio", "predial", "administradora",
               "incorporadora", "construtora", "imoveis", "realty", "lello", "loteamento", "=jll", "alphaville", "eztec",
               "cyrela", "mrv", "tecnisa", "=even", "gafisa", "construcao", "engenharia civil")}),
    ("Saúde e Farmacêutico", {
        "kw": ("hospital", "clinica", "saude", "health", "farma", "pharma", "laborat", "odonto",
               "sandoz", "ache.com", "eurofarma", "unimed", "drogaria", "droga", "medic", "einstein",
               "sirio", "fleury", "dasa", "hapvida", "amil", "bayer", "novartis", "pfizer", "sanofi",
               "roche", "abbott", "thermofisher", "biolab", "vet", "pet", "fisioter", "psicolog",
               "nutri", "estetica", "dermat", "oftalmo", "cardio", "santacasa", "santa casa", "hcor", "=jnj",
               "pmenos", "raiadrogasil", "=dpsp", "strattner", "medical", "biotec", "diagnost", "=amil", "=dasa",
               "boehringer", "astrazeneca", "gsk", "=msd", "libbs", "cristalia", "hypera", "=ems", "medley")}),
    ("Financeiro e Seguros", {
        "kw": ("banco", "bank", "safra", "itau", "bradesco", "santander", "caixa", "seguro", "seguradora",
               "invest", "capital", "credito", "financ", "fontecred", "cred", "sicredi",
               "sicoob", "btg", "nubank", "portoseguro", "sulamerica", "previdencia", "cielo", "=stone",
               "pagseguro", "pagbank", "mastercard", "=visa", "=b3", "asset", "corretora", "consorcio",
               "contabil", "contab", "auditoria", "pwc", "deloitte", "kpmg", "=ey", "=bb", "=pine", "cashme",
               "experian", "serasa", "edenred", "alelo", "ticket", "=vr", "elopar", "=xp", "bancoinc", "banrisul",
               "=citi", "redecard", "jpmorgan", "hsbc", "=ubs", "=rede", "getnet", "mercadopago")}),
    ("Esportes e Academias", {
        "kw": ("esporte", "~(?<!tran)sport", "academia", "fitness", "gym", "futebol", "crossfit", "corrida",
               "running", "atletic", "clube", "club", "arena", "bike", "ciclismo", "natacao", "judo",
               "jiu", "karate", "tenis", "volei", "basquete", "smartfit", "bodytech", "decathlon",
               "centauro", "nike", "adidas", "puma", "olympikus", "asics", "palmeiras", "corinthians", "saopaulofc",
               "santosfc", "flamengo", "=cbf", "=fc", "beach", "surf", "=box", "pilates", "yoga", "funcional")}),
    ("ONGs, Associações e Igrejas", {
        "tld": (".org.br", ".org", ".ong.br"),
        "kw": ("associacao", "instituto", "fundacao", "sindicato", "igreja", "ong", "federacao",
               "confederacao", "ministerio ", "paroquia", "comunidade", "projeto social", "sesc",
               "sebrae", "rotary", "lions", "abrinq", "apae", "cooperativa", "=apmp", "=abap", "=aatb", "=ong")}),
    ("Tecnologia e Telecom", {
        "tld": (".io", ".tech", ".dev", ".ai", ".app", ".cloud"),
        "kw": ("tecnologia", "tech", "software", "sistemas", "telecom", "informatica", "dados", "cloud",
               "solution", "digital", "data", "cyber", "web", "app", "sistema", "network", "tim.com",
               "vivo", "claro", "oi.net", "algar", "totvs", "ibm", "microsoft", "google", "oracle",
               "=sap", "amazon", "mercadolivre", "rappi", "uber", "99app", "locaweb", "samsung", "=lge", "=lg",
               "motorola", "=dell", "=hp", "lenovo", "positivo tecnologia", "qintess", "numenit", "netza",
               "ituran", "=me", "=it", "stefanini", "accenture", "capgemini", "ntt", "vtex", "=tivit", "sonda",
               "hub", "startup", "labs", "=ai", "robot", "automation", "automacao", "games", "game")}),
    ("Indústria e Energia", {
        "kw": ("industria", "industrial", "metal", "quimic", "energia", "energy", "petro", "solar",
               "plast", "textil", "siderurg", "mineracao", "motor", "automot", "auto", "pecas",
               "embalage", "papel", "celulose", "cimento", "agro", "fertiliz", "seeds", "sementes",
               "maquina", "equipament", "eletric", "eletro", "whirlpool", "=deca", "duratex", "gerdau", "vale.com",
               "votorantim", "braskem", "weg", "basf", "3m.com", "copagaz", "ultragaz", "gas",
               "alpargatas", "natura", "boticario", "avon", "cosmetic", "tintas", "suvinil", "=kcc", "kimberly",
               "colpal", "colgate", "multibras", "=dow", "sealedair", "saint-gobain", "granadophebo", "modine",
               "=edp", "edpbr", "carbonext", "bignardi", "madeira", "isapa", "=isa", "=enel", "cpfl", "cemig",
               "=light", "comgas", "shell", "raizen", "ipiranga", "=vibra", "ypê", "=ype", "reckitt", "=pg",
               "procter", "johnson", "henkel", "eurofarma", "=mahle", "bosch", "siemens", "=abb", "honeywell",
               "mercedes", "volkswagen", "=vw", "fiat", "stellantis", "=gm", "toyota", "honda", "hyundai",
               "renault", "nissan", "scania", "volvo", "caoa", "pirelli", "michelin")}),
    ("Varejo, Moda e Atacado", {
        "listas": ("atacadistas",),
        "kw": ("atacad", "distribuid", "revend", "comercio", "comercial", "loja", "store", "shop",
               "moda", "confec", "boutique", "magazine", "vestuario", "calcado", "varejo", "carrefour",
               "assai", "atacadao", "makro", "renner", "riachuelo", "cea", "marisa", "havan", "americanas",
               "casasbahia", "viavarejo", "magalu", "uniforme", "roupa", "camiseta", "malha", "textil",
               "fashion", "wear", "brand", "outlet", "supermercado", "mercado", "import", "export", "=gpa", "gpabr",
               "paodeacucar", "roldao", "wal-mart", "walmart", "luxottica", "iguatemi", "shopping", "pernambucanas",
               "gimba", "enjoei", "=davo", "taniabulhoes", "hering", "arezzo", "lojas", "=cea", "zara", "=reserva",
               "=farm", "animale", "=soma", "=hope", "=dafiti", "netshoes", "kalunga", "leroy", "telhanorte",
               "tokstok", "etna", "=mm", "multicoisas", "=atacarejo", "sams", "=tenda", "=spani")}),
    ("Serviços, Logística e RH", {
        "kw": ("advoga", "advocacia", "juridic", "law", "consult", "rh", "recursos humanos", "recrut",
               "talent", "engenharia", "engenhar", "arquitet", "logistic", "transport", "log.com", "cargo",
               "express", "entrega", "limpeza", "facilities", "servicos", "service", "seguranca",
               "manutencao", "terceiriz", "gestao", "assessoria", "treinamento", "coaching", "jsl",
               "dhl", "fedex", "gps", "gpssa", "jadlog", "shipping", "voeazul", "=azul", "latam", "voegol",
               "mayerbrown", "allis", "=serv", "servic", "hagana", "pierserv", "expertise", "=log", "loggi", "sequoia",
               "=tnt", "=ups", "correio", "mudanca", "frete", "rh", "pessoas", "people", "=hr", "adecco",
               "manpower", "randstad", "=gi", "=ajato", "trabalho", "juridico", "tax", "patentes")}),
]
SEG_OUTROS = "Empresas - Outros Segmentos"
SEG_PF = "Pessoa Física - Sem Empresa"
ORDEM_ABAS_SEG = [s for s, _ in SEGMENTOS] + [SEG_OUTROS, SEG_PF]

# Listas Locaweb que representam responsável/origem e não segmento
LISTAS_RESPONSAVEL = {"rose": "Rose", "leads_debora": "Débora", "jose_luiz": "José Luiz",
                      "vendedoras": "Vendedoras", "store": "Store"}

# Marcadores fortes de empresa no nome salvo na agenda/WhatsApp (contatos sem e-mail corporativo)
EMPRESA_NO_NOME_RE = re.compile(
    r"\b(ltda|eireli|epp|escola|colegio|col[eé]gio|buffet|gr[aá]fica|ag[eê]ncia|condom[ií]nio|"
    r"restaurante|igreja|associa[cç][aã]o|instituto|prefeitura|academia|cl[ií]nica|loja|confec[cç][oõ]es|"
    r"uniformes|brindes|eventos|distribuidora|com[eé]rcio|ind[uú]stria|farm[aá]cia|hospital|hotel|"
    r"transportadora|construtora|imobili[aá]ria|consultoria|advocacia|contabilidade|papelaria|supermercado|"
    r"padaria|pizzaria|lanchonete|oficina|studio|est[uú]dio|marketing|comunica[cç][aã]o|franquia)\b", re.I)

EMPRESA_JUNK = re.compile(
    r"^(-+|\.+|x+|n/?a|na|nao|não|nenhum[a]?|n[aã]o tenho|n[aã]o possuo|n[aã]o se aplica|particular|"
    r"pessoa f[ií]sica|pf|cpf|aut[oô]nom[oa]|eu|eu mesmo|eu mesma|minha|fam[ií]lia|pessoal|"
    r"sem empresa|nenhuma empresa|teste|test|casa|residencial|evento|0+|\d+|a|aa|.|"
    r"anivers[aá]rio.*|casamento.*|festa.*|formatura.*|ch[aá] (de|revela).*|uso pessoal|amigos|turma.*|"
    r"despedida.*|viagem.*|fam[ií]lia .*|meu|minha .*|nao informado|n[aã]o informar|---+)$", re.I)

# --------------------------------------------------------------------------------------
# Utilidades de texto
# --------------------------------------------------------------------------------------

def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def clean_str(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).replace(" ", " ").strip()
    return "" if s.lower() in {"nan", "none", "null"} else s


PARTICULAS = {"da", "de", "do", "das", "dos", "e", "di", "du", "van", "von"}


def clean_name(raw: str) -> str:
    s = clean_str(raw)
    if not s or "@" in s:
        return ""
    # "A d r i a n a N e v e s" -> "Adriana Neves"
    toks = s.split()
    if len(toks) >= 4 and all(len(t) == 1 for t in toks):
        s = re.sub(r"(?<!^)(?=[A-ZÁÉÍÓÚÂÊÔÃÕÇ])", " ", "".join(toks))
    s = re.sub(r"[^\w\s'\-\.&]", " ", s, flags=re.UNICODE)  # remove emojis/símbolos
    s = re.sub(r"[\d_]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" .-'")
    if len(re.sub(r"[^A-Za-zÀ-ÿ]", "", s)) < 2:
        return ""
    out = []
    for i, w in enumerate(s.split(" ")):
        lw = w.lower()
        out.append(lw if (i > 0 and lw in PARTICULAS) else "-".join(p[:1].upper() + p[1:].lower() for p in w.split("-")))
    return " ".join(out)[:120]


def levenshtein(a: str, b: str, maxd: int = 2) -> int:
    """Distância de edição (com transposição de letras vizinhas, ex.: gamil/gmail = 1)."""
    if abs(len(a) - len(b)) > maxd:
        return maxd + 1
    d = [[0] * (len(b) + 1) for _ in range(len(a) + 1)]
    for i in range(len(a) + 1):
        d[i][0] = i
    for j in range(len(b) + 1):
        d[0][j] = j
    for i in range(1, len(a) + 1):
        for j in range(1, len(b) + 1):
            d[i][j] = min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] != b[j - 1]))
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                d[i][j] = min(d[i][j], d[i - 2][j - 2] + 1)
        if min(d[i]) > maxd:
            return maxd + 1
    return d[-1][-1]


# --------------------------------------------------------------------------------------
# E-mail
# --------------------------------------------------------------------------------------

def fix_email(raw: str) -> tuple[str, str, str]:
    """Retorna (email_corrigido, observacao, status). status in {ok, corrigido, invalido, vazio}."""
    s0 = clean_str(raw)
    if not s0:
        return "", "", "vazio"
    s = s0.lower()
    s = re.sub(r"^mailto:", "", s)
    s = s.replace('"', "").replace("\\", "").replace("'", "")
    s = s.strip(" <>()[]{};:,.")
    s = re.sub(r"\s+", "", s)
    s = s.replace(",", ".").replace(";", ".").replace("..", ".").replace("@@", "@")
    s = s.replace("@.", "@").replace(".@", "@")
    s = strip_accents(s)
    notas = []
    if "@" not in s:
        # e-mail sem @: "joaogmail.com" -> "joao@gmail.com"
        for dom in sorted(FREE_DOMAINS | {"gmail.com", "hotmail.com"}, key=len, reverse=True):
            if s.endswith(dom) and len(s) > len(dom) + 1:
                s = s[: -len(dom)] + "@" + dom
                notas.append("inserido @")
                break
    if s.count("@") > 1:  # "nome@gmail.com@gmail.com" ou lixo
        parts = s.split("@")
        s = parts[0] + "@" + parts[-1] if len(parts) == 3 and parts[1] == parts[2] else parts[0] + "@" + parts[1]
        notas.append("@ duplicado")
    if "@" not in s:
        return "", f"sem @: {s0}", "invalido"
    local, dom = s.rsplit("@", 1)
    dom0 = dom
    dom = dom.strip(".-")
    if dom in DOMAIN_FIX:
        dom = DOMAIN_FIX[dom]
    else:
        for pat, rep in TLD_FIX:
            nd = re.sub(pat, rep, dom)
            if nd != dom:
                dom = nd
                break
        if "." not in dom:
            if dom in ("gmail", "hotmail", "outlook", "icloud", "yahoo", "live"):
                dom = dom + ".com"
        # provedores digitados errado (gmial, hotmial, yahho, outlok, iclod...)
        root, _, suf = dom.partition(".")
        if dom not in FREE_DOMAINS and len(root) >= 4 and suf in ("com", "com.br") and root not in PROVEDORES_LEGITIMOS:
            for prov, canon in PROVIDER_ROOTS.items():
                if root != prov and root[0] == prov[0] and levenshtein(root, prov, 2) <= (1 if len(prov) <= 5 else 2):
                    if canon:
                        dom = canon
                    else:
                        dom = prov + "." + (suf if suf in ("com", "com.br") else "com")
                    break
        if dom in DOMAIN_FIX:
            dom = DOMAIN_FIX[dom]
    if dom != dom0:
        notas.append(f"domínio {dom0} → {dom}")
    local = local.strip(".")
    email = f"{local}@{dom}"
    if local in JUNK_LOCALS or not EMAIL_RE.match(email):
        return "", f"inválido: {s0}", "invalido"
    if email != re.sub(r"\s+", "", s0.lower()) and not notas:
        notas.append("caracteres inválidos removidos/substituídos (vírgula, aspas, mailto:, espaços)")
    return email, "; ".join(notas), ("corrigido" if notas else "ok")


# --------------------------------------------------------------------------------------
# Telefone
# --------------------------------------------------------------------------------------

def norm_phone(raw: str):
    """Retorna dict {e164, tipo, ddd_assumido, uf} ou None."""
    s = clean_str(raw)
    if not s:
        return None
    if re.match(r"^\d+(\.0)?$", s) and s.endswith(".0"):
        s = s[:-2]
    plus = s.lstrip().startswith("+")
    d = re.sub(r"\D", "", s)
    if len(d) < 8:
        return None
    if d.startswith("00"):
        d, plus = d[2:], True
    if plus and not d.startswith("55"):
        if 8 <= len(d) <= 15:
            return {"e164": "+" + d, "tipo": "Internacional", "ddd_assumido": False, "uf": ""}
        return None
    if d.startswith("55") and len(d) in (12, 13):
        d = d[2:]
    elif d.startswith("0"):
        if len(d) in (13, 14):      # 0 + operadora (2 díg.) + DDD + número
            d = d[3:]
        elif len(d) in (11, 12):    # 0 + DDD + número
            d = d[1:]
    assumido = False
    if len(d) in (8, 9):
        d, assumido = DDD_PADRAO + d, True
    if len(d) == 10 and d[2] in "6789":   # celular antigo sem o 9º dígito
        d = d[:2] + "9" + d[2:]
    if len(d) not in (10, 11):
        return None
    ddd, num = d[:2], d[2:]
    if ddd not in DDD_UF:
        return None
    if len(num) == 9 and num[0] != "9":
        return None
    if len(num) == 8 and num[0] not in "2345":
        return None
    if len(set(num)) <= 2:  # 99999999, 12121212...
        return None
    return {"e164": f"+55{ddd}{num}", "tipo": "Celular" if len(num) == 9 else "Fixo",
            "ddd_assumido": assumido, "uf": DDD_UF[ddd]}


def fmt_phone(e164: str) -> str:
    if not e164.startswith("+55"):
        return e164
    d = e164[3:]
    ddd, num = d[:2], d[2:]
    return f"+55 ({ddd}) {num[:-4]}-{num[-4:]}"


# --------------------------------------------------------------------------------------
# Leitura das fontes
# --------------------------------------------------------------------------------------
records: list[dict] = []
arquivos_lidos: list[dict] = []
_hashes: dict[str, str] = {}


def is_dup_file(path: Path) -> bool:
    h = hashlib.md5(path.read_bytes()).hexdigest()
    if h in _hashes:
        arquivos_lidos.append({"Arquivo": str(path.relative_to(RAW)), "Registros": 0,
                               "Observação": f"Cópia idêntica de {_hashes[h]} — ignorado"})
        return True
    _hashes[h] = str(path.relative_to(RAW))
    return False


def add(fonte, familia, arquivo, *, nome="", email="", phones=(), empresa="", cidade="",
        data=None, data_tipo="", export_date=None, responsavel="", lista="", status_lw="",
        views=0, clicks=0, bounces=0, interesse="", mensagem="", pagina="", outlook=0):
    records.append(dict(
        fonte=fonte, familia=familia, arquivo=arquivo, nome_raw=clean_str(nome), email_raw=clean_str(email),
        phones_raw=[clean_str(p) for p in phones if clean_str(p)], empresa_raw=clean_str(empresa),
        cidade=clean_str(cidade), data=data if data is not None and not pd.isna(data) else None,
        data_tipo=data_tipo, export_date=export_date, responsavel=responsavel, lista=lista,
        status_lw=status_lw, views=views, clicks=clicks, bounces=bounces,
        interesse=clean_str(interesse), mensagem=clean_str(mensagem), pagina=clean_str(pagina),
        outlook=outlook,
    ))


def read_csv_any(path: Path) -> pd.DataFrame:
    # Exportações do Elementor trazem escapes estilo JSON (\u00e1, \/, \r\n, \")
    txt = path.read_text(encoding="utf-8-sig", errors="replace")
    txt = re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), txt)
    txt = txt.replace("\\r\\n", " ").replace("\\n", " ").replace("\\r", " ").replace("\\/", "/")
    return pd.read_csv(io.StringIO(txt), dtype=str, escapechar="\\", keep_default_na=False)


def load_locaweb():
    for f in sorted((RAW / "locaweb").glob("*.csv")):
        if is_dup_file(f):
            continue
        lista = re.sub(r"_todos.*$", "", f.stem).replace("_copia", "")
        df = read_csv_any(f)
        for r in df.to_dict("records"):
            dt = pd.to_datetime(r.get("data de cadastro"), format="%d/%m/%Y %H:%M", errors="coerce")
            add("Locaweb E-mail Mkt", "Locaweb", f.name, nome=r.get("Nome", ""), email=r.get("email", ""),
                cidade=r.get("Cidade", ""), data=dt, data_tipo="Importação Locaweb",
                export_date=pd.Timestamp("2026-09-18"), responsavel=LISTAS_RESPONSAVEL.get(lista, ""),
                lista=lista, status_lw=r.get("status", ""),
                views=int(r.get("visualizações únicas") or 0), clicks=int(r.get("cliques") or 0),
                bounces=int(r.get("erros de recebimento") or 0))
        arquivos_lidos.append({"Arquivo": f"locaweb/{f.name}", "Registros": len(df), "Observação": f"Lista Locaweb '{lista}'"})


def load_elementor():
    for f in sorted((RAW / "site_forms").glob("*.csv")):
        if is_dup_file(f):
            continue
        df = read_csv_any(f)
        df.columns = [c.strip() for c in df.columns]
        form = df["Nome do Formulário (ID)"].iloc[0] if len(df) else f.stem
        for _, r in df.iterrows():
            g = lambda *cols: next((clean_str(r[c]) for c in cols if c in df.columns and clean_str(r[c])), "")
            nome = " ".join(x for x in [g("Nome"), g("Sobrenome", "sobrenome")] if x) or g("name")
            spam_txt = " ".join(clean_str(r[c]) for c in df.columns if c in ("Nome", "name", "Mensagem", "Empresa"))
            add("Site (Elementor)", "Site", f.name, nome=nome, email=g("E-mail", "Email"),
                phones=[g("Telefone"), g("WhatsApp"), g("DDD + WhatsApp")], empresa=g("Empresa"),
                data=pd.to_datetime(g("Criado em"), errors="coerce"), data_tipo="Envio de formulário",
                export_date=pd.Timestamp("2026-09-18"),
                interesse=" | ".join(x for x in [g("Assunto"), g("Interesse", "Interesse em:"), g("unidades"), g("Quantos Kits")] if x),
                mensagem=g("Mensagem"), pagina=g("Página") or form)
            records[-1]["spam_txt"] = spam_txt
            records[-1]["form"] = form
        arquivos_lidos.append({"Arquivo": f"site_forms/{f.name}", "Registros": len(df), "Observação": f"Formulário '{form}'"})


def load_raiz():
    p = RAW / "raiz"
    # Contatos do celular/WhatsApp da Camiseta Express (Google Contacts)
    f = p / "ContatosCamex24-07.csv"
    if f.exists() and not is_dup_file(f):
        df = read_csv_any(f)
        for r in df.itertuples(index=False):
            add("Contatos WhatsApp Camex", "Agenda", f.name, nome=r[0], phones=[r[1]], email=r[2],
                export_date=pd.Timestamp("2026-07-24"), data_tipo="Exportação")
        arquivos_lidos.append({"Arquivo": "raiz/" + f.name, "Registros": len(df), "Observação": "Agenda/WhatsApp Camex (exportado 24/07/2026)"})
    # Contatos Outlook/celular Débora
    f = p / "ContatosDebora21-09.xlsx"
    if f.exists() and not is_dup_file(f):
        df = pd.read_excel(f, dtype=str)
        for _, r in df.iterrows():
            nome = clean_str(r.get("Display Name")) or " ".join(filter(None, [clean_str(r.get("First Name")), clean_str(r.get("Last Name"))]))
            add("Contatos Débora", "Agenda", f.name, nome=nome,
                phones=[r.get("Mobile Phone"), r.get("Business Phone"), r.get("Home Phone")],
                export_date=pd.Timestamp("2026-09-21"), data_tipo="Exportação", responsavel="Débora")
        arquivos_lidos.append({"Arquivo": "raiz/" + f.name, "Registros": len(df), "Observação": "Agenda Débora (exportado 21/09/2026)"})
    f = p / "ContatosSilmara.xlsx"
    if f.exists() and not is_dup_file(f):
        df = pd.read_excel(f, dtype=str)
        for _, r in df.iterrows():
            nome = clean_str(r.get("Display Name")) or " ".join(filter(None, [clean_str(r.get("First Name")), clean_str(r.get("Last Name"))]))
            add("Contatos Silmara", "Agenda", f.name, nome=nome, phones=[r.get("Mobile Phone")],
                export_date=pd.Timestamp("2026-09-22"), data_tipo="Exportação", responsavel="Silmara")
        arquivos_lidos.append({"Arquivo": "raiz/" + f.name, "Registros": len(df), "Observação": "Agenda Silmara (exportado 22/09/2026)"})
    # E-mails trocados no Outlook (Débora e Rose)
    for fname, resp, dt in [("EmailsDebora22-09.xlsx", "Débora", "2026-09-22"), ("EmailsRose23-09.xlsx", "Rose", "2026-09-23")]:
        f = p / fname
        if not f.exists() or is_dup_file(f):
            continue
        df = pd.read_excel(f, header=None, dtype=str)
        cnt = Counter()
        for v in df[0].dropna():
            for e in re.split(r"[;\s]+", str(v)):
                if "@" in e:
                    cnt[e.strip()] += 1
        for e, n in cnt.items():
            add(f"E-mails Outlook {resp}", "Outlook", fname, email=e, export_date=pd.Timestamp(dt),
                data_tipo="Exportação", responsavel=resp, outlook=n)
        arquivos_lidos.append({"Arquivo": "raiz/" + fname, "Registros": len(df),
                               "Observação": f"Destinatários Outlook {resp} — {sum(cnt.values())} endereços, {len(cnt)} únicos"})
    f = p / "Leads_test_chat.xlsx"
    if f.exists() and not is_dup_file(f):
        df = pd.read_excel(f, dtype=str)
        for _, r in df.iterrows():
            add("Leads Chat do Site", "Chat", f.name, nome=r.get("Nome"), email=r.get("Email"),
                phones=[r.get("Telefone")], cidade=r.get("Cidade"), export_date=pd.Timestamp("2026-09-21"),
                data_tipo="Exportação")
        arquivos_lidos.append({"Arquivo": "raiz/" + f.name, "Registros": len(df), "Observação": "Leads do chat (exportado 21/09/2026)"})
    f = p / "Lista_de_emails_submissoes.xlsx"
    if f.exists() and not is_dup_file(f):
        df = pd.read_excel(f, dtype=str)
        for _, r in df.iterrows():
            add("Site (Lista submissões)", "Site", f.name, nome=r.get("Nome"), email=r.get("Email"),
                export_date=pd.Timestamp("2026-09-18"), data_tipo="Exportação")
        arquivos_lidos.append({"Arquivo": "raiz/" + f.name, "Registros": len(df), "Observação": "Lista de e-mails de submissões do site"})
    f = p / "Atualizacao_Cadastro_B2B.xlsx"
    if f.exists() and not is_dup_file(f):
        df = pd.read_excel(f, dtype=str)
        for _, r in df.iterrows():
            add("Atualização Cadastro B2B", "Site", f.name, nome=r.get("Nome"), email=r.get("E-mail"),
                phones=[r.get("Telefone")], empresa=r.get("Empresa"),
                data=pd.to_datetime(r.get("Data/Hora"), errors="coerce"), data_tipo="Envio de formulário",
                export_date=pd.Timestamp("2026-08-24"), interesse="Atualização de cadastro B2B")
        arquivos_lidos.append({"Arquivo": "raiz/" + f.name, "Registros": len(df), "Observação": "Formulário Atualização B2B"})


# --------------------------------------------------------------------------------------
# Pipeline
# --------------------------------------------------------------------------------------

def is_spam(r) -> bool:
    txt = r.get("spam_txt", "") + " " + r["nome_raw"]
    if re.search(r"[Ѐ-ӿ一-鿿぀-ヿ฀-๿]", txt):
        return True
    em = r["email"] or r["email_raw"].lower()
    if em.endswith(SPAM_TLDS):
        return True
    msg = r["mensagem"].lower()
    if r["familia"] == "Site" and re.search(r"https?://", msg) and not re.search(r"camisetaexpress|drive\.google|wa\.me|instagram", msg) \
            and re.fullmatch(r"[A-Z][a-z]+[A-Z][a-z]+", r["nome_raw"] or ""):
        return True
    if re.search(r"\b(casino|viagra|crypto|bitcoin|seo service|backlink|porn|loan|etsy|procreate|cliparts?)\b", msg):
        return True
    # mensagem em inglês + nome de robô (1 palavra) — site é 100% em português
    palavras = re.findall(r"[a-zà-ú]+", msg)
    en = sum(w in EN_WORDS for w in palavras)
    pt = sum(w in PT_WORDS for w in palavras)
    if len(palavras) >= 5 and en >= 2 and en > 2 * pt and len((r["nome_raw"] or "x").split()) == 1:
        return True
    return False


EN_WORDS = {"the", "and", "for", "with", "your", "you", "custom", "our", "we", "this", "that", "are", "is",
            "free", "best", "more", "much", "hello", "hi", "offer", "website", "business", "price", "please"}
PT_WORDS = {"de", "para", "com", "camiseta", "camisetas", "orcamento", "orçamento", "gostaria", "bom", "dia",
            "boa", "tarde", "ola", "olá", "quero", "uniforme", "uniformes", "pecas", "peças", "preciso", "obrigado",
            "obrigada", "empresa", "que", "um", "uma", "os", "as", "no", "na", "do", "da", "em", "por"}


def main():
    if not RAW.exists():
        sys.exit(f"Pasta de dados não encontrada: {RAW}")
    load_locaweb()
    load_elementor()
    load_raiz()
    print(f"Registros brutos: {len(records)}")

    correcoes = []
    for i, r in enumerate(records):
        r["id"] = i
        email, nota, st = fix_email(r["email_raw"])
        r["email"], r["email_status"] = email, st
        if st in ("corrigido", "invalido"):
            correcoes.append({"E-mail original": r["email_raw"], "E-mail corrigido": email or "(descartado)",
                              "Correção aplicada": nota, "Fonte": r["fonte"], "Arquivo": r["arquivo"]})
        phones = []
        for p in r["phones_raw"]:
            n = norm_phone(p)
            if n and n["e164"] not in [x["e164"] for x in phones]:
                phones.append(n)
        r["phones"] = phones
        r["nome"] = clean_name(r["nome_raw"])
        emp = r["empresa_raw"]
        r["empresa"] = "" if (not emp or EMPRESA_JUNK.match(strip_accents(emp).strip()) or emp.lower() == r["nome_raw"].lower()
                              or not re.search(r"[A-Za-zÀ-ÿ]{2}", emp)) else emp.strip()[:120]

    # --- correção de domínios corporativos digitados errado (por frequência na base) -----
    dom_count = Counter(r["email"].split("@")[1] for r in records if r["email"])
    # só corrige "falta de caractere": domínio raro = domínio frequente sem 1 letra (ex.: erofarma -> eurofarma)
    frequentes = [d for d, c in dom_count.items() if c >= 5 and not is_free(d) and len(d.split(".")[0]) >= 7]
    dom_map = {}
    for d, c in dom_count.items():
        if c > 2 or is_free(d) or d in frequentes:
            continue
        for fd in frequentes:
            if (len(fd) == len(d) + 1 and fd.split(".", 1)[-1] == d.split(".", 1)[-1]
                    and levenshtein(d, fd, 1) == 1):
                dom_map[d] = fd
                break
    for r in records:
        if r["email"] and r["email"].split("@")[1] in dom_map:
            old = r["email"]
            r["email"] = old.split("@")[0] + "@" + dom_map[old.split("@")[1]]
            correcoes.append({"E-mail original": r["email_raw"], "E-mail corrigido": r["email"],
                              "Correção aplicada": f"domínio corporativo {old.split('@')[1]} → {dom_map[old.split('@')[1]]} (similar a domínio frequente na base)",
                              "Fonte": r["fonte"], "Arquivo": r["arquivo"]})

    # --- filtros: interno, teste, spam, sem contato -------------------------------------
    removidos = []
    fone_interno = Counter()
    for r in records:
        if r["email"].endswith(DOMINIO_INTERNO) or EMAILS_INTERNOS_RE.match(r["email"]):
            for p in r["phones"]:
                fone_interno[p["e164"]] += 1
    fones_internos = {p for p, c in fone_interno.items() if c >= 1}
    # e-mail usado em spam no site -> todos os registros dele (inclusive Locaweb) são spam
    emails_spam = {r["email"] for r in records if r["familia"] == "Site" and r["email"] and is_spam(r)}
    validos = []
    for r in records:
        motivo = ""
        if r["email"].endswith(DOMINIO_INTERNO) or EMAILS_INTERNOS_RE.match(r["email"]):
            motivo = "E-mail interno Camiseta Express / equipe"
        elif r["email"] and (SISTEMA_RE.match(r["email"]) or SISTEMA_DOM_RE.search(r["email"].split("@")[1])) and not r["phones"]:
            motivo = "E-mail automático de sistema (no-reply etc.)"
        elif re.search(r"\b(teste|test)\b", (r["nome_raw"] + " " + r["empresa_raw"]).lower()) and r["familia"] in ("Site", "Chat"):
            motivo = "Registro de teste"
        elif (r["familia"] == "Site" and is_spam(r)) or r["email"] in emails_spam:
            motivo = "Spam (formulário do site)"
        elif re.match(r"^(teste|test)[\d._-]*@|teste\d*@|@teste\.", r["email"]):
            motivo = "Registro de teste"
        elif not r["email"] and not r["phones"]:
            motivo = "Sem e-mail válido e sem telefone válido"
        if not motivo and r["phones"]:
            r["phones"] = [p for p in r["phones"] if p["e164"] not in fones_internos]
            if not r["email"] and not r["phones"]:
                motivo = "Telefone da própria Camiseta Express"
        if motivo:
            removidos.append({"Motivo": motivo, "Nome": r["nome_raw"], "E-mail": r["email_raw"],
                              "Telefone": ", ".join(r["phones_raw"]), "Fonte": r["fonte"], "Arquivo": r["arquivo"]})
        else:
            validos.append(r)
    print(f"Válidos: {len(validos)}  Removidos: {len(removidos)}")

    # --- deduplicação (union-find por e-mail e celular) ---------------------------------
    parent = list(range(len(records)))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    by_email, by_phone = defaultdict(list), defaultdict(list)
    for r in validos:
        if r["email"]:
            by_email[r["email"]].append(r["id"])
        for p in r["phones"]:
            by_phone[p["e164"]].append(r["id"])
    for ids in by_email.values():
        for x in ids[1:]:
            union(ids[0], x)
    # telefone compartilhado por muitos e-mails diferentes = PABX/telefone da empresa -> não unificar
    for ph, ids in by_phone.items():
        emails = {records[i]["email"] for i in ids if records[i]["email"]}
        if len(emails) > 2:
            continue
        # telefone fixo só une se no máximo 1 e-mail distinto
        if not ph.startswith("+55") or len(ph) == 13:
            if len(emails) > 1:
                continue
        for x in ids[1:]:
            union(ids[0], x)
    grupos = defaultdict(list)
    for r in validos:
        grupos[find(r["id"])].append(r)
    print(f"Contatos únicos após deduplicação: {len(grupos)}")

    # --- consolidação --------------------------------------------------------------------
    FONTE_PRIOR_NOME = {"Site": 0, "Chat": 1, "Locaweb": 2, "Agenda": 3, "Outlook": 4}
    contatos = []
    for rs in grupos.values():
        emails = Counter(r["email"] for r in rs if r["email"])
        # prioriza e-mail corporativo e o mais frequente
        email_list = sorted(emails, key=lambda e: (is_free(e.split("@")[1]), -emails[e]))
        phones = {}
        for r in rs:
            for p in r["phones"]:
                phones.setdefault(p["e164"], p)
        ph_sorted = sorted(phones.values(), key=lambda p: (p["tipo"] != "Celular", p["ddd_assumido"]))
        celulares = [p for p in ph_sorted if p["tipo"] == "Celular"]
        fixos = [p for p in ph_sorted if p["tipo"] != "Celular"]
        nomes = [(FONTE_PRIOR_NOME.get(r["familia"], 9), -len(r["nome"].split()), r["nome"]) for r in rs if r["nome"]]
        nome = sorted(nomes)[0][2] if nomes else ""
        empresas = Counter(r["empresa"] for r in rs if r["empresa"])
        email_principal = email_list[0] if email_list else ""
        dominio = email_principal.split("@")[1] if email_principal else ""
        corp_dom = next((e.split("@")[1] for e in email_list if not is_free(e.split("@")[1])), "")
        if empresas:
            empresa, emp_origem = empresas.most_common(1)[0][0], "Informada pelo lead"
        elif corp_dom:
            raiz = corp_dom.split(".")[0]
            empresa, emp_origem = raiz.replace("-", " ").title(), "Derivada do domínio do e-mail"
        elif nome and EMPRESA_NO_NOME_RE.search(nome) and not email_list:
            empresa, emp_origem = nome, "Identificada no nome salvo na agenda/WhatsApp"
        else:
            empresa, emp_origem = "", ""
        # datas reais de interação = envios de formulário. A "data de cadastro" da Locaweb é carimbo de
        # importação em lote (out/2024 e 30/07/2026): serve só como data de entrada, não como interação.
        datas_reais = [r["data"] for r in rs if r["data"] is not None and r["data_tipo"] == "Envio de formulário"]
        datas_import = [r["data"] for r in rs if r["data"] is not None and r["data_tipo"] == "Importação Locaweb"]
        exports = [r["export_date"] for r in rs if r["export_date"] is not None]
        ultima = max(datas_reais) if datas_reais else None
        candidatas = [(d, "Envio de formulário") for d in datas_reais] + [(d, "Importação Locaweb") for d in datas_import]
        if candidatas:
            data_entrada, data_entrada_tipo = min(candidatas)
            primeira = data_entrada
        else:
            primeira = None
            data_entrada = min(exports) if exports else None
            data_entrada_tipo = "Data da exportação (estimada)"
        fam = {r["familia"] for r in rs}
        fontes = sorted({r["fonte"] for r in rs})
        listas = sorted({r["lista"] for r in rs if r["lista"]})
        forms = sorted({r.get("form", "") for r in rs if r.get("form")})
        resp = sorted({r["responsavel"] for r in rs if r["responsavel"]})
        status_lw = {r["status_lw"] for r in rs if r["status_lw"]}
        # status só vale para o e-mail principal
        st_principal = {r["status_lw"] for r in rs if r["status_lw"] and r["email"] == email_principal}
        optout = "Descadastrado" in st_principal
        bounce = ("Desativado" in st_principal) and not optout
        clicks = sum(r["clicks"] for r in rs)
        views = max([r["views"] for r in rs] + [0])
        outlook = sum(r["outlook"] for r in rs)
        site_datas = [r["data"] for r in rs if r["familia"] == "Site" and r["data"] is not None]
        interesses = [r["interesse"] for r in sorted(rs, key=lambda r: r["data"] or pd.Timestamp(0), reverse=True) if r["interesse"]]
        mensagens = [r["mensagem"] for r in sorted(rs, key=lambda r: r["data"] or pd.Timestamp(0), reverse=True) if r["mensagem"]]
        cidades = [r["cidade"] for r in rs if r["cidade"]]
        uf = next((p["uf"] for p in ph_sorted if p["uf"]), "")
        c = dict(
            nome=nome, email=email_principal, emails_extra=", ".join(email_list[1:]),
            celular=celulares[0]["e164"] if celulares else "",
            telefone=fixos[0]["e164"] if fixos else (celulares[1]["e164"] if len(celulares) > 1 else ""),
            telefones_extra=", ".join(p["e164"] for p in (celulares[1:] + fixos[1:]) if p["e164"] not in ((celulares[1]["e164"],) if (not fixos and len(celulares) > 1) else ())),
            ddd_assumido=any(p["ddd_assumido"] for p in ph_sorted[:2]),
            empresa=empresa, emp_origem=emp_origem, dominio=corp_dom,
            cidade=cidades[0].title() if cidades else "", uf=uf,
            data_entrada=data_entrada, data_ultima=ultima, data_estimada=primeira is None,
            data_entrada_tipo=data_entrada_tipo,
            fam=fam, fontes=fontes, listas=listas, forms=forms, resp=resp, status_lw=status_lw,
            optout=optout, bounce=bounce, clicks=clicks, views=views, outlook=outlook,
            site_ultimo=max(site_datas) if site_datas else None,
            interesse=interesses[0][:200] if interesses else "",
            mensagem=re.sub(r"\s+", " ", mensagens[0])[:300] if mensagens else "",
            n_registros=len(rs),
            email_corrigido=any(r["email_status"] == "corrigido" and r["email"] == email_principal for r in rs),
        )
        c["segmento"] = classificar(c)
        c["score"], c["temperatura"] = temperatura(c)
        c["periodo"] = periodo(c)
        contatos.append(c)

    gerar_saida(contatos, removidos, correcoes)


def _kw_hit(text: str, kws) -> str:
    """Palavra-chave com até 3 letras ou prefixada com '=' só vale como palavra inteira; '~' = regex."""
    for k in kws:
        if k.startswith("~"):
            if re.search(k[1:], text):
                return k
            continue
        if k.startswith("="):
            k = k[1:]
            if re.search(rf"(^|[\s\.\-_]){re.escape(k)}($|[\s\.\-_])", text):
                return k
        elif len(k) <= 3:
            if re.search(rf"(^|[\s\.\-_]){re.escape(k)}($|[\s\.\-_])", text):
                return k
        elif k in text:
            return k
    return ""


def classificar(c) -> str:
    is_b2b = bool(c["dominio"]) or bool(c["empresa"])
    if not is_b2b:
        return SEG_PF
    dom = c["dominio"]
    raiz_dom = re.sub(r"(\.(com|net|org|gov|edu|ind|coop|art|adv|eng|med|global|group|red|ag|[a-z]{2}))+$", "", dom)
    texto = strip_accents(" ".join([raiz_dom, c["empresa"]]).lower())
    forms_txt = strip_accents(" ".join(c["forms"]).lower())
    # 1) TLD institucional
    for seg, reg in SEGMENTOS:
        if dom and any(dom.endswith(t) for t in reg.get("tld", ())) and seg in ("Governo e Setor Público", "Educação e Escolas"):
            return seg
    # 2) Palavras-chave no domínio / nome da empresa
    for seg, reg in SEGMENTOS:
        if _kw_hit(texto, reg.get("kw", ())):
            return seg
    # 3) Lista de segmento da Locaweb (curadoria do cliente)
    for seg, reg in SEGMENTOS:
        if any(l in reg.get("listas", ()) for l in c["listas"]):
            return seg
    # 4) Formulário de uniforme escolar (atacado) -> escolas
    if "uniforme escolar" in forms_txt:
        return "Educação e Escolas"
    # 5) Demais TLDs (.org, .io...)
    for seg, reg in SEGMENTOS:
        if dom and any(dom.endswith(t) for t in reg.get("tld", ())):
            return seg
    return SEG_OUTROS


def _dias(d):
    return (HOJE - d).days if d is not None else None


def temperatura(c):
    """Score 0-100: recência da data real de interação + intenção da origem + engajamento."""
    s = 0
    dias = _dias(c["data_ultima"])
    if dias is not None:
        s += 45 if dias <= 90 else 35 if dias <= 180 else 25 if dias <= 365 else 12 if dias <= 730 else 5
    fam = c["fam"]
    if "Site" in fam or "Chat" in fam:
        s += 20                                  # lead inbound (pediu orçamento/contato)
    if "Outlook" in fam:
        s += 15 + (10 if c["outlook"] >= 3 else 0)   # troca de e-mails com vendedora
    if "Agenda" in fam:
        s += 12                                  # contato salvo por vendedora/WhatsApp
    if c["clicks"] > 0:
        s += 15
    elif c["views"] > 0:
        s += 5
    s += min(10, 5 * (len(fam) - 1))             # aparece em várias fontes
    if c["segmento"] != SEG_PF:
        s += 5                                   # perfil B2B (foco do cliente)
    if c["bounce"]:
        s -= 10
    s = max(0, min(100, s))
    if c["optout"]:
        return s, "Descadastrado"
    return s, ("Quente" if s >= 55 else "Morno" if s >= 30 else "Frio")


def periodo(c):
    d = c["data_entrada"]
    if d is None:
        return "Sem data"
    if c["data_estimada"]:
        return "Sem data (base exportada em 2026)"
    dias = _dias(d)
    return ("0-3 meses" if dias <= 90 else "3-6 meses" if dias <= 180 else "6-12 meses" if dias <= 365
            else "1-2 anos" if dias <= 730 else "Mais de 2 anos")


# --------------------------------------------------------------------------------------
# Saída
# --------------------------------------------------------------------------------------
COLS = [
    ("Nome", "nome"), ("E-mail", "email"), ("Celular/WhatsApp", "celular"), ("Telefone", "telefone"),
    ("Empresa", "empresa"), ("Origem do nome da empresa", "emp_origem"), ("Site/Domínio da empresa", "dominio"),
    ("Segmento", "segmento"), ("Temperatura", "temperatura"), ("Score (0-100)", "score"),
    ("Período de entrada", "periodo"), ("Data de entrada na base", "data_entrada_fmt"),
    ("Origem da data de entrada", "data_entrada_tipo"), ("Última interação (data real)", "data_ultima_fmt"),
    ("Cidade", "cidade"), ("UF (pelo DDD)", "uf"), ("Origens", "fontes_fmt"), ("Listas Locaweb", "listas_fmt"),
    ("Formulários do site", "forms_fmt"), ("Responsável/Vendedora", "resp_fmt"), ("Interesse (último)", "interesse"),
    ("Mensagem (última)", "mensagem"), ("Status e-mail", "status_email"), ("Cliques e-mail mkt", "clicks"),
    ("Visualizações e-mail mkt", "views"), ("E-mails trocados (Outlook)", "outlook"),
    ("E-mails adicionais", "emails_extra"), ("Telefones adicionais", "telefones_extra"),
    ("DDD 11 assumido?", "ddd_fmt"), ("E-mail corrigido?", "email_corr_fmt"),
    ("Registros unificados", "n_registros"), ("Tags RD Station", "tags"),
]


def slug(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", strip_accents(s).lower()).strip("-")


def to_frame(contatos) -> pd.DataFrame:
    rows = []
    for c in contatos:
        c = dict(c)
        c["data_entrada_fmt"] = c["data_entrada"].strftime("%d/%m/%Y") if c["data_entrada"] is not None else ""
        c["data_ultima_fmt"] = c["data_ultima"].strftime("%d/%m/%Y") if c["data_ultima"] is not None else ""
        c["fontes_fmt"] = "; ".join(c["fontes"])
        c["listas_fmt"] = "; ".join(c["listas"])
        c["forms_fmt"] = "; ".join(c["forms"])
        c["resp_fmt"] = "; ".join(c["resp"])
        c["status_email"] = ("Descadastrado (opt-out)" if c["optout"] else "Bounce/Desativado na Locaweb" if c["bounce"]
                             else "Sem e-mail" if not c["email"] else "Válido")
        c["ddd_fmt"] = "Sim" if c["ddd_assumido"] else "Não"
        c["email_corr_fmt"] = "Sim" if c["email_corrigido"] else "Não"
        tags = [slug(c["segmento"]), "lead-" + slug(c["temperatura"]), "entrada-" + slug(c["periodo"])]
        tags += ["origem-" + slug(f) for f in sorted(c["fam"])]
        tags += ["resp-" + slug(r) for r in c["resp"]]
        c["tags"] = ", ".join(dict.fromkeys(tags))
        c["celular"] = fmt_phone(c["celular"]) if c["celular"] else ""
        c["telefone"] = fmt_phone(c["telefone"]) if c["telefone"] else ""
        c["telefones_extra"] = ", ".join(fmt_phone(p) for p in c["telefones_extra"].split(", ") if p)
        rows.append({h: c[k] for h, k in COLS} | {"_data_sort": c["data_ultima"] or c["data_entrada"] or pd.Timestamp(0)})
    df = pd.DataFrame(rows)
    ordem_temp = {"Quente": 0, "Morno": 1, "Frio": 2, "Descadastrado": 3}
    df["_t"] = df["Temperatura"].map(ordem_temp)
    df = df.sort_values(["_t", "Score (0-100)", "_data_sort"], ascending=[True, False, False]).drop(columns=["_t", "_data_sort"])
    return df.reset_index(drop=True)


def gerar_saida(contatos, removidos, correcoes):
    OUT.mkdir(exist_ok=True)
    df = to_frame(contatos)
    ativos = df[df["Temperatura"] != "Descadastrado"]
    descad = df[df["Temperatura"] == "Descadastrado"]

    # Empresas (1 linha por empresa)
    emp = ativos[ativos["Segmento"] != SEG_PF].copy()
    emp["_key"] = emp["Site/Domínio da empresa"].where(emp["Site/Domínio da empresa"] != "", emp["Empresa"].str.lower().str.strip())
    temp_rank = {"Quente": 3, "Morno": 2, "Frio": 1}
    empresas = (emp.groupby("_key").agg(
        Empresa=("Empresa", lambda s: s.value_counts().index[0]),
        Dominio=("Site/Domínio da empresa", "first"),
        Segmento=("Segmento", lambda s: s.value_counts().index[0]),
        Contatos=("Nome", "size"),
        Quentes=("Temperatura", lambda s: (s == "Quente").sum()),
        Mornos=("Temperatura", lambda s: (s == "Morno").sum()),
        Frios=("Temperatura", lambda s: (s == "Frio").sum()),
        Score_max=("Score (0-100)", "max"),
        UF=("UF (pelo DDD)", lambda s: next((x for x in s if x), "")),
        Responsaveis=("Responsável/Vendedora", lambda s: "; ".join(sorted({y for x in s for y in x.split("; ") if y}))),
        Contatos_lista=("E-mail", lambda s: ", ".join([x for x in s if x][:10])),
    ).reset_index(drop=True))
    empresas["Temperatura da conta"] = empresas.apply(
        lambda r: "Quente" if r.Quentes else "Morno" if r.Mornos else "Frio", axis=1)
    empresas = empresas.rename(columns={"Dominio": "Site/Domínio", "Score_max": "Maior score",
                                        "Responsaveis": "Responsáveis", "Contatos_lista": "E-mails dos contatos (até 10)"})
    empresas = empresas[["Empresa", "Site/Domínio", "Segmento", "Temperatura da conta", "Maior score", "Contatos",
                         "Quentes", "Mornos", "Frios", "UF", "Responsáveis", "E-mails dos contatos (até 10)"]]
    empresas = empresas.sort_values(["Maior score", "Contatos"], ascending=False)

    # Resumo
    resumo_seg = pd.crosstab(ativos["Segmento"], ativos["Temperatura"]).reindex(index=ORDEM_ABAS_SEG, columns=["Quente", "Morno", "Frio"], fill_value=0)
    resumo_seg["Total"] = resumo_seg.sum(axis=1)
    resumo_seg["Com e-mail"] = ativos.groupby("Segmento")["E-mail"].apply(lambda s: (s != "").sum()).reindex(resumo_seg.index).fillna(0).astype(int)
    resumo_seg["Com celular"] = ativos.groupby("Segmento")["Celular/WhatsApp"].apply(lambda s: (s != "").sum()).reindex(resumo_seg.index).fillna(0).astype(int)
    resumo_seg.loc["TOTAL"] = resumo_seg.sum()
    resumo_per = pd.crosstab(ativos["Período de entrada"], ativos["Temperatura"]).reindex(
        index=["0-3 meses", "3-6 meses", "6-12 meses", "1-2 anos", "Mais de 2 anos", "Sem data (base exportada em 2026)"],
        columns=["Quente", "Morno", "Frio"], fill_value=0)
    resumo_per["Total"] = resumo_per.sum(axis=1)
    arquivos = pd.DataFrame(arquivos_lidos)
    total_brutos = len(records)

    kpis = [
        ("Arquivos processados (únicos)", int((arquivos["Registros"] > 0).sum())),
        ("Arquivos ignorados (cópias idênticas)", int((arquivos["Registros"] == 0).sum())),
        ("Registros brutos lidos (todas as fontes)", total_brutos),
        ("Registros removidos (interno/teste/spam/sem contato)", len(removidos)),
        ("Contatos únicos após deduplicação", len(df)),
        ("   • Ativos (importáveis)", len(ativos)),
        ("   • Descadastrados (opt-out — NÃO importar p/ e-mail mkt)", len(descad)),
        ("Contatos com e-mail válido (ativos)", int((ativos["E-mail"] != "").sum())),
        ("Contatos só com telefone (ativos)", int((ativos["E-mail"] == "").sum())),
        ("E-mails com bounce na Locaweb (ativos)", int((ativos["Status e-mail"] == "Bounce/Desativado na Locaweb").sum())),
        ("E-mails corrigidos", sum(1 for c in correcoes if c["E-mail corrigido"] != "(descartado)")),
        ("E-mails inválidos descartados", sum(1 for c in correcoes if c["E-mail corrigido"] == "(descartado)")),
        ("Contatos com DDD 11 assumido", int((ativos["DDD 11 assumido?"] == "Sim").sum())),
        ("Empresas (contas) identificadas", len(empresas)),
        ("Leads Quentes", int((ativos["Temperatura"] == "Quente").sum())),
        ("Leads Mornos", int((ativos["Temperatura"] == "Morno").sum())),
        ("Leads Frios", int((ativos["Temperatura"] == "Frio").sum())),
    ]

    path = OUT / "Base_Clusterizada_Camiseta_Express.xlsx"
    with pd.ExcelWriter(path, engine="xlsxwriter") as xw:
        wb = xw.book
        h = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F3864", "border": 1, "text_wrap": True, "valign": "vcenter"})
        title = wb.add_format({"bold": True, "font_size": 16, "font_color": "#1F3864"})
        sub = wb.add_format({"bold": True, "font_size": 12, "font_color": "#1F3864"})
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        fq = wb.add_format({"bg_color": "#F8CBAD", "font_color": "#9C0006"})
        fm = wb.add_format({"bg_color": "#FFE699", "font_color": "#7F6000"})
        ff = wb.add_format({"bg_color": "#BDD7EE", "font_color": "#1F4E79"})
        fd = wb.add_format({"bg_color": "#D9D9D9", "font_color": "#595959"})
        num = wb.add_format({"num_format": "#,##0"})

        def write_table(df_, name, widths=None):
            df_.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            for j, col in enumerate(df_.columns):
                ws.write(0, j, col, h)
                w = (widths or {}).get(col) or min(45, max(10, int(df_[col].astype(str).str.len().quantile(0.9) if len(df_) else 10) + 2, len(col) + 2))
                ws.set_column(j, j, w)
            ws.freeze_panes(1, 1)
            ws.autofilter(0, 0, max(len(df_), 1), len(df_.columns) - 1)
            if "Temperatura" in df_.columns and len(df_):
                j = list(df_.columns).index("Temperatura")
                for txt, fmt in (("Quente", fq), ("Morno", fm), ("Frio", ff), ("Descadastrado", fd)):
                    ws.conditional_format(1, j, len(df_), j, {"type": "cell", "criteria": "==", "value": f'"{txt}"', "format": fmt})
            return ws

        # 1. LEIA-ME
        ws = wb.add_worksheet("LEIA-ME")
        ws.set_column(0, 0, 58); ws.set_column(1, 1, 110)
        ws.write(0, 0, "Base de Leads Clusterizada — Camiseta Express", title)
        ws.write(1, 0, f"Processado em {datetime.now():%d/%m/%Y %H:%M} • data de referência {HOJE:%d/%m/%Y}")
        linhas = [
            ("COMO USAR", ""),
            ("Base Completa", "Todos os contatos únicos ATIVOS (sem descadastrados), 1 linha por pessoa, já deduplicados."),
            ("Abas de Segmento", "Mesma base separada por segmento de empresa (1 aba por segmento). 'Pessoa Física - Sem Empresa' = e-mail pessoal sem empresa informada."),
            ("Leads Quentes / Mornos / Frios", "Mesma base separada por temperatura, para priorizar a abordagem comercial."),
            ("Empresas (Contas)", "1 linha por empresa (domínio corporativo ou nome informado), com contagem de contatos e temperatura da conta — use para importar Empresas no RD Station CRM."),
            ("Descadastrados (NÃO importar)", "Contatos que pediram descadastro na Locaweb. NÃO devem receber e-mail marketing (LGPD)."),
            ("Correções de E-mail", "Log de todos os e-mails corrigidos ou descartados (original → corrigido)."),
            ("Removidos", "Registros descartados: e-mails internos, testes, spam de formulário e registros sem nenhum contato válido."),
            ("Arquivos Processados", "Lista de todas as planilhas lidas da pasta do Drive e quantidade de registros."),
            ("", ""),
            ("REGRAS DE TRATAMENTO", ""),
            ("Telefones", "Padronizados em +55 (DDD) número. Números sem DDD receberam DDD 11 (operação em SP) — coluna 'DDD 11 assumido?'. Celulares antigos de 8 dígitos receberam o 9º dígito. Removidos prefixos de operadora (0XX15, 0XX21...). Telefones inválidos (DDD inexistente, poucos dígitos, 99999999) descartados."),
            ("E-mails", "Minúsculas, remoção de espaços/caracteres inválidos, @ faltante ou duplicado, extensões erradas (.con, .cm, .com.b, .combr, gmail.com.br...), provedores digitados errado (gmial, hotmial, yahho, outlok...) e domínios corporativos com 1 letra de diferença de um domínio frequente na base. Inválidos foram descartados (contato mantido se tiver telefone)."),
            ("Duplicados", "Contatos unificados quando têm o mesmo e-mail OU o mesmo celular (telefones compartilhados por 3+ e-mails, como PABX, não unificam). Os dados de todas as fontes foram consolidados na mesma linha (coluna 'Registros unificados')."),
            ("Segmento", "1) TLD institucional (.gov.br, .edu.br) → 2) palavras-chave no domínio/nome da empresa → 3) lista de segmento da Locaweb (agências, franquias, buffet...) → 4) demais regras. Classificação automática: revisar os casos de 'Empresas - Outros Segmentos'."),
            ("Data de entrada", "Data mais antiga em que o contato aparece: envio de formulário do site/B2B ou importação na Locaweb (as datas da Locaweb são lotes de importação: 07, 18 e 22/10/2024 e 30/07/2026 — o contato está na base pelo menos desde então). Para contatos que só existem em agendas/WhatsApp e exportações do Outlook (sem data por registro), foi usada a data da exportação. Ver coluna 'Origem da data de entrada'."),
            ("Última interação", "Somente datas reais de ação do lead (envio de formulário). É a base da recência na temperatura."),
            ("Temperatura (score 0-100)", "Recência da última interação real (≤90 dias: 45 pts; ≤180: 35; ≤365: 25; ≤2 anos: 12; >2 anos: 5) + lead inbound do site/chat (20) + troca de e-mails no Outlook com vendedora (15, +10 se 3+ e-mails) + contato salvo na agenda/WhatsApp (12) + clique em e-mail mkt (15) ou visualização (5) + presença em várias fontes (até 10) + perfil B2B (5) − bounce (10)."),
            ("Faixas", "QUENTE ≥ 55  •  MORNO 30–54  •  FRIO < 30  •  DESCADASTRADO = opt-out."),
            ("", ""),
            ("IMPORTAÇÃO NO RD STATION", ""),
            ("RD Station Marketing", "Use o CSV 'rdstation_marketing_contatos.csv' (somente contatos com e-mail válido, sem descadastrados e sem bounce). Mapeie: Nome, Email, Celular, Telefone, Empresa, Cidade, Estado; a coluna 'Tags' já traz segmento, temperatura, período e origem."),
            ("RD Station CRM", "Use 'rdstation_crm_empresas.csv' para Empresas e 'rdstation_crm_contatos.csv' para Contatos (inclui contatos só com telefone)."),
        ]
        for i, (a, b) in enumerate(linhas, start=3):
            ws.write(i, 0, a, sub if not b else wrap)
            ws.write(i, 1, b, wrap)

        # 2. Resumo
        ws = wb.add_worksheet("Resumo")
        xw.sheets["Resumo"] = ws
        ws.set_column(0, 0, 58); ws.set_column(1, 8, 14)
        ws.write(0, 0, "Resumo do processamento", title)
        for i, (k, v) in enumerate(kpis, start=2):
            ws.write(i, 0, k); ws.write_number(i, 1, v, num)
        r0 = len(kpis) + 4
        ws.write(r0, 0, "Contatos ativos por Segmento × Temperatura", sub)
        resumo_seg.reset_index().rename(columns={"index": "Segmento", "Segmento": "Segmento"}).to_excel(xw, sheet_name="Resumo", startrow=r0 + 1, index=False)
        r1 = r0 + len(resumo_seg) + 4
        ws.write(r1, 0, "Contatos ativos por Período de entrada × Temperatura", sub)
        resumo_per.reset_index().rename(columns={"index": "Período de entrada"}).to_excel(xw, sheet_name="Resumo", startrow=r1 + 1, index=False)
        r2 = r1 + len(resumo_per) + 4
        ws.write(r2, 0, "Contatos ativos por Origem", sub)
        orig = ativos["Origens"].str.split("; ").explode().value_counts().rename_axis("Origem").reset_index(name="Contatos")
        orig.to_excel(xw, sheet_name="Resumo", startrow=r2 + 1, index=False)

        write_table(ativos, "Base Completa")
        for t in ("Quente", "Morno", "Frio"):
            write_table(ativos[ativos["Temperatura"] == t], f"Leads {t}s")
        write_table(empresas, "Empresas (Contas)")
        for seg in ORDEM_ABAS_SEG:
            write_table(ativos[ativos["Segmento"] == seg], seg[:31])
        write_table(descad, "Descadastrados (NÃO importar)")
        write_table(pd.DataFrame(correcoes).drop_duplicates(), "Correções de E-mail")
        write_table(pd.DataFrame(removidos), "Removidos")
        write_table(arquivos, "Arquivos Processados")

    # CSVs para o RD Station
    mk = ativos[(ativos["E-mail"] != "") & (ativos["Status e-mail"] == "Válido")]
    rd_mkt = pd.DataFrame({
        "Nome": mk["Nome"], "Email": mk["E-mail"], "Celular": mk["Celular/WhatsApp"], "Telefone": mk["Telefone"],
        "Empresa": mk["Empresa"], "Cidade": mk["Cidade"], "Estado": mk["UF (pelo DDD)"], "Segmento": mk["Segmento"],
        "Temperatura": mk["Temperatura"], "Score": mk["Score (0-100)"], "Data de entrada": mk["Data de entrada na base"],
        "Origem": mk["Origens"], "Responsavel": mk["Responsável/Vendedora"], "Tags": mk["Tags RD Station"],
    })
    rd_mkt.to_csv(OUT / "rdstation_marketing_contatos.csv", index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)
    rd_crm = pd.DataFrame({
        "Nome": ativos["Nome"].where(ativos["Nome"] != "", ativos["E-mail"].str.split("@").str[0]),
        "Email": ativos["E-mail"], "Celular": ativos["Celular/WhatsApp"], "Telefone": ativos["Telefone"],
        "Empresa": ativos["Empresa"], "Segmento": ativos["Segmento"], "Temperatura": ativos["Temperatura"],
        "Score": ativos["Score (0-100)"], "Cidade": ativos["Cidade"], "Estado": ativos["UF (pelo DDD)"],
        "Data de entrada": ativos["Data de entrada na base"], "Origem": ativos["Origens"],
        "Responsavel": ativos["Responsável/Vendedora"], "Interesse": ativos["Interesse (último)"],
        "Status email": ativos["Status e-mail"], "Tags": ativos["Tags RD Station"],
    })
    rd_crm["Nome"] = rd_crm["Nome"].where(rd_crm["Nome"] != "", "Contato " + rd_crm["Celular"])
    rd_crm.to_csv(OUT / "rdstation_crm_contatos.csv", index=False, encoding="utf-8-sig")
    empresas.rename(columns={"Site/Domínio": "Site"}).to_csv(OUT / "rdstation_crm_empresas.csv", index=False, encoding="utf-8-sig")

    print("\n".join(f"{k}: {v}" for k, v in kpis))
    print(resumo_seg.to_string())
    print(resumo_per.to_string())
    print(f"\nArquivo gerado: {path}")


if __name__ == "__main__":
    main()
