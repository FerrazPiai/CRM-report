#!/usr/bin/env python3
"""
Clusterização da base de leads da Camiseta Express.

Lê todas as planilhas baixadas da pasta do cliente no Google Drive (formulários
do site/Elementor, listas da Locaweb, agendas de WhatsApp/Outlook e históricos
de e-mail das vendedoras), limpa e padroniza os contatos, remove duplicados
entre as fontes e gera uma planilha Excel separada por páginas/segmentos, com
tags para filtros e disparos.

Uso:
    python3 clusterizar_leads.py --entrada <pasta_com_planilhas> --saida <pasta_saida>
        [--data-referencia 2026-09-30]

A pasta de entrada pode ser o download (zip descompactado) da pasta do Drive; os
arquivos são localizados recursivamente e identificados pelas colunas/nomes.
"""
import argparse
import collections
import csv
import datetime as dt
import io
import json
import math
import os
import re
import unicodedata
import zipfile
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------------------
# Configuração
# --------------------------------------------------------------------------------------

DDD_PADRAO = "11"  # operação em São Paulo: telefone sem DDD recebe 11

# Qualquer domínio com "camisetaexpress" (inclui erros de digitação, Odoo e disparos Locaweb)
RE_DOMINIO_INTERNO = re.compile(r"camisetas?exp")
RE_EMAIL_SISTEMA = re.compile(r"^(no-?reply|notifications?|mailer-daemon|postmaster|donotreply|bounce)")
# Endereços que aparecem em cópia em praticamente todos os e-mails internos e
# contatos de teste do site. Revisar se algum deles for de fato cliente.
EMAILS_INTERNOS = {
    "thiagosilva.webdev@gmail.com",
    "sr.garcia2015.sg@gmail.com",
    "priscila.stamp2@gmail.com",
}
TELEFONES_INTERNOS = {"+5511952532669"}

# Datas em que cada planilha "sem data por linha" entrou na pasta (nome do arquivo
# ou data de criação no Drive). Usadas como data de entrada desses contatos.
DATAS_SNAPSHOT = {
    "contatoscamex": "2026-07-24",
    "contatosdebora": "2026-09-21",
    "contatossilmara": "2026-09-22",
    "emailsdebora": "2026-09-22",
    "emailsrose": "2026-09-23",
    "leadstestchat": "2026-09-21",
    "camisetaexpressleadstestchat": "2026-09-21",
    "listaemailssubmissoes": "2026-09-18",
    "camisetaexpresslistadeemailssubmissoes": "2026-09-18",
}

PROVEDORES_PESSOAIS = {
    "gmail.com", "googlemail.com", "hotmail.com", "hotmail.com.br", "outlook.com", "outlook.com.br",
    "live.com", "live.com.br", "msn.com", "yahoo.com", "yahoo.com.br", "ymail.com", "rocketmail.com",
    "uol.com.br", "bol.com.br", "terra.com.br", "ig.com.br", "icloud.com", "me.com", "mac.com",
    "globo.com", "globomail.com", "r7.com", "zipmail.com.br", "oi.com.br", "protonmail.com",
    "proton.me", "aol.com", "gmx.com", "yandex.com", "mail.com", "superig.com.br", "pop.com.br",
    "click21.com.br", "itelefonica.com.br", "veloxmail.com.br", "brturbo.com.br", "ibest.com.br",
}
CORRECOES_DOMINIO = {
    "gmail.con": "gmail.com", "gmail.co": "gmail.com", "gmial.com": "gmail.com", "gmai.com": "gmail.com",
    "gamil.com": "gmail.com", "gnail.com": "gmail.com", "gmail.com.br": "gmail.com", "gmaill.com": "gmail.com",
    "hotmail.con": "hotmail.com", "hotmal.com": "hotmail.com", "hotmial.com": "hotmail.com",
    "hotmai.com": "hotmail.com", "hotmil.com": "hotmail.com", "outlook.con": "outlook.com",
    "yahoo.con": "yahoo.com", "yahoo.com.b": "yahoo.com.br", "uol.com": "uol.com.br",
    "hormail.com": "hotmail.com", "hotmail.co": "hotmail.com", "outlok.com": "outlook.com",
}
# Domínios de robôs/spam vistos nos formulários e listas
DOMINIOS_SPAM = {"course-fitness.com", "inrus.top"}
TLDS_SPAM = {"ru", "top", "xyz", "icu", "click", "cn", "su", "online", "site", "store", "shop"}
# Domínios compartilhados por pessoas de organizações diferentes
DOMINIOS_GENERICOS = {"educacao", "prof", "professor", "aluno", "alunos"}
# Rótulos de 2º nível removidos após o TLD (ex.: empresa.com.br, tjsp.jus.br, sp.gov.br)
SEGUNDO_NIVEL = {
    "com", "org", "net", "gov", "edu", "coop", "ind", "art", "adv", "eng", "co", "go", "ac", "mil", "jus",
    "mp", "leg", "ne", "or", "tur", "med", "arq", "srv", "emp", "inf", "blog", "etc", "g12",
    "sp", "rj", "mg", "pr", "rs", "sc", "ba", "pe", "ce", "df", "ppg", "psi", "far", "fot", "rec", "eco",
    "agr", "bio", "vet", "odo", "tmp", "esp", "cnt", "adm", "jor", "mus", "pro", "nom", "log", "imb",
}
# Nomes de exibição para marcas conhecidas (raiz do domínio -> nome)
ALIAS_EMPRESAS = {
    "carrefour": "Carrefour", "gpabr": "GPA (Pão de Açúcar)", "safra": "Banco Safra", "vigor": "Vigor",
    "atacadao": "Atacadão", "assai": "Assaí Atacadista", "nestle": "Nestlé", "ambev": "Ambev",
    "itau": "Itaú", "itauunibanco": "Itaú", "bradesco": "Bradesco", "natura": "Natura",
    "webmotors": "Webmotors", "thermofisher": "Thermo Fisher", "mdiasbranco": "M. Dias Branco",
    "santander": "Santander", "unifesp": "Unifesp", "usp": "USP", "samsung": "Samsung",
    "wellhub": "Wellhub", "ubsbb": "UBS BB", "panco": "Panco", "jica": "JICA", "unicamp": "Unicamp",
    "renovaenergia": "Renova Energia", "sbasite": "SBA", "cooperata": "Cooperata",
    "kcc": "Kimberly-Clark", "colpal": "Colgate-Palmolive", "basf": "BASF", "whirlpool": "Whirlpool",
    "multibras": "Whirlpool", "multivarejogpa": "GPA (Pão de Açúcar)", "itau-unibanco": "Itaú",
    "redecard": "Rede (Itaú)", "voeazul": "Azul Linhas Aéreas", "grupobimbo": "Grupo Bimbo",
    "hnipo": "Hospital Nipo-Brasileiro", "barry-callebaut": "Barry Callebaut", "unilever": "Unilever",
    "heineken": "Heineken", "alpargatas": "Alpargatas", "luxottica": "Luxottica", "fgv": "FGV",
    "senai": "SENAI", "sulamerica": "SulAmérica", "educacao": "Secretaria da Educação (rede pública)",
    "tjsp": "Tribunal de Justiça de SP", "metrosp": "Metrô SP", "btgpactual": "BTG Pactual",
    "bcgbrasil": "BCG Brasil", "amazon": "Amazon", "shopee": "Shopee", "google": "Google",
    "roldao": "Roldão Atacadista", "extrafarma": "Extrafarma", "tirolez": "Tirolez",
}
# Palavras que parecem empresa no nome do contato, mas são bairro/genéricas
NAO_E_EMPRESA_EM_NOME = {"digital", "prefeitura", "alphaville", "palmeiras", "brasil", "grupo", "eventos",
                         "comercial", "escola", "colegio", "igreja", "academia", "studio", "design"}
# Marcas buscadas também no nome dos contatos das agendas (ex.: "Paulo Carrefour")
MARCAS_EM_NOME = {
    "carrefour": "carrefour", "nestle": "nestle", "safra": "safra", "vigor": "vigor",
    "atacadao": "atacadao", "assai": "assai", "ambev": "ambev", "itau": "itau", "bradesco": "bradesco",
    "natura": "natura", "santander": "santander", "webmotors": "webmotors", "gpa": "gpabr",
    "samsung": "samsung", "unilever": "unilever", "heineken": "heineken", "danone": "danone",
    "pepsico": "pepsico", "cocacola": "cocacola", "bayer": "bayer", "raizen": "raizen",
    "petrobras": "petrobras", "magalu": "magalu", "renner": "renner", "riachuelo": "riachuelo",
    "sodexo": "sodexo", "ifood": "ifood", "nubank": "nubank", "votorantim": "votorantim",
}

# Listas da Locaweb -> (vertical, carteira, núcleo sugerido)
LISTAS_LOCAWEB = {
    "agencias_de_mkt_e_publicidade": ("Agências de Marketing & Publicidade", None, "canal"),
    "abap": ("Agências de Marketing & Publicidade", None, "canal"),
    "brasil_promotions": ("Promocional & Brindes", None, "canal"),
    "impressao_digital": ("Gráficas & Impressão Digital", None, "canal"),
    "atacadistas": ("Atacadistas", None, "b2b"),
    "empresas_alimenticias": ("Indústria Alimentícia", None, "b2b"),
    "buffet_infantil": ("Buffet Infantil & Eventos", None, "b2b"),
    "condominios": ("Condomínios", None, "b2b"),
    "escola_de_ingles": ("Escolas de Idiomas", None, "b2b"),
    "franquias": ("Franquias", None, "b2b"),
    "vendedoras": (None, "Vendedoras", None),
    "rose": (None, "Rose", None),
    "leads_debora": (None, "Débora", None),
    "jose_luiz": (None, "José Luiz", None),
    "store": (None, "Loja", "loja"),
    "lead_site_orcamento": (None, None, None),
    "contatos_envios_site_geral": (None, None, None),
}
# Formulários que não indicam intenção de compra
FORMS_SEM_INTENCAO = ("newsletter", "new form", "atualização b2b", "atualizacao b2b")

UF_POR_DDD = {
    **{d: "SP" for d in ("11", "12", "13", "14", "15", "16", "17", "18", "19")},
    **{d: "RJ" for d in ("21", "22", "24")}, **{d: "ES" for d in ("27", "28")},
    **{d: "MG" for d in ("31", "32", "33", "34", "35", "37", "38")},
    **{d: "PR" for d in ("41", "42", "43", "44", "45", "46")}, **{d: "SC" for d in ("47", "48", "49")},
    **{d: "RS" for d in ("51", "53", "54", "55")}, "61": "DF", "62": "GO", "64": "GO", "63": "TO",
    "65": "MT", "66": "MT", "67": "MS", "68": "AC", "69": "RO", "71": "BA", "73": "BA", "74": "BA",
    "75": "BA", "77": "BA", "79": "SE", "81": "PE", "87": "PE", "82": "AL", "83": "PB", "84": "RN",
    "85": "CE", "88": "CE", "86": "PI", "89": "PI", "91": "PA", "93": "PA", "94": "PA", "92": "AM",
    "97": "AM", "95": "RR", "96": "AP", "98": "MA", "99": "MA",
}

RE_EMAIL = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
RE_SPAM = re.compile(
    r"[Ѐ-ӿ぀-ヿ一-鿿가-힯]|\bhi there\b|\bseo\b|backlink|casino|"
    r"crypto|viagra|porn|\bloan\b|bitcoin|\bdear\b|click here|unsubscribe", re.I)
RE_QTD = re.compile(
    r"(\d{1,3}(?:[.\s]\d{3})*|\d+)\s*(?:mil\s*)?(?:un\b|und|unid|unidades|camis|pe[çc]as|p[çc]s|kits?|"
    r"moletons?|polos?|bon[eé]s|uniformes|coletes|jaquetas|pe[çc]a)", re.I)
PAGINAS_GENERICAS = {"orcamento", "catalogo", "faq", "contato", "home", "sobre nos", "blog", "obrigado",
                     "fale conosco", "politica de privacidade"}
STOP_EN = {"the", "and", "for", "with", "your", "you", "our", "we", "is", "are", "this", "that", "of",
           "to", "much", "more", "get", "can", "will", "from", "here", "have", "business", "website"}
STOP_PT = {"de", "para", "com", "que", "uma", "um", "gostaria", "camisetas", "camiseta", "orçamento",
           "orcamento", "empresa", "boa", "bom", "dia", "tarde", "olá", "ola", "preciso", "quero", "valor"}


def eh_spam(nome, msg, email):
    texto = f"{nome} {msg}"
    if RE_SPAM.search(texto):
        return True
    email = (email or "").lower().strip()
    dom = email.rsplit("@", 1)[-1] if "@" in email else ""
    if dom in DOMINIOS_SPAM or dom.rsplit(".", 1)[-1] in TLDS_SPAM:
        return True
    tokens = re.findall(r"[a-zçãõáéíóúâêô']+", msg.lower())
    en = sum(t in STOP_EN for t in tokens)
    pt = sum(t in STOP_PT for t in tokens)
    return en >= 3 and pt <= 1


LIXO_EMPRESA = {
    "", "-", "--", ".", "x", "xx", "na", "n/a", "nao", "não", "nenhuma", "nenhum", "pessoal", "personal",
    "pessoa fisica", "pessoa física", "particular", "autonomo", "autônomo", "autonoma", "autônoma",
    "eu", "teste", "test", "casa", "familia", "família", "nada", "sem empresa", "pf", "cpf", "ok",
    "nome", "empresa", "mei", "0", "00", "sim", "privado", "individual", "pessoa", "minha",
}


# --------------------------------------------------------------------------------------
# Utilidades de limpeza
# --------------------------------------------------------------------------------------

def sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def chave_texto(s):
    return re.sub(r"[^a-z0-9]+", " ", sem_acento(str(s)).lower()).strip()


def decodificar_escapes(s):
    """Elementor exporta alguns campos com \\u00e1 e \\/ literais."""
    s = str(s)
    if "\\u" in s or "\\/" in s or "\\r" in s or "\\n" in s:
        try:
            s = json.loads('"' + s.replace('"', '\\"') + '"')
        except Exception:
            s = s.replace("\\/", "/").replace("\\r", " ").replace("\\n", " ")
    return s


def limpar_nome(nome):
    nome = decodificar_escapes(nome or "").strip()
    if not nome or nome.lower() in ("nan", "none"):
        return ""
    # "A d r i a n a N e v e s" -> "Adriana Neves"
    partes = nome.split(" ")
    if len(partes) >= 4 and sum(len(p) == 1 for p in partes) / len(partes) > 0.8:
        nome = re.sub(r"(?<=\w) (?=\w(?:\s|$))", "", nome)
        nome = re.sub(r"(?<=[a-zà-ú])(?=[A-ZÀ-Ú])", " ", nome)
    nome = re.sub(r"[\U00010000-\U0010ffff☀-➿️]", "", nome)  # emojis
    nome = re.sub(r"\s+", " ", nome).strip(" -_.,;")
    if not re.search(r"[A-Za-zÀ-ú]", nome) or "@" in nome or "http" in nome.lower():
        return ""
    if nome.isupper() or nome.islower():
        nome = " ".join(p.capitalize() if len(p) > 2 else p.lower() for p in nome.split())
        nome = nome[:1].upper() + nome[1:]
    return nome


def nome_de_email(email):
    local = email.split("@")[0]
    partes = [p for p in re.split(r"[._-]+", local) if p.isalpha() and len(p) >= 3]
    if len(partes) >= 2:
        return " ".join(p.capitalize() for p in partes[:3])
    return ""


def normalizar_email(raw):
    raw = decodificar_escapes(raw or "").strip().lower().replace("mailto:", "")
    m = RE_EMAIL.search(raw)
    if not m:
        return None
    email = m.group(0).strip(".'")
    local, dom = email.rsplit("@", 1)
    dom = CORRECOES_DOMINIO.get(dom, dom)
    if ".." in email or len(local) < 1 or dom.count(".") < 1:
        return None
    return f"{local}@{dom}"


def normalizar_telefone(raw):
    """Retorna (e164, tipo, ddd_assumido) ou None. Telefones sem DDD recebem 11."""
    s = str(raw or "").strip()
    if not s or s.lower() in ("nan", "none"):
        return None
    d = re.sub(r"\D", "", s)
    if len(d) < 8:
        return None
    if s.lstrip().startswith("+") and not d.startswith("55"):
        return ("+" + d, "internacional", False) if 10 <= len(d) <= 15 else None
    if d.startswith("00"):
        d = d[2:]
        if not d.startswith("55"):
            return ("+" + d, "internacional", False) if 10 <= len(d) <= 15 else None
    if d.startswith("0") and len(d) in (13, 14):  # 0 + operadora + DDD + número
        d = d[3:]
    elif d.startswith("0") and len(d) in (11, 12):  # 0 + DDD + número
        d = d[1:]
    if d.startswith("55") and len(d) in (12, 13):
        d = d[2:]
    ddd_assumido = False
    if len(d) in (8, 9):
        d = DDD_PADRAO + d
        ddd_assumido = True
    if len(d) not in (10, 11):
        return None
    ddd, num = d[:2], d[2:]
    if ddd not in UF_POR_DDD:
        return None
    if len(num) == 8 and num[0] in "6789":
        num = "9" + num  # celular antigo sem o nono dígito
    if len(num) == 9 and num[0] != "9":
        return None
    if len(set(num)) <= 2 or num in ("912345678", "987654321"):
        return None
    tipo = "celular" if len(num) == 9 else "fixo"
    return ("+55" + ddd + num, tipo, ddd_assumido)


def formatar_telefone(e164):
    if not e164 or not e164.startswith("+55"):
        return e164 or ""
    d = e164[3:]
    ddd, num = d[:2], d[2:]
    return f"({ddd}) {num[:-4]}-{num[-4:]}"


def raiz_dominio(dom):
    labels = dom.lower().split(".")
    if len(labels) > 1:
        labels.pop()  # TLD (.br, .com, .digital...)
    while len(labels) > 1 and labels[-1] in SEGUNDO_NIVEL:
        labels.pop()
    return labels[-1]


def empresa_valida(txt):
    t = decodificar_escapes(txt or "").strip()
    k = chave_texto(t)
    if k in LIXO_EMPRESA or len(k) < 2 or re.fullmatch(r"[\d\s./-]+", t):
        return ""
    if re.search(r"\d{3}\.\d{3}\.\d{3}-\d{2}", t) or "@" in t or "http" in t.lower():
        return ""
    return re.sub(r"\s+", " ", t)[:80]


def parse_data(v):
    v = str(v or "").strip()
    if not v or v.lower() in ("nan", "none", "nat"):
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f",
                "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(v, fmt)
        except ValueError:
            pass
    try:
        return pd.to_datetime(v, dayfirst=True).to_pydatetime()
    except Exception:
        return None


def estimar_volume(*textos):
    melhor = None
    for t in textos:
        t = decodificar_escapes(t or "")
        if not t:
            continue
        if re.search(r"mais de 10[01]", t, re.I):
            melhor = max(melhor or 0, 101)
        elif re.search(r"100 pe[çc]as ou menos", t, re.I):
            melhor = max(melhor or 0, 100)
        elif re.search(r"20 pe[çc]as ou menos", t, re.I):
            melhor = max(melhor or 0, 20)
        elif re.search(r"21 [àa] 50", t, re.I):
            melhor = max(melhor or 0, 50)
        for m in RE_QTD.finditer(t):
            n = int(re.sub(r"\D", "", m.group(1)))
            if "mil" in m.group(0).lower():
                n *= 1000
            if 1 <= n <= 200000:
                melhor = max(melhor or 0, n)
    return melhor


# --------------------------------------------------------------------------------------
# Leitura das fontes
# --------------------------------------------------------------------------------------

def ler_csv(p):
    raw = Path(p).read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            txt = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    try:
        sep = csv.Sniffer().sniff(txt[:5000], delimiters=",;\t").delimiter
    except csv.Error:
        sep = ","
    return pd.read_csv(io.StringIO(txt), sep=sep, dtype=str, keep_default_na=False, on_bad_lines="skip")


def ler_planilhas(pasta):
    for p in sorted(Path(pasta).rglob("*")):
        if p.is_dir() or p.name.startswith((".", "~$")):
            continue
        ext = p.suffix.lower()
        try:
            if ext == ".csv":
                yield p, "", ler_csv(p)
            elif ext in (".xlsx", ".xlsm", ".xls"):
                for aba, df in pd.read_excel(p, sheet_name=None, dtype=str, keep_default_na=False).items():
                    yield p, aba, df
        except Exception as e:  # noqa: BLE001
            print(f"  ! não foi possível ler {p.name}: {e}")


def stem_simples(p):
    return re.sub(r"[^a-z]", "", sem_acento(p.stem.split("_todos")[0]).lower())


def data_snapshot(p):
    s = stem_simples(p)
    for k, v in DATAS_SNAPSHOT.items():
        if s.startswith(k):
            return dt.datetime.fromisoformat(v)
    m = re.search(r"(\d{2})[-/.](\d{2})(?:[-/.](\d{2,4}))?", p.stem)
    if m:
        ano = int(m.group(3)) if m.group(3) else dt.date.today().year
        ano = ano + 2000 if ano < 100 else ano
        try:
            return dt.datetime(ano, int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    return dt.datetime.fromtimestamp(p.stat().st_mtime)


def col(df, *nomes, contem=None):
    mapa = {chave_texto(c): c for c in df.columns}
    for n in nomes:
        if chave_texto(n) in mapa:
            return mapa[chave_texto(n)]
    if contem:
        for k, c in mapa.items():
            if any(x in k for x in contem):
                return c
    return None


def extrair_registros(pasta):
    """Converte cada linha de cada planilha em um registro bruto padronizado."""
    registros = []
    interacoes = collections.Counter()  # email -> nº de e-mails no histórico do Outlook
    dono_interacao = collections.defaultdict(set)
    resumo_fontes = []

    for p, aba, df in ler_planilhas(pasta):
        if df.empty:
            continue
        cols = [chave_texto(c) for c in df.columns]
        stem = stem_simples(p)
        nome_arq = p.name + (f" [{aba}]" if aba else "")
        antes = len(registros)

        # 1) Formulários do site (Elementor)
        if "nome do formulario id" in cols:
            c_form = col(df, "Nome do Formulário (ID)")
            c_data = col(df, "Criado em")
            c_nome, c_sobren = col(df, "Nome", "name"), col(df, "Sobrenome", "sobrenome")
            c_email = col(df, "E-mail", "Email", contem=["mail"])
            c_tel = col(df, "DDD + WhatsApp", "Telefone", contem=["whatsapp", "telefone"])
            c_tel2 = col(df, "WhatsApp")
            c_emp = col(df, "Empresa")
            c_int = col(df, "Interesse em:", "Interesse")
            c_pag = col(df, "Página")
            c_msg = col(df, "Mensagem")
            c_ass = col(df, "Assunto")
            c_qtd = col(df, "unidades", "Quantos Kits", "field_e85e048")
            for _, r in df.iterrows():
                form = re.sub(r"\s*\([0-9a-f]+\)\s*$", "", decodificar_escapes(r[c_form])).strip()
                msg = decodificar_escapes(r[c_msg]) if c_msg else ""
                nome = " ".join(x for x in (r[c_nome] if c_nome else "", r[c_sobren] if c_sobren else "") if x)
                spam = eh_spam(nome, msg, r[c_email] if c_email else "")
                assunto = decodificar_escapes(r[c_ass]) if c_ass else ""
                pagina = decodificar_escapes(r[c_pag]) if c_pag else ""
                interesse = decodificar_escapes(r[c_int]) if c_int else ""
                registros.append(dict(
                    fonte="Site - Formulários", arquivo=nome_arq, detalhe=form,
                    nome=nome, email=r[c_email] if c_email else "",
                    telefones=[r[c_tel] if c_tel else "", r[c_tel2] if c_tel2 else ""],
                    empresa=r[c_emp] if c_emp else "", data=parse_data(r[c_data]), data_tipo="formulário",
                    solicitacao=not any(x in form.lower() for x in FORMS_SEM_INTENCAO),
                    interesse=interesse or (pagina if pagina not in ("Home", "Contato", "Sobre nós") else ""),
                    assunto=assunto, mensagem=msg[:500], spam=spam,
                    volume=estimar_volume(r[c_qtd] if c_qtd else "", msg),
                ))

        # 2) Listas de e-mail marketing Locaweb
        elif {"email", "status", "data de cadastro"} <= set(cols):
            lista = re.sub(r"_todos.*$", "", p.stem)
            vertical, carteira, _ = LISTAS_LOCAWEB.get(lista, (None, None, None))
            for _, r in df.iterrows():
                def n(c):
                    try:
                        return int(float(r[c] or 0))
                    except (ValueError, KeyError):
                        return 0
                registros.append(dict(
                    fonte="Locaweb - E-mail Mkt", arquivo=nome_arq, detalhe=lista,
                    nome=r.get("Nome", ""), email=r["email"], telefones=[], empresa="",
                    data=parse_data(r["data de cadastro"]), data_tipo="cadastro Locaweb",
                    lw_status=r["status"], lw_views=n("visualizações únicas"), lw_cliques=n("cliques"),
                    lw_erros=n("erros de recebimento"), vertical=vertical, carteira=carteira,
                    lista_locaweb=lista, cidade=r.get("Cidade", ""), spam=eh_spam("", "", r["email"]),
                ))

        # 3) Histórico/lista de e-mails das vendedoras (Outlook): 1 coluna com e-mails
        elif stem.startswith("emails") and len(df.columns) <= 3:
            dono = p.stem[6:].split("2")[0].split("0")[0].split("1")[0].strip("_- ").title() or "Outlook"
            dono = {"Debora": "Débora"}.get(dono, dono)
            valores = list(df.columns) + [v for c in df.columns for v in df[c].tolist()]
            vistos = set()
            for v in valores:
                for e in RE_EMAIL.findall(str(v)):
                    e = normalizar_email(e)
                    if not e:
                        continue
                    interacoes[e] += 1
                    dono_interacao[e].add(dono)
                    if e not in vistos:
                        vistos.add(e)
                        registros.append(dict(
                            fonte="Outlook - Histórico de e-mails", arquivo=nome_arq, detalhe=f"Outlook {dono}",
                            nome="", email=e, telefones=[], empresa="", data=data_snapshot(p),
                            data_tipo="entrada na planilha", carteira=dono,
                        ))

        # 4) Agendas / demais planilhas de contatos (WhatsApp, Outlook, chat, Google Sheets do site)
        else:
            c_nome = col(df, "Nome Completo", "Display Name", "Nome", "Name")
            c_first, c_last = col(df, "First Name"), col(df, "Last Name")
            c_email = col(df, "Email", "E-mail", "E-mail Address", contem=["mail"])
            c_tels = [c for c in df.columns if re.search(r"phone|telefone|celular|whats|fone", chave_texto(c))]
            c_emp = col(df, "Empresa", "Company")
            c_cid = col(df, "Cidade", "City")
            c_data = col(df, "Data/Hora", "Data", "Criado em")
            if not (c_email or c_tels):
                print(f"  ? ignorado (sem e-mail/telefone): {nome_arq}")
                continue
            if stem.startswith("contatos"):
                dono = re.sub(r"\d.*$", "", p.stem[8:]).strip("_- ") or "Agenda"
                dono = {"Debora": "Débora", "Camex": "WhatsApp Camex"}.get(dono, dono)
                fonte, detalhe, carteira = "Agendas WhatsApp/Outlook", f"Agenda {dono}", dono
            elif "chat" in stem:
                fonte, detalhe, carteira = "Site - Chat", "Leads chat do site", None
            elif "atualiza" in stem:
                fonte, detalhe, carteira = "Site - Formulários", "Atualização de Cadastro B2B", None
            else:
                fonte, detalhe, carteira = "Site - Formulários", "Lista de e-mails de submissões", None
            snap = data_snapshot(p)
            for _, r in df.iterrows():
                nome = r[c_nome] if c_nome else ""
                if c_first and not nome:
                    nome = f"{r[c_first]} {r[c_last] if c_last else ''}"
                data = parse_data(r[c_data]) if c_data else None
                registros.append(dict(
                    fonte=fonte, arquivo=nome_arq, detalhe=detalhe, nome=nome,
                    email=r[c_email] if c_email else "", telefones=[r[c] for c in c_tels],
                    empresa=r[c_emp] if c_emp else "", data=data or snap,
                    data_tipo="formulário" if data else "entrada na planilha", carteira=carteira,
                    cidade=(r[c_cid] if c_cid and r[c_cid] != "Unknown" else ""),
                    solicitacao=False,
                ))
        resumo_fontes.append(dict(arquivo=nome_arq, linhas=len(df), registros=len(registros) - antes))

    return registros, interacoes, dono_interacao, pd.DataFrame(resumo_fontes)


# --------------------------------------------------------------------------------------
# Deduplicação (union-find por e-mail e telefone)
# --------------------------------------------------------------------------------------

class UnionFind:
    def __init__(self):
        self.pai = {}

    def achar(self, x):
        self.pai.setdefault(x, x)
        while self.pai[x] != x:
            self.pai[x] = self.pai[self.pai[x]]
            x = self.pai[x]
        return x

    def unir(self, a, b):
        ra, rb = self.achar(a), self.achar(b)
        if ra != rb:
            self.pai[rb] = ra


def consolidar(registros, interacoes, dono_interacao, ref):
    # normalização por registro
    for i, r in enumerate(registros):
        r["id"] = i
        r["email_n"] = normalizar_email(r.get("email"))
        tels = []
        for t in r.get("telefones") or []:
            nt = normalizar_telefone(t)
            if nt and nt[0] not in [x[0] for x in tels]:
                tels.append(nt)
        r["tels_n"] = tels
        r["nome_l"] = limpar_nome(r.get("nome"))
        r["empresa_l"] = empresa_valida(r.get("empresa"))
        e = r["email_n"] or ""
        r["interno"] = (bool(RE_DOMINIO_INTERNO.search(e.split("@")[-1])) or bool(RE_EMAIL_SISTEMA.match(e))
                        or e in EMAILS_INTERNOS
                        or any(t[0] in TELEFONES_INTERNOS for t in tels)
                        or chave_texto(r["nome_l"]) in ("teste", "test", "thiago teste"))

    # telefones compartilhados por muitos e-mails diferentes não servem de chave de ligação
    emails_por_tel = collections.defaultdict(set)
    for r in registros:
        for t in r["tels_n"]:
            if r["email_n"]:
                emails_por_tel[t[0]].add(r["email_n"])
    tel_generico = {t for t, es in emails_por_tel.items() if len(es) > 3}

    uf = UnionFind()
    for r in registros:
        chaves = ([f"e:{r['email_n']}"] if r["email_n"] else []) + \
                 [f"t:{t[0]}" for t in r["tels_n"] if t[0] not in tel_generico]
        if not chaves:
            continue
        for k in chaves:
            uf.unir(chaves[0], k)
        r["grupo"] = uf.achar(chaves[0])
    grupos = collections.defaultdict(list)
    for r in registros:
        if "grupo" in r:
            grupos[uf.achar(r["grupo"])].append(r)

    # empresas por domínio: nome mais citado no campo "Empresa" dos formulários
    nomes_por_raiz = collections.defaultdict(collections.Counter)
    contatos_por_raiz = collections.Counter()
    for rs in grupos.values():
        raizes = set()
        for r in rs:
            e = r["email_n"]
            if e and e.split("@")[1] not in PROVEDORES_PESSOAIS and not r["interno"]:
                raiz = raiz_dominio(e.split("@")[1])
                raizes.add(raiz)
                if r["empresa_l"]:
                    nomes_por_raiz[raiz][r["empresa_l"]] += 1
        for rz in raizes:
            contatos_por_raiz[rz] += 1

    def nome_empresa(raiz):
        if raiz in ALIAS_EMPRESAS:
            return ALIAS_EMPRESAS[raiz]
        cand = nomes_por_raiz.get(raiz)
        if cand:
            nome, _ = cand.most_common(1)[0]
            if chave_texto(nome).replace(" ", "")[:4] == raiz[:4] or len(cand) == 1:
                return nome if not nome.islower() else nome.title()
        return raiz.replace("-", " ").title() if len(raiz) > 3 else raiz.upper()

    # vocabulário para achar empresa no nome dos contatos de agenda ("Paulo Carrefour")
    vocab = dict(MARCAS_EM_NOME)
    for raiz, n in contatos_por_raiz.items():
        if n >= 5 and len(raiz) >= 6 and raiz.isalpha() and raiz not in NAO_E_EMPRESA_EM_NOME:
            vocab.setdefault(raiz, raiz)

    contatos = []
    for gid, rs in grupos.items():
        c = montar_contato(rs, interacoes, dono_interacao, ref, nome_empresa, vocab)
        contatos.append(c)
    df = pd.DataFrame(contatos)
    return df


def montar_contato(rs, interacoes, dono_interacao, ref, nome_empresa, vocab):
    rs = sorted(rs, key=lambda r: (r.get("data") or dt.datetime.min))
    emails = []
    for r in reversed(rs):
        if r["email_n"] and r["email_n"] not in emails:
            emails.append(r["email_n"])
    corp = [e for e in emails if e.split("@")[1] not in PROVEDORES_PESSOAIS]
    email_principal = (corp or emails or [""])[0]

    tels = {}
    for r in reversed(rs):
        for t in r["tels_n"]:
            tels.setdefault(t[0], t)
    tels_ord = sorted(tels.values(), key=lambda t: {"celular": 0, "fixo": 1}.get(t[1], 2))
    tel_principal = tels_ord[0] if tels_ord else None

    prioridade_nome = {"Site - Formulários": 0, "Site - Chat": 0, "Locaweb - E-mail Mkt": 2,
                       "Agendas WhatsApp/Outlook": 1, "Outlook - Histórico de e-mails": 3}
    nomes = sorted([r for r in rs if r["nome_l"]],
                   key=lambda r: (prioridade_nome.get(r["fonte"], 5), -len(r["nome_l"].split())))
    nome = nomes[0]["nome_l"] if nomes else ""
    nome_origem = nomes[0]["fonte"] if nomes else ""
    if not nome and email_principal:
        nome = nome_de_email(email_principal)
        nome_origem = "derivado do e-mail" if nome else ""

    # Empresa: formulário > domínio corporativo > marca citada no nome do contato
    empresa, empresa_origem, dominio = "", "", ""
    emp_forms = [r["empresa_l"] for r in rs if r["empresa_l"]]
    if corp:
        dominio = corp[0].split("@")[1]
        if raiz_dominio(dominio) in DOMINIOS_GENERICOS and emp_forms:
            empresa, empresa_origem = emp_forms[-1], "informada no formulário"
        else:
            empresa, empresa_origem = nome_empresa(raiz_dominio(dominio)), "domínio do e-mail"
    if emp_forms and not empresa:
        empresa, empresa_origem = emp_forms[-1], "informada no formulário"
    if not empresa:
        tokens = set(chave_texto(" ".join(r.get("nome") or "" for r in rs)).split())
        achadas = [vocab[t] for t in tokens if t in vocab]
        if achadas:
            empresa, empresa_origem = nome_empresa(achadas[0]), "citada no nome do contato"
            dominio = dominio or ""

    fontes = sorted({r["fonte"] for r in rs})
    detalhes = sorted({r["detalhe"] for r in rs if r.get("detalhe")})
    carteiras = sorted({r["carteira"] for r in rs if r.get("carteira")} |
                       {d for e in emails for d in dono_interacao.get(e, ())})
    verticais = sorted({r["vertical"] for r in rs if r.get("vertical")})
    listas = sorted({r["lista_locaweb"] for r in rs if r.get("lista_locaweb")})
    nucleos_lista = {LISTAS_LOCAWEB.get(l, (None, None, None))[2] for l in listas} - {None}

    solic = [r for r in rs if r.get("solicitacao") and not r.get("spam") and r.get("data")]
    datas = [r["data"] for r in rs if r.get("data")]
    primeira = min(datas) if datas else None
    ultima_solic = max((r["data"] for r in solic), default=None)
    ultima_mov = max(datas) if datas else None
    dias_solic = (ref - ultima_solic).days if ultima_solic else None

    lw = [r for r in rs if r["fonte"].startswith("Locaweb")]
    status_lw = {r["lw_status"] for r in lw}
    optout = "Descadastrado" in status_lw
    bounce = "Desativado" in status_lw
    lw_views = sum(r["lw_views"] for r in lw)
    lw_cliques = sum(r["lw_cliques"] for r in lw)
    outlook = sum(interacoes.get(e, 0) for e in emails)

    interesses = []
    for r in rs:
        v = (r.get("interesse") or "").strip()
        v = re.sub(r"^\d+\s*[–-]\s*", "", v)
        if v and v not in interesses and len(v) < 60 and chave_texto(v) not in PAGINAS_GENERICAS:
            interesses.append(v)
    msgs = [r["mensagem"] for r in reversed(solic) if r.get("mensagem")]
    volume = max((r.get("volume") or 0 for r in rs), default=0) or None
    cidade = next((r["cidade"] for r in rs if r.get("cidade")), "")
    interno = any(r["interno"] for r in rs)
    n_spam = sum(bool(r.get("spam")) for r in rs)
    spam = n_spam > 0 and n_spam >= len(rs) / 2 and not solic

    email_ok = bool(email_principal) and not optout and not (bounce and len(emails) == 1)
    whats_ok = bool(tel_principal) and tel_principal[1] == "celular"

    return dict(
        nome=nome, empresa=empresa, empresa_origem=empresa_origem, dominio=dominio,
        email=email_principal, emails_adicionais="; ".join(e for e in emails if e != email_principal),
        tipo_email=("corporativo" if corp else "pessoal") if emails else "",
        telefone=formatar_telefone(tel_principal[0]) if tel_principal else "",
        whatsapp=(tel_principal[0].lstrip("+") if whats_ok else ""),
        tipo_telefone=tel_principal[1] if tel_principal else "",
        ddd_assumido_11=bool(tel_principal and tel_principal[2]),
        telefones_adicionais="; ".join(formatar_telefone(t[0]) for t in tels_ord[1:]),
        ddd=(tel_principal[0][3:5] if tel_principal and tel_principal[0].startswith("+55") else ""),
        uf=(UF_POR_DDD.get(tel_principal[0][3:5], "") if tel_principal and tel_principal[0].startswith("+55")
            else ("Exterior" if tel_principal else "")),
        cidade=cidade,
        fontes="; ".join(fontes), qtd_fontes=len(fontes), origens="; ".join(detalhes),
        qtd_registros_brutos=len(rs), carteira_vendedora="; ".join(carteiras),
        listas_locaweb="; ".join(listas), vertical="; ".join(verticais),
        data_entrada=primeira, data_entrada_tipo=(rs[0].get("data_tipo") if rs and rs[0].get("data") else ""),
        ultima_solicitacao=ultima_solic, dias_desde_solicitacao=dias_solic, ultima_movimentacao=ultima_mov,
        qtd_solicitacoes=len(solic), interesses="; ".join(interesses[:5]), volume_estimado=volume,
        ultima_mensagem=(msgs[0][:300] if msgs else ""), interacoes_outlook=outlook,
        locaweb_status=("Descadastrado" if optout else "Desativado (bounce)" if bounce else
                        "Ativo" if lw else ""),
        locaweb_aberturas=lw_views if lw else None, locaweb_cliques=lw_cliques if lw else None,
        email_ok=email_ok, whatsapp_ok=whats_ok, optout=optout, email_bounce=bounce,
        interno=interno, spam=spam, _nucleos_lista=nucleos_lista,
    )


# --------------------------------------------------------------------------------------
# Clusterização
# --------------------------------------------------------------------------------------

PAG_CONTAS = "Contas-Chave"
PAG_B2B = "B2B Corporativo"
PAG_CANAL = "Agências-Brindes-Gráficas"
PAG_LOJA = "Loja Física - Varejo"
PAG_ESCOLAR = "Uniforme Escolar"
PAG_QUALIF = "A Qualificar"
PAG_SUPR = "Supressão"

TEMP_QUENTE = "1. Quente (solicitou ≤90 dias)"
TEMP_REL = "2. Relacionamento ativo (Outlook ≥5 e-mails)"
TEMP_MORNO = "3. Morno"
TEMP_FRIO = "4. Frio / Reativação"
TEMP_SUPR = "5. Não contatar"


def clusterizar(df, ref):
    # --- Empresas e porte da conta
    tem_emp = df["empresa"] != ""
    ag = df[tem_emp].groupby("empresa").agg(contatos=("email", "size"), outlook=("interacoes_outlook", "sum"))
    def porte(emp):
        if not emp or emp not in ag.index:
            return ""
        a = ag.loc[emp]
        if a.contatos >= 5 or a.outlook >= 20 or emp in ALIAS_EMPRESAS.values():
            return "Conta-Chave"
        if a.contatos >= 2 or a.outlook >= 5:
            return "Conta Média"
        return "Conta Única"
    df["porte_conta"] = df["empresa"].map(porte)
    df["contatos_na_empresa"] = df["empresa"].map(lambda e: int(ag.loc[e].contatos) if e in ag.index else 0)

    # --- Núcleo (B2B x Loja física) e segmento
    def nucleo_segmento(r):
        orig = r["origens"].lower()
        if "uniforme escolar" in orig:
            return "B2B", "Uniforme Escolar (revenda/kits)"
        if "canal" in r["_nucleos_lista"]:
            return "B2B", "Canal indireto: " + (r["vertical"] or "Agências/Brindes")
        if "b2b" in r["_nucleos_lista"]:
            return "B2B", r["vertical"]
        if r["empresa"]:
            return "B2B", "Empresa final"
        if (r["volume_estimado"] or 0) >= 100:
            return "B2B", "Empresa final (volume ≥100)"
        if "loja" in r["_nucleos_lista"] or (r["tipo_email"] == "pessoal" and r["qtd_solicitacoes"] > 0) \
                or "copa brasil" in orig or "sustentabilidade" in orig:
            return "Loja Física", "Varejo / projetos menores"
        if r["tipo_email"] == "pessoal" and "Site" in r["fontes"]:
            return "Loja Física", "Varejo / projetos menores"
        return "A qualificar", "Sem indicação de empresa"
    ns = df.apply(nucleo_segmento, axis=1, result_type="expand")
    df["nucleo"], df["segmento"] = ns[0], ns[1]

    # --- Temperatura (data de entrada + indicadores de interação e propostas)
    def temperatura(r):
        if r["interno"] or r["spam"] or r["optout"] or not (r["email_ok"] or r["telefone"]):
            return TEMP_SUPR
        d = r["dias_desde_solicitacao"]
        if d is not None and not pd.isna(d) and d <= 90:
            return TEMP_QUENTE
        if r["interacoes_outlook"] >= 5:
            return TEMP_REL
        entrada_recente = pd.notna(r["data_entrada"]) and (ref - r["data_entrada"]).days <= 365 \
            and r["data_entrada_tipo"] != "entrada na planilha"
        if (d is not None and not pd.isna(d) and d <= 365) or r["interacoes_outlook"] >= 1 \
                or (r["locaweb_cliques"] or 0) > 0 or entrada_recente:
            return TEMP_MORNO
        return TEMP_FRIO
    df["temperatura"] = df.apply(temperatura, axis=1)

    def motivo_supressao(r):
        if r["temperatura"] != TEMP_SUPR:
            return ""
        if r["interno"]:
            return "Interno / teste / sistema"
        if r["spam"]:
            return "Spam de formulário"
        if r["optout"]:
            return "Descadastrado (opt-out) na Locaweb"
        if r["email_bounce"]:
            return "E-mail inválido (bounce) e sem telefone"
        return "Sem e-mail válido nem telefone"
    df["motivo_supressao"] = df.apply(motivo_supressao, axis=1)

    # --- Score de prioridade (0-100) para ordenar dentro de cada página
    def score(r):
        if r["temperatura"] == TEMP_SUPR:
            return 0
        s = 0
        d = r["dias_desde_solicitacao"]
        if d is not None and not pd.isna(d):
            s += 35 if d <= 30 else 28 if d <= 90 else 18 if d <= 180 else 10 if d <= 365 else 4
        s += min(15, 5 * r["qtd_solicitacoes"])
        s += min(20, round(6 * math.log2(1 + r["interacoes_outlook"])))
        s += 8 if (r["locaweb_cliques"] or 0) > 0 else 3 if (r["locaweb_aberturas"] or 0) > 0 else 0
        s += 7 if r["tipo_email"] == "corporativo" else 0
        s += {"Conta-Chave": 10, "Conta Média": 5}.get(r["porte_conta"], 0)
        v = r["volume_estimado"] or 0
        s += 5 if v >= 100 else 2 if v >= 30 else 0
        s += 3 if r["whatsapp_ok"] else 0
        s -= 5 if r["email_bounce"] and not r["whatsapp_ok"] else 0
        return max(0, min(100, s))
    df["score_prioridade"] = df.apply(score, axis=1)

    # --- Coorte pela data de entrada
    def coorte(d):
        if pd.isna(d):
            return ""
        return f"{d.year}-T{(d.month - 1) // 3 + 1}"
    df["coorte_entrada"] = df["data_entrada"].map(coorte)

    # --- Página (aba) de destino, mutuamente exclusiva
    def pagina(r):
        if r["temperatura"] == TEMP_SUPR:
            return PAG_SUPR
        if r["porte_conta"] == "Conta-Chave":
            return PAG_CONTAS
        if r["segmento"].startswith("Uniforme Escolar"):
            return PAG_ESCOLAR
        if r["segmento"].startswith("Canal indireto"):
            return PAG_CANAL
        if r["nucleo"] == "B2B":
            return PAG_B2B
        if r["nucleo"] == "Loja Física":
            return PAG_LOJA
        return PAG_QUALIF
    df["pagina"] = df.apply(pagina, axis=1)

    # --- Tags
    def slug(s):
        return re.sub(r"[^a-z0-9]+", "_", chave_texto(s)).strip("_")
    def tags(r):
        t = [f"nucleo:{slug(r['nucleo'])}", f"temp:{slug(r['temperatura'][3:].split('(')[0])}",
             f"pagina:{slug(r['pagina'])}"]
        if r["vertical"]:
            t += [f"vertical:{slug(v)}" for v in r["vertical"].split("; ")]
        if r["porte_conta"]:
            t.append(f"conta:{slug(r['porte_conta'])}")
        if r["coorte_entrada"]:
            t.append(f"entrada:{r['coorte_entrada']}")
        t += [f"fonte:{slug(f)}" for f in r["fontes"].split("; ")]
        t += [f"carteira:{slug(c)}" for c in r["carteira_vendedora"].split("; ") if c]
        t += [f"interesse:{slug(i)}" for i in r["interesses"].split("; ") if i][:3]
        if r["qtd_solicitacoes"]:
            t.append("pediu_orcamento")
        if r["interacoes_outlook"]:
            t.append("historico_outlook")
        if r["tipo_email"] == "corporativo":
            t.append("email_corporativo")
        t.append("canal:email" if r["email_ok"] else "sem_email_valido")
        if r["whatsapp_ok"]:
            t.append("canal:whatsapp")
        if r["ddd_assumido_11"]:
            t.append("ddd_assumido_11")
        if r["optout"]:
            t.append("optout")
        if r["email_bounce"]:
            t.append("email_bounce")
        return "; ".join(t)
    df["tags"] = df.apply(tags, axis=1)

    df = df.sort_values(["pagina", "score_prioridade", "empresa"], ascending=[True, False, True])
    df.insert(0, "id_contato", [f"CE{str(i + 1).zfill(6)}" for i in range(len(df))])
    return df.drop(columns=["_nucleos_lista"])


# --------------------------------------------------------------------------------------
# Saída
# --------------------------------------------------------------------------------------

COLUNAS = [
    ("id_contato", "ID", 10), ("nome", "Nome", 26), ("empresa", "Empresa", 24), ("porte_conta", "Porte da conta", 13),
    ("nucleo", "Núcleo", 12), ("segmento", "Segmento", 26), ("temperatura", "Temperatura", 24),
    ("score_prioridade", "Score", 7), ("email", "E-mail", 32), ("telefone", "Telefone", 16),
    ("whatsapp", "WhatsApp (disparo)", 15), ("uf", "UF", 5), ("coorte_entrada", "Coorte entrada", 10),
    ("data_entrada", "Data de entrada", 12), ("data_entrada_tipo", "Tipo da data", 14),
    ("ultima_solicitacao", "Última solicitação", 12), ("qtd_solicitacoes", "Nº solicitações", 9),
    ("interacoes_outlook", "E-mails trocados (Outlook)", 10), ("locaweb_status", "Status Locaweb", 14),
    ("locaweb_aberturas", "Aberturas", 8), ("locaweb_cliques", "Cliques", 7),
    ("interesses", "Interesses", 24), ("volume_estimado", "Volume estimado (peças)", 10),
    ("carteira_vendedora", "Carteira", 14), ("vertical", "Vertical", 22), ("fontes", "Fontes", 30),
    ("origens", "Origens (arquivo/formulário/lista)", 36), ("tags", "Tags", 60),
    ("tipo_email", "Tipo e-mail", 11), ("emails_adicionais", "Outros e-mails", 26),
    ("telefones_adicionais", "Outros telefones", 20), ("tipo_telefone", "Tipo tel.", 9),
    ("ddd_assumido_11", "DDD 11 assumido", 8), ("dominio", "Domínio", 20),
    ("empresa_origem", "Origem da empresa", 18), ("cidade", "Cidade", 12),
    ("ultima_mensagem", "Última mensagem", 40), ("motivo_supressao", "Motivo supressão", 20),
    ("qtd_fontes", "Nº fontes", 7), ("qtd_registros_brutos", "Nº registros unidos", 8),
]
ORDEM_PAGINAS = [PAG_CONTAS, PAG_B2B, PAG_CANAL, PAG_ESCOLAR, PAG_LOJA, PAG_QUALIF, PAG_SUPR]
DESC_PAGINAS = {
    PAG_CONTAS: "Contatos de empresas com 5+ contatos na base, 20+ e-mails trocados ou marcas estratégicas (Carrefour, Safra, Vigor, Atacadão, Nestlé...). Ordenado por empresa.",
    PAG_B2B: "Empresas finais (e-mail corporativo, empresa informada ou listas verticais: alimentícias, atacadistas, franquias, condomínios, buffets, escolas de idiomas).",
    PAG_CANAL: "Canal indireto: agências de marketing/publicidade (inclui ABAP), promocionais/brindes (Brasil Promotions) e gráficas de impressão digital.",
    PAG_ESCOLAR: "Leads dos formulários de Uniforme Escolar (atacado, varejo e kit municipal) — revenda e kits.",
    PAG_LOJA: "Núcleo loja física: pessoa física / projetos menores (e-mail pessoal sem empresa que pediu orçamento, lista Store, Copa, Sustentabilidade).",
    PAG_QUALIF: "Contatos sem sinal de empresa ou intenção (principalmente agendas de WhatsApp e históricos sem nome de empresa). Qualificar antes de disparar.",
    PAG_SUPR: "NÃO disparar: opt-out na Locaweb, internos/testes, spam de formulário, ou sem e-mail/telefone válido.",
}


def fmt_data(v):
    if v is None or (isinstance(v, float) and math.isnan(v)) or pd.isna(v):
        return None
    return pd.Timestamp(v).to_pydatetime().replace(microsecond=0)


def escrever_excel(df, empresas, fontes_df, resumo_brutos, ref, caminho):
    with pd.ExcelWriter(caminho, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as xw:
        wb = xw.book
        h = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "text_wrap": True,
                           "valign": "top", "border": 1})
        titulo = wb.add_format({"bold": True, "font_size": 16})
        sub = wb.add_format({"bold": True, "font_size": 12, "font_color": "#1F3864"})
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        num = wb.add_format({"num_format": "#,##0"})

        # --- Resumo
        ws = wb.add_worksheet("Resumo")
        ws.set_column(0, 0, 44)
        ws.set_column(1, 9, 14)
        ws.write(0, 0, "Camiseta Express — Base de Leads Clusterizada", titulo)
        ws.write(1, 0, f"Data de referência: {ref:%d/%m/%Y}. Registros brutos lidos: {resumo_brutos['brutos']:,}. "
                       f"Contatos únicos após limpeza e deduplicação: {len(df):,}.".replace(",", "."))
        linha = 3
        def tabela(titulo_t, tab, linha):
            ws.write(linha, 0, titulo_t, sub)
            linha += 1
            tab = tab.reset_index()
            for j, c in enumerate(tab.columns):
                ws.write(linha, j, str(c), h)
            for i, row in enumerate(tab.itertuples(index=False)):
                for j, v in enumerate(row):
                    ws.write(linha + 1 + i, j, v if not isinstance(v, float) or not math.isnan(v) else "",
                             num if isinstance(v, (int, float)) else None)
            return linha + len(tab) + 3

        pag = df.groupby("pagina").agg(contatos=("id_contato", "size"),
                                       com_email_ok=("email_ok", "sum"), com_whatsapp=("whatsapp_ok", "sum"))
        pag = pag.reindex([p for p in ORDEM_PAGINAS if p in pag.index])
        pag["descrição"] = [DESC_PAGINAS[p] for p in pag.index]
        linha = tabela("Contatos por página (aba)", pag, linha)
        linha = tabela("Página × Temperatura", pd.crosstab(df["pagina"], df["temperatura"], margins=True,
                                                            margins_name="Total"), linha)
        linha = tabela("Núcleo × Temperatura", pd.crosstab(df["nucleo"], df["temperatura"], margins=True,
                                                            margins_name="Total"), linha)
        coorte = df[df["coorte_entrada"] != ""]
        linha = tabela("Coorte de entrada × Temperatura (data em que o contato entrou nas planilhas)",
                       pd.crosstab(coorte["coorte_entrada"], coorte["temperatura"], margins=True,
                                   margins_name="Total"), linha)
        seg = df[df["temperatura"] != TEMP_SUPR].groupby(["nucleo", "segmento"]).size().rename("contatos")
        linha = tabela("Segmentos (sem supressão)", seg.to_frame(), linha)
        linha = tabela("Top 30 empresas (por score)", empresas.head(30).set_index("Empresa")[
            ["Contatos", "Porte", "E-mails trocados (Outlook)", "Solicitações", "Temperatura (melhor)"]], linha)
        fontes = df["fontes"].str.split("; ").explode().value_counts().rename("contatos com esta fonte")
        linha = tabela("Contatos por fonte (um contato pode ter várias)", fontes.to_frame(), linha)
        linha = tabela("Arquivos lidos", fontes_df.set_index("arquivo"), linha)

        # --- Base e páginas
        cols = [c for c, _, _ in COLUNAS]
        nomes = {c: n for c, n, _ in COLUNAS}
        def aba(nome, dados, desc=None):
            out = dados[cols].rename(columns=nomes).copy()
            for c in ("Data de entrada", "Última solicitação"):
                out[c] = out[c].map(fmt_data)
            for c in ("Volume estimado (peças)", "Aberturas", "Cliques"):
                out[c] = out[c].astype("Int64")
            out["DDD 11 assumido"] = out["DDD 11 assumido"].map({True: "sim", False: ""})
            inicio = 2 if desc else 0
            out.to_excel(xw, sheet_name=nome, index=False, startrow=inicio)
            w = xw.sheets[nome]
            if desc:
                w.write(0, 0, desc, wb.add_format({"italic": True, "font_color": "#555555"}))
            for j, (_, n, larg) in enumerate(COLUNAS):
                w.write(inicio, j, n, h)
                w.set_column(j, j, larg)
            w.freeze_panes(inicio + 1, 2)
            w.autofilter(inicio, 0, inicio + len(out), len(COLUNAS) - 1)

        aba("Base Consolidada", df, "Todos os contatos únicos. Use os filtros e a coluna Tags para montar listas de disparo.")
        for p in ORDEM_PAGINAS:
            sub_df = df[df["pagina"] == p]
            if p == PAG_CONTAS:
                sub_df = sub_df.sort_values(["empresa", "score_prioridade"], ascending=[True, False])
            aba(p, sub_df, DESC_PAGINAS[p])

        # --- Empresas
        empresas.to_excel(xw, sheet_name="Empresas", index=False, startrow=2)
        w = xw.sheets["Empresas"]
        w.write(0, 0, "Uma linha por empresa identificada (domínio corporativo, campo Empresa dos formulários ou marca citada no nome do contato).",
                wb.add_format({"italic": True, "font_color": "#555555"}))
        for j, c in enumerate(empresas.columns):
            w.write(2, j, c, h)
        w.set_column(0, 0, 30)
        w.set_column(1, len(empresas.columns), 16)
        w.set_column(len(empresas.columns) - 2, len(empresas.columns) - 1, 50)
        w.freeze_panes(3, 1)
        w.autofilter(2, 0, 2 + len(empresas), len(empresas.columns) - 1)

        # --- Dicionário
        w = wb.add_worksheet("Dicionário")
        w.set_column(0, 0, 30)
        w.set_column(1, 1, 110, wrap)
        itens = [
            ("REGRAS DE LIMPEZA", ""),
            ("Telefone", "Somente dígitos → remove +55/0/código de operadora → sem DDD recebe 11 (operação SP) → celular com 8 dígitos ganha o 9. Salvo como (DD) 9XXXX-XXXX; coluna WhatsApp traz 55DDDNÚMERO pronto para disparo."),
            ("E-mail", "Minúsculo, validado, com correção de erros comuns (gmail.con, hotmial.com...)."),
            ("Deduplicação", "Registros de todas as fontes são unidos quando compartilham e-mail OU telefone (telefones usados por 4+ e-mails diferentes não unem, para não fundir pessoas distintas)."),
            ("Empresa", "1) domínio do e-mail corporativo; 2) campo Empresa do formulário; 3) marca citada no nome do contato da agenda (ex.: 'Paulo Carrefour')."),
            ("Data de entrada", "Data mais antiga em que o contato aparece: envio do formulário, cadastro na Locaweb ou, para agendas/históricos sem data, a data da planilha."),
            ("TEMPERATURA", ""),
            (TEMP_QUENTE, "Enviou formulário de orçamento/WhatsApp nos últimos 90 dias → proposta em aberto, prioridade de follow-up."),
            (TEMP_REL, "5+ e-mails trocados com a equipe (histórico Outlook) → cliente/negociação recorrente."),
            (TEMP_MORNO, "Solicitou entre 91 e 365 dias, ou tem 1–4 e-mails trocados, ou clicou em e-mail marketing, ou entrou na base há menos de 1 ano."),
            (TEMP_FRIO, "Sem sinal nos últimos 12 meses → campanhas de reativação."),
            (TEMP_SUPR, "Opt-out, interno/teste, spam ou sem canal válido."),
            ("PORTE DA CONTA", ""),
            ("Conta-Chave", "Empresa com 5+ contatos, 20+ e-mails trocados, ou marca estratégica."),
            ("Conta Média", "2–4 contatos ou 5+ e-mails trocados."),
            ("Conta Única", "1 contato."),
            ("SCORE", "0–100: recência da solicitação (até 35) + nº de solicitações (15) + e-mails trocados (20) + engajamento Locaweb (8) + e-mail corporativo (7) + porte (10) + volume (5) + WhatsApp (3)."),
            ("TAGS", "Separadas por '; ' no formato grupo:valor — nucleo:, temp:, pagina:, vertical:, conta:, entrada:AAAA-Tn, fonte:, carteira:, interesse:, e marcadores pediu_orcamento, historico_outlook, email_corporativo, canal:email, canal:whatsapp, ddd_assumido_11, optout, email_bounce."),
            ("ARQUIVOS DE DISPARO", "A pasta disparos/ traz, por página, um CSV de e-mail (só e-mails aptos) e um de WhatsApp (só celulares), sem a aba Supressão."),
        ]
        for i, (a, b) in enumerate(itens):
            w.write(i, 0, a, sub if not b else wrap)
            w.write(i, 1, b, wrap)


def tabela_empresas(df):
    ativos = df[(df["empresa"] != "") & ~df["motivo_supressao"].isin(["Interno / teste / sistema", "Spam de formulário"])]
    ordem_temp = {t: i for i, t in enumerate([TEMP_QUENTE, TEMP_REL, TEMP_MORNO, TEMP_FRIO, TEMP_SUPR])}
    g = ativos.groupby("empresa")
    emp = pd.DataFrame({
        "Empresa": g.size().index,
        "Contatos": g.size().values,
        "Porte": g["porte_conta"].first().values,
        "Núcleo": g["nucleo"].agg(lambda s: s.mode().iat[0]).values,
        "Segmento": g["segmento"].agg(lambda s: s.mode().iat[0]).values,
        "Temperatura (melhor)": g["temperatura"].agg(lambda s: min(s, key=lambda t: ordem_temp[t])).values,
        "Score máx.": g["score_prioridade"].max().values,
        "E-mails trocados (Outlook)": g["interacoes_outlook"].sum().values,
        "Solicitações": g["qtd_solicitacoes"].sum().values,
        "Última solicitação": g["ultima_solicitacao"].max().map(fmt_data).values,
        "Primeira entrada": g["data_entrada"].min().map(fmt_data).values,
        "Com e-mail apto": g["email_ok"].sum().values,
        "Com WhatsApp": g["whatsapp_ok"].sum().values,
        "Domínio": g["dominio"].agg(lambda s: next((x for x in s if x), "")).values,
        "Carteira": g["carteira_vendedora"].agg(
            lambda s: "; ".join(sorted({x for v in s for x in v.split("; ") if x}))).values,
        "Contatos (nome <e-mail>)": g.apply(lambda x: "; ".join(
            f"{n or '?'} <{e or t}>" for n, e, t in zip(x["nome"], x["email"], x["telefone"]))[:1000]).values,
    })
    return emp.sort_values(["Score máx.", "Contatos"], ascending=False).reset_index(drop=True)


def escrever_disparos(df, pasta):
    pasta = Path(pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    arquivos = []
    for p in ORDEM_PAGINAS:
        if p == PAG_SUPR:
            continue
        sub = df[df["pagina"] == p]
        slug = re.sub(r"[^a-z0-9]+", "_", chave_texto(p)).strip("_")
        em = sub[sub["email_ok"]][["nome", "email", "empresa", "temperatura", "segmento", "tags"]]
        wa = sub[sub["whatsapp_ok"]][["nome", "whatsapp", "empresa", "temperatura", "segmento", "tags"]]
        for tipo, d in (("email", em), ("whatsapp", wa)):
            f = pasta / f"{slug}__{tipo}.csv"
            d.to_csv(f, index=False, encoding="utf-8-sig", sep=";")
            arquivos.append(f)
    return arquivos


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--entrada", required=True)
    ap.add_argument("--saida", required=True)
    ap.add_argument("--data-referencia", default=dt.date.today().isoformat())
    a = ap.parse_args()
    ref = dt.datetime.fromisoformat(a.data_referencia)
    saida = Path(a.saida)
    saida.mkdir(parents=True, exist_ok=True)

    print("Lendo planilhas...")
    registros, interacoes, dono_interacao, fontes_df = extrair_registros(a.entrada)
    print(f"  {len(registros):,} registros brutos de {len(fontes_df)} arquivos/abas")
    print("Consolidando e deduplicando...")
    df = consolidar(registros, interacoes, dono_interacao, ref)
    print(f"  {len(df):,} contatos únicos")
    print("Clusterizando...")
    df = clusterizar(df, ref)
    empresas = tabela_empresas(df)

    xlsx = saida / "Camiseta_Express_Leads_Clusterizados.xlsx"
    escrever_excel(df, empresas, fontes_df, {"brutos": len(registros)}, ref, xlsx)
    arqs = escrever_disparos(df, saida / "disparos")
    with zipfile.ZipFile(saida / "disparos.zip", "w", zipfile.ZIP_DEFLATED) as z:
        for f in arqs:
            z.write(f, f"disparos/{f.name}")
    print(f"OK → {xlsx}")
    print(df.groupby(["pagina", "temperatura"]).size().unstack(fill_value=0).to_string())


if __name__ == "__main__":
    main()
