# -*- coding: utf-8 -*-
"""
RADAR DA TRANSPARÊNCIA MUNICIPAL — Fase 1
=========================================
Cruza contratos públicos (PNCP ou planilha do portal municipal) com:
  - dados de CNPJ dos fornecedores (data de abertura, sócios, porte, situação)
  - eleitos e doadores de campanha (TSE, eleições municipais de 2024)
  - uma lista opcional de servidores/agentes públicos informada por você

e gera ALERTAS de risco. Um alerta é um INDÍCIO para investigar,
nunca uma prova de irregularidade.

Saídas: relatorio_radar.html (painel) e radar_dados.xlsx (todas as tabelas).
"""

import io, json, os, re, time, unicodedata, zipfile
from collections import defaultdict
from datetime import date, datetime

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO PADRÃO (o notebook sobrescreve o que você mudar)
# ---------------------------------------------------------------------------
CONFIG = {
    "municipio": "Governador Newton Bello",
    "uf": "MA",
    "codigo_ibge": "2104651",
    "anos": [2023, 2024, 2025, 2026],
    # CNPJs de órgãos municipais já conhecidos (prefeitura, fundos de saúde,
    # educação, assistência). O radar também tenta descobrir sozinho.
    "cnpjs_orgaos_extra": [],
    # Planilha de contratos/pagamentos baixada do portal da prefeitura ou do TCE
    # (opcional). Colunas aceitas: cnpj/cpf, fornecedor, valor, data, objeto, modalidade.
    "arquivo_contratos_csv": None,
    # Planilha opcional de agentes públicos: colunas nome;cargo
    "arquivo_agentes_csv": None,
    "usar_pncp": True,
    # Portal da transparência da prefeitura (lista pública de contratos)
    "usar_portal": True,
    "portal_prefeitura": "https://www.governadornewtonbello.ma.gov.br",
    "usar_tse": False,
    "pasta_saida": ".",
}

PNCP = "https://pncp.gov.br/api/consulta"
CNPJ_APIS = ["https://brasilapi.com.br/api/cnpj/v1/{}", "https://minhareceita.org/{}"]
TSE_CAND = "https://cdn.tse.jus.br/estatistica/sead/odsele/consulta_cand/consulta_cand_2024.zip"
TSE_RECEITAS = ("https://cdn.tse.jus.br/estatistica/sead/odsele/prestacao_contas/"
                "prestacao_de_contas_eleitorais_candidatos_2024.zip")

MODALIDADES = {1: "Leilão eletrônico", 2: "Diálogo competitivo", 3: "Concurso",
               4: "Concorrência eletrônica", 5: "Concorrência presencial",
               6: "Pregão eletrônico", 7: "Pregão presencial", 8: "Dispensa",
               9: "Inexigibilidade", 10: "Manifestação de interesse",
               11: "Pré-qualificação", 12: "Credenciamento", 13: "Leilão presencial"}
SEM_LICITACAO = {"Dispensa", "Inexigibilidade"}

PESO = {"ALTO": 3, "MÉDIO": 2, "BAIXO": 1}

# Sobrenomes muito comuns no Brasil: coincidência neles não indica parentesco.
SOBRENOMES_COMUNS = set("""
SILVA SANTOS SOUSA SOUZA OLIVEIRA PEREIRA LIMA COSTA FERREIRA RODRIGUES ALVES
NASCIMENTO ARAUJO CARVALHO GOMES MARTINS RIBEIRO ROCHA BARBOSA MELO MELLO CASTRO
CARDOSO SANTANA MOURA MENDES NUNES MOREIRA VIEIRA LOPES MARQUES FREITAS CORREIA
DIAS TEIXEIRA CAVALCANTE BEZERRA FERNANDES MACHADO MONTEIRO PINTO REIS JESUS
BRITO MIRANDA ANDRADE BATISTA RAMOS CRUZ FARIAS FARIA AZEVEDO GONCALVES MOTA
MATOS PAIVA CAMPOS LEITE CUNHA SALES DUARTE SOARES XAVIER AGUIAR VIANA BARROS
FILHO FILHA JUNIOR NETO NETA SOBRINHO
""".split())
PARTICULAS = {"DA", "DE", "DO", "DAS", "DOS", "E", "D"}


# ---------------------------------------------------------------------------
# UTILITÁRIOS
# ---------------------------------------------------------------------------
LOG_FN = None  # o app web troca por uma função própria
ULTIMO_RESULTADO = None


def log(msg):
    if LOG_FN:
        LOG_FN(msg)
    else:
        print(msg, flush=True)


def norm(txt):
    """Maiúsculas, sem acento, só letras e espaços."""
    if txt is None or (isinstance(txt, float) and pd.isna(txt)):
        return ""
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    txt = re.sub(r"[^A-Za-z ]", " ", txt).upper()
    return re.sub(r"\s+", " ", txt).strip()


def so_digitos(txt):
    return re.sub(r"\D", "", str(txt or ""))


def norm_end(info):
    """Endereço normalizado (com número) para comparar fornecedores; vazio se sem número."""
    numero = str(info.get("numero") or "").strip().upper()
    if numero in ("", "SN", "S/N", "S N", "0"):
        return ""
    txt = unicodedata.normalize("NFKD", f"{info.get('logradouro', '')} {numero} {info.get('municipio', '')}")
    txt = re.sub(r"[^A-Z0-9 ]", " ", txt.encode("ascii", "ignore").decode().upper())
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt if len(txt) > 8 else ""


def sobrenomes_raros(nome):
    partes = [p for p in norm(nome).split() if p not in PARTICULAS]
    return {p for p in partes[1:] if p not in SOBRENOMES_COMUNS and len(p) > 2}


def num(v):
    """Converte '1.234,56' ou 1234.56 em float."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("R$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def data(v):
    if not v or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v)[:10]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y%m%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def cnpj_fmt(c):
    original, c = str(c or ""), so_digitos(c)
    if "*" in original:
        return original  # já mascarado pela fonte
    if len(c) == 14:
        return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
    if len(c) == 11:
        return f"***.{c[3:6]}.{c[6:9]}-**"  # CPF mascarado (LGPD)
    return c


SESSION = requests.Session()
SESSION.headers["User-Agent"] = "radar-transparencia/1.0 (controle social)"


def http_get(url, params=None, tentativas=4):
    """GET com novas tentativas. Devolve dict/list, None (vazio) ou levanta erro."""
    for i in range(tentativas):
        try:
            r = SESSION.get(url, params=params, timeout=60)
            if r.status_code == 204:
                return None
            if r.status_code == 404:
                return None
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(2 * (i + 1))
                continue
            r.raise_for_status()
            return r.json() if r.content else None
        except requests.RequestException:
            if i == tentativas - 1:
                raise
            time.sleep(2 * (i + 1))
    return None


def g(d, *chaves, padrao=None):
    """Pega o primeiro campo existente (o PNCP varia a grafia de alguns campos)."""
    for k in chaves:
        cur = d
        ok = True
        for parte in k.split("."):
            if isinstance(cur, dict) and parte in cur:
                cur = cur[parte]
            else:
                ok = False
                break
        if ok and cur not in (None, ""):
            return cur
    return padrao


# ---------------------------------------------------------------------------
# 1. COLETA DE CONTRATOS
# ---------------------------------------------------------------------------
def intervalo(ano):
    fim = min(date(ano, 12, 31), date.today())
    return f"{ano}0101", fim.strftime("%Y%m%d")


def paginar(endpoint, params, tamanho=50):
    pagina, itens = 1, []
    while True:
        resp = http_get(f"{PNCP}{endpoint}", {**params, "pagina": pagina, "tamanhoPagina": tamanho})
        if not resp or not resp.get("data"):
            break
        itens.extend(resp["data"])
        if resp.get("paginasRestantes", 0) <= 0:
            break
        pagina += 1
        time.sleep(0.3)
    return itens


def coletar_pncp(cfg):
    ibge, anos = cfg["codigo_ibge"], [a for a in cfg["anos"] if a <= date.today().year]
    compras, orgaos = {}, {so_digitos(c): "" for c in cfg["cnpjs_orgaos_extra"]}

    log("PNCP · buscando licitações e dispensas do município…")
    for ano in anos:
        ini, fim = intervalo(ano)
        for cod, nome in MODALIDADES.items():
            try:
                itens = paginar("/v1/contratacoes/publicacao",
                                {"dataInicial": ini, "dataFinal": fim,
                                 "codigoModalidadeContratacao": cod,
                                 "codigoMunicipioIbge": ibge})
            except Exception as e:
                log(f"  aviso: {ano}/{nome}: {e}")
                continue
            for c in itens:
                compras[g(c, "numeroControlePNCP")] = {
                    "modalidade": nome, "objeto_compra": g(c, "objetoCompra", padrao="")}
                cnpj_org = so_digitos(g(c, "orgaoEntidade.cnpj"))
                if cnpj_org:
                    orgaos[cnpj_org] = g(c, "orgaoEntidade.razaoSocial", "orgaoEntidade.razaosocial", padrao="")
            if itens:
                log(f"  {ano} · {nome}: {len(itens)}")
    log(f"  órgãos encontrados: {len(orgaos)}")

    linhas = []
    log("PNCP · buscando contratos de cada órgão…")
    for cnpj_org, nome_org in orgaos.items():
        for ano in anos:
            ini, fim = intervalo(ano)
            try:
                itens = paginar("/v1/contratos", {"dataInicial": ini, "dataFinal": fim,
                                                  "cnpjOrgao": cnpj_org}, tamanho=100)
            except Exception as e:
                log(f"  aviso: contratos {cnpj_org}/{ano}: {e}")
                continue
            for c in itens:
                ctrl_compra = g(c, "numeroControlePNCPCompra", "numeroControlePncpCompra")
                compra = compras.get(ctrl_compra, {})
                ano_c, seq = g(c, "anoContrato"), g(c, "sequencialContrato")
                linhas.append({
                    "fonte": "PNCP",
                    "orgao": g(c, "orgaoEntidade.razaoSocial", padrao=nome_org),
                    "cnpj_orgao": cnpj_org,
                    "doc_fornecedor": so_digitos(g(c, "niFornecedor")),
                    "fornecedor": g(c, "nomeRazaoSocialFornecedor", padrao=""),
                    "valor": num(g(c, "valorGlobal", "valorInicial", padrao=0)),
                    "data": data(g(c, "dataAssinatura", "dataVigenciaInicio", "dataPublicacaoPncp")),
                    "objeto": g(c, "objetoContrato", padrao=compra.get("objeto_compra", "")),
                    "modalidade": compra.get("modalidade", "Não informada"),
                    "link": (f"https://pncp.gov.br/app/contratos/{cnpj_org}/{ano_c}/{seq}"
                             if ano_c and seq else ""),
                })
            if itens:
                log(f"  {nome_org or cnpj_org} · {ano}: {len(itens)} contratos")
    return pd.DataFrame(linhas)


# ---------------------------------------------------------------------------
# 1b. PORTAL DA PREFEITURA (páginas públicas de contratos)
# ---------------------------------------------------------------------------
RE_CNPJ = re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")
RE_CPF = re.compile(r"[\d\*]{3}\.?[\d\*]{3}\.?[\d\*]{3}-[\d\*]{2}")
RE_DATA = re.compile(r"\d{2}/\d{2}/\d{4}")
RE_VALOR = re.compile(r"R\$\s*([\d\.]+,\d{2})")
RE_NUM = re.compile(r"\b(\d{1,4}/\d{4})\b")
TIPOS_CONTRATO = ("CONTRATO ORIGINAL", "ADITIVO", "TERMO DE RETIFICA", "APOSTILAMENTO")


def _linhas(celula):
    return [t.strip() for t in celula.get_text("\n").split("\n") if t.strip()]


def _parse_linha(tr, base):
    tds = tr.find_all("td")
    if len(tds) < 4:
        return None
    texto = " | ".join(" ".join(_linhas(td)) for td in tds)
    doc_m = RE_CNPJ.search(texto) or RE_CPF.search(texto)
    val_m = RE_VALOR.search(texto)
    if not doc_m or not val_m:
        return None
    link = ""
    for a in tr.find_all("a", href=True):
        if "contratos.php?id=" in a["href"]:
            link = requests.compat.urljoin(base + "/", a["href"])
            break
    # célula do contratado = a que contém o documento; a seguinte traz secretaria + objeto
    forn, objeto, secretaria, tipo, numero = "", "", "", "", ""
    idx_doc = None
    for i, td in enumerate(tds):
        ls = _linhas(td)
        junto = " ".join(ls)
        if idx_doc is None and doc_m.group(0) in junto:
            idx_doc = i
            forn = " ".join(l for l in ls if doc_m.group(0) not in l)
            forn = re.sub(r"\b(CNPJ|CPF)\s*:?\s*$", "", forn, flags=re.I).strip(" -:")
        elif not tipo and any(t in junto.upper() for t in TIPOS_CONTRATO):
            tipo = next((l for l in ls if any(t in l.upper() for t in TIPOS_CONTRATO)), "")
            n = RE_NUM.search(junto)
            numero = n.group(1) if n else ""
    if idx_doc is not None and idx_doc + 1 < len(tds):
        ls = [l for l in _linhas(tds[idx_doc + 1]) if not l.lower().startswith("veja mais")]
        if ls:
            secretaria, objeto = ls[0], " ".join(ls[1:]) or ls[0]
    # data do contrato = a data na mesma célula do valor
    data_c = None
    for td in tds:
        junto = " ".join(_linhas(td))
        if RE_VALOR.search(junto):
            d = RE_DATA.search(junto)
            data_c = data(d.group(0)) if d else None
            break
    doc = doc_m.group(0)
    return {
        "fonte": "Portal da prefeitura",
        "orgao": secretaria,
        "cnpj_orgao": "",
        "doc_fornecedor": so_digitos(doc) if RE_CNPJ.fullmatch(doc) else doc,
        "fornecedor": forn,
        "valor": num(val_m.group(1)),
        "data": data_c,
        "objeto": objeto,
        "modalidade": "Não informada",
        "tipo_registro": tipo or "Não informado",
        "numero": numero,
        "link": link,
    }


def coletar_portal(cfg, max_paginas=200):
    from bs4 import BeautifulSoup
    base = cfg["portal_prefeitura"].rstrip("/")
    url = f"{base}/contratos.php"
    log(f"Portal · lendo a lista pública de contratos em {url}")
    vistos, linhas, sem_novos = set(), [], 0
    for pagina in range(0, max_paginas):
        try:
            r = SESSION.get(url, params={"pagina": pagina} if pagina else None, timeout=60)
            r.raise_for_status()
        except requests.RequestException as e:
            log(f"  aviso: página {pagina}: {e}")
            break
        soup = BeautifulSoup(r.text, "html.parser")
        novos = 0
        for tr in soup.find_all("tr"):
            item = _parse_linha(tr, base)
            if not item:
                continue
            chave = item["link"] or (item["doc_fornecedor"], item["valor"], item["data"], item["numero"])
            if chave in vistos:
                continue
            vistos.add(chave)
            linhas.append(item)
            novos += 1
        if pagina == 0 and novos == 0:
            caminho = os.path.join(cfg["pasta_saida"], "portal_pagina0.html")
            with open(caminho, "w", encoding="utf-8") as fh:
                fh.write(r.text)
            log(f"  aviso: não reconheci a tabela de contratos. Página salva em {caminho} para ajuste.")
            break
        if novos == 0:
            sem_novos += 1
            if sem_novos >= 2:
                break
        else:
            sem_novos = 0
            log(f"  página {pagina + 1}: {novos} contratos")
        time.sleep(1.0)  # pausa para não sobrecarregar o site da prefeitura
    df = pd.DataFrame(linhas)
    if len(df):
        aditivos = df["tipo_registro"].str.upper().str.contains("ADITIVO", na=False).sum()
        log(f"  total no portal: {len(df)} registros ({aditivos} aditivos)")
    return df


COLS_ALTERNATIVAS = {
    "doc_fornecedor": ["cnpj", "cpf_cnpj", "cnpj_cpf", "cnpj/cpf", "cpf/cnpj", "documento", "cnpj_fornecedor", "ni_fornecedor"],
    "fornecedor": ["fornecedor", "credor", "razao_social", "nome", "favorecido", "contratado"],
    "valor": ["valor", "valor_pago", "valor_total", "valor_contrato", "valor_global", "pago", "empenhado"],
    "data": ["data", "data_assinatura", "data_pagamento", "data_empenho", "dt"],
    "objeto": ["objeto", "descricao", "historico", "especificacao"],
    "modalidade": ["modalidade", "modalidade_licitacao", "tipo_licitacao"],
    "orgao": ["orgao", "unidade", "secretaria", "unidade_gestora"],
}


def carregar_planilha_contratos(caminho):
    log(f"Planilha · lendo {caminho}")
    if caminho.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(caminho, dtype=str)
    else:
        df = None
        for enc in ("utf-8", "latin1"):
            try:
                df = pd.read_csv(caminho, sep=None, engine="python", dtype=str, encoding=enc)
                break
            except UnicodeDecodeError:
                continue
    cols = {norm(c).lower().replace(" ", "_"): c for c in df.columns}
    out = pd.DataFrame()
    for alvo, opcoes in COLS_ALTERNATIVAS.items():
        achou = next((cols[o] for o in opcoes if o in cols), None)
        out[alvo] = df[achou] if achou else ""
    out["doc_fornecedor"] = out["doc_fornecedor"].map(so_digitos)
    out["valor"] = out["valor"].map(num)
    out["data"] = out["data"].map(data)
    out["modalidade"] = out["modalidade"].map(
        lambda m: "Dispensa" if "DISPENSA" in norm(m) else
                  "Inexigibilidade" if "INEXIG" in norm(m) else (m or "Não informada"))
    out["fonte"], out["cnpj_orgao"], out["link"] = "Planilha", "", ""
    return out


# ---------------------------------------------------------------------------
# 2. DADOS DE CNPJ
# ---------------------------------------------------------------------------
def consultar_cnpjs(cnpjs, pasta):
    cache_path = os.path.join(pasta, "cache_cnpj.json")
    cache = json.load(open(cache_path)) if os.path.exists(cache_path) else {}
    faltam = [c for c in cnpjs if c not in cache]
    log(f"CNPJ · {len(cnpjs)} fornecedores PJ ({len(faltam)} para consultar)")
    for i, c in enumerate(faltam, 1):
        info = None
        for api in CNPJ_APIS:
            try:
                info = http_get(api.format(c), tentativas=3)
                if info:
                    break
            except Exception:
                continue
        cache[c] = info or {}
        if i % 10 == 0:
            log(f"  {i}/{len(faltam)}")
            json.dump(cache, open(cache_path, "w"))
        time.sleep(0.6)
    json.dump(cache, open(cache_path, "w"))
    return cache


# ---------------------------------------------------------------------------
# 3. TSE — ELEITOS E DOADORES DE 2024
# ---------------------------------------------------------------------------
def ler_tse(url, padrao_arquivo, municipio, pasta):
    destino = os.path.join(pasta, os.path.basename(url))
    if not os.path.exists(destino):
        log(f"TSE · baixando {os.path.basename(url)} (pode demorar alguns minutos)…")
        with SESSION.get(url, stream=True, timeout=600) as r:
            r.raise_for_status()
            with open(destino, "wb") as f:
                for parte in r.iter_content(1 << 20):
                    f.write(parte)
    alvo = norm(municipio)
    with zipfile.ZipFile(destino) as z:
        nome = next(n for n in z.namelist() if re.search(padrao_arquivo, n, re.I))
        pedacos = []
        with z.open(nome) as f:
            for bloco in pd.read_csv(f, sep=";", encoding="latin1", dtype=str, chunksize=200_000):
                if "NM_UE" in bloco.columns:
                    bloco = bloco[bloco["NM_UE"].map(norm) == alvo]
                pedacos.append(bloco)
    return pd.concat(pedacos) if pedacos else pd.DataFrame()


def carregar_agentes(cfg):
    agentes = []
    if cfg["usar_tse"]:
        try:
            cand = ler_tse(TSE_CAND, rf"consulta_cand_2024_{cfg['uf']}\.csv", cfg["municipio"], cfg["pasta_saida"])
            situ = cand.get("DS_SIT_TOT_TURNO", pd.Series("", index=cand.index)).map(norm)
            eleitos = cand[situ.str.contains("ELEITO") & ~situ.str.contains("NAO ELEITO")]
            for _, r in eleitos.iterrows():
                agentes.append({"nome": r.get("NM_CANDIDATO", ""), "cargo": f"{r.get('DS_CARGO', '')} (eleito 2024)"})
            log(f"TSE · {len(eleitos)} eleitos em 2024")
        except Exception as e:
            log(f"  aviso: não consegui ler candidatos do TSE ({e})")
    if cfg["arquivo_agentes_csv"]:
        df = pd.read_csv(cfg["arquivo_agentes_csv"], sep=None, engine="python", dtype=str)
        df.columns = [norm(c).lower() for c in df.columns]
        for _, r in df.iterrows():
            agentes.append({"nome": r.get("nome", ""), "cargo": r.get("cargo", "Servidor/agente")})
        log(f"Agentes · {len(df)} nomes da sua planilha")
    return pd.DataFrame(agentes, columns=["nome", "cargo"])


def carregar_doacoes(cfg):
    if not cfg["usar_tse"]:
        return pd.DataFrame()
    try:
        rec = ler_tse(TSE_RECEITAS, rf"receitas_candidatos_2024_{cfg['uf']}\.csv", cfg["municipio"], cfg["pasta_saida"])
    except Exception as e:
        log(f"  aviso: não consegui ler doações do TSE ({e})")
        return pd.DataFrame()
    def col(*nomes):
        for n in nomes:
            if n in rec.columns:
                return rec[n].fillna("")
        return pd.Series("", index=rec.index)
    out = pd.DataFrame({
        "doc_doador": col("NR_CPF_CNPJ_DOADOR").map(so_digitos),
        "doador": col("NM_DOADOR", "NM_DOADOR_RFB"),
        "candidato": col("NM_CANDIDATO"),
        "cargo": col("DS_CARGO"),
        "valor": col("VR_RECEITA").map(num),
    })
    log(f"TSE · {len(out)} doações de campanha no município")
    return out


# ---------------------------------------------------------------------------
# 4. ALERTAS
# ---------------------------------------------------------------------------
def gerar_alertas(contratos, cnpj_info, agentes, doacoes):
    alertas = []

    def add(doc, nivel, tipo, texto):
        alertas.append({"doc_fornecedor": doc, "nivel": nivel, "tipo": tipo, "evidencia": texto})

    total_geral = contratos["valor"].sum() or 1
    agentes_norm = [(norm(a.nome), sobrenomes_raros(a.nome), a.nome, a.cargo) for a in agentes.itertuples()]
    doa_por_doc = defaultdict(list)
    doa_por_nome = defaultdict(list)
    for d in doacoes.itertuples():
        doa_por_doc[d.doc_doador].append(d)
        doa_por_nome[norm(d.doador)].append(d)

    # índices para achar fornecedores ligados entre si (mesmo sócio ou mesmo endereço)
    docs = [d for d in contratos["doc_fornecedor"].unique() if d and len(d) == 14]
    por_socio, por_end = defaultdict(set), defaultdict(set)
    for d in docs:
        inf = cnpj_info.get(d) or {}
        for s in inf.get("qsa") or []:
            n = norm(s.get("nome_socio"))
            if len(n) > 5:
                por_socio[n].add(d)
        end = norm_end(inf)
        if end:
            por_end[end].add(d)
    nome_de = lambda d: (cnpj_info.get(d) or {}).get("razao_social") or d

    for doc, grupo in contratos.groupby("doc_fornecedor"):
        if not doc:
            continue
        info = cnpj_info.get(doc, {}) if len(doc) == 14 else {}
        total = grupo["valor"].sum()

        # A7 — fornecedores ligados entre si
        for s in info.get("qsa") or []:
            outros = sorted(por_socio.get(norm(s.get("nome_socio")), set()) - {doc})
            if outros:
                add(doc, "MÉDIO", "Sócio em outro fornecedor",
                    f"{s.get('nome_socio')} também é sócio de {', '.join(map(nome_de, outros))}, "
                    f"que também tem contrato com a prefeitura.")
        end = norm_end(info)
        mesmos = sorted(por_end.get(end, set()) - {doc}) if end else []
        if mesmos:
            add(doc, "MÉDIO", "Mesmo endereço de outro fornecedor",
                f"Endereço na Receita igual ao de {', '.join(map(nome_de, mesmos))}.")
        primeira = grupo["data"].dropna().min() if grupo["data"].notna().any() else None
        socios = [s.get("nome_socio", "") for s in (info.get("qsa") or [])]
        if len(doc) != 14:
            socios = [grupo["fornecedor"].iloc[0]]  # pessoa física: o próprio contratado

        # A1 — empresa recém-aberta
        abertura = data(info.get("data_inicio_atividade"))
        if abertura and primeira:
            dias = (primeira - abertura).days
            if 0 <= dias < 180:
                add(doc, "ALTO", "Empresa nova",
                    f"Aberta em {abertura:%d/%m/%Y}, só {dias} dias antes do 1º contrato ({primeira:%d/%m/%Y}).")
            elif 180 <= dias < 365:
                add(doc, "MÉDIO", "Empresa nova",
                    f"Aberta em {abertura:%d/%m/%Y}, {dias} dias antes do 1º contrato ({primeira:%d/%m/%Y}).")

        # A2 — vínculo com agente público (nome igual ou sobrenome raro em comum)
        for socio in socios:
            sn, ss = norm(socio), sobrenomes_raros(socio)
            for an, asob, anome, acargo in agentes_norm:
                if not an:
                    continue
                if sn == an:
                    add(doc, "ALTO", "Agente público no quadro",
                        f"{socio} (sócio/contratado) tem o mesmo nome de {acargo}. Confirmar se é a mesma pessoa.")
                else:
                    comuns = ss & asob
                    if len(comuns) >= 2:
                        add(doc, "MÉDIO", "Possível parentesco",
                            f"Sócio {socio} e {anome} ({acargo}) compartilham os sobrenomes {', '.join(sorted(comuns))}.")
                    elif len(comuns) == 1:
                        add(doc, "BAIXO", "Possível parentesco",
                            f"Sócio {socio} e {anome} ({acargo}) compartilham o sobrenome {next(iter(comuns))}.")

        # A3 — doador de campanha
        for d in doa_por_doc.get(doc, []):
            add(doc, "ALTO", "Doador de campanha",
                f"O próprio fornecedor doou {brl(d.valor)} para {d.candidato} ({d.cargo}) em 2024.")
        for socio in socios:
            for d in doa_por_nome.get(norm(socio), []):
                if d.doc_doador == doc:
                    continue
                add(doc, "ALTO", "Doador de campanha",
                    f"{socio} (sócio/contratado) aparece como doador de {brl(d.valor)} para "
                    f"{d.candidato} ({d.cargo}) em 2024. Conferir se não é homônimo.")

        # A4 — perfil incompatível
        situ = norm(info.get("descricao_situacao_cadastral"))
        if situ and situ != "ATIVA":
            add(doc, "ALTO", "Situação cadastral", f"Situação atual na Receita: {situ}.")
        if info.get("opcao_pelo_mei"):
            com_data = grupo.dropna(subset=["data"])
            por_ano = com_data.groupby(com_data["data"].map(lambda d: d.year))["valor"].sum()
            for ano, v in por_ano.items():
                if v > 81_000:
                    add(doc, "MÉDIO", "MEI acima do limite",
                        f"É MEI (faturamento máximo R$ 81 mil/ano) e recebeu {brl(v)} em contratos em {ano}.")
        cap = num(info.get("capital_social"))
        if info and total > 50_000 and cap < 0.10 * total:
            add(doc, "MÉDIO", "Capital baixo",
                f"Capital social de {brl(cap)} para {brl(total)} em contratos (menos de 10%).")

        # A5 — contratações repetidas sem licitação
        sem_lic = grupo[grupo["modalidade"].isin(SEM_LICITACAO)].dropna(subset=["data"])
        if len(sem_lic):
            for ano, gr in sem_lic.groupby(sem_lic["data"].map(lambda d: d.year)):
                if len(gr) >= 3:
                    add(doc, "MÉDIO", "Contratos sem licitação repetidos",
                        f"{len(gr)} dispensas/inexigibilidades em {ano}, somando {brl(gr['valor'].sum())}. "
                        f"Pode indicar fracionamento de despesa.")

        # A5b — vários registros no mesmo dia (possível fracionamento)
        com_data = grupo.dropna(subset=["data"])
        for dia, gr in com_data.groupby("data"):
            if len(gr) >= 3:
                add(doc, "MÉDIO", "Vários contratos no mesmo dia",
                    f"{len(gr)} contratos em {dia:%d/%m/%Y}, somando {brl(gr['valor'].sum())}. "
                    f"Pode indicar fracionamento de despesa.")

        # A5c — aditivos de acréscimo repetidos
        if "tipo_registro" in grupo.columns:
            adit = grupo[grupo["tipo_registro"].fillna("").str.upper().str.contains("ACR")]
            if len(adit) >= 2:
                add(doc, "BAIXO", "Aditivos de acréscimo",
                    f"{len(adit)} aditivos de acréscimo, somando {brl(adit['valor'].sum())}. "
                    f"Vale comparar com o valor original e a justificativa.")

        # A6 — concentração
        share = total / total_geral
        if share >= 0.15 and contratos["doc_fornecedor"].nunique() >= 5:
            add(doc, "BAIXO", "Concentração",
                f"Recebeu {share:.0%} de todo o valor contratado no período analisado.")

    return pd.DataFrame(alertas, columns=["doc_fornecedor", "nivel", "tipo", "evidencia"])


def montar_fornecedores(contratos, cnpj_info, alertas):
    linhas = []
    for doc, grupo in contratos.groupby("doc_fornecedor"):
        if not doc:
            continue
        info = cnpj_info.get(doc, {}) if len(doc) == 14 else {}
        al = alertas[alertas["doc_fornecedor"] == doc]
        end = ", ".join(x for x in [info.get("logradouro"), info.get("numero"), info.get("bairro"),
                                    info.get("municipio"), info.get("uf")] if x)
        linhas.append({
            "doc_fornecedor": doc,
            "documento": cnpj_fmt(doc),
            "fornecedor": info.get("razao_social") or grupo["fornecedor"].iloc[0],
            "tipo": "PJ" if len(doc) == 14 else "PF",
            "contratos": len(grupo),
            "valor_total": grupo["valor"].sum(),
            "abertura": info.get("data_inicio_atividade", ""),
            "porte": "MEI" if info.get("opcao_pelo_mei") else info.get("porte", ""),
            "situacao": info.get("descricao_situacao_cadastral", ""),
            "atividade": info.get("cnae_fiscal_descricao", ""),
            "endereco": end,
            "socios": "; ".join(s.get("nome_socio", "") for s in (info.get("qsa") or [])),
            "pontuacao": int(al["nivel"].map(PESO).sum()),
            "alertas": len(al),
        })
    df = pd.DataFrame(linhas)
    return df.sort_values(["pontuacao", "valor_total"], ascending=False) if len(df) else df


# ---------------------------------------------------------------------------
# 5. RELATÓRIO HTML
# ---------------------------------------------------------------------------
def esc(t):
    return (str(t if t is not None else "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def gerar_html(cfg, contratos, fornecedores, alertas, fontes_usadas, caminho):
    hoje = date.today().strftime("%d/%m/%Y")
    n_alto = (alertas["nivel"] == "ALTO").sum()
    n_med = (alertas["nivel"] == "MÉDIO").sum()
    periodo = (f"{contratos['data'].dropna().min():%d/%m/%Y} a {contratos['data'].dropna().max():%d/%m/%Y}"
               if contratos["data"].notna().any() else "—")

    aviso_vazio = ""
    if contratos.empty:
        aviso_vazio = """<div class="aviso"><b>Nenhum contrato encontrado.</b> Municípios com até 20 mil
        habitantes têm até 1º de abril de 2027 para integrar seus sistemas ao PNCP, então os dados podem
        ainda não estar lá. Baixe a lista de despesas ou contratos no portal da transparência da prefeitura
        (ou no TCE-MA) em CSV/Excel e informe o arquivo em <code>arquivo_contratos_csv</code>.</div>"""

    cards = []
    for f in fornecedores.itertuples():
        al = alertas[alertas["doc_fornecedor"] == f.doc_fornecedor]
        if al.empty:
            continue
        al = al.assign(p=al["nivel"].map(PESO)).sort_values("p", ascending=False)
        itens = "".join(f'<li><span class="tag {a.nivel[0]}">{esc(a.nivel)}</span> <b>{esc(a.tipo)}</b> — {esc(a.evidencia)}</li>'
                        for a in al.itertuples())
        cts = contratos[contratos["doc_fornecedor"] == f.doc_fornecedor].sort_values("data")
        linhas = "".join(
            f"<tr><td>{c.data.strftime('%d/%m/%Y') if c.data else '—'}</td><td>{esc(c.modalidade)}</td>"
            f"<td>{esc(str(c.objeto)[:140])}</td><td class='n'>{brl(c.valor)}</td>"
            f"<td>{f'<a href={chr(34)}{c.link}{chr(34)} target=_blank>ver</a>' if c.link else ''}</td></tr>"
            for c in cts.itertuples())
        cards.append(f"""
        <details class="card"><summary>
          <span class="score">{f.pontuacao}</span>
          <span class="nome">{esc(f.fornecedor)}<small>{esc(f.documento)} · {f.contratos} contrato(s) · {brl(f.valor_total)}</small></span>
        </summary>
          <ul class="alertas">{itens}</ul>
          <p class="meta"><b>Abertura:</b> {esc(f.abertura) or '—'} · <b>Porte:</b> {esc(f.porte) or '—'} ·
             <b>Situação:</b> {esc(f.situacao) or '—'}<br><b>Atividade:</b> {esc(f.atividade) or '—'}<br>
             <b>Endereço:</b> {esc(f.endereco) or '—'}<br><b>Sócios:</b> {esc(f.socios) or '—'}</p>
          <table><tr><th>Data</th><th>Modalidade</th><th>Objeto</th><th>Valor</th><th></th></tr>{linhas}</table>
        </details>""")

    ranking = "".join(
        f"<tr><td>{esc(f.fornecedor)}<br><small>{esc(f.documento)}</small></td><td class='n'>{f.contratos}</td>"
        f"<td class='n'>{brl(f.valor_total)}</td><td class='n'><b>{f.pontuacao}</b></td></tr>"
        for f in fornecedores.head(15).itertuples())

    html = f"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Radar da Transparência — {esc(cfg['municipio'])}</title>
<style>
:root{{--bg:#f6f5f1;--card:#fff;--ink:#1d211f;--mut:#69706b;--line:#e3e1da;--alto:#b3261e;--med:#b26a00;--baixo:#55705f;--acc:#22433a}}
*{{box-sizing:border-box}}body{{margin:0;font:15px/1.5 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;background:var(--bg);color:var(--ink)}}
main{{max-width:960px;margin:auto;padding:24px 16px 64px}}
h1{{font-size:26px;margin:0}}h2{{font-size:18px;margin:36px 0 12px;border-bottom:1px solid var(--line);padding-bottom:6px}}
.sub{{color:var(--mut);margin:4px 0 20px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}}
.kpi{{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px}}
.kpi b{{display:block;font-size:22px}}.kpi span{{color:var(--mut);font-size:13px}}
.aviso{{background:#fff7e6;border:1px solid #f0d9a8;border-radius:10px;padding:12px 14px;margin:16px 0}}
.nota{{background:#eef3f0;border-left:4px solid var(--acc);padding:12px 14px;border-radius:6px}}
table{{width:100%;border-collapse:collapse;font-size:13.5px;background:var(--card)}}
th,td{{text-align:left;padding:7px 8px;border-bottom:1px solid var(--line);vertical-align:top}}th{{color:var(--mut);font-weight:600}}
td.n{{text-align:right;white-space:nowrap}}small{{color:var(--mut)}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:10px;margin:10px 0;padding:0 14px}}
.card summary{{cursor:pointer;display:flex;gap:12px;align-items:center;padding:12px 0;list-style:none}}
.card summary::-webkit-details-marker{{display:none}}
.score{{min-width:38px;height:38px;border-radius:50%;background:var(--acc);color:#fff;display:grid;place-items:center;font-weight:700}}
.nome{{font-weight:600}}.nome small{{display:block;font-weight:400}}
.alertas{{padding-left:18px}}.alertas li{{margin:6px 0}}
.tag{{font-size:11px;font-weight:700;padding:2px 6px;border-radius:4px;color:#fff}}
.tag.A{{background:var(--alto)}}.tag.M{{background:var(--med)}}.tag.B{{background:var(--baixo)}}
.meta{{font-size:13.5px;color:#3c423e}}.card table{{margin-bottom:14px}}
a{{color:var(--acc)}}
</style></head><body><main>
<h1>Radar da Transparência</h1>
<p class="sub">{esc(cfg['municipio'])} – {esc(cfg['uf'])} · gerado em {hoje} · contratos de {periodo}</p>

<div class="nota"><b>Como ler este relatório.</b> Cada alerta é um <b>indício</b> calculado automaticamente
a partir de dados públicos. Ele não prova irregularidade. Homônimos, erros de cadastro e explicações legítimas
são comuns. Antes de publicar qualquer coisa: confira os documentos originais, procure o órgão e o fornecedor
para que se manifestem e use linguagem de indício ("chama atenção", "não foi possível confirmar").</div>
{aviso_vazio}

<h2>Resumo</h2>
<div class="kpis">
 <div class="kpi"><b>{len(contratos)}</b><span>contratos analisados</span></div>
 <div class="kpi"><b>{brl(contratos['valor'].sum())}</b><span>valor total</span></div>
 <div class="kpi"><b>{len(fornecedores)}</b><span>fornecedores</span></div>
 <div class="kpi"><b>{n_alto}</b><span>alertas altos</span></div>
 <div class="kpi"><b>{n_med}</b><span>alertas médios</span></div>
</div>

<h2>Fornecedores com alertas</h2>
<p class="sub">Ordenados pela pontuação de risco (alto = 3, médio = 2, baixo = 1). Toque para abrir.</p>
{''.join(cards) or '<p>Nenhum alerta gerado.</p>'}

<h2>Maiores pontuações</h2>
<table><tr><th>Fornecedor</th><th>Contratos</th><th>Valor</th><th>Pontos</th></tr>{ranking}</table>

<h2>Método</h2>
<p><b>Fontes:</b> {esc(', '.join(fontes_usadas))}.</p>
<ul>
<li><b>Empresa nova:</b> aberta menos de 180 dias (alto) ou 365 dias (médio) antes do primeiro contrato.</li>
<li><b>Agente público / parentesco:</b> sócio com nome idêntico a eleito ou servidor (alto), ou com sobrenomes
pouco comuns em comum (médio com dois, baixo com um). Sobrenomes muito frequentes são ignorados.</li>
<li><b>Doador de campanha:</b> fornecedor ou sócio que aparece como doador nas eleições municipais de 2024.</li>
<li><b>Perfil incompatível:</b> situação cadastral diferente de ativa, MEI com contratos acima de R$ 81 mil no
ano, ou capital social abaixo de 10% do valor contratado.</li>
<li><b>Sem licitação repetida:</b> três ou mais dispensas/inexigibilidades para o mesmo fornecedor no ano.</li>
<li><b>Concentração:</b> fornecedor com 15% ou mais de todo o valor contratado.</li>
<li><b>Fornecedores ligados:</b> empresas diferentes com sócio em comum ou com o mesmo endereço na Receita.</li>
<li><b>Vários contratos no mesmo dia</b> e <b>aditivos de acréscimo repetidos</b> para o mesmo fornecedor.</li>
</ul>
<p><small>CPFs aparecem mascarados. Dados cadastrais refletem a situação atual na Receita Federal, que pode ser
diferente da época do contrato.</small></p>
</main></body></html>"""
    with open(caminho, "w", encoding="utf-8") as fh:
        fh.write(html)


# ---------------------------------------------------------------------------
# EXECUÇÃO
# ---------------------------------------------------------------------------
def executar(cfg_usuario=None):
    cfg = {**CONFIG, **(cfg_usuario or {})}
    os.makedirs(cfg["pasta_saida"], exist_ok=True)
    fontes, partes = [], []

    if cfg["usar_pncp"]:
        df = coletar_pncp(cfg)
        if len(df):
            fontes.append("PNCP")
        partes.append(df)
    if cfg.get("usar_portal") and cfg.get("portal_prefeitura"):
        dfp = coletar_portal(cfg)
        if len(dfp):
            fontes.append("Portal da transparência da prefeitura")
        partes.append(dfp)
    if cfg["arquivo_contratos_csv"]:
        partes.append(carregar_planilha_contratos(cfg["arquivo_contratos_csv"]))
        fontes.append("planilha do portal municipal")
    partes = [p for p in partes if len(p)]
    contratos = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(
        columns=["fonte", "orgao", "cnpj_orgao", "doc_fornecedor", "fornecedor", "valor", "data", "objeto", "modalidade", "link"])
    if len(contratos):
        # o mesmo contrato pode aparecer no PNCP e no portal: mantém um só
        contratos["_k"] = contratos["doc_fornecedor"].astype(str) + "|" + contratos["valor"].round(2).astype(str) + "|" + contratos["data"].astype(str)
        antes = len(contratos)
        contratos = contratos.sort_values("fonte").drop_duplicates("_k").drop(columns="_k").reset_index(drop=True)
        if antes > len(contratos):
            log(f"\n{antes - len(contratos)} registros repetidos entre fontes foram unificados")
    log(f"\nTotal: {len(contratos)} contratos")

    pjs = sorted({d for d in contratos["doc_fornecedor"] if len(d) == 14})
    cnpj_info = consultar_cnpjs(pjs, cfg["pasta_saida"]) if pjs else {}
    if pjs:
        fontes.append("Receita Federal (via BrasilAPI/Minha Receita)")

    agentes = carregar_agentes(cfg)
    doacoes = carregar_doacoes(cfg)
    if len(agentes) or len(doacoes):
        fontes.append("TSE – eleições 2024")
    if cfg["arquivo_agentes_csv"]:
        fontes.append("lista de agentes públicos informada pelo usuário")

    alertas = gerar_alertas(contratos, cnpj_info, agentes, doacoes)
    fornecedores = montar_fornecedores(contratos, cnpj_info, alertas)

    html_path = os.path.join(cfg["pasta_saida"], "relatorio_radar.html")
    xlsx_path = os.path.join(cfg["pasta_saida"], "radar_dados.xlsx")
    gerar_html(cfg, contratos, fornecedores, alertas, fontes or ["nenhuma"], html_path)

    with pd.ExcelWriter(xlsx_path) as xw:
        fornecedores.drop(columns=["doc_fornecedor"], errors="ignore").to_excel(xw, sheet_name="Fornecedores", index=False)
        alertas.assign(documento=alertas["doc_fornecedor"].map(cnpj_fmt)).drop(columns="doc_fornecedor") \
               .to_excel(xw, sheet_name="Alertas", index=False)
        contratos.assign(documento=contratos["doc_fornecedor"].map(cnpj_fmt)).drop(columns="doc_fornecedor") \
                 .to_excel(xw, sheet_name="Contratos", index=False)
        agentes.to_excel(xw, sheet_name="Agentes", index=False)
        if len(doacoes):
            doacoes.assign(doc_doador=doacoes["doc_doador"].map(cnpj_fmt)).to_excel(xw, sheet_name="Doacoes", index=False)

    global ULTIMO_RESULTADO
    ULTIMO_RESULTADO = {"cfg": cfg, "contratos": contratos, "fornecedores": fornecedores,
                        "alertas": alertas, "agentes": agentes, "doacoes": doacoes,
                        "fontes": fontes, "html": html_path, "xlsx": xlsx_path}
    log(f"\nPronto! {len(alertas)} alertas ({(alertas['nivel'] == 'ALTO').sum()} altos).")
    log(f"Relatório: {html_path}\nPlanilha:  {xlsx_path}")
    return html_path, xlsx_path


if __name__ == "__main__":
    executar()
