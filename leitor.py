# -*- coding: utf-8 -*-
"""
Leitor de planilhas e relatórios para o Radar da Transparência.

Aceita CSV, Excel (.xlsx/.xls) e PDF com tabela. Acha a linha de cabeçalho
mesmo quando há títulos antes dela, reconhece os nomes de coluna mais comuns
dos portais e, quando a coluna não existe, procura CNPJ/CPF, valor (R$) e data
dentro da própria linha. Gera um CSV limpo com as colunas que o radar entende.
"""
import csv
import os
import re
import unicodedata

import pandas as pd

RE_CNPJ = re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")
RE_CPF = re.compile(r"[\d\*]{3}\.?[\d\*]{3}\.?[\d\*]{3}-[\d\*]{2}")
RE_VALOR = re.compile(r"(?:R\$\s*)?(\d{1,3}(?:\.\d{3})*,\d{2})")
RE_DATA = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")

ALIAS = {
    "cnpj": ["cnpj", "cpf", "cpf cnpj", "cnpj cpf", "documento", "cnpj fornecedor", "cnpj credor",
             "cnpj do contratado", "cpf cnpj do contratado", "cnpj contratado", "ni fornecedor"],
    "fornecedor": ["fornecedor", "credor", "razao social", "nome", "favorecido", "contratado", "contratada",
                   "nome do contratado", "nome do credor", "nome contratado", "beneficiario"],
    "valor": ["valor", "valor pago", "valor total", "valor contrato", "valor do contrato", "valor global",
              "pago", "empenhado", "valor empenhado", "valor liquidado", "valor r", "valor inicial"],
    "data": ["data", "data assinatura", "data da assinatura", "data pagamento", "data empenho",
             "data do empenho", "data contrato", "data do contrato", "data inicio", "inicio vigencia", "dt"],
    "objeto": ["objeto", "secretaria objeto", "descricao", "historico", "especificacao", "descricao do objeto", "objeto do contrato"],
    "modalidade": ["modalidade", "modalidade licitacao", "modalidade de licitacao", "tipo licitacao",
                   "tipo", "tipo de contrato"],
    "orgao": ["orgao", "unidade", "secretaria", "unidade gestora", "entidade", "contratante"],
}
ALIAS_INV = {a: k for k, lst in ALIAS.items() for a in lst}


def _chave(txt):
    txt = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    txt = re.sub(r"[^A-Za-z ]", " ", txt).lower()
    return re.sub(r"\s+", " ", txt).strip()


def _celula(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _ler_linhas(caminho):
    """Devolve uma lista de tabelas; cada tabela é uma lista de linhas (listas de células)."""
    ext = os.path.splitext(caminho)[1].lower()
    if ext == ".pdf":
        import pdfplumber
        linhas = []
        with pdfplumber.open(caminho) as pdf:
            for pagina in pdf.pages:
                tabelas = pagina.extract_tables() or []
                if not tabelas:  # PDF sem linhas de grade: tenta pelo alinhamento do texto
                    tabelas = pagina.extract_tables({"vertical_strategy": "text",
                                                     "horizontal_strategy": "text"}) or []
                for t in tabelas:
                    linhas.extend([[_celula(c) for c in l] for l in t])
        return [linhas]
    if ext in (".xlsx", ".xls"):
        abas = pd.read_excel(caminho, sheet_name=None, header=None, dtype=str)
        return [[[_celula(c) for c in l] for l in df.values.tolist()] for df in abas.values()]
    for enc in ("utf-8-sig", "latin1"):
        try:
            with open(caminho, encoding=enc, newline="") as fh:
                amostra = fh.read(4096)
                fh.seek(0)
                sep = ";" if amostra.count(";") >= amostra.count(",") else ","
                return [[[_celula(c) for c in l] for l in csv.reader(fh, delimiter=sep)]]
        except UnicodeDecodeError:
            continue
    return [[]]


def _mapear_cabecalho(linha):
    mapa = {}
    for i, c in enumerate(linha):
        k = ALIAS_INV.get(_chave(c))
        if k and k not in mapa:
            mapa[k] = i
    return mapa


def normalizar(caminho, pasta_saida):
    """Lê o arquivo e grava um CSV com colunas padronizadas. Devolve (caminho_csv, resumo)."""
    melhor = None
    for tabela in _ler_linhas(caminho):
        for i, linha in enumerate(tabela[:60]):
            mapa = _mapear_cabecalho(linha)
            if melhor is None or len(mapa) > len(melhor[2]):
                melhor = (tabela, i, mapa)
    if melhor is None or not melhor[0]:
        raise ValueError("Não encontrei nenhuma tabela no arquivo.")
    tabela, idx, mapa = melhor
    cab = tabela[idx] if mapa else []
    dados = tabela[idx + 1:] if len(mapa) >= 2 else tabela

    saida = []
    for linha in dados:
        if not any(linha) or linha == cab:
            continue  # linhas vazias e cabeçalho repetido em outra página do PDF
        texto = " | ".join(linha)
        pega = lambda k: linha[mapa[k]] if k in mapa and mapa[k] < len(linha) else ""
        doc = pega("cnpj")
        m_doc = RE_CNPJ.search(doc or texto) or RE_CPF.search(doc or texto)
        valor = pega("valor")
        if not RE_VALOR.search(valor or ""):
            m_val = RE_VALOR.search(texto)
            valor = m_val.group(1) if m_val else ""
        data = pega("data")
        if not RE_DATA.search(data or ""):
            m_dt = RE_DATA.search(texto)
            data = m_dt.group(0) if m_dt else ""
        forn = pega("fornecedor")
        if not forn and m_doc:  # nome e documento na mesma célula
            celula = next((c for c in linha if m_doc.group(0) in c), "")
            forn = celula.replace(m_doc.group(0), "").strip(" -–:|")
        if forn and m_doc:
            forn = forn.replace(m_doc.group(0), "").strip(" -–:|")
        if not (m_doc or forn) or not valor:
            continue
        saida.append({"cnpj": m_doc.group(0) if m_doc else "", "fornecedor": forn, "valor": valor,
                      "data": RE_DATA.search(data).group(0) if RE_DATA.search(data or "") else data,
                      "objeto": pega("objeto") or "Não informado", "modalidade": pega("modalidade") or "Não informada",
                      "orgao": pega("orgao")})

    if not saida:
        lidas = ", ".join(c for c in cab if c)[:300] or "nenhuma"
        raise ValueError(f"Não encontrei linhas com fornecedor e valor. Colunas lidas: {lidas}.")
    destino = os.path.join(pasta_saida, "planilha_normalizada.csv")
    pd.DataFrame(saida).to_csv(destino, sep=";", index=False, encoding="utf-8")
    resumo = (f"{len(saida)} linhas lidas de {os.path.basename(caminho)}"
              f" (colunas reconhecidas: {', '.join(sorted(mapa)) or 'nenhuma, usei busca na linha'})")
    return destino, resumo
