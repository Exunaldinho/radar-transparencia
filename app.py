# -*- coding: utf-8 -*-
"""
Radar da Transparência — app web
Rodar no PC:   streamlit run app.py
"""
import os, tempfile
from datetime import date

import pandas as pd
import streamlit as st

import leitor
import radar

st.set_page_config(page_title="Radar da Transparência", page_icon="🔎", layout="wide")

st.markdown("""
<style>
.block-container{padding-top:2rem;max-width:1150px}
.nota{background:#eef3f0;border-left:4px solid #22433a;padding:12px 16px;border-radius:6px;margin:0 0 18px}
.tag{font-size:11px;font-weight:700;padding:2px 7px;border-radius:4px;color:#fff;margin-right:6px}
.ALTO{background:#b3261e}.MÉDIO{background:#b26a00}.BAIXO{background:#55705f}
.alerta{margin:6px 0;line-height:1.5}
</style>""", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# BARRA LATERAL — CONFIGURAÇÃO
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("Configuração")
    municipio = st.text_input("Município", radar.CONFIG["municipio"])
    c1, c2 = st.columns([1, 2])
    uf = c1.text_input("UF", radar.CONFIG["uf"]).upper()
    ibge = c2.text_input("Código IBGE", radar.CONFIG["codigo_ibge"])
    ano_atual = date.today().year
    anos = st.multiselect("Anos", list(range(2021, ano_atual + 1)),
                          default=[a for a in radar.CONFIG["anos"] if a <= ano_atual])

    st.subheader("Fontes")
    usar_portal = st.toggle("Buscar no portal da prefeitura", True,
                            help="Lê a lista pública de contratos do portal da transparência.")
    portal = st.text_input("Endereço do portal", radar.CONFIG["portal_prefeitura"], disabled=not usar_portal)
    usar_pncp = st.toggle("Buscar contratos no PNCP", True)
    usar_tse = st.toggle("Cruzar com o TSE (eleitos e doadores 2024)", False,
                         help="Baixa arquivos grandes do TSE (centenas de MB). Na versão online gratuita "
                              "pode faltar memória; se travar, rode essa parte no PC.")
    cnpjs_extra = st.text_area("CNPJs de órgãos municipais (opcional)",
                               help="Um por linha: prefeitura, fundos de saúde, educação etc. "
                                    "O radar também tenta descobrir sozinho.")
    arq_contratos = st.file_uploader("Planilha ou relatório de despesas/contratos (opcional)",
                                     type=["csv", "xlsx", "xls", "pdf"],
                                     help="CSV, Excel ou PDF baixado do portal da prefeitura ou do TCE-MA. "
                                          "Não é preciso enviar a lista de contratos do portal: o radar já lê sozinho.")
    arq_agentes = st.file_uploader("Lista de servidores/agentes (opcional)", type=["csv"],
                                   help="Colunas: nome;cargo")
    rodar = st.button("🔎 Rodar radar", type="primary", width="stretch")
    st.caption("Enquanto o radar roda, não mexa nos campos: qualquer mudança reinicia a análise.")

# --------------------------------------------------------------------------
# CABEÇALHO
# --------------------------------------------------------------------------
st.title("Radar da Transparência")
st.caption(f"{municipio} – {uf} · contratos do portal da prefeitura e do PNCP cruzados com CNPJ e TSE")
st.markdown("""<div class="nota"><b>Alerta não é prova.</b> Cada alerta é um indício calculado a partir de
dados públicos. Homônimos, erros de cadastro e explicações legítimas são comuns. Antes de publicar,
confira os documentos originais, ouça o órgão e o fornecedor e use linguagem de indício.</div>""",
            unsafe_allow_html=True)


def salvar_upload(arquivo, pasta):
    if not arquivo:
        return None
    caminho = os.path.join(pasta, arquivo.name)
    with open(caminho, "wb") as f:
        f.write(arquivo.getbuffer())
    return caminho


# --------------------------------------------------------------------------
# EXECUÇÃO
# --------------------------------------------------------------------------
if rodar:
    if not usar_pncp and not usar_portal and not arq_contratos:
        st.error("Ative o portal da prefeitura, o PNCP ou envie uma planilha de contratos.")
        st.stop()
    pasta = os.path.join(tempfile.gettempdir(), f"radar_{radar.so_digitos(ibge) or 'municipio'}")
    os.makedirs(pasta, exist_ok=True)  # pasta fixa por município: o cache de CNPJ/TSE é reaproveitado
    cfg = {
        "municipio": municipio, "uf": uf, "codigo_ibge": ibge, "anos": sorted(anos),
        "cnpjs_orgaos_extra": [c for c in cnpjs_extra.split() if radar.so_digitos(c)],
        "arquivo_contratos_csv": None,
        "arquivo_agentes_csv": salvar_upload(arq_agentes, pasta),
        "usar_portal": usar_portal, "portal_prefeitura": portal.strip(),
        "usar_pncp": usar_pncp, "usar_tse": usar_tse, "pasta_saida": pasta,
    }
    if arq_contratos:
        try:
            cfg["arquivo_contratos_csv"], resumo = leitor.normalizar(salvar_upload(arq_contratos, pasta), pasta)
            st.success(resumo)
        except Exception as e:
            st.error(f"Não consegui ler {arq_contratos.name}: {e}")
            st.stop()
    with st.status("Coletando e cruzando dados…", expanded=True) as status:
        linhas = []
        caixa = st.empty()

        def mostrar(msg):
            linhas.append(msg.strip())
            caixa.code("\n".join(l for l in linhas if l)[-3000:], language=None)

        radar.LOG_FN = mostrar
        try:
            radar.executar(cfg)
            st.session_state["res"] = radar.ULTIMO_RESULTADO
            status.update(label="Análise concluída", state="complete", expanded=False)
        except Exception as e:
            status.update(label="Erro na análise", state="error")
            st.exception(e)
        finally:
            radar.LOG_FN = None

res = st.session_state.get("res")
if not res:
    st.info("Ajuste a configuração na barra lateral e toque em **Rodar radar**. "
            "No celular, a barra lateral abre pelo ícone › no canto superior esquerdo.")
    st.stop()

contratos, fornecedores, alertas = res["contratos"], res["fornecedores"], res["alertas"]

if contratos.empty:
    st.warning("Nenhum contrato encontrado. Municípios com até 20 mil habitantes têm até 1º de abril de 2027 "
               "para integrar seus sistemas ao PNCP. Baixe a lista de despesas no portal da prefeitura ou "
               "do TCE-MA e envie pela barra lateral.")
    st.stop()

# --------------------------------------------------------------------------
# RESUMO
# --------------------------------------------------------------------------
k = st.columns(5)
k[0].metric("Contratos", f"{len(contratos):,}".replace(",", "."))
def brl_curto(v):
    for lim, suf in ((1e9, "bi"), (1e6, "mi"), (1e3, "mil")):
        if v >= lim:
            return f"R$ {v / lim:.1f} {suf}".replace(".", ",")
    return radar.brl(v)


k[1].metric("Valor total", brl_curto(contratos["valor"].sum()), help=radar.brl(contratos["valor"].sum()))
k[2].metric("Fornecedores", len(fornecedores))
k[3].metric("Alertas altos", int((alertas["nivel"] == "ALTO").sum()))
k[4].metric("Alertas médios", int((alertas["nivel"] == "MÉDIO").sum()))

aba_alertas, aba_graficos, aba_contratos, aba_baixar = st.tabs(
    ["🚩 Alertas", "📊 Gráficos", "📄 Contratos", "⬇️ Baixar"])

# --------------------------------------------------------------------------
# ALERTAS
# --------------------------------------------------------------------------
with aba_alertas:
    f1, f2, f3 = st.columns(3)
    niveis = f1.multiselect("Nível", ["ALTO", "MÉDIO", "BAIXO"], default=["ALTO", "MÉDIO", "BAIXO"])
    tipos = f2.multiselect("Tipo de alerta", sorted(alertas["tipo"].unique()))
    busca = f3.text_input("Buscar fornecedor ou CNPJ")

    filtrados = alertas[alertas["nivel"].isin(niveis)]
    if tipos:
        filtrados = filtrados[filtrados["tipo"].isin(tipos)]
    docs = set(filtrados["doc_fornecedor"])
    lista = fornecedores[fornecedores["doc_fornecedor"].isin(docs)]
    if busca:
        b = radar.norm(busca)
        dig = radar.so_digitos(busca)
        lista = lista[lista["fornecedor"].map(radar.norm).str.contains(b, na=False) |
                      (lista["doc_fornecedor"].str.contains(dig) if dig else False)]

    st.caption(f"{len(lista)} fornecedor(es) com alertas · ordenados pela pontuação "
               "(alto = 3, médio = 2, baixo = 1)")
    for f in lista.itertuples():
        al = filtrados[filtrados["doc_fornecedor"] == f.doc_fornecedor].copy()
        al["p"] = al["nivel"].map(radar.PESO)
        al = al.sort_values("p", ascending=False)
        with st.expander(f"**{f.pontuacao} pts** · {f.fornecedor} · {f.documento} · {radar.brl(f.valor_total)}"):
            for a in al.itertuples():
                st.markdown(f'<div class="alerta"><span class="tag {a.nivel}">{a.nivel}</span>'
                            f'<b>{radar.esc(a.tipo)}</b> — {radar.esc(a.evidencia)}</div>', unsafe_allow_html=True)
            st.markdown(f"**Abertura:** {f.abertura or '—'} · **Porte:** {f.porte or '—'} · "
                        f"**Situação:** {f.situacao or '—'}  \n**Atividade:** {f.atividade or '—'}  \n"
                        f"**Endereço:** {f.endereco or '—'}  \n**Sócios:** {f.socios or '—'}")
            cts = contratos[contratos["doc_fornecedor"] == f.doc_fornecedor].sort_values("data")
            st.dataframe(cts[["data", "modalidade", "objeto", "valor", "link"]], hide_index=True,
                         width="stretch",
                         column_config={"valor": st.column_config.NumberColumn("valor", format="R$ %.2f"),
                                        "link": st.column_config.LinkColumn("Link", display_text="abrir")})

# --------------------------------------------------------------------------
# GRÁFICOS
# --------------------------------------------------------------------------
with aba_graficos:
    g1, g2 = st.columns(2)
    top = fornecedores.nlargest(10, "valor_total")[["fornecedor", "valor_total"]].set_index("fornecedor")
    g1.subheader("Maiores fornecedores")
    g1.bar_chart(top, horizontal=True, color="#22433a")
    g2.subheader("Valor por modalidade")
    g2.bar_chart(contratos.groupby("modalidade")["valor"].sum().sort_values(), horizontal=True, color="#b26a00")
    com_data = contratos.dropna(subset=["data"]).copy()
    if len(com_data):
        com_data["mes"] = pd.to_datetime(com_data["data"]).dt.to_period("M").dt.to_timestamp()
        st.subheader("Valor contratado por mês")
        st.caption("Picos em ano eleitoral ou no fim de mandato merecem atenção.")
        st.bar_chart(com_data.groupby("mes")["valor"].sum(), color="#22433a")
    st.subheader("Alertas por tipo")
    st.bar_chart(alertas.groupby(["tipo", "nivel"]).size().unstack(fill_value=0), horizontal=True,
                 color=["#b3261e", "#55705f", "#b26a00"] if alertas["nivel"].nunique() == 3 else None)

# --------------------------------------------------------------------------
# CONTRATOS
# --------------------------------------------------------------------------
with aba_contratos:
    vis = contratos.assign(documento=contratos["doc_fornecedor"].map(radar.cnpj_fmt))
    st.dataframe(vis[["data", "orgao", "fornecedor", "documento", "modalidade", "objeto", "valor", "fonte", "link"]]
                 .sort_values("data", ascending=False), hide_index=True, width="stretch", height=560,
                 column_config={"valor": st.column_config.NumberColumn("valor", format="R$ %.2f"),
                                "link": st.column_config.LinkColumn("Link", display_text="abrir")})

# --------------------------------------------------------------------------
# DOWNLOADS
# --------------------------------------------------------------------------
with aba_baixar:
    st.write("Baixe o relatório para guardar, imprimir ou compartilhar, e a planilha com todos os dados.")
    with open(res["html"], "rb") as fh:
        st.download_button("Relatório (HTML)", fh.read(), "relatorio_radar.html", "text/html")
    with open(res["xlsx"], "rb") as fh:
        st.download_button("Planilha completa (Excel)", fh.read(), "radar_dados.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("Fontes: " + ", ".join(res["fontes"] or ["nenhuma"]))
