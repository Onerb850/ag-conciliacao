import time
import streamlit as st
import pandas as pd
from datetime import date, timedelta

from comum import (
    ARQUIVO_DE_MATERIAL,
    PASTA_PROJETO,
    carregar,
    ler_aba_historico,
    salvar_aba_historico,
    acumular_historico,
    limpa_mapa,
    normalizar_codigo,
    fator_conversao_caixas,
    cor_linha_status,
    formata_qtd_fisica,
    com_apelido,
    rotulo_familia_vazio,
    montar_lookup_ag_por_codigo,
    familia_tipo_por_codigo,
)

# Planilha de referência Data+Mapa->PA (colunas DATA, MAPA, PONTO DE APOIO). Sobe na
# mesma pasta do Drive sempre com o nome "CONC.csv" — carregar() busca por esse
# prefixo, igual aos outros arquivos (De Material, 03.07.13).
ARQUIVO_MAPA_PA = PASTA_PROJETO / "CONC.csv"


def limpar_numero_robusto(v) -> str:
    """Como limpa_mapa() do comum.py, mas também limpa strings tipo '257781.0' —
    sobra de quando um número volta do Excel como float depois de acumular no
    histórico. limpa_mapa() sozinha não dá conta disso (int('257781.0') falha),
    então usa essa aqui pra QUALQUER leitura que já passou por um round-trip no
    Excel (df_mapa_pa e df_mapas_ag_sem_filtro_data), não só no primeiro carregamento."""
    s = str(v).strip()
    if s.lower() in ("", "nan", "none"):
        return ""
    try:
        return str(int(float(s)))
    except Exception:
        return s

st.set_page_config(page_title="Conciliação de Mapas (AG)", layout="wide")
st.title("⚖️ Conciliação de Mapas (AG)")
st.caption("_\"Balança enganosa é abominação ao SENHOR, mas o peso justo lhe é agradável.\" — Provérbios 11:1_")

# Aba separada no historico_ag.xlsx só pra dados de simulação — nunca mistura com a
# aba "VazioPA" de produção. Ativar o modo simulação (sidebar) troca de qual aba o
# app lê, sem apagar nem sobrescrever nada real.
NOME_ABA_SIMULACAO = "VazioPA_Simulacao"
NOME_ABA_LOTE_SIMULACAO = "VazioPALote_Simulacao"

REGRAS_VAZIO = {
    "300ml": {"garrafas_por_cx": 23, "garrafeiras_por_cx": 1},
    "600ml": {"garrafas_por_cx": 24, "garrafeiras_por_cx": 1},
    "Verde 600": {"garrafas_por_cx": 24, "garrafeiras_por_cx": 1},
    "1L": {"garrafas_por_cx": 12, "garrafeiras_por_cx": 1},
}

# Pros formulários de CONFERÊNCIA (Manual e Lote) — não a Previsão, que continua
# mostrando Âmbar e Verde separados por cor. Aqui o conferente digita um total só de
# "600" (não separa fisicamente Âmbar de Verde), e o app soma a Saída dos dois
# códigos (600ml + Verde 600) antes de comparar com esse total único.
FAMILIAS_CONFERENCIA = ["300ml", "600ml", "1L"]


def normaliza_600(fam: str) -> str:
    """Trata 600ml e Verde 600 como a mesma família pra fins de conferência PA —
    conferente não separa fisicamente, então Saída e Retorno são comparados juntos."""
    return "600ml" if fam == "Verde 600" else fam


def rotulo_conferencia(fam: str) -> str:
    """Rótulo pros campos dos formulários de conferência — "600ml" vira o total
    combinado (Âmbar+Verde), não o rótulo padrão "600 AMBAR" (que sugeriria só Âmbar)."""
    if fam == "600ml":
        return "600 (Âmbar + Verde)"
    return rotulo_familia_vazio(fam)

CORES_RESUMO = {
    "verde": ("#EAF3DE", "#173404"),
    "vermelho": ("#FCEBEB", "#501313"),
    "amarelo": ("#FFF4D4", "#5A4000"),
    "azul": ("#CCE5FF", "#004085"),
    "cinza": ("#E9ECEF", "#495057"),
}
ICONES_RESUMO = {
    "verde": "✅",
    "vermelho": "❌",
    "amarelo": "⚠️",
    "azul": "🔎",
    "cinza": "⏳",
}


def renderizar_cards_resumo(itens: list[tuple[str, int, str]] | list[tuple[str, int, str, list[str] | None]]) -> None:
    colunas = st.columns(len(itens))
    for col, item in zip(colunas, itens):
        rotulo, valor, cor = item[0], item[1], item[2]
        detalhes = item[3] if len(item) > 3 else None
        bg, fg = CORES_RESUMO[cor]
        icone = ICONES_RESUMO.get(cor, "•")

        linhas_html = ""
        if detalhes:
            linhas_html = "".join(
                f'<div style="font-size:11px; color:{fg}; opacity:0.85; padding:2px 0; '
                f'border-top:1px solid {fg}18; margin-top:4px;">{d}</div>'
                for d in detalhes[:5]
            )
            if len(detalhes) > 5:
                linhas_html += (
                    f'<div style="font-size:10.5px; color:{fg}; opacity:0.6; margin-top:2px;">'
                    f'+{len(detalhes) - 5} outro(s)…</div>'
                )

        html_card = (
            f'<div style="background-color:{bg}; border-radius:10px; padding:8px 10px;">'
            f'<div style="display:flex; align-items:center; gap:6px;">'
            f'<span style="font-size:1.3em; line-height:1;">{icone}</span>'
            f'<span style="font-size:22px; font-weight:800; color:{fg}; line-height:1;">{valor}</span>'
            f'<span style="font-size:11px; color:{fg}; opacity:0.75; font-weight:600; '
            f'text-transform:uppercase; letter-spacing:0.3px;">{rotulo}</span>'
            f'</div>'
            f'{linhas_html}'
            f'</div>'
        )
        col.markdown(html_card, unsafe_allow_html=True)


CORES_FAROL_FAMILIA = {
    "600ml": ("#1F3B57", "#EAF1F8"),
    "Verde 600": ("#1F4720", "#E9F5EA"),
    "300ml": ("#8A6D1B", "#FBF3DF"),
    "1L": ("#2B2B2B", "#ECECEC"),
}
ORDEM_FAROL_FAMILIA = ["600ml", "Verde 600", "300ml", "1L"]


def renderizar_farol_previsao(dados_familia: dict, dados_outros: dict) -> None:
    familias_com_dado = [f for f in ORDEM_FAROL_FAMILIA if f in dados_familia]
    if familias_com_dado:
        cols = st.columns(len(familias_com_dado))
        for col, fam in zip(cols, familias_com_dado):
            d = dados_familia[fam]
            cor_header, cor_body = CORES_FAROL_FAMILIA[fam]
            rotulo = rotulo_familia_vazio(fam)
            html = (
                f'<div style="border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.08);">'
                f'<div style="background:{cor_header}; color:white; padding:9px 12px; display:flex; align-items:center; gap:7px;">'
                f'<span style="font-size:1.15em;">🍺</span>'
                f'<span style="font-weight:700; font-size:13px; letter-spacing:0.3px;">{rotulo}</span>'
                f'</div>'
                f'<div style="background:{cor_body}; padding:14px 12px;">'
                f'<div style="font-size:26px; font-weight:800; color:{cor_header}; line-height:1;">{d["caixas"]}'
                f'<span style="font-size:13px; font-weight:600;"> cx</span></div>'
                f'<div style="font-size:12px; color:{cor_header}; opacity:0.8; margin-top:4px;">+ {d["soltas"]} gf soltas</div>'
                f'</div>'
                f'</div>'
            )
            col.markdown(html, unsafe_allow_html=True)

    if dados_outros:
        st.write("")
        cols2 = st.columns(len(dados_outros))
        for col, (rotulo, valor) in zip(cols2, dados_outros.items()):
            html2 = (
                f'<div style="background:#3A3A38; border-radius:10px; padding:11px 10px; text-align:center;">'
                f'<div style="font-size:11px; color:#D8D6CE; font-weight:600; text-transform:uppercase; letter-spacing:0.4px;">{rotulo}</div>'
                f'<div style="font-size:21px; font-weight:800; color:white; margin-top:3px;">{valor} <span style="font-size:12px; font-weight:600;">un</span></div>'
                f'</div>'
            )
            col.markdown(html2, unsafe_allow_html=True)


def extrair_dados_farol(df_subset: pd.DataFrame, lookup_map: dict) -> tuple[dict, dict]:
    if df_subset.empty:
        return {}, {}

    previsao_agg = df_subset.groupby("Material")["Qtde_Saida"].sum().reset_index()
    previsao_agg = previsao_agg.rename(columns={"Qtde_Saida": "P Vazia"})
    previsao_agg = previsao_agg[previsao_agg["P Vazia"] > 0]

    if previsao_agg.empty:
        return {}, {}

    if "Descricao" in df_subset.columns:
        desc_previsao = df_subset.drop_duplicates(subset=["Material"])[["Material", "Descricao"]].rename(columns={"Descricao": "Desc_Previsao"})
        previsao_agg = previsao_agg.merge(desc_previsao, on="Material", how="left")
        previsao_agg["AG"] = [
            com_apelido(cod, str(desc)) for cod, desc in zip(previsao_agg["Material"], previsao_agg["Desc_Previsao"].fillna(""))
        ]
    else:
        previsao_agg["AG"] = previsao_agg["Material"]

    fam_tipo_previsao = previsao_agg["Material"].apply(lambda c: familia_tipo_por_codigo(c, lookup_map))
    previsao_agg["Familia"] = fam_tipo_previsao.apply(lambda ft: ft[0])
    previsao_agg["Tipo"] = fam_tipo_previsao.apply(lambda ft: ft[1])
    previsao_agg["P Vazia"] = previsao_agg["P Vazia"].round(0).astype(int)

    dados_familia_farol: dict[str, dict[str, int]] = {}
    dados_outros_farol: dict[str, int] = {}

    for _, r in previsao_agg.iterrows():
        fam, tipo, qtd, ag_label = r["Familia"], r["Tipo"], int(r["P Vazia"]), r["AG"]
        if tipo == "Garrafa" and fam in REGRAS_VAZIO:
            fator = int(fator_conversao_caixas(fam))
            acc = dados_familia_farol.setdefault(fam, {"caixas": 0, "soltas": 0})
            acc["caixas"] += qtd // fator
            acc["soltas"] += qtd % fator
        else:
            dados_outros_farol[ag_label] = dados_outros_farol.get(ag_label, 0) + qtd

    return dados_familia_farol, dados_outros_farol


def _chave_cor_status(status: str) -> str:
    s = str(status)
    if "✅" in s: return "verde"
    if "❌" in s: return "vermelho"
    if "⚠️" in s: return "amarelo"
    if "🔎" in s: return "azul"
    if "⏳" in s: return "cinza"
    return "cinza"


def _pill_status(status: str) -> str:
    cor_key = _chave_cor_status(status)
    bg, fg = CORES_RESUMO[cor_key]
    rotulo = str(status)
    for emoji in ("✅", "❌", "⚠️", "🔎", "⏳"):
        rotulo = rotulo.replace(emoji, "")
    rotulo = rotulo.strip()
    return (
        f'<span style="background:{bg}; color:{fg}; font-size:11.5px; font-weight:600; '
        f'padding:3px 11px; border-radius:999px; white-space:nowrap; display:inline-block;">− {rotulo}</span>'
    )


def _cor_diferenca(texto) -> str:
    t = str(texto).strip()
    if t.startswith("+"): return "#0F6E56"
    if t.startswith("-"): return "#A32D2D"
    return "#888780"


def renderizar_tabela_limpa(df: pd.DataFrame, colunas: list[str], col_status: str = "Status") -> None:
    if df.empty:
        st.caption("Nenhum registro.")
        return

    cols_dado = [c for c in colunas if c != col_status]

    cabecalho = "".join(
        f'<th style="padding:0 12px 8px; text-align:{"left" if i == 0 else "right"}; '
        f'font-size:11.5px; color:#888780; font-weight:600; white-space:nowrap;">{c}</th>'
        for i, c in enumerate(cols_dado)
    )
    if col_status in colunas:
        cabecalho += (
            '<th style="padding:0 12px 8px; text-align:right; font-size:11.5px; '
            'color:#888780; font-weight:600;">Status</th>'
        )

    linhas_html = []
    for _, row in df.iterrows():
        celulas = []
        for i, col in enumerate(cols_dado):
            val = row[col]
            alinhado = "left" if i == 0 else "right"
            cor = _cor_diferenca(val) if col == "Diferença" else "inherit"
            peso = "600" if col == "Diferença" else "400"
            celulas.append(
                f'<td style="padding:10px 12px; text-align:{alinhado}; font-size:13px; '
                f'color:{cor}; font-weight:{peso}; white-space:nowrap;">{val}</td>'
            )
        if col_status in colunas:
            celulas.append(f'<td style="padding:10px 12px; text-align:right;">{_pill_status(row[col_status])}</td>')
        linhas_html.append(f'<tr style="border-top:1px solid #E9ECEF;">{"".join(celulas)}</tr>')

    html_tabela = (
        '<div style="overflow-x:auto;">'
        '<table style="width:100%; border-collapse:collapse;">'
        f'<thead><tr>{cabecalho}</tr></thead>'
        f'<tbody>{"".join(linhas_html)}</tbody>'
        '</table>'
        '</div>'
    )
    st.markdown(html_tabela, unsafe_allow_html=True)


with st.sidebar:
    st.caption("Fonte: 02.05.01.csv (atualiza sozinho a cada 5 min)")
    if st.button("🔄 Recarregar tela", width="stretch"):
        st.rerun()
    intervalo_datas = st.date_input(
        "Considerar mapas do CONC.csv no período:",
        value=(date(2026, 8, 1), date.today()),
    )
    if isinstance(intervalo_datas, tuple) and len(intervalo_datas) == 2:
        data_inicio, data_fim = intervalo_datas
    else:
        data_inicio = intervalo_datas[0] if isinstance(intervalo_datas, tuple) else intervalo_datas
        data_fim = date.today()

    st.divider()
    modo_simulacao = st.checkbox(
        "🧪 Modo simulação",
        value=False,
        key="modo_simulacao_ativo",
        help="Isola TUDO (individual + lote) numa aba de teste separada, nunca mexe nos dados reais de produção.",
    )

if modo_simulacao:
    st.warning("🧪 **MODO SIMULAÇÃO ATIVO** — os dados de conferência abaixo são de teste, não reais. Desative na sidebar pra voltar ao normal.")

# --- De Material ---
df_de_material = carregar(ARQUIVO_DE_MATERIAL)
if df_de_material is not None and "Promax" in df_de_material.columns:
    df_de_material["Promax"] = normalizar_codigo(df_de_material["Promax"])
lookup_ag = montar_lookup_ag_por_codigo(df_de_material) if df_de_material is not None else {}

# --- Mapa PA ---
df_mapa_pa = carregar(ARQUIVO_MAPA_PA)
MAPA_CONSOLIDADO_LOOKUP: dict[str, str] = {}
if df_mapa_pa is not None:
    df_mapa_pa = df_mapa_pa.copy()
    df_mapa_pa.columns = df_mapa_pa.columns.str.strip()
    _renomear_mapa_pa = {}
    for _col in df_mapa_pa.columns:
        _col_upper = _col.strip().upper()
        if _col_upper == "DATA":
            _renomear_mapa_pa[_col] = "Data"
        elif _col_upper == "MAPA":
            _renomear_mapa_pa[_col] = "Mapa"
        elif _col_upper in ("PONTO DE APOIO", "PA"):
            _renomear_mapa_pa[_col] = "PA"
        elif _col_upper == "MAPA CONSOLIDADO":
            _renomear_mapa_pa[_col] = "MapaConsolidado"
    df_mapa_pa = df_mapa_pa.rename(columns=_renomear_mapa_pa)

    if "Mapa" in df_mapa_pa.columns:
        df_mapa_pa["Mapa"] = df_mapa_pa["Mapa"].apply(limpar_numero_robusto)
    if "Data" in df_mapa_pa.columns:
        _dt_mapa_pa = pd.to_datetime(df_mapa_pa["Data"], dayfirst=True, errors="coerce")
        df_mapa_pa["Data"] = _dt_mapa_pa.dt.strftime("%d/%m/%Y")
    if "PA" in df_mapa_pa.columns:
        df_mapa_pa["PA"] = df_mapa_pa["PA"].astype(str).str.strip()
    if "MapaConsolidado" in df_mapa_pa.columns:
        df_mapa_pa["MapaConsolidado"] = df_mapa_pa["MapaConsolidado"].apply(
            lambda v: "" if str(v).strip().lower() in ("", "nan", "none", "0") else limpar_numero_robusto(v)
        )
    else:
        df_mapa_pa["MapaConsolidado"] = ""


@st.cache_data(show_spinner=False, ttl=300)
def _ingerir_conc_no_historico(_df_conc_limpo: pd.DataFrame) -> pd.DataFrame:
    colunas_manter = [c for c in ["Data", "Mapa", "PA", "MapaConsolidado"] if c in _df_conc_limpo.columns]
    return acumular_historico(_df_conc_limpo[colunas_manter], "MapaPAHistorico", ["Data", "Mapa"])


if df_mapa_pa is not None and not df_mapa_pa.empty:
    df_mapa_pa = _ingerir_conc_no_historico(df_mapa_pa)
else:
    df_mapa_pa = ler_aba_historico("MapaPAHistorico")

if df_mapa_pa is not None and not df_mapa_pa.empty:
    if "Mapa" in df_mapa_pa.columns:
        df_mapa_pa["Mapa"] = df_mapa_pa["Mapa"].apply(limpar_numero_robusto)
    if "MapaConsolidado" in df_mapa_pa.columns:
        df_mapa_pa["MapaConsolidado"] = df_mapa_pa["MapaConsolidado"].apply(
            lambda v: "" if str(v).strip().lower() in ("", "nan", "none", "0") else limpa_mapa(v)
        )
    else:
        df_mapa_pa["MapaConsolidado"] = ""
    df_mapa_pa = df_mapa_pa.dropna(subset=["Mapa"])
    df_mapa_pa = df_mapa_pa[df_mapa_pa["Mapa"].str.lower() != "nan"]

if df_mapa_pa is not None and not df_mapa_pa.empty:
    for _, _linha in df_mapa_pa.iterrows():
        if _linha.get("MapaConsolidado") and _linha["MapaConsolidado"] != _linha["Mapa"]:
            MAPA_CONSOLIDADO_LOOKUP[_linha["Mapa"]] = _linha["MapaConsolidado"]


def resolver_mapa(mapa: str) -> str:
    return MAPA_CONSOLIDADO_LOOKUP.get(mapa, mapa)


def resolver_mapas(mapas) -> list[str]:
    resolvidos: list[str] = []
    vistos: set[str] = set()
    for m in mapas:
        alvo = resolver_mapa(m)
        if alvo not in vistos:
            vistos.add(alvo)
            resolvidos.append(alvo)
    return resolvidos


REVERSE_MAPA_CONSOLIDADO: dict[str, list[str]] = {}
for _orig, _cons in MAPA_CONSOLIDADO_LOOKUP.items():
    REVERSE_MAPA_CONSOLIDADO.setdefault(_cons, []).append(_orig)


def rotulo_mapa(mapa: str) -> str:
    originais = REVERSE_MAPA_CONSOLIDADO.get(mapa)
    if not originais:
        return mapa
    ordenados = sorted(originais, key=lambda m: int(m) if str(m).isdigit() else 0)
    return f"{'+'.join(ordenados)} (→{mapa})"


def buscar_mapas_por_data_pa(data_alvo, pa_alvo: str) -> list[str]:
    if df_mapa_pa is None or df_mapa_pa.empty or "Data" not in df_mapa_pa.columns or "PA" not in df_mapa_pa.columns:
        return []
    data_str = data_alvo.strftime("%d/%m/%Y")
    pa_alvo_norm = _PA_NORMALIZADO.get(pa_alvo.strip().upper(), pa_alvo).upper()
    pa_bate = df_mapa_pa["PA"].apply(lambda v: _PA_NORMALIZADO.get(str(v).strip().upper(), v).upper() == pa_alvo_norm)
    sub = df_mapa_pa[(df_mapa_pa["Data"] == data_str) & pa_bate]
    return sorted(sub["Mapa"].dropna().unique().tolist(), key=lambda m: int(m) if str(m).isdigit() else 0)


ARQUIVO_020501 = PASTA_PROJETO / "02.05.01.csv"
NOME_ABA_020501_HISTORICO = "Relatorio020501Historico"


def parse_qtde_entrada_robusta(serie: pd.Series) -> pd.Series:
    s = serie.astype(str).str.strip()
    s = s.str.replace(".", "", regex=False)
    s = s.str.replace("/", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0)


@st.cache_data(show_spinner=False, ttl=300)
def _ingerir_020501_no_historico(_df_020501_limpo: pd.DataFrame) -> pd.DataFrame:
    return acumular_historico(_df_020501_limpo, NOME_ABA_020501_HISTORICO, ["Data", "Mapa", "Material"])


df_020501 = carregar(ARQUIVO_020501)
if df_020501 is not None and not df_020501.empty:
    df_020501 = df_020501.copy()
    df_020501.columns = df_020501.columns.str.strip()
    if "Código Operação" in df_020501.columns:
        df_020501 = df_020501[pd.to_numeric(df_020501["Código Operação"], errors="coerce") == 554]
    if not df_020501.empty:
        df_020501["Mapa"] = df_020501["Mapa"].apply(limpar_numero_robusto)
        df_020501["Material"] = df_020501["Item"].apply(limpar_numero_robusto)
        df_020501["Descricao"] = df_020501["Descrição"].astype(str).str.strip() if "Descrição" in df_020501.columns else ""
        df_020501["Qtde_Saida"] = parse_qtde_entrada_robusta(df_020501["Qtde Entrada"])
        _dt_020501 = pd.to_datetime(df_020501["Data"], dayfirst=True, errors="coerce")
        df_020501["Data"] = _dt_020501.dt.strftime("%d/%m/%Y")
        df_020501 = df_020501.groupby(["Data", "Mapa", "Material", "Descricao"], as_index=False)["Qtde_Saida"].sum()
        df_020501_historico = _ingerir_020501_no_historico(df_020501)
    else:
        df_020501_historico = ler_aba_historico(NOME_ABA_020501_HISTORICO)
else:
    df_020501_historico = ler_aba_historico(NOME_ABA_020501_HISTORICO)

if df_020501_historico is not None and not df_020501_historico.empty:
    for _col_020501 in ["Mapa", "Material"]:
        if _col_020501 in df_020501_historico.columns:
            df_020501_historico[_col_020501] = df_020501_historico[_col_020501].apply(limpar_numero_robusto)
    if "Qtde_Saida" in df_020501_historico.columns:
        df_020501_historico["Qtde_Saida"] = pd.to_numeric(df_020501_historico["Qtde_Saida"], errors="coerce").fillna(0)


def calcular_saida_familias(mapas_resolvidos: list[str]) -> dict[str, int]:
    if df_020501_historico is None or df_020501_historico.empty or not mapas_resolvidos:
        return {}
    fam_tipo_calc = df_020501_historico["Material"].apply(lambda c: familia_tipo_por_codigo(c, lookup_ag))
    df_calc = df_020501_historico.copy()
    df_calc["Familia"] = fam_tipo_calc.apply(lambda ft: normaliza_600(ft[0]))
    df_calc["Tipo"] = fam_tipo_calc.apply(lambda ft: ft[1])
    df_calc = df_calc[(df_calc["Familia"] != "Outro") & (df_calc["Tipo"] != "Garrafeira")]
    df_calc = df_calc[df_calc["Mapa"].isin(mapas_resolvidos)]
    if df_calc.empty:
        return {}
    return df_calc.groupby("Familia")["Qtde_Saida"].sum().round(0).astype(int).to_dict()


def formata_un_em_cx(qtd: int, familia: str) -> str:
    fator = int(fator_conversao_caixas(familia)) or 1
    qtd = int(qtd)
    cx, un = qtd // fator, qtd % fator
    if cx == 0 and un == 0:
        return "0"
    partes = []
    if cx > 0: partes.append(f"{cx} cx")
    if un > 0: partes.append(f"{un} un")
    return " + ".join(partes)


def formata_dif_em_cx(dif: int, familia: str) -> str:
    sinal = "+" if dif > 0 else ("-" if dif < 0 else "")
    return f"{sinal}{formata_un_em_cx(abs(int(dif)), familia)}" if dif != 0 else "0"


_PA_NORMALIZADO = {"TIANGUÁ": "Tianguá", "TIANGUA": "Tianguá", "GRANJA": "Granja", "SEDE": "Sede"}

df_mapa_pa_periodo = df_mapa_pa
if df_mapa_pa is not None and not df_mapa_pa.empty and "Data" in df_mapa_pa.columns:
    _dt_conc = pd.to_datetime(df_mapa_pa["Data"], dayfirst=True, errors="coerce")
    df_mapa_pa_periodo = df_mapa_pa[(_dt_conc >= pd.Timestamp(data_inicio)) & (_dt_conc <= pd.Timestamp(data_fim))]

MAPA_PA_CLASSIFICACAO: dict[str, str] = {}
MAPA_DATA_CONC: dict[str, str] = {}
if df_mapa_pa_periodo is not None and not df_mapa_pa_periodo.empty and "PA" in df_mapa_pa_periodo.columns:
    for _, _linha_conc in df_mapa_pa_periodo.iterrows():
        _mapa_resolvido = resolver_mapa(_linha_conc["Mapa"])
        _pa_bruto = str(_linha_conc["PA"]).strip().upper()
        MAPA_PA_CLASSIFICACAO[_mapa_resolvido] = _PA_NORMALIZADO.get(_pa_bruto, _linha_conc["PA"])
        if "Data" in _linha_conc:
            MAPA_DATA_CONC[_mapa_resolvido] = _linha_conc["Data"]

MAPAS_PA_CONC = {m for m, pa in MAPA_PA_CLASSIFICACAO.items() if pa in ("Tianguá", "Granja")}
MAPAS_SEDE_CONC = {m for m, pa in MAPA_PA_CLASSIFICACAO.items() if pa == "Sede"}

_periodo_str = f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
with st.sidebar:
    if df_mapa_pa is None or df_mapa_pa.empty:
        st.error(f"Não encontrei '{ARQUIVO_MAPA_PA.name}' no Google Drive nem histórico acumulado ainda.")
    else:
        st.success(
            f"{ARQUIVO_020501.name}: {len(df_020501_historico) if df_020501_historico is not None else 0} linha(s) acumuladas. "
            f"CONC.csv: {len(MAPA_PA_CLASSIFICACAO)} mapa(s) no período ({_periodo_str})."
        )

        mapas_no_020501 = set(df_020501_historico["Mapa"].unique()) if df_020501_historico is not None and not df_020501_historico.empty else set()
        mapas_conc_sem_relatorio = set(MAPA_PA_CLASSIFICACAO.keys()) - mapas_no_020501

        if mapas_conc_sem_relatorio:
            st.warning(f"⚠️ {len(mapas_conc_sem_relatorio)} mapa(s) do CONC.csv ainda não estão no 02.05.01.")
            with st.expander("Ver quais (agrupado por PA)"):
                mapas_faltantes_por_pa: dict[str, list[str]] = {}
                for m in mapas_conc_sem_relatorio:
                    mapas_faltantes_por_pa.setdefault(MAPA_PA_CLASSIFICACAO.get(m, "?"), []).append(m)
                for pa_nome, lista_mapas in sorted(mapas_faltantes_por_pa.items()):
                    lista_ordenada = sorted(lista_mapas, key=lambda x: int(str(x)) if str(x).isdigit() else 0)
                    st.markdown(f"**{pa_nome}** ({len(lista_ordenada)}):")
                    st.caption(", ".join(lista_ordenada))
        else:
            st.caption("✅ Todos os mapas do CONC.csv já estão no 02.05.01.")


def gerar_simulacao_perfeita(data_alvo) -> pd.DataFrame:
    if df_mapa_pa is None or df_mapa_pa.empty or df_020501_historico is None or df_020501_historico.empty:
        return pd.DataFrame()

    data_str = data_alvo.strftime("%d/%m/%Y")
    _pa_normalizada_sim = df_mapa_pa["PA"].apply(lambda v: _PA_NORMALIZADO.get(str(v).strip().upper(), v))
    mapas_pa_sim = df_mapa_pa[
        (df_mapa_pa["Data"] == data_str) & (_pa_normalizada_sim.isin(["Tianguá", "Granja"]))
    ][["Mapa", "PA"]].drop_duplicates().copy()
    if mapas_pa_sim.empty:
        return pd.DataFrame()

    mapas_pa_sim["MapaResolvido"] = mapas_pa_sim["Mapa"].apply(resolver_mapa)
    mapas_pa_sim["PA"] = mapas_pa_sim["PA"].apply(lambda v: _PA_NORMALIZADO.get(str(v).strip().upper(), v))
    pa_lookup_sim = mapas_pa_sim.groupby("MapaResolvido")["PA"].first().to_dict()

    familia_tipo_sim = df_020501_historico["Material"].apply(lambda c: familia_tipo_por_codigo(c, lookup_ag))
    df_sim = df_020501_historico.copy()
    df_sim["Familia"] = familia_tipo_sim.apply(lambda ft: normaliza_600(ft[0]))
    df_sim["Tipo"] = familia_tipo_sim.apply(lambda ft: ft[1])
    df_sim = df_sim[(df_sim["Familia"] != "Outro") & (df_sim["Tipo"] != "Garrafeira")]
    df_sim = df_sim[df_sim["Mapa"].isin(mapas_pa_sim["MapaResolvido"].unique())]

    saida_sim = df_sim.groupby(["Mapa", "Familia"])["Qtde_Saida"].sum().reset_index()
    saida_sim = saida_sim[saida_sim["Qtde_Saida"] > 0]

    linhas = []
    for _, r in saida_sim.iterrows():
        linhas.append({
            "Data": data_str,
            "PA": pa_lookup_sim.get(r["Mapa"], "Tianguá"),
            "Mapa": r["Mapa"],
            "Familia": r["Familia"],
            "Caixas": 0,
            "Garrafas": int(r["Qtde_Saida"]),
            "Garrafeiras": 0,
            "Unidades": 0,
        })
    return pd.DataFrame(linhas)


aba_vazio_pa, aba_conciliacao_sede, aba_historico = st.tabs(
    ["Vazio por PA", "Previsão Sede", "📈 Histórico Gerencial"]
)

def mapas_da_lote(mapas_str: str) -> list[str]:
    return [limpa_mapa(m) for m in str(mapas_str).split(";") if str(m).strip()]


_hist_vazio_pa_bruto = ler_aba_historico(NOME_ABA_SIMULACAO if modo_simulacao else "VazioPA")
ABA_VAZIO_PA_ATIVA = NOME_ABA_SIMULACAO if modo_simulacao else "VazioPA"
ABA_LOTE_ATIVA = NOME_ABA_LOTE_SIMULACAO if modo_simulacao else "VazioPALote"
if not _hist_vazio_pa_bruto.empty and "Mapa" in _hist_vazio_pa_bruto.columns:
    MAPAS_INDIVIDUAIS = set(_hist_vazio_pa_bruto["Mapa"].apply(limpar_numero_robusto).unique())
else:
    MAPAS_INDIVIDUAIS = set()

_hist_lote_bruto = ler_aba_historico(ABA_LOTE_ATIVA)
MAPAS_EM_LOTE = set()
if not _hist_lote_bruto.empty and "Mapas" in _hist_lote_bruto.columns:
    for _mapas_str in _hist_lote_bruto["Mapas"].unique():
        MAPAS_EM_LOTE.update(mapas_da_lote(_mapas_str))


# =========================================================================
# ABA VAZIO POR PA (conferência física digitada pelo conferente)
# =========================================================================
with aba_vazio_pa:
    st.caption("Conferência do vazio por PA e mapa.")

    # =====================================================================
    # 🚦 NOVO: PAINEL DE PENDÊNCIAS DE LANÇAMENTO (TIANGUÁ / GRANJA)
    # =====================================================================
    st.markdown("### 🚦 Painel de Pendências de Lançamento (PAs)")
    st.caption("Acompanhe quais mapas de PA da data abaixo já foram conferidos e quais ainda faltam.")
    
    col_dt_chk, col_info_chk = st.columns([1, 3])
    data_painel_chk = col_dt_chk.date_input("Data da Descarga / Fechamento", value=date.today(), key="data_painel_pendencias")
    data_painel_chk_str = data_painel_chk.strftime("%d/%m/%Y")

    mapas_conc_dia = {}
    if df_mapa_pa is not None and not df_mapa_pa.empty and "Data" in df_mapa_pa.columns and "PA" in df_mapa_pa.columns:
        sub_conc_painel = df_mapa_pa[df_mapa_pa["Data"] == data_painel_chk_str].copy()
        for _, r_c in sub_conc_painel.iterrows():
            pa_nome = _PA_NORMALIZADO.get(str(r_c["PA"]).strip().upper(), str(r_c["PA"]).strip())
            if pa_nome in ("Tianguá", "Granja"):
                mapas_conc_dia.setdefault(pa_nome, set()).add(limpar_numero_robusto(r_c["Mapa"]))

    mapas_lancados_dia = set()
    
    hist_lote_chk = ler_aba_historico(ABA_LOTE_ATIVA)
    if not hist_lote_chk.empty and "Data" in hist_lote_chk.columns and "Mapas" in hist_lote_chk.columns:
        sub_lote_chk = hist_lote_chk[hist_lote_chk["Data"] == data_painel_chk_str]
        for m_str in sub_lote_chk["Mapas"].unique():
            mapas_lancados_dia.update(mapas_da_lote(m_str))

    hist_indiv_chk = ler_aba_historico(ABA_VAZIO_PA_ATIVA)
    if not hist_indiv_chk.empty and "Data" in hist_indiv_chk.columns and "Mapa" in hist_indiv_chk.columns:
        sub_indiv_chk = hist_indiv_chk[hist_indiv_chk["Data"] == data_painel_chk_str]
        for m_indiv in sub_indiv_chk["Mapa"].unique():
            mapas_lancados_dia.add(limpar_numero_robusto(m_indiv))

    if not mapas_conc_dia:
        st.info(f"Nenhum mapa cadastrado para Tianguá ou Granja em {data_painel_chk_str} no CONC.csv.")
    else:
        total_esperado = sum(len(m_set) for m_set in mapas_conc_dia.values())
        todos_mapas_pa_dia = set.union(*mapas_conc_dia.values())
        total_lancados = len(todos_mapas_pa_dia & mapas_lancados_dia)
        total_pendentes = total_esperado - total_lancados

        renderizar_cards_resumo([
            ("Total de Mapas PA", total_esperado, "azul"),
            ("Mapas Lançados", total_lancados, "verde"),
            ("Mapas Pendentes", total_pendentes, "vermelho" if total_pendentes > 0 else "verde"),
        ])

        st.write("")
        linhas_pendencias = []
        cols_pa_chk = st.columns(len(mapas_conc_dia))
        for col_p, (nome_pa_chk, conjunto_mapas) in zip(cols_pa_chk, sorted(mapas_conc_dia.items())):
            mapas_ordenados = sorted(conjunto_mapas, key=lambda x: int(x) if str(x).isdigit() else 0)
            lancados_pa = [m for m in mapas_ordenados if m in mapas_lancados_dia]
            pendentes_pa = [m for m in mapas_ordenados if m not in mapas_lancados_dia]
            
            linhas_pendencias.append({
                "Data": data_painel_chk_str,
                "PA": nome_pa_chk,
                "Total Mapas": len(mapas_ordenados),
                "Lançados": len(lancados_pa),
                "Pendentes": len(pendentes_pa),
                "Lista Pendentes": ", ".join(pendentes_pa) if pendentes_pa else "Nenhum"
            })
            
            with col_p:
                st.markdown(f"**🚛 {nome_pa_chk}** ({len(lancados_pa)}/{len(mapas_ordenados)} lançados)")
                badges_html = []
                for m in lancados_pa:
                    badges_html.append(f'<span style="background:#EAF3DE; color:#173404; font-size:12px; font-weight:700; padding:3px 9px; border-radius:6px; margin:2px; display:inline-block;">✅ {m}</span>')
                for m in pendentes_pa:
                    badges_html.append(f'<span style="background:#FCEBEB; color:#501313; font-size:12px; font-weight:700; padding:3px 9px; border-radius:6px; margin:2px; display:inline-block;">⏳ {m}</span>')
                st.markdown(f'<div style="background:#F8F9FA; padding:10px; border-radius:8px; border:1px solid #E9ECEF;">{"".join(badges_html)}</div>', unsafe_allow_html=True)

        if linhas_pendencias:
            st.write("")
            if st.button("💾 Salvar Pendências no Histórico", key="btn_save_pend"):
                acumular_historico(pd.DataFrame(linhas_pendencias), "Snap_Pendencias", ["Data", "PA"])
                st.success("✅ Salvo na aba 'Snap_Pendencias' do histórico no Drive!")

    st.divider()
    st.markdown("### ✍️ Conferência Manual (um ou vários mapas)")
    st.caption("Digite um mapa só, ou vários separados por vírgula — o total informado abaixo vale pra soma de todos eles juntos.")

    col_data_manual, col_mapas_manual = st.columns([1, 2])
    data_manual = col_data_manual.date_input("Data da Descarga", value=data_painel_chk, key="data_vazio_manual")
    mapas_texto_manual = col_mapas_manual.text_input(
        "Números dos Mapas (separados por vírgula)",
        placeholder="ex: 257828, 257829, 257847",
        key="mapas_texto_manual",
    )

    mapas_manual_limpos = sorted(set(limpa_mapa(m) for m in mapas_texto_manual.split(",") if m.strip()))
    mapas_manual_resolvidos = resolver_mapas(mapas_manual_limpos) if mapas_manual_limpos else []
    pa_por_mapa_manual = {m: MAPA_PA_CLASSIFICACAO.get(m, "Desconhecido") for m in mapas_manual_resolvidos}
    pas_distintas_manual = sorted(set(pa_por_mapa_manual.values()))
    pa_combinado_manual = pas_distintas_manual[0] if len(pas_distintas_manual) == 1 else " + ".join(pas_distintas_manual)
    data_manual_str = data_manual.strftime("%d/%m/%Y")
    mapas_chave_manual = ";".join(mapas_manual_limpos)

    valores_existentes_manual = {}
    if mapas_chave_manual:
        hist_manual_atual = ler_aba_historico(ABA_LOTE_ATIVA)
        if not hist_manual_atual.empty:
            filtro_manual = (
                (hist_manual_atual["Data"] == data_manual_str)
                & (hist_manual_atual["PA"] == pa_combinado_manual)
                & (hist_manual_atual["Mapas"] == mapas_chave_manual)
            )
            for _, r in hist_manual_atual[filtro_manual].iterrows():
                valores_existentes_manual[r["Familia"]] = r

    if mapas_manual_limpos:
        st.caption(" · ".join(f"{m} → {pa_por_mapa_manual.get(m, '?')}" for m in mapas_manual_resolvidos))
        if "Desconhecido" in pas_distintas_manual:
            st.warning("Algum mapa digitado não está no CONC.csv — confira o número, ou ele fica classificado como 'Desconhecido'.")

        saida_esperada_manual = calcular_saida_familias(mapas_manual_resolvidos)
        if not valores_existentes_manual:
            if saida_esperada_manual:
                resumo_saida = ", ".join(f"{fam}: {formata_un_em_cx(qtd, fam)}" for fam, qtd in saida_esperada_manual.items())
                st.info(f"📤 Ainda não há retorno salvo pra esses mapas. Saída esperada: {resumo_saida}")
            else:
                st.warning("Nenhuma Saída encontrada pra esses mapas no 02.05.01 ainda.")
        else:
            st.markdown("**Status atual (com o que já está salvo):**")
            linhas_status = []
            for familia in FAMILIAS_CONFERENCIA:
                saida_fam = saida_esperada_manual.get(familia, 0)
                retorno_fam = int(valores_existentes_manual[familia]["Garrafas"]) if familia in valores_existentes_manual else 0
                dif = retorno_fam - saida_fam
                if saida_fam == 0 and retorno_fam == 0:
                    continue
                status_txt = "✅ Bateu" if dif == 0 else ("❌ Faltou" if dif < 0 else "⚠️ Sobrou")
                linhas_status.append({
                    "Item": rotulo_conferencia(familia),
                    "Saída": formata_un_em_cx(saida_fam, familia),
                    "Retorno": formata_un_em_cx(retorno_fam, familia),
                    "Diferença": formata_dif_em_cx(dif, familia),
                    "Status": status_txt,
                })
            if linhas_status:
                renderizar_tabela_limpa(pd.DataFrame(linhas_status), ["Item", "Saída", "Retorno", "Diferença", "Status"])

    with st.form("form_vazio_pa", clear_on_submit=True):
        if valores_existentes_manual:
            resumo_manual = ", ".join(
                f"{fam}: {int(r['Caixas'])} cx" if fam in REGRAS_VAZIO else f"{fam}: {int(r['Unidades'])} un"
                for fam, r in valores_existentes_manual.items()
            )
            st.info(f"📋 Total já salvo pra esses mapas — {resumo_manual}. Digite só a quantidade NOVA, o sistema soma sozinho.")

        st.markdown("**Caixas Físicas que Retornaram (total dos mapas acima)**")
        valores_familia_pa = {fam: st.number_input(rotulo_conferencia(fam), min_value=0, step=1, key=f"cx_pa_{fam}") for fam in FAMILIAS_CONFERENCIA}

        st.markdown("**Outros AG (sem conversão — já em unidade final)**")
        c1, c2, c3, c4, c5 = st.columns(5)
        chapatex_pa = c1.number_input("Chapatex (Und)", min_value=0, step=1, key="outros_pa_chapatex")
        pbr1_pa = c2.number_input("Pallet PBR1", min_value=0, step=1, key="outros_pa_pbr1")
        pbr2_pa = c3.number_input("Pallet PBR2", min_value=0, step=1, key="outros_pa_pbr2")
        barril30_pa = c4.number_input("Barril 30L", min_value=0, step=1, key="outros_pa_barril30")
        barril50_pa = c5.number_input("Barril 50L", min_value=0, step=1, key="outros_pa_barril50")

        if st.form_submit_button("Salvar conferência"):
            if not mapas_manual_limpos:
                st.error("Informe pelo menos um número de mapa antes de salvar.")
            else:
                data_str_pa = data_manual_str
                gf_600 = valores_familia_pa.get("600ml", 0) + valores_familia_pa.get("Verde 600", 0)
                linhas_pa = []

                for familia, qtd_cx_nova in valores_familia_pa.items():
                    caixas_existentes = int(valores_existentes_manual[familia]["Caixas"]) if familia in valores_existentes_manual else 0
                    qtd_cx = caixas_existentes + qtd_cx_nova
                    if qtd_cx > 0:
                        r = REGRAS_VAZIO[familia]
                        gf = gf_600 if familia == "600ml" else (0 if familia == "Verde 600" else qtd_cx * r["garrafeiras_por_cx"])
                        linhas_pa.append({
                            "Data": data_str_pa,
                            "PA": pa_combinado_manual,
                            "Mapas": mapas_chave_manual,
                            "Familia": familia,
                            "Caixas": qtd_cx,
                            "Garrafas": qtd_cx * r["garrafas_por_cx"],
                            "Garrafeiras": gf,
                            "Unidades": 0,
                        })

                for familia_outros, qtd_un_nova in [
                    ("Chapatex", chapatex_pa), ("Pallet PBR1", pbr1_pa), ("Pallet PBR2", pbr2_pa),
                    ("Barril 30L", barril30_pa), ("Barril 50L", barril50_pa),
                ]:
                    unidades_existentes = int(valores_existentes_manual[familia_outros]["Unidades"]) if familia_outros in valores_existentes_manual else 0
                    qtd_un = unidades_existentes + qtd_un_nova
                    if qtd_un > 0:
                        linhas_pa.append({
                            "Data": data_str_pa,
                            "PA": pa_combinado_manual,
                            "Mapas": mapas_chave_manual,
                            "Familia": familia_outros,
                            "Caixas": 0,
                            "Garrafas": 0,
                            "Garrafeiras": 0,
                            "Unidades": qtd_un,
                        })

                if linhas_pa:
                    acumular_historico(pd.DataFrame(linhas_pa), ABA_LOTE_ATIVA, ["Data", "PA", "Mapas", "Familia"])
                    st.success(f"✅ Somado ao total de {len(mapas_manual_limpos)} mapa(s) ({', '.join(mapas_manual_limpos)})!")
                    st.rerun()
                else:
                    st.warning("Nenhuma quantidade foi informada para salvar.")

    st.divider()
    st.markdown("### 📦 Conferência em Lote (vários mapas conferidos juntos)")
    st.caption("Use quando só souber o TOTAL, sem separar por mapa.")

    col_data_l, col_pa_l = st.columns(2)
    data_lote = col_data_l.date_input("Data da Descarga", value=data_painel_chk, key="data_lote")
    pa_lote = col_pa_l.selectbox("PA", ["Tianguá", "Granja"], key="pa_lote")

    mapas_lote_auto = buscar_mapas_por_data_pa(data_lote, pa_lote)
    data_lote_str_atual = data_lote.strftime("%d/%m/%Y")
    mapas_chave_atual = ";".join(mapas_lote_auto) if mapas_lote_auto else ""

    valores_existentes_lote = {}
    if mapas_chave_atual:
        hist_lote_atual = ler_aba_historico(ABA_LOTE_ATIVA)
        if not hist_lote_atual.empty:
            filtro_lote_atual = (
                (hist_lote_atual["Data"] == data_lote_str_atual)
                & (hist_lote_atual["PA"] == pa_lote)
                & (hist_lote_atual["Mapas"] == mapas_chave_atual)
            )
            for _, r in hist_lote_atual[filtro_lote_atual].iterrows():
                valores_existentes_lote[r["Familia"]] = r

    if df_mapa_pa is None or df_mapa_pa.empty:
        st.error(f"Não encontrei '{ARQUIVO_MAPA_PA.name}' no Google Drive nem histórico acumulado ainda.")
    elif mapas_lote_auto:
        st.success(f"{len(mapas_lote_auto)} mapa(s) de {pa_lote} em {data_lote.strftime('%d/%m/%Y')}: {', '.join(mapas_lote_auto)}")

        mapas_lote_resolvidos_status = resolver_mapas(mapas_lote_auto)
        saida_esperada_lote = calcular_saida_familias(mapas_lote_resolvidos_status)
        if not valores_existentes_lote:
            if saida_esperada_lote:
                resumo_saida_lote = ", ".join(f"{fam}: {formata_un_em_cx(qtd, fam)}" for fam, qtd in saida_esperada_lote.items())
                st.info(f"📤 Ainda não há retorno salvo pra esse lote. Saída esperada: {resumo_saida_lote}")
        else:
            st.markdown("**Status atual (com o que já está salvo):**")
            linhas_status_lote = []
            for familia in FAMILIAS_CONFERENCIA:
                saida_fam = saida_esperada_lote.get(familia, 0)
                retorno_fam = int(valores_existentes_lote[familia]["Garrafas"]) if familia in valores_existentes_lote else 0
                dif = retorno_fam - saida_fam
                if saida_fam == 0 and retorno_fam == 0:
                    continue
                status_txt = "✅ Bateu" if dif == 0 else ("❌ Faltou" if dif < 0 else "⚠️ Sobrou")
                linhas_status_lote.append({
                    "Item": rotulo_conferencia(familia),
                    "Saída": formata_un_em_cx(saida_fam, familia),
                    "Retorno": formata_un_em_cx(retorno_fam, familia),
                    "Diferença": formata_dif_em_cx(dif, familia),
                    "Status": status_txt,
                })
            if linhas_status_lote:
                renderizar_tabela_limpa(pd.DataFrame(linhas_status_lote), ["Item", "Saída", "Retorno", "Diferença", "Status"])
    else:
        st.warning(f"Nenhum mapa cadastrado pra {pa_lote} em {data_lote.strftime('%d/%m/%Y')} na planilha 'Mapa PA'.")

    with st.form("form_vazio_pa_lote", clear_on_submit=True):
        if valores_existentes_lote:
            resumo_atual = ", ".join(
                f"{fam}: {int(r['Caixas'])} cx" if fam in REGRAS_VAZIO else f"{fam}: {int(r['Unidades'])} un"
                for fam, r in valores_existentes_lote.items()
            )
            st.info(
                f"📋 Total já salvo pra esse PA/Data — {resumo_atual}. "
                "Digite abaixo só a quantidade NOVA que está lançando agora — o sistema soma automaticamente."
            )

        st.markdown("**Caixas Físicas que Retornaram (quantidade NOVA, não o total)**")
        valores_familia_lote = {
            fam: st.number_input(rotulo_conferencia(fam), min_value=0, step=1, key=f"cx_lote_{fam}")
            for fam in FAMILIAS_CONFERENCIA
        }

        st.markdown("**Outros AG (quantidade NOVA, sem conversão)**")
        cl1, cl2, cl3, cl4, cl5 = st.columns(5)
        chapatex_lote = cl1.number_input("Chapatex (Und)", min_value=0, step=1, key="outros_lote_chapatex")
        pbr1_lote = cl2.number_input("Pallet PBR1", min_value=0, step=1, key="outros_lote_pbr1")
        pbr2_lote = cl3.number_input("Pallet PBR2", min_value=0, step=1, key="outros_lote_pbr2")
        barril30_lote = cl4.number_input("Barril 30L", min_value=0, step=1, key="outros_lote_barril30")
        barril50_lote = cl5.number_input("Barril 50L", min_value=0, step=1, key="outros_lote_barril50")

        if st.form_submit_button("Salvar conferência em lote"):
            if len(mapas_lote_auto) < 2:
                st.error("Não há pelo menos 2 mapas cadastrados pra essa Data/PA na planilha 'Mapa PA' — confira se ela está atualizada.")
            else:
                mapas_chave = ";".join(mapas_lote_auto)
                data_str_lote = data_lote.strftime("%d/%m/%Y")
                gf_600_lote = valores_familia_lote.get("600ml", 0) + valores_familia_lote.get("Verde 600", 0)
                linhas_lote = []

                for familia, qtd_cx_nova in valores_familia_lote.items():
                    caixas_existentes = int(valores_existentes_lote[familia]["Caixas"]) if familia in valores_existentes_lote else 0
                    qtd_cx = caixas_existentes + qtd_cx_nova
                    if qtd_cx > 0:
                        r = REGRAS_VAZIO[familia]
                        gf = gf_600_lote if familia == "600ml" else (0 if familia == "Verde 600" else qtd_cx * r["garrafeiras_por_cx"])
                        linhas_lote.append({
                            "Data": data_str_lote, "PA": pa_lote, "Mapas": mapas_chave, "Familia": familia,
                            "Caixas": qtd_cx, "Garrafas": qtd_cx * r["garrafas_por_cx"], "Garrafeiras": gf, "Unidades": 0,
                        })

                for familia_outros, qtd_un_nova in [
                    ("Chapatex", chapatex_lote), ("Pallet PBR1", pbr1_lote), ("Pallet PBR2", pbr2_lote),
                    ("Barril 30L", barril30_lote), ("Barril 50L", barril50_lote),
                ]:
                    unidades_existentes = int(valores_existentes_lote[familia_outros]["Unidades"]) if familia_outros in valores_existentes_lote else 0
                    qtd_un = unidades_existentes + qtd_un_nova
                    if qtd_un > 0:
                        linhas_lote.append({
                            "Data": data_str_lote, "PA": pa_lote, "Mapas": mapas_chave, "Familia": familia_outros,
                            "Caixas": 0, "Garrafas": 0, "Garrafeiras": 0, "Unidades": qtd_un,
                        })

                if linhas_lote:
                    acumular_historico(pd.DataFrame(linhas_lote), ABA_LOTE_ATIVA, ["Data", "PA", "Mapas", "Familia"])
                    st.success(f"✅ Somado ao total de {len(mapas_lote_auto)} mapas ({', '.join(mapas_lote_auto)})!")
                    st.rerun()
                else:
                    st.warning("Nenhuma quantidade foi informada para salvar.")

    # =====================================================================
    # 📋 RESUMO ACUMULADO DO DIA COM JUSTIFICATIVA (ST.DATA_EDITOR)
    # =====================================================================
    st.divider()
    st.markdown("### 📋 Resumo Acumulado do Dia (Saída vs Retorno)")
    st.caption("Acompanhe o status das conferências da data abaixo. Digite a justificativa direto na tabela e clique em Salvar.")
    
    col_res1, col_res2 = st.columns([1, 3])
    data_resumo = col_res1.date_input("Data do Resumo", value=data_painel_chk, key="data_resumo_diario")
    data_resumo_str = data_resumo.strftime("%d/%m/%Y")

    linhas_resumo_diario = []
    
    # Prepara dicionário de justificativas existentes para preencher a tabela
    hist_justificativas = ler_aba_historico("Justificativas")
    dict_justif = {}
    if not hist_justificativas.empty:
        for _, r in hist_justificativas.iterrows():
            dict_justif[(str(r.get("Data", "")), str(r.get("Mapas/Lote", "")), str(r.get("Item", "")))] = str(r.get("Justificativa", ""))

    hist_lote_resumo = ler_aba_historico(ABA_LOTE_ATIVA)
    if not hist_lote_resumo.empty and "Data" in hist_lote_resumo.columns:
        hist_lote_dia = hist_lote_resumo[hist_lote_resumo["Data"] == data_resumo_str]
        for (pa_r, mapas_r), group in hist_lote_dia.groupby(["PA", "Mapas"]):
            mapas_limpos = mapas_da_lote(mapas_r)
            mapas_resolvidos = resolver_mapas(mapas_limpos)
            saida_esperada = calcular_saida_familias(mapas_resolvidos)
            
            for familia in FAMILIAS_CONFERENCIA:
                saida_fam = saida_esperada.get(familia, 0)
                linha_fam = group[group["Familia"] == familia]
                if not linha_fam.empty:
                    retorno_fam = int(pd.to_numeric(linha_fam["Garrafas"]).sum()) + int(pd.to_numeric(linha_fam["Unidades"]).sum())
                else:
                    retorno_fam = 0
                
                if saida_fam == 0 and retorno_fam == 0:
                    continue
                    
                dif = retorno_fam - saida_fam
                status_txt = "✅ Bateu" if dif == 0 else ("❌ Faltou" if dif < 0 else "⚠️ Sobrou")
                
                chave_justif = (data_resumo_str, str(mapas_r), rotulo_conferencia(familia))
                
                linhas_resumo_diario.append({
                    "PA": pa_r,
                    "Mapas/Lote": str(mapas_r).replace(";", ", "),
                    "Item": rotulo_conferencia(familia),
                    "Saída": formata_un_em_cx(saida_fam, familia),
                    "Retorno": formata_un_em_cx(retorno_fam, familia),
                    "Diferença": formata_dif_em_cx(dif, familia),
                    "Status": status_txt,
                    "Justificativa": dict_justif.get(chave_justif, "")
                })

    hist_indiv_resumo = ler_aba_historico(ABA_VAZIO_PA_ATIVA)
    if not hist_indiv_resumo.empty and "Data" in hist_indiv_resumo.columns:
        hist_indiv_dia = hist_indiv_resumo[hist_indiv_resumo["Data"] == data_resumo_str]
        for (pa_r, mapa_r), group in hist_indiv_dia.groupby(["PA", "Mapa"]):
            mapas_limpos = [limpa_mapa(mapa_r)]
            mapas_resolvidos = resolver_mapas(mapas_limpos)
            saida_esperada = calcular_saida_familias(mapas_resolvidos)
            
            for familia in FAMILIAS_CONFERENCIA:
                saida_fam = saida_esperada.get(familia, 0)
                linha_fam = group[group["Familia"] == familia]
                if not linha_fam.empty:
                    retorno_fam = int(pd.to_numeric(linha_fam["Garrafas"]).sum()) + int(pd.to_numeric(linha_fam["Unidades"]).sum())
                else:
                    retorno_fam = 0
                
                if saida_fam == 0 and retorno_fam == 0:
                    continue
                    
                dif = retorno_fam - saida_fam
                status_txt = "✅ Bateu" if dif == 0 else ("❌ Faltou" if dif < 0 else "⚠️ Sobrou")
                
                ja_existe = any(str(x["Mapas/Lote"]) == str(mapa_r) and x["Item"] == rotulo_conferencia(familia) for x in linhas_resumo_diario)
                if not ja_existe:
                    chave_justif = (data_resumo_str, str(mapa_r), rotulo_conferencia(familia))
                    linhas_resumo_diario.append({
                        "PA": pa_r,
                        "Mapas/Lote": str(mapa_r),
                        "Item": rotulo_conferencia(familia),
                        "Saída": formata_un_em_cx(saida_fam, familia),
                        "Retorno": formata_un_em_cx(retorno_fam, familia),
                        "Diferença": formata_dif_em_cx(dif, familia),
                        "Status": status_txt,
                        "Justificativa": dict_justif.get(chave_justif, "")
                    })

    if linhas_resumo_diario:
        df_resumo = pd.DataFrame(linhas_resumo_diario)
        df_resumo = df_resumo.sort_values(by=["PA", "Mapas/Lote", "Item"])
        
        edited_resumo = st.data_editor(
            df_resumo,
            use_container_width=True,
            hide_index=True,
            disabled=["PA", "Mapas/Lote", "Item", "Saída", "Retorno", "Diferença", "Status"],
            column_config={
                "Justificativa": st.column_config.TextColumn("📝 Justificativa", help="Clique aqui para digitar", max_chars=250)
            }
        )
        
        st.write("")
        if st.button("💾 Salvar Resumo e Justificativas no Histórico", key="btn_save_resumo"):
            # Salvar Justificativas
            df_justif_save = edited_resumo[["Mapas/Lote", "Item", "Justificativa"]].copy()
            df_justif_save.insert(0, "Data", data_resumo_str)
            acumular_historico(df_justif_save, "Justificativas", ["Data", "Mapas/Lote", "Item"])
            
            # Salvar Snapshot Resumo
            df_resumo_save = edited_resumo.copy()
            df_resumo_save.insert(0, "Data", data_resumo_str)
            acumular_historico(df_resumo_save, "Snap_ResumoAcumulado", ["Data", "Mapas/Lote", "Item"])
            st.success("✅ Resumo e Justificativas salvos nas abas 'Snap_ResumoAcumulado' e 'Justificativas' do Drive!")
    else:
        st.info(f"Nenhuma conferência encontrada para a data {data_resumo_str}.")

    # =====================================================================
    # EDIÇÃO E EXCLUSÃO
    # =====================================================================
    df_vazio_pa = ler_aba_historico(ABA_VAZIO_PA_ATIVA)
    if not df_vazio_pa.empty:
        st.divider()
        st.markdown("### 🚚 Detalhe de Retorno por Mapa")

        if "Unidades" not in df_vazio_pa.columns:
            df_vazio_pa["Unidades"] = 0

        colunas_exibicao = ["Data", "PA", "Mapa", "Familia", "Caixas", "Garrafas", "Garrafeiras", "Unidades"]
        df_exibicao = df_vazio_pa[colunas_exibicao].copy()

        for col in ["Caixas", "Garrafas", "Garrafeiras", "Unidades"]:
            df_exibicao[col] = pd.to_numeric(df_exibicao[col], errors='coerce').fillna(0).astype(int)

        df_exibicao["Data_Sort"] = pd.to_datetime(df_exibicao["Data"], format="%d/%m/%Y", errors="coerce")
        df_exibicao = df_exibicao.sort_values(by=["Data_Sort", "Mapa"], ascending=[False, True]).drop(columns=["Data_Sort"])

        st.dataframe(df_exibicao, width='stretch', hide_index=True)

        st.markdown("### 📊 Resumo Físico por PA e Família")
        resumo_pa_familia = df_vazio_pa.groupby(["PA", "Familia"])[["Caixas", "Garrafas", "Garrafeiras", "Unidades"]].sum().reset_index()
        resumo_pa_familia[["Caixas", "Garrafas", "Garrafeiras", "Unidades"]] = resumo_pa_familia[["Caixas", "Garrafas", "Garrafeiras", "Unidades"]].astype(int)
        st.dataframe(resumo_pa_familia, width='stretch', hide_index=True)

        st.divider()
        st.markdown("### 🛠️ Editar ou Apagar Registros")

        with st.expander("✏️ Editar um item específico (sem apagar o mapa inteiro)", expanded=False):
            c_data_e, c_mapa_e = st.columns(2)
            datas_existentes_e = sorted(df_vazio_pa["Data"].unique(), key=lambda d: pd.to_datetime(d, dayfirst=True), reverse=True)
            edit_data = c_data_e.selectbox("Data da descarga", datas_existentes_e, key="edit_data_pa")

            mapas_na_data_e = sorted(df_vazio_pa[df_vazio_pa["Data"] == edit_data]["Mapa"].astype(str).unique())
            edit_mapa = c_mapa_e.selectbox("Número do Mapa", mapas_na_data_e, key="edit_mapa_pa")

            df_mapa_editar = df_vazio_pa[(df_vazio_pa["Data"] == edit_data) & (df_vazio_pa["Mapa"].astype(str) == edit_mapa)]

            if df_mapa_editar.empty:
                st.info("Nenhum item encontrado pra esse mapa/data.")
            else:
                TODAS_FAMILIAS = list(REGRAS_VAZIO.keys()) + ["Chapatex", "Pallet PBR1", "Pallet PBR2", "Barril 30L", "Barril 50L"]
                familias_ja_digitadas = set(df_mapa_editar["Familia"].unique())
                familias_disponiveis = sorted(TODAS_FAMILIAS, key=lambda f: (f not in familias_ja_digitadas, f))
                edit_familia = st.selectbox("Item (Família) para editar ou adicionar", familias_disponiveis, key="edit_familia_pa")

                linhas_familia = df_mapa_editar[df_mapa_editar["Familia"] == edit_familia]
                pa_padrao = df_mapa_editar["PA"].iloc[0]
                if not linhas_familia.empty:
                    linha_atual = linhas_familia.iloc[0]
                    pa_atual = linha_atual["PA"]
                    st.caption(f"Mapa {edit_mapa} · {pa_atual} · {edit_data} · {edit_familia}")
                else:
                    linha_atual = {}
                    pa_atual = pa_padrao
                    st.caption(f"Novo: Mapa {edit_mapa} · {pa_atual} · {edit_data} · {edit_familia}")

                st.caption("💡 **Atenção:** Para bebidas (300ml, 600ml, 1L), altere apenas as **Caixas**. As garrafas serão recalculadas sozinhas ao salvar.")
                ce1, ce2, ce3, ce4 = st.columns(4)
                novo_caixas = ce1.number_input("Caixas", min_value=0, step=1, value=int(linha_atual.get("Caixas", 0)) if isinstance(linha_atual, pd.Series) else 0, key="edit_caixas")
                novo_garrafas = ce2.number_input("Garrafas", min_value=0, step=1, value=int(linha_atual.get("Garrafas", 0)) if isinstance(linha_atual, pd.Series) else 0, key="edit_garrafas")
                novo_garrafeiras = ce3.number_input("Garrafeiras", min_value=0, step=1, value=int(linha_atual.get("Garrafeiras", 0)) if isinstance(linha_atual, pd.Series) else 0, key="edit_garrafeiras")
                novo_unidades = ce4.number_input("Unidades", min_value=0, step=1, value=int(linha_atual.get("Unidades", 0)) if isinstance(linha_atual, pd.Series) else 0, key="edit_unidades")

                if st.button("💾 Salvar edição", type="primary", key="salvar_edicao_pa"):
                    if edit_familia in REGRAS_VAZIO:
                        r = REGRAS_VAZIO[edit_familia]
                        cx_salvar = novo_caixas
                        gf_salvar = novo_caixas * r["garrafas_por_cx"]
                        gfr_salvar = 0 if edit_familia == "Verde 600" else (novo_caixas * r["garrafeiras_por_cx"])
                        un_salvar = 0
                    else:
                        cx_salvar = 0
                        gf_salvar = 0
                        gfr_salvar = 0
                        un_salvar = novo_unidades

                    nova_linha = pd.DataFrame([{
                        "Data": edit_data, "PA": pa_atual, "Mapa": edit_mapa, "Familia": edit_familia,
                        "Caixas": cx_salvar, "Garrafas": gf_salvar, "Garrafeiras": gfr_salvar, "Unidades": un_salvar,
                    }])
                    acumular_historico(nova_linha, ABA_VAZIO_PA_ATIVA, ["Data", "PA", "Mapa", "Familia"])
                    st.success(f"✅ Item '{edit_familia}' do mapa {edit_mapa} atualizado!")
                    st.rerun()

        with st.expander("🗑️ Selecionar Mapa para Exclusão", expanded=False):
            c_data, c_mapa, c_btn = st.columns([1, 1, 1])
            datas_existentes = sorted(df_vazio_pa["Data"].unique(), key=lambda d: pd.to_datetime(d, dayfirst=True), reverse=True)
            del_data = c_data.selectbox("Data da descarga", datas_existentes, key="del_data_pa")

            mapas_na_data = sorted(df_vazio_pa[df_vazio_pa["Data"] == del_data]["Mapa"].astype(str).unique())
            del_mapa = c_mapa.selectbox("Número do Mapa", mapas_na_data, key="del_mapa_pa")

            c_btn.write("")
            c_btn.write("")
            if c_btn.button("🗑️ Apagar este Mapa", type="primary", use_container_width=True):
                d_alvo = str(del_data).strip()
                m_alvo = str(del_mapa).strip()
                
                mascara_excluir = df_vazio_pa.apply(
                    lambda r: str(r["Data"]).strip() == d_alvo and str(r["Mapa"]).strip() == m_alvo, 
                    axis=1
                )
                
                df_restante = df_vazio_pa[~mascara_excluir]
                salvar_aba_historico(ABA_VAZIO_PA_ATIVA, df_restante)
                
                time.sleep(1)
                st.rerun()

    df_vazio_lote = ler_aba_historico(ABA_LOTE_ATIVA)
    if not df_vazio_lote.empty:
        st.divider()
        st.markdown("#### 📦 Lotes conferidos")
        df_lote_exib = df_vazio_lote.copy()
        for col in ["Caixas", "Garrafas", "Garrafeiras", "Unidades"]:
            if col in df_lote_exib.columns:
                df_lote_exib[col] = pd.to_numeric(df_lote_exib[col], errors="coerce").fillna(0).astype(int)
        df_lote_exib["Mapas"] = df_lote_exib["Mapas"].apply(lambda m: ", ".join(mapas_da_lote(m)))
        st.dataframe(
            df_lote_exib[["Data", "PA", "Mapas", "Familia", "Caixas", "Garrafas", "Garrafeiras", "Unidades"]],
            width='stretch', hide_index=True,
        )

        with st.expander("✏️ Editar um item do lote (sem apagar o lote inteiro)", expanded=False):
            cel1, cel2 = st.columns(2)
            datas_lote_edit = sorted(df_vazio_lote["Data"].unique(), key=lambda d: pd.to_datetime(d, dayfirst=True), reverse=True)
            edit_data_lote = cel1.selectbox("Data", datas_lote_edit, key="edit_data_lote")

            lotes_na_data_edit = sorted(df_vazio_lote[df_vazio_lote["Data"] == edit_data_lote]["Mapas"].unique())
            edit_lote_chave = cel2.selectbox(
                "Lote", lotes_na_data_edit, format_func=lambda m: ", ".join(mapas_da_lote(m)), key="edit_lote_chave"
            )

            df_lote_editar = df_vazio_lote[(df_vazio_lote["Data"] == edit_data_lote) & (df_vazio_lote["Mapas"] == edit_lote_chave)]

            if df_lote_editar.empty:
                st.info("Nenhum item encontrado pra esse lote/data.")
            else:
                TODAS_FAMILIAS_LOTE = list(REGRAS_VAZIO.keys()) + ["Chapatex", "Pallet PBR1", "Pallet PBR2", "Barril 30L", "Barril 50L"]
                familias_ja_no_lote = set(df_lote_editar["Familia"].unique())
                familias_disp_lote = sorted(TODAS_FAMILIAS_LOTE, key=lambda f: (f not in familias_ja_no_lote, f))
                edit_familia_lote = st.selectbox("Item (Família) para editar ou adicionar", familias_disp_lote, key="edit_familia_lote")

                linhas_familia_lote = df_lote_editar[df_lote_editar["Familia"] == edit_familia_lote]
                pa_padrao_lote = df_lote_editar["PA"].iloc[0]
                if not linhas_familia_lote.empty:
                    linha_atual_lote = linhas_familia_lote.iloc[0]
                    pa_atual_lote = linha_atual_lote["PA"]
                    st.caption(f"Lote {', '.join(mapas_da_lote(edit_lote_chave))} · {pa_atual_lote} · {edit_data_lote} · {edit_familia_lote}")
                else:
                    linha_atual_lote = {}
                    pa_atual_lote = pa_padrao_lote
                    st.caption(f"Novo: Lote {', '.join(mapas_da_lote(edit_lote_chave))} · {pa_atual_lote} · {edit_data_lote} · {edit_familia_lote}")

                st.caption("💡 **Atenção:** Para bebidas (300ml, 600ml, 1L), altere apenas as **Caixas**. As garrafas serão recalculadas sozinhas ao salvar.")
                cel3, cel4, cel5, cel6 = st.columns(4)
                novo_caixas_lote = cel3.number_input("Caixas", min_value=0, step=1, value=int(linha_atual_lote.get("Caixas", 0)) if isinstance(linha_atual_lote, pd.Series) else 0, key="edit_caixas_lote")
                novo_garrafas_lote = cel4.number_input("Garrafas", min_value=0, step=1, value=int(linha_atual_lote.get("Garrafas", 0)) if isinstance(linha_atual_lote, pd.Series) else 0, key="edit_garrafas_lote")
                novo_garrafeiras_lote = cel5.number_input("Garrafeiras", min_value=0, step=1, value=int(linha_atual_lote.get("Garrafeiras", 0)) if isinstance(linha_atual_lote, pd.Series) else 0, key="edit_garrafeiras_lote")
                novo_unidades_lote = cel6.number_input("Unidades", min_value=0, step=1, value=int(linha_atual_lote.get("Unidades", 0)) if isinstance(linha_atual_lote, pd.Series) else 0, key="edit_unidades_lote")

                if st.button("💾 Salvar edição do lote", type="primary", key="salvar_edicao_lote"):
                    if edit_familia_lote in REGRAS_VAZIO:
                        r = REGRAS_VAZIO[edit_familia_lote]
                        cx_salvar = novo_caixas_lote
                        gf_salvar = novo_caixas_lote * r["garrafas_por_cx"]
                        gfr_salvar = 0 if edit_familia_lote == "Verde 600" else (novo_caixas_lote * r["garrafeiras_por_cx"])
                        un_salvar = 0
                    else:
                        cx_salvar = 0
                        gf_salvar = 0
                        gfr_salvar = 0
                        un_salvar = novo_unidades_lote

                    nova_linha_lote = pd.DataFrame([{
                        "Data": edit_data_lote, "PA": pa_atual_lote, "Mapas": edit_lote_chave, "Familia": edit_familia_lote,
                        "Caixas": cx_salvar, "Garrafas": gf_salvar, "Garrafeiras": gfr_salvar, "Unidades": un_salvar,
                    }])
                    acumular_historico(nova_linha_lote, ABA_LOTE_ATIVA, ["Data", "PA", "Mapas", "Familia"])
                    st.success(f"✅ Item '{edit_familia_lote}' do lote atualizado!")
                    st.rerun()

        with st.expander("🗑️ Apagar um lote", expanded=False):
            cdl1, cdl2, cdl3 = st.columns([1, 2, 1])
            datas_lote_exist = sorted(df_vazio_lote["Data"].unique(), key=lambda d: pd.to_datetime(d, dayfirst=True), reverse=True)
            del_data_lote = cdl1.selectbox("Data", datas_lote_exist, key="del_data_lote")
            lotes_na_data = sorted(df_vazio_lote[df_vazio_lote["Data"] == del_data_lote]["Mapas"].unique())
            del_lote_chave = cdl2.selectbox(
                "Lote", lotes_na_data, format_func=lambda m: ", ".join(mapas_da_lote(m)), key="del_lote_chave"
            )
            cdl3.write("")
            cdl3.write("")
            if cdl3.button("🗑️ Apagar", type="primary", use_container_width=True, key="btn_del_lote"):
                d_alvo = str(del_data_lote).strip()
                m_alvo = str(del_lote_chave).strip()
                
                mascara_excluir = df_vazio_lote.apply(
                    lambda r: str(r["Data"]).strip() == d_alvo and str(r["Mapas"]).strip() == m_alvo, 
                    axis=1
                )
                
                df_restante_lote = df_vazio_lote[~mascara_excluir]
                salvar_aba_historico(ABA_LOTE_ATIVA, df_restante_lote)
                
                time.sleep(1) 
                st.rerun()

    st.divider()
    with st.expander("🧹 Limpar mapa 'fantasma' do histórico do CONC", expanded=False):
        st.caption(
            "O app acumula todo mapa já visto no CONC.csv, pra Previsão de Contagem "
            "conseguir consultar dias anteriores mesmo depois do arquivo de hoje substituir o de ontem. "
            "Se um mapa veio de um teste/erro de digitação corrigido depois, remova aqui."
        )
        _hist_conc_bruto_limpeza = ler_aba_historico("MapaPAHistorico")
        if _hist_conc_bruto_limpeza.empty or "Mapa" not in _hist_conc_bruto_limpeza.columns:
            st.info("Nenhum histórico de CONC registrado ainda.")
        else:
            _hist_conc_bruto_limpeza["Mapa"] = _hist_conc_bruto_limpeza["Mapa"].apply(limpar_numero_robusto)
            mapas_disponiveis_limpeza = sorted(
                _hist_conc_bruto_limpeza["Mapa"].dropna().unique().tolist(),
                key=lambda m: int(m) if str(m).isdigit() else 0,
            )
            mapas_pra_remover = st.multiselect(
                "Mapas a apagar do registro histórico (não afeta o VazioPA/VazioPALote nem a conciliação)",
                mapas_disponiveis_limpeza,
                key="mapas_pra_remover_conc_hist",
            )
            if st.button("🗑️ Remover selecionados do registro", disabled=not mapas_pra_remover):
                hist_conc_restante = _hist_conc_bruto_limpeza[~_hist_conc_bruto_limpeza["Mapa"].isin(mapas_pra_remover)]
                salvar_aba_historico("MapaPAHistorico", hist_conc_restante)
                st.success(f"{len(mapas_pra_remover)} mapa(s) removido(s) do registro.")
                st.rerun()

    st.divider()
    with st.expander("🧪 Modo Simulação (dados de teste)", expanded=False):
        st.caption("Gera retorno = saída perfeita pra todos os mapas Tianguá/Granja de uma data (via CONC.csv). Grava numa aba separada, nunca mistura com produção.")
        data_simulacao = st.date_input("Data pra simular", value=date.today() - timedelta(days=1), key="data_gerar_simulacao")

        col_sim1, col_sim2 = st.columns(2)
        if col_sim1.button("🧪 Gerar simulação", use_container_width=True):
            df_simulado = gerar_simulacao_perfeita(data_simulacao)
            if df_simulado.empty:
                st.warning("Nenhum mapa Tianguá/Granja encontrado pra essa data (confira o CONC.csv e o 02.05.01).")
            else:
                salvar_aba_historico(NOME_ABA_SIMULACAO, df_simulado)
                st.success(f"{len(df_simulado)} linha(s) simuladas geradas.")
                st.rerun()

        if col_sim2.button("🗑️ Apagar simulação", use_container_width=True):
            salvar_aba_historico(NOME_ABA_SIMULACAO, pd.DataFrame(columns=["Data", "PA", "Mapa", "Familia", "Caixas", "Garrafas", "Garrafeiras", "Unidades"]))
            st.success("Dados de simulação apagados.")
            st.rerun()


# =========================================================================
# ABA DE PREVISÃO DE CONTAGEM (TOTAL, SEDE E PAs)
# =========================================================================
with aba_conciliacao_sede:
    st.header("📅 Previsão de Contagem do AG")
    st.caption("Quanto deveria voltar vazio, baseado no que saiu (Geral, Sede e por PA).")
    df_concil_sede = pd.DataFrame()  # fallback

    data_previsao = st.date_input(
        "Data em que a rota saiu",
        value=date.today() - timedelta(days=1),
        key="data_previsao_contagem",
    )

    if df_mapa_pa is None or df_mapa_pa.empty:
        st.error(f"Não encontrei '{ARQUIVO_MAPA_PA.name}' no Google Drive — a previsão usa essa planilha pra saber QUAIS mapas considerar naquele dia.")
    elif df_020501_historico is None or df_020501_historico.empty:
        st.info("⚠️ Aguardando dados do relatório 02.05.01.")
    else:
        data_previsao_str = data_previsao.strftime("%d/%m/%Y")
        
        sub_conc_data = df_mapa_pa[df_mapa_pa["Data"] == data_previsao_str].copy()
        
        if sub_conc_data.empty:
            st.warning(f"Nenhum mapa cadastrado em {data_previsao_str} na planilha '{ARQUIVO_MAPA_PA.name}'.")
        else:
            pa_por_mapa_data = {}
            for _, r in sub_conc_data.iterrows():
                m_res = resolver_mapa(str(r["Mapa"]))
                pa_bruto = str(r.get("PA", "Sede")).strip().upper()
                pa_norm = _PA_NORMALIZADO.get(pa_bruto, str(r.get("PA", "Sede")).strip())
                pa_por_mapa_data[m_res] = pa_norm

            mapas_previsao_originais = sorted(sub_conc_data["Mapa"].dropna().unique().tolist(), key=lambda m: int(m) if str(m).isdigit() else 0)
            mapas_previsao = resolver_mapas(mapas_previsao_originais)

            df_previsao = df_020501_historico[df_020501_historico["Mapa"].isin(mapas_previsao)].copy()
            mapas_encontrados = set(df_previsao["Mapa"].unique())
            mapas_faltando = [m for m in mapas_previsao if m not in mapas_encontrados]

            if mapas_faltando:
                st.warning(f"{len(mapas_faltando)} mapa(s) ainda não estão no relatório: {', '.join(mapas_faltando)}. Previsão incompleta.")
            else:
                st.caption(f"{len(mapas_previsao)} mapa(s) encontrados no total.")

            if df_previsao.empty:
                st.info("Nenhum dos mapas dessa data foi encontrado no relatório ainda.")
            else:
                df_previsao["PA"] = df_previsao["Mapa"].apply(lambda m: pa_por_mapa_data.get(m, "Sede"))
                
                pas_no_dia = sorted(df_previsao["PA"].unique().tolist(), key=lambda p: (0 if p == "Sede" else 1, p))

                titulos_abas = ["🌐 Total Geral"] + [f"🏢 {p}" if p == "Sede" else f"🚛 {p}" for p in pas_no_dia]
                sub_abas_previsao = st.tabs(titulos_abas)

                linhas_prev_save = []

                # 1. SUB-ABA: TOTAL GERAL
                with sub_abas_previsao[0]:
                    st.markdown(f"#### 🌐 Previsão Total Geral ({len(mapas_previsao)} mapas)")
                    dados_fam_total, dados_outros_total = extrair_dados_farol(df_previsao, lookup_ag)
                    if dados_fam_total or dados_outros_total:
                        renderizar_farol_previsao(dados_fam_total, dados_outros_total)
                    else:
                        st.info("Não houve movimentação de vasilhame no total geral.")

                # 2. SUB-ABAS: POR PA / SEDE
                for idx_pa, nome_pa in enumerate(pas_no_dia):
                    with sub_abas_previsao[idx_pa + 1]:
                        df_pa_especifico = df_previsao[df_previsao["PA"] == nome_pa]
                        mapas_pa_especifico = sorted(df_pa_especifico["Mapa"].unique(), key=lambda m: int(m) if str(m).isdigit() else 0)
                        
                        icone_pa = "🏢" if nome_pa == "Sede" else "🚛"
                        st.markdown(f"#### {icone_pa} Previsão — {nome_pa} ({len(mapas_pa_especifico)} mapas)")
                        st.caption(f"Mapas: {', '.join(mapas_pa_especifico)}")
                        
                        dados_fam_pa, dados_outros_pa = extrair_dados_farol(df_pa_especifico, lookup_ag)
                        
                        for fam, d in dados_fam_pa.items():
                            linhas_prev_save.append({
                                "Data": data_previsao_str, "PA": nome_pa,
                                "Item": rotulo_familia_vazio(fam),
                                "Caixas": d["caixas"], "Garrafas_Soltas": d["soltas"]
                            })
                        for item_nome, qtd in dados_outros_pa.items():
                            linhas_prev_save.append({
                                "Data": data_previsao_str, "PA": nome_pa,
                                "Item": item_nome,
                                "Caixas": 0, "Garrafas_Soltas": qtd
                            })
                            
                        if dados_fam_pa or dados_outros_pa:
                            renderizar_farol_previsao(dados_fam_pa, dados_outros_pa)
                        else:
                            st.info(f"Não houve movimentação de vasilhame em {nome_pa} nesta data.")

                if linhas_prev_save:
                    st.write("")
                    if st.button("💾 Salvar Previsão no Histórico", key="btn_save_previsao"):
                        acumular_historico(pd.DataFrame(linhas_prev_save), "Snap_PrevisaoAG", ["Data", "PA", "Item"])
                        st.success("✅ Salvo na aba 'Snap_PrevisaoAG' do histórico no Drive!")

# =========================================================================
# ABA DE HISTÓRICO GERENCIAL (CONSULTA DOS ÚLTIMOS 3 MESES)
# =========================================================================
with aba_historico:
    st.header("📈 Histórico Gerencial")
    st.caption("Consulte os retratos de conferência salvos nos últimos 90 dias.")

    data_limite_inicio = date.today() - timedelta(days=90)
    
    col_hist1, _ = st.columns([1, 2])
    data_filtro_hist = col_hist1.date_input(
        "Filtrar período do histórico:",
        value=(data_limite_inicio, date.today()),
        max_value=date.today(),
        key="filtro_datas_historico"
    )

    if isinstance(data_filtro_hist, tuple) and len(data_filtro_hist) == 2:
        dt_inicio_hist, dt_fim_hist = data_filtro_hist
    else:
        dt_inicio_hist = data_filtro_hist[0] if isinstance(data_filtro_hist, tuple) else data_filtro_hist
        dt_fim_hist = date.today()

    aba_h_resumo, aba_h_pendencias, aba_h_previsao = st.tabs(["Resumo Acumulado", "Pendências (PA)", "Previsão (Sede/PA)"])

    # --- HISTÓRICO: RESUMO ACUMULADO ---
    with aba_h_resumo:
        df_snap_resumo = ler_aba_historico("Snap_ResumoAcumulado")
        if df_snap_resumo.empty:
            st.info("Nenhum histórico de Resumo Acumulado foi salvo ainda.")
        else:
            df_snap_resumo["Data_Real"] = pd.to_datetime(df_snap_resumo["Data"], format="%d/%m/%Y", errors="coerce")
            mascara = (df_snap_resumo["Data_Real"].dt.date >= dt_inicio_hist) & (df_snap_resumo["Data_Real"].dt.date <= dt_fim_hist)
            df_filtrado_resumo = df_snap_resumo[mascara].drop(columns=["Data_Real"]).sort_values(by=["Data", "PA"], ascending=[False, True])
            
            if df_filtrado_resumo.empty:
                st.warning("Nenhum registro encontrado neste período.")
            else:
                st.dataframe(df_filtrado_resumo, width='stretch', hide_index=True)

    # --- HISTÓRICO: PENDÊNCIAS ---
    with aba_h_pendencias:
        df_snap_pend = ler_aba_historico("Snap_Pendencias")
        if df_snap_pend.empty:
            st.info("Nenhum histórico de Pendências foi salvo ainda.")
        else:
            df_snap_pend["Data_Real"] = pd.to_datetime(df_snap_pend["Data"], format="%d/%m/%Y", errors="coerce")
            mascara = (df_snap_pend["Data_Real"].dt.date >= dt_inicio_hist) & (df_snap_pend["Data_Real"].dt.date <= dt_fim_hist)
            df_filtrado_pend = df_snap_pend[mascara].drop(columns=["Data_Real"]).sort_values(by=["Data", "PA"], ascending=[False, True])
            
            if df_filtrado_pend.empty:
                st.warning("Nenhum registro encontrado neste período.")
            else:
                st.dataframe(df_filtrado_pend, width='stretch', hide_index=True)

    # --- HISTÓRICO: PREVISÃO ---
    with aba_h_previsao:
        df_snap_prev = ler_aba_historico("Snap_PrevisaoAG")
        if df_snap_prev.empty:
            st.info("Nenhum histórico de Previsão foi salvo ainda.")
        else:
            df_snap_prev["Data_Real"] = pd.to_datetime(df_snap_prev["Data"], format="%d/%m/%Y", errors="coerce")
            mascara = (df_snap_prev["Data_Real"].dt.date >= dt_inicio_hist) & (df_snap_prev["Data_Real"].dt.date <= dt_fim_hist)
            df_filtrado_prev = df_snap_prev[mascara].drop(columns=["Data_Real"]).sort_values(by=["Data", "PA"], ascending=[False, True])
            
            if df_filtrado_prev.empty:
                st.warning("Nenhum registro encontrado neste período.")
            else:
                st.dataframe(df_filtrado_prev, width='stretch', hide_index=True)
