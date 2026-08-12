import streamlit as st
import pandas as pd
from datetime import date, timedelta

from comum import (
    ARQUIVO_DE_MATERIAL,
    ARQUIVO_MAPAS_AG,
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
    formata_diferenca_fisica,
    com_apelido,
    rotulo_familia_vazio,
    montar_lookup_ag_por_codigo,
    familia_tipo_por_codigo,
    CATEGORIAS_AG_EXTRA,
)

# Planilha de referência Data+Mapa->PA (colunas DATA, MAPA, PONTO DE APOIO). Sobe na
# mesma pasta do Drive sempre com o nome "CONC.csv" — carregar() busca por esse
# prefixo, igual aos outros arquivos (De Material, 03.07.13).
ARQUIVO_MAPA_PA = PASTA_PROJETO / "CONC.csv"

st.set_page_config(page_title="Conciliação de Mapas (AG)", layout="wide")
st.title("⚖️ Conciliação de Mapas (AG)")
st.caption("_\"Balança enganosa é abominação ao SENHOR, mas o peso justo lhe é agradável.\" — Provérbios 11:1_")

# Aba separada no historico_ag.xlsx só pra dados de simulação — nunca mistura com a
# aba "VazioPA" de produção. Ativar o modo simulação (sidebar) troca de qual aba o
# app lê, sem apagar nem sobrescrever nada real.
NOME_ABA_SIMULACAO = "VazioPA_Simulacao"

REGRAS_VAZIO = {
    "300ml": {"garrafas_por_cx": 23, "garrafeiras_por_cx": 1},
    "600ml": {"garrafas_por_cx": 24, "garrafeiras_por_cx": 1},
    "Verde 600": {"garrafas_por_cx": 24, "garrafeiras_por_cx": 1},
    "1L": {"garrafas_por_cx": 12, "garrafeiras_por_cx": 1},
}

# Mesma paleta usada em cor_linha_status — cada status vira o fundo pastel e o emoji
# gigante do "cartão de resultado" nos resumos "pra enviar", e também alimenta a
# tabela "estilo limpo" (pill de status) usada nos blocos de detalhe.
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
    """itens: lista de (rotulo, valor, cor) ou (rotulo, valor, cor, detalhes) — cor é uma
    chave de CORES_RESUMO. detalhes (opcional) é uma lista de linhas curtas mostradas
    dentro do card quando valor > 0 (ex: qual mapa/item está causando aquele número).
    Visual compacto: badge inline (ícone+número+rótulo numa linha só), detalhes como
    texto pequeno logo abaixo, sem cartão gigante."""
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


# Paleta inspirada no "FAROL AG" do usuário: cada família de garrafa tem uma cor de
# identidade própria (faixa escura no topo do card + fundo pastel no corpo).
CORES_FAROL_FAMILIA = {
    "600ml": ("#1F3B57", "#EAF1F8"),
    "Verde 600": ("#1F4720", "#E9F5EA"),
    "300ml": ("#8A6D1B", "#FBF3DF"),
    "1L": ("#2B2B2B", "#ECECEC"),
}
ORDEM_FAROL_FAMILIA = ["600ml", "Verde 600", "300ml", "1L"]


def renderizar_farol_previsao(dados_familia: dict, dados_outros: dict) -> None:
    """Visual tipo 'FAROL AG': um card colorido por família de garrafa (300/600/Verde/
    Litrão) mostrando Caixas + garrafas soltas, e uma segunda fileira de cards escuros
    pros itens que não convertem em caixa (Garrafeira, Pallet, Chapatex, Barril),
    mostrando só a unidade. dados_familia: {familia: {"caixas": int, "soltas": int}}.
    dados_outros: {rótulo: unidades}."""
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


def _chave_cor_status(status: str) -> str:
    """Resolve a chave de CORES_RESUMO a partir do emoji presente no texto de status —
    mesma lógica de cor_linha_status(), mas devolvendo a chave em vez do CSS pronto."""
    s = str(status)
    if "✅" in s: return "verde"
    if "❌" in s: return "vermelho"
    if "⚠️" in s: return "amarelo"
    if "🔎" in s: return "azul"
    if "⏳" in s: return "cinza"
    return "cinza"


def _pill_status(status: str) -> str:
    """Status como pill arredondado (fundo pastel, texto curto) em vez de célula/linha
    inteira colorida — usado por renderizar_tabela_limpa()."""
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
    """Verde pra entrada/sobra (+), vermelho pra falta (-) — aplicado só na coluna Diferença."""
    t = str(texto).strip()
    if t.startswith("+"): return "#0F6E56"
    if t.startswith("-"): return "#A32D2D"
    return "#888780"


def renderizar_tabela_limpa(df: pd.DataFrame, colunas: list[str], col_status: str = "Status") -> None:
    """Tabela HTML em estilo 'planilha limpa': cabeçalho discreto em cinza, linhas finas
    separadas por hairline, primeira coluna alinhada à esquerda e as demais à direita,
    coluna Diferença colorida por sinal, e Status como pill em vez de linha/célula
    inteira pintada. Substitui st.dataframe(...).style.map(cor_linha_status) nos blocos
    de 'itens com diferença' das abas de conciliação."""
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


def status_por_mapa(df: pd.DataFrame, ordem_prioridade: list[str]) -> pd.Series:
    """Resume os status de TODAS as linhas (famílias/itens) de cada mapa num status só
    por mapa — usando a ordem de prioridade dada (primeiro item = pior/mais urgente,
    último = status "bom"). Um mapa com qualquer linha "Faltou" conta como Faltou,
    mesmo que as outras famílias tenham batido — só conta como Bateu se TODAS baterem."""
    def _pior(serie_status: pd.Series) -> str:
        valores = set(serie_status)
        for status in ordem_prioridade:
            if status in valores:
                return status
        return ordem_prioridade[-1]
    return df.groupby("Mapa")["Status"].apply(_pior)


with st.sidebar:
    st.caption(f"Fonte: {ARQUIVO_MAPAS_AG.name} (atualiza sozinho a cada 5 min)")
    if st.button("🔄 Recarregar tela", width="stretch"):
        st.rerun()
    intervalo_datas = st.date_input(
        "Considerar mapas do CONC.csv no período:",
        value=(date(2026, 8, 1), date.today()),
    )
    # date_input com range só retorna as duas datas depois que o usuário escolhe as duas
    # no calendário — enquanto só a primeira estiver selecionada, vem uma tupla de 1 item.
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
        help="Mostra dados de teste (gerados na aba 'Vazio por PA') em vez dos reais. Não mexe nos dados de produção.",
    )

if modo_simulacao:
    st.warning("🧪 **MODO SIMULAÇÃO ATIVO** — os dados de conferência abaixo são de teste, não reais. Desative na sidebar pra voltar ao normal.")

# --- De Material: usado pra classificar por Código e pra filtrar itens válidos de AG ---
df_de_material = carregar(ARQUIVO_DE_MATERIAL)
if df_de_material is not None and "Promax" in df_de_material.columns:
    df_de_material["Promax"] = normalizar_codigo(df_de_material["Promax"])
lookup_ag = montar_lookup_ag_por_codigo(df_de_material) if df_de_material is not None else {}

# --- Mapa PA: diz de qual PA é cada mapa numa data — alimenta a busca automática de
# mapas na Conferência em Lote, pra o conferente não precisar digitar número nenhum.
# Coluna opcional "MAPA CONSOLIDADO": quando vários mapas viram um só no relatório
# (ex: 257682 + 257685 → 257693), essa coluna traz o número final pra cada original.
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
        df_mapa_pa["Mapa"] = df_mapa_pa["Mapa"].apply(limpa_mapa)
    if "Data" in df_mapa_pa.columns:
        _dt_mapa_pa = pd.to_datetime(df_mapa_pa["Data"], dayfirst=True, errors="coerce")
        df_mapa_pa["Data"] = _dt_mapa_pa.dt.strftime("%d/%m/%Y")
    if "PA" in df_mapa_pa.columns:
        df_mapa_pa["PA"] = df_mapa_pa["PA"].astype(str).str.strip()
    if "MapaConsolidado" in df_mapa_pa.columns:
        def _limpar_mapa_consolidado(v) -> str:
            """Igual limpa_mapa(), mas trata o caso de vir como float (ex: 257690.0)
            — acontece quando a coluna tem células vazias misturadas com números,
            o que faz o pandas/Excel guardar tudo como decimal."""
            s = str(v).strip()
            if s.lower() in ("", "nan", "none", "0"):
                return ""
            try:
                return str(int(float(s)))
            except Exception:
                return s
        df_mapa_pa["MapaConsolidado"] = df_mapa_pa["MapaConsolidado"].apply(_limpar_mapa_consolidado)
    else:
        df_mapa_pa["MapaConsolidado"] = ""


@st.cache_data(show_spinner=False, ttl=300)
def _ingerir_conc_no_historico(_df_conc_limpo: pd.DataFrame) -> pd.DataFrame:
    """Acumula o CONC.csv atual (já limpo: Data/Mapa/PA/MapaConsolidado) num histórico
    permanente — assim, mesmo substituindo o CONC.csv todo dia, um mapa de um dia
    anterior que ainda não foi conferido continua aparecendo nas conciliações, em vez
    de sumir quando o arquivo do dia é trocado. Cacheado 5 min (igual carregar()) pra
    não gravar no Drive a cada interação do usuário na tela."""
    colunas_manter = [c for c in ["Data", "Mapa", "PA", "MapaConsolidado"] if c in _df_conc_limpo.columns]
    return acumular_historico(_df_conc_limpo[colunas_manter], "MapaPAHistorico", ["Data", "Mapa"])


if df_mapa_pa is not None and not df_mapa_pa.empty:
    df_mapa_pa = _ingerir_conc_no_historico(df_mapa_pa)
else:
    df_mapa_pa = ler_aba_historico("MapaPAHistorico")

if df_mapa_pa is not None and not df_mapa_pa.empty:
    for _, _linha in df_mapa_pa.iterrows():
        if _linha.get("MapaConsolidado") and _linha["MapaConsolidado"] != _linha["Mapa"]:
            MAPA_CONSOLIDADO_LOOKUP[_linha["Mapa"]] = _linha["MapaConsolidado"]


def resolver_mapa(mapa: str) -> str:
    """Se o mapa foi consolidado em outro (planilha CONC, coluna MAPA CONSOLIDADO),
    devolve o número final — senão devolve o próprio mapa sem alteração."""
    return MAPA_CONSOLIDADO_LOOKUP.get(mapa, mapa)


def resolver_mapas(mapas) -> list[str]:
    """Aplica resolver_mapa numa lista, sem duplicar quando vários originais caem no
    mesmo mapa consolidado (ex: 257682 e 257685 os dois viram 257693 uma vez só)."""
    resolvidos: list[str] = []
    vistos: set[str] = set()
    for m in mapas:
        alvo = resolver_mapa(m)
        if alvo not in vistos:
            vistos.add(alvo)
            resolvidos.append(alvo)
    return resolvidos


# Sentido inverso — dado o mapa consolidado, quais originais viraram ele. Usado só pra
# deixar a exibição clara ("257682+257685 (→257693)") sem mudar o cálculo.
REVERSE_MAPA_CONSOLIDADO: dict[str, list[str]] = {}
for _orig, _cons in MAPA_CONSOLIDADO_LOOKUP.items():
    REVERSE_MAPA_CONSOLIDADO.setdefault(_cons, []).append(_orig)


def rotulo_mapa(mapa: str) -> str:
    """Número do mapa pra exibição — se ele é um consolidado de vários originais,
    mostra 'orig1+orig2 (→consolidado)'; senão mostra o número normalmente."""
    originais = REVERSE_MAPA_CONSOLIDADO.get(mapa)
    if not originais:
        return mapa
    ordenados = sorted(originais, key=lambda m: int(m) if str(m).isdigit() else 0)
    return f"{'+'.join(ordenados)} (→{mapa})"


def buscar_mapas_por_data_pa(data_alvo, pa_alvo: str) -> list[str]:
    """Consulta df_mapa_pa e devolve os números de mapa cadastrados pra essa Data+PA —
    é isso que substitui o campo de digitação manual na Conferência em Lote. Devolve os
    números ORIGINAIS (não resolvidos) — a resolução de consolidação acontece só na
    hora de buscar a Saída, não aqui."""
    if df_mapa_pa is None or df_mapa_pa.empty or "Data" not in df_mapa_pa.columns or "PA" not in df_mapa_pa.columns:
        return []
    data_str = data_alvo.strftime("%d/%m/%Y")
    sub = df_mapa_pa[(df_mapa_pa["Data"] == data_str) & (df_mapa_pa["PA"].str.upper() == pa_alvo.upper())]
    return sorted(sub["Mapa"].dropna().unique().tolist(), key=lambda m: int(m) if str(m).isdigit() else 0)


# =========================================================================
# CONC.csv é a fonte única de verdade de QUAIS MAPAS e QUAIS DATAS entram em
# cada conciliação — o relatório 03.07.13 só é usado pra consultar valores por
# número de mapa, nunca pra decidir se um mapa entra ou não. O filtro de
# período da sidebar agora escopa o CONC.csv (não mais o relatório).
# =========================================================================
_PA_NORMALIZADO = {"TIANGUÁ": "Tianguá", "TIANGUA": "Tianguá", "GRANJA": "Granja", "SEDE": "Sede"}

df_mapa_pa_periodo = df_mapa_pa
if df_mapa_pa is not None and not df_mapa_pa.empty and "Data" in df_mapa_pa.columns:
    _dt_conc = pd.to_datetime(df_mapa_pa["Data"], dayfirst=True, errors="coerce")
    df_mapa_pa_periodo = df_mapa_pa[(_dt_conc >= pd.Timestamp(data_inicio)) & (_dt_conc <= pd.Timestamp(data_fim))]

# {mapa_resolvido: "Tianguá"/"Granja"/"Sede"} — dita o roteamento entre as abas.
MAPA_PA_CLASSIFICACAO: dict[str, str] = {}
if df_mapa_pa_periodo is not None and not df_mapa_pa_periodo.empty and "PA" in df_mapa_pa_periodo.columns:
    for _, _linha_conc in df_mapa_pa_periodo.iterrows():
        _mapa_resolvido = resolver_mapa(_linha_conc["Mapa"])
        _pa_bruto = str(_linha_conc["PA"]).strip().upper()
        MAPA_PA_CLASSIFICACAO[_mapa_resolvido] = _PA_NORMALIZADO.get(_pa_bruto, _linha_conc["PA"])

MAPAS_PA_CONC = {m for m, pa in MAPA_PA_CLASSIFICACAO.items() if pa in ("Tianguá", "Granja")}
MAPAS_SEDE_CONC = {m for m, pa in MAPA_PA_CLASSIFICACAO.items() if pa == "Sede"}


# --- 03.07.13: carrega e filtra pelo período escolhido, sem gravar nada no Drive ---
df_mapas_ag = carregar(ARQUIVO_MAPAS_AG)
if df_mapas_ag is not None:
    df_mapas_ag = df_mapas_ag.copy()
    df_mapas_ag.columns = df_mapas_ag.columns.str.strip()
    if "Material" in df_mapas_ag.columns:
        df_mapas_ag["Material"] = df_mapas_ag["Material"].apply(limpa_mapa)
    if "Mapa" in df_mapas_ag.columns:
        df_mapas_ag["Mapa"] = df_mapas_ag["Mapa"].apply(limpa_mapa)
    if "Descricao" in df_mapas_ag.columns:
        df_mapas_ag["Descricao"] = df_mapas_ag["Descricao"].astype(str).str.strip()

    colunas_numericas_relatorio = ["P Vazia", "R Vazio"] + [
        c for _, cp, cr in CATEGORIAS_AG_EXTRA for c in (cp, cr)
    ]
    for col_qtd in colunas_numericas_relatorio:
        if col_qtd in df_mapas_ag.columns:
            df_mapas_ag[col_qtd] = pd.to_numeric(df_mapas_ag[col_qtd], errors="coerce").fillna(0)

    if df_de_material is not None and "Material" in df_mapas_ag.columns:
        codigos_validos = set(df_de_material["Promax"].unique())
        df_mapas_ag = df_mapas_ag[df_mapas_ag["Material"].isin(codigos_validos)]

    # O relatório 03.07.13 nunca é filtrado por período — o CONC.csv já dita quais
    # mapas e datas entram em cada conciliação; aqui só se consulta por número de mapa.
    df_mapas_ag_sem_filtro_data = df_mapas_ag
else:
    df_mapas_ag_sem_filtro_data = None

_periodo_str = f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
with st.sidebar:
    if df_mapas_ag is None:
        st.error(f"Não encontrei '{ARQUIVO_MAPAS_AG.name}' no Google Drive.")
    elif df_mapa_pa is None or df_mapa_pa.empty:
        st.error(f"Não encontrei '{ARQUIVO_MAPA_PA.name}' no Google Drive nem histórico acumulado ainda.")
    else:
        st.success(
            f"{ARQUIVO_MAPAS_AG.name}: {len(df_mapas_ag)} linha(s). "
            f"CONC.csv: {len(MAPA_PA_CLASSIFICACAO)} mapa(s) no período ({_periodo_str})."
        )


def gerar_simulacao_perfeita(data_alvo) -> pd.DataFrame:
    """Pra cada mapa Tianguá/Granja do CONC.csv na data escolhida, monta uma linha de
    retorno = saída exata (garrafas soltas, sem caixas/garrafeiras/unidades) — simula
    uma conferência 100% perfeita, só pra teste visual. Ignora consolidação de mapas
    (rara em Tianguá/Granja) usando direto o número resolvido."""
    if df_mapa_pa is None or df_mapa_pa.empty or df_mapas_ag_sem_filtro_data is None or df_mapas_ag_sem_filtro_data.empty:
        return pd.DataFrame()

    data_str = data_alvo.strftime("%d/%m/%Y")
    mapas_pa_sim = df_mapa_pa[
        (df_mapa_pa["Data"] == data_str) & (df_mapa_pa["PA"].str.upper().isin(["TIANGUÁ", "GRANJA"]))
    ][["Mapa", "PA"]].drop_duplicates().copy()
    if mapas_pa_sim.empty:
        return pd.DataFrame()

    mapas_pa_sim["MapaResolvido"] = mapas_pa_sim["Mapa"].apply(resolver_mapa)
    pa_normalizado = {"TIANGUÁ": "Tianguá", "GRANJA": "Granja"}
    mapas_pa_sim["PA"] = mapas_pa_sim["PA"].str.upper().map(pa_normalizado)
    pa_lookup_sim = mapas_pa_sim.groupby("MapaResolvido")["PA"].first().to_dict()

    familia_tipo_sim = df_mapas_ag_sem_filtro_data["Material"].apply(lambda c: familia_tipo_por_codigo(c, lookup_ag))
    df_sim = df_mapas_ag_sem_filtro_data.copy()
    df_sim["Familia"] = familia_tipo_sim.apply(lambda ft: ft[0])
    df_sim["Tipo"] = familia_tipo_sim.apply(lambda ft: ft[1])
    df_sim = df_sim[(df_sim["Familia"] != "Outro") & (df_sim["Tipo"] != "Garrafeira")]
    df_sim = df_sim[df_sim["Mapa"].isin(mapas_pa_sim["MapaResolvido"].unique())]

    saida_sim = df_sim.groupby(["Mapa", "Familia"])["P Vazia"].sum().reset_index()
    saida_sim = saida_sim[saida_sim["P Vazia"] > 0]

    linhas = []
    for _, r in saida_sim.iterrows():
        linhas.append({
            "Data": data_str,
            "PA": pa_lookup_sim.get(r["Mapa"], "Tianguá"),
            "Mapa": r["Mapa"],
            "Familia": r["Familia"],
            "Caixas": 0,
            "Garrafas": int(r["P Vazia"]),
            "Garrafeiras": 0,
            "Unidades": 0,
        })
    return pd.DataFrame(linhas)

aba_vazio_pa, aba_conciliacao, aba_conciliacao_sede, aba_categorias_extra, aba_fechamento = st.tabs(
    ["Vazio por PA", "Conciliação Mapas PA", "Conciliação Mapas Sede", "Outras Categorias", "📊 Fechamento"]
)

# Roteamento entre as abas agora vem do CONC.csv (MAPA_PA_CLASSIFICACAO), não mais de
# "foi digitado ou não". MAPAS_INDIVIDUAIS/MAPAS_EM_LOTE continuam existindo só pra
# saber o que já foi conferido (usado no cruzamento de Retorno), não pra roteamento.
def mapas_da_lote(mapas_str: str) -> list[str]:
    """'257379;257386;257402' -> ['257379','257386','257402'] — ';' é o separador usado
    pra guardar o conjunto de mapas de uma conferência em lote numa única célula."""
    return [limpa_mapa(m) for m in str(mapas_str).split(";") if str(m).strip()]


_hist_vazio_pa_bruto = ler_aba_historico(NOME_ABA_SIMULACAO if modo_simulacao else "VazioPA")
# Usado em todo o formulário/tabela/edição/exclusão da aba "Vazio por PA" — com a
# simulação ativa, a aba inteira vira um sandbox isolado (lê e grava só na aba de
# simulação); desativada, volta a mexer só na aba real "VazioPA".
ABA_VAZIO_PA_ATIVA = NOME_ABA_SIMULACAO if modo_simulacao else "VazioPA"
if not _hist_vazio_pa_bruto.empty and "Mapa" in _hist_vazio_pa_bruto.columns:
    MAPAS_INDIVIDUAIS = set(_hist_vazio_pa_bruto["Mapa"].apply(limpa_mapa).unique())
else:
    MAPAS_INDIVIDUAIS = set()

_hist_lote_bruto = ler_aba_historico("VazioPALote")
MAPAS_EM_LOTE = set()
if not _hist_lote_bruto.empty and "Mapas" in _hist_lote_bruto.columns:
    for _mapas_str in _hist_lote_bruto["Mapas"].unique():
        MAPAS_EM_LOTE.update(mapas_da_lote(_mapas_str))



# =========================================================================
# ABA VAZIO POR PA (conferência física digitada pelo conferente)
# =========================================================================
with aba_vazio_pa:
    st.caption("Conferência do vazio por PA e mapa.")

    with st.form("form_vazio_pa", clear_on_submit=True):
        col_data, col_pa, col_mapa = st.columns(3)
        data_pa = col_data.date_input("Data da Descarga", value=date.today(), key="data_vazio_pa")
        pa_escolhido = col_pa.selectbox("PA", ["Tianguá", "Granja"], key="pa_vazio_pa")
        mapa_texto = col_mapa.text_input("Número do Mapa (um por vez)")

        st.markdown("**Caixas Físicas que Retornaram**")
        valores_familia_pa = {fam: st.number_input(rotulo_familia_vazio(fam), min_value=0, step=1, key=f"cx_pa_{fam}") for fam in REGRAS_VAZIO}

        st.markdown("**Outros AG (sem conversão — já em unidade final)**")
        c1, c2, c3, c4, c5 = st.columns(5)
        chapatex_pa = c1.number_input("Chapatex (Und)", min_value=0, step=1, key="outros_pa_chapatex")
        pbr1_pa = c2.number_input("Pallet PBR1", min_value=0, step=1, key="outros_pa_pbr1")
        pbr2_pa = c3.number_input("Pallet PBR2", min_value=0, step=1, key="outros_pa_pbr2")
        barril30_pa = c4.number_input("Barril 30L", min_value=0, step=1, key="outros_pa_barril30")
        barril50_pa = c5.number_input("Barril 50L", min_value=0, step=1, key="outros_pa_barril50")

        if st.form_submit_button("Salvar conferência"):
            mapa_numero = limpa_mapa(mapa_texto.strip())
            if not mapa_texto.strip():
                st.error("Informe o número do mapa antes de salvar.")
            elif "," in mapa_texto:
                st.error("Um mapa por vez — se tiver mais de um, salve cada um separadamente (o formulário limpa sozinho depois de salvar).")
            else:
                data_str_pa = data_pa.strftime("%d/%m/%Y")
                gf_600 = valores_familia_pa.get("600ml", 0) + valores_familia_pa.get("Verde 600", 0)
                linhas_pa = []

                for familia, qtd_cx in valores_familia_pa.items():
                    if qtd_cx > 0:
                        r = REGRAS_VAZIO[familia]
                        gf = gf_600 if familia == "600ml" else (0 if familia == "Verde 600" else qtd_cx * r["garrafeiras_por_cx"])
                        linhas_pa.append({
                            "Data": data_str_pa,
                            "PA": pa_escolhido,
                            "Mapa": mapa_numero,
                            "Familia": familia,
                            "Caixas": qtd_cx,
                            "Garrafas": qtd_cx * r["garrafas_por_cx"],
                            "Garrafeiras": gf,
                            "Unidades": 0,
                        })

                for familia_outros, qtd_un in [
                    ("Chapatex", chapatex_pa), ("Pallet PBR1", pbr1_pa), ("Pallet PBR2", pbr2_pa),
                    ("Barril 30L", barril30_pa), ("Barril 50L", barril50_pa),
                ]:
                    if qtd_un > 0:
                        linhas_pa.append({
                            "Data": data_str_pa,
                            "PA": pa_escolhido,
                            "Mapa": mapa_numero,
                            "Familia": familia_outros,
                            "Caixas": 0,
                            "Garrafas": 0,
                            "Garrafeiras": 0,
                            "Unidades": qtd_un,
                        })

                if linhas_pa:
                    acumular_historico(pd.DataFrame(linhas_pa), ABA_VAZIO_PA_ATIVA, ["Data", "PA", "Mapa", "Familia"])
                    st.success(f"✅ Retorno do mapa {mapa_numero} salvo com sucesso!")
                else:
                    st.warning("Nenhuma quantidade foi informada para salvar.")

    st.divider()
    st.markdown("### 📦 Conferência em Lote (vários mapas conferidos juntos)")
    st.caption("Use quando só souber o TOTAL, sem separar por mapa.")

    col_data_l, col_pa_l = st.columns(2)
    data_lote = col_data_l.date_input("Data da Descarga", value=date.today(), key="data_lote")
    pa_lote = col_pa_l.selectbox("PA", ["Tianguá", "Granja"], key="pa_lote")

    mapas_lote_auto = buscar_mapas_por_data_pa(data_lote, pa_lote)
    if df_mapa_pa is None or df_mapa_pa.empty:
        st.error(f"Não encontrei '{ARQUIVO_MAPA_PA.name}' no Google Drive nem histórico acumulado ainda — sem isso não dá pra buscar os mapas automaticamente.")
    elif mapas_lote_auto:
        st.success(f"{len(mapas_lote_auto)} mapa(s) de {pa_lote} em {data_lote.strftime('%d/%m/%Y')}: {', '.join(mapas_lote_auto)}")
    else:
        st.warning(f"Nenhum mapa cadastrado pra {pa_lote} em {data_lote.strftime('%d/%m/%Y')} na planilha 'Mapa PA'.")

    with st.form("form_vazio_pa_lote", clear_on_submit=True):
        st.markdown("**Caixas Físicas que Retornaram (TOTAL do lote)**")
        valores_familia_lote = {
            fam: st.number_input(rotulo_familia_vazio(fam), min_value=0, step=1, key=f"cx_lote_{fam}")
            for fam in REGRAS_VAZIO
        }

        st.markdown("**Outros AG (TOTAL do lote, sem conversão)**")
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

                for familia, qtd_cx in valores_familia_lote.items():
                    if qtd_cx > 0:
                        r = REGRAS_VAZIO[familia]
                        gf = gf_600_lote if familia == "600ml" else (0 if familia == "Verde 600" else qtd_cx * r["garrafeiras_por_cx"])
                        linhas_lote.append({
                            "Data": data_str_lote, "PA": pa_lote, "Mapas": mapas_chave, "Familia": familia,
                            "Caixas": qtd_cx, "Garrafas": qtd_cx * r["garrafas_por_cx"], "Garrafeiras": gf, "Unidades": 0,
                        })

                for familia_outros, qtd_un in [
                    ("Chapatex", chapatex_lote), ("Pallet PBR1", pbr1_lote), ("Pallet PBR2", pbr2_lote),
                    ("Barril 30L", barril30_lote), ("Barril 50L", barril50_lote),
                ]:
                    if qtd_un > 0:
                        linhas_lote.append({
                            "Data": data_str_lote, "PA": pa_lote, "Mapas": mapas_chave, "Familia": familia_outros,
                            "Caixas": 0, "Garrafas": 0, "Garrafeiras": 0, "Unidades": qtd_un,
                        })

                if linhas_lote:
                    acumular_historico(pd.DataFrame(linhas_lote), "VazioPALote", ["Data", "PA", "Mapas", "Familia"])
                    st.success(f"✅ Lote de {len(mapas_lote_auto)} mapas ({', '.join(mapas_lote_auto)}) salvo com sucesso!")
                else:
                    st.warning("Nenhuma quantidade foi informada para salvar.")

    # =====================================================================
    # A PARTIR DAQUI: só telas de conferência (tabelas, edição, exclusão) —
    # os dois formulários de lançamento (individual e lote) ficam sempre no
    # topo da aba, acima desta linha.
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
                # Todas as famílias possíveis — não só as que já têm registro pra esse mapa,
                # assim dá pra ADICIONAR um item que nunca foi digitado (ex: Chapatex),
                # não só editar o que já existe.
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

                ce1, ce2, ce3, ce4 = st.columns(4)
                novo_caixas = ce1.number_input("Caixas", min_value=0, step=1, value=int(linha_atual.get("Caixas", 0)) if isinstance(linha_atual, pd.Series) else 0, key="edit_caixas")
                novo_garrafas = ce2.number_input("Garrafas", min_value=0, step=1, value=int(linha_atual.get("Garrafas", 0)) if isinstance(linha_atual, pd.Series) else 0, key="edit_garrafas")
                novo_garrafeiras = ce3.number_input("Garrafeiras", min_value=0, step=1, value=int(linha_atual.get("Garrafeiras", 0)) if isinstance(linha_atual, pd.Series) else 0, key="edit_garrafeiras")
                novo_unidades = ce4.number_input("Unidades", min_value=0, step=1, value=int(linha_atual.get("Unidades", 0)) if isinstance(linha_atual, pd.Series) else 0, key="edit_unidades")

                if st.button("💾 Salvar edição", type="primary", key="salvar_edicao_pa"):
                    nova_linha = pd.DataFrame([{
                        "Data": edit_data, "PA": pa_atual, "Mapa": edit_mapa, "Familia": edit_familia,
                        "Caixas": novo_caixas, "Garrafas": novo_garrafas, "Garrafeiras": novo_garrafeiras, "Unidades": novo_unidades,
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
                df_restante = df_vazio_pa[~((df_vazio_pa["Data"] == del_data) & (df_vazio_pa["Mapa"].astype(str) == del_mapa))]
                salvar_aba_historico(ABA_VAZIO_PA_ATIVA, df_restante)
                st.rerun()

    df_vazio_lote = ler_aba_historico("VazioPALote")
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
                df_restante_lote = df_vazio_lote[
                    ~((df_vazio_lote["Data"] == del_data_lote) & (df_vazio_lote["Mapas"] == del_lote_chave))
                ]
                salvar_aba_historico("VazioPALote", df_restante_lote)
                st.rerun()

    st.divider()
    with st.expander("🧪 Modo Simulação (dados de teste)", expanded=False):
        st.caption("Gera retorno = saída perfeita pra todos os mapas Tianguá/Granja de uma data (via CONC.csv). Grava numa aba separada, nunca mistura com produção.")
        data_simulacao = st.date_input("Data pra simular", value=date.today() - timedelta(days=1), key="data_gerar_simulacao")

        col_sim1, col_sim2 = st.columns(2)
        if col_sim1.button("🧪 Gerar simulação", use_container_width=True):
            df_simulado = gerar_simulacao_perfeita(data_simulacao)
            if df_simulado.empty:
                st.warning("Nenhum mapa Tianguá/Granja encontrado pra essa data (confira o CONC.csv e o 03.07.13).")
            else:
                salvar_aba_historico(NOME_ABA_SIMULACAO, df_simulado)
                st.success(f"{len(df_simulado)} linha(s) simuladas geradas. Ative '🧪 Modo simulação' na sidebar pra ver.")

        if col_sim2.button("🗑️ Apagar simulação", use_container_width=True):
            salvar_aba_historico(NOME_ABA_SIMULACAO, pd.DataFrame(columns=["Data", "PA", "Mapa", "Familia", "Caixas", "Garrafas", "Garrafeiras", "Unidades"]))
            st.success("Dados de simulação apagados.")




# =========================================================================
# ABA DE CONCILIAÇÃO POR MAPA PA (VENDA x RETORNO CONFERENTE)
# =========================================================================
with aba_conciliacao:
    st.header("⚖️ Conciliação de Mapas PA (Saída vs. Retorno conferente)")
    st.caption("Mapas Tianguá/Granja do CONC.csv — aparecem mesmo sem conferência ainda.")

    df_concil = pd.DataFrame()  # fallback — usado pela aba Fechamento mesmo sem dados aqui
    if df_mapas_ag_sem_filtro_data is None or df_mapas_ag_sem_filtro_data.empty or not MAPAS_PA_CONC:
        st.info("⚠️ Aguardando dados. É necessário ter o relatório 03.07.13 e o CONC.csv (com mapas Tianguá/Granja) carregados.")
    else:
        # 1. VENDA (SAÍDA) — classificada pelo Código do Material via De Material.xlsx.
        # Usa df_mapas_ag_sem_filtro_data (relatório sem filtro de data — só o número do
        # mapa importa) e MAPAS_PA_CONC (o CONC.csv é quem dita quais mapas entram aqui,
        # independente de já terem sido conferidos ou não).
        # Só entram garrafa/barril soltos (não garrafeira), igual ao Retorno digitado
        # manualmente, que também só conta Garrafas+Unidades (nunca Garrafeiras).
        familia_tipo_venda = df_mapas_ag_sem_filtro_data["Material"].apply(lambda c: familia_tipo_por_codigo(c, lookup_ag))
        df_venda_ag = df_mapas_ag_sem_filtro_data.copy()
        df_venda_ag["Familia"] = familia_tipo_venda.apply(lambda ft: ft[0])
        df_venda_ag["Tipo"] = familia_tipo_venda.apply(lambda ft: ft[1])
        df_venda_ag = df_venda_ag[(df_venda_ag["Familia"] != "Outro") & (df_venda_ag["Tipo"] != "Garrafeira")]

        venda_agg_todos = df_venda_ag.groupby(["Mapa", "Familia"])["P Vazia"].sum().reset_index()
        venda_agg_todos.rename(columns={"P Vazia": "Qtd_Saida_Unidades"}, inplace=True)
        # Mapas em lote são tratados à parte mais abaixo — tira eles daqui pra não
        # duplicar (uma vez como linha individual, outra dentro da linha do lote).
        _mapas_em_lote_resolvidos = set(resolver_mapas(MAPAS_EM_LOTE))
        venda_agg = venda_agg_todos[venda_agg_todos["Mapa"].isin(MAPAS_PA_CONC - _mapas_em_lote_resolvidos)]

        # 2. RETORNO DO PA
        hist_vazio_pa = _hist_vazio_pa_bruto.copy()
        hist_vazio_pa["Mapa"] = hist_vazio_pa["Mapa"].apply(limpa_mapa)
        # Se o mapa foi consolidado (coluna MAPA CONSOLIDADO do CONC.csv), soma o
        # retorno de todos os originais que caem no mesmo consolidado antes de comparar
        # — senão cada um bateria errado sozinho contra a Saída combinada dos dois.
        hist_vazio_pa["Mapa"] = hist_vazio_pa["Mapa"].apply(resolver_mapa)

        if "Garrafas" not in hist_vazio_pa.columns: hist_vazio_pa["Garrafas"] = 0
        if "Unidades" not in hist_vazio_pa.columns: hist_vazio_pa["Unidades"] = 0

        hist_vazio_pa["Qtd_Retorno_Unidades"] = pd.to_numeric(hist_vazio_pa["Garrafas"], errors='coerce').fillna(0) + \
                                                pd.to_numeric(hist_vazio_pa["Unidades"], errors='coerce').fillna(0)

        # Alerta: um mapa não pode estar registrado individualmente E dentro de um lote
        # ao mesmo tempo — a Saída dele é excluída daqui (fica só no lote), então o
        # retorno individual apareceria inteiro como "Sobrou AG" sem explicação.
        _mapas_conflito = set(hist_vazio_pa["Mapa"].unique()) & _mapas_em_lote_resolvidos
        if _mapas_conflito:
            st.error(
                f"⚠️ Mapa(s) registrados tanto individualmente quanto em lote — isso zera a "
                f"Saída do lado individual: {', '.join(sorted(_mapas_conflito, key=lambda m: int(m) if m.isdigit() else 0))}. "
                "Apague o registro individual OU o lote pra esses mapas (seção 'Editar ou Apagar "
                "Registros' / 'Apagar um lote', mais abaixo nesta aba)."
            )

        # PA "dono" de cada mapa (já resolvido) — usado pra não deixar uma família
        # faltante "sumir" sob um rótulo genérico quando o conferente não digitou
        # retorno pra ela.
        mapa_pa_lookup = hist_vazio_pa.groupby("Mapa")["PA"].first().to_dict()

        colunas_agrupamento_vazio = ["Mapa", "PA", "Familia"]
        tem_data_vazio_pa = "Data" in hist_vazio_pa.columns
        if tem_data_vazio_pa:
            colunas_agrupamento_vazio.append("Data")
        vazio_agg = hist_vazio_pa.groupby(colunas_agrupamento_vazio)["Qtd_Retorno_Unidades"].sum().reset_index()

        # 3. CRUZAMENTO (MERGE) E CÁLCULO FÍSICO
        df_concil = pd.merge(venda_agg, vazio_agg, on=["Mapa", "Familia"], how="outer").fillna(0)

        df_concil["PA"] = df_concil.apply(
            lambda r: MAPA_PA_CLASSIFICACAO.get(r["Mapa"], mapa_pa_lookup.get(r["Mapa"], "Aguardando Retorno")) if r["PA"] == 0 else r["PA"],
            axis=1,
        )
        if tem_data_vazio_pa:
            df_concil["Data"] = df_concil["Data"].replace(0, "-")

        # Troca o número puro do Mapa pelo rótulo consolidado quando aplicável — só
        # muda a exibição, o cálculo acima já foi feito com o número resolvido.
        df_concil["Mapa"] = df_concil["Mapa"].apply(rotulo_mapa)

        # ==================== LOTE (vários mapas conferidos juntos) ====================
        # Pra cada lote+família digitado, soma a Saída de TODOS os mapas do lote (na
        # venda_agg_todos, sem o filtro de individuais) e compara com o total único
        # informado — gera 1 linha por lote+família, não 1 por mapa.
        if not _hist_lote_bruto.empty and "Mapas" in _hist_lote_bruto.columns:
            hist_lote = _hist_lote_bruto.copy()
            if "Garrafas" not in hist_lote.columns: hist_lote["Garrafas"] = 0
            if "Unidades" not in hist_lote.columns: hist_lote["Unidades"] = 0
            hist_lote["Qtd_Retorno_Unidades"] = pd.to_numeric(hist_lote["Garrafas"], errors="coerce").fillna(0) + \
                                                 pd.to_numeric(hist_lote["Unidades"], errors="coerce").fillna(0)

            colunas_grp_lote = ["Mapas", "PA", "Familia"] + (["Data"] if "Data" in hist_lote.columns else [])
            lote_retorno_agg = hist_lote.groupby(colunas_grp_lote)["Qtd_Retorno_Unidades"].sum().reset_index()

            linhas_lote_concil = []
            for _, linha_lote in lote_retorno_agg.iterrows():
                mapas_lote = mapas_da_lote(linha_lote["Mapas"])
                mapas_lote_resolvidos = resolver_mapas(mapas_lote)
                saida_lote = venda_agg_todos[
                    venda_agg_todos["Mapa"].isin(mapas_lote_resolvidos) & (venda_agg_todos["Familia"] == linha_lote["Familia"])
                ]["Qtd_Saida_Unidades"].sum()
                rotulo_mapas_lote = ", ".join(mapas_lote)
                if mapas_lote_resolvidos != mapas_lote:
                    rotulo_mapas_lote += f" (consolidados: {', '.join(mapas_lote_resolvidos)})"
                linha_final = {
                    "Mapa": f"Lote: {rotulo_mapas_lote}",
                    "PA": linha_lote["PA"],
                    "Familia": linha_lote["Familia"],
                    "Qtd_Saida_Unidades": saida_lote,
                    "Qtd_Retorno_Unidades": linha_lote["Qtd_Retorno_Unidades"],
                }
                if tem_data_vazio_pa:
                    linha_final["Data"] = linha_lote["Data"] if "Data" in linha_lote else "-"
                linhas_lote_concil.append(linha_final)

            if linhas_lote_concil:
                df_concil = pd.concat([df_concil, pd.DataFrame(linhas_lote_concil)], ignore_index=True).fillna(0)

        df_concil["Fator"] = df_concil["Familia"].apply(fator_conversao_caixas)

        df_concil["Caixas_Saida"] = df_concil["Qtd_Saida_Unidades"] // df_concil["Fator"]
        df_concil["Soltas_Saida"] = df_concil["Qtd_Saida_Unidades"] % df_concil["Fator"]

        df_concil["Caixas_Retorno"] = df_concil["Qtd_Retorno_Unidades"] // df_concil["Fator"]
        df_concil["Soltas_Retorno"] = df_concil["Qtd_Retorno_Unidades"] % df_concil["Fator"]

        df_concil["Diferença_Unidades"] = df_concil["Qtd_Retorno_Unidades"] - df_concil["Qtd_Saida_Unidades"]

        # Só garrafa de verdade converte pra caixa+garrafa solta — qualquer outra coisa
        # (Pallet, Chapatex, Barril) é sempre unidade, sem "gf" nenhum.
        FAMILIAS_GARRAFA = ("300ml", "600ml", "Verde 600", "1L")

        # 4. CRIAÇÃO DOS TEXTOS FORMATADOS
        def formata_cx_un(cx, un, fam):
            if fam not in FAMILIAS_GARRAFA:
                total = int(cx) + int(un)
                return f"{total} un" if total > 0 else "0"
            if cx == 0 and un == 0: return "0"
            res = []
            if cx > 0: res.append(f"{int(cx)} cx")
            if un > 0: res.append(f"{int(un)} gf")
            return " + ".join(res)

        df_concil["Saída"] = df_concil.apply(lambda r: formata_cx_un(r["Caixas_Saida"], r["Soltas_Saida"], r["Familia"]), axis=1)
        df_concil["Retorno"] = df_concil.apply(lambda r: formata_cx_un(r["Caixas_Retorno"], r["Soltas_Retorno"], r["Familia"]), axis=1)

        def formata_dif(dif, fam):
            item = "gf" if fam in FAMILIAS_GARRAFA else "un"
            if dif == 0: return "0"
            if dif > 0: return f"+{int(dif)} {item}"
            return f"{int(dif)} {item}"

        df_concil["Diferença"] = df_concil.apply(lambda r: formata_dif(r["Diferença_Unidades"], r["Familia"]), axis=1)

        # LÓGICA DE STATUS — 3 regras de negócio:
        # 1. Saiu (Previsto) e o conferente não digitou (ou digitou menos)  -> Faltou AG
        # 2. Não saiu e o conferente também não digitou                     -> nenhuma linha é gerada (ok)
        # 3. Não saiu e o conferente digitou (ou digitou mais que saiu)     -> Sobrou AG
        def status_conciliacao(row):
            dif = row["Diferença_Unidades"]
            if dif == 0:
                return "✅ Bateu"
            elif dif < 0:
                return "❌ Faltou AG"
            else:
                return "⚠️ Sobrou AG"

        df_concil["Status"] = df_concil.apply(status_conciliacao, axis=1)

        # 5. FILTROS E EXIBIÇÃO
        if tem_data_vazio_pa:
            col_filtro0, col_filtro1, col_filtro2, col_filtro3 = st.columns([1, 1, 1, 2])
            datas_disponiveis_pa = sorted(
                {d for d in df_concil["Data"].unique() if d != "-"},
                key=lambda d: pd.to_datetime(d, dayfirst=True, errors="coerce"),
                reverse=True,
            )
            data_filter = col_filtro0.selectbox("Filtrar por Data:", ["Todas"] + datas_disponiveis_pa, key="filtro_data_pa")
        else:
            col_filtro1, col_filtro2, col_filtro3 = st.columns([1, 1, 2])
            data_filter = "Todas"

        lista_pas = ["Todos"] + sorted(df_concil["PA"].unique().tolist())
        pa_filter = col_filtro1.selectbox("Filtrar por PA:", lista_pas)
        status_filter = col_filtro2.selectbox("Filtrar por Status:", ["Todos", "❌ Faltou AG", "⚠️ Sobrou AG", "✅ Bateu"])
        mapa_search = col_filtro3.text_input("🔍 Pesquisar Mapa Específico (opcional):", "")

        df_display = df_concil.copy()

        if tem_data_vazio_pa and data_filter != "Todas":
            df_display = df_display[df_display["Data"] == data_filter]

        if pa_filter != "Todos":
            df_display = df_display[df_display["PA"] == pa_filter]

        if status_filter != "Todos":
            df_display = df_display[df_display["Status"] == status_filter]

        if mapa_search.strip() != "":
            df_display = df_display[df_display["Mapa"].str.contains(limpa_mapa(mapa_search))]

        colunas_exibir_pa = ["Mapa"] + (["Data"] if tem_data_vazio_pa else []) + ["PA", "Familia", "Saída", "Retorno", "Diferença", "Status"]
        df_display = df_display[colunas_exibir_pa]
        df_display = df_display.sort_values(by=["Mapa", "Familia"])

        # =================================================================
        # RESUMO PRA ENVIAR (WhatsApp/print) — números por PA, contando
        # MAPAS (não itens/famílias). Um mapa só conta como "Bateu" se TODAS
        # as suas famílias bateram; se qualquer uma faltou, o mapa inteiro
        # conta como "Faltou" (Faltou > Sobrou > Bateu em prioridade).
        # =================================================================
        st.divider()
        titulo_resumo = "📋 Resumo pra enviar"
        if tem_data_vazio_pa and data_filter != "Todas":
            titulo_resumo += f" — {data_filter}"
        st.markdown(f"### {titulo_resumo}")

        qtd_mapas = df_display["Mapa"].nunique()
        st.caption(f"{qtd_mapas} mapa(s) conferido(s) nesse recorte.")

        pas_no_resumo = sorted(df_display["PA"].unique().tolist())
        ORDEM_STATUS_PA = ["❌ Faltou AG", "⚠️ Sobrou AG", "✅ Bateu"]

        # Por enquanto, sobra de Pallet/Chapatex não derruba o status do mapa no resumo
        # (só Garrafa/Garrafeira/Barril contam pra isso) — o item continua aparecendo
        # normal na tabela de detalhe, só não conta contra o mapa no card.
        FAMILIAS_IGNORAR_SOBRA_RESUMO = {"Pallet PBR1", "Pallet PBR2", "Chapatex"}

        for pa_nome in pas_no_resumo:
            df_pa_atual = df_display[df_display["PA"] == pa_nome].copy()
            st.markdown(
                f"""<div style="display:inline-block; background-color:#e2e6ea; color:#212529; font-weight:700; font-size:15px; padding:6px 16px; border-radius:20px; margin-bottom:10px;">📍 {pa_nome}</div>""",
                unsafe_allow_html=True,
            )
            df_pa_atual["Status_Resumo"] = df_pa_atual.apply(
                lambda r: "✅ Bateu" if (r["Status"] == "⚠️ Sobrou AG" and r["Familia"] in FAMILIAS_IGNORAR_SOBRA_RESUMO) else r["Status"],
                axis=1,
            )
            status_mapas_pa = df_pa_atual.groupby("Mapa")["Status_Resumo"].apply(
                lambda s: next((st_ for st_ in ORDEM_STATUS_PA if st_ in set(s)), ORDEM_STATUS_PA[-1])
            )
            bateu_pa = int((status_mapas_pa == "✅ Bateu").sum())
            faltou_pa = int((status_mapas_pa == "❌ Faltou AG").sum())
            sobrou_pa = int((status_mapas_pa == "⚠️ Sobrou AG").sum())

            # Linhas de detalhe (Mapa · Família · Diferença) mostradas dentro do card,
            # pra saber de cara QUAL item está causando a diferença sem abrir a tabela.
            itens_visiveis_pa = df_pa_atual[df_pa_atual["Status_Resumo"] != "✅ Bateu"]
            detalhes_faltou = [
                f"Mapa {r['Mapa']} · {r['Familia']} · {r['Diferença']}"
                for _, r in itens_visiveis_pa[itens_visiveis_pa["Status"] == "❌ Faltou AG"].iterrows()
            ]
            detalhes_sobrou = [
                f"Mapa {r['Mapa']} · {r['Familia']} · {r['Diferença']}"
                for _, r in itens_visiveis_pa[itens_visiveis_pa["Status"] == "⚠️ Sobrou AG"].iterrows()
            ]

            renderizar_cards_resumo([
                ("Bateram", bateu_pa, "verde", None),
                ("Faltou", faltou_pa, "vermelho", detalhes_faltou),
                ("Sobrou", sobrou_pa, "amarelo", detalhes_sobrou),
            ])
            st.write("")

        itens_problema = df_display[df_display["Status"] != "✅ Bateu"]
        if itens_problema.empty:
            st.success("🎉 Nenhuma diferença — tudo bateu certinho!")
        else:
            eh_lote = itens_problema["Mapa"].str.startswith("Lote:")
            itens_problema_individual = itens_problema[~eh_lote]
            itens_problema_lote = itens_problema[eh_lote].copy()

            if not itens_problema_individual.empty:
                st.markdown("**Itens com diferença (mapas individuais):**")
                colunas_resumo_prob = ["Mapa"] + (["Data"] if tem_data_vazio_pa else []) + ["PA", "Familia", "Diferença", "Status"]
                renderizar_tabela_limpa(itens_problema_individual[colunas_resumo_prob], colunas_resumo_prob)

            if not itens_problema_lote.empty:
                st.markdown("**Itens com diferença (lotes):**")
                itens_problema_lote["Mapas"] = itens_problema_lote["Mapa"].str.replace("Lote: ", "", regex=False)
                colunas_resumo_lote = ["Mapas"] + (["Data"] if tem_data_vazio_pa else []) + ["PA", "Familia", "Diferença", "Status"]
                renderizar_tabela_limpa(itens_problema_lote[colunas_resumo_lote], colunas_resumo_lote)

        with st.expander("📄 Ver tabela completa (todos os itens, inclusive os que bateram)"):
            renderizar_tabela_limpa(df_display, colunas_exibir_pa)

        st.caption("Faltou AG = saiu e não retornou (ou retornou menos). Sobrou AG = retornou mais que o previsto.")


# =========================================================================
# ABA DE CONCILIAÇÃO POR MAPA SEDE (Previsto x Realizado — sem conferente físico)
# =========================================================================
with aba_conciliacao_sede:
    st.header("🏢 Conciliação de Mapas Sede (Previsto vs. Realizado)")
    st.caption("Total Previsto x Total Realizado. Só mapas classificados como SEDE no CONC.csv.")
    df_concil_sede = pd.DataFrame()  # fallback — usado pela aba Fechamento mesmo sem dados aqui

    # =========================================================================
    # PREVISÃO DE CONTAGEM DO AG — quanto deveria estar de volta no armazém,
    # baseado no que saiu pra rota num dia (normalmente volta vazio no dia
    # seguinte). Considera TODOS os mapas do dia (Sede + Tianguá + Granja
    # juntos) e usa o relatório INTEIRO, sem o filtro de período da sidebar —
    # só a Data escolhida aqui importa.
    # =========================================================================
    st.markdown("### 📅 Previsão de Contagem do AG")
    st.caption("Quanto deveria voltar vazio, baseado no que saiu.")
    data_previsao = st.date_input(
        "Data em que a rota saiu",
        value=date.today() - timedelta(days=1),
        key="data_previsao_contagem",
    )

    if df_mapa_pa is None or df_mapa_pa.empty:
        st.error(f"Não encontrei '{ARQUIVO_MAPA_PA.name}' no Google Drive — a previsão usa essa planilha pra saber QUAIS mapas considerar naquele dia.")
    elif df_mapas_ag_sem_filtro_data is None or df_mapas_ag_sem_filtro_data.empty:
        st.info("⚠️ Aguardando dados do relatório 03.07.13.")
    else:
        data_previsao_str = data_previsao.strftime("%d/%m/%Y")
        mapas_previsao_originais = sorted(
            df_mapa_pa[df_mapa_pa["Data"] == data_previsao_str]["Mapa"].dropna().unique().tolist(),
            key=lambda m: int(m) if str(m).isdigit() else 0,
        )
        # Resolvido: mapas consolidados (ex: 257682+257685→257693) contam uma vez só —
        # a Saída deles só existe sob o número final no relatório.
        mapas_previsao = resolver_mapas(mapas_previsao_originais)

        if not mapas_previsao:
            st.warning(f"Nenhum mapa cadastrado em {data_previsao_str} na planilha '{ARQUIVO_MAPA_PA.name}'.")
        else:
            df_previsao = df_mapas_ag_sem_filtro_data[df_mapas_ag_sem_filtro_data["Mapa"].isin(mapas_previsao)].copy()
            mapas_encontrados = set(df_previsao["Mapa"].unique())
            mapas_faltando = [m for m in mapas_previsao if m not in mapas_encontrados]

            if mapas_faltando:
                st.warning(f"{len(mapas_faltando)} mapa(s) ainda não estão no relatório: {', '.join(mapas_faltando)}. Previsão incompleta.")
            else:
                st.caption(f"{len(mapas_previsao)} mapa(s) encontrados.")

            if df_previsao.empty:
                st.info("Nenhum dos mapas dessa data foi encontrado no relatório ainda.")
            else:
                previsao_agg = df_previsao.groupby("Material")["P Vazia"].sum().reset_index()
                previsao_agg = previsao_agg[previsao_agg["P Vazia"] > 0]

                if previsao_agg.empty:
                    st.info(f"Não houve saída de Vazio em {data_previsao_str}.")
                else:
                    if "Descricao" in df_previsao.columns:
                        desc_previsao = df_previsao.drop_duplicates(subset=["Material"])[["Material", "Descricao"]].rename(columns={"Descricao": "Desc_Previsao"})
                        previsao_agg = previsao_agg.merge(desc_previsao, on="Material", how="left")
                        previsao_agg["AG"] = [
                            com_apelido(cod, str(desc)) for cod, desc in zip(previsao_agg["Material"], previsao_agg["Desc_Previsao"].fillna(""))
                        ]
                    else:
                        previsao_agg["AG"] = previsao_agg["Material"]

                    fam_tipo_previsao = previsao_agg["Material"].apply(lambda c: familia_tipo_por_codigo(c, lookup_ag))
                    previsao_agg["Familia"] = fam_tipo_previsao.apply(lambda ft: ft[0])
                    previsao_agg["Tipo"] = fam_tipo_previsao.apply(lambda ft: ft[1])
                    previsao_agg["P Vazia"] = previsao_agg["P Vazia"].round(0).astype(int)
                    previsao_agg["Previsão de Retorno"] = previsao_agg.apply(
                        lambda r: formata_qtd_fisica(r["P Vazia"], r["Tipo"], r["Familia"]), axis=1
                    )

                    # Monta os dados pro visual "farol": famílias de garrafa (300/600/
                    # Verde/Litrão) viram caixas+soltas; tudo mais (garrafeira, pallet,
                    # chapatex, barril) vira um card escuro só com a unidade.
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

                    renderizar_farol_previsao(dados_familia_farol, dados_outros_farol)

    st.divider()

    if df_mapas_ag_sem_filtro_data is None or df_mapas_ag_sem_filtro_data.empty:
        st.info("⚠️ Aguardando dados. É necessário ter o relatório 03.07.13 carregado para cruzar.")
    elif not MAPAS_SEDE_CONC:
        st.info("⚠️ Nenhum mapa classificado como SEDE no CONC.csv (no período selecionado).")
    else:
        colunas_p = ["P Vazia"] + [cp for _, cp, cr in CATEGORIAS_AG_EXTRA if cp in df_mapas_ag_sem_filtro_data.columns]
        colunas_r = ["R Vazio"] + [cr for _, cp, cr in CATEGORIAS_AG_EXTRA if cr in df_mapas_ag_sem_filtro_data.columns]

        # Só os mapas que o CONC.csv classifica como SEDE entram aqui — não mais "tudo
        # que não foi digitado no PA". Isso evita mapa de Tianguá/Granja vazando pra cá
        # quando ele ainda não tem nenhuma conferência registrada.
        df_totais = df_mapas_ag_sem_filtro_data[df_mapas_ag_sem_filtro_data["Mapa"].isin(MAPAS_SEDE_CONC)].copy()
        df_totais["Qtd_Saida_554"] = df_totais[colunas_p].sum(axis=1)
        df_totais["Qtd_Retorno_654"] = df_totais[colunas_r].sum(axis=1)

        col_desc_rep = "Descricao" if "Descricao" in df_totais.columns else None
        desc_por_material = None
        if col_desc_rep:
            desc_por_material = df_totais.drop_duplicates(subset=["Material"])[["Material", col_desc_rep]].rename(columns={col_desc_rep: "Desc_AG"})

        data_por_mapa = {}
        if "Data" in df_totais.columns:
            tmp = df_totais.copy()
            tmp["_dt"] = pd.to_datetime(tmp["Data"], dayfirst=True, errors="coerce")
            tmp = tmp.dropna(subset=["_dt"])
            if not tmp.empty:
                idx = tmp.groupby("Mapa")["_dt"].idxmax()
                data_por_mapa = tmp.loc[idx].set_index("Mapa")["Data"].to_dict()

        df_concil_sede = df_totais.groupby(["Mapa", "Material"])[["Qtd_Saida_554", "Qtd_Retorno_654"]].sum().reset_index()

        df_concil_sede["Data"] = df_concil_sede["Mapa"].map(data_por_mapa).fillna("-")

        if desc_por_material is not None:
            df_concil_sede = df_concil_sede.merge(desc_por_material, on="Material", how="left")
            df_concil_sede["AG"] = [
                com_apelido(cod, str(desc)) for cod, desc in zip(df_concil_sede["Material"], df_concil_sede["Desc_AG"].fillna(""))
            ]
        else:
            df_concil_sede["AG"] = df_concil_sede["Material"]

        df_concil_sede["Qtd_Saida_554"] = df_concil_sede["Qtd_Saida_554"].round(0).astype(int)
        df_concil_sede["Qtd_Retorno_654"] = df_concil_sede["Qtd_Retorno_654"].round(0).astype(int)
        df_concil_sede["Diferença_Num"] = df_concil_sede["Qtd_Retorno_654"] - df_concil_sede["Qtd_Saida_554"]

        familia_tipo_serie = df_concil_sede["Material"].apply(lambda c: familia_tipo_por_codigo(c, lookup_ag))
        df_concil_sede["Familia"] = familia_tipo_serie.apply(lambda ft: ft[0])
        df_concil_sede["Tipo"] = familia_tipo_serie.apply(lambda ft: ft[1])

        df_concil_sede["Saída (Total)"] = df_concil_sede.apply(lambda r: formata_qtd_fisica(r["Qtd_Saida_554"], r["Tipo"], r["Familia"]), axis=1)
        df_concil_sede["Retorno (Total)"] = df_concil_sede.apply(lambda r: formata_qtd_fisica(r["Qtd_Retorno_654"], r["Tipo"], r["Familia"]), axis=1)
        df_concil_sede["Diferença"] = df_concil_sede.apply(lambda r: formata_diferenca_fisica(r["Diferença_Num"], r["Tipo"], r["Familia"]), axis=1)

        def status_sede(row):
            dif = row["Diferença_Num"]
            saida = row["Qtd_Saida_554"]
            retorno = row["Qtd_Retorno_654"]
            if saida == 0 and retorno > 0:
                return "🔎 Sem Saída"
            if saida > 0 and retorno == 0:
                return "⏳ Aguardando Retorno"
            if dif == 0:
                return "✅ Bateu"
            if dif < 0:
                return "❌ Faltou (não retornou)"
            return "⚠️ Sobrou no Retorno"

        df_concil_sede["Status"] = df_concil_sede.apply(status_sede, axis=1)

        mostrar_so_divergencias = st.checkbox("🔍 Mostrar só o que tem diferença (recomendado)", value=True, key="so_divergencias_sede")

        col_f1, col_f2, col_f3 = st.columns([1, 1, 2])
        status_filter_sede = col_f1.selectbox(
            "Filtrar por Status:",
            ["Todos", "❌ Faltou (não retornou)", "⚠️ Sobrou no Retorno", "🔎 Sem Saída", "⏳ Aguardando Retorno", "✅ Bateu"],
            key="status_sede",
        )
        mapa_search_sede = col_f2.text_input("🔍 Pesquisar Mapa:", "", key="mapa_sede")
        material_search_sede = col_f3.text_input("🔍 Pesquisar Material/AG:", "", key="material_sede")

        df_display_sede = df_concil_sede.copy()
        if mostrar_so_divergencias:
            df_display_sede = df_display_sede[df_display_sede["Status"] != "✅ Bateu"]
        if status_filter_sede != "Todos":
            df_display_sede = df_display_sede[df_display_sede["Status"] == status_filter_sede]
        if mapa_search_sede.strip():
            df_display_sede = df_display_sede[df_display_sede["Mapa"].str.contains(limpa_mapa(mapa_search_sede))]
        if material_search_sede.strip():
            df_display_sede = df_display_sede[df_display_sede["AG"].str.contains(material_search_sede, case=False, na=False)]

        colunas_exibir_sede = ["Mapa", "Data", "AG", "Saída (Total)", "Retorno (Total)", "Diferença", "Status"]
        df_display_sede = df_display_sede[colunas_exibir_sede].sort_values(by=["Mapa", "AG"])

        # =================================================================
        # RESUMO PRA ENVIAR (WhatsApp/print) — usa df_concil_sede inteiro
        # (não os filtros da tela), pra sempre dar o total real do recorte.
        # Conta MAPAS (não itens): cada mapa recebe o pior status entre
        # todos os seus itens (Faltou > Sobrou > Sem Saída > Aguardando >
        # Bateu), e só conta como Bateu se TODOS os itens baterem.
        # =================================================================
        st.divider()
        st.markdown("### 📋 Resumo pra enviar")

        qtd_mapas_sede = df_concil_sede["Mapa"].nunique()
        st.caption(f"{qtd_mapas_sede} mapa(s) nesse recorte.")

        ORDEM_STATUS_SEDE = [
            "❌ Faltou (não retornou)", "⚠️ Sobrou no Retorno",
            "🔎 Sem Saída", "⏳ Aguardando Retorno", "✅ Bateu",
        ]
        status_mapas_sede = status_por_mapa(df_concil_sede, ORDEM_STATUS_SEDE)

        bateu_sede = int((status_mapas_sede == "✅ Bateu").sum())
        faltou_sede = int((status_mapas_sede == "❌ Faltou (não retornou)").sum())
        sobrou_sede = int((status_mapas_sede == "⚠️ Sobrou no Retorno").sum())
        sem_saida_sede = int((status_mapas_sede == "🔎 Sem Saída").sum())
        aguardando_sede = int((status_mapas_sede == "⏳ Aguardando Retorno").sum())

        # Linhas de detalhe (Mapa · AG · Diferença) dentro dos cards — mesmo padrão
        # visual da aba Conciliação Mapas PA (badge de local + cards com detalhe).
        itens_visiveis_sede = df_concil_sede[df_concil_sede["Status"] != "✅ Bateu"]
        detalhes_faltou_sede = [
            f"Mapa {r['Mapa']} · {r['AG']} · {r['Diferença']}"
            for _, r in itens_visiveis_sede[itens_visiveis_sede["Status"] == "❌ Faltou (não retornou)"].iterrows()
        ]
        detalhes_sobrou_sede = [
            f"Mapa {r['Mapa']} · {r['AG']} · {r['Diferença']}"
            for _, r in itens_visiveis_sede[itens_visiveis_sede["Status"] == "⚠️ Sobrou no Retorno"].iterrows()
        ]
        detalhes_sem_saida_sede = [
            f"Mapa {r['Mapa']} · {r['AG']}"
            for _, r in itens_visiveis_sede[itens_visiveis_sede["Status"] == "🔎 Sem Saída"].iterrows()
        ]
        detalhes_aguardando_sede = [
            f"Mapa {r['Mapa']} · {r['AG']}"
            for _, r in itens_visiveis_sede[itens_visiveis_sede["Status"] == "⏳ Aguardando Retorno"].iterrows()
        ]

        st.markdown(
            """<div style="display:inline-block; background-color:#e2e6ea; color:#212529; font-weight:700; font-size:15px; padding:6px 16px; border-radius:20px; margin-bottom:10px;">📍 Sede</div>""",
            unsafe_allow_html=True,
        )
        renderizar_cards_resumo([
            ("Bateram", bateu_sede, "verde", None),
            ("Faltou", faltou_sede, "vermelho", detalhes_faltou_sede),
            ("Sobrou", sobrou_sede, "amarelo", detalhes_sobrou_sede),
            ("Sem Saída", sem_saida_sede, "azul", detalhes_sem_saida_sede),
            ("Aguardando", aguardando_sede, "cinza", detalhes_aguardando_sede),
        ])
        st.write("")

        itens_problema_sede = df_concil_sede[df_concil_sede["Status"] != "✅ Bateu"]
        if itens_problema_sede.empty:
            st.success("🎉 Nenhuma diferença — tudo bateu certinho!")
        else:
            st.markdown("**Itens com diferença:**")
            renderizar_tabela_limpa(
                itens_problema_sede[["Mapa", "Data", "AG", "Diferença", "Status"]],
                ["Mapa", "Data", "AG", "Diferença", "Status"],
            )

        with st.expander("📄 Ver tabela completa (respeitando os filtros da tela)"):
            renderizar_tabela_limpa(df_display_sede, colunas_exibir_sede)

        st.caption("Pra ver por espécie, use a aba 'Outras Categorias'.")


# =========================================================================
# ABA DE OUTRAS CATEGORIAS (Comodato, Devolução, Troca, Consignação, Rec. Consignação)
# =========================================================================
with aba_categorias_extra:
    st.header("📋 Divergências por Categoria")
    st.caption("Previsto x Realizado por categoria (Comodato, Devolução, Troca, Consignação).")

    MAPAS_CONC_TODOS = MAPAS_PA_CONC | MAPAS_SEDE_CONC
    if df_mapas_ag_sem_filtro_data is None or df_mapas_ag_sem_filtro_data.empty:
        st.info("⚠️ Aguardando dados do relatório 03.07.13.")
    elif not MAPAS_CONC_TODOS:
        st.info("⚠️ Nenhum mapa no CONC.csv (no período selecionado).")
    else:
        df_mapas_cat = df_mapas_ag_sem_filtro_data[df_mapas_ag_sem_filtro_data["Mapa"].isin(MAPAS_CONC_TODOS)]
        col_desc_cat = "Descricao" if "Descricao" in df_mapas_cat.columns else None

        linhas_cat = []
        for nome_cat, col_p, col_r in CATEGORIAS_AG_EXTRA:
            if col_p not in df_mapas_cat.columns or col_r not in df_mapas_cat.columns:
                continue
            agg = df_mapas_cat.groupby(["Mapa", "Material"])[[col_p, col_r]].sum().reset_index()
            agg = agg[(agg[col_p] != 0) | (agg[col_r] != 0)]
            if agg.empty:
                continue
            agg = agg.rename(columns={col_p: "Previsto", col_r: "Realizado"})
            agg["Categoria"] = nome_cat
            linhas_cat.append(agg[["Mapa", "Material", "Categoria", "Previsto", "Realizado"]])

        if not linhas_cat:
            st.info("Nenhum movimento registrado ainda em Comodato, Devolução, Troca, Consignação ou Rec. Consignação.")
        else:
            df_cat = pd.concat(linhas_cat, ignore_index=True)
            df_cat["Previsto"] = df_cat["Previsto"].round(0).astype(int)
            df_cat["Realizado"] = df_cat["Realizado"].round(0).astype(int)
            df_cat["Diferença_Num"] = df_cat["Realizado"] - df_cat["Previsto"]
            df_cat["Diferença"] = df_cat["Diferença_Num"].apply(lambda d: f"+{d}" if d > 0 else (f"{d}" if d < 0 else "0"))

            if col_desc_cat:
                desc_lookup_cat = df_mapas_cat.drop_duplicates(subset=["Material"]).set_index("Material")[col_desc_cat].to_dict()
                df_cat["AG"] = [com_apelido(cod, str(desc_lookup_cat.get(cod, cod))) for cod in df_cat["Material"]]
            else:
                df_cat["AG"] = df_cat["Material"]

            def status_categoria(row):
                if row["Diferença_Num"] == 0:
                    return "✅ Bateu"
                elif row["Diferença_Num"] < 0:
                    return "❌ Faltou"
                else:
                    return "⚠️ Sobrou"

            df_cat["Status"] = df_cat.apply(status_categoria, axis=1)

            col_fc1, col_fc2, col_fc3 = st.columns([1, 1, 2])
            categoria_filter = col_fc1.selectbox("Filtrar por Categoria:", ["Todas"] + [c[0] for c in CATEGORIAS_AG_EXTRA])
            status_filter_cat = col_fc2.selectbox("Filtrar por Status:", ["Todos", "❌ Faltou", "⚠️ Sobrou", "✅ Bateu"])
            mapa_search_cat = col_fc3.text_input("🔍 Pesquisar Mapa:", "", key="mapa_cat_extra")

            df_cat_display = df_cat.copy()
            if categoria_filter != "Todas":
                df_cat_display = df_cat_display[df_cat_display["Categoria"] == categoria_filter]
            if status_filter_cat != "Todos":
                df_cat_display = df_cat_display[df_cat_display["Status"] == status_filter_cat]
            if mapa_search_cat.strip():
                df_cat_display = df_cat_display[df_cat_display["Mapa"].str.contains(limpa_mapa(mapa_search_cat))]

            df_cat_display = df_cat_display[["Mapa", "AG", "Categoria", "Previsto", "Realizado", "Diferença", "Status"]].sort_values(["Mapa", "Categoria"])

            renderizar_tabela_limpa(
                df_cat_display,
                ["Mapa", "AG", "Categoria", "Previsto", "Realizado", "Diferença", "Status"],
            )


# =========================================================================
# ABA FECHAMENTO — junta PA + Sede numa visão única: Top 10 Faltas/Sobras,
# Justificativas (texto livre por item) e Mapa de Calor (variação diária).
# =========================================================================
with aba_fechamento:
    st.header("📊 Fechamento da Conciliação")
    st.caption("Visão única PA + Sede — Top Faltas/Sobras, justificativas e variação diária.")

    # --- Unifica PA (por Família) e Sede (por Item/AG) numa tabela só de diferenças ---
    partes_fechamento = []
    if not df_concil.empty and "Diferença_Unidades" in df_concil.columns:
        pa_unif = df_concil[df_concil["Diferença_Unidades"] != 0][["Mapa", "Familia", "Diferença_Unidades"]].copy()
        pa_unif = pa_unif.rename(columns={"Familia": "Item", "Diferença_Unidades": "Diferença"})
        pa_unif["Data"] = df_concil.get("Data", "-")
        pa_unif["Origem"] = "PA"
        partes_fechamento.append(pa_unif)
    if not df_concil_sede.empty and "Diferença_Num" in df_concil_sede.columns:
        sede_unif = df_concil_sede[df_concil_sede["Diferença_Num"] != 0][["Mapa", "AG", "Diferença_Num", "Data"]].copy()
        sede_unif = sede_unif.rename(columns={"AG": "Item", "Diferença_Num": "Diferença"})
        sede_unif["Origem"] = "Sede"
        partes_fechamento.append(sede_unif)

    df_fechamento = pd.concat(partes_fechamento, ignore_index=True) if partes_fechamento else pd.DataFrame(columns=["Mapa", "Item", "Diferença", "Data", "Origem"])

    if df_fechamento.empty:
        st.success("🎉 Nenhuma diferença registrada — nada pra fechar hoje.")
    else:
        # ================= TOP 10 FALTAS / TOP 10 SOBRAS (por quantidade) =================
        st.markdown("### 🔻🔺 Top 10 Faltas e Sobras (por quantidade)")
        totais_item = df_fechamento.groupby("Item")["Diferença"].sum().reset_index()

        top_faltas = totais_item[totais_item["Diferença"] < 0].sort_values("Diferença").head(10)
        top_sobras = totais_item[totais_item["Diferença"] > 0].sort_values("Diferença", ascending=False).head(10)

        col_falta, col_sobra = st.columns(2)
        with col_falta:
            st.markdown("**🔻 TOP 10 Faltas**")
            if top_faltas.empty:
                st.caption("Nenhuma falta no recorte.")
            else:
                df_tf = top_faltas.rename(columns={"Diferença": "Qtd"}).copy()
                df_tf["Qtd"] = df_tf["Qtd"].apply(lambda v: f"{int(v)}")
                renderizar_tabela_limpa(df_tf[["Item", "Qtd"]], ["Item", "Qtd"], col_status="")
        with col_sobra:
            st.markdown("**🔺 TOP 10 Sobras**")
            if top_sobras.empty:
                st.caption("Nenhuma sobra no recorte.")
            else:
                df_ts = top_sobras.rename(columns={"Diferença": "Qtd"}).copy()
                df_ts["Qtd"] = df_ts["Qtd"].apply(lambda v: f"+{int(v)}")
                renderizar_tabela_limpa(df_ts[["Item", "Qtd"]], ["Item", "Qtd"], col_status="")

        # ================= JUSTIFICATIVAS (texto livre por Mapa+Item+Data) =================
        st.divider()
        st.markdown("### 📝 Justificativas")

        df_justif_hist = ler_aba_historico("Justificativas")
        justif_lookup = {}
        if not df_justif_hist.empty:
            for _, r in df_justif_hist.iterrows():
                justif_lookup[(str(r.get("Data", "")), str(r.get("Mapa", "")), str(r.get("Item", "")))] = r.get("Justificativa", "")

        df_fechamento["Justificativa"] = df_fechamento.apply(
            lambda r: justif_lookup.get((str(r["Data"]), str(r["Mapa"]), str(r["Item"])), ""), axis=1
        )

        opcoes_justif = [
            f"{r['Mapa']} · {r['Item']} · {r['Diferença']:+.0f} ({r['Data']})"
            for _, r in df_fechamento.iterrows()
        ]
        if opcoes_justif:
            col_j1, col_j2 = st.columns([2, 3])
            item_escolhido = col_j1.selectbox("Item com diferença", opcoes_justif, key="justif_item_escolhido")
            texto_justif = col_j2.text_input("Justificativa", key="justif_texto")
            if st.button("💾 Salvar justificativa"):
                idx = opcoes_justif.index(item_escolhido)
                linha = df_fechamento.iloc[idx]
                nova_justif = pd.DataFrame([{
                    "Data": linha["Data"], "Mapa": linha["Mapa"], "Item": linha["Item"],
                    "Diferença": linha["Diferença"], "Justificativa": texto_justif,
                }])
                acumular_historico(nova_justif, "Justificativas", ["Data", "Mapa", "Item"])
                st.success("Justificativa salva.")
                st.rerun()

        df_fechamento_exib = df_fechamento.copy()
        df_fechamento_exib["Diferença"] = df_fechamento_exib["Diferença"].apply(lambda v: f"+{int(v)}" if v > 0 else f"{int(v)}")
        df_fechamento_exib["Justificativa"] = df_fechamento_exib["Justificativa"].replace("", "—")
        renderizar_tabela_limpa(
            df_fechamento_exib[["Mapa", "Data", "Origem", "Item", "Diferença", "Justificativa"]],
            ["Mapa", "Data", "Origem", "Item", "Diferença", "Justificativa"],
            col_status="",
        )

        # ================= MAPA DE CALOR (variação diária dos itens mais voláteis) =================
        st.divider()
        st.markdown("### 🌡️ Mapa de Calor — variação diária")
        st.caption("Diferença (Faltou/Sobrou) por item e por dia — os itens com mais impacto no período aparecem primeiro.")

        if df_fechamento["Data"].nunique() < 2:
            st.caption("Precisa de pelo menos 2 dias com diferença registrada pra montar o mapa de calor.")
        else:
            impacto_item = df_fechamento.groupby("Item")["Diferença"].apply(lambda s: s.abs().sum()).sort_values(ascending=False)
            itens_top_calor = impacto_item.head(10).index.tolist()

            pivot = df_fechamento[df_fechamento["Item"].isin(itens_top_calor)].pivot_table(
                index="Item", columns="Data", values="Diferença", aggfunc="sum", fill_value=0
            )
            datas_ordenadas = sorted(pivot.columns, key=lambda d: pd.to_datetime(d, dayfirst=True, errors="coerce"))
            pivot = pivot.reindex(columns=datas_ordenadas).reindex(itens_top_calor)

            maior_abs = max(pivot.abs().max().max(), 1)

            def _cor_celula(v: float) -> tuple[str, str]:
                intensidade = min(abs(v) / maior_abs, 1.0)
                if v < 0:
                    return f"rgba(163,45,45,{0.12 + intensidade * 0.6:.2f})", "#501313"
                if v > 0:
                    return f"rgba(15,110,86,{0.12 + intensidade * 0.6:.2f})", "#0F6E56"
                return "#F2F2F0", "#888780"

            cabecalho_calor = "".join(
                f'<th style="padding:6px 10px; font-size:11.5px; color:#888780; font-weight:600; text-align:center;">{d}</th>'
                for d in datas_ordenadas
            )
            linhas_calor = []
            for item_nome in itens_top_calor:
                celulas = f'<td style="padding:8px 10px; font-size:12.5px; white-space:nowrap;">{item_nome}</td>'
                for d in datas_ordenadas:
                    v = pivot.loc[item_nome, d] if item_nome in pivot.index and d in pivot.columns else 0
                    bg, fg = _cor_celula(v)
                    texto_v = f"+{int(v)}" if v > 0 else (f"{int(v)}" if v < 0 else "·")
                    celulas += f'<td style="padding:8px 10px; text-align:center; background:{bg}; color:{fg}; font-weight:600; font-size:12.5px;">{texto_v}</td>'
                linhas_calor.append(f"<tr>{celulas}</tr>")

            html_calor = (
                '<div style="overflow-x:auto;"><table style="width:100%; border-collapse:collapse;">'
                f'<thead><tr><th style="padding:6px 10px; text-align:left; font-size:11.5px; color:#888780;">Item</th>{cabecalho_calor}</tr></thead>'
                f'<tbody>{"".join(linhas_calor)}</tbody></table></div>'
            )
            st.markdown(html_calor, unsafe_allow_html=True)
