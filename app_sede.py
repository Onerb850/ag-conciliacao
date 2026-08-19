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


def extrair_dados_farol(df_subset: pd.DataFrame, lookup_map: dict) -> tuple[dict, dict]:
    """Processa um subconjunto de dados de saída e devolve os dicionários
    de garrafas (caixas+soltas) e outros itens (unidades) para o farol."""
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


with st.sidebar:
    st.caption("Fonte: 02.05.01.csv (atualiza sozinho a cada 5 min)")
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
        help="Isola TUDO (individual + lote) numa aba de teste separada, nunca mexe nos dados reais de produção.",
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
    """Acumula o CONC.csv atual (já limpo: Data/Mapa/PA/MapaConsolidado) num histórico
    permanente — assim, mesmo substituindo o CONC.csv todo dia, um mapa de um dia
    anterior que ainda não foi conferido continua aparecendo nas conciliações, em vez
    de sumir quando o arquivo do dia é trocado. Cacheado 5 min (igual carregar()) pra
    não gravar no Drive a cada interação do usuário na tela."""
    colunas_manter = [c for c in ["Data", "Mapa", "PA", "MapaConsolidado"] if c in _df_conc_limpo.columns]
    return acumular_historico(_df_conc_limpo[colunas_manter], "MapaPAHistorico", ["Data", "Mapa"])


if df_mapa_pa is not None and not df_mapa_pa.empty:
    # Acumula no histórico (MapaPAHistorico) — necessário pra Previsão de Contagem e
    # o Fechamento conseguirem consultar dias anteriores, mesmo depois do CONC.csv de
    # hoje substituir o de ontem. O risco de "mapa fantasma" (erro/teste de uma versão
    # antiga que nunca mais sai do histórico) é tratado à parte: o alerta de
    # integridade (sidebar) avisa quando isso acontece, e a ferramenta "Limpar
    # registro antigo" (aba Vazio por PA) deixa você remover na hora, se precisar.
    df_mapa_pa = _ingerir_conc_no_historico(df_mapa_pa)
else:
    df_mapa_pa = ler_aba_historico("MapaPAHistorico")

# O Excel guarda/relê "Mapa" e "MapaConsolidado" como número quando o texto parece um
# número puro — isso corrompe o tipo (vira NaN/float em vez de string) e quebra tudo
# que espera string. Relimpa sempre que o histórico vem de volta do Drive.
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
    # Normaliza os dois lados (planilha às vezes vem "TIANGUA" sem acento) antes de
    # comparar, senão "TIANGUA" != "Tianguá" e o mapa nunca é encontrado.
    pa_alvo_norm = _PA_NORMALIZADO.get(pa_alvo.strip().upper(), pa_alvo).upper()
    pa_bate = df_mapa_pa["PA"].apply(lambda v: _PA_NORMALIZADO.get(str(v).strip().upper(), v).upper() == pa_alvo_norm)
    sub = df_mapa_pa[(df_mapa_pa["Data"] == data_str) & pa_bate]
    return sorted(sub["Mapa"].dropna().unique().tolist(), key=lambda m: int(m) if str(m).isdigit() else 0)


# =========================================================================
# 02.05.01.csv — usado SÓ como fonte de Saída pra Previsão de Contagem (é puro
# "o que saiu tem que voltar", código de Operação 554). Sede e Outras Categorias
# continuam usando o 03.07.13, que traz Retorno e as categorias extras que o
# 02.05.01 não tem. Acumula num histórico próprio, mesmo padrão do CONC/03.07.13,
# pra sobreviver à substituição diária do arquivo.
# =========================================================================
ARQUIVO_020501 = PASTA_PROJETO / "02.05.01.csv"
NOME_ABA_020501_HISTORICO = "Relatorio020501Historico"


def parse_qtde_entrada_robusta(serie: pd.Series) -> pd.Series:
    """'2.592/00' -> 2592.00 — formato do 02.05.01: '.' separa milhar, '/' faz as
    vezes de vírgula decimal."""
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
    """Soma a Saída (02.05.01) de um conjunto de mapas, por família (300ml/600ml/1L,
    já com 600 Âmbar+Verde combinados) — usado pra mostrar Bateu/Faltou/Sobrou na hora,
    direto na tela de conferência, sem precisar ir na aba de Conciliação."""
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
    """Converte unidades soltas em 'X cx' ou 'X cx + Y un' — mesmo padrão usado no
    resto do app, pro status na hora não mostrar número gigante de garrafa solta."""
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
    """Igual formata_un_em_cx, mas com sinal — pra coluna Diferença."""
    sinal = "+" if dif > 0 else ("-" if dif < 0 else "")
    return f"{sinal}{formata_un_em_cx(abs(int(dif)), familia)}" if dif != 0 else "0"


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
# {mapa_resolvido: "dd/mm/aaaa"} — data do mapa segundo o CONC.csv, independente de
# ter conferência lançada ou não. Usada no Fechamento pra achar mapas sem retorno
# ainda (que não têm data de retorno pra filtrar), evitando que sumam da tela.
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

        # Verificação de integridade: todo mapa do CONC.csv (Tianguá, Granja e Sede)
        # deveria existir no 02.05.01 — única fonte de Saída do app agora. Se não
        # existir, a Saída conta como 0 (não "falta", só invisível), o que pode
        # mascarar ou distorcer números de conciliação/previsão.
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
    """Pra cada mapa Tianguá/Granja do CONC.csv na data escolhida, monta uma linha de
    retorno = saída exata (garrafas soltas, sem caixas/garrafeiras/unidades) — simula
    uma conferência 100% perfeita, só pra teste visual. Ignora consolidação de mapas
    (rara em Tianguá/Granja) usando direto o número resolvido."""
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

aba_vazio_pa, aba_conciliacao, aba_conciliacao_sede = st.tabs(
    ["Vazio por PA", "Conciliação Mapas PA", "Previsão Sede"]
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
# simulação); desativada, volta a mexer só na aba real "VazioPA". O Lote segue o
# mesmo toggle agora — simulação precisa estar 100% isolada de dado real (Lote
# incluso), senão um lote de teste antigo contamina silenciosamente a conta.
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

    st.markdown("### ✍️ Conferência Manual (um ou vários mapas)")
    st.caption("Digite um mapa só, ou vários separados por vírgula (pode misturar PA e Sede) — o total informado abaixo vale pra soma de todos eles juntos. A PA de cada mapa é detectada sozinha pelo CONC.csv.")

    col_data_manual, col_mapas_manual = st.columns([1, 2])
    data_manual = col_data_manual.date_input("Data da Descarga", value=date.today(), key="data_vazio_manual")
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

    # Já busca aqui fora (não só dentro do form) o que já está salvo pra esse
    # conjunto exato de mapas, e a Saída real deles — assim dá pra mostrar
    # Bateu/Faltou/Sobrou na hora, sem precisar salvar antes ou ir noutra aba.
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
    data_lote = col_data_l.date_input("Data da Descarga", value=date.today(), key="data_lote")
    pa_lote = col_pa_l.selectbox("PA", ["Tianguá", "Granja"], key="pa_lote")

    mapas_lote_auto = buscar_mapas_por_data_pa(data_lote, pa_lote)
    data_lote_str_atual = data_lote.strftime("%d/%m/%Y")
    mapas_chave_atual = ";".join(mapas_lote_auto) if mapas_lote_auto else ""

    # Busca aqui fora (não só dentro do form) o que já está salvo, e a Saída real —
    # assim dá pra mostrar Bateu/Faltou/Sobrou na hora.
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
        st.error(f"Não encontrei '{ARQUIVO_MAPA_PA.name}' no Google Drive nem histórico acumulado ainda — sem isso não dá pra buscar os mapas automaticamente.")
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
    # 📋 NOVO: RESUMO ACUMULADO DO DIA (LOGO ABAIXO DOS FORMULÁRIOS)
    # =====================================================================
    st.divider()
    st.markdown("### 📋 Resumo Acumulado do Dia (Saída vs Retorno)")
    st.caption("Acompanhe o status de todas as conferências que você já salvou para a data abaixo.")
    
    col_res1, col_res2 = st.columns([1, 3])
    data_resumo = col_res1.date_input("Data do Resumo", value=data_manual, key="data_resumo_diario")
    data_resumo_str = data_resumo.strftime("%d/%m/%Y")

    linhas_resumo_diario = []

    # 1. Pega do histórico de LOTES (que agora também recebe os lançamentos manuais)
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
                
                linhas_resumo_diario.append({
                    "PA": pa_r,
                    "Mapas/Lote": str(mapas_r).replace(";", ", "),
                    "Item": rotulo_conferencia(familia),
                    "Saída": formata_un_em_cx(saida_fam, familia),
                    "Retorno": formata_un_em_cx(retorno_fam, familia),
                    "Diferença": formata_dif_em_cx(dif, familia),
                    "Status": status_txt,
                })

    # 2. Pega do histórico INDIVIDUAL (legado/edições feitas na aba de exclusão)
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
                
                # Evita duplicar na tela se por acaso estiver salvo em ambos os formatos
                ja_existe = any(str(x["Mapas/Lote"]) == str(mapa_r) and x["Item"] == rotulo_conferencia(familia) for x in linhas_resumo_diario)
                if not ja_existe:
                    linhas_resumo_diario.append({
                        "PA": pa_r,
                        "Mapas/Lote": str(mapa_r),
                        "Item": rotulo_conferencia(familia),
                        "Saída": formata_un_em_cx(saida_fam, familia),
                        "Retorno": formata_un_em_cx(retorno_fam, familia),
                        "Diferença": formata_dif_em_cx(dif, familia),
                        "Status": status_txt,
                    })

    if linhas_resumo_diario:
        df_resumo = pd.DataFrame(linhas_resumo_diario)
        # Organiza para ficar bonito
        df_resumo = df_resumo.sort_values(by=["PA", "Mapas/Lote", "Item"])
        renderizar_tabela_limpa(df_resumo, ["PA", "Mapas/Lote", "Item", "Saída", "Retorno", "Diferença", "Status"])
    else:
        st.info(f"Nenhuma conferência (Saída ou Retorno) encontrada para a data {data_resumo_str}.")

    # =====================================================================
    # A PARTIR DAQUI: só telas de conferência (tabelas, edição, exclusão)
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

                st.caption("💡 **Atenção:** Para bebidas (300ml, 600ml, 1L), altere apenas as **Caixas**. As garrafas serão recalculadas sozinhas ao salvar. Para outros itens, altere as **Unidades**.")
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

                st.caption("💡 **Atenção:** Para bebidas (300ml, 600ml, 1L), altere apenas as **Caixas**. As garrafas serão recalculadas sozinhas ao salvar. Para outros itens, altere as **Unidades**.")
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
    with st.expander("☑️ Checklist de mapas lançados", expanded=False):
        st.caption("Marca sozinho quem já tem retorno digitado (individual ou lote) — os demais você marca manualmente conforme for conferindo.")
        col_chk1, col_chk2 = st.columns(2)
        data_checklist = col_chk1.date_input("Data", value=date.today(), key="data_checklist")
        pa_checklist = col_chk2.selectbox("PA", ["Tianguá", "Granja", "Sede"], key="pa_checklist")

        data_checklist_str = data_checklist.strftime("%d/%m/%Y")
        mapas_checklist = []
        if df_mapa_pa is not None and not df_mapa_pa.empty:
            pa_norm_alvo_chk = _PA_NORMALIZADO.get(pa_checklist.strip().upper(), pa_checklist).upper()
            pa_bate_chk = df_mapa_pa["PA"].apply(lambda v: _PA_NORMALIZADO.get(str(v).strip().upper(), v).upper() == pa_norm_alvo_chk)
            mapas_checklist = sorted(
                df_mapa_pa[(df_mapa_pa["Data"] == data_checklist_str) & pa_bate_chk]["Mapa"].dropna().unique().tolist(),
                key=lambda m: int(m) if str(m).isdigit() else 0,
            )

        if not mapas_checklist:
            st.info(f"Nenhum mapa cadastrado pra {pa_checklist} em {data_checklist_str} no CONC.csv.")
        else:
            MAPAS_JA_LANCADOS = MAPAS_INDIVIDUAIS | MAPAS_EM_LOTE
            df_checklist_hist = ler_aba_historico("MapaChecklist")
            checklist_manual = set()
            if not df_checklist_hist.empty and "Data" in df_checklist_hist.columns:
                checklist_manual = set(df_checklist_hist[df_checklist_hist["Data"] == data_checklist_str]["Mapa"].astype(str))

            estados_novos = {}
            cols_chk = st.columns(4)
            for i, mapa_c in enumerate(mapas_checklist):
                marcado_auto = mapa_c in MAPAS_JA_LANCADOS
                valor_inicial = marcado_auto or (mapa_c in checklist_manual)
                col = cols_chk[i % 4]
                estados_novos[mapa_c] = col.checkbox(
                    mapa_c, value=valor_inicial, key=f"chk_mapa_{data_checklist_str}_{pa_checklist}_{mapa_c}",
                    disabled=marcado_auto,
                    help="Já tem retorno lançado" if marcado_auto else "Marcação manual — só pra acompanhamento",
                )

            qtd_marcados = sum(estados_novos.values())
            st.caption(f"{qtd_marcados} de {len(mapas_checklist)} mapa(s) marcados.")

            if st.button("💾 Salvar checklist", key="salvar_checklist"):
                linhas_checklist = [
                    {"Data": data_checklist_str, "PA": pa_checklist, "Mapa": m, "Checado": 1}
                    for m, marcado in estados_novos.items() if marcado and m not in MAPAS_JA_LANCADOS
                ]
                if linhas_checklist:
                    acumular_historico(pd.DataFrame(linhas_checklist), "MapaChecklist", ["Data", "Mapa"])
                desmarcados = [m for m, marcado in estados_novos.items() if not marcado and m in checklist_manual]
                if desmarcados and not df_checklist_hist.empty:
                    df_checklist_restante = df_checklist_hist[
                        ~((df_checklist_hist["Data"] == data_checklist_str) & (df_checklist_hist["Mapa"].astype(str).isin(desmarcados)))
                    ]
                    salvar_aba_historico("MapaChecklist", df_checklist_restante)
                st.success("Checklist salvo.")
                st.rerun()

    st.divider()
    with st.expander("🧹 Limpar mapa 'fantasma' do histórico do CONC", expanded=False):
        st.caption(
            "O app acumula todo mapa já visto no CONC.csv, pra Previsão de Contagem e o Fechamento "
            "conseguirem consultar dias anteriores mesmo depois do arquivo de hoje substituir o de ontem. "
            "Se um mapa veio de um teste/erro de digitação corrigido depois, ele fica 'preso' aqui achando "
            "que ainda existe — o alerta de integridade (sidebar) avisa quando isso acontece. Remova aqui."
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
# ABA DE CONCILIAÇÃO POR MAPA PA (VENDA x RETORNO CONFERENTE)
# =========================================================================
with aba_conciliacao:
    st.header("⚖️ Conciliação de Mapas PA (Saída 02.05.01 vs. Retorno conferente)")
    st.caption("Mapas Tianguá/Granja do CONC.csv — aparecem mesmo sem conferência ainda.")

    df_concil = pd.DataFrame()  # fallback
    if df_020501_historico is None or df_020501_historico.empty or not MAPAS_PA_CONC:
        st.info("⚠️ Aguardando dados. É necessário ter o relatório 02.05.01 e o CONC.csv (com mapas Tianguá/Granja) carregados.")
    else:
        # 1. VENDA (SAÍDA)
        familia_tipo_venda = df_020501_historico["Material"].apply(lambda c: familia_tipo_por_codigo(c, lookup_ag))
        df_venda_ag = df_020501_historico.copy()
        df_venda_ag["Familia"] = familia_tipo_venda.apply(lambda ft: normaliza_600(ft[0]))
        df_venda_ag["Tipo"] = familia_tipo_venda.apply(lambda ft: ft[1])
        df_venda_ag = df_venda_ag[(df_venda_ag["Familia"] != "Outro") & (df_venda_ag["Tipo"] != "Garrafeira")]

        venda_agg_todos = df_venda_ag.groupby(["Mapa", "Familia"])["Qtde_Saida"].sum().reset_index()
        venda_agg_todos.rename(columns={"Qtde_Saida": "Qtd_Saida_Unidades"}, inplace=True)
        _mapas_em_lote_resolvidos = set(resolver_mapas(MAPAS_EM_LOTE))
        venda_agg = venda_agg_todos[venda_agg_todos["Mapa"].isin(MAPAS_PA_CONC - _mapas_em_lote_resolvidos)]

        # 2. RETORNO DO PA
        hist_vazio_pa = _hist_vazio_pa_bruto.copy()
        if "Mapa" not in hist_vazio_pa.columns:
            hist_vazio_pa = pd.DataFrame(columns=["Data", "PA", "Mapa", "Familia", "Caixas", "Garrafas", "Garrafeiras", "Unidades"])
        else:
            hist_vazio_pa["Mapa"] = hist_vazio_pa["Mapa"].apply(limpar_numero_robusto)
            hist_vazio_pa["Mapa"] = hist_vazio_pa["Mapa"].apply(resolver_mapa)

        if "Garrafas" not in hist_vazio_pa.columns: hist_vazio_pa["Garrafas"] = 0
        if "Unidades" not in hist_vazio_pa.columns: hist_vazio_pa["Unidades"] = 0

        hist_vazio_pa["Qtd_Retorno_Unidades"] = pd.to_numeric(hist_vazio_pa["Garrafas"], errors='coerce').fillna(0) + \
                                                pd.to_numeric(hist_vazio_pa["Unidades"], errors='coerce').fillna(0)

        _mapas_conflito = set(hist_vazio_pa["Mapa"].unique()) & _mapas_em_lote_resolvidos
        if _mapas_conflito:
            st.error(
                f"⚠️ Mapa(s) registrados tanto individualmente quanto em lote — isso zera a "
                f"Saída do lado individual: {', '.join(sorted(_mapas_conflito, key=lambda m: int(str(m)) if str(m).isdigit() else 0))}. "
                "Apague o registro individual OU o lote pra esses mapas."
            )

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

        df_concil["Mapa"] = df_concil["Mapa"].apply(rotulo_mapa)

        # ==================== LOTE ====================
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

        FAMILIAS_GARRAFA = ("300ml", "600ml", "Verde 600", "1L")

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

        # Se houver pesquisa, aprofundar na análise do mapa
        if mapa_search.strip() != "":
            df_display = df_display[df_display["Mapa"].str.contains(limpa_mapa(mapa_search))]

        colunas_exibir_pa = ["Mapa"] + (["Data"] if tem_data_vazio_pa else []) + ["PA", "Familia", "Saída", "Retorno", "Diferença", "Status"]
        df_display = df_display[colunas_exibir_pa]
        df_display = df_display.sort_values(by=["Mapa", "Familia"])

        # =====================================================================
        # 🔍 NOVO: RAIO-X DO MAPA (Deep Dive)
        # =====================================================================
        if mapa_search.strip() != "":
            mapa_limpo = limpa_mapa(mapa_search)
            st.divider()
            st.markdown(f"### 🔍 Raio-X do Mapa: `{mapa_limpo}`")
            
            if df_display.empty:
                st.warning("Este mapa não possui registros de conciliação no período/filtros selecionados.")
            else:
                st.markdown("**⚖️ Resultado da Conciliação Físico (Agrupado por Família)**")
                renderizar_tabela_limpa(df_display, colunas_exibir_pa)
                
                st.write("")
                col_saida, col_retorno = st.columns(2)
                
                with col_saida:
                    st.markdown("**📤 Saída Detalhada (Relatório 02.05.01)**")
                    mapa_res = resolver_mapa(mapa_limpo)
                    
                    df_saida_mapa = df_020501_historico[df_020501_historico["Mapa"] == mapa_res].copy() if (df_020501_historico is not None and not df_020501_historico.empty) else pd.DataFrame()
                    
                    if not df_saida_mapa.empty:
                        if "Descricao" in df_saida_mapa.columns:
                            df_saida_mapa["Item"] = df_saida_mapa.apply(lambda r: com_apelido(r["Material"], r["Descricao"]), axis=1)
                        else:
                            df_saida_mapa["Item"] = df_saida_mapa["Material"]
                            
                        df_saida_mapa = df_saida_mapa[["Item", "Qtde_Saida"]].rename(columns={"Qtde_Saida": "Qtd (un)"})
                        df_saida_mapa["Qtd (un)"] = df_saida_mapa["Qtd (un)"].astype(int)
                        df_saida_mapa = df_saida_mapa.sort_values("Qtd (un)", ascending=False)
                        st.dataframe(df_saida_mapa, width='stretch', hide_index=True)
                    else:
                        st.caption("Nenhum item registrado na saída para este mapa no relatório atual.")

                with col_retorno:
                    st.markdown("**🚚 Retorno Digitado (Físico)**")
                    encontrou_retorno = False
                    
                    def mapa_in_lote(mapas_str):
                        return mapa_limpo in mapas_da_lote(mapas_str)
                        
                    if not _hist_lote_bruto.empty and "Mapas" in _hist_lote_bruto.columns:
                        lotes_encontrados = _hist_lote_bruto[_hist_lote_bruto["Mapas"].apply(mapa_in_lote)]
                        if not lotes_encontrados.empty:
                            encontrou_retorno = True
                            st.info(f"📦 Conferido num lote junto com: {lotes_encontrados.iloc[0]['Mapas'].replace(';', ', ')}")
                            df_ret_exib = lotes_encontrados[["Familia", "Caixas", "Garrafas", "Garrafeiras", "Unidades"]].copy()
                            for c in ["Caixas", "Garrafas", "Garrafeiras", "Unidades"]:
                                df_ret_exib[c] = pd.to_numeric(df_ret_exib[c], errors='coerce').fillna(0).astype(int)
                            st.dataframe(df_ret_exib, width='stretch', hide_index=True)
                            
                    if not encontrou_retorno:
                        if not _hist_vazio_pa_bruto.empty and "Mapa" in _hist_vazio_pa_bruto.columns:
                            indiv_encontrados = _hist_vazio_pa_bruto[_hist_vazio_pa_bruto["Mapa"].astype(str) == str(mapa_limpo)]
                            if not indiv_encontrados.empty:
                                encontrou_retorno = True
                                df_ret_exib = indiv_encontrados[["Familia", "Caixas", "Garrafas", "Garrafeiras", "Unidades"]].copy()
                                for c in ["Caixas", "Garrafas", "Garrafeiras", "Unidades"]:
                                    df_ret_exib[c] = pd.to_numeric(df_ret_exib[c], errors='coerce').fillna(0).astype(int)
                                st.dataframe(df_ret_exib, width='stretch', hide_index=True)

                    if not encontrou_retorno:
                        st.caption("Nenhum retorno físico foi digitado para este mapa ainda.")

        else:
            with st.expander("📄 Ver tabela completa de conciliações", expanded=True):
                renderizar_tabela_limpa(df_display, colunas_exibir_pa)


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
        
        # Filtra mapas do dia no CONC.csv
        sub_conc_data = df_mapa_pa[df_mapa_pa["Data"] == data_previsao_str].copy()
        
        if sub_conc_data.empty:
            st.warning(f"Nenhum mapa cadastrado em {data_previsao_str} na planilha '{ARQUIVO_MAPA_PA.name}'.")
        else:
            # Mapeamento de cada mapa resolvido para seu Ponto de Apoio (Sede, Tianguá, Granja...)
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
                # Vincula a PA em cada registro de saída
                df_previsao["PA"] = df_previsao["Mapa"].apply(lambda m: pa_por_mapa_data.get(m, "Sede"))
                
                # Lista de PAs presentes no dia (ex: Sede, Tianguá, Granja)
                pas_no_dia = sorted(df_previsao["PA"].unique().tolist(), key=lambda p: (0 if p == "Sede" else 1, p))

                # Monta as sub-abas: Geral (Total) + Cada PA individualmente
                titulos_abas = ["🌐 Total Geral"] + [f"🏢 {p}" if p == "Sede" else f"🚛 {p}" for p in pas_no_dia]
                sub_abas_previsao = st.tabs(titulos_abas)

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
                        if dados_fam_pa or dados_outros_pa:
                            renderizar_farol_previsao(dados_fam_pa, dados_outros_pa)
                        else:
                            st.info(f"Não houve movimentação de vasilhame em {nome_pa} nesta data.")
