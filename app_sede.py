import streamlit as st
import pandas as pd
from datetime import date

from comum import (
    ARQUIVO_DE_MATERIAL,
    ARQUIVO_MAPAS_AG,
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
    GARRAFEIRA_FAMILIA,
    familia_normalizada_600,
    montar_lookup_ag_por_codigo,
    familia_tipo_por_codigo,
    CATEGORIAS_AG_EXTRA,
)

st.set_page_config(page_title="Conciliação de Mapas (AG)", layout="wide")
st.title("⚖️ Conciliação de Mapas (AG)")
st.caption("_\"Balança enganosa é abominação ao SENHOR, mas o peso justo lhe é agradável.\" — Provérbios 11:1_")

REGRAS_VAZIO = {
    "300ml": {"garrafas_por_cx": 23, "garrafeiras_por_cx": 1},
    "600ml": {"garrafas_por_cx": 24, "garrafeiras_por_cx": 1},
    "Verde 600": {"garrafas_por_cx": 24, "garrafeiras_por_cx": 1},
    "1L": {"garrafas_por_cx": 12, "garrafeiras_por_cx": 1},
}

# Mesma paleta usada em cor_linha_status, só que em cartão em vez de célula de tabela —
# usada nos resumos "pra enviar" das duas abas de conciliação.
CORES_RESUMO = {
    "verde": ("#EAF3DE", "#173404"),
    "vermelho": ("#FCEBEB", "#501313"),
    "amarelo": ("#FFF4D4", "#5A4000"),
    "azul": ("#CCE5FF", "#004085"),
    "cinza": ("#E9ECEF", "#495057"),
}


def renderizar_cards_resumo(itens: list[tuple[str, int, str]]) -> None:
    """itens: lista de (rotulo, valor, cor) — cor é uma chave de CORES_RESUMO."""
    colunas = st.columns(len(itens))
    for col, (rotulo, valor, cor) in zip(colunas, itens):
        bg, fg = CORES_RESUMO[cor]
        col.markdown(
            f"""<div style="background-color:{bg}; border-radius:10px; padding:14px 12px; text-align:center;">
                <div style="font-size:13px; color:{fg}; opacity:0.85; margin-bottom:2px;">{rotulo}</div>
                <div style="font-size:26px; font-weight:700; color:{fg}; line-height:1.2;">{valor}</div>
            </div>""",
            unsafe_allow_html=True,
        )


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
        "Considerar movimentações no período:",
        value=(date(2026, 8, 1), date.today()),
    )
    # date_input com range só retorna as duas datas depois que o usuário escolhe as duas
    # no calendário — enquanto só a primeira estiver selecionada, vem uma tupla de 1 item.
    if isinstance(intervalo_datas, tuple) and len(intervalo_datas) == 2:
        data_inicio, data_fim = intervalo_datas
    else:
        data_inicio = intervalo_datas[0] if isinstance(intervalo_datas, tuple) else intervalo_datas
        data_fim = date.today()

# --- De Material: usado pra classificar por Código e pra filtrar itens válidos de AG ---
df_de_material = carregar(ARQUIVO_DE_MATERIAL)
if df_de_material is not None and "Promax" in df_de_material.columns:
    df_de_material["Promax"] = normalizar_codigo(df_de_material["Promax"])
lookup_ag = montar_lookup_ag_por_codigo(df_de_material) if df_de_material is not None else {}

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

    if "Data" in df_mapas_ag.columns:
        _dt = pd.to_datetime(df_mapas_ag["Data"], dayfirst=True, errors="coerce")
        df_mapas_ag = df_mapas_ag[(_dt >= pd.Timestamp(data_inicio)) & (_dt <= pd.Timestamp(data_fim))]

_periodo_str = f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
with st.sidebar:
    if df_mapas_ag is None:
        st.error(f"Não encontrei '{ARQUIVO_MAPAS_AG.name}' no Google Drive.")
    elif df_mapas_ag.empty:
        st.warning(f"'{ARQUIVO_MAPAS_AG.name}' carregado, mas nenhuma linha entre {_periodo_str}.")
    else:
        st.success(f"{ARQUIVO_MAPAS_AG.name}: {len(df_mapas_ag)} linha(s) entre {_periodo_str}.")

aba_vazio_pa, aba_conciliacao, aba_conciliacao_sede, aba_categorias_extra = st.tabs(
    ["Vazio por PA", "Conciliação Mapas PA", "Conciliação Mapas Sede", "Outras Categorias"]
)

# Roteamento: um mapa só entra na Conciliação Mapas PA se foi digitado na aba 'Vazio por
# PA' pelo conferente; todo o resto cai na 'Conciliação Mapas Sede'. Esse histórico
# (VazioPA) é pequeno — só o que o conferente digitou — e continua sendo salvo normalmente.
_hist_vazio_pa_bruto = ler_aba_historico("VazioPA")
if not _hist_vazio_pa_bruto.empty and "Mapa" in _hist_vazio_pa_bruto.columns:
    MAPAS_COM_CONFERENCIA_PA = set(_hist_vazio_pa_bruto["Mapa"].apply(limpa_mapa).unique())
else:
    MAPAS_COM_CONFERENCIA_PA = set()


# =========================================================================
# ABA VAZIO POR PA (conferência física digitada pelo conferente)
# =========================================================================
with aba_vazio_pa:
    st.caption("Conferência do vazio por PA e mapa. Alimenta a Conciliação Mapas PA, ao lado.")

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
                    acumular_historico(pd.DataFrame(linhas_pa), "VazioPA", ["Data", "PA", "Mapa", "Familia"])
                    st.success(f"✅ Retorno do mapa {mapa_numero} salvo com sucesso!")
                else:
                    st.warning("Nenhuma quantidade foi informada para salvar.")

    df_vazio_pa = ler_aba_historico("VazioPA")
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
                    st.caption(f"Editando: Mapa {edit_mapa} · {pa_atual} · {edit_data} · {edit_familia}")
                else:
                    linha_atual = {}
                    pa_atual = pa_padrao
                    st.caption(f"Adicionando item novo: Mapa {edit_mapa} · {pa_atual} · {edit_data} · {edit_familia} (ainda não digitado)")

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
                    acumular_historico(nova_linha, "VazioPA", ["Data", "PA", "Mapa", "Familia"])
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
                salvar_aba_historico("VazioPA", df_restante)
                st.rerun()


# =========================================================================
# ABA DE CONCILIAÇÃO POR MAPA PA (VENDA x RETORNO CONFERENTE)
# =========================================================================
with aba_conciliacao:
    st.header("⚖️ Conciliação de Mapas PA (Saída vs. Retorno conferente)")
    st.caption("Cruza as quantidades físicas previstas (saída, do relatório 03.07.13) com o que foi conferido no retorno do PA. Só entram aqui mapas que foram digitados na aba 'Vazio por PA'.")

    if df_mapas_ag is None or df_mapas_ag.empty or _hist_vazio_pa_bruto.empty:
        st.info("⚠️ Aguardando dados. É necessário ter o relatório 03.07.13 carregado e algum retorno digitado na aba 'Vazio por PA' para fazer o cruzamento.")
    else:
        # 1. VENDA (SAÍDA) — classificada pelo Código do Material via De Material.xlsx.
        # Só entram garrafa/barril soltos (não garrafeira), igual ao Retorno digitado
        # manualmente, que também só conta Garrafas+Unidades (nunca Garrafeiras).
        familia_tipo_venda = df_mapas_ag["Material"].apply(lambda c: familia_tipo_por_codigo(c, lookup_ag))
        df_venda_ag = df_mapas_ag.copy()
        df_venda_ag["Familia"] = familia_tipo_venda.apply(lambda ft: ft[0])
        df_venda_ag["Tipo"] = familia_tipo_venda.apply(lambda ft: ft[1])
        df_venda_ag = df_venda_ag[(df_venda_ag["Familia"] != "Outro") & (df_venda_ag["Tipo"] != "Garrafeira")]

        venda_agg = df_venda_ag.groupby(["Mapa", "Familia"])["P Vazia"].sum().reset_index()
        venda_agg.rename(columns={"P Vazia": "Qtd_Saida_Unidades"}, inplace=True)
        venda_agg = venda_agg[venda_agg["Mapa"].isin(MAPAS_COM_CONFERENCIA_PA)]

        # 2. RETORNO DO PA
        hist_vazio_pa = _hist_vazio_pa_bruto.copy()
        hist_vazio_pa["Mapa"] = hist_vazio_pa["Mapa"].apply(limpa_mapa)

        if "Garrafas" not in hist_vazio_pa.columns: hist_vazio_pa["Garrafas"] = 0
        if "Unidades" not in hist_vazio_pa.columns: hist_vazio_pa["Unidades"] = 0

        hist_vazio_pa["Qtd_Retorno_Unidades"] = pd.to_numeric(hist_vazio_pa["Garrafas"], errors='coerce').fillna(0) + \
                                                pd.to_numeric(hist_vazio_pa["Unidades"], errors='coerce').fillna(0)

        # PA "dono" de cada mapa — usado pra não deixar uma família faltante "sumir" sob
        # um rótulo genérico quando o conferente não digitou retorno pra ela.
        mapa_pa_lookup = hist_vazio_pa.groupby("Mapa")["PA"].first().to_dict()

        colunas_agrupamento_vazio = ["Mapa", "PA", "Familia"]
        tem_data_vazio_pa = "Data" in hist_vazio_pa.columns
        if tem_data_vazio_pa:
            colunas_agrupamento_vazio.append("Data")
        vazio_agg = hist_vazio_pa.groupby(colunas_agrupamento_vazio)["Qtd_Retorno_Unidades"].sum().reset_index()

        # 3. CRUZAMENTO (MERGE) E CÁLCULO FÍSICO
        df_concil = pd.merge(venda_agg, vazio_agg, on=["Mapa", "Familia"], how="outer").fillna(0)

        df_concil["PA"] = df_concil.apply(
            lambda r: mapa_pa_lookup.get(r["Mapa"], "Aguardando Retorno") if r["PA"] == 0 else r["PA"],
            axis=1,
        )
        if tem_data_vazio_pa:
            df_concil["Data"] = df_concil["Data"].replace(0, "-")

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
            data_filter = col_filtro0.selectbox("Filtrar por Data:", ["Todas"] + datas_disponiveis_pa)
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
            renderizar_cards_resumo([
                ("Bateram", bateu_pa, "verde"),
                ("Faltou", faltou_pa, "vermelho"),
                ("Sobrou", sobrou_pa, "amarelo"),
            ])
            st.write("")

        itens_problema = df_display[df_display["Status"] != "✅ Bateu"]
        if itens_problema.empty:
            st.success("🎉 Nenhuma diferença — tudo bateu certinho!")
        else:
            st.markdown("**Itens com diferença:**")
            colunas_resumo_prob = ["Mapa"] + (["Data"] if tem_data_vazio_pa else []) + ["PA", "Familia", "Diferença", "Status"]
            st.dataframe(
                itens_problema[colunas_resumo_prob].style.map(cor_linha_status, subset=["Status"]),
                use_container_width=True, hide_index=True,
            )

        with st.expander("📄 Ver tabela completa (todos os itens, inclusive os que bateram)"):
            st.dataframe(
                df_display.style.map(cor_linha_status, subset=["Status"]),
                use_container_width=True, hide_index=True
            )

        st.caption("Nota 1: Mapas com status 'Faltou AG' saíram no Previsto e não tiveram (ou tiveram menos) retorno digitado na aba 'Vazio por PA'.")
        st.caption("Nota 2: Mapas com status 'Sobrou AG' foram conferidos no PA com quantidade maior do que o Previsto (incluindo casos em que nada saiu, mas algo foi digitado). A diferença é sempre exibida na menor unidade física (garrafas ou unidades soltas).")


# =========================================================================
# ABA DE CONCILIAÇÃO POR MAPA SEDE (Previsto x Realizado — sem conferente físico)
# =========================================================================
with aba_conciliacao_sede:
    st.header("🏢 Conciliação de Mapas Sede (Previsto vs. Realizado)")
    st.caption("Cruza item a item o Total Previsto com o Total Realizado (soma de Vazio + Comodato + Devolução + Troca + Consignação + Rec. Consignação) — não importa em qual espécie o item saiu ou voltou, só o total. Só mapas NÃO digitados na aba 'Vazio por PA'.")

    if df_mapas_ag is None or df_mapas_ag.empty:
        st.info("⚠️ Aguardando dados. É necessário ter o relatório 03.07.13 carregado para cruzar.")
    else:
        colunas_p = ["P Vazia"] + [cp for _, cp, cr in CATEGORIAS_AG_EXTRA if cp in df_mapas_ag.columns]
        colunas_r = ["R Vazio"] + [cr for _, cp, cr in CATEGORIAS_AG_EXTRA if cr in df_mapas_ag.columns]

        df_totais = df_mapas_ag.copy()
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
        df_concil_sede = df_concil_sede[~df_concil_sede["Mapa"].isin(MAPAS_COM_CONFERENCIA_PA)]

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

        renderizar_cards_resumo([
            ("Bateram", bateu_sede, "verde"),
            ("Faltou", faltou_sede, "vermelho"),
            ("Sobrou", sobrou_sede, "amarelo"),
            ("Sem Saída", sem_saida_sede, "azul"),
            ("Aguardando", aguardando_sede, "cinza"),
        ])

        itens_problema_sede = df_concil_sede[df_concil_sede["Status"] != "✅ Bateu"]
        if itens_problema_sede.empty:
            st.success("🎉 Nenhuma diferença — tudo bateu certinho!")
        else:
            st.markdown("**Itens com diferença:**")
            st.dataframe(
                itens_problema_sede[["Mapa", "Data", "AG", "Diferença", "Status"]].style.map(cor_linha_status, subset=["Status"]),
                use_container_width=True, hide_index=True,
            )

        with st.expander("📄 Ver tabela completa (respeitando os filtros da tela)"):
            st.dataframe(
                df_display_sede.style.map(cor_linha_status, subset=["Status"]),
                use_container_width=True, hide_index=True,
            )

        st.caption("Nota: Saída/Retorno somam Vazio + Comodato + Devolução + Troca + Consignação + Rec. Consignação. Pra ver em qual espécie está a diferença, use a aba 'Outras Categorias' filtrando pelo mesmo mapa.")

        # =================================================================
        # CONFERÊNCIA CRUZADA: GARRAFA × GARRAFEIRA, POR MAPA
        # =================================================================
        st.divider()
        st.markdown("### ⚠️ Conferência Garrafa × Garrafeira (por Mapa)")
        st.caption(
            "Compara, por mapa, quantas caixas de garrafa (convertidas a partir da quantidade bruta) "
            "saíram/retornaram contra quantas garrafeiras saíram/retornaram — mesmo com Saída = Retorno "
            "em cada item, pode haver descompasso entre os dois (ex: faturar 1 garrafeira pra 48 garrafas)."
        )

        linhas_diverg = []
        for mapa_val, grupo in df_concil_sede.groupby("Mapa"):
            garrafas_grp = grupo[grupo["Tipo"] == "Garrafa"].copy()
            garrafeiras_grp = grupo[grupo["Tipo"] == "Garrafeira"].copy()
            if garrafas_grp.empty and garrafeiras_grp.empty:
                continue

            garrafas_grp["FamiliaNorm"] = garrafas_grp["Familia"].apply(familia_normalizada_600)
            garrafeiras_grp["FamiliaNorm"] = garrafeiras_grp["Material"].map(GARRAFEIRA_FAMILIA)

            familias_presentes = set(garrafas_grp["FamiliaNorm"].dropna()) | set(garrafeiras_grp["FamiliaNorm"].dropna())
            for fam in familias_presentes:
                if fam == "Outro" or fam is None:
                    continue
                fator = int(fator_conversao_caixas(fam))
                sg_saida = garrafas_grp.loc[garrafas_grp["FamiliaNorm"] == fam, "Qtd_Saida_554"].sum()
                sg_retorno = garrafas_grp.loc[garrafas_grp["FamiliaNorm"] == fam, "Qtd_Retorno_654"].sum()
                sgf_saida = garrafeiras_grp.loc[garrafeiras_grp["FamiliaNorm"] == fam, "Qtd_Saida_554"].sum()
                sgf_retorno = garrafeiras_grp.loc[garrafeiras_grp["FamiliaNorm"] == fam, "Qtd_Retorno_654"].sum()

                for fluxo, qtd_garrafa, qtd_garrafeira in [
                    ("Saída", sg_saida, sgf_saida), ("Retorno", sg_retorno, sgf_retorno)
                ]:
                    if qtd_garrafa == 0 and qtd_garrafeira == 0:
                        continue
                    cx_equiv = int(qtd_garrafa) // fator
                    gf_soltas = int(qtd_garrafa) % fator
                    dif = int(qtd_garrafeira) - cx_equiv

                    if dif == 0:
                        status = "✅ Bateu"
                    elif dif > 0:
                        status = "⚠️ Garrafeira a mais"
                    else:
                        status = "⚠️ Garrafeira a menos"

                    linhas_diverg.append({
                        "Mapa": mapa_val, "Família": fam, "Fluxo": fluxo,
                        "Garrafas (un)": int(qtd_garrafa), "Caixas equiv.": cx_equiv,
                        "Garrafas soltas": gf_soltas, "Garrafeiras": int(qtd_garrafeira),
                        "Diferença": dif, "Status": status,
                    })

        if linhas_diverg:
            df_diverg = pd.DataFrame(linhas_diverg).sort_values(["Mapa", "Família", "Fluxo"])

            status_filter_diverg = st.selectbox(
                "Filtrar por Status:",
                ["Todos", "⚠️ Garrafeira a mais", "⚠️ Garrafeira a menos", "✅ Bateu"],
                key="status_diverg",
            )
            df_diverg_display = df_diverg if status_filter_diverg == "Todos" else df_diverg[df_diverg["Status"] == status_filter_diverg]

            st.dataframe(
                df_diverg_display.style.map(cor_linha_status, subset=["Status"]),
                use_container_width=True, hide_index=True,
            )
            st.caption(
                "Diferença > 0 = saiu/retornou garrafeira a mais do que caixas de garrafa fechadas. "
                "Diferença < 0 = faltou garrafeira pra cobrir as caixas de garrafa."
            )
        else:
            st.caption("Nenhum mapa com garrafa e/ou garrafeira pra comparar ainda.")


# =========================================================================
# ABA DE OUTRAS CATEGORIAS (Comodato, Devolução, Troca, Consignação, Rec. Consignação)
# =========================================================================
with aba_categorias_extra:
    st.header("📋 Divergências por Categoria")
    st.caption(
        "Previsto x Realizado, direto do relatório 03.07.13, pra cada categoria além de Vazio "
        "(essas não passam pela conferência manual do 'Vazio por PA' — comparação direta do relatório)."
    )

    if df_mapas_ag is None or df_mapas_ag.empty:
        st.info("⚠️ Aguardando dados do relatório 03.07.13.")
    else:
        col_desc_cat = "Descricao" if "Descricao" in df_mapas_ag.columns else None

        linhas_cat = []
        for nome_cat, col_p, col_r in CATEGORIAS_AG_EXTRA:
            if col_p not in df_mapas_ag.columns or col_r not in df_mapas_ag.columns:
                continue
            agg = df_mapas_ag.groupby(["Mapa", "Material"])[[col_p, col_r]].sum().reset_index()
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
            df_cat["Diferença"] = df_cat["Realizado"] - df_cat["Previsto"]

            if col_desc_cat:
                desc_lookup_cat = df_mapas_ag.drop_duplicates(subset=["Material"]).set_index("Material")[col_desc_cat].to_dict()
                df_cat["AG"] = [com_apelido(cod, str(desc_lookup_cat.get(cod, cod))) for cod in df_cat["Material"]]
            else:
                df_cat["AG"] = df_cat["Material"]

            def status_categoria(row):
                if row["Diferença"] == 0:
                    return "✅ Bateu"
                elif row["Diferença"] < 0:
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

            st.dataframe(
                df_cat_display.style.map(cor_linha_status, subset=["Status"]),
                use_container_width=True, hide_index=True,
            )

            st.caption("Cada linha é um Mapa+Item+Categoria com movimento previsto e/ou realizado — itens com Previsto=Realizado=0 nessa categoria não aparecem aqui.")
