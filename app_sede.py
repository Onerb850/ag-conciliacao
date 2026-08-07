import streamlit as st
import pandas as pd
from datetime import date

from comum import (
    ARQUIVO_DE_MATERIAL,
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
    valor_mais_recente_por_grupo,
    montar_total_previsto_realizado,
    CATEGORIAS_AG_EXTRA,
    NOME_ABA_CATEGORIAS_EXTRA,
)

st.set_page_config(page_title="Conciliação de Mapas (AG)", layout="wide")
st.title("⚖️ Conciliação de Mapas (AG)")
st.caption(
    "Este app só lê o histórico já salvo (historico_ag.xlsx) — não processa CSVs brutos, por isso é bem mais leve. "
    "Pra alimentar Venda (Previsto x Realizado), use o app 'Operacional'."
)

REGRAS_VAZIO = {
    "300ml": {"garrafas_por_cx": 23, "garrafeiras_por_cx": 1},
    "600ml": {"garrafas_por_cx": 24, "garrafeiras_por_cx": 1},
    "Verde 600": {"garrafas_por_cx": 24, "garrafeiras_por_cx": 1},
    "1L": {"garrafas_por_cx": 12, "garrafeiras_por_cx": 1},
}

# Lookup Código -> (Familia, Tipo), montado a partir da descrição mestre do De
# Material.xlsx — é lido aqui direto (app leve, arquivo pequeno) pra classificar
# os itens de Venda/Retorno pelo Código, em vez de interpretar a descrição
# abreviada de cada relatório (mais confiável, mesma fonte usada no app Operacional).
df_de_material = carregar(ARQUIVO_DE_MATERIAL)
if df_de_material is not None and "Promax" in df_de_material.columns:
    df_de_material["Promax"] = normalizar_codigo(df_de_material["Promax"])
lookup_ag = montar_lookup_ag_por_codigo(df_de_material) if df_de_material is not None else {}

aba_vazio_pa, aba_conciliacao, aba_conciliacao_sede, aba_categorias_extra = st.tabs(
    ["Vazio por PA", "Conciliação Mapas PA", "Conciliação Mapas Sede", "Outras Categorias"]
)

# Roteamento entre as duas conciliações: um mapa só entra na Conciliação Mapas PA se
# ele foi digitado ali pelo conferente (aba Vazio por PA); todo o resto (mapas que nunca
# tiveram digitação de PA) cai automaticamente na Conciliação Mapas Sede (Previsto vs. Realizado).
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
    st.caption("Cruza as quantidades físicas previstas (saída) com o que foi conferido no retorno do PA. Só entram aqui mapas que foram digitados na aba 'Vazio por PA' — os demais são conciliados na aba 'Conciliação Mapas Sede'.")

    hist_venda = ler_aba_historico("Venda")
    hist_vazio_pa = ler_aba_historico("VazioPA")

    if hist_venda.empty or hist_vazio_pa.empty:
        st.info("⚠️ Aguardando dados. É necessário ter histórico de Venda (Previsto) e de Retorno (Vazio por PA) salvos para fazer o cruzamento.")
    else:
        # 1. VENDA (SAÍDA) — classificada pelo Código do Material via De Material.xlsx,
        # não pela descrição abreviada do relatório (mais confiável).
        hist_venda = hist_venda.copy()
        # 1. VENDA (SAÍDA) — pega o valor mais recente por Mapa+Material (não soma todas
        # as datas: o 03.07.13 já traz o total acumulado do mapa, então somar histórico
        # de dias diferentes infla o número se o mapa foi processado mais de uma vez).
        # Classificação de Familia/Tipo pelo Código, via De Material.xlsx.
        hist_venda_latest = valor_mais_recente_por_grupo(hist_venda, ["Mapa", "Material"], "Data", "Qtd. Vendida/Movimentada")
        familia_tipo_venda = hist_venda_latest["Material"].apply(lambda c: familia_tipo_por_codigo(c, lookup_ag))
        hist_venda_latest["Familia"] = familia_tipo_venda.apply(lambda ft: ft[0])
        hist_venda_latest["Tipo"] = familia_tipo_venda.apply(lambda ft: ft[1])
        hist_venda_latest["Mapa"] = hist_venda_latest["Mapa"].apply(limpa_mapa)

        # Só entram na soma por família os itens "soltos" (Garrafa ou Barril) — igual ao
        # Retorno digitado manualmente, que também só conta Garrafas+Unidades e nunca a
        # coluna Garrafeiras. Incluir garrafeira aqui infla a Saída e faz o mapa parecer
        # "faltando" mesmo quando bateu perfeito (garrafeira é conferida à parte, na seção
        # "Conferência Garrafa × Garrafeira" da Conciliação Mapas Sede).
        df_venda_ag = hist_venda_latest[(hist_venda_latest["Familia"] != "Outro") & (hist_venda_latest["Tipo"] != "Garrafeira")].copy()
        venda_agg = df_venda_ag.groupby(["Mapa", "Familia"])["Qtd. Vendida/Movimentada"].sum().reset_index()
        venda_agg.rename(columns={"Qtd. Vendida/Movimentada": "Qtd_Saida_Unidades"}, inplace=True)
        venda_agg = venda_agg[venda_agg["Mapa"].isin(MAPAS_COM_CONFERENCIA_PA)]

        # 2. RETORNO DO PA
        hist_vazio_pa["Mapa"] = hist_vazio_pa["Mapa"].apply(limpa_mapa)

        if "Garrafas" not in hist_vazio_pa.columns: hist_vazio_pa["Garrafas"] = 0
        if "Unidades" not in hist_vazio_pa.columns: hist_vazio_pa["Unidades"] = 0

        hist_vazio_pa["Qtd_Retorno_Unidades"] = pd.to_numeric(hist_vazio_pa["Garrafas"], errors='coerce').fillna(0) + \
                                                pd.to_numeric(hist_vazio_pa["Unidades"], errors='coerce').fillna(0)

        # PA "dono" de cada mapa: mesmo que uma família específica ainda não tenha sido
        # conferida (e portanto não apareça em vazio_agg pra essa família), o mapa como um
        # todo já pertence a um PA conhecido (Tianguá/Granja) — usado logo abaixo pra não
        # deixar a família faltante "sumir" sob um PA genérico.
        mapa_pa_lookup = hist_vazio_pa.groupby("Mapa")["PA"].first().to_dict()

        vazio_agg = hist_vazio_pa.groupby(["Mapa", "PA", "Familia"])["Qtd_Retorno_Unidades"].sum().reset_index()

        # 3. CRUZAMENTO (MERGE) E CÁLCULO FÍSICO
        df_concil = pd.merge(venda_agg, vazio_agg, on=["Mapa", "Familia"], how="outer").fillna(0)

        # Se uma família saiu na Previsto mas o conferente não digitou nenhum retorno pra ela
        # (situação real de "faltou"), a linha continua no PA correto do mapa (em vez de
        # cair num rótulo genérico "Aguardando Retorno" que ficava invisível ao filtrar
        # por um PA específico).
        df_concil["PA"] = df_concil.apply(
            lambda r: mapa_pa_lookup.get(r["Mapa"], "Aguardando Retorno") if r["PA"] == 0 else r["PA"],
            axis=1,
        )

        df_concil["Fator"] = df_concil["Familia"].apply(fator_conversao_caixas)

        df_concil["Caixas_Saida"] = df_concil["Qtd_Saida_Unidades"] // df_concil["Fator"]
        df_concil["Soltas_Saida"] = df_concil["Qtd_Saida_Unidades"] % df_concil["Fator"]

        df_concil["Caixas_Retorno"] = df_concil["Qtd_Retorno_Unidades"] // df_concil["Fator"]
        df_concil["Soltas_Retorno"] = df_concil["Qtd_Retorno_Unidades"] % df_concil["Fator"]

        df_concil["Diferença_Unidades"] = df_concil["Qtd_Retorno_Unidades"] - df_concil["Qtd_Saida_Unidades"]

        # 4. CRIAÇÃO DOS TEXTOS FORMATADOS
        def formata_cx_un(cx, un, fam):
            if "Barril" in fam:
                return f"{int(un)} un" if un > 0 else "0"
            else:
                if cx == 0 and un == 0: return "0"
                res = []
                if cx > 0: res.append(f"{int(cx)} cx")
                if un > 0: res.append(f"{int(un)} gf")
                return " + ".join(res)

        df_concil["Saída"] = df_concil.apply(lambda r: formata_cx_un(r["Caixas_Saida"], r["Soltas_Saida"], r["Familia"]), axis=1)
        df_concil["Retorno"] = df_concil.apply(lambda r: formata_cx_un(r["Caixas_Retorno"], r["Soltas_Retorno"], r["Familia"]), axis=1)

        def formata_dif(dif, fam):
            item = "un" if "Barril" in fam else "gf"
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
            else:  # dif > 0 (retornou mais do que saiu, incluindo quando não saiu nada)
                return "⚠️ Sobrou AG"

        df_concil["Status"] = df_concil.apply(status_conciliacao, axis=1)

        # 5. FILTROS E EXIBIÇÃO
        col_filtro1, col_filtro2, col_filtro3 = st.columns([1, 1, 2])

        lista_pas = ["Todos"] + sorted(df_concil["PA"].unique().tolist())
        pa_filter = col_filtro1.selectbox("Filtrar por PA:", lista_pas)
        status_filter = col_filtro2.selectbox("Filtrar por Status:", ["Todos", "❌ Faltou AG", "⚠️ Sobrou AG", "✅ Bateu"])
        mapa_search = col_filtro3.text_input("🔍 Pesquisar Mapa Específico (opcional):", "")

        df_display = df_concil.copy()

        if pa_filter != "Todos":
            df_display = df_display[df_display["PA"] == pa_filter]

        if status_filter != "Todos":
            df_display = df_display[df_display["Status"] == status_filter]

        if mapa_search.strip() != "":
            df_display = df_display[df_display["Mapa"].str.contains(limpa_mapa(mapa_search))]

        df_display = df_display[["Mapa", "PA", "Familia", "Saída", "Retorno", "Diferença", "Status"]]
        df_display = df_display.sort_values(by=["Mapa", "Familia"])

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
    st.caption("Cruza item a item o que estava previsto (saída) com o que foi realizado (retorno) — para todo mapa que NÃO foi digitado na aba 'Vazio por PA' (esses ficam na 'Conciliação Mapas PA').")

    hist_venda_item = ler_aba_historico("Venda")
    hist_retorno_654 = ler_aba_historico("Retorno654")
    hist_categorias_sede = ler_aba_historico(NOME_ABA_CATEGORIAS_EXTRA)

    if hist_venda_item.empty or hist_retorno_654.empty:
        st.info("⚠️ Aguardando dados. É necessário ter histórico de Venda (Previsto) e Retorno (Realizado) salvos para cruzar.")
    else:
        hist_venda_item = hist_venda_item.copy()
        hist_venda_item["Mapa"] = hist_venda_item["Mapa"].apply(limpa_mapa)
        hist_retorno_654 = hist_retorno_654.copy()
        hist_retorno_654["Mapa"] = hist_retorno_654["Mapa"].apply(limpa_mapa)
        if not hist_categorias_sede.empty:
            hist_categorias_sede = hist_categorias_sede.copy()
            hist_categorias_sede["Mapa"] = hist_categorias_sede["Mapa"].apply(limpa_mapa)

        # Total combinado: não importa em qual "espécie" o item saiu ou voltou (Vazio,
        # Comodato, Devolução, Troca, Consignação, Rec. Consignação) — só interessa se o
        # TOTAL previsto bateu com o TOTAL realizado. Uma diferença só aparece aqui quando
        # o saldo geral do item não fecha (ver aba "Outras Categorias" pra investigar em
        # qual espécie especificamente está o furo).
        df_concil_sede = montar_total_previsto_realizado(hist_venda_item, hist_retorno_654, hist_categorias_sede)
        df_concil_sede = df_concil_sede.rename(columns={"Total_Previsto": "Qtd_Saida_554", "Total_Realizado": "Qtd_Retorno_654"})
        df_concil_sede = df_concil_sede[~df_concil_sede["Mapa"].isin(MAPAS_COM_CONFERENCIA_PA)]

        # Valor mais recente (por data) de uma coluna qualquer, por mapa — usado só pra Data
        # (o 03.07.13 não tem coluna de Depósito, então essa informação não existe mais aqui).
        def valor_mais_recente_por_mapa(df, coluna):
            if df.empty or "Data" not in df.columns or coluna not in df.columns:
                return {}
            tmp = df.copy()
            tmp["_dt"] = pd.to_datetime(tmp["Data"], dayfirst=True, errors="coerce")
            tmp = tmp.dropna(subset=["_dt"])
            if tmp.empty:
                return {}
            idx = tmp.groupby("Mapa")["_dt"].idxmax()
            return tmp.loc[idx].set_index("Mapa")[coluna].to_dict()

        data_saida_por_mapa = valor_mais_recente_por_mapa(hist_venda_item, "Data")
        data_retorno_por_mapa = valor_mais_recente_por_mapa(hist_retorno_654, "Data")

        # Descrição: junta a descrição de quem tiver (Venda e/ou Retorno654), pra nenhum item ficar em branco
        col_desc_venda = next((c for c in ["Descrição", "Descricao"] if c in hist_venda_item.columns), None)
        col_desc_retorno = next((c for c in ["Descrição", "Descricao"] if c in hist_retorno_654.columns), None)

        partes_desc = []
        if col_desc_venda:
            partes_desc.append(hist_venda_item[["Material", col_desc_venda]].rename(columns={col_desc_venda: "Desc_AG"}))
        if col_desc_retorno:
            partes_desc.append(hist_retorno_654[["Material", col_desc_retorno]].rename(columns={col_desc_retorno: "Desc_AG"}))

        desc_por_material = None
        if partes_desc:
            desc_por_material = pd.concat(partes_desc, ignore_index=True)
            desc_por_material = desc_por_material[desc_por_material["Desc_AG"].astype(str).str.strip() != ""]
            desc_por_material = desc_por_material.drop_duplicates(subset=["Material"], keep="first")

        df_concil_sede["Data Saída"] = df_concil_sede["Mapa"].map(data_saida_por_mapa).fillna("-")
        df_concil_sede["Data Retorno"] = df_concil_sede["Mapa"].map(data_retorno_por_mapa).fillna("-")

        if desc_por_material is not None:
            df_concil_sede = df_concil_sede.merge(desc_por_material, on="Material", how="left")
            df_concil_sede["AG"] = [
                com_apelido(cod, str(desc)) for cod, desc in zip(df_concil_sede["Material"], df_concil_sede["Desc_AG"].fillna(""))
            ]
        else:
            df_concil_sede["AG"] = df_concil_sede["Material"]

        # Quantidades sempre em número inteiro (sem casas decimais sobrando)
        df_concil_sede["Qtd_Saida_554"] = df_concil_sede["Qtd_Saida_554"].round(0).astype(int)
        df_concil_sede["Qtd_Retorno_654"] = df_concil_sede["Qtd_Retorno_654"].round(0).astype(int)
        df_concil_sede["Diferença_Num"] = df_concil_sede["Qtd_Retorno_654"] - df_concil_sede["Qtd_Saida_554"]

        # Classifica o item pelo Código (De Material.xlsx) — Tipo (Garrafa/Garrafeira/Barril/
        # Outro) e Familia (300ml/600ml/Verde 600/1L/...), em vez de interpretar a descrição
        # abreviada do relatório (mais confiável, mesma fonte usada em toda a conciliação).
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

        datas_disponiveis_sede = sorted(
            {d for d in pd.concat([df_concil_sede["Data Saída"], df_concil_sede["Data Retorno"]]) if d != "-"},
            key=lambda d: pd.to_datetime(d, dayfirst=True, errors="coerce"),
            reverse=True,
        )

        mostrar_so_divergencias = st.checkbox("🔍 Mostrar só o que tem diferença (recomendado)", value=True, key="so_divergencias_sede")

        col_f0, col_f1, col_f2, col_f3 = st.columns([1, 1, 1, 2])
        data_filter_sede = col_f0.selectbox(
            "Filtrar por Data:",
            ["Todas"] + datas_disponiveis_sede,
            key="data_sede",
        )
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
        if data_filter_sede != "Todas":
            df_display_sede = df_display_sede[
                (df_display_sede["Data Saída"] == data_filter_sede) | (df_display_sede["Data Retorno"] == data_filter_sede)
            ]
        if status_filter_sede != "Todos":
            df_display_sede = df_display_sede[df_display_sede["Status"] == status_filter_sede]
        if mapa_search_sede.strip():
            df_display_sede = df_display_sede[df_display_sede["Mapa"].str.contains(limpa_mapa(mapa_search_sede))]
        if material_search_sede.strip():
            df_display_sede = df_display_sede[df_display_sede["AG"].str.contains(material_search_sede, case=False, na=False)]

        colunas_exibir_sede = ["Mapa", "Data Saída", "Data Retorno", "AG", "Saída (Total)", "Retorno (Total)", "Diferença", "Status"]
        df_display_sede = df_display_sede[colunas_exibir_sede].sort_values(by=["Mapa", "AG"])

        st.dataframe(
            df_display_sede.style.map(cor_linha_status, subset=["Status"]),
            use_container_width=True, hide_index=True,
        )

        st.caption("Nota: Saída/Retorno somam Vazio + Comodato + Devolução + Troca + Consignação + Rec. Consignação — não importa em qual espécie o item saiu ou voltou, só o total. Pra ver em qual espécie está a diferença, use a aba 'Outras Categorias' filtrando pelo mesmo mapa.")

        # =================================================================
        # CONFERÊNCIA CRUZADA: GARRAFA × GARRAFEIRA, POR MAPA
        # Pega o caso de faturar quantidade de garrafeira desproporcional à
        # quantidade de garrafa (ex: 1 garrafeira pra 48 garrafas) — mesmo que
        # cada item, individualmente, bata perfeito entre Saída e Retorno.
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
# Comparação direta Previsto x Realizado do próprio relatório 03.07.13, categoria
# por categoria — não passa pela conferência manual do Vazio por PA (essa só
# cobre a categoria Vazio). Aplica-se a TODO mapa com movimento nessas categorias,
# independente de já ter sido conferido ou não na aba Vazio por PA.
# =========================================================================
with aba_categorias_extra:
    st.header("📋 Divergências por Categoria")
    st.caption(
        "Previsto x Realizado, direto do relatório 03.07.13, pra cada categoria além de Vazio "
        "(essas não passam pela conferência manual do 'Vazio por PA' — comparação direta do relatório)."
    )

    hist_categorias = ler_aba_historico(NOME_ABA_CATEGORIAS_EXTRA)
    if hist_categorias.empty:
        st.info(
            "⚠️ Ainda não há dados dessas categorias no histórico. Rode o relatório 03.07.13 "
            "na aba 'Venda' do app Operacional pra alimentar esta aba."
        )
    else:
        hist_categorias = hist_categorias.copy()
        hist_categorias["Mapa"] = hist_categorias["Mapa"].apply(limpa_mapa)
        col_desc_cat = next((c for c in ["Descrição", "Descricao"] if c in hist_categorias.columns), None)

        linhas_cat = []
        for nome_cat, col_p, col_r in CATEGORIAS_AG_EXTRA:
            if col_p not in hist_categorias.columns or col_r not in hist_categorias.columns:
                continue
            previsto_latest = valor_mais_recente_por_grupo(hist_categorias, ["Mapa", "Material"], "Data", col_p)
            realizado_latest = valor_mais_recente_por_grupo(hist_categorias, ["Mapa", "Material"], "Data", col_r)
            cruzado = pd.merge(previsto_latest, realizado_latest, on=["Mapa", "Material"], how="outer").fillna(0)
            cruzado = cruzado[(cruzado[col_p] != 0) | (cruzado[col_r] != 0)]
            if cruzado.empty:
                continue
            cruzado = cruzado.rename(columns={col_p: "Previsto", col_r: "Realizado"})
            cruzado["Categoria"] = nome_cat
            linhas_cat.append(cruzado[["Mapa", "Material", "Categoria", "Previsto", "Realizado"]])

        if not linhas_cat:
            st.info("Nenhum movimento registrado ainda em Comodato, Devolução, Troca, Consignação ou Rec. Consignação.")
        else:
            df_cat = pd.concat(linhas_cat, ignore_index=True)
            df_cat["Previsto"] = df_cat["Previsto"].round(0).astype(int)
            df_cat["Realizado"] = df_cat["Realizado"].round(0).astype(int)
            df_cat["Diferença"] = df_cat["Realizado"] - df_cat["Previsto"]

            if col_desc_cat:
                desc_lookup_cat = hist_categorias.drop_duplicates(subset=["Material"]).set_index("Material")[col_desc_cat].to_dict()
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
