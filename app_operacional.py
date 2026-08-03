import streamlit as st
import pandas as pd
from datetime import date
from matplotlib.colors import LinearSegmentedColormap

from comum import (
    PASTA_PROJETO, ARQUIVO_DE_MATERIAL, ARQUIVO_PRESTACAO, ARQUIVO_COMODATO,
    ARQUIVO_MOVIMENTACAO, ARQUIVO_RET,
    carregar, salvar_aba_historico, ler_aba_historico, normalizar_codigo,
    limpa_mapa, numerizar, parse_qtde_entrada, exibir_seguro,
    padronizar_familia, fator_conversao_caixas, converter_cheio_em_ag,
    localizar_grade_mais_recente, extrair_data_do_nome_arquivo,
    acumular_historico, codigos_fora_do_depara,
    coletar_datas_disponiveis, MAPA_APELIDOS, com_apelido, calcular_totais_por_familia,
    classificar_tipo_generico, gdrive_ativo, listar_arquivos_pasta,
)

st.set_page_config(page_title="AG - Operacional", layout="wide")
st.title("Ativos de Giro (AG) — Operacional")
st.caption("Painel, Venda (554/654), Cheio e Vazio (cada um já com sua própria Variação por data). Vazio por PA e Conciliação de Mapas ficam no app separado 'Conciliação de Mapas'.")

# --- CARREGAMENTO DE ARQUIVOS (SIDEBAR) ---
with st.sidebar:
    if st.button("🔄 Atualizar dados", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.caption("Os arquivos brutos só são relidos quando você aperta esse botão.")

    if gdrive_ativo():
        with st.expander("🔍 Diagnóstico Google Drive"):
            try:
                arquivos_pasta = listar_arquivos_pasta()
                if arquivos_pasta:
                    st.success(f"{len(arquivos_pasta)} arquivo(s) visíveis na pasta configurada:")
                    for a in arquivos_pasta:
                        st.caption(f"• {a['name']}  ({a.get('modifiedTime', '?')})")
                else:
                    st.warning(
                        "A conexão com o Drive funcionou, mas a pasta apareceu vazia pra essa conta de serviço. "
                        "Confira se o pasta_id no secrets é exatamente o desta pasta, e se ela foi mesmo "
                        "compartilhada com o e-mail da conta de serviço (não um arquivo individual, a PASTA)."
                    )
            except Exception as e:
                st.error(f"Erro ao consultar o Drive: {e}")

    st.divider()

    st.header("Bases (carregadas automaticamente)")

    df_de_material = None
    try:
        df_de_material = carregar(ARQUIVO_DE_MATERIAL)
        if df_de_material is not None:
            df_de_material["Promax"] = normalizar_codigo(df_de_material["Promax"])
            st.success(f"De Material.xlsx ({len(df_de_material)} linhas)")
        else:
            st.error(f"Não encontrei '{ARQUIVO_DE_MATERIAL.name}'.")
    except Exception as e:
        st.error(f"Erro ao ler '{ARQUIVO_DE_MATERIAL.name}': {e}")

    df_depara_dedup = None
    if df_de_material is not None:
        duplicados = df_de_material[df_de_material.duplicated(subset=["Promax"], keep=False)]
        if not duplicados.empty:
            st.caption(f"{duplicados['Promax'].nunique()} código(s) duplicado(s) no de-para.")
        df_depara_dedup = df_de_material.drop_duplicates(subset=["Promax"], keep="first")

    alertas_fora_depara = []

    df_prestacao = None
    try:
        df_prestacao = carregar(ARQUIVO_PRESTACAO)
        if df_prestacao is not None:
            df_prestacao["Material"] = normalizar_codigo(df_prestacao["Material"])
            df_prestacao["Mapa"] = normalizar_codigo(df_prestacao["Mapa"])
            st.success(f"{ARQUIVO_PRESTACAO.name} — mapas da rota ({len(df_prestacao)} linhas)")
    except Exception as e:
        st.error(f"Erro ao ler '{ARQUIVO_PRESTACAO.name}': {e}")

    if df_prestacao is not None and df_de_material is not None:
        resumo_faltantes = codigos_fora_do_depara(df_prestacao, "Material", df_de_material, ARQUIVO_PRESTACAO.name)
        if resumo_faltantes is not None:
            alertas_fora_depara.append(resumo_faltantes)
        codigos_validos = set(df_de_material["Promax"].unique())
        df_prestacao = df_prestacao[df_prestacao["Material"].isin(codigos_validos)]

    df_comodato = None
    try:
        df_comodato = carregar(ARQUIVO_COMODATO)
        if df_comodato is not None:
            df_comodato.columns = df_comodato.columns.str.strip()
            if "Codigo Produto" in df_comodato.columns:
                df_comodato["Codigo Produto"] = normalizar_codigo(df_comodato["Codigo Produto"])
            st.success(f"{ARQUIVO_COMODATO.name} — comodato ({len(df_comodato)} linhas)")
    except Exception as e:
        pass

    df_movimentacao = None
    try:
        df_movimentacao = carregar(ARQUIVO_MOVIMENTACAO)
        if df_movimentacao is not None:
            df_movimentacao.columns = df_movimentacao.columns.str.strip()
            if "Item" in df_movimentacao.columns:
                df_movimentacao["Item"] = normalizar_codigo(df_movimentacao["Item"])
            if "Mapa" in df_movimentacao.columns:
                df_movimentacao["Mapa"] = normalizar_codigo(df_movimentacao["Mapa"])
            if "Código Operação" in df_movimentacao.columns:
                df_movimentacao["Código Operação"] = normalizar_codigo(df_movimentacao["Código Operação"])
            if "Qtde Entrada" in df_movimentacao.columns:
                df_movimentacao["Qtde Entrada"] = parse_qtde_entrada(df_movimentacao["Qtde Entrada"])
            st.success(f"{ARQUIVO_MOVIMENTACAO.name} — movimentação ({len(df_movimentacao)} linhas)")
    except Exception as e:
        pass

    df_estoque_cheio = None
    data_grade_cheio = None
    arquivo_cheio_encontrado = localizar_grade_mais_recente("02.03.04")
    try:
        if arquivo_cheio_encontrado is not None:
            df_estoque_cheio = carregar(arquivo_cheio_encontrado)
        if df_estoque_cheio is not None:
            df_estoque_cheio.columns = df_estoque_cheio.columns.str.strip()
            if "UN" in df_estoque_cheio.columns:
                df_estoque_cheio["UN"] = df_estoque_cheio["UN"].astype(str).str.strip()
            if "Cod" in df_estoque_cheio.columns:
                df_estoque_cheio["Cod"] = normalizar_codigo(df_estoque_cheio["Cod"])
            if "Inicial" in df_estoque_cheio.columns and "Ent." in df_estoque_cheio.columns:
                df_estoque_cheio["Inicial"] = numerizar(df_estoque_cheio["Inicial"])
                df_estoque_cheio["Ent."] = numerizar(df_estoque_cheio["Ent."])
                df_estoque_cheio["Qtd_Cheio"] = df_estoque_cheio["Inicial"] + df_estoque_cheio["Ent."]
            data_grade_cheio = extrair_data_do_nome_arquivo(arquivo_cheio_encontrado, "02.03.04")
            if data_grade_cheio is None:
                data_grade_cheio = date.today().strftime("%d/%m/%Y")
            st.success(f"{arquivo_cheio_encontrado.name} — estoque cheio ({len(df_estoque_cheio)} linhas)")
    except Exception as e:
        pass

    df_ret = None
    try:
        df_ret = carregar(ARQUIVO_RET)
        if df_ret is not None:
            df_ret.columns = df_ret.columns.str.strip()
            if "CodProd" in df_ret.columns:
                df_ret["CodProd"] = normalizar_codigo(df_ret["CodProd"])
            if "QTD_SKU" in df_ret.columns:
                df_ret["QTD_SKU"] = numerizar(df_ret["QTD_SKU"])
            st.success(f"{ARQUIVO_RET.name} — cadastro retornáveis ({len(df_ret)} linhas)")
    except Exception as e:
        pass

if alertas_fora_depara:
    df_alerta = pd.concat(alertas_fora_depara, ignore_index=True)
    st.warning("Códigos apareceram nas bases mas não estão no De Material.")


ORDEM_FAMILIA = {"300ml": 0, "Verde 600": 1, "600ml": 1, "1L": 2, "Barril 30L": 3, "Barril 50L": 3, "Barril": 3}
ORDEM_TIPO_600 = {"Garrafa_Verde 600": 0, "Garrafa_600ml": 1, "Garrafeira_Verde 600": 2, "Garrafeira_600ml": 2}


def chave_ordenacao(rotulo_completo: str) -> tuple:
    familia = padronizar_familia(rotulo_completo) or "ZZZ"
    tipo = classificar_tipo_generico(rotulo_completo)
    fam_prioridade = ORDEM_FAMILIA.get(familia, 9)
    chave_600 = f"{tipo}_{familia}"
    sub_prioridade = ORDEM_TIPO_600.get(chave_600, 0 if tipo in ("Garrafa", "Barril") else 1 if tipo == "Garrafeira" else 9)
    return (fam_prioridade, sub_prioridade, rotulo_completo)


def mostrar_mapa_calor(pivot: pd.DataFrame, rotulo_metrica: str):
    try: pivot = pivot[sorted(pivot.columns, key=lambda d: pd.to_datetime(d, dayfirst=True))].fillna(0)
    except Exception: pivot = pivot.fillna(0)
    cmap = LinearSegmentedColormap.from_list("azul", ["#F3F8FD", "#CFE4F7", "#9AC7EE", "#5FA3DE", "#2E75B8"])
    st.markdown(f"**{rotulo_metrica}**")
    st.dataframe(pivot.style.background_gradient(cmap=cmap, axis=None).format(precision=0, thousands="."), width='stretch')


def renderizar_mapa_calor(nome_aba_excel, coluna_valor, rotulo_metrica):
    historico = ler_aba_historico(nome_aba_excel)
    if historico.empty or coluna_valor not in historico.columns:
        st.info(f"Ainda não há histórico pra '{rotulo_metrica}'.")
        return
    colunas_desc = [c for c in ["Descrição", "Descricao"] if c in historico.columns]
    rotulo = historico["Material"].astype(str)
    if colunas_desc: rotulo = rotulo + " - " + historico[colunas_desc[0]].astype(str)
    historico["AG"] = [com_apelido(cod, r) for cod, r in zip(historico["Material"].astype(str), rotulo)]
    pivot = historico.pivot_table(index="AG", columns="Data", values=coluna_valor, aggfunc="sum")
    pivot = pivot.reindex(sorted(pivot.index, key=chave_ordenacao))
    mostrar_mapa_calor(pivot, rotulo_metrica)
    if st.button(f"Limpar histórico de: {rotulo_metrica}"):
        salvar_aba_historico(nome_aba_excel, pd.DataFrame())
        st.rerun()


def renderizar_mapa_calor_unificado(nome_aba_excel, rotulo_fonte):
    historico = ler_aba_historico(nome_aba_excel)
    if historico.empty:
        st.info(f"Ainda não há histórico pra '{rotulo_fonte}'.")
        return

    col_gf = "Garrafeiras_ou_Barris" if "Garrafeiras_ou_Barris" in historico.columns else "Garrafeiras"
    linhas = {}

    def pegar(fam, col):
        sub = historico[historico["Material"].astype(str) == fam]
        return sub.groupby("Data")[col].sum() if not sub.empty and col in sub.columns else None

    for rot, fam in [("Garrafa Litrinho", "300ml"), ("Garrafa 600 Verde", "Verde 600"), ("Garrafa 600 Normal", "600ml"), ("Garrafa 1L", "1L")]:
        if (s := pegar(fam, "Garrafas")) is not None: linhas[rot] = s

    for rot, fam in [("Garrafeira Litrinho", "300ml"), ("Garrafeira 1L", "1L")]:
        if (s := pegar(fam, col_gf if col_gf == "Garrafeiras" else "Garrafeiras")) is not None: linhas[rot] = s

    s_600, s_vd = pegar("600ml", "Garrafeiras"), pegar("Verde 600", "Garrafeiras")
    if s_600 is not None or s_vd is not None:
        base = pd.Series(dtype=float).add(s_600 if s_600 is not None else 0, fill_value=0)
        linhas["Garrafeira 600ml"] = base.add(s_vd if s_vd is not None else 0, fill_value=0)

    for rot, fam in [("Barril", "Barril"), ("Barril 30L", "Barril 30L"), ("Barril 50L", "Barril 50L")]:
        col_busca = "Unidades" if fam != "Barril" else col_gf
        if (s := pegar(fam, col_busca)) is not None: linhas[rot] = s

    if not linhas: return
    pivot = pd.DataFrame(linhas).T.reindex(sorted(pd.DataFrame(linhas).T.index, key=chave_ordenacao))
    mostrar_mapa_calor(pivot, rotulo_fonte)
    if st.button(f"Limpar histórico de: {rotulo_fonte}"):
        salvar_aba_historico(nome_aba_excel, pd.DataFrame())
        st.rerun()


aba_painel, aba_movimentacao, aba_cheio, aba_vazio, aba_dados = st.tabs(
    ["Painel", "Venda", "Cheio", "Vazio", "Dados"]
)


with aba_movimentacao:
    st.caption("Movimentação (vendido/retornado) por AG — Filtrado exclusivamente pelos códigos presentes no 'De Material'.")

    st.markdown("#### 📈 Variação (histórico por dia)")
    renderizar_mapa_calor("Venda", "Qtd. Vendida/Movimentada", "Venda (Operação 554)")
    st.divider()
    st.markdown("#### 📋 Resumo (arquivo carregado agora)")

    if df_movimentacao is not None:
        colunas_necessarias = {"Item", "Código Operação", "Qtde Entrada"}
        if colunas_necessarias.issubset(df_movimentacao.columns):
            codigo_limpo = df_movimentacao["Código Operação"].astype(str).str.extract(r"(\d+)")[0]
            mov_venda = df_movimentacao[codigo_limpo == "554"]
            mov_retorno_654 = df_movimentacao[codigo_limpo == "654"]

            codigos_validos_ag = set(df_depara_dedup["Promax"].unique()) if df_depara_dedup is not None else None
            if codigos_validos_ag is not None:
                if not mov_venda.empty:
                    mov_venda = mov_venda[mov_venda["Item"].isin(codigos_validos_ag)]
                if not mov_retorno_654.empty:
                    mov_retorno_654 = mov_retorno_654[mov_retorno_654["Item"].isin(codigos_validos_ag)]

            colunas_data_mov = ["Data"] if "Data" in df_movimentacao.columns else []
            colunas_mapa_mov = ["Mapa"] if "Mapa" in df_movimentacao.columns else []
            colunas_desc_mov = [c for c in ["Descrição", "Descricao"] if c in df_movimentacao.columns]
            desc_material = None
            if colunas_desc_mov:
                desc_material = df_movimentacao.rename(columns={"Item": "Material"}).drop_duplicates(subset=["Material"])[["Material"] + colunas_desc_mov]

            # --- Operação 554 (saída) ---
            if mov_venda.empty:
                st.warning("Após os filtros, nenhuma venda válida de AG foi encontrada para a operação 554.")
            else:
                resumo_venda = mov_venda.groupby(["Item"] + colunas_data_mov + colunas_mapa_mov)["Qtde Entrada"].sum().reset_index().rename(columns={"Item": "Material", "Qtde Entrada": "Qtd. Vendida/Movimentada"})

                if desc_material is not None:
                    resumo_venda = resumo_venda.merge(desc_material, on="Material", how="left")

                resumo_venda["Qtd. Vendida/Movimentada"] = resumo_venda["Qtd. Vendida/Movimentada"].round(0).astype(int)

                if "Data" in resumo_venda.columns:
                    colunas_historico = ["Material", "Data"] + colunas_mapa_mov + colunas_desc_mov
                    chave_historico = ["Material", "Data"] + colunas_mapa_mov
                    acumular_historico(resumo_venda[colunas_historico + ["Qtd. Vendida/Movimentada"]], "Venda", chave_historico)

                st.markdown("**Detalhe por produto — Saída (554)**")
                st.dataframe(resumo_venda.sort_values("Qtd. Vendida/Movimentada", ascending=False), width='stretch')

            # --- Operação 654 (retorno) ---
            if mov_retorno_654.empty:
                st.caption("Nenhum retorno (Operação 654) encontrado nesta base.")
            else:
                resumo_retorno = mov_retorno_654.groupby(["Item"] + colunas_data_mov + colunas_mapa_mov)["Qtde Entrada"].sum().reset_index().rename(columns={"Item": "Material", "Qtde Entrada": "Qtd_Retorno_654"})

                if desc_material is not None:
                    resumo_retorno = resumo_retorno.merge(desc_material, on="Material", how="left")

                resumo_retorno["Qtd_Retorno_654"] = resumo_retorno["Qtd_Retorno_654"].round(0).astype(int)

                if "Data" in resumo_retorno.columns:
                    colunas_historico_ret = ["Material", "Data"] + colunas_mapa_mov + colunas_desc_mov
                    chave_historico_ret = ["Material", "Data"] + colunas_mapa_mov
                    acumular_historico(resumo_retorno[colunas_historico_ret + ["Qtd_Retorno_654"]], "Retorno654", chave_historico_ret)

                st.markdown("**Detalhe por produto — Retorno (654)**")
                st.dataframe(resumo_retorno.sort_values("Qtd_Retorno_654", ascending=False), width='stretch')


with aba_cheio:
    st.caption("Estoque Cheio — Filtrado exclusivamente pela coluna de Retornabilidade do RET.csv.")

    st.markdown("#### 📈 Variação (histórico por dia)")
    renderizar_mapa_calor_unificado("Cheio", "Cheio")
    st.divider()
    st.markdown("#### 📋 Resumo (arquivo carregado agora)")

    if df_estoque_cheio is not None and df_ret is not None:
        if {"Cod", "UN", "Qtd_Cheio", "Descricao"}.issubset(df_estoque_cheio.columns):

            colunas_ret = ["CodProd", "Embalagem", "QTD_SKU"]
            col_ret = next((c for c in df_ret.columns if "RETORN" in c.upper() or "RET" == c.upper() or "TIPO" in c.upper()), None)
            if col_ret and col_ret not in colunas_ret:
                colunas_ret.append(col_ret)

            df_cheio = df_estoque_cheio.merge(
                df_ret[colunas_ret].drop_duplicates(subset=["CodProd"]),
                left_on="Cod", right_on="CodProd", how="inner"
            )

            if col_ret:
                condicao_ret = df_cheio[col_ret].astype(str).str.upper().str.contains("RET|SIM|^S", na=False, regex=True)
                df_cheio = df_cheio[condicao_ret]
            else:
                st.warning("A coluna 'Retornabilidade' não foi encontrada no RET.csv. O filtro de segurança não pôde ser aplicado.")

            if not df_cheio.empty:
                df_cheio["Familia"] = df_cheio["Descricao"].apply(padronizar_familia)
                df_cheio = df_cheio[df_cheio["Familia"] != "Outro"]

                conversoes = df_cheio.apply(converter_cheio_em_ag, axis=1)
                df_cheio = pd.concat([df_cheio, conversoes], axis=1)

                totais = df_cheio.groupby("Familia")[["Garrafas", "Garrafeiras", "Barris"]].sum().round(0).astype(int)
                st.markdown("### Total de AG por família (neste arquivo)")
                st.dataframe(totais, width='stretch')

                historico_cheio_dia = totais.reset_index().rename(columns={"Familia": "Material"})
                historico_cheio_dia["Descricao"] = historico_cheio_dia["Material"]
                historico_cheio_dia["Garrafeiras_ou_Barris"] = historico_cheio_dia["Garrafeiras"] + historico_cheio_dia["Barris"]
                historico_cheio_dia["Data"] = data_grade_cheio or date.today().strftime("%d/%m/%Y")

                acumular_historico(historico_cheio_dia, "Cheio", ["Material", "Data"])

                st.markdown("**Detalhe por produto**")
                colunas_detalhe_cheio = ["Cod", "Descricao", "UN", "Inicial", "Ent.", "Qtd_Cheio", "Embalagem", "Familia", "Garrafas", "Garrafeiras", "Barris"]
                colunas_detalhe_cheio = [c for c in colunas_detalhe_cheio if c in df_cheio.columns]
                df_cheio_exibir = df_cheio[colunas_detalhe_cheio].copy()
                st.dataframe(df_cheio_exibir, width='stretch')
    else:
        st.info("Carregue o estoque cheio (02.03.04.csv) e o cadastro de retornáveis (RET.csv) na barra lateral.")


with aba_vazio:
    st.caption("Digitação manual do relatório diário 'FAROL AG'.")
    REGRAS_VAZIO = {
        "300ml": {"garrafas_por_cx": 23, "garrafeiras_por_cx": 1},
        "600ml": {"garrafas_por_cx": 24, "garrafeiras_por_cx": 1},
        "Verde 600": {"garrafas_por_cx": 24, "garrafeiras_por_cx": 1},
        "1L": {"garrafas_por_cx": 12, "garrafeiras_por_cx": 1},
    }

    with st.form("form_farol_ag"):
        data_farol = st.date_input("Data da grade (FAROL AG)", value=date.today())
        st.markdown("**Caixas por família**")
        valores_familia = {}
        for familia in REGRAS_VAZIO:
            valores_familia[familia] = st.number_input(familia, min_value=0, step=1, key=f"cx_{familia}")

        st.markdown("**Outros AG (sem conversão — já em unidade final)**")
        c1, c2, c3, c4, c5 = st.columns(5)
        chapatex_und = c1.number_input("Chapatex (Und)", min_value=0, step=1)
        pbr1 = c2.number_input("Pallet PBR1", min_value=0, step=1)
        pbr2 = c3.number_input("Pallet PBR2", min_value=0, step=1)
        barril_30l = c4.number_input("Barril 30L", min_value=0, step=1)
        barril_50l = c5.number_input("Barril 50L", min_value=0, step=1)

        enviado = st.form_submit_button("Salvar grade do dia")

    if enviado:
        data_str = data_farol.strftime("%d/%m/%Y")
        garrafeiras_600_unificadas = valores_familia.get("600ml", 0) + valores_familia.get("Verde 600", 0)
        linhas = []
        for familia, qtd_cx in valores_familia.items():
            regra = REGRAS_VAZIO[familia]
            garrafeiras = garrafeiras_600_unificadas if familia == "600ml" else (0 if familia == "Verde 600" else qtd_cx * regra["garrafeiras_por_cx"])
            linhas.append({
                "Material": familia, "Descrição": familia, "Data": data_str,
                "Caixas": qtd_cx, "Garrafas": qtd_cx * regra["garrafas_por_cx"], "Garrafeiras": garrafeiras,
            })
        for material, qtd in [
            ("Chapatex", chapatex_und), ("Pallet PBR1", pbr1), ("Pallet PBR2", pbr2),
            ("Barril 30L", barril_30l), ("Barril 50L", barril_50l),
        ]:
            linhas.append({
                "Material": material, "Descrição": material, "Data": data_str,
                "Caixas": 0, "Garrafas": 0, "Garrafeiras": 0, "Unidades": qtd,
            })
        acumular_historico(pd.DataFrame(linhas), "Vazio", ["Material", "Data"])
        st.success(f"Grade de {data_str} salva.")

    historico_vazio = ler_aba_historico("Vazio")
    if not historico_vazio.empty:
        st.divider()
        st.markdown("#### 📈 Variação (histórico por dia)")
        renderizar_mapa_calor_unificado("Vazio", "Vazio (FAROL AG)")
        st.divider()
        st.markdown("#### 📋 Resumo (histórico salvo)")
        st.dataframe(historico_vazio.sort_values("Data", ascending=False), width='stretch', hide_index=True)


with aba_dados:
    if df_de_material is not None:
        st.subheader("De Material (de-para)"); st.dataframe(exibir_seguro(df_de_material), width='stretch')
    if df_prestacao is not None:
        st.subheader("Prestação de Contas"); st.dataframe(exibir_seguro(df_prestacao), width='stretch')
    if df_comodato is not None:
        st.subheader("Estoque / Comodato"); st.dataframe(exibir_seguro(df_comodato), width='stretch')
    if df_movimentacao is not None:
        st.subheader("Movimentação"); st.dataframe(exibir_seguro(df_movimentacao), width='stretch')
    if df_estoque_cheio is not None:
        st.subheader("Estoque Cheio"); st.dataframe(exibir_seguro(df_estoque_cheio), width='stretch')
    if df_ret is not None:
        st.subheader("Cadastro Retornáveis"); st.dataframe(exibir_seguro(df_ret), width='stretch')


# =========================================================================
# BLOCO DO PAINEL EXECUTIVO
# =========================================================================

with aba_painel:
    col_header, col_toggle, col_destaque = st.columns([1, 1, 1])
    with col_header:
        st.header("Visão Geral do AG")
    with col_toggle:
        visao = st.radio("Selecione a Métrica:", ["Caixas / Unidades", "Hectolitros (HL)"], horizontal=True)
    with col_destaque:
        destaque_cards = st.selectbox("Destacar nos números grandes:", ["Cheio", "Vazio", "Venda", "Total"])

    datas_disponiveis = coletar_datas_disponiveis("Cheio", "Venda", "Vazio")
    if not datas_disponiveis:
        st.info("Ainda não há histórico suficiente (Cheio, Venda ou Vazio) pra montar o painel.")
        st.stop()

    datas_dt = pd.to_datetime(datas_disponiveis, dayfirst=True)
    data_escolhida = st.date_input("Data", value=datas_dt.max().date(), min_value=datas_dt.min().date(), max_value=date.today()).strftime("%d/%m/%Y")

    if data_escolhida not in datas_disponiveis:
        st.warning(f"Não há histórico registrado pra {data_escolhida}.")

    familias_exibicao = ["300ml", "600ml", "Verde 600", "1L", "Barril 30L", "Barril 50L"]
    dict_cheio, dict_venda, dict_vazio = calcular_totais_por_familia(data_escolhida, familias_exibicao)

    linhas_painel = []
    for fam in familias_exibicao:
        fator_hl = {"300ml": 0.003, "600ml": 0.006, "Verde 600": 0.006, "1L": 0.01, "Barril 30L": 0.3, "Barril 50L": 0.5}[fam]
        fator_cx = fator_conversao_caixas(fam)

        c_val = dict_cheio[fam] * fator_hl if visao == "Hectolitros (HL)" else dict_cheio[fam] / fator_cx
        v_val = dict_venda[fam] * fator_hl if visao == "Hectolitros (HL)" else dict_venda[fam] / fator_cx
        vz_val = dict_vazio[fam] * fator_hl if visao == "Hectolitros (HL)" else dict_vazio[fam] / fator_cx

        linhas_painel.append({"Família": fam, "Cheio": c_val, "Venda": v_val, "Vazio": vz_val})

    df_painel = pd.DataFrame(linhas_painel)
    df_painel["Total"] = df_painel[["Cheio", "Venda", "Vazio"]].sum(axis=1)

    def formatar_numero(valor, vis):
        if pd.isna(valor) or valor == 0: return "0"
        return f"{int(valor):_}".replace("_", ".") if vis == "Caixas / Unidades" else f"{valor:_.2f}".replace(".", ",").replace("_", ".")

    unidade_str = "Cxs/Und" if visao == "Caixas / Unidades" else "HL"

    st.markdown(f"### 🎯 FAROL DO AG — {data_escolhida}")
    st.write("")

    colunas_metricas = st.columns(3)
    for i, row in df_painel.iterrows():
        colunas_metricas[i % 3].metric(
            label=f"{row['Família']} | {destaque_cards} ({unidade_str})",
            value=formatar_numero(row[destaque_cards], visao)
        )

    st.divider()
    st.dataframe(
        df_painel.style.map(lambda _: "background-color: #D6E4F9; color: #1E4D8C; font-weight: 600", subset=["Total"]).format(
            {c: lambda x: formatar_numero(x, visao) for c in ["Cheio", "Venda", "Vazio", "Total"]}
        ), use_container_width=True, hide_index=True
    )

    st.divider()
    st.markdown("### 📈 Total por dia (Cheio + Venda + Vazio)")

    linhas_evolucao = []
    for d_loop in datas_disponiveis:
        c_loop, v_loop, vz_loop = calcular_totais_por_familia(d_loop, familias_exibicao)
        for fam in familias_exibicao:
            fator_hl = {"300ml": 0.003, "600ml": 0.006, "Verde 600": 0.006, "1L": 0.01, "Barril 30L": 0.3, "Barril 50L": 0.5}[fam]
            fator_cx = fator_conversao_caixas(fam)
            total_dia = (c_loop[fam] + v_loop[fam] + vz_loop[fam]) * (fator_hl if visao == "Hectolitros (HL)" else 1/fator_cx)
            linhas_evolucao.append({"Família": fam, "Data": d_loop, "Total": total_dia})

    pivot_total = pd.DataFrame(linhas_evolucao).pivot_table(index="Família", columns="Data", values="Total", aggfunc="sum")
    try: pivot_total = pivot_total[sorted(pivot_total.columns, key=lambda d: pd.to_datetime(d, dayfirst=True))].reindex(familias_exibicao)
    except Exception: pivot_total = pivot_total.reindex(familias_exibicao)

    coluna_hoje = date.today().strftime("%d/%m/%Y")
    if coluna_hoje not in pivot_total.columns: coluna_hoje = datas_dt.max().date().strftime("%d/%m/%Y")

    st.dataframe(
        pivot_total.style.background_gradient(cmap=LinearSegmentedColormap.from_list("azul", ["#F3F8FD", "#CFE4F7", "#9AC7EE", "#5FA3DE", "#2E75B8"]), axis=None)
        .apply(lambda s: ["border: 2px solid #1E4D8C; font-weight: 600" if s.name == coluna_hoje else ""] * len(s), axis=0)
        .format(lambda x: formatar_numero(x, visao)),
        width='stretch',
    )
