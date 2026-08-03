"""
Funções e constantes compartilhadas entre os apps de AG:
- app_operacional.py  (Painel, Venda, Cheio, Vazio, Vazio por PA, Variação, Dados)
- app_conciliacao.py  (Conciliação Mapas PA, Conciliação Mapas Sede)

Ambos os apps devem rodar na MESMA pasta que este arquivo, junto com
historico_ag.xlsx, De Material.xlsx, 02.05.01.csv, RET.csv, etc.
"""

import streamlit as st
import pandas as pd
import re
from pathlib import Path
from datetime import date

PASTA_PROJETO = Path(__file__).parent
ARQUIVO_DE_MATERIAL = PASTA_PROJETO / "De Material.xlsx"
ARQUIVO_PRESTACAO = PASTA_PROJETO / "03.07.13.csv"   # mapas da rota
ARQUIVO_COMODATO = PASTA_PROJETO / "02.02.20.csv"    # comodato (emprestado)
ARQUIVO_MOVIMENTACAO = PASTA_PROJETO / "02.05.01.csv"  # movimentações 554/654
ARQUIVO_RET = PASTA_PROJETO / "RET.csv"  # cadastro de produtos retornáveis
ARQUIVO_HISTORICO_EXCEL = PASTA_PROJETO / "historico_ag.xlsx"  # planilha única, uma aba por origem


# =========================================================================
# LEITURA/ESCRITA DE ARQUIVOS
# =========================================================================

def deduplicar_nomes_coluna(nomes: list[str]) -> list[str]:
    vistos: dict[str, int] = {}
    resultado = []
    for i, nome in enumerate(nomes):
        base = nome.strip() if nome.strip() else f"col_sem_nome_{i}"
        if base in vistos:
            vistos[base] += 1
            resultado.append(f"{base}____{vistos[base]}")
        else:
            vistos[base] = 0
            resultado.append(base)
    return resultado


def ler_csv_robusto(caminho: Path) -> pd.DataFrame:
    encodings = ["utf-8-sig", "latin1", "cp1252"]
    separadores = [";", ",", "\t"]
    ultimo_erro = None
    for enc in encodings:
        for sep in separadores:
            try:
                with open(caminho, encoding=enc) as f:
                    primeira_linha = f.readline().rstrip("\n\r")
                colunas = deduplicar_nomes_coluna(primeira_linha.split(sep))
                return pd.read_csv(
                    caminho, sep=sep, encoding=enc, thousands=".", decimal=",",
                    header=0, names=colunas,
                )
            except Exception as e:
                ultimo_erro = e
    raise ultimo_erro


@st.cache_data(show_spinner=False)
def carregar(caminho: Path) -> pd.DataFrame | None:
    if not caminho.exists():
        return None
    if caminho.suffix.lower() == ".csv":
        return ler_csv_robusto(caminho)
    return pd.read_excel(caminho)


def salvar_aba_historico(nome_aba: str, df: pd.DataFrame) -> None:
    if ARQUIVO_HISTORICO_EXCEL.exists():
        with pd.ExcelWriter(ARQUIVO_HISTORICO_EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=nome_aba, index=False)
    else:
        with pd.ExcelWriter(ARQUIVO_HISTORICO_EXCEL, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name=nome_aba, index=False)


def ler_aba_historico(nome_aba: str) -> pd.DataFrame:
    if not ARQUIVO_HISTORICO_EXCEL.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(ARQUIVO_HISTORICO_EXCEL, sheet_name=nome_aba)
    except ValueError:
        return pd.DataFrame()


def normalizar_codigo(serie: pd.Series) -> pd.Series:
    def conv(v):
        if pd.isna(v): return None
        if isinstance(v, float) and v.is_integer(): return str(int(v))
        return str(v).strip()
    return serie.apply(conv)


def limpa_mapa(m):
    """Remove zeros à esquerda e espaços para garantir que os mapas casem perfeitamente no cruzamento."""
    m = str(m).strip()
    try:
        return str(int(m))
    except Exception:
        return m


def numerizar(serie: pd.Series) -> pd.Series:
    return pd.to_numeric(serie, errors="coerce").fillna(0)


def parse_qtde_entrada(serie: pd.Series) -> pd.Series:
    s = serie.astype(str).str.strip()
    s = s.str.replace(".", "", regex=False)
    s = s.str.replace("/", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0)


def exibir_seguro(df: pd.DataFrame) -> pd.DataFrame:
    df_seguro = df.copy()
    for col in df_seguro.select_dtypes(include=["object"]).columns:
        df_seguro[col] = df_seguro[col].astype(str)
    return df_seguro


# =========================================================================
# MOTOR CENTRAL DE CLASSIFICAÇÃO E CONVERSÃO
# =========================================================================

def padronizar_familia(desc: str) -> str:
    """Classifica os itens já filtrados na base mestre em suas respectivas famílias de volume."""
    d = str(desc).upper()

    if "300C23" in d or "300C24" in d or "GFVD300" in d or "LITRINHO" in d or "300ML" in d or "330ML" in d:
        return "300ml"

    marcas_verde = ["SPTN", "SPATEN", "STELLA", "S ARTOIS", "STARTPG", "BECKS", "HEINEKEN"]
    if ("600" in d or "635" in d) and any(m in d for m in marcas_verde):
        return "Verde 600"

    if "600" in d or "635" in d:
        return "600ml"

    if "1000" in d or "1L" in d or "1 L" in d or "LITRAO" in d or "LITRÃO" in d or "GFVD1000" in d:
        return "1L"

    if "BARRIL" in d or "KEG" in d or "CHOPP" in d or "CHP" in d:
        if "30" in d: return "Barril 30L"
        if "50" in d: return "Barril 50L"
        return "Barril"

    return "Outro"


def fator_conversao_caixas(fam: str) -> float:
    """Fator universal para converter garrafas soltas vendidas em caixas físicas."""
    if fam == "300ml": return 23.0
    if fam in ["600ml", "Verde 600"]: return 24.0
    if fam == "1L": return 12.0
    if "Barril" in fam: return 1.0
    return 1.0


def converter_cheio_em_ag(row: pd.Series) -> pd.Series:
    familia = row["Familia"]
    disp = float(row.get("Qtd_Cheio", 0))
    un = str(row.get("UN", "")).strip().upper()
    qtd_sku = float(row.get("QTD_SKU", 0))

    garrafas = garrafeiras = barris = 0.0

    if familia == "300ml":
        if qtd_sku <= 0: qtd_sku = 23
        if un == "CX":
            garrafas = disp * qtd_sku
            garrafeiras = disp
        else:
            garrafas = disp
            garrafeiras = disp / qtd_sku

    elif familia in ("600ml", "Verde 600"):
        if un == "DZ":
            garrafas = disp * 12
            garrafeiras = disp / 2
        elif un == "CX":
            if qtd_sku <= 0: qtd_sku = 24
            garrafas = disp * qtd_sku
            garrafeiras = disp
        else:
            garrafas = disp
            garrafeiras = disp / 24

    elif familia == "1L":
        if un == "DZ":
            garrafas = disp * 12
            garrafeiras = disp
        elif un == "CX":
            if qtd_sku <= 0: qtd_sku = 12
            garrafas = disp * qtd_sku
            garrafeiras = disp
        else:
            garrafas = disp
            garrafeiras = disp / 12

    elif familia.startswith("Barril"):
        if un == "L":
            if "50" in familia:
                barris = disp / 50
            elif "30" in familia:
                barris = disp / 30
            else:
                barris = disp
        else:
            barris = disp

    return pd.Series({"Garrafas": garrafas, "Garrafeiras": garrafeiras, "Barris": barris})


def encontrar_arquivo_por_prefixo(pasta: Path, prefixo: str) -> Path | None:
    candidatos = sorted(pasta.glob(f"{prefixo}*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0] if candidatos else None


def extrair_data_do_nome_arquivo(caminho: Path, prefixo: str) -> str | None:
    resto = caminho.stem[len(prefixo):].strip("._ ")
    m = re.search(r"(\d{2})[_.\-](\d{2})[_.\-](\d{2,4})", resto)
    if not m:
        return None
    dia, mes, ano = m.groups()
    if len(ano) == 2:
        ano = "20" + ano
    return f"{dia}/{mes}/{ano}"


def acumular_historico(df_dia: pd.DataFrame, nome_aba: str, colunas_chave: list[str]) -> pd.DataFrame:
    historico = ler_aba_historico(nome_aba)
    if not historico.empty:
        for c in colunas_chave:
            if c in historico.columns:
                historico[c] = historico[c].astype(str)
        combinado = pd.concat([historico, df_dia], ignore_index=True)
    else:
        combinado = df_dia.copy()
    combinado = combinado.drop_duplicates(subset=colunas_chave, keep="last")
    salvar_aba_historico(nome_aba, combinado)
    return combinado


def codigos_fora_do_depara(df: pd.DataFrame, coluna_codigo: str, df_de_material: pd.DataFrame, nome_base: str) -> pd.DataFrame | None:
    if df is None or df_de_material is None or coluna_codigo not in df.columns: return None
    codigos_validos = set(df_de_material["Promax"].unique())
    faltantes = df[~df[coluna_codigo].isin(codigos_validos)]
    if faltantes.empty: return None
    colunas_desc = [c for c in ["Descricao", "Desc 2", "Desc"] if c in df.columns]
    resumo = faltantes.groupby([coluna_codigo] + colunas_desc, dropna=False).size().reset_index(name="Linhas afetadas")
    resumo.insert(0, "Base", nome_base)
    return resumo


def cor_linha_status(val):
    """Colore a coluna Status nas abas de conciliação (PA e Sede) — mesma paleta pras duas."""
    if "✅" in str(val): return "color: #173404; font-weight: bold; background-color: #EAF3DE"
    if "❌" in str(val): return "color: #501313; font-weight: bold; background-color: #FCEBEB"
    if "⚠️" in str(val): return "color: #5A4000; font-weight: bold; background-color: #FFF4D4"
    if "🔎" in str(val): return "color: #004085; font-weight: bold; background-color: #CCE5FF"
    return ""


def classificar_tipo_generico(desc: str) -> str:
    """Classifica o item em Garrafa / Garrafeira / Barril / Outro, pra decidir como exibir a quantidade física."""
    d = str(desc).upper()
    if "GARRAFEIRA" in d: return "Garrafeira"
    if "BARRIL" in d or "KEG" in d: return "Barril"
    if "GFA" in d or "GARRAFA" in d or "VIDRO" in d: return "Garrafa"
    return "Outro"


def formata_qtd_fisica(qtd, tipo: str, familia: str) -> str:
    """Garrafa (300ml/600ml/Verde 600/1L) vira 'X cx + Y gf'; qualquer outro tipo (Garrafeira, Barril, Palete, etc.) vira 'X un'."""
    qtd = int(qtd)
    if qtd == 0:
        return "0"
    if tipo == "Garrafa" and familia != "Outro":
        fator = int(fator_conversao_caixas(familia))
        cx, gf = qtd // fator, qtd % fator
        partes = []
        if cx > 0: partes.append(f"{cx} cx")
        if gf > 0: partes.append(f"{gf} gf")
        return " + ".join(partes) if partes else "0"
    return f"{qtd} un"


def formata_diferenca_fisica(dif, tipo: str, familia: str) -> str:
    """Diferença sempre na menor unidade física: 'gf' pra garrafa, 'un' pra tudo mais — com sinal."""
    dif = int(dif)
    if dif == 0:
        return "0"
    sinal = "+" if dif > 0 else "-"
    unidade = "gf" if (tipo == "Garrafa" and familia != "Outro") else "un"
    return f"{sinal}{abs(dif)} {unidade}"


# =========================================================================
# APELIDOS (nomenclatura própria do usuário, exibida ao lado da descrição de origem)
# =========================================================================

MAPA_APELIDOS = {
    "198214": "GARRAFA LITRINHO", "786238": "GARRAFA 600 VERDE", "27983": "GARRAFA 600 VERDE",
    "188006": "GARRAFA 1L", "101490": "BARRIL 50L", "188005": "GARRAFEIRA 1L",
    "863059": "GARRAFEIRA LITRINHO", "899599": "GARRAFEIRA LITRINHO", "104195": "PALLET PBR1",
    "42069": "PALLET PBR2",
}


def com_apelido(codigo: str, rotulo_base: str) -> str:
    apelido = MAPA_APELIDOS.get(str(codigo).strip())
    if apelido: return f"{rotulo_base} ({apelido})"
    return rotulo_base


# =========================================================================
# APOIO PARA O PAINEL EXECUTIVO (usado só pelo app operacional)
# =========================================================================

def coletar_datas_disponiveis(*nomes_abas: str) -> list[str]:
    datas = set()
    for nome_aba in nomes_abas:
        hist = ler_aba_historico(nome_aba)
        if "Data" in hist.columns:
            datas.update(hist["Data"].dropna().astype(str).unique())
    try:
        return sorted(datas, key=lambda d: pd.to_datetime(d, dayfirst=True), reverse=True)
    except Exception:
        return sorted(datas, reverse=True)


def calcular_totais_por_familia(data_str: str, familias_exibicao: list[str]) -> tuple[dict, dict, dict]:
    dict_cheio = {f: 0 for f in familias_exibicao}
    dict_venda = {f: 0 for f in familias_exibicao}
    dict_vazio = {f: 0 for f in familias_exibicao}

    hist_cheio = ler_aba_historico("Cheio")
    if not hist_cheio.empty:
        hist_cheio = hist_cheio[hist_cheio["Data"].astype(str) == data_str]
        for _, row in hist_cheio.iterrows():
            fam = str(row.get("Material", ""))
            if fam in dict_cheio:
                if "Barril" in fam:
                    dict_cheio[fam] += row.get("Barris", 0)
                else:
                    dict_cheio[fam] += row.get("Garrafas", 0)

    hist_venda = ler_aba_historico("Venda")
    if not hist_venda.empty:
        hist_venda = hist_venda[hist_venda["Data"].astype(str) == data_str]
        col_desc = next((c for c in ["Descrição", "Descricao"] if c in hist_venda.columns), None)
        for _, row in hist_venda.iterrows():
            desc = str(row.get(col_desc, "")) if col_desc else ""
            fam = padronizar_familia(desc)
            if fam in dict_venda:
                dict_venda[fam] += row.get("Qtd. Vendida/Movimentada", 0)

    hist_vazio = ler_aba_historico("Vazio")
    if not hist_vazio.empty:
        hist_vazio = hist_vazio[hist_vazio["Data"].astype(str) == data_str]
        for _, row in hist_vazio.iterrows():
            fam = str(row.get("Material", ""))
            if fam in dict_vazio:
                if "Barril" in fam:
                    dict_vazio[fam] += row.get("Unidades", 0)
                else:
                    dict_vazio[fam] += row.get("Garrafas", 0)

    return dict_cheio, dict_venda, dict_vazio
