"""
Funções e constantes compartilhadas entre os apps de AG:
- app_operacional.py  (Painel, Venda, Cheio, Vazio, Vazio por PA, Variação, Dados)
- app_conciliacao.py  (Conciliação Mapas PA, Conciliação Mapas Sede)

Modo de armazenamento: LOCAL (arquivos na mesma pasta do projeto) ou GOOGLE DRIVE
(pra deploy na nuvem, onde não existe uma pasta local persistente). Controlado
pelo secrets.toml — ver instruções no final deste arquivo.
"""

import io
import streamlit as st
import pandas as pd
import re
from pathlib import Path
from datetime import date

PASTA_PROJETO = Path(__file__).parent
ARQUIVO_DE_MATERIAL = PASTA_PROJETO / "De Material.xlsx"
ARQUIVO_PRESTACAO = PASTA_PROJETO / "03.07.13.csv"   # relatório de mapas: Previsto (P) x Realizado (R) por item/mapa
ARQUIVO_MAPAS_AG = ARQUIVO_PRESTACAO  # alias — mesmo arquivo, nome mais claro pra quem lê o código depois de 2026-08
ARQUIVO_COMODATO = PASTA_PROJETO / "02.02.20.csv"    # comodato (emprestado)
ARQUIVO_RET = PASTA_PROJETO / "RET.csv"  # cadastro de produtos retornáveis
ARQUIVO_HISTORICO_EXCEL = PASTA_PROJETO / "historico_ag.xlsx"  # planilha única, uma aba por origem
NOME_HISTORICO_EXCEL = "historico_ag.xlsx"

# NOTA (2026-08): o antigo ARQUIVO_MOVIMENTACAO (02.05.01.csv) foi aposentado —
# o relatório 03.07.13 (ARQUIVO_PRESTACAO/ARQUIVO_MAPAS_AG) já traz Previsto (saída)
# e Realizado (retorno) lado a lado por Mapa+Material, então não é mais necessário
# processar o 02.05.01 (mais pesado e com códigos de operação 554/654 pra decifrar).


# =========================================================================
# GOOGLE DRIVE — ativado via secrets.toml (ver instruções no final do arquivo)
# =========================================================================

def gdrive_ativo() -> bool:
    """True se o secrets.toml tiver [gdrive] configurado com ativo=true."""
    try:
        return bool(st.secrets.get("gdrive", {}).get("ativo", False))
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def _servico_drive():
    from googleapiclient.discovery import build

    if "gdrive_oauth" in st.secrets:
        # Autentica como a própria conta Google do usuário (tem cota de armazenamento normal).
        # Necessário pra ESCREVER — contas de serviço não têm cota própria em Drive pessoal.
        from google.oauth2.credentials import Credentials
        info = st.secrets["gdrive_oauth"]
        credenciais = Credentials(
            token=None,
            refresh_token=info["refresh_token"],
            client_id=info["client_id"],
            client_secret=info["client_secret"],
            token_uri="https://oauth2.googleapis.com/token",
            scopes=["https://www.googleapis.com/auth/drive"],
        )
    else:
        # Conta de serviço: funciona só para LEITURA em Drive pessoal (não tem cota pra escrever).
        from google.oauth2 import service_account
        info = dict(st.secrets["gdrive_service_account"])
        credenciais = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/drive"]
        )
    return build("drive", "v3", credentials=credenciais)


def _pasta_id() -> str:
    return st.secrets["gdrive"]["pasta_id"]


def _com_retry(func, tentativas: int = 3, espera_inicial: float = 1.0):
    """Executa func() com novas tentativas em caso de falha de rede transitória
    (SSL, timeout, conexão), com espera crescente entre elas. Erros de permissão/
    autenticação não são deste tipo e propagam na primeira tentativa mesmo assim
    (não adianta tentar de novo)."""
    import ssl
    import socket
    import time
    erros_transitorios = (ssl.SSLError, socket.timeout, socket.error, ConnectionError, TimeoutError, OSError)
    ultimo_erro = None
    espera = espera_inicial
    for tentativa in range(tentativas):
        try:
            return func()
        except erros_transitorios as e:
            ultimo_erro = e
            if tentativa < tentativas - 1:
                time.sleep(espera)
                espera *= 2
    raise ultimo_erro


def listar_arquivos_pasta(nome_contem: str | None = None) -> list[dict]:
    """Lista arquivos da pasta configurada, mais recentes primeiro. nome_contem filtra por substring do nome."""
    def _fazer():
        servico = _servico_drive()
        query = f"'{_pasta_id()}' in parents and trashed = false"
        if nome_contem:
            query += f" and name contains '{nome_contem}'"
        resultado = servico.files().list(
            q=query, fields="files(id, name, modifiedTime)", orderBy="modifiedTime desc"
        ).execute()
        return resultado.get("files", [])
    return _com_retry(_fazer)


def _baixar_bytes_drive(file_id: str) -> bytes:
    def _fazer():
        from googleapiclient.http import MediaIoBaseDownload
        servico = _servico_drive()
        buffer = io.BytesIO()
        request = servico.files().get_media(fileId=file_id)
        downloader = MediaIoBaseDownload(buffer, request)
        concluido = False
        while not concluido:
            _, concluido = downloader.next_chunk()
        buffer.seek(0)
        return buffer.read()
    return _com_retry(_fazer)


def _subir_bytes_drive(nome_arquivo: str, conteudo: bytes, mimetype: str) -> None:
    """Sobe uma nova versão do arquivo (atualiza se já existir com esse nome na pasta, senão cria)."""
    from googleapiclient.http import MediaIoBaseUpload
    servico = _servico_drive()
    existentes = listar_arquivos_pasta(nome_contem=nome_arquivo)
    existentes = [a for a in existentes if a["name"] == nome_arquivo]
    media = MediaIoBaseUpload(io.BytesIO(conteudo), mimetype=mimetype, resumable=False)
    if existentes:
        servico.files().update(fileId=existentes[0]["id"], media_body=media).execute()
    else:
        metadata = {"name": nome_arquivo, "parents": [_pasta_id()]}
        servico.files().create(body=metadata, media_body=media).execute()


@st.cache_data(show_spinner=False, ttl=60)
def _baixar_arquivo_mais_recente_drive(nome_contem: str) -> bytes | None:
    """Busca o arquivo mais recente cujo nome contenha nome_contem e baixa seu conteúdo.
    Cache de 60s: várias interações seguidas na mesma sessão não batem na API repetidamente,
    mas uma troca de versão no Drive é detectada em até 1 minuto."""
    arquivos = listar_arquivos_pasta(nome_contem)
    if not arquivos:
        return None
    return _baixar_bytes_drive(arquivos[0]["id"])


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


def _ler_csv_bytes(dados: bytes) -> pd.DataFrame:
    encodings = ["utf-8-sig", "latin1", "cp1252"]
    separadores = [";", ",", "\t"]
    ultimo_erro = None
    for enc in encodings:
        for sep in separadores:
            try:
                primeira_linha = dados.decode(enc).split("\n", 1)[0].rstrip("\r")
                colunas = deduplicar_nomes_coluna(primeira_linha.split(sep))
                return pd.read_csv(
                    io.BytesIO(dados), sep=sep, encoding=enc, thousands=".", decimal=",",
                    header=0, names=colunas,
                )
            except Exception as e:
                ultimo_erro = e
    raise ultimo_erro


def ler_csv_robusto(caminho: Path) -> pd.DataFrame:
    with open(caminho, "rb") as f:
        return _ler_csv_bytes(f.read())


@st.cache_data(show_spinner=False)
def carregar(caminho: Path) -> pd.DataFrame | None:
    """Modo LOCAL: lê do disco pelo caminho de sempre.
    Modo DRIVE: ignora o caminho e busca na pasta do Drive um arquivo cujo nome
    contenha o mesmo prefixo do caminho local (ex. 'De Material', '03.07.13', 'RET')."""
    if gdrive_ativo():
        nome_busca = caminho.stem  # "De Material", "03.07.13", "RET", "02.02.20"
        dados = _baixar_arquivo_mais_recente_drive(nome_busca)
        if dados is None:
            return None
        if caminho.suffix.lower() == ".csv":
            return _ler_csv_bytes(dados)
        return pd.read_excel(io.BytesIO(dados))

    if not caminho.exists():
        return None
    if caminho.suffix.lower() == ".csv":
        return ler_csv_robusto(caminho)
    return pd.read_excel(caminho)


def salvar_aba_historico(nome_aba: str, df: pd.DataFrame, nome_arquivo: str = None) -> None:
    nome_arquivo = nome_arquivo or NOME_HISTORICO_EXCEL
    caminho_local = PASTA_PROJETO / nome_arquivo

    if gdrive_ativo():
        historico_atual = _baixar_arquivo_mais_recente_drive(nome_arquivo)
        buffer_saida = io.BytesIO()
        if historico_atual:
            # reabre todas as abas existentes e substitui/adiciona a que mudou
            abas_existentes = pd.read_excel(io.BytesIO(historico_atual), sheet_name=None)
            abas_existentes[nome_aba] = df
            with pd.ExcelWriter(buffer_saida, engine="openpyxl") as writer:
                for aba, conteudo in abas_existentes.items():
                    conteudo.to_excel(writer, sheet_name=aba, index=False)
        else:
            with pd.ExcelWriter(buffer_saida, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=nome_aba, index=False)
        _subir_bytes_drive(
            nome_arquivo, buffer_saida.getvalue(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        _baixar_arquivo_mais_recente_drive.clear()
        return

    if caminho_local.exists():
        with pd.ExcelWriter(caminho_local, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=nome_aba, index=False)
    else:
        with pd.ExcelWriter(caminho_local, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name=nome_aba, index=False)


def ler_aba_historico(nome_aba: str, nome_arquivo: str = None) -> pd.DataFrame:
    nome_arquivo = nome_arquivo or NOME_HISTORICO_EXCEL
    caminho_local = PASTA_PROJETO / nome_arquivo

    if gdrive_ativo():
        dados = _baixar_arquivo_mais_recente_drive(nome_arquivo)
        if dados is None:
            return pd.DataFrame()
        try:
            return pd.read_excel(io.BytesIO(dados), sheet_name=nome_aba)
        except ValueError:
            return pd.DataFrame()

    if not caminho_local.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(caminho_local, sheet_name=nome_aba)
    except ValueError:
        return pd.DataFrame()


# =========================================================================
# ARQUIVAMENTO — mantém historico_ag.xlsx enxuto movendo dados antigos pra
# um segundo arquivo (historico_ag_arquivo.xlsx), consultável sob demanda
# =========================================================================

NOME_HISTORICO_ARQUIVO = "historico_ag_arquivo.xlsx"

ABAS_ARQUIVAVEIS = ["Venda", "Retorno654", "Cheio", "Vazio", "VazioPA"]


def arquivar_dados_antigos(nome_aba: str, meses_manter: int = 6) -> tuple[int, int]:
    """Move linhas com Data mais antiga que `meses_manter` meses do historico_ag.xlsx
    pro historico_ag_arquivo.xlsx, mantendo o arquivo ativo enxuto e rápido.
    Retorna (linhas que ficaram ativas, linhas movidas pro arquivo)."""
    historico = ler_aba_historico(nome_aba)
    if historico.empty or "Data" not in historico.columns:
        return (len(historico), 0)

    historico = historico.copy()
    historico["_dt"] = pd.to_datetime(historico["Data"], dayfirst=True, errors="coerce")
    limite = pd.Timestamp.today() - pd.DateOffset(months=meses_manter)

    recentes = historico[historico["_dt"] >= limite].drop(columns=["_dt"])
    antigos = historico[historico["_dt"] < limite].drop(columns=["_dt"])

    if antigos.empty:
        return (len(recentes), 0)

    arquivo_existente = ler_aba_historico(nome_aba, nome_arquivo=NOME_HISTORICO_ARQUIVO)
    if not arquivo_existente.empty:
        combinado_arquivo = pd.concat([arquivo_existente, antigos], ignore_index=True).drop_duplicates()
    else:
        combinado_arquivo = antigos

    salvar_aba_historico(nome_aba, combinado_arquivo, nome_arquivo=NOME_HISTORICO_ARQUIVO)
    salvar_aba_historico(nome_aba, recentes)
    return (len(recentes), len(antigos))


def normalizar_codigo(serie: pd.Series) -> pd.Series:
    def conv(v):
        if pd.isna(v): return None
        if isinstance(v, float) and v.is_integer(): return str(int(v))
        return str(v).strip()
    return serie.apply(conv)


def limpa_mapa(m):
    """Remove zeros à esquerda e espaços — usada tanto pra número de mapa quanto pra
    código de material zero-padded (ex: '027983' -> '27983'), garantindo que tudo
    casa perfeitamente nos cruzamentos, não importa o relatório de origem."""
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

    marcas_verde = ["SPTN", "SPATEN", "STELLA", "S ARTOIS", "STARTPG", "BECKS", "HEINEKEN", "VERDE"]
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


def localizar_grade_mais_recente(prefixo: str = "02.03.04") -> Path | None:
    """Acha a grade de estoque cheio mais recente (nome com data grudada, ex. 02.03.04.30_07_26.csv).
    Modo LOCAL: procura na pasta do projeto. Modo DRIVE: procura na pasta configurada do Drive.
    Retorna um Path (real ou só com o nome, usado depois só pra extrair stem/suffix e chamar carregar())."""
    if gdrive_ativo():
        arquivos = listar_arquivos_pasta(nome_contem=prefixo)
        return Path(arquivos[0]["name"]) if arquivos else None
    return encontrar_arquivo_por_prefixo(PASTA_PROJETO, prefixo)


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
    if "⏳" in str(val): return "color: #495057; font-weight: bold; background-color: #E9ECEF"
    return ""


def classificar_tipo_generico(desc: str) -> str:
    """Classifica o item em Garrafa / Garrafeira / Barril / Outro, pra decidir como exibir a quantidade física."""
    d = str(desc).upper()
    if "GARRAFEIRA" in d: return "Garrafeira"
    if "BARRIL" in d or "KEG" in d: return "Barril"
    if "GFA" in d or "GARRAFA" in d or "VIDRO" in d: return "Garrafa"
    return "Outro"


def formata_qtd_fisica(qtd, tipo: str, familia: str) -> str:
    """Garrafa (300ml/600ml/Verde 600/1L) vira 'X un (Y cx + Z gf)' — mostra a quantidade bruta E a
    conversão em caixa lado a lado, já que são duas leituras diferentes da mesma coisa.
    Qualquer outro tipo (Garrafeira, Barril, Palete, etc.) vira só 'X un', pois cada unidade ali já
    corresponde fisicamente a uma caixa/unidade só, sem conversão a fazer."""
    qtd = int(qtd)
    if qtd == 0:
        return "0"
    if tipo == "Garrafa" and familia != "Outro":
        fator = int(fator_conversao_caixas(familia))
        cx, gf = qtd // fator, qtd % fator
        partes = []
        if cx > 0: partes.append(f"{cx} cx")
        if gf > 0: partes.append(f"{gf} gf")
        texto_cx = " + ".join(partes) if partes else "0 cx"
        return f"{qtd} un ({texto_cx})"
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
    "198214": "GARRAFA LITRINHO", "786238": "GARRAFA 600 VERDE", "27983": "GARRAFA 600 AMBAR",
    "188006": "GARRAFA 1L", "101490": "BARRIL 50L", "188005": "GARRAFEIRA 1L",
    "863059": "GARRAFEIRA LITRINHO", "899599": "GARRAFEIRA 600", "104195": "PALLET PBR1",
    "42069": "PALLET PBR2",
}

# A descrição da garrafeira nem sempre menciona "300"/"600" (ex: "LITRINHO" sozinho no
# De Material), então padronizar_familia() não consegue classificá-la pela descrição —
# precisa saber pelo código. Também força o Tipo="Garrafeira" em montar_lookup_ag_por_codigo().
GARRAFEIRA_FAMILIA = {"863059": "300ml", "899599": "600ml", "188005": "1L"}

# Rótulo exibido nos formulários de digitação (Vazio, Vazio por PA) — a família continua
# sendo "300ml"/"600ml"/"Verde 600"/"1L" por trás, isso é só o texto que aparece na tela.
RUTULO_FAMILIA_VAZIO = {
    "300ml": "LITRINHO",
    "600ml": "600 AMBAR",
    "Verde 600": "600 VERDE",
    "1L": "LITRÃO",
}


def rotulo_familia_vazio(familia: str) -> str:
    return RUTULO_FAMILIA_VAZIO.get(familia, familia)


def familia_normalizada_600(familia: str) -> str:
    """Agrupa 600ml e Verde 600 numa família só — a mesma garrafeira física (899599) serve às duas cores."""
    return "600ml" if familia in ("600ml", "Verde 600") else familia


# =========================================================================
# CLASSIFICAÇÃO POR CÓDIGO (De Material.xlsx) — 2026-08
# Em vez de tentar interpretar a descrição abreviada de cada relatório (que muda de
# formato entre 02.05.01, 03.07.13 etc., e já causou bug de classificação errada),
# a fonte de verdade agora é sempre a descrição mestre do De Material.xlsx (coluna
# "Desc 2"), consultada pelo Código do item — muito mais estável.
# =========================================================================

def montar_lookup_ag_por_codigo(df_de_material: pd.DataFrame) -> dict:
    """Monta um dicionário Código(Promax) -> (Familia, Tipo), usando a descrição mestre
    do De Material.xlsx. Códigos de garrafeira cuja descrição mestre não menciona volume
    (ex: 863059='LITRINHO', sem "300"/"GARRAFEIRA" no texto) usam o override manual
    GARRAFEIRA_FAMILIA, que também força Tipo='Garrafeira'."""
    lookup: dict[str, tuple[str, str]] = {}
    if df_de_material is not None and "Promax" in df_de_material.columns:
        col_desc = next((c for c in ["Desc 2", "Descricao", "Descrição"] if c in df_de_material.columns), None)
        if col_desc:
            for _, linha in df_de_material.iterrows():
                codigo = limpa_mapa(linha["Promax"])
                desc = str(linha[col_desc])
                lookup[codigo] = (padronizar_familia(desc), classificar_tipo_generico(desc))
    for codigo, familia in GARRAFEIRA_FAMILIA.items():
        lookup[codigo] = (familia, "Garrafeira")
    return lookup


def familia_tipo_por_codigo(codigo, lookup: dict) -> tuple:
    """Consulta o lookup Código->(Familia,Tipo) montado por montar_lookup_ag_por_codigo().
    Cai pra ('Outro','Outro') se o código não estiver cadastrado em nenhum dos dois."""
    return lookup.get(limpa_mapa(codigo), ("Outro", "Outro"))


# =========================================================================
# CATEGORIAS EXTRAS DO RELATÓRIO 03.07.13 — além de Vazio (que já vira Venda/
# Retorno654 e alimenta a Conciliação Mapas PA/Sede principal). Cada tupla:
# (nome de exibição, coluna Previsto no CSV, coluna Realizado no CSV). Guardadas
# numa aba própria do histórico (NOME_ABA_CATEGORIAS_EXTRA), sem passar pela
# conferência manual do Vazio por PA — comparação direta Previsto x Realizado
# do próprio relatório.
# =========================================================================
CATEGORIAS_AG_EXTRA = [
    ("Comodato", "P V Comodato", "R V Comodato"),
    ("Devolução", "P Devol", "R Devol"),
    ("Troca", "P Troca", "R Troca"),
    ("Consignação", "P Consignacao", "R Consignacao"),
    ("Rec. Consignação", "P Rec Consignacao", "R Rec Consignacao"),
]
NOME_ABA_CATEGORIAS_EXTRA = "MapasCategorias"


# =========================================================================
# NOMES DE DEPÓSITO — tabela fixa do Promax (mesma pra todos os armazéns 1/2/3)
# =========================================================================
DEPOSITO_NOMES = {
    "1": "Central",
    "2": "Varejo",
    "3": "Analise",
    "4": "Terceiros",
    "5": "Faltas",
    "6": "Devolucao",
    "8": "Vazio",
    "20": "PECAS E MATERIAIS FROTA E LOG",
    "21": "PECAS/MATERIAIS ADM/VENDA/SEG.",
    "22": "Deposito PNC",
    "23": "Equipamentos SOPIV",
    "24": "Vazio",
}


def nome_deposito(codigo) -> str:
    """Formata o código do depósito com o nome, ex: '1 - Central'. Mantém o código
    visível mesmo com o nome, e cai pro código sozinho se não reconhecer (ou for vazio/'-')."""
    if codigo is None or str(codigo).strip() in ("", "-", "nan", "0"):
        return "-"
    codigo_str = str(codigo).strip()
    if codigo_str.endswith(".0"):
        codigo_str = codigo_str[:-2]
    nome = DEPOSITO_NOMES.get(codigo_str)
    return f"{codigo_str} - {nome}" if nome else codigo_str


def com_apelido(codigo: str, rotulo_base: str) -> str:
    apelido = MAPA_APELIDOS.get(str(codigo).strip())
    if apelido: return f"{rotulo_base} ({apelido})"
    return rotulo_base


# =========================================================================
# APOIO PARA O PAINEL EXECUTIVO (usado só pelo app operacional)
# =========================================================================

def valor_mais_recente_por_grupo(df: pd.DataFrame, colunas_grupo: list[str], coluna_data: str, coluna_valor: str) -> pd.DataFrame:
    """Pra fontes tipo 'retrato' (ex: 03.07.13 — Previsto/Realizado já é o total
    acumulado daquele mapa, não um incremento diário pra somar), pega o valor da
    DATA MAIS RECENTE por grupo em vez de somar todas as datas do histórico — evita
    inflar o total quando o mesmo Mapa+Material foi processado em dias diferentes
    (ex: uma vez pela fonte antiga, outra pela nova, ou reprocessado após correção).
    Retorna um DataFrame só com colunas_grupo + coluna_valor."""
    tmp = df.copy()
    tmp["_dt"] = pd.to_datetime(tmp[coluna_data], dayfirst=True, errors="coerce")
    tmp = tmp.sort_values("_dt")
    tmp = tmp.drop_duplicates(subset=colunas_grupo, keep="last")
    return tmp[colunas_grupo + [coluna_valor]].reset_index(drop=True)


def montar_total_previsto_realizado(hist_venda_vazio: pd.DataFrame, hist_retorno_vazio: pd.DataFrame, hist_categorias: pd.DataFrame) -> pd.DataFrame:
    """Soma Vazio + todas as categorias extras (Comodato, Devolução, Troca, Consignação,
    Rec. Consignação) num único Total Previsto e Total Realizado por Mapa+Material — não
    importa em qual 'espécie' o item saiu ou voltou, só interessa se o total bateu ou não.
    Usa sempre o valor mais recente de cada fonte (não soma histórico de dias diferentes).
    Retorna colunas: Mapa, Material, Total_Previsto, Total_Realizado."""
    partes_p, partes_r = [], []

    if hist_venda_vazio is not None and not hist_venda_vazio.empty and "Qtd. Vendida/Movimentada" in hist_venda_vazio.columns:
        vp = valor_mais_recente_por_grupo(hist_venda_vazio, ["Mapa", "Material"], "Data", "Qtd. Vendida/Movimentada")
        partes_p.append(vp.rename(columns={"Qtd. Vendida/Movimentada": "Valor"}))

    if hist_retorno_vazio is not None and not hist_retorno_vazio.empty and "Qtd_Retorno_654" in hist_retorno_vazio.columns:
        vr = valor_mais_recente_por_grupo(hist_retorno_vazio, ["Mapa", "Material"], "Data", "Qtd_Retorno_654")
        partes_r.append(vr.rename(columns={"Qtd_Retorno_654": "Valor"}))

    if hist_categorias is not None and not hist_categorias.empty:
        for _, col_p, col_r in CATEGORIAS_AG_EXTRA:
            if col_p in hist_categorias.columns:
                cp = valor_mais_recente_por_grupo(hist_categorias, ["Mapa", "Material"], "Data", col_p)
                partes_p.append(cp.rename(columns={col_p: "Valor"}))
            if col_r in hist_categorias.columns:
                cr = valor_mais_recente_por_grupo(hist_categorias, ["Mapa", "Material"], "Data", col_r)
                partes_r.append(cr.rename(columns={col_r: "Valor"}))

    df_p = (
        pd.concat(partes_p, ignore_index=True).groupby(["Mapa", "Material"])["Valor"].sum().reset_index().rename(columns={"Valor": "Total_Previsto"})
        if partes_p else pd.DataFrame(columns=["Mapa", "Material", "Total_Previsto"])
    )
    df_r = (
        pd.concat(partes_r, ignore_index=True).groupby(["Mapa", "Material"])["Valor"].sum().reset_index().rename(columns={"Valor": "Total_Realizado"})
        if partes_r else pd.DataFrame(columns=["Mapa", "Material", "Total_Realizado"])
    )

    return pd.merge(df_p, df_r, on=["Mapa", "Material"], how="outer").fillna(0)


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


def calcular_totais_por_familia(data_str: str, familias_exibicao: list[str], lookup_codigo: dict | None = None) -> tuple[dict, dict, dict]:
    """lookup_codigo (opcional): dict Código->(Familia,Tipo) de montar_lookup_ag_por_codigo().
    Quando fornecido, a Venda é classificada pelo Código do Material (mais confiável);
    sem ele, cai pro método antigo de interpretar o texto da Descrição."""
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
            if lookup_codigo is not None and "Material" in hist_venda.columns:
                fam = familia_tipo_por_codigo(row.get("Material", ""), lookup_codigo)[0]
            else:
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
