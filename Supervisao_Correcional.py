import time
import random
import uuid
import re
import json
from typing import Any
from http.client import RemoteDisconnected
from datetime import datetime, date, time as dtime
from zoneinfo import ZoneInfo

import requests
import streamlit as st
import pandas as pd
import gspread
from gspread.exceptions import APIError, SpreadsheetNotFound, WorksheetNotFound
from google.oauth2.service_account import Credentials


# ==========================================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================================
st.set_page_config(
    page_title="SUPERVISÃO CORRECIONAL",
    page_icon="📋",
    layout="wide"
)


# ==========================================================
# CONFIGURAÇÕES GERAIS
# ==========================================================
NOME_APP = "SUPERVISÃO CORRECIONAL"
ID_PLANILHA = "1bVz1b4B-0avm59LO9Iju69v6E46yAJlZn2cHjp-FzHs"
NOME_PLANILHA = "Supervisao_Correcional"
ABA_CADASTRO = "CADASTRO"
FUSO_BR = ZoneInfo("America/Sao_Paulo")

ADMIN_USUARIO = "Adm"
ADMIN_SENHA = "123"

UNIDADES_SERVICO = [
    "1ª DPJM", "2ª DPJM", "3ª DPJM", "4ª DPJM", "5ª DPJM",
    "6ª DPJM", "7ª DPJM", "8ª DPJM", "CGPM",
]

POSTOS_GRADUACOES = [
    "",
    "CEL",
    "TCEL",
    "MAJ",
    "CAP",
    "1º TEN",
    "2ª TEN",
    "ASP",
    "CADETE",
    "SUBTEN",
    "1º SGT",
    "2º SGT",
    "3º SGT",
    "CB",
    "SD",
    "ALUNO CFSD",
    "ALUNO CURSO CORRECIONAL",
]

COLUNAS_CADASTRO = [
    "ID", "ORDEM", "POSTO_GRADUACAO", "RG", "NOME_DE_ESCALA",
    "SENHA", "AUTORIZADO", "ATUALIZADO_EM", "VERSAO",
]

COLUNAS_SERVICO = [
    "ID_SERVICO", "STATUS", "UNIDADE", "DATA", "INICIO_SERVICO",
    "TERMINO_SERVICO", "VIATURA", "KM_INICIAL", "KM_FINAL",
    "NUM_AREAS", "NOMES_AREAS_JSON", "ABAS_AREAS_CRIADAS_JSON",
    "SUPERVISOR", "MOTORISTA", "SEGURANCA_1", "SEGURANCA_2",
    "SEGURANCA_3", "OBSERVACOES_GERAIS", "CRIADO_POR_ID",
    "CRIADO_POR_RG", "CRIADO_POR_NOME", "CRIADO_EM", "ATUALIZADO_EM",
    "VERSAO",
]

COLUNAS_AREA = [
    "ID_SERVICO", "UNIDADE_ORIGEM", "NOME_AREA", "DATA_SERVICO",
    "STATUS_SERVICO", "CRIADO_EM",
]


# ==========================================================
# ESTILO
# ==========================================================
st.markdown(
    """
    <style>
        .main-title {
            text-align: center;
            font-size: 36px;
            font-weight: 900;
            margin-bottom: 4px;
            letter-spacing: 0.5px;
        }
        .sub-title {
            text-align: center;
            font-size: 16px;
            color: #555;
            margin-bottom: 28px;
        }
        .service-title {
            text-align: center;
            font-size: 34px;
            font-weight: 900;
            margin-top: 8px;
            margin-bottom: 18px;
            letter-spacing: 0.8px;
        }
        .status-box {
            padding: 14px 16px;
            border-radius: 14px;
            background-color: #f8f9fa;
            border: 1px solid #dee2e6;
            margin-bottom: 18px;
            font-weight: 650;
            line-height: 1.5;
        }
        .blocked-box {
            padding: 18px;
            border-radius: 14px;
            background-color: #fff3cd;
            border: 1px solid #ffecb5;
            color: #664d03;
            font-weight: 750;
            margin-top: 20px;
            line-height: 1.6;
        }
        .confirm-box {
            padding: 16px;
            border-radius: 14px;
            background-color: #fff7ed;
            border: 1px solid #fed7aa;
            color: #7c2d12;
            font-weight: 700;
            margin-top: 14px;
            margin-bottom: 14px;
            line-height: 1.5;
        }
        div.stButton > button,
        div[data-testid="stFormSubmitButton"] button {
            width: 100%;
            min-height: 43px;
            border-radius: 12px;
            font-weight: 800;
            border: 1px solid #d0d7de;
            box-shadow: 0 1px 4px rgba(0,0,0,0.08);
            transition: all 0.15s ease-in-out;
        }
        div.stButton > button:hover,
        div[data-testid="stFormSubmitButton"] button:hover {
            transform: translateY(-1px);
            box-shadow: 0 3px 10px rgba(0,0,0,0.12);
        }
        div[data-testid="stTextInput"] input,
        div[data-testid="stNumberInput"] input {
            border-radius: 10px;
            min-height: 42px;
        }
        button[data-baseweb="tab"] {
            font-weight: 800;
            font-size: 15px;
        }
        .small-muted {
            color: #666;
            font-size: 13px;
            margin-top: -8px;
            margin-bottom: 10px;
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ==========================================================
# FUNÇÕES AUXILIARES
# ==========================================================
def agora_br() -> str:
    return datetime.now(FUSO_BR).strftime("%Y-%m-%d %H:%M:%S")


def gerar_id_usuario() -> str:
    data_atual = datetime.now(FUSO_BR).strftime("%Y%m%d%H%M%S")
    sufixo = uuid.uuid4().hex[:8].upper()
    return f"USR-{data_atual}-{sufixo}"


def gerar_id_servico() -> str:
    data_atual = datetime.now(FUSO_BR).strftime("%Y%m%d%H%M%S")
    sufixo = uuid.uuid4().hex[:8].upper()
    return f"SRV-{data_atual}-{sufixo}"


def normalizar_texto(valor: Any) -> str:
    if valor is None:
        return ""
    return str(valor).strip()


def normalizar_rg(valor: Any) -> str:
    if valor is None:
        return ""
    return str(valor).strip()


def apenas_digitos(valor: Any) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def para_int(valor: Any, padrao: int = 0) -> int:
    try:
        if valor is None:
            return padrao
        valor_str = str(valor).strip()
        if not valor_str:
            return padrao
        return int(float(valor_str))
    except Exception:
        return padrao


def normalizar_autorizado(valor: Any) -> str:
    valor = normalizar_texto(valor).upper()
    return "SIM" if valor == "SIM" else "NÃO"


def texto_caixa_alta(valor: Any) -> str:
    return normalizar_texto(valor).upper()


def data_para_texto(valor: Any) -> str:
    if valor is None:
        return ""
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    return normalizar_texto(valor)


def hora_para_texto(valor: Any) -> str:
    if valor is None:
        return ""
    if isinstance(valor, dtime):
        return valor.strftime("%H:%M")
    return normalizar_texto(valor)


def texto_para_data(valor: Any):
    """
    Converte a data salva no Google Sheets para objeto date,
    permitindo que o date_input volte preenchido ao editar.
    Aceita formatos DD/MM/YYYY e YYYY-MM-DD.
    """
    if valor is None:
        return None

    if isinstance(valor, date) and not isinstance(valor, datetime):
        return valor

    if isinstance(valor, datetime):
        return valor.date()

    valor_str = normalizar_texto(valor)

    if not valor_str:
        return None

    formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]

    for formato in formatos:
        try:
            return datetime.strptime(valor_str, formato).date()
        except Exception:
            pass

    return None


def texto_para_hora(valor: Any):
    """
    Converte a hora salva no Google Sheets para objeto time,
    permitindo que o time_input volte preenchido ao editar.
    Aceita HH:MM e HH:MM:SS.
    """
    if valor is None:
        return None

    if isinstance(valor, dtime):
        return valor

    if isinstance(valor, datetime):
        return valor.time().replace(second=0, microsecond=0)

    valor_str = normalizar_texto(valor)

    if not valor_str:
        return None

    formatos = ["%H:%M", "%H:%M:%S"]

    for formato in formatos:
        try:
            return datetime.strptime(valor_str, formato).time()
        except Exception:
            pass

    return None



def limpar_nome_aba(nome: str) -> str:
    nome = texto_caixa_alta(nome)
    nome = re.sub(r"[\[\]\:\*\?\/\\]", "-", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    if not nome:
        nome = "AREA"
    return nome[:95]


def localizar_indices_por_rg(df: pd.DataFrame, rg_busca: str) -> list:
    termo = normalizar_rg(rg_busca)
    termo_digitos = apenas_digitos(termo)

    if df.empty or not termo:
        return []

    exatos = []
    parciais = []

    for indice, linha in df.iterrows():
        rg_atual = normalizar_rg(linha.get("RG", ""))
        rg_atual_digitos = apenas_digitos(rg_atual)

        if rg_atual == termo or (termo_digitos and rg_atual_digitos == termo_digitos):
            exatos.append(indice)
            continue

        if termo.lower() in rg_atual.lower() or (
            termo_digitos and termo_digitos in rg_atual_digitos
        ):
            parciais.append(indice)

    return exatos if exatos else parciais


def json_dumps_lista(lista: list) -> str:
    return json.dumps(lista, ensure_ascii=False)


def json_loads_lista(valor: Any) -> list:
    try:
        if not valor:
            return []
        resultado = json.loads(str(valor))
        if isinstance(resultado, list):
            return resultado
        return []
    except Exception:
        return []


def indice_opcao(opcoes: list, valor: Any) -> int:
    valor = normalizar_texto(valor)
    if valor in opcoes:
        return opcoes.index(valor)
    return 0


def montar_opcoes_posto_graduacao(valor_atual: Any = "") -> list:
    """
    Monta a lista do campo Posto/Graduação.
    Se houver valor antigo já salvo fora da lista, ele é mantido temporariamente
    para permitir visualizar/editar o cadastro sem apagar informação existente.
    """
    opcoes = POSTOS_GRADUACOES.copy()
    valor_atual = normalizar_texto(valor_atual).upper()

    if valor_atual and valor_atual not in opcoes:
        opcoes.append(valor_atual)

    return opcoes


# ==========================================================
# GOOGLE SHEETS COM RETRY E CACHE
# ==========================================================
def executar_com_retry(funcao, *args, tentativas=6, espera_inicial=1, **kwargs):
    ultimo_erro = None

    for tentativa in range(1, tentativas + 1):
        try:
            return funcao(*args, **kwargs)

        except (
            APIError,
            requests.exceptions.ConnectionError,
            requests.exceptions.Timeout,
            requests.exceptions.RequestException,
            RemoteDisconnected,
            ConnectionResetError,
            TimeoutError,
        ) as erro:
            ultimo_erro = erro
            mensagem = str(erro)

            erro_temporario = (
                "429" in mensagem
                or "Quota" in mensagem
                or "Quota exceeded" in mensagem
                or "RESOURCE_EXHAUSTED" in mensagem
                or "RemoteDisconnected" in mensagem
                or "Connection aborted" in mensagem
                or "Connection reset" in mensagem
                or "Read timed out" in mensagem
                or isinstance(
                    erro,
                    (
                        requests.exceptions.ConnectionError,
                        requests.exceptions.Timeout,
                        requests.exceptions.RequestException,
                        RemoteDisconnected,
                        ConnectionResetError,
                        TimeoutError,
                    ),
                )
            )

            if tentativa == tentativas:
                raise

            if erro_temporario:
                if "429" in mensagem or "Quota exceeded" in mensagem or "Quota" in mensagem:
                    espera = min(65, 10 * tentativa)
                else:
                    espera = espera_inicial * (2 ** (tentativa - 1)) + random.uniform(0, 1.5)
                time.sleep(espera)
            else:
                raise

    if ultimo_erro:
        raise ultimo_erro


@st.cache_resource
def conectar_google_sheets():
    info_credenciais = dict(st.secrets["gcp_service_account"])

    if "private_key" in info_credenciais:
        info_credenciais["private_key"] = info_credenciais["private_key"].replace("\\n", "\n")

    escopos = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    credenciais = Credentials.from_service_account_info(
        info_credenciais,
        scopes=escopos,
    )

    cliente = gspread.authorize(credenciais)
    return cliente


@st.cache_resource
def abrir_planilha_cacheada():
    cliente = conectar_google_sheets()
    return executar_com_retry(cliente.open_by_key, ID_PLANILHA)


def abrir_planilha():
    try:
        return abrir_planilha_cacheada()

    except SpreadsheetNotFound:
        st.error(
            "❌ A planilha não foi encontrada pelo ID informado.\n\n"
            "Confira se o ID da planilha está correto e se ela foi compartilhada "
            "com o e-mail da Service Account."
        )
        st.stop()

    except APIError as erro:
        mensagem = str(erro)

        if "429" in mensagem or "Quota exceeded" in mensagem or "Quota" in mensagem:
            st.error(
                "❌ O Google Sheets bloqueou temporariamente por excesso de leituras.\n\n"
                "Aguarde cerca de 1 minuto e tente novamente. "
                "Esta versão já reduz as leituras ao clicar nas guias e ao usar Assunção do Serviço."
            )
        else:
            st.error(
                "❌ Erro ao abrir a planilha pelo ID.\n\n"
                "Verifique se a planilha foi compartilhada com a Service Account como Editor."
            )

        st.exception(erro)
        st.stop()


def cache_headers_ok() -> set:
    if "_headers_ok" not in st.session_state:
        st.session_state["_headers_ok"] = set()
    return st.session_state["_headers_ok"]


def encontrar_worksheet(planilha, nome_aba: str):
    try:
        return executar_com_retry(planilha.worksheet, nome_aba)
    except WorksheetNotFound:
        pass

    # Só procura ignorando maiúsculas/minúsculas quando a busca exata falha.
    abas = executar_com_retry(planilha.worksheets)
    nome_normalizado = nome_aba.strip().lower()

    for aba in abas:
        if aba.title.strip().lower() == nome_normalizado:
            return aba

    return None


def garantir_worksheet(nome_aba: str, colunas: list, verificar_cabecalho: bool = True):
    planilha = abrir_planilha()
    ws = encontrar_worksheet(planilha, nome_aba)

    if ws is None:
        ws = executar_com_retry(
            planilha.add_worksheet,
            title=nome_aba,
            rows=1000,
            cols=max(len(colunas), 10),
        )
        intervalo = f"A1:{chr(64 + len(colunas))}1"
        executar_com_retry(ws.update, intervalo, [colunas])
        cache_headers_ok().add(nome_aba)
        return ws

    if verificar_cabecalho and nome_aba not in cache_headers_ok():
        cabecalho = [normalizar_texto(c) for c in executar_com_retry(ws.row_values, 1)]
        if cabecalho[:len(colunas)] != colunas:
            intervalo = f"A1:{chr(64 + len(colunas))}1"
            executar_com_retry(ws.update, intervalo, [colunas])
        cache_headers_ok().add(nome_aba)

    return ws


# ==========================================================
# CADASTRO
# ==========================================================
def migrar_estrutura_cadastro(ws):
    if st.session_state.get("_cadastro_migrado", False):
        return

    cabecalho_atual = [normalizar_texto(c) for c in executar_com_retry(ws.row_values, 1)]

    if not cabecalho_atual:
        executar_com_retry(ws.update, "A1:I1", [COLUNAS_CADASTRO])
        st.session_state["_cadastro_migrado"] = True
        cache_headers_ok().add(ABA_CADASTRO)
        return

    if cabecalho_atual == COLUNAS_CADASTRO:
        st.session_state["_cadastro_migrado"] = True
        cache_headers_ok().add(ABA_CADASTRO)
        return

    valores = executar_com_retry(ws.get_all_values)

    if not valores:
        executar_com_retry(ws.update, "A1:I1", [COLUNAS_CADASTRO])
        st.session_state["_cadastro_migrado"] = True
        cache_headers_ok().add(ABA_CADASTRO)
        return

    cabecalho_atual = [normalizar_texto(c) for c in valores[0]]
    linhas = valores[1:]
    mapa_colunas = {nome: idx for idx, nome in enumerate(cabecalho_atual)}

    novos_registros = []
    ids_usados = set()

    for indice_linha, linha in enumerate(linhas, start=1):
        if not any(normalizar_texto(c) for c in linha):
            continue

        registro = {}
        for coluna in COLUNAS_CADASTRO:
            if coluna in mapa_colunas:
                idx = mapa_colunas[coluna]
                registro[coluna] = linha[idx] if idx < len(linha) else ""
            else:
                registro[coluna] = ""

        id_usuario = normalizar_texto(registro.get("ID", ""))
        if not id_usuario or id_usuario in ids_usados:
            id_usuario = gerar_id_usuario()
        ids_usados.add(id_usuario)

        ordem = para_int(registro.get("ORDEM", ""), indice_linha)
        if ordem <= 0:
            ordem = indice_linha

        atualizado_em = normalizar_texto(registro.get("ATUALIZADO_EM", "")) or agora_br()
        versao = para_int(registro.get("VERSAO", ""), 1)
        if versao <= 0:
            versao = 1

        novos_registros.append({
            "ID": id_usuario,
            "ORDEM": ordem,
            "POSTO_GRADUACAO": normalizar_texto(registro.get("POSTO_GRADUACAO", "")),
            "RG": normalizar_rg(registro.get("RG", "")),
            "NOME_DE_ESCALA": normalizar_texto(registro.get("NOME_DE_ESCALA", "")),
            "SENHA": normalizar_texto(registro.get("SENHA", "")),
            "AUTORIZADO": normalizar_autorizado(registro.get("AUTORIZADO", "NÃO")),
            "ATUALIZADO_EM": atualizado_em,
            "VERSAO": versao,
        })

    valores_novos = [COLUNAS_CADASTRO]
    for i, registro in enumerate(novos_registros, start=1):
        if para_int(registro.get("ORDEM", ""), 0) <= 0:
            registro["ORDEM"] = i
        valores_novos.append([registro[coluna] for coluna in COLUNAS_CADASTRO])

    executar_com_retry(ws.clear)
    executar_com_retry(ws.update, "A1", valores_novos)

    try:
        linhas_necessarias = max(1000, len(valores_novos) + 20)
        executar_com_retry(ws.resize, rows=linhas_necessarias, cols=len(COLUNAS_CADASTRO))
    except Exception:
        pass

    st.session_state["_cadastro_migrado"] = True
    cache_headers_ok().add(ABA_CADASTRO)


def obter_worksheet_cadastro():
    planilha = abrir_planilha()
    ws = encontrar_worksheet(planilha, ABA_CADASTRO)

    if ws is None:
        ws = executar_com_retry(
            planilha.add_worksheet,
            title=ABA_CADASTRO,
            rows=1000,
            cols=len(COLUNAS_CADASTRO),
        )
        executar_com_retry(ws.update, "A1:I1", [COLUNAS_CADASTRO])
        st.session_state["_cadastro_migrado"] = True
        cache_headers_ok().add(ABA_CADASTRO)
        return ws

    migrar_estrutura_cadastro(ws)
    return ws


def preparar_dataframe_cadastro(df: pd.DataFrame) -> pd.DataFrame:
    for coluna in COLUNAS_CADASTRO:
        if coluna not in df.columns:
            df[coluna] = ""

    df = df[COLUNAS_CADASTRO].copy()

    if df.empty:
        return df

    df["ID"] = df["ID"].astype(str).fillna("").str.strip()
    df["ORDEM"] = pd.to_numeric(df["ORDEM"], errors="coerce").fillna(0).astype(int)
    df["VERSAO"] = pd.to_numeric(df["VERSAO"], errors="coerce").fillna(1).astype(int)

    for coluna in [
        "POSTO_GRADUACAO", "RG", "NOME_DE_ESCALA", "SENHA",
        "AUTORIZADO", "ATUALIZADO_EM",
    ]:
        df[coluna] = df[coluna].astype(str).fillna("").str.strip()

    df["AUTORIZADO"] = df["AUTORIZADO"].apply(normalizar_autorizado)
    df = df[df["ID"].astype(str).str.strip() != ""].copy()

    df["_RG_NUM"] = df["RG"].apply(lambda x: para_int(apenas_digitos(x), 999999999))
    df = df.sort_values(["_RG_NUM", "RG", "NOME_DE_ESCALA"], na_position="last").reset_index(drop=True)
    df = df.drop(columns=["_RG_NUM"], errors="ignore")

    return df


@st.cache_data(ttl=45)
def carregar_cadastros() -> pd.DataFrame:
    ws = obter_worksheet_cadastro()
    registros = executar_com_retry(ws.get_all_records)
    df = pd.DataFrame(registros)
    return preparar_dataframe_cadastro(df)


def limpar_cache():
    carregar_cadastros.clear()


def registro_cadastro_para_linha(registro: dict) -> list:
    return [registro.get(coluna, "") for coluna in COLUNAS_CADASTRO]


def localizar_linha_cadastro_por_id(id_usuario: str):
    ws = obter_worksheet_cadastro()
    valores = executar_com_retry(ws.get_all_values)

    if not valores:
        return None, None

    cabecalho = [normalizar_texto(c) for c in valores[0]]
    if "ID" not in cabecalho:
        return None, None

    idx_id = cabecalho.index("ID")

    for numero_linha, linha in enumerate(valores[1:], start=2):
        linha_pad = linha + [""] * (len(cabecalho) - len(linha))

        if idx_id < len(linha_pad) and normalizar_texto(linha_pad[idx_id]) == id_usuario:
            registro = {}
            for coluna in COLUNAS_CADASTRO:
                if coluna in cabecalho:
                    idx = cabecalho.index(coluna)
                    registro[coluna] = linha_pad[idx] if idx < len(linha_pad) else ""
                else:
                    registro[coluna] = ""

            registro["AUTORIZADO"] = normalizar_autorizado(registro.get("AUTORIZADO", "NÃO"))
            registro["VERSAO"] = para_int(registro.get("VERSAO", ""), 1)
            registro["ORDEM"] = para_int(registro.get("ORDEM", ""), 0)
            return numero_linha, registro

    return None, None


def rg_ja_existe(df: pd.DataFrame, rg: str, ignorar_id: str | None = None) -> bool:
    rg = normalizar_rg(rg)
    if not rg or df.empty:
        return False

    for _, linha in df.iterrows():
        id_atual = normalizar_texto(linha.get("ID", ""))
        if ignorar_id is not None and id_atual == ignorar_id:
            continue
        if normalizar_rg(linha.get("RG", "")) == rg:
            return True

    return False


def proxima_ordem(df: pd.DataFrame) -> int:
    if df.empty:
        return 1
    ordens = pd.to_numeric(df["ORDEM"], errors="coerce").fillna(0).astype(int)
    if ordens.empty:
        return 1
    return int(ordens.max()) + 1


def adicionar_usuario(posto: str, rg: str, nome: str, senha: str):
    limpar_cache()
    df_atual = carregar_cadastros()

    if rg_ja_existe(df_atual, rg):
        return False, "❌ Já existe cadastro com esse RG."

    ws = obter_worksheet_cadastro()

    novo_registro = {
        "ID": gerar_id_usuario(),
        "ORDEM": proxima_ordem(df_atual),
        "POSTO_GRADUACAO": posto,
        "RG": rg,
        "NOME_DE_ESCALA": nome,
        "SENHA": senha,
        "AUTORIZADO": "NÃO",
        "ATUALIZADO_EM": agora_br(),
        "VERSAO": 1,
    }

    executar_com_retry(
        ws.append_row,
        registro_cadastro_para_linha(novo_registro),
        value_input_option="RAW",
    )

    limpar_cache()
    return True, novo_registro["ID"]


def atualizar_usuario_por_id(id_usuario: str, versao_esperada: int, posto: str, rg: str, nome: str, senha_nova: str):
    limpar_cache()
    df_atual = carregar_cadastros()

    if rg_ja_existe(df_atual, rg, ignorar_id=id_usuario):
        return False, "❌ Já existe outro cadastro com esse RG."

    ws = obter_worksheet_cadastro()
    numero_linha, registro_atual = localizar_linha_cadastro_por_id(id_usuario)

    if numero_linha is None or registro_atual is None:
        return False, "⚠️ Este cadastro não foi encontrado. Ele pode ter sido excluído por outro usuário."

    versao_atual = para_int(registro_atual.get("VERSAO", ""), 1)

    if versao_atual != para_int(versao_esperada, 1):
        return False, "⚠️ Este cadastro foi alterado por outro usuário ou pelo Administrador. Cancele e abra novamente antes de salvar."

    senha_final = senha_nova if senha_nova else normalizar_texto(registro_atual.get("SENHA", ""))

    registro_atualizado = {
        "ID": id_usuario,
        "ORDEM": para_int(registro_atual.get("ORDEM", ""), 0),
        "POSTO_GRADUACAO": posto,
        "RG": rg,
        "NOME_DE_ESCALA": nome,
        "SENHA": senha_final,
        "AUTORIZADO": normalizar_autorizado(registro_atual.get("AUTORIZADO", "NÃO")),
        "ATUALIZADO_EM": agora_br(),
        "VERSAO": versao_atual + 1,
    }

    executar_com_retry(ws.update, f"A{numero_linha}:I{numero_linha}", [registro_cadastro_para_linha(registro_atualizado)])
    limpar_cache()
    return True, "✅ Cadastro atualizado com sucesso."


def excluir_usuario_por_id(id_usuario: str, senha_digitada: str):
    ws = obter_worksheet_cadastro()
    numero_linha, registro_atual = localizar_linha_cadastro_por_id(id_usuario)

    if numero_linha is None or registro_atual is None:
        return False, "⚠️ Este cadastro não foi encontrado. Ele pode já ter sido excluído."

    senha_salva = normalizar_texto(registro_atual.get("SENHA", ""))
    if normalizar_texto(senha_digitada) != senha_salva:
        return False, "❌ Senha incorreta. Exclusão não autorizada."

    executar_com_retry(ws.delete_rows, numero_linha)
    limpar_cache()
    return True, "✅ Cadastro excluído com sucesso."


def atualizar_autorizacao_por_id(id_usuario: str, autorizado: bool, versao_esperada: int):
    ws = obter_worksheet_cadastro()
    numero_linha, registro_atual = localizar_linha_cadastro_por_id(id_usuario)

    if numero_linha is None or registro_atual is None:
        return False, "Cadastro não encontrado."

    versao_atual = para_int(registro_atual.get("VERSAO", ""), 1)
    if versao_atual != para_int(versao_esperada, 1):
        return False, "Cadastro alterado por outro usuário."

    registro_atualizado = {
        "ID": id_usuario,
        "ORDEM": para_int(registro_atual.get("ORDEM", ""), 0),
        "POSTO_GRADUACAO": normalizar_texto(registro_atual.get("POSTO_GRADUACAO", "")),
        "RG": normalizar_rg(registro_atual.get("RG", "")),
        "NOME_DE_ESCALA": normalizar_texto(registro_atual.get("NOME_DE_ESCALA", "")),
        "SENHA": normalizar_texto(registro_atual.get("SENHA", "")),
        "AUTORIZADO": "SIM" if autorizado else "NÃO",
        "ATUALIZADO_EM": agora_br(),
        "VERSAO": versao_atual + 1,
    }

    executar_com_retry(ws.update, f"A{numero_linha}:I{numero_linha}", [registro_cadastro_para_linha(registro_atualizado)])
    limpar_cache()
    return True, "OK"


def validar_usuario_logado(df_atual: pd.DataFrame):
    if not st.session_state.usuario_logado:
        return

    dados = st.session_state.dados_usuario_logado
    if not dados:
        st.session_state.usuario_logado = False
        st.session_state.dados_usuario_logado = None
        return

    id_usuario = normalizar_texto(dados.get("ID", ""))
    rg = normalizar_rg(dados.get("RG", ""))
    senha = normalizar_texto(dados.get("SENHA", ""))

    if df_atual.empty:
        st.session_state.usuario_logado = False
        st.session_state.dados_usuario_logado = None
        return

    if id_usuario:
        usuario = df_atual[df_atual["ID"].astype(str).str.strip() == id_usuario]
    else:
        usuario = df_atual[
            (df_atual["RG"].astype(str).str.strip() == rg)
            &
            (df_atual["SENHA"].astype(str).str.strip() == senha)
        ]

    if usuario.empty:
        st.session_state.usuario_logado = False
        st.session_state.dados_usuario_logado = None
        return

    dados_atualizados = usuario.iloc[0].to_dict()

    if normalizar_texto(dados_atualizados.get("SENHA", "")) != senha:
        st.session_state.usuario_logado = False
        st.session_state.dados_usuario_logado = None
        return

    autorizado = normalizar_texto(dados_atualizados.get("AUTORIZADO", "")).upper()
    if autorizado != "SIM":
        st.session_state.usuario_logado = False
        st.session_state.dados_usuario_logado = None
        return

    st.session_state.dados_usuario_logado = dados_atualizados


# ==========================================================
# SERVIÇO
# ==========================================================
def registro_servico_para_linha(registro: dict) -> list:
    return [registro.get(coluna, "") for coluna in COLUNAS_SERVICO]


def localizar_linha_servico_por_id(unidade: str, id_servico: str):
    ws = garantir_worksheet(unidade, COLUNAS_SERVICO)
    valores = executar_com_retry(ws.get_all_values)

    if not valores:
        return None, None

    cabecalho = [normalizar_texto(c) for c in valores[0]]
    if "ID_SERVICO" not in cabecalho:
        return None, None

    idx_id = cabecalho.index("ID_SERVICO")

    for numero_linha, linha in enumerate(valores[1:], start=2):
        linha_pad = linha + [""] * (len(cabecalho) - len(linha))
        if idx_id < len(linha_pad) and normalizar_texto(linha_pad[idx_id]) == id_servico:
            registro = {}
            for coluna in COLUNAS_SERVICO:
                if coluna in cabecalho:
                    idx = cabecalho.index(coluna)
                    registro[coluna] = linha_pad[idx] if idx < len(linha_pad) else ""
                else:
                    registro[coluna] = ""
            registro["VERSAO"] = para_int(registro.get("VERSAO", ""), 1)
            return numero_linha, registro

    return None, None


def criar_ou_obter_aba_area(nome_area: str, id_servico: str, unidade: str, data_servico: str, nomes_existentes: set):
    planilha = abrir_planilha()
    nome_base = limpar_nome_aba(nome_area)

    if nome_base in nomes_existentes:
        ws = encontrar_worksheet(planilha, nome_base)
        criada = False
        nome_final = nome_base
    else:
        nome_final = nome_base
        ws = executar_com_retry(
            planilha.add_worksheet,
            title=nome_final,
            rows=1000,
            cols=len(COLUNAS_AREA),
        )
        executar_com_retry(ws.update, "A1:F1", [COLUNAS_AREA])
        nomes_existentes.add(nome_final)
        cache_headers_ok().add(nome_final)
        criada = True

    if not criada and nome_final not in cache_headers_ok():
        cabecalho = [normalizar_texto(c) for c in executar_com_retry(ws.row_values, 1)]
        if cabecalho[:len(COLUNAS_AREA)] != COLUNAS_AREA:
            executar_com_retry(ws.update, "A1:F1", [COLUNAS_AREA])
        cache_headers_ok().add(nome_final)

    linha_area = [
        id_servico,
        unidade,
        texto_caixa_alta(nome_area),
        data_servico,
        "ABERTO",
        agora_br(),
    ]
    executar_com_retry(ws.append_row, linha_area, value_input_option="RAW")
    return nome_final, criada


def obter_nomes_abas_existentes_set() -> set:
    planilha = abrir_planilha()
    abas = executar_com_retry(planilha.worksheets)
    return {aba.title for aba in abas}


def excluir_abas_areas_criadas(nomes_abas: list):
    planilha = abrir_planilha()

    for nome_aba in nomes_abas:
        nome_aba = normalizar_texto(nome_aba)
        if not nome_aba:
            continue
        if nome_aba in [ABA_CADASTRO] + UNIDADES_SERVICO:
            continue
        ws = encontrar_worksheet(planilha, nome_aba)
        if ws is not None:
            executar_com_retry(planilha.del_worksheet, ws)
            cache_headers_ok().discard(nome_aba)


def adicionar_servico(dados: dict):
    unidade = dados["UNIDADE"]
    id_servico = gerar_id_servico()
    data_servico = dados["DATA"]
    areas = dados["AREAS"]

    # IMPORTANTE: a aba da unidade só é criada/acessada no SALVAR.
    # O clique na guia ASSUNÇÃO DO SERVIÇO não cria nem verifica todas as abas.
    nomes_existentes = obter_nomes_abas_existentes_set()
    abas_areas_criadas = []

    for area in areas:
        nome_aba_area, criada = criar_ou_obter_aba_area(
            nome_area=area,
            id_servico=id_servico,
            unidade=unidade,
            data_servico=data_servico,
            nomes_existentes=nomes_existentes,
        )
        if criada:
            abas_areas_criadas.append(nome_aba_area)

    usuario = st.session_state.dados_usuario_logado or {}

    registro = {
        "ID_SERVICO": id_servico,
        "STATUS": "ABERTO",
        "UNIDADE": unidade,
        "DATA": dados["DATA"],
        "INICIO_SERVICO": dados["INICIO_SERVICO"],
        "TERMINO_SERVICO": dados["TERMINO_SERVICO"],
        "VIATURA": dados["VIATURA"],
        "KM_INICIAL": dados["KM_INICIAL"],
        "KM_FINAL": dados["KM_FINAL"],
        "NUM_AREAS": dados["NUM_AREAS"],
        "NOMES_AREAS_JSON": json_dumps_lista(areas),
        "ABAS_AREAS_CRIADAS_JSON": json_dumps_lista(abas_areas_criadas),
        "SUPERVISOR": dados["SUPERVISOR"],
        "MOTORISTA": dados["MOTORISTA"],
        "SEGURANCA_1": dados["SEGURANCA_1"],
        "SEGURANCA_2": dados["SEGURANCA_2"],
        "SEGURANCA_3": dados["SEGURANCA_3"],
        "OBSERVACOES_GERAIS": dados["OBSERVACOES_GERAIS"],
        "CRIADO_POR_ID": normalizar_texto(usuario.get("ID", "")),
        "CRIADO_POR_RG": normalizar_rg(usuario.get("RG", "")),
        "CRIADO_POR_NOME": normalizar_texto(usuario.get("NOME_DE_ESCALA", "")),
        "CRIADO_EM": agora_br(),
        "ATUALIZADO_EM": agora_br(),
        "VERSAO": 1,
    }

    ws = garantir_worksheet(unidade, COLUNAS_SERVICO)
    executar_com_retry(ws.append_row, registro_servico_para_linha(registro), value_input_option="RAW")
    return True, registro


def atualizar_servico(dados: dict):
    unidade = dados["UNIDADE"]
    id_servico = dados["ID_SERVICO"]
    versao_esperada = para_int(dados["VERSAO"], 1)

    numero_linha, registro_atual = localizar_linha_servico_por_id(unidade, id_servico)
    if numero_linha is None or registro_atual is None:
        return False, "⚠️ Serviço não encontrado. Ele pode ter sido excluído."

    # Serviço concluído também pode ser editado nesta tela, desde que tenha sido selecionado
    # e o usuário faça parte da supervisão. Mantemos o STATUS original ao salvar.
    status_atual_servico = normalizar_texto(registro_atual.get("STATUS", "ABERTO")).upper() or "ABERTO"

    versao_atual = para_int(registro_atual.get("VERSAO", ""), 1)
    if versao_atual != versao_esperada:
        return False, "⚠️ Este serviço foi alterado por outro usuário. Cancele e abra novamente."

    abas_antigas = json_loads_lista(registro_atual.get("ABAS_AREAS_CRIADAS_JSON", "[]"))
    excluir_abas_areas_criadas(abas_antigas)

    nomes_existentes = obter_nomes_abas_existentes_set()
    areas = dados["AREAS"]
    abas_areas_criadas = []

    for area in areas:
        nome_aba_area, criada = criar_ou_obter_aba_area(
            nome_area=area,
            id_servico=id_servico,
            unidade=unidade,
            data_servico=dados["DATA"],
            nomes_existentes=nomes_existentes,
        )
        if criada:
            abas_areas_criadas.append(nome_aba_area)

    registro_atualizado = {
        "ID_SERVICO": id_servico,
        "STATUS": status_atual_servico,
        "UNIDADE": unidade,
        "DATA": dados["DATA"],
        "INICIO_SERVICO": dados["INICIO_SERVICO"],
        "TERMINO_SERVICO": dados["TERMINO_SERVICO"],
        "VIATURA": dados["VIATURA"],
        "KM_INICIAL": dados["KM_INICIAL"],
        "KM_FINAL": dados["KM_FINAL"],
        "NUM_AREAS": dados["NUM_AREAS"],
        "NOMES_AREAS_JSON": json_dumps_lista(areas),
        "ABAS_AREAS_CRIADAS_JSON": json_dumps_lista(abas_areas_criadas),
        "SUPERVISOR": dados["SUPERVISOR"],
        "MOTORISTA": dados["MOTORISTA"],
        "SEGURANCA_1": dados["SEGURANCA_1"],
        "SEGURANCA_2": dados["SEGURANCA_2"],
        "SEGURANCA_3": dados["SEGURANCA_3"],
        "OBSERVACOES_GERAIS": dados["OBSERVACOES_GERAIS"],
        "CRIADO_POR_ID": registro_atual.get("CRIADO_POR_ID", ""),
        "CRIADO_POR_RG": registro_atual.get("CRIADO_POR_RG", ""),
        "CRIADO_POR_NOME": registro_atual.get("CRIADO_POR_NOME", ""),
        "CRIADO_EM": registro_atual.get("CRIADO_EM", ""),
        "ATUALIZADO_EM": agora_br(),
        "VERSAO": versao_atual + 1,
    }

    ws = garantir_worksheet(unidade, COLUNAS_SERVICO)
    executar_com_retry(ws.update, f"A{numero_linha}:X{numero_linha}", [registro_servico_para_linha(registro_atualizado)])
    limpar_cache_servicos()
    return True, registro_atualizado


def excluir_servico(unidade: str, id_servico: str):
    numero_linha, registro_atual = localizar_linha_servico_por_id(unidade, id_servico)
    if numero_linha is None or registro_atual is None:
        return False, "⚠️ Serviço não encontrado."

    if normalizar_texto(registro_atual.get("STATUS", "")).upper() == "CONCLUIDO":
        return False, "⚠️ Serviço concluído não pode ser excluído."

    abas_criadas = json_loads_lista(registro_atual.get("ABAS_AREAS_CRIADAS_JSON", "[]"))
    excluir_abas_areas_criadas(abas_criadas)

    ws = garantir_worksheet(unidade, COLUNAS_SERVICO)
    executar_com_retry(ws.delete_rows, numero_linha)
    limpar_cache_servicos()
    return True, "✅ Serviço excluído com sucesso."


def concluir_servico(unidade: str, id_servico: str, versao_esperada: int):
    numero_linha, registro_atual = localizar_linha_servico_por_id(unidade, id_servico)
    if numero_linha is None or registro_atual is None:
        return False, "⚠️ Serviço não encontrado."

    if normalizar_texto(registro_atual.get("STATUS", "")).upper() == "CONCLUIDO":
        return False, "⚠️ Este serviço já estava concluído."

    versao_atual = para_int(registro_atual.get("VERSAO", ""), 1)
    if versao_atual != para_int(versao_esperada, 1):
        return False, "⚠️ Serviço alterado por outro usuário. Atualize antes de concluir."

    km_final = normalizar_texto(registro_atual.get("KM_FINAL", ""))
    if not km_final or not apenas_digitos(km_final):
        return False, (
            "⚠️ O KM FINAL é obrigatório para concluir um serviço aberto. "
            "Clique em Editar, preencha o KM FINAL, salve e depois clique em Concluir Serviço."
        )

    registro_atual["KM_FINAL"] = apenas_digitos(km_final)
    registro_atual["STATUS"] = "CONCLUIDO"
    registro_atual["ATUALIZADO_EM"] = agora_br()
    registro_atual["VERSAO"] = versao_atual + 1

    ws = garantir_worksheet(unidade, COLUNAS_SERVICO)
    executar_com_retry(ws.update, f"A{numero_linha}:X{numero_linha}", [registro_servico_para_linha(registro_atual)])
    limpar_cache_servicos()
    return True, "✅ Serviço concluído com sucesso."


# ==========================================================
# CONSULTA DE SERVIÇOS SALVOS POR UNIDADE/DATA
# ==========================================================
def preparar_dataframe_servicos(df: pd.DataFrame, unidade: str) -> pd.DataFrame:
    for coluna in COLUNAS_SERVICO:
        if coluna not in df.columns:
            df[coluna] = ""

    df = df[COLUNAS_SERVICO].copy()

    if df.empty:
        return df

    for coluna in COLUNAS_SERVICO:
        df[coluna] = df[coluna].astype(str).fillna("").str.strip()

    df = df[df["ID_SERVICO"].astype(str).str.strip() != ""].copy()
    df["UNIDADE"] = df["UNIDADE"].replace("", unidade)
    df["VERSAO"] = pd.to_numeric(df["VERSAO"], errors="coerce").fillna(1).astype(int)

    # Mantém serviços abertos primeiro e os mais recentes acima quando houver empate.
    df["_STATUS_ORDEM"] = df["STATUS"].apply(lambda x: 0 if normalizar_texto(x).upper() == "ABERTO" else 1)
    df = df.sort_values(["_STATUS_ORDEM", "DATA", "CRIADO_EM"], ascending=[True, False, False]).reset_index(drop=True)
    df = df.drop(columns=["_STATUS_ORDEM"], errors="ignore")
    return df


@st.cache_data(ttl=20)
def carregar_servicos_unidade(unidade: str) -> pd.DataFrame:
    unidade = normalizar_texto(unidade)

    if not unidade:
        return pd.DataFrame(columns=COLUNAS_SERVICO)

    planilha = abrir_planilha()
    ws = encontrar_worksheet(planilha, unidade)

    if ws is None:
        return pd.DataFrame(columns=COLUNAS_SERVICO)

    registros = executar_com_retry(ws.get_all_records)
    df = pd.DataFrame(registros)
    return preparar_dataframe_servicos(df, unidade)


def limpar_cache_servicos():
    try:
        carregar_servicos_unidade.clear()
    except Exception:
        pass


def montar_opcoes_datas_servico(df_servicos: pd.DataFrame):
    """
    Retorna lista de IDs e mapa ID -> rótulo para a listbox de datas cadastradas.

    Revisão: exibe TODAS as datas já cadastradas na aba da unidade
    selecionada. Serviços concluídos continuam aparecendo identificados
    no rótulo, mas a continuidade será bloqueada somente quando o usuário
    logado não fizer parte da equipe daquela supervisão.
    """
    if df_servicos.empty:
        return [""], {"": "Nenhuma data cadastrada"}

    df_validos = df_servicos.copy()
    df_validos["ID_SERVICO"] = df_validos["ID_SERVICO"].astype(str).str.strip()
    df_validos["DATA"] = df_validos["DATA"].astype(str).str.strip()
    df_validos = df_validos[(df_validos["ID_SERVICO"] != "") & (df_validos["DATA"] != "")].copy()

    if df_validos.empty:
        return [""], {"": "Nenhuma data cadastrada"}

    opcoes = [""]
    rotulos = {"": "Selecione uma data cadastrada"}

    contagem_datas = df_validos["DATA"].astype(str).str.strip().value_counts().to_dict()

    for _, linha in df_validos.iterrows():
        id_servico = normalizar_texto(linha.get("ID_SERVICO", ""))
        data_servico = normalizar_texto(linha.get("DATA", ""))
        inicio = normalizar_texto(linha.get("INICIO_SERVICO", ""))
        supervisor = normalizar_texto(linha.get("SUPERVISOR", ""))
        status = normalizar_texto(linha.get("STATUS", "")).upper() or "ABERTO"

        if not id_servico or not data_servico:
            continue

        partes = [data_servico]

        if contagem_datas.get(data_servico, 0) > 1:
            if inicio:
                partes.append(inicio)
            if supervisor:
                partes.append(supervisor)

        if status == "CONCLUIDO":
            partes.append("CONCLUÍDO")

        rotulo = " — ".join(partes)
        opcoes.append(id_servico)
        rotulos[id_servico] = rotulo

    if len(opcoes) == 1:
        return [""], {"": "Nenhuma data cadastrada"}

    return opcoes, rotulos


def obter_servico_df_por_id(df_servicos: pd.DataFrame, id_servico: str) -> dict | None:
    id_servico = normalizar_texto(id_servico)

    if df_servicos.empty or not id_servico:
        return None

    encontrado = df_servicos[df_servicos["ID_SERVICO"].astype(str).str.strip() == id_servico]

    if encontrado.empty:
        return None

    registro = encontrado.iloc[0].to_dict()
    registro["VERSAO"] = para_int(registro.get("VERSAO", ""), 1)
    return registro


def usuario_logado_participa_servico(registro: dict) -> bool:
    usuario = st.session_state.dados_usuario_logado or {}
    rg_logado = apenas_digitos(usuario.get("RG", ""))

    if not rg_logado:
        return False

    campos_equipe = [
        "SUPERVISOR",
        "MOTORISTA",
        "SEGURANCA_1",
        "SEGURANCA_2",
        "SEGURANCA_3",
    ]

    for campo in campos_equipe:
        rg_campo = apenas_digitos(registro.get(campo, ""))

        if rg_campo and rg_logado == rg_campo:
            return True

    return False


# ==========================================================
# ESTADO DA SESSÃO
# ==========================================================
def iniciar_estado():
    padroes = {
        "indice_atual": 0,
        "modo_cadastro": "visualizar",
        "admin_logado": False,
        "usuario_logado": False,
        "dados_usuario_logado": None,
        "msg_cadastro": "",
        "tipo_msg_cadastro": "info",
        "id_em_edicao": None,
        "versao_em_edicao": None,
        "busca_rg_cadastro": "",
        "busca_rg_admin": "",
        "filtro_rg_admin": "",
        "confirmacao_pendente": None,
        "modo_servico": "visualizar",
        "msg_servico": "",
        "tipo_msg_servico": "info",
        "servico_atual": None,
        "serv_unidade": "",
        "serv_data": None,
        "serv_inicio": None,
        "serv_termino": None,
        "serv_viatura": "",
        "serv_km_inicial": "",
        "serv_km_final": "",
        "serv_num_areas": 0,
        "serv_supervisor": "",
        "serv_motorista": "",
        "serv_seguranca_1": "",
        "serv_seguranca_2": "",
        "serv_seguranca_3": "",
        "serv_observacoes": "",
        "serv_data_cadastrada": "",
        "serv_data_cadastrada_carregada": "",
        "servico_pendente_carregar": None,
        "serv_form_version": 0,
        "area_menu_ativo": False,
        "area_menu_abas": [],
        "area_menu_msg": "",
        "assuncao_ativa": False,
        "area_menu_tipo_msg": "info",
        "pagina_atual": "🔐 Login",
        "_cadastro_migrado": False,
        "_headers_ok": set(),
    }

    for chave, valor in padroes.items():
        if chave not in st.session_state:
            st.session_state[chave] = valor

    for i in range(1, 6):
        chave_area = f"serv_area_{i}"
        if chave_area not in st.session_state:
            st.session_state[chave_area] = ""


iniciar_estado()


# ==========================================================
# CONFIRMAÇÃO
# ==========================================================
def solicitar_confirmacao(acao: str, mensagem: str, dados: dict | None = None):
    st.session_state.confirmacao_pendente = {
        "acao": acao,
        "mensagem": mensagem,
        "dados": dados or {},
    }


def limpar_confirmacao():
    st.session_state.confirmacao_pendente = None


def renderizar_confirmacao(prefixos: list[str]):
    confirmacao = st.session_state.confirmacao_pendente
    if not confirmacao:
        return None, None

    acao = confirmacao.get("acao", "")
    if not any(acao.startswith(prefixo) for prefixo in prefixos):
        return None, None

    mensagem = confirmacao.get("mensagem", "")
    dados = confirmacao.get("dados", {})

    st.markdown(
        f"""
        <div class='confirm-box'>
            ⚠️ {mensagem}
            <br><br>
            Confirma a execução dessa ação?
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_conf_1, col_conf_2 = st.columns(2)

    with col_conf_1:
        confirmar = st.button(
            "✅ Confirmar",
            key=f"confirmar_{acao}",
            type="primary",
            use_container_width=True,
        )

    with col_conf_2:
        cancelar = st.button(
            "❌ Cancelar",
            key=f"cancelar_{acao}",
            use_container_width=True,
        )

    if cancelar:
        limpar_confirmacao()
        st.info("Operação cancelada.")
        st.rerun()

    if confirmar:
        limpar_confirmacao()
        return acao, dados

    return None, None


# ==========================================================
# BLOQUEIO, RESET E LIMPEZA
# ==========================================================
def cadastro_em_operacao() -> bool:
    return st.session_state.modo_cadastro in ["novo", "editar"]


def servico_em_operacao() -> bool:
    return st.session_state.modo_servico in ["novo", "editar"]


def mostrar_bloqueio(mensagem: str):
    st.markdown(
        f"""
        <div class='blocked-box'>
            🔒 {mensagem}
        </div>
        """,
        unsafe_allow_html=True,
    )


def resetar_operacao_cadastro():
    st.session_state.modo_cadastro = "visualizar"
    st.session_state.id_em_edicao = None
    st.session_state.versao_em_edicao = None


def limpar_busca_cadastro():
    st.session_state.busca_rg_cadastro = ""
    st.session_state.msg_cadastro = "🧹 Busca limpa."
    st.session_state.tipo_msg_cadastro = "info"


def limpar_busca_admin():
    st.session_state.busca_rg_admin = ""
    st.session_state.filtro_rg_admin = ""


def limpar_campos_servico():
    st.session_state.modo_servico = "visualizar"
    st.session_state.servico_atual = None
    st.session_state.area_menu_ativo = False
    st.session_state.area_menu_abas = []
    st.session_state.area_menu_msg = ""
    st.session_state.area_menu_tipo_msg = "info"
    st.session_state.serv_unidade = ""
    st.session_state.serv_data = None
    st.session_state.serv_inicio = None
    st.session_state.serv_termino = None
    st.session_state.serv_viatura = ""
    st.session_state.serv_km_inicial = ""
    st.session_state.serv_km_final = ""
    st.session_state.serv_num_areas = 0
    st.session_state.serv_supervisor = ""
    st.session_state.serv_motorista = ""
    st.session_state.serv_seguranca_1 = ""
    st.session_state.serv_seguranca_2 = ""
    st.session_state.serv_seguranca_3 = ""
    st.session_state.serv_observacoes = ""
    st.session_state.serv_data_cadastrada = ""
    st.session_state.serv_data_cadastrada_carregada = ""
    st.session_state.servico_pendente_carregar = None
    st.session_state.serv_form_version = st.session_state.get("serv_form_version", 0) + 1

    for i in range(1, 6):
        st.session_state[f"serv_area_{i}"] = ""


def limpar_dados_visuais_servico_mantendo_unidade():
    """
    Limpa os campos do serviço mostrado na tela quando o usuário troca a unidade.
    Mantém apenas a unidade selecionada para que a lista de datas seja recarregada
    conforme a nova unidade.
    """
    st.session_state.servico_atual = None
    st.session_state.serv_data = None
    st.session_state.serv_inicio = None
    st.session_state.serv_termino = None
    st.session_state.serv_viatura = ""
    st.session_state.serv_km_inicial = ""
    st.session_state.serv_km_final = ""
    st.session_state.serv_num_areas = 0
    st.session_state.serv_supervisor = ""
    st.session_state.serv_motorista = ""
    st.session_state.serv_seguranca_1 = ""
    st.session_state.serv_seguranca_2 = ""
    st.session_state.serv_seguranca_3 = ""
    st.session_state.serv_observacoes = ""
    st.session_state.serv_form_version = st.session_state.get("serv_form_version", 0) + 1

    for i in range(1, 6):
        st.session_state[f"serv_area_{i}"] = ""



def limpar_campos_servico_mantendo_registro_atual():
    """
    Limpa somente os dados exibidos na tela da Assunção, mantendo o último
    serviço salvo em memória para permitir Editar, Excluir ou Continuar Supervisão.

    Uso: depois de editar/salvar/cancelar/terminar supervisão, a tela não deve
    continuar mostrando dados antigos. Os dados só voltam a aparecer quando a
    unidade e uma data cadastrada forem selecionadas novamente ou quando Editar
    for acionado.
    """
    registro_atual = st.session_state.servico_atual

    st.session_state.modo_servico = "visualizar"
    st.session_state.servico_atual = registro_atual
    st.session_state.serv_data = None
    st.session_state.serv_inicio = None
    st.session_state.serv_termino = None
    st.session_state.serv_viatura = ""
    st.session_state.serv_km_inicial = ""
    st.session_state.serv_km_final = ""
    st.session_state.serv_num_areas = 0
    st.session_state.serv_supervisor = ""
    st.session_state.serv_motorista = ""
    st.session_state.serv_seguranca_1 = ""
    st.session_state.serv_seguranca_2 = ""
    st.session_state.serv_seguranca_3 = ""
    st.session_state.serv_observacoes = ""
    st.session_state.serv_data_cadastrada = ""
    st.session_state.serv_data_cadastrada_carregada = ""
    st.session_state.servico_pendente_carregar = None
    st.session_state.serv_form_version = st.session_state.get("serv_form_version", 0) + 1

    for i in range(1, 6):
        st.session_state[f"serv_area_{i}"] = ""

def ao_alterar_unidade_servico():
    """
    Força a atualização da tela quando a unidade é alterada.

    Ao trocar a unidade, a data anteriormente selecionada deixa de valer, o serviço
    carregado na tela é limpo e o cache de serviços é liberado para permitir leitura
    atualizada da aba correspondente à nova unidade.
    """
    st.session_state.serv_data_cadastrada = ""
    st.session_state.serv_data_cadastrada_carregada = ""
    st.session_state.servico_pendente_carregar = None

    if st.session_state.modo_servico == "visualizar":
        limpar_dados_visuais_servico_mantendo_unidade()

    limpar_cache_servicos()

    unidade = normalizar_texto(st.session_state.serv_unidade)
    if unidade:
        st.session_state.msg_servico = f"ℹ️ Unidade selecionada: {unidade}. A lista de datas foi atualizada."
        st.session_state.tipo_msg_servico = "info"


def ao_alterar_data_cadastrada_servico():
    """
    Prepara o carregamento dos dados quando uma data cadastrada é selecionada.
    O carregamento efetivo ocorre no início do render da Assunção, antes da criação
    dos widgets, evitando erro de alteração de st.session_state após instanciar campos.
    """
    st.session_state.serv_data_cadastrada_carregada = ""


def carregar_servico_na_tela(registro: dict, modo_destino: str = "visualizar"):
    """
    Carrega na tela todos os dados do serviço salvo.

    Esta função é essencial para o botão Editar: quando a supervisão é
    encerrada pelo menu TERMINAR SUPERVISÃO e o usuário volta para
    ASSUNÇÃO DO SERVIÇO, o último serviço salvo permanece em memória.
    Ao clicar em Editar, todos os campos voltam preenchidos com os dados
    desse serviço para alteração.
    """
    st.session_state.servico_atual = registro
    st.session_state.modo_servico = modo_destino
    st.session_state.serv_unidade = normalizar_texto(registro.get("UNIDADE", ""))
    st.session_state.serv_data = texto_para_data(registro.get("DATA", ""))
    st.session_state.serv_inicio = texto_para_hora(registro.get("INICIO_SERVICO", ""))
    st.session_state.serv_termino = texto_para_hora(registro.get("TERMINO_SERVICO", ""))
    st.session_state.serv_viatura = normalizar_texto(registro.get("VIATURA", ""))
    st.session_state.serv_km_inicial = normalizar_texto(registro.get("KM_INICIAL", ""))
    st.session_state.serv_km_final = normalizar_texto(registro.get("KM_FINAL", ""))
    st.session_state.serv_num_areas = para_int(registro.get("NUM_AREAS", ""), 0)
    st.session_state.serv_supervisor = normalizar_texto(registro.get("SUPERVISOR", ""))
    st.session_state.serv_motorista = normalizar_texto(registro.get("MOTORISTA", ""))
    st.session_state.serv_seguranca_1 = normalizar_texto(registro.get("SEGURANCA_1", ""))
    st.session_state.serv_seguranca_2 = normalizar_texto(registro.get("SEGURANCA_2", ""))
    st.session_state.serv_seguranca_3 = normalizar_texto(registro.get("SEGURANCA_3", ""))
    st.session_state.serv_observacoes = normalizar_texto(registro.get("OBSERVACOES_GERAIS", ""))

    areas = json_loads_lista(registro.get("NOMES_AREAS_JSON", "[]"))
    for i in range(1, 6):
        st.session_state[f"serv_area_{i}"] = areas[i - 1] if i <= len(areas) else ""

    st.session_state.serv_form_version = st.session_state.get("serv_form_version", 0) + 1


def ativar_menu_areas_do_servico(registro: dict):
    """
    Após salvar a assunção, mostra no menu apenas as abas das áreas e a aba TERMINAR SUPERVISÃO.
    As demais abas ficam ocultas até o usuário clicar em TERMINAR SUPERVISÃO.
    """
    areas = json_loads_lista(registro.get("NOMES_AREAS_JSON", "[]"))
    areas = [texto_caixa_alta(area) for area in areas if normalizar_texto(area)]

    # Garante abas únicas preservando a ordem digitada pelo usuário.
    abas = []
    for area in areas:
        if area not in abas:
            abas.append(area)

    st.session_state.area_menu_ativo = True
    st.session_state.area_menu_abas = abas
    st.session_state.assuncao_ativa = True

    if abas:
        st.session_state.pagina_atual = abas[0]
    else:
        st.session_state.pagina_atual = "TERMINAR SUPERVISÃO"


def desativar_menu_areas():
    st.session_state.area_menu_ativo = False
    st.session_state.area_menu_abas = []


def encerrar_assuncao_sem_logout():
    """
    Encerra a tela/fluxo da Assunção do Serviço sem deslogar o usuário.
    Isso faz a aba Login voltar a aparecer no menu e deixa o foco no Login.
    """
    limpar_campos_servico()
    st.session_state.assuncao_ativa = False
    st.session_state.area_menu_ativo = False
    st.session_state.area_menu_abas = []
    st.session_state.pagina_atual = "🔐 Login"


# ==========================================================
# VALIDAÇÃO DO SERVIÇO
# ==========================================================
def montar_opcoes_usuarios(df_usuarios: pd.DataFrame) -> list:
    opcoes = [""]

    if df_usuarios.empty:
        return opcoes

    df_temp = df_usuarios.copy()
    df_temp["_RG_NUM"] = df_temp["RG"].apply(lambda x: para_int(apenas_digitos(x), 999999999))
    df_temp = df_temp.sort_values(["_RG_NUM", "RG", "NOME_DE_ESCALA"]).reset_index(drop=True)

    for _, linha in df_temp.iterrows():
        rg = normalizar_rg(linha.get("RG", ""))
        nome = normalizar_texto(linha.get("NOME_DE_ESCALA", ""))
        posto = normalizar_texto(linha.get("POSTO_GRADUACAO", ""))
        rotulo = f"{rg} - {posto} - {nome}".strip(" -")
        opcoes.append(rotulo)

    return opcoes


def coletar_dados_servico():
    unidade = normalizar_texto(st.session_state.serv_unidade)
    data_servico = data_para_texto(st.session_state.serv_data)
    inicio = hora_para_texto(st.session_state.serv_inicio)
    termino = hora_para_texto(st.session_state.serv_termino)
    viatura = texto_caixa_alta(st.session_state.serv_viatura)
    km_inicial = normalizar_texto(st.session_state.serv_km_inicial)
    km_final = normalizar_texto(st.session_state.serv_km_final)
    num_areas = para_int(st.session_state.serv_num_areas, 0)

    supervisor = normalizar_texto(st.session_state.serv_supervisor)
    motorista = normalizar_texto(st.session_state.serv_motorista)
    seguranca_1 = normalizar_texto(st.session_state.serv_seguranca_1)
    seguranca_2 = normalizar_texto(st.session_state.serv_seguranca_2)
    seguranca_3 = normalizar_texto(st.session_state.serv_seguranca_3)
    observacoes = texto_caixa_alta(st.session_state.serv_observacoes)

    erros = []

    if not unidade:
        erros.append("Selecione a unidade.")
    if not data_servico:
        erros.append("Selecione a data.")
    if not inicio:
        erros.append("Informe o início do serviço.")
    if not termino:
        erros.append("Informe o término do serviço.")
    if not viatura:
        erros.append("Informe a viatura.")
    if not km_inicial or not apenas_digitos(km_inicial):
        erros.append("Informe o KM inicial usando somente números.")
    # KM FINAL não é obrigatório no salvamento inicial/edição comum.
    # Ele só será exigido no momento de CONCLUIR um serviço ABERTO.
    if km_final and not apenas_digitos(km_final):
        erros.append("Informe o KM final usando somente números, ou deixe em branco até a conclusão.")
    if num_areas < 1 or num_areas > 5:
        erros.append("Informe o número de áreas supervisionadas entre 1 e 5.")
    # Na equipe de serviço, somente o Supervisor é obrigatório.
    # Motorista e Seguranças podem ficar em branco.
    if not supervisor:
        erros.append("Selecione o Supervisor.")

    areas = []
    for i in range(1, num_areas + 1):
        nome_area = texto_caixa_alta(st.session_state.get(f"serv_area_{i}", ""))
        if not nome_area:
            erros.append(f"Informe o nome da área {i}.")
        else:
            areas.append(nome_area)

    if len(set(areas)) != len(areas):
        erros.append("Não repita nomes de áreas.")

    nomes_reservados_menu = {"TERMINAR SUPERVISÃO", "TERMINAR SUPERVISAO"}
    if any(area in nomes_reservados_menu for area in areas):
        erros.append("O nome TERMINAR SUPERVISÃO é reservado para a aba do menu. Use outro nome para a área.")

    dados = {
        "UNIDADE": unidade,
        "DATA": data_servico,
        "INICIO_SERVICO": inicio,
        "TERMINO_SERVICO": termino,
        "VIATURA": viatura,
        "KM_INICIAL": apenas_digitos(km_inicial),
        "KM_FINAL": apenas_digitos(km_final),
        "NUM_AREAS": num_areas,
        "AREAS": areas,
        "SUPERVISOR": supervisor,
        "MOTORISTA": motorista,
        "SEGURANCA_1": seguranca_1,
        "SEGURANCA_2": seguranca_2,
        "SEGURANCA_3": seguranca_3,
        "OBSERVACOES_GERAIS": observacoes,
    }

    return erros, dados


# ==========================================================
# CABEÇALHO
# ==========================================================
st.markdown(f"<div class='main-title'>📋 {NOME_APP}</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='sub-title'>Sistema inicial de cadastro, autorização, acesso e assunção do serviço</div>",
    unsafe_allow_html=True,
)


# ==========================================================
# ABAS VISÍVEIS
# ==========================================================
def obter_abas_visiveis():
    """
    Controla exatamente quais abas aparecem no menu superior.

    Regras atuais:
    - Sem login: Login, Cadastro e Administrador.
    - Admin logado: somente Administrador.
    - Cadastro em Novo/Editar: somente Cadastro.
    - Usuário logado e sem Assunção ativa: Login e ASSUNÇÃO DO SERVIÇO.
    - Ao clicar/entrar em ASSUNÇÃO DO SERVIÇO: Login fica oculto e aparece somente ASSUNÇÃO DO SERVIÇO.
    - Após salvar a assunção: somente as áreas digitadas + TERMINAR SUPERVISÃO.
    - Ao clicar em TERMINAR SUPERVISÃO: volta somente para ASSUNÇÃO DO SERVIÇO.
    - Ao clicar em Concluir Serviço na Assunção: cancela/conclui o fluxo da assunção e o Login volta a aparecer com foco.
    """
    if st.session_state.admin_logado:
        return ["👨‍💼 Administrador"]

    if cadastro_em_operacao():
        return ["📝 Cadastro"]

    if st.session_state.usuario_logado:
        if st.session_state.area_menu_ativo:
            abas_unicas = []

            for nome_area in st.session_state.area_menu_abas or []:
                nome_area = texto_caixa_alta(nome_area)
                if nome_area and nome_area not in {"TERMINAR SUPERVISÃO", "TERMINAR SUPERVISAO"} and nome_area not in abas_unicas:
                    abas_unicas.append(nome_area)

            st.session_state.area_menu_abas = abas_unicas

            if abas_unicas:
                return abas_unicas + ["TERMINAR SUPERVISÃO"]

            return ["TERMINAR SUPERVISÃO"]

        if st.session_state.assuncao_ativa:
            return ["🕘 ASSUNÇÃO DO SERVIÇO"]

        return ["🔐 Login", "🕘 ASSUNÇÃO DO SERVIÇO"]

    return ["🔐 Login", "📝 Cadastro", "👨‍💼 Administrador"]


def renderizar_menu_superior(abas_visiveis: list[str]) -> str:
    """
    Menu superior controlado por estado.
    Diferente de st.tabs, ele permite ocultar/mostrar opções e definir foco programaticamente.

    O widget do menu usa uma chave própria, diferente de pagina_atual, para evitar erro do
    Streamlit ao alterar session_state depois que o widget já foi instanciado.
    """
    if not abas_visiveis:
        return ""

    if st.session_state.get("pagina_atual") not in abas_visiveis:
        st.session_state.pagina_atual = abas_visiveis[0]

    indice_atual = abas_visiveis.index(st.session_state.pagina_atual)
    chave_menu = "menu_widget_" + str(abs(hash("|".join(abas_visiveis))))

    pagina_escolhida = st.radio(
        "Menu",
        options=abas_visiveis,
        index=indice_atual,
        key=chave_menu,
        horizontal=True,
        label_visibility="collapsed",
    )

    if pagina_escolhida != st.session_state.pagina_atual:
        st.session_state.pagina_atual = pagina_escolhida

    if (
        pagina_escolhida == "🕘 ASSUNÇÃO DO SERVIÇO"
        and st.session_state.usuario_logado
        and not st.session_state.area_menu_ativo
        and not st.session_state.assuncao_ativa
    ):
        st.session_state.assuncao_ativa = True
        st.rerun()

    return pagina_escolhida


abas_visiveis = obter_abas_visiveis()
pagina_selecionada = renderizar_menu_superior(abas_visiveis)


# ==========================================================
# RENDERIZAÇÃO: LOGIN
# ==========================================================
def render_login():
    st.subheader("🔐 Acesso ao App")

    acao_confirmada, dados_confirmados = renderizar_confirmacao(["login_"])

    if acao_confirmada == "login_entrar":
        limpar_cache()
        df_login = carregar_cadastros()
        rg_login = normalizar_rg(dados_confirmados.get("rg", ""))
        senha_login = normalizar_texto(dados_confirmados.get("senha", ""))

        usuario = df_login[
            (df_login["RG"].astype(str).str.strip() == rg_login)
            &
            (df_login["SENHA"].astype(str).str.strip() == senha_login)
        ]

        if usuario.empty:
            st.error("❌ RG ou Senha inválidos.")
        else:
            dados = usuario.iloc[0].to_dict()
            autorizado = normalizar_texto(dados.get("AUTORIZADO", "")).upper()
            if autorizado != "SIM":
                st.warning("⚠️ Usuário cadastrado, mas ainda não autorizado pelo Administrador.")
            else:
                st.session_state.usuario_logado = True
                st.session_state.dados_usuario_logado = dados
                st.success(f"✅ Acesso liberado. Bem-vindo, {dados.get('NOME_DE_ESCALA', '')}!")
                st.rerun()

    if acao_confirmada == "login_sair":
        st.session_state.usuario_logado = False
        st.session_state.dados_usuario_logado = None
        limpar_campos_servico()
        st.rerun()

    if st.session_state.admin_logado:
        mostrar_bloqueio("A aba Login está bloqueada enquanto o Administrador estiver logado.<br>Para voltar ao uso normal, saia da aba Administrador.")
        return

    if cadastro_em_operacao():
        mostrar_bloqueio("A aba Login está bloqueada porque há uma operação de Cadastro em andamento.<br>Finalize com Salvar ou abandone com Cancelar na aba Cadastro.")
        return

    df_login = carregar_cadastros()
    validar_usuario_logado(df_login)

    if not st.session_state.usuario_logado:
        with st.form("form_login"):
            rg_login = st.text_input("RG")
            senha_login = st.text_input("Senha", type="password")
            st.markdown("")
            entrar = st.form_submit_button("🔓 Entrar", type="primary", use_container_width=True)

        if entrar:
            rg_login = normalizar_rg(rg_login)
            senha_login = normalizar_texto(senha_login)
            if not rg_login or not senha_login:
                st.warning("⚠️ Informe RG e Senha.")
            else:
                solicitar_confirmacao("login_entrar", "Você está prestes a realizar login no sistema.", {"rg": rg_login, "senha": senha_login})
                st.rerun()

    if st.session_state.usuario_logado and st.session_state.dados_usuario_logado:
        dados = st.session_state.dados_usuario_logado
        st.markdown("---")
        st.success("✅ Usuário logado no sistema.")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.info(f"**Posto/Graduação:** {dados.get('POSTO_GRADUACAO', '')}")
        with col2:
            st.info(f"**RG:** {dados.get('RG', '')}")
        with col3:
            st.info(f"**Nome de Escala:** {dados.get('NOME_DE_ESCALA', '')}")

        st.markdown("")
        if st.button("🚪 Sair do Login", use_container_width=True):
            solicitar_confirmacao("login_sair", "Você está prestes a sair do login.", {})
            st.rerun()


# ==========================================================
# RENDERIZAÇÃO: CADASTRO
# ==========================================================
def render_cadastro():
    st.subheader("📝 Cadastro de Usuários")
    acao_confirmada, dados_confirmados = renderizar_confirmacao(["cadastro_"])

    if acao_confirmada == "cadastro_novo":
        st.session_state.modo_cadastro = "novo"
        st.session_state.id_em_edicao = None
        st.session_state.versao_em_edicao = None
        st.session_state.msg_cadastro = "🆕 Campos liberados para novo cadastro."
        st.session_state.tipo_msg_cadastro = "info"
        st.rerun()

    if acao_confirmada == "cadastro_cancelar":
        resetar_operacao_cadastro()
        st.session_state.msg_cadastro = "↩️ Operação cancelada."
        st.session_state.tipo_msg_cadastro = "info"
        st.rerun()

    if acao_confirmada == "cadastro_editar":
        st.session_state.modo_cadastro = "editar"
        st.session_state.id_em_edicao = dados_confirmados.get("id", "")
        st.session_state.versao_em_edicao = dados_confirmados.get("versao", 1)
        st.session_state.msg_cadastro = "✏️ Campos liberados para edição."
        st.session_state.tipo_msg_cadastro = "success"
        st.rerun()

    if acao_confirmada == "cadastro_salvar":
        modo_atual = dados_confirmados.get("modo", "visualizar")
        posto_graduacao = normalizar_texto(dados_confirmados.get("posto", ""))
        rg = normalizar_rg(dados_confirmados.get("rg", ""))
        nome_escala = normalizar_texto(dados_confirmados.get("nome", ""))
        senha = normalizar_texto(dados_confirmados.get("senha", ""))

        if modo_atual == "novo":
            sucesso, resultado = adicionar_usuario(posto=posto_graduacao, rg=rg, nome=nome_escala, senha=senha)
            if not sucesso:
                st.session_state.msg_cadastro = resultado
                st.session_state.tipo_msg_cadastro = "error"
                st.rerun()

            limpar_cache()
            df_pos = carregar_cadastros()
            novo_id = resultado
            indices = df_pos.index[df_pos["ID"].astype(str).str.strip() == novo_id].tolist()
            st.session_state.indice_atual = int(indices[0]) if indices else max(0, len(df_pos) - 1)
            resetar_operacao_cadastro()
            st.session_state.msg_cadastro = "✅ Cadastro salvo com sucesso. Aguarde autorização do Administrador."
            st.session_state.tipo_msg_cadastro = "success"
            st.rerun()

        if modo_atual == "editar":
            id_edicao = normalizar_texto(dados_confirmados.get("id", ""))
            versao_edicao = para_int(dados_confirmados.get("versao", 1), 1)
            sucesso, mensagem = atualizar_usuario_por_id(id_edicao, versao_edicao, posto_graduacao, rg, nome_escala, senha)
            if not sucesso:
                st.session_state.msg_cadastro = mensagem
                st.session_state.tipo_msg_cadastro = "warning"
                st.rerun()

            limpar_cache()
            df_pos = carregar_cadastros()
            indices = df_pos.index[df_pos["ID"].astype(str).str.strip() == id_edicao].tolist()
            if indices:
                st.session_state.indice_atual = int(indices[0])
            resetar_operacao_cadastro()
            st.session_state.msg_cadastro = mensagem
            st.session_state.tipo_msg_cadastro = "success"
            st.rerun()

    if acao_confirmada == "cadastro_excluir":
        id_excluido = dados_confirmados.get("id", "")
        senha_digitada = dados_confirmados.get("senha", "")
        sucesso, mensagem = excluir_usuario_por_id(id_excluido, senha_digitada)
        if not sucesso:
            st.session_state.msg_cadastro = mensagem
            st.session_state.tipo_msg_cadastro = "error"
            st.rerun()

        limpar_cache()
        df_pos = carregar_cadastros()
        if st.session_state.indice_atual >= len(df_pos):
            st.session_state.indice_atual = max(0, len(df_pos) - 1)
        resetar_operacao_cadastro()
        st.session_state.msg_cadastro = mensagem
        st.session_state.tipo_msg_cadastro = "success"
        st.rerun()

    df = carregar_cadastros()
    if df.empty:
        st.info("Nenhum usuário cadastrado ainda.")
        st.session_state.indice_atual = 0
    else:
        st.session_state.indice_atual = max(0, min(st.session_state.indice_atual, len(df) - 1))

    modo = st.session_state.modo_cadastro
    registro_atual = {
        "ID": "", "ORDEM": "", "POSTO_GRADUACAO": "", "RG": "",
        "NOME_DE_ESCALA": "", "SENHA": "", "AUTORIZADO": "NÃO",
        "ATUALIZADO_EM": "", "VERSAO": 1,
    }

    if modo != "novo" and not df.empty:
        registro_atual = df.iloc[st.session_state.indice_atual].to_dict()

    campos_desabilitados = modo == "visualizar"

    if modo == "visualizar":
        st.markdown("<div class='status-box'>Modo atual: VISUALIZAÇÃO — os campos ficam bloqueados. Digite a senha apenas para Editar ou Excluir o cadastro selecionado.</div>", unsafe_allow_html=True)
    elif modo == "novo":
        st.markdown("<div class='status-box'>Modo atual: NOVO CADASTRO — as demais abas estão ocultas. Preencha todos os campos e clique em Salvar, ou clique em Cancelar.</div>", unsafe_allow_html=True)
    elif modo == "editar":
        st.markdown("<div class='status-box'>Modo atual: EDIÇÃO — as demais abas estão ocultas. Altere os dados desejados e clique em Salvar, ou clique em Cancelar. Se deixar a Senha em branco, a senha antiga será mantida.</div>", unsafe_allow_html=True)

    if st.session_state.msg_cadastro:
        if st.session_state.tipo_msg_cadastro == "success":
            st.success(st.session_state.msg_cadastro)
        elif st.session_state.tipo_msg_cadastro == "error":
            st.error(st.session_state.msg_cadastro)
        elif st.session_state.tipo_msg_cadastro == "warning":
            st.warning(st.session_state.msg_cadastro)
        else:
            st.info(st.session_state.msg_cadastro)

    with st.container(border=True):
        st.markdown("### 🔎 Localizar usuário por RG")
        st.markdown("<div class='small-muted'>Digite o RG e clique em Localizar para carregar o cadastro correspondente.</div>", unsafe_allow_html=True)
        col_busca1, col_busca2, col_busca3 = st.columns([5, 1.15, 1.15])
        with col_busca1:
            st.text_input("RG para localização", key="busca_rg_cadastro", disabled=(df.empty or modo != "visualizar"), label_visibility="collapsed", placeholder="Digite o RG para localizar")
        with col_busca2:
            botao_localizar_cadastro = st.button("🔎 Localizar", disabled=(df.empty or modo != "visualizar"), key="btn_localizar_cadastro", use_container_width=True, type="primary")
        with col_busca3:
            st.button("🧹 Limpar", disabled=(modo != "visualizar"), key="btn_limpar_busca_cadastro", on_click=limpar_busca_cadastro, use_container_width=True)

    if botao_localizar_cadastro:
        termo_busca = normalizar_rg(st.session_state.busca_rg_cadastro)
        if not termo_busca:
            st.session_state.msg_cadastro = "⚠️ Digite um RG para localizar."
            st.session_state.tipo_msg_cadastro = "warning"
            st.rerun()

        indices_encontrados = localizar_indices_por_rg(df, termo_busca)
        if not indices_encontrados:
            st.session_state.msg_cadastro = f"❌ Nenhum usuário encontrado com o RG: {termo_busca}"
            st.session_state.tipo_msg_cadastro = "error"
            st.rerun()

        st.session_state.indice_atual = int(indices_encontrados[0])
        st.session_state.msg_cadastro = f"✅ Usuário localizado pelo RG: {termo_busca}"
        st.session_state.tipo_msg_cadastro = "success"
        st.rerun()

    total = len(df)
    with st.container(border=True):
        col_nav1, col_nav2, col_nav3 = st.columns([1.2, 2, 1.2])
        with col_nav1:
            voltar = st.button("⬅️ Anterior", disabled=(df.empty or modo != "visualizar"), use_container_width=True)
        with col_nav2:
            st.markdown("### Registro 0 de 0" if df.empty else f"### Registro {st.session_state.indice_atual + 1} de {total}")
        with col_nav3:
            avancar = st.button("Próximo ➡️", disabled=(df.empty or modo != "visualizar"), use_container_width=True)

    if voltar:
        st.session_state.indice_atual = max(0, st.session_state.indice_atual - 1)
        st.session_state.msg_cadastro = ""
        st.rerun()

    if avancar:
        st.session_state.indice_atual = min(len(df) - 1, st.session_state.indice_atual + 1)
        st.session_state.msg_cadastro = ""
        st.rerun()

    st.markdown("---")
    with st.form("form_cadastro"):
        st.markdown("### 📄 Dados do Cadastro")
        col_a, col_b = st.columns(2)
        with col_a:
            posto_atual = normalizar_texto(registro_atual.get("POSTO_GRADUACAO", "")).upper()
            opcoes_posto = montar_opcoes_posto_graduacao(posto_atual)
            posto_graduacao = st.selectbox(
                "Posto/Graduação",
                options=opcoes_posto,
                index=indice_opcao(opcoes_posto, posto_atual),
                disabled=campos_desabilitados,
            )
            rg = st.text_input("RG", value=normalizar_texto(registro_atual.get("RG", "")), disabled=campos_desabilitados)
        with col_b:
            nome_escala = st.text_input("Nome de Escala", value=normalizar_texto(registro_atual.get("NOME_DE_ESCALA", "")), disabled=campos_desabilitados)
            if modo == "visualizar":
                senha = st.text_input("Senha", value="", type="password", help="Digite a senha para Editar ou Excluir o cadastro selecionado.")
            elif modo == "editar":
                senha = st.text_input("Senha", value="", type="password", help="Digite uma nova senha apenas se quiser alterar. Se deixar em branco, mantém a senha antiga.")
            else:
                senha = st.text_input("Senha", value="", type="password")

        st.markdown("---")
        st.markdown("### ⚙️ Comandos")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            botao_novo = st.form_submit_button("🆕 Novo", use_container_width=True)
        with col2:
            botao_editar = st.form_submit_button("✏️ Editar", use_container_width=True)
        with col3:
            botao_salvar = st.form_submit_button("💾 Salvar", type="primary", use_container_width=True)
        with col4:
            botao_cancelar = st.form_submit_button("↩️ Cancelar", use_container_width=True)
        with col5:
            botao_excluir = st.form_submit_button("🗑️ Excluir", use_container_width=True)

    if botao_novo:
        solicitar_confirmacao("cadastro_novo", "Você está prestes a iniciar um novo cadastro.", {})
        st.rerun()

    if botao_cancelar:
        solicitar_confirmacao("cadastro_cancelar", "Você está prestes a cancelar a operação atual.", {})
        st.rerun()

    if botao_editar:
        if df.empty:
            st.session_state.msg_cadastro = "⚠️ Não há cadastro para editar."
            st.session_state.tipo_msg_cadastro = "warning"
            st.rerun()

        senha_digitada = normalizar_texto(senha)
        senha_salva = normalizar_texto(registro_atual.get("SENHA", ""))
        rg_digitado = normalizar_rg(rg)
        rg_salvo = normalizar_rg(registro_atual.get("RG", ""))

        if not rg_digitado or rg_digitado != rg_salvo:
            st.session_state.msg_cadastro = "⚠️ Para editar, o RG do cadastro selecionado precisa estar correto."
            st.session_state.tipo_msg_cadastro = "warning"
            st.rerun()

        if not senha_digitada:
            st.session_state.msg_cadastro = "⚠️ Para editar, informe a senha do cadastro selecionado."
            st.session_state.tipo_msg_cadastro = "warning"
            st.rerun()

        if senha_digitada != senha_salva:
            st.session_state.msg_cadastro = "❌ Senha incorreta. Edição não autorizada."
            st.session_state.tipo_msg_cadastro = "error"
            st.rerun()

        solicitar_confirmacao(
            "cadastro_editar",
            "Você está prestes a liberar este cadastro para edição.",
            {"id": normalizar_texto(registro_atual.get("ID", "")), "versao": para_int(registro_atual.get("VERSAO", ""), 1)},
        )
        st.rerun()

    if botao_salvar:
        modo_atual = st.session_state.modo_cadastro
        posto_graduacao = normalizar_texto(posto_graduacao)
        rg = normalizar_rg(rg)
        nome_escala = normalizar_texto(nome_escala)
        senha = normalizar_texto(senha)

        if modo_atual == "visualizar":
            st.session_state.msg_cadastro = "⚠️ Clique em Novo ou Editar antes de salvar."
            st.session_state.tipo_msg_cadastro = "warning"
            st.rerun()

        if not posto_graduacao or not rg or not nome_escala:
            st.session_state.msg_cadastro = "⚠️ Preencha Posto/Graduação, RG e Nome de Escala."
            st.session_state.tipo_msg_cadastro = "warning"
            st.rerun()

        if modo_atual == "novo" and not senha:
            st.session_state.msg_cadastro = "⚠️ Informe uma senha para o novo cadastro."
            st.session_state.tipo_msg_cadastro = "warning"
            st.rerun()

        solicitar_confirmacao(
            "cadastro_salvar",
            "Você está prestes a salvar este cadastro.",
            {"modo": modo_atual, "id": normalizar_texto(st.session_state.id_em_edicao), "versao": para_int(st.session_state.versao_em_edicao, 1), "posto": posto_graduacao, "rg": rg, "nome": nome_escala, "senha": senha},
        )
        st.rerun()

    if botao_excluir:
        if df.empty:
            st.session_state.msg_cadastro = "⚠️ Não há cadastro para excluir."
            st.session_state.tipo_msg_cadastro = "warning"
            st.rerun()

        if st.session_state.modo_cadastro != "visualizar":
            st.session_state.msg_cadastro = "⚠️ A exclusão só pode ser feita no modo Visualização."
            st.session_state.tipo_msg_cadastro = "warning"
            st.rerun()

        senha_digitada = normalizar_texto(senha)
        if not senha_digitada:
            st.session_state.msg_cadastro = "⚠️ Para excluir, informe a senha do cadastro selecionado."
            st.session_state.tipo_msg_cadastro = "warning"
            st.rerun()

        solicitar_confirmacao(
            "cadastro_excluir",
            "Você está prestes a excluir este cadastro. Essa ação não poderá ser desfeita.",
            {"id": normalizar_texto(registro_atual.get("ID", "")), "senha": senha_digitada},
        )
        st.rerun()


# ==========================================================
# RENDERIZAÇÃO: ADMIN
# ==========================================================
def render_admin():
    st.subheader("👨‍💼 Acesso do Administrador")
    acao_confirmada, dados_confirmados = renderizar_confirmacao(["admin_"])

    if acao_confirmada == "admin_entrar":
        usuario_admin = dados_confirmados.get("usuario", "")
        senha_admin = dados_confirmados.get("senha", "")
        if usuario_admin == ADMIN_USUARIO and senha_admin == ADMIN_SENHA:
            st.session_state.admin_logado = True
            st.session_state.usuario_logado = False
            st.session_state.dados_usuario_logado = None
            resetar_operacao_cadastro()
            st.success("✅ Administrador logado com sucesso.")
            st.rerun()
        else:
            st.error("❌ Usuário ou senha do Administrador inválidos.")

    if acao_confirmada == "admin_sair":
        st.session_state.admin_logado = False
        resetar_operacao_cadastro()
        st.rerun()

    if acao_confirmada == "admin_salvar_autorizacoes":
        linhas_editadas = dados_confirmados.get("linhas", [])
        df_admin = carregar_cadastros()
        alterados = 0
        conflitos = 0

        for linha_editada in linhas_editadas:
            id_usuario = normalizar_texto(linha_editada.get("ID", ""))
            if not id_usuario:
                continue
            linha_original_df = df_admin[df_admin["ID"].astype(str).str.strip() == id_usuario]
            if linha_original_df.empty:
                conflitos += 1
                continue

            linha_original = linha_original_df.iloc[0]
            autorizado_original = normalizar_texto(linha_original.get("AUTORIZADO", "")).upper() == "SIM"
            autorizado_novo = bool(linha_editada.get("LIBERAR_ACESSO", False))
            if autorizado_original == autorizado_novo:
                continue

            versao_esperada = para_int(linha_original.get("VERSAO", ""), 1)
            sucesso, _ = atualizar_autorizacao_por_id(id_usuario, autorizado_novo, versao_esperada)
            if sucesso:
                alterados += 1
            else:
                conflitos += 1

        limpar_cache()
        if alterados > 0 and conflitos == 0:
            st.success(f"✅ Autorizações salvas com sucesso. Alterações realizadas: {alterados}.")
        elif alterados > 0 and conflitos > 0:
            st.warning(f"⚠️ Algumas autorizações foram salvas, mas {conflitos} cadastro(s) foram alterados por outro usuário antes do salvamento.")
        elif alterados == 0 and conflitos == 0:
            st.info("ℹ️ Nenhuma autorização foi alterada.")
        else:
            st.error("❌ Não foi possível salvar algumas autorizações.")
        st.rerun()

    if st.session_state.usuario_logado:
        mostrar_bloqueio("A aba Administrador está bloqueada enquanto houver um usuário logado.<br>Para acessar o Administrador, primeiro saia do Login.")
        return

    if cadastro_em_operacao():
        mostrar_bloqueio("A aba Administrador está bloqueada porque há uma operação de Cadastro em andamento.<br>Finalize com Salvar ou abandone com Cancelar na aba Cadastro.")
        return

    if not st.session_state.admin_logado:
        with st.form("form_admin_login"):
            usuario_admin = st.text_input("Usuário do Administrador")
            senha_admin = st.text_input("Senha do Administrador", type="password")
            st.markdown("")
            entrar_admin = st.form_submit_button("🔐 Entrar como Administrador", type="primary", use_container_width=True)

        if entrar_admin:
            solicitar_confirmacao("admin_entrar", "Você está prestes a entrar como Administrador.", {"usuario": usuario_admin, "senha": senha_admin})
            st.rerun()
        return

    st.success("✅ Administrador logado. As demais abas estão ocultas até sair do Administrador.")
    if st.button("🚪 Sair do Administrador", use_container_width=True):
        solicitar_confirmacao("admin_sair", "Você está prestes a sair do Administrador.", {})
        st.rerun()

    st.markdown("---")
    df_admin = carregar_cadastros()
    if df_admin.empty:
        st.info("Nenhum usuário cadastrado.")
        return

    with st.container(border=True):
        st.markdown("### 🔎 Localizar usuário por RG")
        st.markdown("<div class='small-muted'>Digite o RG para filtrar a tabela de autorização.</div>", unsafe_allow_html=True)
        col_admin_busca1, col_admin_busca2, col_admin_busca3 = st.columns([5, 1.15, 1.15])
        with col_admin_busca1:
            st.text_input("RG para localização na tabela", key="busca_rg_admin", label_visibility="collapsed", placeholder="Digite o RG para localizar na tabela")
        with col_admin_busca2:
            botao_localizar_admin = st.button("🔎 Localizar", key="btn_localizar_admin", use_container_width=True, type="primary")
        with col_admin_busca3:
            st.button("🧹 Limpar", key="btn_limpar_admin", on_click=limpar_busca_admin, use_container_width=True)

    if botao_localizar_admin:
        termo_admin = normalizar_rg(st.session_state.busca_rg_admin)
        if not termo_admin:
            st.warning("⚠️ Digite um RG para localizar.")
        else:
            st.session_state.filtro_rg_admin = termo_admin
            st.rerun()

    filtro_rg_admin = normalizar_rg(st.session_state.filtro_rg_admin)
    if filtro_rg_admin:
        indices_admin = localizar_indices_por_rg(df_admin, filtro_rg_admin)
        if indices_admin:
            df_admin_visivel = df_admin.loc[indices_admin].copy().reset_index(drop=True)
            st.success(f"✅ Resultado da busca por RG: {filtro_rg_admin} — {len(df_admin_visivel)} registro(s) encontrado(s).")
        else:
            df_admin_visivel = df_admin.iloc[0:0].copy()
            st.error(f"❌ Nenhum usuário encontrado com o RG: {filtro_rg_admin}")
    else:
        df_admin_visivel = df_admin.copy()

    st.markdown("### Validação de acesso dos usuários")
    if df_admin_visivel.empty:
        st.info("Nenhum usuário para exibir com o filtro atual.")
        return

    df_editor = df_admin_visivel.copy()
    df_editor["LIBERAR_ACESSO"] = df_editor["AUTORIZADO"].astype(str).str.strip().str.upper().eq("SIM")
    colunas_editor = ["ID", "ORDEM", "POSTO_GRADUACAO", "RG", "NOME_DE_ESCALA", "SENHA", "AUTORIZADO", "ATUALIZADO_EM", "VERSAO", "LIBERAR_ACESSO"]
    df_editor = df_editor[colunas_editor]
    chave_filtro = apenas_digitos(filtro_rg_admin) if filtro_rg_admin else "todos"

    df_editado = st.data_editor(
        df_editor,
        use_container_width=True,
        hide_index=True,
        disabled=["ID", "ORDEM", "POSTO_GRADUACAO", "RG", "NOME_DE_ESCALA", "SENHA", "AUTORIZADO", "ATUALIZADO_EM", "VERSAO"],
        column_config={
            "ID": st.column_config.TextColumn("ID"),
            "ORDEM": st.column_config.NumberColumn("ORDEM"),
            "POSTO_GRADUACAO": st.column_config.TextColumn("POSTO/GRADUAÇÃO"),
            "RG": st.column_config.TextColumn("RG"),
            "NOME_DE_ESCALA": st.column_config.TextColumn("NOME DE ESCALA"),
            "SENHA": st.column_config.TextColumn("SENHA"),
            "AUTORIZADO": st.column_config.TextColumn("AUTORIZADO ATUAL"),
            "ATUALIZADO_EM": st.column_config.TextColumn("ATUALIZADO EM"),
            "VERSAO": st.column_config.NumberColumn("VERSÃO"),
            "LIBERAR_ACESSO": st.column_config.CheckboxColumn("LIBERAR ACESSO", help="Marcado = SIM | Desmarcado = NÃO"),
        },
        key=f"editor_autorizacao_{chave_filtro}",
    )

    st.markdown("")
    if st.button("💾 Salvar Autorizações", type="primary", use_container_width=True):
        solicitar_confirmacao("admin_salvar_autorizacoes", "Você está prestes a salvar as autorizações alteradas.", {"linhas": df_editado.to_dict(orient="records")})
        st.rerun()


# ==========================================================
# RENDERIZAÇÃO: ASSUNÇÃO DO SERVIÇO
# ==========================================================
def render_assuncao_servico():
    # Ao entrar/usar a Assunção, o Login deverá ficar oculto até clicar em Concluir Serviço nesta aba.
    if st.session_state.usuario_logado:
        st.session_state.assuncao_ativa = True

    # Quando uma data cadastrada é selecionada na listbox, o serviço correspondente
    # é carregado aqui, ANTES da criação dos widgets, evitando erro de alteração
    # de st.session_state depois que o campo já foi instanciado pelo Streamlit.
    if st.session_state.servico_pendente_carregar:
        registro_pendente = st.session_state.servico_pendente_carregar
        carregar_servico_na_tela(registro_pendente, modo_destino="visualizar")
        st.session_state.serv_data_cadastrada_carregada = normalizar_texto(registro_pendente.get("ID_SERVICO", ""))
        st.session_state.servico_pendente_carregar = None
        st.session_state.msg_servico = "✅ Dados do serviço selecionado foram carregados na tela."
        st.session_state.tipo_msg_servico = "success"

    acao_confirmada, dados_confirmados = renderizar_confirmacao(["servico_"])

    if acao_confirmada == "servico_novo":
        limpar_campos_servico()
        st.session_state.modo_servico = "novo"
        st.session_state.msg_servico = "🆕 Campos liberados para nova assunção de serviço."
        st.session_state.tipo_msg_servico = "info"
        st.rerun()

    if acao_confirmada == "servico_cancelar":
        if st.session_state.servico_atual:
            limpar_campos_servico_mantendo_registro_atual()
        else:
            limpar_campos_servico()
        st.session_state.msg_servico = "↩️ Operação cancelada. A tela foi limpa; os dados só serão exibidos novamente após selecionar Unidade e Data."
        st.session_state.tipo_msg_servico = "info"
        st.rerun()

    if acao_confirmada == "servico_editar":
        if not st.session_state.servico_atual:
            st.session_state.msg_servico = "⚠️ Não há serviço salvo para editar."
            st.session_state.tipo_msg_servico = "warning"
            st.rerun()
        # Recarrega explicitamente o último serviço salvo para dentro dos widgets
        # antes de liberar a edição. Assim DATA, horários, viatura, áreas, equipe
        # e observações aparecem preenchidos para alteração.
        carregar_servico_na_tela(st.session_state.servico_atual, modo_destino="editar")
        st.session_state.msg_servico = "✏️ Serviço liberado para edição com os dados do último serviço carregados."
        st.session_state.tipo_msg_servico = "success"
        st.rerun()

    if acao_confirmada == "servico_salvar":
        dados = dados_confirmados
        if dados.get("MODO") == "novo":
            sucesso, resultado = adicionar_servico(dados)
            if not sucesso:
                st.session_state.msg_servico = resultado
                st.session_state.tipo_msg_servico = "error"
                st.rerun()
            carregar_servico_na_tela(resultado)
            st.session_state.assuncao_ativa = True
            ativar_menu_areas_do_servico(resultado)
            st.session_state.area_menu_msg = "✅ Serviço salvo com sucesso. As abas das áreas foram exibidas no menu."
            st.session_state.area_menu_tipo_msg = "success"
            st.session_state.msg_servico = "✅ Serviço salvo com sucesso. As abas das áreas foram exibidas no menu."
            st.session_state.tipo_msg_servico = "success"
            st.rerun()

        if dados.get("MODO") == "editar":
            sucesso, resultado = atualizar_servico(dados)
            if not sucesso:
                st.session_state.msg_servico = resultado
                st.session_state.tipo_msg_servico = "warning"
                st.rerun()

            # Mantém o serviço atualizado em memória para Editar/Excluir/Continuar,
            # mas limpa a tela da Assunção. Os dados só voltam a aparecer ao
            # selecionar Unidade + Data cadastrada, ou ao clicar em Editar.
            st.session_state.servico_atual = resultado
            st.session_state.assuncao_ativa = True
            desativar_menu_areas()
            limpar_campos_servico_mantendo_registro_atual()
            st.session_state.msg_servico = "✅ Serviço atualizado com sucesso. A tela foi limpa; selecione Unidade e Data para visualizar novamente."
            st.session_state.tipo_msg_servico = "success"
            st.rerun()

    if acao_confirmada == "servico_excluir":
        registro = st.session_state.servico_atual
        if not registro:
            st.session_state.msg_servico = "⚠️ Não há serviço salvo para excluir."
            st.session_state.tipo_msg_servico = "warning"
            st.rerun()
        sucesso, mensagem = excluir_servico(registro.get("UNIDADE", ""), registro.get("ID_SERVICO", ""))
        if not sucesso:
            st.session_state.msg_servico = mensagem
            st.session_state.tipo_msg_servico = "error"
            st.rerun()
        encerrar_assuncao_sem_logout()
        st.session_state.msg_servico = mensagem + " O Login voltou a ficar visível."
        st.session_state.tipo_msg_servico = "success"
        st.rerun()

    if acao_confirmada == "servico_continuar_supervisao":
        registro = st.session_state.servico_atual

        if not registro:
            st.session_state.msg_servico = "⚠️ Não há serviço salvo para continuar a supervisão."
            st.session_state.tipo_msg_servico = "warning"
            st.rerun()

        ativar_menu_areas_do_servico(registro)
        st.session_state.area_menu_msg = "✅ Supervisão retomada. As abas das áreas foram reexibidas no menu."
        st.session_state.area_menu_tipo_msg = "success"
        st.rerun()

    if acao_confirmada == "servico_concluir":
        registro = st.session_state.servico_atual

        # O botão Concluir Serviço da aba Assunção fica sempre disponível.
        # Se houver serviço novo/edição em andamento, cancela tudo que estiver na tela e reexibe o Login.
        if st.session_state.modo_servico in ["novo", "editar"]:
            encerrar_assuncao_sem_logout()
            st.session_state.msg_servico = "✅ Serviço em andamento cancelado. O Login voltou a ficar visível."
            st.session_state.tipo_msg_servico = "success"
            st.rerun()

        # Se houver serviço salvo aberto, conclui o registro no Sheets e depois limpa a tela/reexibe o Login.
        if registro:
            status = normalizar_texto(registro.get("STATUS", "")).upper()

            if status != "CONCLUIDO":
                sucesso, mensagem = concluir_servico(
                    unidade=registro.get("UNIDADE", ""),
                    id_servico=registro.get("ID_SERVICO", ""),
                    versao_esperada=para_int(registro.get("VERSAO", ""), 1),
                )

                if not sucesso:
                    st.session_state.msg_servico = mensagem
                    st.session_state.tipo_msg_servico = "error"
                    st.rerun()

                encerrar_assuncao_sem_logout()
                st.session_state.msg_servico = mensagem + " O Login voltou a ficar visível."
                st.session_state.tipo_msg_servico = "success"
                st.rerun()

        # Se não houver serviço salvo, apenas encerra a tela da Assunção e reexibe o Login.
        encerrar_assuncao_sem_logout()
        st.session_state.msg_servico = "✅ Assunção encerrada. O Login voltou a ficar visível."
        st.session_state.tipo_msg_servico = "success"
        st.rerun()


    if st.session_state.msg_servico:
        if st.session_state.tipo_msg_servico == "success":
            st.success(st.session_state.msg_servico)
        elif st.session_state.tipo_msg_servico == "error":
            st.error(st.session_state.msg_servico)
        elif st.session_state.tipo_msg_servico == "warning":
            st.warning(st.session_state.msg_servico)
        else:
            st.info(st.session_state.msg_servico)

    # IMPORTANTE: aqui NÃO verificamos/criamos as abas das unidades.
    # A leitura do Sheets nesta guia fica restrita ao cadastro já cacheado dos usuários.
    df_usuarios = carregar_cadastros()
    opcoes_usuarios = montar_opcoes_usuarios(df_usuarios)

    modo = st.session_state.modo_servico
    servico_atual = st.session_state.servico_atual

    # Importante: não recarregamos automaticamente os dados de servico_atual.
    # A tela deve ficar limpa depois de editar/salvar/cancelar/terminar supervisão.
    # Os dados só aparecem quando o usuário selecionar Unidade + Data cadastrada
    # ou quando clicar em Editar para abrir o último serviço em modo de edição.

    campos_desabilitados = modo not in ["novo", "editar"]

    if modo == "visualizar" and not servico_atual:
        st.markdown("<div class='status-box'>Nenhum serviço aberto na tela. Clique em <b>Novo</b> para lançar uma nova assunção.</div>", unsafe_allow_html=True)
    elif modo == "visualizar" and servico_atual:
        st.markdown("<div class='status-box'>Há um serviço salvo em memória. A tela permanece sem dados; selecione Unidade e Data para visualizar, ou use Editar/Excluir/Continuar Supervisão.</div>", unsafe_allow_html=True)
    elif modo == "novo":
        st.markdown("<div class='status-box'>Modo atual: NOVA ASSUNÇÃO — preencha os dados e clique em Salvar.</div>", unsafe_allow_html=True)
    elif modo == "editar":
        st.markdown("<div class='status-box'>Modo atual: EDIÇÃO DA ASSUNÇÃO — altere os dados e clique em Salvar.</div>", unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown("### 🧭 Unidade e Continuidade do Serviço")
        st.markdown(
            "<div class='small-muted'>Selecione a unidade para visualizar as datas de serviços abertos já cadastrados. "
            "Ao escolher uma data, os dados do serviço serão carregados na tela.</div>",
            unsafe_allow_html=True,
        )

        unidade_options = [""] + UNIDADES_SERVICO
        unidade_atual = normalizar_texto(st.session_state.serv_unidade)

        col_unidade, col_data_cadastrada, col_escolher_data = st.columns([1.5, 1.5, 1])

        with col_unidade:
            st.selectbox(
                "Unidade",
                options=unidade_options,
                index=indice_opcao(unidade_options, unidade_atual),
                key="serv_unidade",
                disabled=False,
                placeholder="Selecione a unidade",
                on_change=ao_alterar_unidade_servico,
            )

        unidade_para_datas = normalizar_texto(st.session_state.serv_unidade)
        df_servicos_unidade = carregar_servicos_unidade(unidade_para_datas) if unidade_para_datas else pd.DataFrame(columns=COLUNAS_SERVICO)
        opcoes_datas, rotulos_datas = montar_opcoes_datas_servico(df_servicos_unidade)

        if st.session_state.serv_data_cadastrada not in opcoes_datas:
            st.session_state.serv_data_cadastrada = ""
            st.session_state.serv_data_cadastrada_carregada = ""

        with col_data_cadastrada:
            st.selectbox(
                "Datas cadastradas",
                options=opcoes_datas,
                key="serv_data_cadastrada",
                format_func=lambda valor: rotulos_datas.get(valor, "Selecione uma data cadastrada"),
                disabled=(not unidade_para_datas or modo in ["novo", "editar"]),
                on_change=ao_alterar_data_cadastrada_servico,
            )

        id_data_selecionada = normalizar_texto(st.session_state.serv_data_cadastrada)

        if id_data_selecionada and id_data_selecionada != normalizar_texto(st.session_state.serv_data_cadastrada_carregada):
            registro_selecionado = obter_servico_df_por_id(df_servicos_unidade, id_data_selecionada)

            if registro_selecionado:
                st.session_state.servico_pendente_carregar = registro_selecionado
                st.rerun()

        with col_escolher_data:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            botao_escolher_data = st.button(
                "📅 Escolher data",
                disabled=(not id_data_selecionada or modo in ["novo", "editar"]),
                use_container_width=True,
                type="primary",
            )

        if botao_escolher_data:
            registro_escolhido = obter_servico_df_por_id(df_servicos_unidade, id_data_selecionada)

            if not registro_escolhido:
                st.session_state.msg_servico = "⚠️ Não foi possível localizar o serviço dessa data. Atualize e tente novamente."
                st.session_state.tipo_msg_servico = "warning"
                st.rerun()

            # A continuidade pela data cadastrada não será bloqueada pelo STATUS
            # do serviço. A única trava solicitada aqui é: o usuário logado
            # precisa fazer parte da equipe lançada naquela supervisão.
            if not usuario_logado_participa_servico(registro_escolhido):
                st.session_state.msg_servico = "❌ Impossível dar continuidade ao serviço porque este usuário não faz parte daquela supervisão."
                st.session_state.tipo_msg_servico = "error"
                st.rerun()

            st.session_state.servico_atual = registro_escolhido
            st.session_state.modo_servico = "visualizar"
            st.session_state.assuncao_ativa = True
            ativar_menu_areas_do_servico(registro_escolhido)
            st.session_state.area_menu_msg = "✅ Data escolhida. Supervisão retomada nas abas das áreas."
            st.session_state.area_menu_tipo_msg = "success"
            st.rerun()

    # ======================================================
    # QUANTIDADE DE ÁREAS SUPERVISIONADAS
    # ======================================================
    # Este campo fica FORA do st.form para que, ao alterar a quantidade,
    # o Streamlit atualize imediatamente a tela e exiba somente a quantidade
    # correta de caixas de nome de área. Os demais campos continuam dentro
    # do form para manter a edição estável.
    with st.container(border=True):
        st.markdown("### 📍 Quantidade de Áreas Supervisionadas")
        st.number_input(
            "Nº DE ÁREAS SUPERVISIONADAS",
            min_value=0,
            max_value=5,
            step=1,
            key="serv_num_areas",
            disabled=campos_desabilitados,
        )

    # ======================================================
    # FORMULÁRIO PRINCIPAL DO SERVIÇO
    # ======================================================
    # CORREÇÃO EFETIVA DO SALVAR:
    # - Os campos usam chaves próprias do formulário, versionadas por serv_form_version.
    # - O botão Salvar grava DIRETO no Google Sheets no mesmo clique.
    # - Não há confirmação intermediária para Salvar, pois ela fazia a tela rerenderizar
    #   e os valores editados podiam ser perdidos/reaproveitados do registro anterior.
    form_version = st.session_state.get("serv_form_version", 0)

    with st.form("form_servico_principal", clear_on_submit=False):
        with st.container(border=True):
            st.markdown("### 📍 Áreas Supervisionadas")
            qtd_areas_visiveis = para_int(st.session_state.serv_num_areas, 0)
            areas_form = {}

            if qtd_areas_visiveis <= 0:
                st.info("Informe o número de áreas supervisionadas para abrir os campos correspondentes.")
            else:
                st.markdown("#### Nome das Áreas")
                st.markdown(
                    "<div class='small-muted'>Serão exibidas somente as caixas correspondentes ao número informado. "
                    "Os nomes serão salvos em CAIXA ALTA.</div>",
                    unsafe_allow_html=True,
                )

                col_area_1, col_area_2 = st.columns(2)

                for i in range(1, qtd_areas_visiveis + 1):
                    coluna_destino = col_area_1 if i % 2 == 1 else col_area_2
                    with coluna_destino:
                        areas_form[i] = st.text_input(
                            f"NOME DA ÁREA {i}",
                            value=st.session_state.get(f"serv_area_{i}", ""),
                            key=f"form_serv_area_{i}_{form_version}",
                            disabled=campos_desabilitados,
                            placeholder="Será salvo em CAIXA ALTA",
                        )

        with st.container(border=True):
            st.markdown("### 🗓️ Dados do Serviço")
            col1, col2, col3 = st.columns(3)
            with col1:
                serv_data_form = st.date_input(
                    "DATA",
                    value=st.session_state.serv_data,
                    format="DD/MM/YYYY",
                    disabled=campos_desabilitados,
                    key=f"form_serv_data_{form_version}",
                )
            with col2:
                serv_inicio_form = st.time_input(
                    "INÍCIO DO SERVIÇO",
                    value=st.session_state.serv_inicio,
                    disabled=campos_desabilitados,
                    key=f"form_serv_inicio_{form_version}",
                )
            with col3:
                serv_termino_form = st.time_input(
                    "TÉRMINO DO SERVIÇO",
                    value=st.session_state.serv_termino,
                    disabled=campos_desabilitados,
                    key=f"form_serv_termino_{form_version}",
                )

            col4, col5, col6 = st.columns(3)
            with col4:
                serv_viatura_form = st.text_input(
                    "VIATURA",
                    value=st.session_state.serv_viatura,
                    disabled=campos_desabilitados,
                    placeholder="Ex.: D-0000",
                    key=f"form_serv_viatura_{form_version}",
                )
            with col5:
                serv_km_inicial_form = st.text_input(
                    "KM INICIAL",
                    value=st.session_state.serv_km_inicial,
                    disabled=campos_desabilitados,
                    placeholder="Somente números",
                    key=f"form_serv_km_inicial_{form_version}",
                )
            with col6:
                serv_km_final_form = st.text_input(
                    "KM FINAL",
                    value=st.session_state.serv_km_final,
                    disabled=campos_desabilitados,
                    placeholder="Obrigatório apenas ao concluir serviço aberto",
                    key=f"form_serv_km_final_{form_version}",
                )

        with st.container(border=True):
            st.markdown("### 👥 Equipe de Serviço")
            col_eq1, col_eq2 = st.columns(2)
            with col_eq1:
                serv_supervisor_form = st.selectbox(
                    "Supervisor",
                    options=opcoes_usuarios,
                    index=indice_opcao(opcoes_usuarios, st.session_state.serv_supervisor),
                    disabled=campos_desabilitados,
                    key=f"form_serv_supervisor_{form_version}",
                )
                serv_motorista_form = st.selectbox(
                    "Motorista",
                    options=opcoes_usuarios,
                    index=indice_opcao(opcoes_usuarios, st.session_state.serv_motorista),
                    disabled=campos_desabilitados,
                    key=f"form_serv_motorista_{form_version}",
                )
                serv_seguranca_1_form = st.selectbox(
                    "Segurança1",
                    options=opcoes_usuarios,
                    index=indice_opcao(opcoes_usuarios, st.session_state.serv_seguranca_1),
                    disabled=campos_desabilitados,
                    key=f"form_serv_seguranca_1_{form_version}",
                )
            with col_eq2:
                serv_seguranca_2_form = st.selectbox(
                    "Segurança2",
                    options=opcoes_usuarios,
                    index=indice_opcao(opcoes_usuarios, st.session_state.serv_seguranca_2),
                    disabled=campos_desabilitados,
                    key=f"form_serv_seguranca_2_{form_version}",
                )
                serv_seguranca_3_form = st.selectbox(
                    "Segurança3",
                    options=opcoes_usuarios,
                    index=indice_opcao(opcoes_usuarios, st.session_state.serv_seguranca_3),
                    disabled=campos_desabilitados,
                    key=f"form_serv_seguranca_3_{form_version}",
                )

        with st.container(border=True):
            st.markdown("### 📝 Observações Gerais")
            serv_observacoes_form = st.text_area(
                "Observações gerais",
                value=st.session_state.serv_observacoes,
                disabled=campos_desabilitados,
                placeholder="Descreva qualquer alteração ocorrida no início do serviço. Será salvo em CAIXA ALTA.",
                height=140,
                key=f"form_serv_observacoes_{form_version}",
            )

        st.markdown("---")
        st.markdown("### ⚙️ Comandos do Serviço")
        colb1, colb2, colb3, colb4, colb5, colb6, colb7 = st.columns(7)

        novo_desabilitado = modo in ["novo", "editar"] or servico_atual is not None
        editar_desabilitado = modo in ["novo", "editar"] or servico_atual is None
        salvar_desabilitado = modo not in ["novo", "editar"]
        cancelar_desabilitado = modo not in ["novo", "editar"]
        excluir_desabilitado = modo in ["novo", "editar"] or servico_atual is None
        continuar_desabilitado = modo in ["novo", "editar"] or servico_atual is None or not json_loads_lista((servico_atual or {}).get("NOMES_AREAS_JSON", "[]"))
        concluir_desabilitado = False

        with colb1:
            botao_serv_novo = st.form_submit_button("🆕 Novo", disabled=novo_desabilitado, use_container_width=True)
        with colb2:
            botao_serv_editar = st.form_submit_button("✏️ Editar", disabled=editar_desabilitado, use_container_width=True)
        with colb3:
            botao_serv_salvar = st.form_submit_button("💾 Salvar", type="primary", disabled=salvar_desabilitado, use_container_width=True)
        with colb4:
            botao_serv_cancelar = st.form_submit_button("↩️ Cancelar", disabled=cancelar_desabilitado, use_container_width=True)
        with colb5:
            botao_serv_excluir = st.form_submit_button("🗑️ Excluir", disabled=excluir_desabilitado, use_container_width=True)
        with colb6:
            botao_serv_continuar = st.form_submit_button("▶️ Continuar Supervisão", disabled=continuar_desabilitado, use_container_width=True)
        with colb7:
            botao_serv_concluir = st.form_submit_button("✅ Concluir Serviço", disabled=concluir_desabilitado, use_container_width=True)

    if botao_serv_novo:
        solicitar_confirmacao("servico_novo", "Você está prestes a iniciar uma nova assunção de serviço.", {})
        st.rerun()

    if botao_serv_editar:
        solicitar_confirmacao("servico_editar", "Você está prestes a editar este serviço.", {})
        st.rerun()

    if botao_serv_cancelar:
        solicitar_confirmacao("servico_cancelar", "Você está prestes a cancelar a operação atual.", {})
        st.rerun()

    if botao_serv_excluir:
        solicitar_confirmacao("servico_excluir", "Você está prestes a excluir este serviço e as abas de áreas criadas para ele. Essa ação não poderá ser desfeita.", {})
        st.rerun()

    if botao_serv_continuar:
        solicitar_confirmacao("servico_continuar_supervisao", "Você está prestes a reabrir as abas das áreas e continuar a supervisão de onde parou.", {})
        st.rerun()

    if botao_serv_concluir:
        solicitar_confirmacao(
            "servico_concluir",
            "Você está prestes a concluir/encerrar a Assunção do Serviço. Se houver lançamento em andamento, ele será cancelado; se houver serviço salvo aberto, ele será concluído no Sheets. O Login voltará a ficar visível.",
            {},
        )
        st.rerun()

    if botao_serv_salvar:
        unidade_form = normalizar_texto(st.session_state.serv_unidade)
        data_servico = data_para_texto(serv_data_form)
        inicio = hora_para_texto(serv_inicio_form)
        termino = hora_para_texto(serv_termino_form)
        viatura = texto_caixa_alta(serv_viatura_form)
        km_inicial = normalizar_texto(serv_km_inicial_form)
        km_final = normalizar_texto(serv_km_final_form)
        num_areas = para_int(st.session_state.serv_num_areas, 0)
        supervisor = normalizar_texto(serv_supervisor_form)
        motorista = normalizar_texto(serv_motorista_form)
        seguranca_1 = normalizar_texto(serv_seguranca_1_form)
        seguranca_2 = normalizar_texto(serv_seguranca_2_form)
        seguranca_3 = normalizar_texto(serv_seguranca_3_form)
        observacoes = texto_caixa_alta(serv_observacoes_form)

        erros = []
        if not unidade_form:
            erros.append("Selecione a unidade.")
        if not data_servico:
            erros.append("Selecione a data.")
        if not inicio:
            erros.append("Informe o início do serviço.")
        if not termino:
            erros.append("Informe o término do serviço.")
        if not viatura:
            erros.append("Informe a viatura.")
        if not km_inicial or not apenas_digitos(km_inicial):
            erros.append("Informe o KM inicial usando somente números.")
        if km_final and not apenas_digitos(km_final):
            erros.append("Informe o KM final usando somente números, ou deixe em branco até a conclusão.")
        if num_areas < 1 or num_areas > 5:
            erros.append("Informe o número de áreas supervisionadas entre 1 e 5.")
        if not supervisor:
            erros.append("Selecione o Supervisor.")

        areas = []
        for i in range(1, num_areas + 1):
            nome_area = texto_caixa_alta(areas_form.get(i, ""))
            if not nome_area:
                erros.append(f"Informe o nome da área {i}.")
            else:
                areas.append(nome_area)

        if len(set(areas)) != len(areas):
            erros.append("Não repita nomes de áreas.")

        nomes_reservados_menu = {"TERMINAR SUPERVISÃO", "TERMINAR SUPERVISAO"}
        if any(area in nomes_reservados_menu for area in areas):
            erros.append("O nome TERMINAR SUPERVISÃO é reservado para a aba do menu. Use outro nome para a área.")

        if erros:
            for erro in erros:
                st.warning(f"⚠️ {erro}")
            st.stop()

        dados = {
            "UNIDADE": unidade_form,
            "DATA": data_servico,
            "INICIO_SERVICO": inicio,
            "TERMINO_SERVICO": termino,
            "VIATURA": viatura,
            "KM_INICIAL": apenas_digitos(km_inicial),
            "KM_FINAL": apenas_digitos(km_final),
            "NUM_AREAS": num_areas,
            "AREAS": areas,
            "SUPERVISOR": supervisor,
            "MOTORISTA": motorista,
            "SEGURANCA_1": seguranca_1,
            "SEGURANCA_2": seguranca_2,
            "SEGURANCA_3": seguranca_3,
            "OBSERVACOES_GERAIS": observacoes,
            "MODO": modo,
        }

        # Sincroniza também o estado visual com os dados submetidos.
        st.session_state.serv_data = serv_data_form
        st.session_state.serv_inicio = serv_inicio_form
        st.session_state.serv_termino = serv_termino_form
        st.session_state.serv_viatura = viatura
        st.session_state.serv_km_inicial = apenas_digitos(km_inicial)
        st.session_state.serv_km_final = apenas_digitos(km_final)
        st.session_state.serv_supervisor = supervisor
        st.session_state.serv_motorista = motorista
        st.session_state.serv_seguranca_1 = seguranca_1
        st.session_state.serv_seguranca_2 = seguranca_2
        st.session_state.serv_seguranca_3 = seguranca_3
        st.session_state.serv_observacoes = observacoes
        for i in range(1, 6):
            st.session_state[f"serv_area_{i}"] = areas[i - 1] if i <= len(areas) else ""

        if modo == "novo":
            sucesso, resultado = adicionar_servico(dados)
            if not sucesso:
                st.session_state.msg_servico = resultado
                st.session_state.tipo_msg_servico = "error"
                st.rerun()

            carregar_servico_na_tela(resultado)
            st.session_state.assuncao_ativa = True
            ativar_menu_areas_do_servico(resultado)
            st.session_state.area_menu_msg = "✅ Serviço salvo com sucesso. As abas das áreas foram exibidas no menu."
            st.session_state.area_menu_tipo_msg = "success"
            st.session_state.msg_servico = "✅ Serviço salvo com sucesso. As abas das áreas foram exibidas no menu."
            st.session_state.tipo_msg_servico = "success"
            st.rerun()

        if modo == "editar":
            if not servico_atual:
                st.warning("⚠️ Não há serviço aberto para editar.")
                st.stop()

            dados["ID_SERVICO"] = servico_atual.get("ID_SERVICO", "")
            dados["VERSAO"] = para_int(servico_atual.get("VERSAO", ""), 1)

            sucesso, resultado = atualizar_servico(dados)
            if not sucesso:
                st.session_state.msg_servico = resultado
                st.session_state.tipo_msg_servico = "warning"
                st.rerun()

            st.session_state.servico_atual = resultado
            st.session_state.assuncao_ativa = True
            desativar_menu_areas()
            limpar_campos_servico_mantendo_registro_atual()
            limpar_cache_servicos()
            st.session_state.msg_servico = "✅ Serviço atualizado com sucesso no Google Sheets. A tela foi limpa; selecione Unidade e Data para visualizar novamente."
            st.session_state.tipo_msg_servico = "success"
            st.rerun()


# ==========================================================
# RENDERIZAÇÃO: ABAS DAS ÁREAS E CONCLUIR
# ==========================================================
def render_area_menu(nome_area: str):
    st.markdown(f"<div class='service-title'>{nome_area}</div>", unsafe_allow_html=True)

    if st.session_state.area_menu_msg:
        if st.session_state.area_menu_tipo_msg == "success":
            st.success(st.session_state.area_menu_msg)
        elif st.session_state.area_menu_tipo_msg == "error":
            st.error(st.session_state.area_menu_msg)
        elif st.session_state.area_menu_tipo_msg == "warning":
            st.warning(st.session_state.area_menu_msg)
        else:
            st.info(st.session_state.area_menu_msg)

    registro = st.session_state.servico_atual or {}

    with st.container(border=True):
        st.markdown("### 📍 Área supervisionada")
        st.info(f"**Área:** {nome_area}")

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.write(f"**Unidade:** {registro.get('UNIDADE', '')}")
        with col_b:
            st.write(f"**Data:** {registro.get('DATA', '')}")
        with col_c:
            st.write(f"**Status:** {registro.get('STATUS', 'ABERTO')}")

    with st.container(border=True):
        st.markdown("### ℹ️ Próxima etapa")
        st.write(
            "Esta aba foi aberta automaticamente a partir das áreas informadas na Assunção do Serviço. "
            "Na próxima fase, podemos inserir aqui os campos específicos de fiscalização/supervisão dessa área."
        )


def render_terminar_supervisao_menu():
    st.markdown("<div class='service-title'>TERMINAR SUPERVISÃO</div>", unsafe_allow_html=True)

    acao_confirmada, _ = renderizar_confirmacao(["area_menu_"])

    if acao_confirmada == "area_menu_terminar_supervisao":
        # Esta ação NÃO conclui o serviço no Sheets.
        # Ela apenas fecha/oculta as abas das áreas e retorna para a aba Assunção do Serviço.
        # A tela da Assunção volta limpa; os dados só aparecem novamente se Unidade/Data
        # forem selecionadas ou se o usuário clicar em Editar.
        desativar_menu_areas()
        st.session_state.assuncao_ativa = True
        limpar_campos_servico_mantendo_registro_atual()
        st.session_state.pagina_atual = "🕘 ASSUNÇÃO DO SERVIÇO"
        st.session_state.msg_servico = "✅ Supervisão finalizada no menu. As abas das áreas foram ocultadas e a Assunção foi reexibida sem dados na tela."
        st.session_state.tipo_msg_servico = "success"
        st.rerun()

    registro = st.session_state.servico_atual or {}
    areas = st.session_state.area_menu_abas or []

    with st.container(border=True):
        st.markdown("### ✅ Terminar Supervisão")
        st.write(
            "Ao terminar a supervisão, as abas das áreas serão ocultadas e você retornará para a aba ASSUNÇÃO DO SERVIÇO. "
            "Esta ação não conclui nem altera o serviço salvo no Google Sheets."
        )

        col_1, col_2, col_3 = st.columns(3)
        with col_1:
            st.info(f"**Unidade:** {registro.get('UNIDADE', '')}")
        with col_2:
            st.info(f"**Data:** {registro.get('DATA', '')}")
        with col_3:
            st.info(f"**Áreas:** {len(areas)}")

        if areas:
            st.markdown("**Áreas abertas no menu:**")
            st.write(", ".join(areas))

    if st.button("✅ TERMINAR SUPERVISÃO", type="primary", use_container_width=True):
        solicitar_confirmacao(
            "area_menu_terminar_supervisao",
            "Você está prestes a ocultar as abas das áreas e retornar para a Assunção do Serviço.",
            {},
        )
        st.rerun()


# ==========================================================
# EXECUÇÃO DAS ABAS VISÍVEIS
# ==========================================================
if pagina_selecionada == "🔐 Login":
    render_login()
elif pagina_selecionada == "📝 Cadastro":
    render_cadastro()
elif pagina_selecionada == "👨‍💼 Administrador":
    render_admin()
elif pagina_selecionada == "🕘 ASSUNÇÃO DO SERVIÇO":
    render_assuncao_servico()
elif pagina_selecionada == "TERMINAR SUPERVISÃO":
    render_terminar_supervisao_menu()
elif pagina_selecionada in (st.session_state.area_menu_abas or []):
    render_area_menu(pagina_selecionada)
