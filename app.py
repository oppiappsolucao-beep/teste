import base64
import html
import io
import json
import os
import re
import unicodedata
import uuid
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components
from google.oauth2.service_account import Credentials
from gspread.exceptions import SpreadsheetNotFound, WorksheetNotFound


# =========================================================
# CONFIGURAÇÕES GERAIS
# =========================================================
st.set_page_config(
    page_title="Dashboard Oppi Comercial",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

SHEET_ID = "1GAbrca0NSiJfPXaSte1qGxXCsGkQPacoRsm0PVB51gE"
WORKSHEET_NAME = "Folha1"
CACHE_TTL_SECONDS = 120

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Status usados no dashboard já organizados para a nova estrutura da planilha.
# A coluna T da planilha é o Status WhatsApp e a coluna V é o Status Ligação.
# Cada coluna tem sua própria lista, igual aos filtros configurados no Google Sheets.
STATUS_WHATSAPP_OPTIONS = [
    "Novo Lead",
    "Chamado Whats",
    "Conversando",
    "Reunião",
    "Proposta",
    "Sem interesse",
    "Fechado",
    "Sem Resposta",
    "Sem Whatsapp",
    "Retornar",
]

STATUS_LIGACAO_OPTIONS = [
    "Ligação - Conversando Whats",
    "Ligação não atende/cx",
    "Ligação Numero errado",
    "Ligação retornar",
    "Proposta",
    "Reunião",
    "Sem interesse",
]

STATUS_OPTIONS = list(dict.fromkeys(STATUS_WHATSAPP_OPTIONS + STATUS_LIGACAO_OPTIONS))
STATUS_WHATSAPP_SELECT_OPTIONS = ["Sem status"] + STATUS_WHATSAPP_OPTIONS
STATUS_LIGACAO_SELECT_OPTIONS = ["Sem status"] + STATUS_LIGACAO_OPTIONS
STATUS_SELECT_OPTIONS = ["Sem status"] + STATUS_OPTIONS

# Cards da Visão Geral seguindo exatamente os status do filtro único.
# Não agrupa os status de ligação e não soma categorias diferentes.
DASHBOARD_STATUS_OPTIONS = [
    "Novo Lead",
    "Chamado Whats",
    "Conversando",
    "Reunião",
    "Proposta",
    "Sem interesse",
    "Fechado",
    "Sem Resposta",
    "Sem Whatsapp",
    "Retornar",
    "Ligação - Conversando Whats",
    "Ligação não atende/cx",
    "Ligação Numero errado",
    "Ligação retornar",
]

STATUS_COLORS = {
    "Novo Lead": ("#E8F0FF", "#5C8BFF"),
    "Chamado Whats": ("#E8FFF0", "#00C853"),
    "Conversando": ("#F8EFE6", "#B37A2A"),
    "Sem interesse": ("#E9F8FA", "#2F9FB3"),
    "Não responde": ("#FBECEF", "#DA5C78"),
    "Sem Resposta": ("#FBECEF", "#DA5C78"),
    "Fechado": ("#EAF8EF", "#58B97A"),
    "Proposta": ("#EAF2FF", "#5C9DFF"),
    "Reunião": ("#F3EAFE", "#A65BDB"),
    "Ligação": ("#EAF8FF", "#3C92A8"),
    "Ligação - Conversando Whats": ("#E8FFF0", "#3C92A8"),
    "Ligação não atende/cx": ("#EAF8FF", "#1F6B7A"),
    "Ligação Numero errado": ("#FFE8E8", "#C40000"),
    "Ligação retornar": ("#EAF2FF", "#2F6BBA"),
    "Retornar": ("#EAF2FF", "#2F6BBA"),
    "Sem Whatsapp": ("#FFF3E6", "#8B4A00"),
}


# =========================================================
# ESTADO DA SESSÃO
# =========================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "auth_error" not in st.session_state:
    st.session_state.auth_error = ""

if "selected_page" not in st.session_state:
    st.session_state.selected_page = "Visão Geral"

if "selected_cadastro_subpage" not in st.session_state:
    st.session_state.selected_cadastro_subpage = "Novo cadastro"

if "selected_contract_sheet_row" not in st.session_state:
    st.session_state.selected_contract_sheet_row = None

if "navigation_session_token" not in st.session_state:
    st.session_state.navigation_session_token = ""


# =========================================================
# UTILITÁRIOS
# =========================================================
def render_html(content: str) -> None:
    """Renderiza HTML sem que o Streamlit o transforme em bloco de código."""
    clean_content = " ".join(
        line.strip()
        for line in content.splitlines()
        if line.strip()
    )
    st.markdown(clean_content, unsafe_allow_html=True)


def normalize_text(value) -> str:
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    return str(value).strip()


def normalize_search_text(value) -> str:
    text = normalize_text(value).lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(
        character
        for character in text
        if not unicodedata.combining(character)
    )
    return re.sub(r"\s+", " ", text).strip()


def flexible_search_match(search_value, target_value) -> bool:
    """
    Busca por empresa/telefone sem trazer empresas erradas.
    - Se digitar uma frase, precisa encontrar a frase inteira ou todos os termos.
    - Se digitar uma palavra só, encontra por essa palavra.
    - Se digitar telefone/CNPJ, compara pelos números.
    """
    term = normalize_search_text(search_value)
    target = normalize_search_text(target_value)

    if not term:
        return True

    if not target:
        return False

    if term in target:
        return True

    term_digits = normalize_digits(term)
    target_digits = normalize_digits(target)

    if term_digits and term_digits in target_digits:
        return True

    tokens = [
        token
        for token in re.split(r"\s+", term)
        if len(token) >= 3
    ]

    if not tokens:
        return False

    # Uma palavra só: pode encontrar por essa palavra.
    if len(tokens) == 1:
        return tokens[0] in target

    # Mais de uma palavra: precisa bater todos os termos digitados.
    # Isso evita que "Marmoraria Topazio" traga qualquer outra marmoraria.
    return all(token in target for token in tokens)


def infer_niche_from_company_name(value) -> str:
    """Identifica automaticamente o nicho usando palavras presentes no nome da empresa."""
    company_name = normalize_search_text(value)

    if not company_name:
        return "Não identificado"

    niche_keywords = [
        ("Marmoraria", [
            "marmoraria", "marmore", "marmores", "granito", "granitos",
            "pedra", "pedras", "revestimento", "revestimentos", "travertino",
        ]),
        ("Marcenaria", [
            "marcenaria", "marceneiro", "moveis", "movel", "planejados",
            "planejado", "armarios", "armario",
        ]),
        ("Academia", [
            "academia", "fitness", "gym", "crossfit", "jiu jitsu", "muay thai",
        ]),
        ("Clínica", [
            "clinica", "consultorio", "odontologia", "odontologica", "dental",
            "saude", "estetica",
        ]),
        ("Pet shop", [
            "pet shop", "petshop", "pet", "veterinaria", "veterinario",
        ]),
        ("Construção civil", [
            "construtora", "construcao", "engenharia", "arquitetura", "obra",
        ]),
        ("Restaurante", [
            "restaurante", "pizzaria", "lanchonete", "hamburgueria", "bar", "cafe",
        ]),
        ("Loja", [
            "loja", "comercio", "varejo", "store",
        ]),
        ("Serviços", [
            "servicos", "servico", "solucoes", "consultoria",
        ]),
    ]

    for niche_name, keywords in niche_keywords:
        if any(keyword in company_name for keyword in keywords):
            return niche_name

    return "Outros"


def infer_state_from_address(value) -> str:
    """Extrai a UF do endereço. Reconhece siglas e nomes completos dos estados brasileiros."""
    address_original = normalize_text(value)

    if not address_original:
        return "Não identificado"

    valid_ufs = [
        "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA",
        "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN",
        "RS", "RO", "RR", "SC", "SP", "SE", "TO",
    ]

    upper_address = address_original.upper()
    uf_matches = re.findall(r"(?<![A-Z])(" + "|".join(valid_ufs) + r")(?![A-Z])", upper_address)

    if uf_matches:
        return uf_matches[-1]

    normalized_address = normalize_search_text(address_original)
    state_names = {
        "acre": "AC",
        "alagoas": "AL",
        "amapa": "AP",
        "amazonas": "AM",
        "bahia": "BA",
        "ceara": "CE",
        "distrito federal": "DF",
        "espirito santo": "ES",
        "goias": "GO",
        "maranhao": "MA",
        "mato grosso do sul": "MS",
        "mato grosso": "MT",
        "minas gerais": "MG",
        "para": "PA",
        "paraiba": "PB",
        "parana": "PR",
        "pernambuco": "PE",
        "piaui": "PI",
        "rio de janeiro": "RJ",
        "rio grande do norte": "RN",
        "rio grande do sul": "RS",
        "rondonia": "RO",
        "roraima": "RR",
        "santa catarina": "SC",
        "sao paulo": "SP",
        "sergipe": "SE",
        "tocantins": "TO",
    }

    for state_name, uf in sorted(state_names.items(), key=lambda item: len(item[0]), reverse=True):
        if state_name in normalized_address:
            return uf

    return "Não identificado"


def parse_money(value) -> float:
    text = normalize_text(value)

    if not text:
        return 0.0

    text = text.replace("R$", "").replace(" ", "")

    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", "")

    try:
        return float(text)
    except Exception:
        return 0.0


def format_money(value) -> str:
    try:
        number = float(value)
    except Exception:
        number = 0.0

    return (
        f"R$ {number:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def parse_date(value):
    text = normalize_text(value)

    if not text:
        return pd.NaT

    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def normalize_period_filter(value):
    """
    Normaliza o retorno do st.date_input.
    Quando o usuário escolhe apenas um dia, o Streamlit pode retornar um único
    date em vez de uma tupla. Nesse caso, filtramos exatamente aquele dia.
    """
    if isinstance(value, (tuple, list)):
        cleaned_dates = [item for item in value if item is not None]

        if len(cleaned_dates) >= 2:
            start_date = cleaned_dates[0]
            end_date = cleaned_dates[1]
        elif len(cleaned_dates) == 1:
            start_date = cleaned_dates[0]
            end_date = cleaned_dates[0]
        else:
            return None, None
    elif isinstance(value, date):
        start_date = value
        end_date = value
    else:
        return None, None

    if start_date > end_date:
        start_date, end_date = end_date, start_date

    return start_date, end_date


def apply_period_filter(df: pd.DataFrame, date_column: str, period_value) -> pd.DataFrame:
    """
    Aplica o filtro por período respeitando a coluna Data do chamado.
    Para datas específicas, retorna somente registros daquele dia.
    Para intervalos, retorna somente registros entre início e fim.
    Linhas sem data não entram quando o usuário escolhe um período.
    """
    start_date, end_date = normalize_period_filter(period_value)

    if start_date is None or end_date is None or date_column not in df.columns:
        return df.copy()

    valid_dates = df[date_column].notna()

    return df[
        valid_dates
        & (df[date_column].dt.date >= start_date)
        & (df[date_column].dt.date <= end_date)
    ].copy()


def make_unique_headers(headers: list[str]) -> list[str]:
    result = []
    counter = {}

    for index, header in enumerate(headers):
        clean_header = normalize_text(header)

        if not clean_header:
            clean_header = f"Coluna {index + 1}"

        if clean_header in counter:
            counter[clean_header] += 1
            clean_header = f"{clean_header}_{counter[clean_header]}"
        else:
            counter[clean_header] = 1

        result.append(clean_header)

    return result


def first_existing_column(
    df: pd.DataFrame,
    possible_names: list[str],
) -> Optional[str]:
    normalized_columns = {
        normalize_search_text(column): column
        for column in df.columns
    }

    for name in possible_names:
        normalized_name = normalize_search_text(name)

        if normalized_name in normalized_columns:
            return normalized_columns[normalized_name]

    return None



def existing_column_by_occurrence(
    df: pd.DataFrame,
    possible_names: list[str],
    occurrence: int = 1,
) -> Optional[str]:
    """
    Encontra uma coluna considerando cabeçalhos repetidos da planilha.
    Exemplo: se a planilha tiver várias colunas chamadas "Telefone",
    o pandas recebe "Telefone", "Telefone_2", "Telefone_3".
    """
    normalized_aliases = {normalize_search_text(name) for name in possible_names}
    found = 0

    for column in df.columns:
        normalized_column = normalize_search_text(column)
        normalized_column_base = re.sub(r"_\d+$", "", normalized_column)

        if normalized_column in normalized_aliases or normalized_column_base in normalized_aliases:
            found += 1

            if found == occurrence:
                return column

    return None


def safe_series(
    df: pd.DataFrame,
    column: Optional[str],
    default_value="",
) -> pd.Series:
    if column and column in df.columns:
        return df[column]

    return pd.Series(
        [default_value] * len(df),
        index=df.index,
    )


def status_group(value: str) -> str:
    """Agrupa os status da nova planilha nos cards principais do dashboard."""
    status = normalize_search_text(value)

    if not status:
        return "Novo Lead"

    if any(word in status for word in ["chamado whats", "chamado whatsapp", "chamando whats", "chamando whatsapp"]):
        return "Chamado Whats"

    if any(word in status for word in ["sem whatsapp", "sem whats", "sem whats app"]):
        return "Sem Whatsapp"

    if any(word in status for word in ["reuniao", "reuniao marcada", "reuniao agendada"]):
        return "Reunião"

    if "proposta" in status:
        return "Proposta"

    if any(word in status for word in ["fechado", "ganho", "cliente"]):
        return "Fechado"

    if any(word in status for word in ["sem resposta", "nao responde", "nao respondeu", "nao atendeu", "nao atende"]):
        return "Não responde"

    if any(word in status for word in ["sem interesse", "nao tem interesse", "não tem interesse"]):
        return "Sem interesse"

    if status == "retornar" or "ligacao retornar" in status or "retornar" in status:
        return "Retornar"

    if any(word in status for word in ["ligacao", "ligando", "telefonema", "telefone"]):
        return "Ligação"

    if any(word in status for word in ["conversando", "contato", "negoci", "andamento"]):
        return "Conversando"

    if any(word in status for word in ["novo", "lead"]):
        return "Novo Lead"

    return normalize_text(value)


def dashboard_status_from_rows(status_whatsapp: str, status_ligacao: str) -> str:
    """
    Define o status principal do dashboard usando as duas colunas novas:
    1. Status WhatsApp, quando preenchido.
    2. Status Ligação, quando o WhatsApp ainda está vazio.
    3. Novo Lead, quando ambos estão vazios.
    """
    whatsapp_text = normalize_text(status_whatsapp)
    ligacao_text = normalize_text(status_ligacao)

    if whatsapp_text:
        return status_group(whatsapp_text)

    if ligacao_text:
        return status_group(ligacao_text)

    return "Novo Lead"


def row_matches_status_filter(row, selected_status: str) -> bool:
    """
    Filtro único de Status: procura o status escolhido nas duas colunas da planilha,
    sem somar e sem agrupar.

    Exemplo: se escolher "Proposta", retorna linhas com:
    - Status WhatsApp = Proposta
    OU
    - Status Ligação = Proposta
    """
    status_value = normalize_text(selected_status)

    if not status_value or status_value == "Todos os status":
        return True

    normalized_filter = normalize_search_text(status_value)
    whatsapp_status = normalize_search_text(row.get("_status_whatsapp_original", ""))
    ligacao_status = normalize_search_text(row.get("_status_ligacao_original", ""))

    return whatsapp_status == normalized_filter or ligacao_status == normalized_filter




def row_matches_dashboard_card(row, selected_status: str) -> bool:
    """
    Cards da Visão Geral: usa o mesmo filtro único das duas colunas.
    A única exceção é Novo Lead, que também considera linhas sem nenhum status preenchido.
    """
    status_value = normalize_text(selected_status)

    if not status_value:
        return True

    if row_matches_status_filter(row, status_value):
        return True

    if normalize_search_text(status_value) == "novo lead":
        whatsapp_status = normalize_text(row.get("_status_whatsapp_original", ""))
        ligacao_status = normalize_text(row.get("_status_ligacao_original", ""))
        return not whatsapp_status and not ligacao_status

    return False


def count_dashboard_status(df: pd.DataFrame, status_name: str) -> int:
    if df.empty:
        return 0

    return int(df.apply(lambda row: row_matches_dashboard_card(row, status_name), axis=1).sum())


def calculate_score(row: pd.Series, columns: dict) -> int:
    score = 0

    if normalize_text(row.get(columns.get("telefone_b2b", ""), "")):
        score += 15

    if normalize_text(row.get(columns.get("email", ""), "")):
        score += 10

    if normalize_text(row.get(columns.get("site", ""), "")):
        score += 10

    if normalize_text(row.get(columns.get("instagram", ""), "")):
        score += 10

    if normalize_text(row.get(columns.get("linkedin", ""), "")):
        score += 5

    if normalize_text(row.get(columns.get("socio_1", ""), "")):
        score += 10

    capital_value = parse_money(row.get(columns.get("capital", ""), ""))

    if capital_value >= 100000:
        score += 20
    elif capital_value >= 50000:
        score += 15
    elif capital_value > 0:
        score += 8

    grouped_status = status_group(row.get(columns.get("status", ""), ""))

    if grouped_status == "Fechado":
        score += 20
    elif grouped_status == "Proposta":
        score += 16
    elif grouped_status == "Reunião":
        score += 14
    elif grouped_status == "Conversando":
        score += 12
    elif grouped_status == "Ligação":
        score += 10
    elif grouped_status == "Novo Lead":
        score += 6

    return min(score, 100)


def score_classification(score: int) -> str:
    if score >= 70:
        return "Lead Quente"

    if score >= 40:
        return "Lead Morno"

    return "Lead Frio"




def get_logo_data_uri() -> str:
    """Usa a logo exatamente como a imagem original, sem remover fundo nem aplicar recorte."""
    possible_paths = [
        Path(__file__).parent / "logo_oppi.png",
        Path(__file__).parent / "logo.png",
        Path(__file__).parent / "assets" / "logo_oppi.png",
        Path(__file__).parent / "assets" / "logo.png",
    ]

    for file_path in possible_paths:
        if file_path.exists():
            mime = "image/png" if file_path.suffix.lower() == ".png" else "image/jpeg"
            encoded = base64.b64encode(file_path.read_bytes()).decode("utf-8")
            return f"data:{mime};base64,{encoded}"

    return ""
# =========================================================
# CONEXÃO COM GOOGLE SHEETS
# =========================================================
def get_runtime_setting(name: str, default: str = "") -> str:
    """
    Busca primeiro uma variável de ambiente do EasyPanel.
    Caso ela não exista, mantém compatibilidade com os Secrets do Streamlit Cloud.
    """
    environment_value = os.getenv(name)

    if normalize_text(environment_value):
        return str(environment_value)

    try:
        return str(st.secrets.get(name, default))
    except Exception:
        return default


def _decode_service_account_b64(raw_value: str) -> dict:
    """Converte o JSON da conta de serviço armazenado em Base64 no EasyPanel."""
    value = normalize_text(raw_value)

    if value.startswith("GCP_SERVICE_ACCOUNT_B64="):
        value = value.split("=", 1)[1].strip()

    if value.startswith("GOOGLE_SERVICE_ACCOUNT_B64="):
        value = value.split("=", 1)[1].strip()

    if not value:
        raise RuntimeError("A variável Base64 da conta de serviço está vazia.")

    try:
        decoded_json = base64.b64decode(value).decode("utf-8")
        credentials_info = json.loads(decoded_json)
    except Exception as error:
        raise RuntimeError(
            "Não consegui converter a variável GCP_SERVICE_ACCOUNT_B64 em JSON. "
            "Gere novamente o Base64 usando o arquivo JSON original da conta de serviço."
        ) from error

    if not isinstance(credentials_info, dict):
        raise RuntimeError("O conteúdo decodificado da conta de serviço não é um JSON válido.")

    return credentials_info


def _normalize_google_private_key(value: str) -> str:
    """Normaliza quebras de linha da chave privada sem alterar seu conteúdo."""
    private_key = str(value or "").strip()

    # Quando o valor foi salvo com aspas no painel, remove somente as aspas externas.
    if (private_key.startswith('"') and private_key.endswith('"')) or (
        private_key.startswith("'") and private_key.endswith("'")
    ):
        private_key = private_key[1:-1].strip()

    private_key = private_key.replace("\\n", "\n")

    if private_key and not private_key.endswith("\n"):
        private_key += "\n"

    return private_key


def _load_google_credentials_info() -> dict:
    """
    Prioridade de leitura:
    1. JSON completo em Base64 no EasyPanel;
    2. variáveis GOOGLE_* separadas;
    3. Secrets do Streamlit Cloud.

    Usar o JSON completo em Base64 evita misturar private_key, private_key_id e
    client_email de credenciais diferentes.
    """
    b64_credentials = (
        os.getenv("GCP_SERVICE_ACCOUNT_B64", "").strip()
        or os.getenv("GOOGLE_SERVICE_ACCOUNT_B64", "").strip()
    )

    if b64_credentials:
        credentials_info = _decode_service_account_b64(b64_credentials)
    else:
        credentials_info = {
            "type": os.getenv("GOOGLE_TYPE", ""),
            "project_id": os.getenv("GOOGLE_PROJECT_ID", ""),
            "private_key_id": os.getenv("GOOGLE_PRIVATE_KEY_ID", ""),
            "private_key": os.getenv("GOOGLE_PRIVATE_KEY", ""),
            "client_email": os.getenv("GOOGLE_CLIENT_EMAIL", ""),
            "client_id": os.getenv("GOOGLE_CLIENT_ID", ""),
            "auth_uri": os.getenv("GOOGLE_AUTH_URI", ""),
            "token_uri": os.getenv("GOOGLE_TOKEN_URI", ""),
            "auth_provider_x509_cert_url": os.getenv(
                "GOOGLE_AUTH_PROVIDER_X509_CERT_URL",
                "",
            ),
            "client_x509_cert_url": (
                os.getenv("GOOGLE_CLIENT_X509_CERT_URL", "")
                or os.getenv("_CLIENT_X509_CERT_URL", "")
            ),
            "universe_domain": os.getenv("GOOGLE_UNIVERSE_DOMAIN", "googleapis.com"),
        }

        required_env_fields = [
            "type",
            "project_id",
            "private_key_id",
            "private_key",
            "client_email",
            "client_id",
            "auth_uri",
            "token_uri",
            "auth_provider_x509_cert_url",
            "client_x509_cert_url",
        ]

        has_all_separate_env_values = all(
            normalize_text(credentials_info.get(field, ""))
            for field in required_env_fields
        )

        if not has_all_separate_env_values:
            try:
                credentials_info = dict(st.secrets["gcp_service_account"])
            except Exception as error:
                raise RuntimeError(
                    "Não encontrei credenciais completas do Google. No EasyPanel, "
                    "configure preferencialmente uma única variável chamada "
                    "GCP_SERVICE_ACCOUNT_B64 com o JSON completo convertido em Base64."
                ) from error

    required_fields = [
        "type",
        "project_id",
        "private_key_id",
        "private_key",
        "client_email",
        "client_id",
        "auth_uri",
        "token_uri",
        "auth_provider_x509_cert_url",
        "client_x509_cert_url",
    ]

    missing_fields = [
        field
        for field in required_fields
        if not normalize_text(credentials_info.get(field, ""))
    ]

    if missing_fields:
        raise RuntimeError(
            "A credencial do Google está incompleta. Campos ausentes: "
            + ", ".join(missing_fields)
        )

    credentials_info["private_key"] = _normalize_google_private_key(
        credentials_info.get("private_key", "")
    )

    return credentials_info


@st.cache_resource
def get_gsheet_client():
    """Conecta ao Google Sheets usando uma única credencial consistente."""
    credentials_info = _load_google_credentials_info()

    try:
        credentials = Credentials.from_service_account_info(
            credentials_info,
            scopes=SCOPES,
        )

        return gspread.authorize(credentials)
    except Exception as error:
        raise RuntimeError(
            "Não consegui preparar a credencial do Google. Gere uma nova chave JSON "
            "da conta de serviço e atualize a variável GCP_SERVICE_ACCOUNT_B64 no EasyPanel."
        ) from error


@st.cache_data(show_spinner=False, ttl=CACHE_TTL_SECONDS)
def load_sheet_data() -> pd.DataFrame:
    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SHEET_ID)
    worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    values = worksheet.get_all_values()

    if not values:
        return pd.DataFrame()

    headers = make_unique_headers(values[0])
    rows = values[1:]

    df = pd.DataFrame(rows, columns=headers)
    df["_sheet_row"] = list(range(2, len(rows) + 2))

    for column in df.columns:
        if column != "_sheet_row":
            df[column] = df[column].astype(str).str.strip()

    data_columns = [column for column in df.columns if column != "_sheet_row"]
    df = df[
        df[data_columns].apply(
            lambda row: any(normalize_text(value) for value in row),
            axis=1,
        )
    ].copy()

    return df.reset_index(drop=True)


def update_statuses_in_sheet(
    changes: list[dict],
    status_column_name: str,
    updated_at_column_name: Optional[str] = None,
) -> None:
    """Atualiza os status editados diretamente na planilha do Google Sheets."""
    if not changes:
        return

    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SHEET_ID)
    worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    headers = worksheet.row_values(1)

    if status_column_name not in headers:
        raise RuntimeError(
            f"Não encontrei a coluna '{status_column_name}' na planilha."
        )

    status_column_index = headers.index(status_column_name) + 1
    updated_at_column_index = None

    if updated_at_column_name and updated_at_column_name in headers:
        updated_at_column_index = headers.index(updated_at_column_name) + 1

    cells = []
    now_text = pd.Timestamp.now(tz="America/Sao_Paulo").strftime("%d/%m/%Y %H:%M")

    for change in changes:
        sheet_row = int(change["sheet_row"])
        new_status = normalize_text(change["status"])

        if new_status == "Sem status":
            new_status = ""

        if new_status and new_status not in STATUS_OPTIONS and new_status not in DASHBOARD_STATUS_OPTIONS:
            raise RuntimeError(f"Status inválido: {new_status}")

        cells.append(gspread.Cell(sheet_row, status_column_index, new_status))

        if updated_at_column_index:
            cells.append(gspread.Cell(sheet_row, updated_at_column_index, now_text))

    worksheet.update_cells(cells, value_input_option="USER_ENTERED")
    st.cache_data.clear()


def _set_sheet_value_by_header(
    row_values: list[str],
    headers: list[str],
    aliases: list[str],
    value,
    occurrence: int = 1,
) -> bool:
    """Preenche uma coluna da planilha procurando pelo nome do cabeçalho."""
    normalized_aliases = {normalize_search_text(alias) for alias in aliases}
    found = 0

    for index, header in enumerate(headers):
        if normalize_search_text(header) in normalized_aliases:
            found += 1

            if found == occurrence:
                row_values[index] = normalize_text(value)
                return True

    return False


class DuplicateRegistrationError(RuntimeError):
    """Impede o cadastro quando telefone, CPF ou CNPJ já existe na planilha."""


def normalize_digits(value) -> str:
    """Mantém somente números para comparar telefones e CPFs independentemente da máscara."""
    return re.sub(r"\D", "", normalize_text(value))


def normalize_phone_for_duplicate(value) -> str:
    """Normaliza telefone brasileiro, removendo DDI 55 quando informado."""
    digits = normalize_digits(value)

    if digits.startswith("55") and len(digits) in (12, 13):
        digits = digits[2:]

    return digits if len(digits) >= 8 else ""


def normalize_cpf_for_duplicate(value) -> str:
    """Normaliza CPF para comparação, ignorando campos vazios ou incompletos."""
    digits = normalize_digits(value)
    return digits if len(digits) == 11 else ""


def normalize_cnpj_for_duplicate(value) -> str:
    """Normaliza CNPJ para comparação, ignorando campos vazios ou incompletos."""
    digits = normalize_digits(value)
    return digits if len(digits) == 14 else ""


def _header_matches_any(header: str, aliases: list[str]) -> bool:
    normalized_header = normalize_search_text(header)
    return any(alias in normalized_header for alias in aliases)


def validate_unique_company_registration(payload: dict, worksheet, ignore_sheet_row: Optional[int] = None) -> None:
    """
    Bloqueia cadastro ou edição quando qualquer telefone, CPF ou CNPJ informado já existe
    em outra linha da planilha. A leitura é feita diretamente da aba para evitar
    duplicidade mesmo quando o cache ainda não atualizou.
    """
    values = worksheet.get_all_values()

    if not values:
        return

    headers = values[0]
    rows = values[1:]

    phone_column_indexes = [
        index
        for index, header in enumerate(headers)
        if _header_matches_any(header, ["telefone", "celular", "whatsapp", "fone"])
    ]

    cpf_column_indexes = [
        index
        for index, header in enumerate(headers)
        if normalize_search_text(header) == "cpf"
        or _header_matches_any(header, ["cpf do", "cpf socio", "cpf sócio"])
    ]

    cnpj_column_indexes = [
        index
        for index, header in enumerate(headers)
        if normalize_search_text(header) == "cnpj"
        or _header_matches_any(header, ["cnpj da empresa", "cnpj empresa"])
    ]

    submitted_phones = {
        normalize_phone_for_duplicate(payload.get(field))
        for field in [
            "telefone_b2b",
            "telefone_fixo",
            "telefone_alternativo",
            "telefone_socio_1",
            "telefone_socio_2",
            "telefone_socio_3",
        ]
    }
    submitted_phones.discard("")

    submitted_cpfs = {
        normalize_cpf_for_duplicate(payload.get(field))
        for field in [
            "cpf_socio_1",
            "cpf_socio_2",
            "cpf_socio_3",
        ]
    }
    submitted_cpfs.discard("")

    submitted_cnpjs = {
        normalize_cnpj_for_duplicate(payload.get("cnpj"))
    }
    submitted_cnpjs.discard("")

    duplicate_phones = set()
    duplicate_cpfs = set()
    duplicate_cnpjs = set()

    for row_offset, row in enumerate(rows, start=2):
        if ignore_sheet_row is not None and int(row_offset) == int(ignore_sheet_row):
            continue

        for index in phone_column_indexes:
            if index >= len(row):
                continue

            existing_phone = normalize_phone_for_duplicate(row[index])

            if existing_phone and existing_phone in submitted_phones:
                duplicate_phones.add(existing_phone)

        for index in cpf_column_indexes:
            if index >= len(row):
                continue

            existing_cpf = normalize_cpf_for_duplicate(row[index])

            if existing_cpf and existing_cpf in submitted_cpfs:
                duplicate_cpfs.add(existing_cpf)

        for index in cnpj_column_indexes:
            if index >= len(row):
                continue

            existing_cnpj = normalize_cnpj_for_duplicate(row[index])

            if existing_cnpj and existing_cnpj in submitted_cnpjs:
                duplicate_cnpjs.add(existing_cnpj)

    if not duplicate_phones and not duplicate_cpfs and not duplicate_cnpjs:
        return

    messages = []

    if duplicate_phones:
        phones_text = ", ".join(sorted(duplicate_phones))
        messages.append(f"Telefone já cadastrado: {phones_text}")

    if duplicate_cpfs:
        cpfs_text = ", ".join(sorted(duplicate_cpfs))
        messages.append(f"CPF já cadastrado: {cpfs_text}")

    if duplicate_cnpjs:
        cnpjs_text = ", ".join(sorted(duplicate_cnpjs))
        messages.append(f"CNPJ já cadastrado: {cnpjs_text}")

    raise DuplicateRegistrationError(
        "Não foi possível salvar porque já existe outro cadastro com os mesmos dados. " + " | ".join(messages)
    )


def append_company_to_sheet(payload: dict) -> None:
    """Adiciona uma nova empresa na aba principal respeitando a estrutura atual da planilha."""
    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SHEET_ID)
    worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    headers = worksheet.row_values(1)

    validate_unique_company_registration(payload, worksheet)

    if not headers:
        raise RuntimeError("A primeira linha da planilha precisa conter os cabeçalhos.")

    row_values = [""] * len(headers)

    _set_sheet_value_by_header(row_values, headers, ["Nome Empresas", "Nome da empresa", "Empresa", "Nome Empresa", "Nome empresas", "Nome Empresa(s)"], payload.get("empresa"))
    _set_sheet_value_by_header(row_values, headers, ["Data de abertura", "Data abertura"], payload.get("data_abertura"))
    _set_sheet_value_by_header(row_values, headers, ["Capital", "Capital social"], payload.get("capital"))
    _set_sheet_value_by_header(row_values, headers, ["CNPJ"], payload.get("cnpj"))
    _set_sheet_value_by_header(row_values, headers, ["Endereço", "Endereco"], payload.get("endereco"))
    _set_sheet_value_by_header(row_values, headers, ["Email", "E-mail"], payload.get("email_empresa"))
    _set_sheet_value_by_header(row_values, headers, ["Site empresa", "Site", "Website"], payload.get("site"))

    _set_sheet_value_by_header(row_values, headers, ["Telefone (b2b)", "Telefone b2b"], payload.get("telefone_b2b"))
    _set_sheet_value_by_header(row_values, headers, ["Telefone fixo", "Fixo"], payload.get("telefone_fixo"))
    _set_sheet_value_by_header(row_values, headers, ["Telefone lemitt", "Telefone alternativo", "Outro telefone"], payload.get("telefone_alternativo"))

    _set_sheet_value_by_header(row_values, headers, ["Sócio 1", "Socio 1", "Sócio1", "Socio1"], payload.get("socio_1"))
    _set_sheet_value_by_header(row_values, headers, ["CPF"], payload.get("cpf_socio_1"), occurrence=1)
    _set_sheet_value_by_header(row_values, headers, ["E-mail Sócio 1", "Email Sócio 1", "E-mail Socio 1", "Email Socio 1"], payload.get("email_socio_1"))
    _set_sheet_value_by_header(row_values, headers, ["Telefone"], payload.get("telefone_socio_1"), occurrence=1)

    _set_sheet_value_by_header(row_values, headers, ["Sócio 2", "Socio 2", "Sócio2", "Socio2"], payload.get("socio_2"))
    _set_sheet_value_by_header(row_values, headers, ["Telefone sócio 2", "Telefone socio 2", "Telefone do sócio 2", "Telefone do socio 2", "Telefone"], payload.get("telefone_socio_2"), occurrence=2)
    _set_sheet_value_by_header(row_values, headers, ["CPF"], payload.get("cpf_socio_2"), occurrence=2)

    _set_sheet_value_by_header(row_values, headers, ["Sócio 3", "Socio 3", "Sócio3", "Socio3"], payload.get("socio_3"))
    _set_sheet_value_by_header(row_values, headers, ["Telefone sócio 3", "Telefone socio 3", "Telefone do sócio 3", "Telefone do socio 3", "Telefone"], payload.get("telefone_socio_3"), occurrence=3)
    _set_sheet_value_by_header(row_values, headers, ["CPF"], payload.get("cpf_socio_3"), occurrence=3)

    _set_sheet_value_by_header(row_values, headers, ["Instagram"], payload.get("instagram"))
    _set_sheet_value_by_header(row_values, headers, ["Linkedin", "LinkedIn"], payload.get("linkedin"))
    _set_sheet_value_by_header(row_values, headers, ["Vendedor", "Responsável", "Responsavel"], payload.get("vendedor"))
    _set_sheet_value_by_header(row_values, headers, ["Status", "Etapa"], payload.get("status"))
    _set_sheet_value_by_header(row_values, headers, ["Data do chamado", "Data chamado"], payload.get("data_chamado"))
    _set_sheet_value_by_header(row_values, headers, ["Última atualização", "Ultima atualização", "Ultima atualizacao"], payload.get("ultima_atualizacao"))
    _set_sheet_value_by_header(row_values, headers, ["Observações", "Observacoes", "Observação", "Observacao"], payload.get("observacoes"))

    worksheet.append_row(
        row_values,
        value_input_option="USER_ENTERED",
        insert_data_option="INSERT_ROWS",
    )

    st.cache_data.clear()


def update_company_in_sheet(sheet_row: int, payload: dict) -> None:
    """Atualiza uma empresa diretamente na planilha, preservando as demais colunas da linha."""
    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SHEET_ID)
    worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    headers = worksheet.row_values(1)

    if not headers:
        raise RuntimeError("A primeira linha da planilha precisa conter os cabeçalhos.")

    validate_unique_company_registration(
        payload,
        worksheet,
        ignore_sheet_row=int(sheet_row),
    )

    current_row = worksheet.row_values(int(sheet_row))
    row_values = list(current_row) + [""] * max(0, len(headers) - len(current_row))
    row_values = row_values[:len(headers)]

    _set_sheet_value_by_header(row_values, headers, ["Nome Empresas", "Nome da empresa", "Empresa", "Nome Empresa", "Nome empresas", "Nome Empresa(s)"], payload.get("empresa"))
    _set_sheet_value_by_header(row_values, headers, ["Data de abertura", "Data abertura"], payload.get("data_abertura"))
    _set_sheet_value_by_header(row_values, headers, ["Capital", "Capital social"], payload.get("capital"))
    _set_sheet_value_by_header(row_values, headers, ["CNPJ"], payload.get("cnpj"))
    _set_sheet_value_by_header(row_values, headers, ["Endereço", "Endereco"], payload.get("endereco"))
    _set_sheet_value_by_header(row_values, headers, ["Email", "E-mail"], payload.get("email_empresa"))
    _set_sheet_value_by_header(row_values, headers, ["Site empresa", "Site", "Website"], payload.get("site"))

    _set_sheet_value_by_header(row_values, headers, ["Telefone (b2b)", "Telefone b2b"], payload.get("telefone_b2b"))
    _set_sheet_value_by_header(row_values, headers, ["Telefone fixo", "Fixo"], payload.get("telefone_fixo"))
    _set_sheet_value_by_header(row_values, headers, ["Telefone lemitt", "Telefone alternativo", "Outro telefone"], payload.get("telefone_alternativo"))

    _set_sheet_value_by_header(row_values, headers, ["Sócio 1", "Socio 1", "Sócio1", "Socio1"], payload.get("socio_1"))
    _set_sheet_value_by_header(row_values, headers, ["CPF"], payload.get("cpf_socio_1"), occurrence=1)
    _set_sheet_value_by_header(row_values, headers, ["E-mail Sócio 1", "Email Sócio 1", "E-mail Socio 1", "Email Socio 1"], payload.get("email_socio_1"))
    _set_sheet_value_by_header(row_values, headers, ["Telefone"], payload.get("telefone_socio_1"), occurrence=1)

    _set_sheet_value_by_header(row_values, headers, ["Sócio 2", "Socio 2", "Sócio2", "Socio2"], payload.get("socio_2"))
    _set_sheet_value_by_header(row_values, headers, ["Telefone sócio 2", "Telefone socio 2", "Telefone do sócio 2", "Telefone do socio 2", "Telefone"], payload.get("telefone_socio_2"), occurrence=2)
    _set_sheet_value_by_header(row_values, headers, ["CPF"], payload.get("cpf_socio_2"), occurrence=2)

    _set_sheet_value_by_header(row_values, headers, ["Sócio 3", "Socio 3", "Sócio3", "Socio3"], payload.get("socio_3"))
    _set_sheet_value_by_header(row_values, headers, ["Telefone sócio 3", "Telefone socio 3", "Telefone do sócio 3", "Telefone do socio 3", "Telefone"], payload.get("telefone_socio_3"), occurrence=3)
    _set_sheet_value_by_header(row_values, headers, ["CPF"], payload.get("cpf_socio_3"), occurrence=3)

    _set_sheet_value_by_header(row_values, headers, ["Instagram"], payload.get("instagram"))
    _set_sheet_value_by_header(row_values, headers, ["Linkedin", "LinkedIn"], payload.get("linkedin"))
    _set_sheet_value_by_header(row_values, headers, ["Vendedor", "Responsável", "Responsavel"], payload.get("vendedor"))
    _set_sheet_value_by_header(row_values, headers, ["Status", "Etapa"], payload.get("status"))
    _set_sheet_value_by_header(row_values, headers, ["Data do chamado", "Data chamado"], payload.get("data_chamado"))
    _set_sheet_value_by_header(row_values, headers, ["Última atualização", "Ultima atualização", "Ultima atualizacao"], payload.get("ultima_atualizacao"))
    _set_sheet_value_by_header(row_values, headers, ["Observações", "Observacoes", "Observação", "Observacao"], payload.get("observacoes"))

    changed_cells = []

    for column_index, new_value in enumerate(row_values, start=1):
        old_value = current_row[column_index - 1] if column_index - 1 < len(current_row) else ""

        if normalize_text(old_value) != normalize_text(new_value):
            changed_cells.append(gspread.Cell(int(sheet_row), column_index, normalize_text(new_value)))

    if changed_cells:
        worksheet.update_cells(changed_cells, value_input_option="USER_ENTERED")

    st.cache_data.clear()


# =========================================================
# IDENTIFICAÇÃO DAS COLUNAS
# =========================================================
def identify_columns(df: pd.DataFrame) -> dict:
    return {
        "empresa": first_existing_column(df, ["Nome Empresas", "Nome da empresa", "Empresa", "Nome Empresa", "Nome empresas", "Nome Empresa(s)"]),
        "data_abertura": first_existing_column(df, ["Data de abertura", "Data abertura"]),
        "capital": first_existing_column(df, ["Capital", "Capital social"]),
        "cnpj": first_existing_column(df, ["CNPJ"]),
        "endereco": first_existing_column(df, ["Endereço", "Endereco"]),
        "email": first_existing_column(df, ["Email", "E-mail"]),
        "site": first_existing_column(df, ["Site empresa", "Site", "Website"]),
        "telefone_b2b": first_existing_column(df, ["Telefone (b2b)", "Telefone b2b", "Telefone"]),
        "telefone_fixo": first_existing_column(df, ["Telefone fixo", "Fixo"]),
        "telefone_alternativo": first_existing_column(df, ["Telefone lemitt", "Telefone alternativo", "Outro telefone"]),
        "socio_1": first_existing_column(df, ["Sócio 1", "Socio 1", "Sócio1", "Socio1"]),
        "cpf_socio_1": first_existing_column(df, ["CPF"]),
        "email_socio_1": first_existing_column(df, ["E-mail Sócio 1", "Email Sócio 1", "E-mail Socio 1", "Email Socio 1"]),
        "telefone_socio_1": (
            first_existing_column(df, ["Telefone sócio 1", "Telefone socio 1", "Telefone cliente"])
            or existing_column_by_occurrence(df, ["Telefone"], occurrence=1)
        ),
        "socio_2": first_existing_column(df, ["Sócio 2", "Socio 2", "Sócio2", "Socio2"]),
        "telefone_socio_2": (
            first_existing_column(df, ["Telefone sócio 2", "Telefone socio 2", "Telefone do sócio 2", "Telefone do socio 2"])
            or existing_column_by_occurrence(df, ["Telefone"], occurrence=2)
        ),
        "cpf_socio_2": first_existing_column(df, ["CPF_2"]),
        "socio_3": first_existing_column(df, ["Sócio 3", "Socio 3", "Sócio3", "Socio3"]),
        "telefone_socio_3": (
            first_existing_column(df, ["Telefone sócio 3", "Telefone socio 3", "Telefone do sócio 3", "Telefone do socio 3"])
            or existing_column_by_occurrence(df, ["Telefone"], occurrence=3)
        ),
        "cpf_socio_3": first_existing_column(df, ["CPF_3"]),
        "instagram": first_existing_column(df, ["Instagram"]),
        "linkedin": first_existing_column(df, ["Linkedin", "LinkedIn"]),
        "vendedor": first_existing_column(df, ["Vendedor", "Responsável", "Responsavel"]),
        "status_whatsapp": first_existing_column(df, ["Status WhatsApp", "Status Whatsapp", "Status Whats", "Status Whats App"]),
        "status_ligacao": first_existing_column(df, ["Status Ligação", "Status Ligacao", "Status da Ligação", "Status da Ligacao"]),
        "status": first_existing_column(df, ["Status WhatsApp", "Status Whatsapp", "Status Whats", "Status Whats App", "Status", "Etapa"]),
        "data_chamado": first_existing_column(df, ["Data do chamado", "Data chamado"]),
        "ultima_atualizacao": first_existing_column(df, ["Última atualização", "Ultima atualização", "Ultima atualizacao"]),
        "observacoes": first_existing_column(df, ["Observações", "Observacoes", "Observação", "Observacao"]),
    }


def prepare_data(df: pd.DataFrame, columns: dict) -> pd.DataFrame:
    result = df.copy()

    empresa_column = columns.get("empresa") or first_existing_column(result, ["Nome Empresas", "Nome da empresa", "Empresa", "Nome Empresa"])
    result["_empresa"] = safe_series(result, empresa_column)
    result["_capital_num"] = safe_series(result, columns.get("capital")).apply(parse_money)
    result["_status_whatsapp_original"] = safe_series(result, columns.get("status_whatsapp") or columns.get("status"))
    result["_status_ligacao_original"] = safe_series(result, columns.get("status_ligacao"))
    result["_status_original"] = result["_status_whatsapp_original"].replace("", "Novo Lead")
    result["_status_grupo"] = result.apply(
        lambda row: dashboard_status_from_rows(
            row.get("_status_whatsapp_original", ""),
            row.get("_status_ligacao_original", ""),
        ),
        axis=1,
    )
    result["_vendedor"] = safe_series(result, columns.get("vendedor")).replace("", "Sem vendedor")
    result["_telefone"] = safe_series(result, columns.get("telefone_b2b"))
    result["_nicho"] = result["_empresa"].apply(infer_niche_from_company_name)
    result["_estado"] = safe_series(result, columns.get("endereco")).apply(infer_state_from_address)
    result["_data_chamado"] = safe_series(result, columns.get("data_chamado")).apply(parse_date)
    result["_ultima_atualizacao"] = safe_series(result, columns.get("ultima_atualizacao")).apply(parse_date)
    result["_pontuacao"] = result.apply(lambda row: calculate_score(row, columns), axis=1)
    result["_classificacao"] = result["_pontuacao"].apply(score_classification)

    return result


# =========================================================
# CSS DA TELA DE LOGIN
# =========================================================
def apply_login_css() -> None:
    render_html(
        """
        <style>
            .stApp {
                background:
                    radial-gradient(circle at 72% 47%, rgba(213, 56, 255, 0.22), transparent 26%),
                    radial-gradient(circle at 18% 82%, rgba(125, 0, 255, 0.10), transparent 18%),
                    linear-gradient(90deg, #05050C 0%, #090819 37%, #140A2E 68%, #170A2A 100%);
            }

            header[data-testid="stHeader"] {
                background: transparent !important;
            }

            section[data-testid="stSidebar"],
            [data-testid="collapsedControl"] {
                display: none !important;
            }

            .block-container {
                max-width: 1320px !important;
                padding-top: 1rem !important;
                padding-bottom: 1rem !important;
            }

            .login-brand-panel {
                min-height: 620px;
                border-radius: 36px;
                border: 1px solid rgba(255,255,255,0.06);
                padding: 42px 38px;
                position: relative;
                overflow: hidden;
                background:
                    linear-gradient(180deg, rgba(0,0,0,0.73), rgba(2,2,14,0.96)),
                    linear-gradient(145deg, #090910, #030309);
                box-shadow: 0 24px 60px rgba(0,0,0,0.30);
            }

            .login-brand-panel::before {
                content: "";
                position: absolute;
                left: -16%;
                right: -18%;
                bottom: -6%;
                height: 190px;
                background:
                    radial-gradient(circle at 18% 70%, rgba(255, 42, 154, 0.35), transparent 22%),
                    radial-gradient(circle at 50% 85%, rgba(119, 30, 255, 0.30), transparent 24%),
                    radial-gradient(circle at 82% 75%, rgba(255, 42, 154, 0.22), transparent 20%);
                filter: blur(18px);
                opacity: 0.90;
            }

            .login-logo-wrap {
                margin: 6px 0 38px 0;
            }

            .login-logo-img {
                width: 116px;
                height: 116px;
                object-fit: contain;
                display: block;
                filter: drop-shadow(0 18px 42px rgba(203, 38, 255, 0.24));
            }

            .login-logo-fallback {
                width: 116px;
                height: 116px;
                border-radius: 50%;
                background: linear-gradient(145deg, #FF4BAA 10%, #9C19FF 88%);
                position: relative;
                box-shadow: 0 18px 42px rgba(203, 38, 255, 0.24);
            }

            .login-logo-fallback::before {
                content: "";
                position: absolute;
                width: 40px;
                height: 40px;
                border-radius: 50%;
                background: #06060B;
                top: 28px;
                left: 38px;
            }

            .login-logo-fallback::after {
                content: "";
                position: absolute;
                width: 42px;
                height: 42px;
                left: 4px;
                bottom: 2px;
                transform: rotate(-7deg);
                clip-path: polygon(0 100%, 26% 26%, 100% 0);
                background: linear-gradient(145deg, #FF4BAA 10%, #9C19FF 88%);
                border-bottom-left-radius: 18px;
            }

            .login-brand-title {
                font-size: 2.95rem;
                line-height: 1.04;
                font-weight: 950;
                color: #FFFFFF;
                letter-spacing: -0.05em;
            }

            .login-brand-highlight {
                display: block;
                background: linear-gradient(90deg, #FF4BAA 0%, #D73AFF 50%, #8C2BFF 100%);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                background-clip: text;
            }

            .login-brand-subtitle {
                margin-top: 14px;
                color: rgba(255,255,255,0.84);
                font-size: 1.0rem;
            }

            .login-accent-line {
                width: 86px;
                height: 4px;
                border-radius: 999px;
                margin: 34px 0 34px 0;
                background: linear-gradient(90deg, #FF4BAA, #A62CFF);
            }

            .login-benefit {
                display: flex;
                align-items: center;
                gap: 18px;
                max-width: 320px;
                color: rgba(255,255,255,0.82);
                font-size: 0.94rem;
                line-height: 1.55;
            }

            .login-benefit-icon {
                width: 50px;
                height: 50px;
                min-width: 50px;
                border-radius: 15px;
                border: 2px solid rgba(183, 75, 255, 0.54);
                display: flex;
                align-items: center;
                justify-content: center;
                color: #D44BFF;
                font-size: 1.2rem;
            }

            .login-right-spacer {
                height: 66px;
            }

            [data-testid="stForm"] {
                position: relative !important;
                isolation: isolate !important;
                overflow: visible !important;
                background: #FFFFFF !important;
                border: none !important;
                border-radius: 30px !important;
                padding: 28px 32px 24px 32px !important;
                box-shadow:
                    0 28px 70px rgba(0,0,0,0.30),
                    0 0 0 1px rgba(255,255,255,0.55) !important;
                max-width: 640px !important;
                margin: 0 auto !important;
            }

            [data-testid="stForm"]::before {
                content: "";
                position: absolute;
                inset: -24px;
                border-radius: 42px;
                background:
                    radial-gradient(circle at 18% 52%, rgba(255, 72, 170, 0.34), transparent 34%),
                    radial-gradient(circle at 82% 50%, rgba(151, 42, 255, 0.34), transparent 36%),
                    radial-gradient(circle at 50% 100%, rgba(233, 56, 193, 0.24), transparent 32%);
                filter: blur(24px);
                z-index: -2;
                opacity: 1;
            }

            [data-testid="stForm"]::after {
                content: "";
                position: absolute;
                inset: -1px;
                border-radius: 31px;
                background: linear-gradient(135deg, rgba(255,255,255,0.90), rgba(255,255,255,0.72));
                z-index: -1;
            }

            .login-top-icon {
                width: 66px;
                height: 66px;
                border-radius: 50%;
                background: #F4EAFB;
                display: flex;
                align-items: center;
                justify-content: center;
                margin: 0 auto 12px auto;
                color: #A640FF;
                font-size: 1.55rem;
            }

            .login-card-title {
                text-align: center;
                color: #1E2230;
                font-size: 1.55rem;
                font-weight: 850;
                line-height: 1.28;
            }

            .login-card-subtitle {
                text-align: center;
                color: #7B8090;
                font-size: 1rem;
                margin-top: 5px;
                margin-bottom: 18px;
            }

            [data-testid="stForm"] label {
                color: #1F2430 !important;
                font-size: 0.98rem !important;
                font-weight: 750 !important;
            }

            [data-testid="stForm"] [data-baseweb="input"] {
                min-height: 54px !important;
                border: 1px solid #D7DAE2 !important;
                border-radius: 15px !important;
                background: #FFFFFF !important;
                box-shadow: none !important;
            }

            [data-testid="stForm"] input {
                color: #1F2330 !important;
                font-size: 1rem !important;
            }

            [data-testid="stForm"] .stButton > button,
            [data-testid="stForm"] button[kind="secondaryFormSubmit"] {
                width: 100% !important;
                min-height: 54px !important;
                border: none !important;
                border-radius: 15px !important;
                color: #FFFFFF !important;
                font-size: 1.03rem !important;
                font-weight: 850 !important;
                letter-spacing: 0.01em !important;
                background: linear-gradient(90deg, #FF4BAA 0%, #D73AFF 54%, #8C2BFF 100%) !important;
                box-shadow:
                    0 16px 32px rgba(188, 32, 255, 0.28),
                    0 6px 18px rgba(255, 75, 170, 0.18) !important;
                margin-top: 0.45rem !important;
                transition: transform 0.15s ease, box-shadow 0.15s ease !important;
            }

            [data-testid="stForm"] .stButton > button:hover,
            [data-testid="stForm"] button[kind="secondaryFormSubmit"]:hover {
                transform: translateY(-1px);
                box-shadow:
                    0 18px 36px rgba(188, 32, 255, 0.32),
                    0 8px 22px rgba(255, 75, 170, 0.20) !important;
            }

            .login-forgot-row {
                display: grid;
                grid-template-columns: 1fr auto 1fr;
                gap: 18px;
                align-items: center;
                margin-top: 18px;
            }

            .login-forgot-line {
                height: 1px;
                background: #E7E7EC;
            }

            .login-forgot-text {
                color: #A23BFF;
                font-weight: 750;
                font-size: 0.92rem;
            }

            .login-error {
                max-width: 640px;
                margin: 14px auto 0 auto;
                padding: 12px 14px;
                border-radius: 14px;
                background: #FFF0F3;
                color: #A02B42;
                border: 1px solid #FFC7D0;
                font-weight: 650;
                font-size: 0.95rem;
            }

            @media (max-width: 1050px) {
                .login-brand-panel {
                    min-height: auto;
                    padding: 30px;
                }

                .login-logo-wrap {
                    margin-top: 0;
                    margin-bottom: 24px;
                }

                .login-brand-title {
                    font-size: 2.2rem;
                }

                .login-right-spacer {
                    height: 0;
                }
            }
        </style>
        """
    )


# =========================================================
# CSS DO DASHBOARD
# =========================================================
def apply_dashboard_css() -> None:
    render_html(
        """
        <style>
            .stApp {
                background:
                    radial-gradient(circle at 78% 10%, rgba(202, 0, 255, 0.15), transparent 20%),
                    linear-gradient(120deg, #04040A 0%, #090915 34%, #140B2A 68%, #0A071A 100%);
            }

            header[data-testid="stHeader"] {
                background: transparent !important;
            }

            .block-container {
                max-width: 1600px !important;
                padding-top: 1.25rem !important;
                padding-bottom: 1.8rem !important;
            }

            section[data-testid="stSidebar"] {
                background:
                    radial-gradient(circle at top left, rgba(255,255,255,0.92) 0%, rgba(255,255,255,0.00) 32%),
                    radial-gradient(circle at bottom right, rgba(208,212,223,0.72) 0%, rgba(208,212,223,0.00) 34%),
                    linear-gradient(180deg, #F7F8FC 0%, #ECEEF4 38%, #DCE0E9 72%, #CED3DE 100%);
                border-right: 1px solid rgba(63, 53, 83, 0.12);
                box-shadow: 10px 0 34px rgba(0,0,0,0.16);
            }

            section[data-testid="stSidebar"] * {
                color: #20192F;
            }

            section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
                padding-top: 0.35rem !important;
            }

            .side-logo-wrap {
                width: 124px;
                height: 124px;
                display: flex;
                align-items: center;
                justify-content: center;
                overflow: visible;
                margin: -4px 0 8px -10px;
            }

            .side-logo-img {
                width: 116px;
                height: 116px;
                max-width: none;
                object-fit: contain;
                object-position: center;
                display: block;
                border-radius: 0;
                background: transparent;
                overflow: visible;
                filter: drop-shadow(0 12px 24px rgba(188, 45, 255, 0.18));
            }

            .side-logo-fallback {
                width: 92px;
                height: 92px;
                border-radius: 50%;
                background: linear-gradient(145deg, #FF4BAA 10%, #9C19FF 88%);
                position: relative;
                box-shadow: 0 12px 24px rgba(188, 45, 255, 0.20);
            }

            .side-logo-fallback::before {
                content: "";
                position: absolute;
                width: 32px;
                height: 32px;
                border-radius: 50%;
                background: #1B1725;
                top: 24px;
                left: 30px;
            }

            .side-logo-fallback::after {
                content: "";
                position: absolute;
                width: 32px;
                height: 32px;
                left: 2px;
                bottom: 1px;
                clip-path: polygon(0 100%, 25% 26%, 100% 0);
                background: linear-gradient(145deg, #FF4BAA 10%, #9C19FF 88%);
            }

            .side-title {
                color: #211A30;
                font-size: 1.16rem;
                font-weight: 900;
                line-height: 1.15;
            }

            .side-highlight {
                display: block;
                background: linear-gradient(90deg, #FF4BAA, #AE26FF);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                background-clip: text;
            }

            .side-subtitle {
                color: rgba(33,26,48,0.76);
                margin-top: 6px;
                font-size: 0.90rem;
            }

            .side-line {
                width: 70px;
                height: 4px;
                border-radius: 999px;
                margin: 18px 0 18px 0;
                background: linear-gradient(90deg, #FF4BAA, #AE26FF);
            }

            .side-tip {
                display: flex;
                gap: 12px;
                align-items: center;
                margin: 14px 0 18px 0;
                padding: 14px;
                border-radius: 16px;
                background: rgba(255,255,255,0.44);
                border: 1px solid rgba(90,76,118,0.12);
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.72);
            }

            .side-tip-icon {
                width: 44px;
                height: 44px;
                min-width: 44px;
                display: flex;
                align-items: center;
                justify-content: center;
                color: #D54BFF;
                border: 1px solid rgba(184, 70, 255, 0.55);
                border-radius: 14px;
            }

            .side-tip-text {
                font-size: 0.82rem;
                line-height: 1.48;
                color: rgba(33,26,48,0.82);
                font-weight: 700;
            }

            section[data-testid="stSidebar"] div[role="radiogroup"] {
                gap: 4px !important;
            }

            section[data-testid="stSidebar"] div[role="radiogroup"] > label {
                display: flex !important;
                align-items: center !important;
                margin: 0 !important;
                padding: 7px 0 !important;
                background: transparent !important;
                border: none !important;
                box-shadow: none !important;
            }

            section[data-testid="stSidebar"] div[role="radiogroup"] > label p,
            section[data-testid="stSidebar"] div[role="radiogroup"] > label span {
                color: #241C34 !important;
                font-weight: 800 !important;
            }

            section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover p,
            section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover span {
                color: #7D2DFF !important;
            }

            section[data-testid="stSidebar"] div[role="radiogroup"] > label input {
                accent-color: #FF4BAA !important;
            }

            .page-title {
                color: #FFFFFF;
                font-size: 2.55rem;
                line-height: 1.08;
                font-weight: 950;
                letter-spacing: -0.04em;
            }

            .page-subtitle {
                margin-top: 7px;
                margin-bottom: 16px;
                color: rgba(255,255,255,0.70);
                font-size: 0.94rem;
            }

            .metric-card {
                height: 188px;
                min-height: 188px;
                padding: 17px;
                border-radius: 20px;
                border: 1px solid rgba(255,255,255,0.06);
                background: linear-gradient(145deg, rgba(22,20,42,0.98), rgba(10,9,25,0.98));
                box-shadow: 0 18px 46px rgba(0,0,0,0.22);
                box-sizing: border-box;
                display: flex;
                flex-direction: column;
            }

            .metric-icon {
                width: 44px;
                height: 44px;
                display: flex;
                align-items: center;
                justify-content: center;
                border-radius: 13px;
                color: #FFFFFF;
                font-size: 1.1rem;
                margin-bottom: 14px;
            }

            .metric-label {
                min-height: 38px;
                color: rgba(255,255,255,0.78);
                font-size: 0.94rem;
                font-weight: 750;
                line-height: 1.18;
                display: flex;
                align-items: flex-end;
            }

            .metric-value {
                margin-top: 5px;
                color: #FFFFFF;
                font-size: 1.95rem;
                font-weight: 950;
                line-height: 1;
            }

            .metric-note {
                margin-top: 8px;
                color: #55DF7D;
                font-size: 0.84rem;
                font-weight: 700;
            }

            .section-heading {
                color: #FFFFFF;
                font-size: 1.45rem;
                font-weight: 900;
                margin-bottom: 4px;
            }

            .section-subtitle {
                color: rgba(255,255,255,0.68);
                font-size: 0.92rem;
                margin-bottom: 12px;
            }

            .status-wrap {
                display: flex;
                flex-direction: column;
                gap: 10px;
            }

            .status-row {
                display: grid;
                grid-template-columns: 1fr auto auto;
                align-items: center;
                gap: 12px;
                padding: 11px 12px;
                border-radius: 14px;
                background: rgba(255,255,255,0.04);
                border: 1px solid rgba(255,255,255,0.04);
            }

            .status-left {
                color: #FFFFFF;
                font-size: 0.92rem;
                font-weight: 750;
            }

            .status-count {
                color: #FFFFFF;
                font-weight: 850;
            }

            .status-percent {
                color: rgba(255,255,255,0.66);
                font-weight: 750;
            }

            /* Filtros da visão geral: todos no mesmo estilo escuro, mesma altura e sem borda branca extra */
            div[data-testid="stSelectbox"] > div[data-baseweb="select"] > div,
            div[data-testid="stTextInput"] div[data-baseweb="input"] > div,
            div[data-testid="stDateInput"] div[data-baseweb="input"] > div {
                min-height: 54px !important;
                height: 54px !important;
                border-radius: 15px !important;
                border: 1px solid rgba(255, 75, 170, 0.72) !important;
                background: rgba(8, 7, 24, 0.92) !important;
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.03), 0 10px 30px rgba(0,0,0,0.16) !important;
                color: #FFFFFF !important;
                outline: none !important;
            }

            div[data-testid="stSelectbox"] > div[data-baseweb="select"] > div:hover,
            div[data-testid="stTextInput"] div[data-baseweb="input"] > div:hover,
            div[data-testid="stDateInput"] div[data-baseweb="input"] > div:hover {
                border-color: rgba(255, 75, 170, 0.88) !important;
            }

            div[data-testid="stSelectbox"] > div[data-baseweb="select"] > div:focus-within,
            div[data-testid="stTextInput"] div[data-baseweb="input"] > div:focus-within,
            div[data-testid="stDateInput"] div[data-baseweb="input"] > div:focus-within {
                border-color: rgba(255, 75, 170, 1) !important;
                box-shadow: 0 0 0 1px rgba(255, 75, 170, 0.35), 0 0 22px rgba(169, 28, 255, 0.16) !important;
                background: rgba(8, 7, 24, 0.98) !important;
            }

            /* Remove bordas extras/brancas internas especificamente do Período e Busca */
            div[data-testid="stTextInput"] [data-baseweb="base-input"],
            div[data-testid="stTextInput"] [data-baseweb="input"],
            div[data-testid="stTextInput"] [data-baseweb="input"] > div,
            div[data-testid="stDateInput"] [data-baseweb="base-input"],
            div[data-testid="stDateInput"] [data-baseweb="input"],
            div[data-testid="stDateInput"] [data-baseweb="input"] > div,
            div[data-testid="stDateInput"] > div,
            div[data-testid="stTextInput"] > div {
                border: none !important;
                outline: none !important;
            }

            div[data-testid="stTextInput"] [data-baseweb="input"],
            div[data-testid="stTextInput"] [data-baseweb="input"] > div,
            div[data-testid="stDateInput"] [data-baseweb="input"],
            div[data-testid="stDateInput"] [data-baseweb="input"] > div {
                min-height: 54px !important;
                height: 54px !important;
                box-shadow: none !important;
                box-sizing: border-box !important;
                overflow: visible !important;
            }

            /* Evita corte da borda superior, principalmente no campo Buscar empresa ou telefone */
            div[data-testid="stTextInput"],
            div[data-testid="stDateInput"] {
                padding-top: 3px !important;
                overflow: visible !important;
            }

            div[data-testid="stTextInput"] > div,
            div[data-testid="stDateInput"] > div,
            div[data-testid="stTextInput"] [data-baseweb="base-input"],
            div[data-testid="stDateInput"] [data-baseweb="base-input"] {
                overflow: visible !important;
                box-sizing: border-box !important;
            }

            label {
                color: rgba(255,255,255,0.88) !important;
                font-weight: 700 !important;
            }

            div[data-testid="stSelectbox"] * {
                color: #FFFFFF !important;
            }

            div[data-testid="stTextInput"] [data-baseweb="input"],
            div[data-testid="stTextInput"] [data-baseweb="input"] > div,
            div[data-testid="stTextInput"] input,
            div[data-testid="stDateInput"] [data-baseweb="input"],
            div[data-testid="stDateInput"] [data-baseweb="input"] > div,
            div[data-testid="stDateInput"] input {
                background: transparent !important;
                color: #FFFFFF !important;
                -webkit-text-fill-color: #FFFFFF !important;
                caret-color: #FF4BAA !important;
                min-height: 54px !important;
                height: 54px !important;
                line-height: 54px !important;
                padding-top: 0 !important;
                padding-bottom: 0 !important;
            }

            div[data-testid="stTextInput"] input::placeholder,
            div[data-testid="stDateInput"] input::placeholder {
                color: rgba(255,255,255,0.54) !important;
                -webkit-text-fill-color: rgba(255,255,255,0.54) !important;
            }

            div[data-testid="stDateInput"] button,
            div[data-testid="stDateInput"] svg {
                color: #FFFFFF !important;
                fill: #FFFFFF !important;
            }

            div[data-testid="stTextInput"] input:-webkit-autofill,
            div[data-testid="stTextInput"] input:-webkit-autofill:hover,
            div[data-testid="stTextInput"] input:-webkit-autofill:focus,
            div[data-testid="stDateInput"] input:-webkit-autofill,
            div[data-testid="stDateInput"] input:-webkit-autofill:hover,
            div[data-testid="stDateInput"] input:-webkit-autofill:focus {
                -webkit-box-shadow: 0 0 0 1000px #0B0918 inset !important;
                -webkit-text-fill-color: #FFFFFF !important;
                caret-color: #FF4BAA !important;
            }

            .stButton > button {
                min-height: 48px !important;
                border: none !important;
                border-radius: 15px !important;
                color: #FFFFFF !important;
                font-weight: 800 !important;
                background: linear-gradient(90deg, #FF4BAA 0%, #A91CFF 100%) !important;
            }

            /* Seta para reabrir o menu lateral: cinza e visível em todas as páginas. */
            [data-testid="collapsedControl"],
            [data-testid="stSidebarCollapsedControl"] {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                align-items: center !important;
                justify-content: center !important;
                width: 40px !important;
                height: 40px !important;
                border-radius: 11px !important;
                border: 1px solid rgba(55,65,81,0.20) !important;
                background: #D1D5DB !important;
                box-shadow: 0 8px 20px rgba(0,0,0,0.20) !important;
                z-index: 1000001 !important;
            }

            [data-testid="collapsedControl"] button,
            [data-testid="stSidebarCollapsedControl"] button {
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
                width: 100% !important;
                height: 100% !important;
                padding: 0 !important;
                border: none !important;
                background: transparent !important;
                box-shadow: none !important;
            }

            [data-testid="collapsedControl"] svg,
            [data-testid="stSidebarCollapsedControl"] svg {
                color: #4B5563 !important;
                fill: #4B5563 !important;
                stroke: #4B5563 !important;
                opacity: 1 !important;
            }

            [data-testid="collapsedControl"]:hover,
            [data-testid="stSidebarCollapsedControl"]:hover {
                background: #E5E7EB !important;
                transform: scale(1.06) !important;
            }

            .latest-calls-shell {
                margin-top: 18px;
                margin-bottom: 14px;
                padding: 22px 24px 18px 24px;
                border-radius: 26px;
                background: linear-gradient(145deg, rgba(22,20,42,0.98), rgba(10,9,25,0.98));
                border: 1px solid rgba(255,255,255,0.06);
                box-shadow: 0 18px 46px rgba(0,0,0,0.22);
            }

            .latest-filter-title {
                color: #FFFFFF;
                font-size: 1.08rem;
                font-weight: 900;
                line-height: 1.2;
                margin-bottom: 4px;
            }

            .latest-filter-subtitle {
                color: rgba(255,255,255,0.68);
                font-size: 0.88rem;
                line-height: 1.45;
            }

            .latest-filter-spacer {
                height: 2px;
            }

            .latest-calls-head {
                display: flex;
                align-items: flex-start;
                justify-content: space-between;
                gap: 16px;
                margin-bottom: 4px;
            }

            .latest-calls-title {
                color: #FFFFFF;
                font-size: 1.05rem;
                font-weight: 900;
                line-height: 1.2;
                margin-bottom: 4px;
            }

            .latest-calls-subtitle {
                color: rgba(255,255,255,0.68);
                font-size: 0.88rem;
                line-height: 1.45;
            }

            .latest-calls-chip {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                padding: 8px 14px;
                min-width: 88px;
                border-radius: 999px;
                background: rgba(255, 246, 217, 0.08);
                border: 1px solid rgba(232, 194, 67, 0.92);
                color: #E8C243;
                font-size: 0.78rem;
                font-weight: 900;
                letter-spacing: 0.04em;
                text-transform: uppercase;
            }

            .latest-status-card {
                min-height: 132px;
                height: 132px;
                padding: 14px 12px 12px 12px;
                border-radius: 20px;
                background: linear-gradient(145deg, rgba(22,20,42,0.98), rgba(10,9,25,0.98));
                border: 1px solid rgba(255,255,255,0.06);
                box-shadow: 0 18px 46px rgba(0,0,0,0.22);
                margin-bottom: 12px;
            }

            .latest-status-top {
                display: flex;
                align-items: center;
                gap: 10px;
                margin-bottom: 8px;
            }

            .latest-status-icon {
                width: 36px;
                height: 36px;
                min-width: 36px;
                border-radius: 12px;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 0.95rem;
                font-weight: 900;
            }

            .latest-status-name {
                color: #FFFFFF;
                font-size: 0.82rem;
                font-weight: 850;
                line-height: 1.2;
            }

            .latest-status-number {
                color: #FFFFFF;
                font-size: 1.15rem;
                line-height: 1;
                font-weight: 950;
                margin-top: 4px;
            }

            .latest-status-caption {
                color: #55DF7D;
                font-size: 0.72rem;
                font-weight: 700;
                margin-top: 6px;
            }

            .latest-table-card {
                margin-top: 30px;
                padding: 0;
                border-radius: 26px;
                background: linear-gradient(135deg, rgba(255,75,170,0.96), rgba(169,28,255,0.96));
                box-shadow:
                    0 20px 52px rgba(0,0,0,0.26),
                    0 0 34px rgba(169,28,255,0.18);
            }

            .latest-table-card-inner {
                margin: 1px;
                padding: 18px 18px 16px 18px;
                border-radius: 25px;
                background:
                    radial-gradient(circle at 100% 0%, rgba(169,28,255,0.16), transparent 28%),
                    radial-gradient(circle at 0% 100%, rgba(255,75,170,0.12), transparent 30%),
                    linear-gradient(145deg, rgba(22,20,42,0.99), rgba(10,9,25,0.99));
                border: 1px solid rgba(255,255,255,0.06);
            }

            .latest-table-head {
                display: flex;
                align-items: flex-start;
                justify-content: space-between;
                gap: 14px;
                margin-bottom: 12px;
            }

            .latest-table-title-wrap {
                display: flex;
                align-items: center;
                gap: 12px;
            }

            .latest-table-icon {
                width: 42px;
                height: 42px;
                min-width: 42px;
                display: flex;
                align-items: center;
                justify-content: center;
                border-radius: 14px;
                background: linear-gradient(135deg, rgba(255,75,170,0.22), rgba(169,28,255,0.24));
                border: 1px solid rgba(255,75,170,0.34);
                color: #FF8CCC;
                font-size: 1.05rem;
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.06);
            }

            .latest-table-title {
                color: #FFFFFF;
                font-size: 1.04rem;
                font-weight: 900;
                line-height: 1.2;
                margin-bottom: 4px;
            }

            .latest-table-subtitle {
                color: rgba(255,255,255,0.66);
                font-size: 0.84rem;
                line-height: 1.45;
            }

            .latest-table-badges {
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                justify-content: flex-end;
                gap: 8px;
            }

            .latest-table-badge,
            .latest-table-status-badge {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                padding: 7px 12px;
                border-radius: 999px;
                font-size: 0.76rem;
                font-weight: 900;
                white-space: nowrap;
            }

            /* Ao clicar em “Ver nomes”, exibe somente a planilha editável. */
            div[data-testid="stDataEditor"] {
                margin-top: 24px !important;
            }

            .latest-table-badge {
                background: rgba(255, 246, 217, 0.08);
                border: 1px solid rgba(232, 194, 67, 0.92);
                color: #E8C243;
            }

            .latest-table-status-badge {
                background: rgba(255,75,170,0.10);
                border: 1px solid rgba(255,75,170,0.44);
                color: #FF8CCC;
            }

            .latest-company-fields {
                display: flex;
                flex-wrap: wrap;
                gap: 8px;
                margin-top: 10px;
            }

            .latest-company-field {
                display: inline-flex;
                align-items: center;
                gap: 6px;
                padding: 6px 9px;
                border-radius: 999px;
                background: rgba(255,255,255,0.045);
                border: 1px solid rgba(255,255,255,0.07);
                color: rgba(255,255,255,0.72);
                font-size: 0.72rem;
                font-weight: 750;
            }

            .latest-editor-help {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 16px;
                margin: 14px 0 12px 0;
                padding: 14px 16px;
                border-radius: 16px;
                background:
                    linear-gradient(90deg, rgba(255,75,170,0.10), rgba(169,28,255,0.10)),
                    rgba(13,11,31,0.94);
                border: 1px solid rgba(255,75,170,0.30);
                color: rgba(255,255,255,0.76);
                font-size: 0.84rem;
                line-height: 1.45;
                box-shadow: 0 12px 30px rgba(0,0,0,0.16);
            }

            .latest-editor-help strong {
                color: #FF8CCC;
            }

            .latest-sync-badge {
                display: inline-flex;
                align-items: center;
                gap: 7px;
                flex-shrink: 0;
                padding: 7px 11px;
                border-radius: 999px;
                background: rgba(85,223,125,0.10);
                border: 1px solid rgba(85,223,125,0.38);
                color: #55DF7D;
                font-size: 0.74rem;
                font-weight: 900;
                letter-spacing: 0.02em;
            }

            .latest-status-legend {
                display: flex;
                flex-wrap: wrap;
                gap: 8px;
                margin: 10px 0 14px 0;
            }

            .latest-status-pill {
                display: inline-flex;
                align-items: center;
                gap: 7px;
                padding: 6px 10px;
                border-radius: 999px;
                background: rgba(255,255,255,0.045);
                border: 1px solid rgba(255,255,255,0.08);
                color: rgba(255,255,255,0.86);
                font-size: 0.73rem;
                font-weight: 800;
            }

            .latest-status-dot {
                width: 8px;
                height: 8px;
                border-radius: 50%;
                display: inline-block;
                box-shadow: 0 0 12px currentColor;
            }

            /* Tabela comercial compacta com botão Copiar */
            .premium-inline-table-header {
                margin-top: 4px;
                margin-bottom: 3px;
                padding: 7px 8px;
                border-radius: 9px;
                background: linear-gradient(90deg, rgba(255,75,170,0.15), rgba(169,28,255,0.15));
                border: 1px solid rgba(255,75,170,0.24);
                color: rgba(255,255,255,0.94);
                font-size: 0.73rem;
                font-weight: 850;
            }

            .premium-inline-cell {
                min-height: 30px;
                display: flex;
                align-items: center;
                padding: 4px 7px;
                border-radius: 7px;
                background: rgba(255,255,255,0.97);
                border: 1px solid rgba(169,28,255,0.08);
                color: #261C35;
                font-size: 0.77rem;
                line-height: 1.16;
                word-break: break-word;
            }

            .premium-inline-cell.phone {
                color: #5C2A83;
                font-weight: 850;
            }

            .premium-inline-cell.date {
                justify-content: center;
                color: #5B5369;
                font-size: 0.73rem;
            }

            .premium-inline-cell.muted {
                color: #6E667A;
            }

            .premium-inline-hint {
                margin: 5px 0 7px 0;
                padding: 8px 11px;
                border-radius: 10px;
                background: linear-gradient(90deg, rgba(255,75,170,0.07), rgba(169,28,255,0.07));
                border: 1px solid rgba(255,75,170,0.17);
                color: rgba(255,255,255,0.72);
                font-size: 0.74rem;
                line-height: 1.30;
            }

            .premium-inline-hint strong {
                color: #FF79C4;
            }

            /* Linhas compactas: sem espaços exagerados entre empresas */
            .st-key-compact_inline_table div[data-testid="stHorizontalBlock"] {
                gap: 0.34rem !important;
                margin-bottom: 0 !important;
            }

            .st-key-compact_inline_table div[data-testid="stVerticalBlock"] {
                gap: 0.20rem !important;
            }

            .st-key-compact_inline_table div[data-testid="stElementContainer"] {
                margin-bottom: 0 !important;
            }

            .st-key-compact_inline_table div[data-testid="stSelectbox"] {
                margin-bottom: 0 !important;
            }

            .st-key-compact_inline_table div[data-testid="stSelectbox"] > div[data-baseweb="select"] {
                min-height: 34px !important;
                height: 34px !important;
            }

            .st-key-compact_inline_table div[data-testid="stSelectbox"] > div[data-baseweb="select"] > div {
                min-height: 34px !important;
                height: 34px !important;
                border-radius: 7px !important;
                padding-top: 0 !important;
                padding-bottom: 0 !important;
                display: flex !important;
                align-items: center !important;
                overflow: visible !important;
            }

            .st-key-compact_inline_table div[data-testid="stSelectbox"] span,
            .st-key-compact_inline_table div[data-testid="stSelectbox"] p {
                line-height: 1.15 !important;
                white-space: nowrap !important;
                overflow: visible !important;
                text-overflow: clip !important;
            }

            .st-key-compact_inline_table iframe {
                min-height: 30px !important;
                height: 30px !important;
            }

            /* A tabela não deve ampliar no hover */
            .premium-inline-table-header,
            .premium-inline-table-header:hover,
            .premium-inline-cell,
            .premium-inline-cell:hover,
            .st-key-compact_inline_table,
            .st-key-compact_inline_table * {
                transform: none !important;
                transition:
                    border-color 0.16s ease,
                    background 0.16s ease !important;
            }

            /* Planilha editável: detalhes em rosa e roxo, sem zoom */
            div[data-testid="stDataEditor"] {
                overflow: hidden;
                border-radius: 20px;
                border: 1px solid rgba(255,75,170,0.52);
                background: linear-gradient(180deg, rgba(255,255,255,0.995), rgba(249,247,255,0.995));
                box-shadow:
                    0 18px 44px rgba(0,0,0,0.24),
                    0 0 0 1px rgba(169,28,255,0.10),
                    0 0 32px rgba(169,28,255,0.14);
                transform: none !important;
                transition:
                    border-color 0.22s ease,
                    box-shadow 0.22s ease !important;
            }

            div[data-testid="stDataEditor"] [role="grid"] {
                border-radius: 20px;
                overflow: hidden;
                transform: none !important;
            }

            div[data-testid="stDataEditor"]:hover {
                transform: none !important;
                border-color: rgba(255,75,170,0.80);
                box-shadow:
                    0 20px 48px rgba(0,0,0,0.25),
                    0 0 0 1px rgba(169,28,255,0.22),
                    0 0 36px rgba(255,75,170,0.16) !important;
            }

            div[data-testid="stDataEditor"] * {
                transform: none !important;
            }

            div[data-testid="stDataEditor"] [role="columnheader"] {
                background: linear-gradient(90deg, rgba(255,75,170,0.22), rgba(169,28,255,0.22)) !important;
                border-bottom: 1px solid rgba(169,28,255,0.24) !important;
                color: #2A183E !important;
                font-weight: 900 !important;
                letter-spacing: 0.01em !important;
            }

            div[data-testid="stDataEditor"] [role="gridcell"] {
                border-color: rgba(169,28,255,0.10) !important;
                color: #261C35 !important;
                background: rgba(255,255,255,0.99) !important;
            }

            div[data-testid="stDataEditor"] [role="row"]:nth-child(even) [role="gridcell"] {
                background: rgba(169,28,255,0.045) !important;
            }

            div[data-testid="stDataEditor"] [role="row"]:hover [role="gridcell"] {
                background: linear-gradient(90deg, rgba(255,75,170,0.095), rgba(169,28,255,0.065)) !important;
            }

            div[data-testid="stDataEditor"] [role="gridcell"]:focus,
            div[data-testid="stDataEditor"] [role="gridcell"]:focus-within {
                outline: 2px solid rgba(255,75,170,0.82) !important;
                outline-offset: -2px !important;
                background: rgba(255,75,170,0.10) !important;
            }

            div[data-testid="stDataEditor"] button,
            div[data-testid="stDataEditor"] svg {
                color: #A91CFF !important;
            }

            div[data-testid="stDataFrame"] {
                overflow: hidden;
                border-radius: 16px;
                border: 1px solid rgba(20,16,36,0.10);
                box-shadow: 0 8px 18px rgba(14, 13, 27, 0.04);
            }

            /* Lista de nomes em Todos os cadastros: visual preto, rosa e roxo */
            .contracts-names-count-card {
                margin: 14px 0 14px 0;
                padding: 13px 16px;
                border-radius: 14px;
                color: rgba(38,31,53,0.82);
                font-size: 0.86rem;
                line-height: 1.4;
                background:
                    radial-gradient(circle at 100% 0%, rgba(255,255,255,0.80), transparent 35%),
                    linear-gradient(90deg, rgba(247,248,252,0.98), rgba(220,224,233,0.98));
                border: 1px solid rgba(255,75,170,0.30);
                box-shadow: 0 12px 30px rgba(0,0,0,0.14), 0 0 14px rgba(169,28,255,0.06);
            }

            .contracts-names-count-card strong {
                color: #FF79C4;
                font-weight: 950;
            }

            .contracts-names-table {
                margin-top: 10px;
                overflow: hidden;
                border-radius: 18px;
                border: 1px solid rgba(255,75,170,0.44);
                background:
                    radial-gradient(circle at 100% 0%, rgba(169,28,255,0.14), transparent 34%),
                    linear-gradient(145deg, rgba(13,11,31,0.99), rgba(7,6,18,0.99));
                box-shadow:
                    0 20px 46px rgba(0,0,0,0.30),
                    0 0 0 1px rgba(169,28,255,0.08),
                    0 0 26px rgba(169,28,255,0.12);
            }

            .contracts-names-table-header {
                padding: 14px 18px;
                color: #FFFFFF;
                font-size: 0.90rem;
                font-weight: 950;
                letter-spacing: 0.02em;
                text-transform: uppercase;
                background:
                    linear-gradient(90deg, rgba(255,75,170,0.34), rgba(169,28,255,0.32)),
                    rgba(12,10,28,0.98);
                border-bottom: 1px solid rgba(255,75,170,0.34);
            }

            .contracts-names-table-row {
                padding: 12px 18px;
                color: rgba(255,255,255,0.90);
                font-size: 0.90rem;
                font-weight: 650;
                line-height: 1.25;
                background: rgba(11,10,27,0.96);
                border-bottom: 1px solid rgba(255,75,170,0.10);
                transition:
                    background 0.18s ease,
                    color 0.18s ease,
                    padding-left 0.18s ease,
                    box-shadow 0.18s ease !important;
            }

            .contracts-names-table-row:nth-child(odd) {
                background: rgba(18,13,38,0.97);
            }

            .contracts-names-table-row:last-child {
                border-bottom: none;
            }

            .contracts-names-table-row:hover {
                padding-left: 22px;
                color: #FFFFFF;
                background: linear-gradient(90deg, rgba(255,75,170,0.18), rgba(169,28,255,0.13));
                box-shadow: inset 4px 0 0 #FF4BAA;
            }

            .contracts-filter-summary-grid {
                display: grid;
                grid-template-columns: repeat(7, minmax(0, 1fr));
                gap: 10px;
                margin: 16px 0 4px 0;
            }

            .contracts-filter-summary-card {
                min-height: 110px;
                padding: 13px 12px;
                border-radius: 18px;
                border: 1px solid rgba(255,255,255,0.06);
                background: linear-gradient(145deg, rgba(22,20,42,0.98), rgba(10,9,25,0.98));
                box-shadow: 0 14px 34px rgba(0,0,0,0.20);
                transition: transform 0.20s ease, box-shadow 0.20s ease, border-color 0.20s ease !important;
            }

            .contracts-filter-summary-card:hover {
                transform: scale(1.025);
                border-color: rgba(255,75,170,0.40);
                box-shadow: 0 18px 42px rgba(0,0,0,0.24), 0 0 20px rgba(169,28,255,0.10);
            }

            .contracts-filter-summary-top {
                display: flex;
                align-items: center;
                gap: 8px;
                margin-bottom: 7px;
            }

            .contracts-filter-summary-icon {
                width: 34px;
                height: 34px;
                min-width: 34px;
                border-radius: 11px;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 0.90rem;
                font-weight: 900;
            }

            .contracts-filter-summary-name {
                color: #FFFFFF;
                font-size: 0.77rem;
                font-weight: 850;
                line-height: 1.15;
            }

            .contracts-filter-summary-count {
                color: #FFFFFF;
                font-size: 1.18rem;
                line-height: 1;
                font-weight: 950;
            }

            .contracts-filter-summary-caption {
                margin-top: 7px;
                color: #55DF7D;
                font-size: 0.69rem;
                line-height: 1.25;
                font-weight: 800;
            }

            @media (max-width: 1200px) {
                .contracts-filter-summary-grid {
                    grid-template-columns: repeat(4, minmax(0, 1fr));
                }
            }

            /* Animações suaves de zoom ao passar o mouse */
            .metric-card,
            .latest-calls-shell,
            .latest-status-card,
            .latest-table-card,
            .latest-placeholder-card,
            .status-row,
            .side-tip,
            section[data-testid="stSidebar"] div[role="radiogroup"] > label,
            .stButton > button,
            div[data-testid="stSelectbox"] > div[data-baseweb="select"] > div,
            div[data-testid="stTextInput"] div[data-baseweb="input"] > div,
            div[data-testid="stDateInput"] div[data-baseweb="input"] > div {
                transition:
                    transform 0.22s ease,
                    box-shadow 0.22s ease,
                    border-color 0.22s ease,
                    filter 0.22s ease,
                    background 0.22s ease !important;
                transform-origin: center center;
                will-change: transform;
            }

            .metric-card:hover,
            .latest-calls-shell:hover,
            .latest-status-card:hover,
            .latest-table-card:hover,
            .latest-placeholder-card:hover,
            .status-row:hover,
            .side-tip:hover {
                transform: scale(1.025);
                box-shadow: 0 22px 54px rgba(0,0,0,0.28), 0 0 24px rgba(169, 28, 255, 0.12) !important;
                border-color: rgba(255, 75, 170, 0.34) !important;
                z-index: 3;
            }

            section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
                transform: scale(1.035);
            }

            .stButton > button:hover {
                transform: scale(1.035);
                filter: brightness(1.06);
                box-shadow: 0 14px 30px rgba(169, 28, 255, 0.28) !important;
            }

            div[data-testid="stSelectbox"] > div[data-baseweb="select"] > div:hover,
            div[data-testid="stTextInput"] div[data-baseweb="input"] > div:hover,
            div[data-testid="stDateInput"] div[data-baseweb="input"] > div:hover {
                transform: scale(1.018);
            }

            /* A tabela é a única área sem animação de zoom */
            div[data-testid="stDataEditor"],
            div[data-testid="stDataEditor"]:hover,
            div[data-testid="stDataEditor"] *,
            div[data-testid="stDataEditor"] *:hover,
            div[data-testid="stDataFrame"],
            div[data-testid="stDataFrame"]:hover,
            div[data-testid="stDataFrame"] *,
            div[data-testid="stDataFrame"] *:hover {
                transform: none !important;
                will-change: auto !important;
            }

            /* Menu lateral preservado com submenu flutuante lateral no Cadastro */
            .oppi-side-nav {
                display: flex;
                flex-direction: column;
                gap: 4px;
                margin: 2px 0 14px 0;
                overflow: visible !important;
            }

            .oppi-nav-link,
            .oppi-nav-summary {
                min-height: 38px;
                display: flex;
                align-items: center;
                gap: 10px;
                width: 100%;
                padding: 7px 0;
                color: #241C34 !important;
                text-decoration: none !important;
                font-size: 0.93rem;
                font-weight: 800;
                line-height: 1;
                cursor: pointer;
                list-style: none;
                transition: color 0.16s ease, transform 0.16s ease;
            }

            .oppi-nav-link:hover,
            .oppi-nav-summary:hover {
                color: #7D2DFF !important;
                transform: translateX(3px);
            }

            .oppi-nav-summary::-webkit-details-marker {
                display: none;
            }

            .oppi-nav-dot {
                width: 16px;
                height: 16px;
                min-width: 16px;
                border-radius: 50%;
                border: 1px solid rgba(70, 62, 90, 0.34);
                background: rgba(255,255,255,0.84);
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.84);
            }

            .oppi-nav-link.active .oppi-nav-dot,
            .oppi-cadastro-details.active .oppi-nav-dot {
                border: 4px solid #FF5C64;
                background: #FFFFFF;
                box-shadow: none;
            }

            .oppi-nav-arrow {
                margin-left: auto;
                padding-right: 8px;
                color: #241C34;
                font-size: 1.35rem;
                font-weight: 950;
                line-height: 1;
            }

            .oppi-cadastro-details {
                position: relative;
                overflow: visible !important;
            }

            .oppi-cadastro-flyout {
                display: none;
                position: fixed;
                left: 304px;
                top: 222px;
                width: 275px;
                z-index: 999999;
                overflow: hidden;
                border-radius: 0 12px 12px 0;
                border: 1px solid rgba(63,53,83,0.16);
                border-left: 4px solid #A91CFF;
                background:
                    radial-gradient(circle at top left, rgba(255,255,255,0.96) 0%, rgba(255,255,255,0.00) 34%),
                    radial-gradient(circle at bottom right, rgba(208,212,223,0.78) 0%, rgba(208,212,223,0.00) 40%),
                    linear-gradient(180deg, #F7F8FC 0%, #ECEEF4 42%, #DCE0E9 100%);
                box-shadow: 0 22px 48px rgba(0,0,0,0.22), 0 0 18px rgba(169,28,255,0.10);
            }

            .oppi-cadastro-details[open] .oppi-cadastro-flyout {
                display: block;
            }

            .oppi-flyout-title {
                padding: 16px 18px 13px 18px;
                color: #241C34;
                font-size: 1rem;
                font-weight: 900;
                border-bottom: 1px solid rgba(63,53,83,0.14);
                background: rgba(255,255,255,0.46);
            }

            .oppi-flyout-link {
                display: block;
                min-height: 46px;
                padding: 14px 18px 12px 18px;
                color: #241C34 !important;
                text-decoration: none !important;
                font-size: 0.88rem;
                font-weight: 700;
                background: transparent;
                transition: background 0.16s ease, padding-left 0.16s ease, color 0.16s ease;
            }

            .oppi-flyout-link:hover,
            .oppi-flyout-link.active {
                padding-left: 22px;
                color: #5F1DB8 !important;
                background: linear-gradient(90deg, rgba(255,75,170,0.14), rgba(169,28,255,0.13));
                box-shadow: inset 3px 0 0 #A91CFF;
            }

            @media (max-width: 900px) {
                .oppi-cadastro-flyout {
                    left: 286px;
                    top: 210px;
                    width: 245px;
                }
            }

            @media (prefers-reduced-motion: reduce) {
                .metric-card,
                .latest-calls-shell,
                .latest-status-card,
                .latest-table-card,
                .latest-placeholder-card,
                .status-row,
                .side-tip,
                section[data-testid="stSidebar"] div[role="radiogroup"] > label,
                .stButton > button,
                div[data-testid="stSelectbox"] > div[data-baseweb="select"] > div,
                div[data-testid="stTextInput"] div[data-baseweb="input"] > div,
                div[data-testid="stDateInput"] div[data-baseweb="input"] > div {
                    transition: none !important;
                    transform: none !important;
                }
            }
        </style>
        """
    )


def apply_registration_css() -> None:
    render_html(
        """
        <style>
            .registration-header-card {
                margin-bottom: 18px;
                padding: 24px 26px 22px 26px;
                border-radius: 24px;
                background:
                    radial-gradient(circle at 100% 0%, rgba(169,28,255,0.20), transparent 32%),
                    radial-gradient(circle at 0% 100%, rgba(255,75,170,0.12), transparent 34%),
                    linear-gradient(145deg, rgba(22,20,42,0.99), rgba(10,9,25,0.99));
                border: 1px solid rgba(255,75,170,0.32);
                box-shadow: 0 18px 46px rgba(0,0,0,0.22), 0 0 26px rgba(169,28,255,0.10);
            }

            .registration-kicker {
                color: #FF79C4;
                font-size: 0.76rem;
                font-weight: 900;
                letter-spacing: 0.16em;
                text-transform: uppercase;
                margin-bottom: 10px;
            }

            .registration-title {
                color: #FFFFFF;
                font-size: 2rem;
                font-weight: 950;
                letter-spacing: -0.035em;
                line-height: 1.05;
            }

            .registration-subtitle {
                margin-top: 9px;
                color: rgba(255,255,255,0.68);
                font-size: 0.94rem;
                line-height: 1.45;
            }

            .registration-section {
                margin: 8px 0 14px 0;
                padding: 14px 16px;
                border-radius: 16px;
                background: rgba(255,255,255,0.96);
                border: 1px solid rgba(255,75,170,0.38);
                box-shadow:
                    0 8px 18px rgba(169,28,255,0.08),
                    inset 0 1px 0 rgba(255,255,255,0.92);
            }

            .registration-section-title {
                color: #1E1729;
                font-size: 0.95rem;
                font-weight: 900;
                letter-spacing: 0.015em;
            }

            .registration-section-text {
                margin-top: 4px;
                color: rgba(30,23,41,0.72);
                font-size: 0.80rem;
                line-height: 1.4;
            }

            .registration-note {
                margin: 0 0 14px 0;
                padding: 12px 14px;
                border-radius: 14px;
                color: rgba(38,31,53,0.78);
                background: rgba(255,255,255,0.48);
                border: 1px solid rgba(90,76,118,0.12);
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.68);
                font-size: 0.83rem;
                line-height: 1.45;
            }

            .registration-note strong {
                color: #FF79C4;
            }

            [data-testid="stForm"] {
                padding: 20px !important;
                border-radius: 24px !important;
                background:
                    radial-gradient(circle at top left, rgba(255,255,255,0.92) 0%, rgba(255,255,255,0.00) 32%),
                    radial-gradient(circle at bottom right, rgba(208,212,223,0.72) 0%, rgba(208,212,223,0.00) 34%),
                    linear-gradient(180deg, #F7F8FC 0%, #ECEEF4 38%, #DCE0E9 72%, #CED3DE 100%) !important;
                border: 1px solid rgba(63,53,83,0.14) !important;
                box-shadow:
                    0 18px 46px rgba(0,0,0,0.18),
                    inset 0 1px 0 rgba(255,255,255,0.80) !important;
                overflow: visible !important;
            }

            [data-testid="stForm"] form,
            [data-testid="stForm"] div[data-testid="stVerticalBlock"],
            [data-testid="stForm"] div[data-testid="column"],
            [data-testid="stForm"] div[data-testid="stElementContainer"],
            [data-testid="stForm"] div[data-testid="stTextInput"],
            [data-testid="stForm"] div[data-testid="stTextArea"],
            [data-testid="stForm"] div[data-testid="stSelectbox"],
            [data-testid="stForm"] div[data-testid="stDateInput"] {
                overflow: visible !important;
            }

            [data-testid="stForm"] div[data-testid="stTextInput"],
            [data-testid="stForm"] div[data-testid="stTextArea"],
            [data-testid="stForm"] div[data-testid="stSelectbox"],
            [data-testid="stForm"] div[data-testid="stDateInput"] {
                padding: 4px 0 !important;
                position: relative !important;
            }

            [data-testid="stForm"] label {
                color: #1E1828 !important;
                -webkit-text-fill-color: #1E1828 !important;
                font-size: 0.86rem !important;
                font-weight: 800 !important;
            }

            [data-testid="stForm"] div[data-baseweb="input"] > div,
            [data-testid="stForm"] div[data-baseweb="select"] > div {
                min-height: 48px !important;
                border-radius: 13px !important;
                border: 1px solid rgba(255,75,170,0.58) !important;
                background: #FFFFFF !important;
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.88) !important;
                color: #1E1828 !important;
                transition:
                    transform 0.20s ease,
                    box-shadow 0.20s ease,
                    border-color 0.20s ease !important;
                transform-origin: center center;
                position: relative !important;
                z-index: 1 !important;
            }

            [data-testid="stForm"] div[data-baseweb="input"] > div:hover,
            [data-testid="stForm"] div[data-baseweb="select"] > div:hover {
                transform: scale(1.012) !important;
                border-color: rgba(255,75,170,0.82) !important;
                box-shadow:
                    0 0 0 1px rgba(255,75,170,0.12),
                    0 0 14px rgba(169,28,255,0.12),
                    0 8px 18px rgba(169,28,255,0.10),
                    inset 0 1px 0 rgba(255,255,255,0.03) !important;
                z-index: 8 !important;
            }

            [data-testid="stForm"] div[data-testid="stTextInput"]:hover,
            [data-testid="stForm"] div[data-testid="stSelectbox"]:hover,
            [data-testid="stForm"] div[data-testid="stDateInput"]:hover {
                z-index: 15 !important;
            }

            [data-testid="stForm"] div[data-baseweb="textarea"] {
                border: none !important;
                background: transparent !important;
                box-shadow: none !important;
            }

            [data-testid="stForm"] div[data-testid="stTextArea"]:hover {
                z-index: 15 !important;
            }

            [data-testid="stForm"] div[data-testid="stTextArea"] > div,
            [data-testid="stForm"] div[data-testid="stTextArea"] > div > div {
                overflow: visible !important;
                border-radius: 13px !important;
                background: transparent !important;
                box-shadow: none !important;
            }

            [data-testid="stForm"] div[data-baseweb="textarea"] > div {
                min-height: 124px !important;
                border-radius: 13px !important;
                border: 1px solid rgba(255,75,170,0.58) !important;
                background: #FFFFFF !important;
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.88) !important;
                transition:
                    transform 0.18s ease,
                    box-shadow 0.18s ease,
                    border-color 0.18s ease !important;
                transform-origin: center center;
                position: relative !important;
                z-index: 1 !important;
                overflow: visible !important;
            }

            [data-testid="stForm"] div[data-baseweb="textarea"] > div:hover,
            [data-testid="stForm"] div[data-baseweb="textarea"] > div:focus-within {
                transform: scale(1.012) !important;
                border-color: rgba(255,75,170,0.82) !important;
                box-shadow:
                    0 0 0 1px rgba(255,75,170,0.12),
                    0 0 14px rgba(169,28,255,0.12),
                    0 8px 18px rgba(169,28,255,0.10),
                    inset 0 1px 0 rgba(255,255,255,0.03) !important;
                z-index: 8 !important;
            }

            [data-testid="stForm"] div[data-testid="stTextArea"] textarea,
            [data-testid="stForm"] div[data-baseweb="textarea"] textarea {
                min-height: 124px !important;
                border: none !important;
                border-radius: 13px !important;
                outline: none !important;
                background: transparent !important;
                box-shadow: none !important;
                color: #1E1828 !important;
                position: relative !important;
                z-index: 1 !important;
                resize: vertical !important;
            }

            [data-testid="stForm"] div[data-baseweb="textarea"] > div:focus-within,
            [data-testid="stForm"] div[data-baseweb="input"] > div:focus-within,
            [data-testid="stForm"] div[data-baseweb="select"] > div:focus-within {
                border: 1px solid rgba(255,75,170,0.88) !important;
                box-shadow:
                    0 0 0 1px rgba(255,75,170,0.14),
                    0 0 16px rgba(169,28,255,0.14),
                    0 8px 18px rgba(169,28,255,0.10) !important;
                z-index: 10 !important;
            }

            /* Labels dos campos em preto para permanecerem legíveis sobre o fundo cinza */
            [data-testid="stForm"] div[data-testid="stSelectbox"] label,
            [data-testid="stForm"] div[data-testid="stSelectbox"] label p,
            [data-testid="stForm"] div[data-testid="stSelectbox"] label span,
            [data-testid="stForm"] div[data-testid="stSelectbox"] [data-testid="stWidgetLabel"],
            [data-testid="stForm"] div[data-testid="stSelectbox"] [data-testid="stWidgetLabel"] p {
                color: #1E1828 !important;
                -webkit-text-fill-color: #1E1828 !important;
            }

            /* Campos do formulário: caixas brancas e texto preto sobre o fundo cinza */
            [data-testid="stForm"] div[data-baseweb="input"] input,
            [data-testid="stForm"] div[data-baseweb="select"] *,
            [data-testid="stForm"] div[data-testid="stDateInput"] input {
                color: #1E1828 !important;
                -webkit-text-fill-color: #1E1828 !important;
            }

            [data-testid="stForm"] div[data-baseweb="input"] input::placeholder,
            [data-testid="stForm"] div[data-testid="stDateInput"] input::placeholder {
                color: rgba(30,24,40,0.52) !important;
                -webkit-text-fill-color: rgba(30,24,40,0.52) !important;
            }

            [data-testid="stForm"] div[data-testid="stTextArea"] div[data-baseweb="textarea"] > div {
                background: #FFFFFF !important;
            }

            [data-testid="stForm"] div[data-testid="stTextArea"] textarea {
                background: transparent !important;
                color: #1E1828 !important;
                -webkit-text-fill-color: #1E1828 !important;
            }

            [data-testid="stForm"] div[data-testid="stTextArea"] textarea::placeholder {
                color: rgba(30,24,40,0.52) !important;
                -webkit-text-fill-color: rgba(30,24,40,0.52) !important;
            }


            /* Lista clicável das empresas cadastradas com fundo cinza igual ao menu */
            .st-key-contracts_names_list {
                margin-top: 10px !important;
                overflow: hidden !important;
                border-radius: 18px !important;
                border: 1px solid rgba(255,75,170,0.34) !important;
                background:
                    radial-gradient(circle at top left, rgba(255,255,255,0.94) 0%, rgba(255,255,255,0.00) 34%),
                    radial-gradient(circle at bottom right, rgba(208,212,223,0.72) 0%, rgba(208,212,223,0.00) 36%),
                    linear-gradient(180deg, #F7F8FC 0%, #ECEEF4 42%, #DCE0E9 72%, #CED3DE 100%) !important;
                box-shadow:
                    0 18px 42px rgba(0,0,0,0.18),
                    0 0 0 1px rgba(169,28,255,0.08),
                    0 0 18px rgba(169,28,255,0.08) !important;
            }

            .st-key-contracts_names_list div[data-testid="stVerticalBlock"] {
                gap: 0 !important;
            }

            .st-key-contracts_names_list div[data-testid="stElementContainer"] {
                margin: 0 !important;
                padding: 0 !important;
            }

            .st-key-contracts_names_list .stButton > button {
                width: 100% !important;
                min-height: 45px !important;
                margin: 0 !important;
                padding: 11px 18px !important;
                justify-content: flex-start !important;
                border: none !important;
                border-bottom: 1px solid rgba(255,75,170,0.10) !important;
                border-radius: 0 !important;
                color: #211A30 !important;
                background: rgba(255,255,255,0.92) !important;
                box-shadow: none !important;
                font-size: 0.90rem !important;
                font-weight: 700 !important;
                line-height: 1.25 !important;
                text-align: left !important;
                transition:
                    background 0.18s ease,
                    color 0.18s ease,
                    padding-left 0.18s ease,
                    box-shadow 0.18s ease !important;
            }

            .st-key-contracts_names_list div[data-testid="stElementContainer"]:nth-child(even) .stButton > button {
                background: rgba(232,235,242,0.96) !important;
            }

            .st-key-contracts_names_list .stButton > button:hover {
                transform: none !important;
                padding-left: 22px !important;
                color: #211A30 !important;
                background: linear-gradient(90deg, rgba(255,75,170,0.16), rgba(169,28,255,0.11), rgba(255,255,255,0.94)) !important;
                box-shadow: inset 4px 0 0 #FF4BAA, 0 0 20px rgba(169,28,255,0.12) !important;
            }

            /* Mantém a lista igual à tabela aprovada: linhas juntas e nomes alinhados à esquerda */
            .st-key-contracts_names_list,
            .st-key-contracts_names_list > div,
            .st-key-contracts_names_list div[data-testid="stVerticalBlock"],
            .st-key-contracts_names_list div[data-testid="stElementContainer"],
            .st-key-contracts_names_list .stButton {
                margin-top: 0 !important;
                margin-bottom: 0 !important;
                padding-top: 0 !important;
                padding-bottom: 0 !important;
                gap: 0 !important;
                row-gap: 0 !important;
            }

            .st-key-contracts_names_list .stButton > button {
                min-height: 45px !important;
                height: 45px !important;
                display: flex !important;
                align-items: center !important;
                justify-content: flex-start !important;
                padding: 0 18px !important;
                margin: 0 !important;
                border-radius: 0 !important;
            }

            .st-key-contracts_names_list .stButton > button p,
            .st-key-contracts_names_list .stButton > button span,
            .st-key-contracts_names_list .stButton > button div {
                width: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
                text-align: left !important;
                justify-content: flex-start !important;
                line-height: 1.15 !important;
            }

            .contracts-names-clickable-header {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 12px;
                padding: 10px 12px 10px 18px;
                color: #FFFFFF;
                font-size: 0.90rem;
                font-weight: 950;
                letter-spacing: 0.02em;
                text-transform: uppercase;
                background:
                    linear-gradient(90deg, rgba(255,75,170,0.42), rgba(169,28,255,0.42)),
                    rgba(20,15,43,0.98);
                border-bottom: 1px solid rgba(255,75,170,0.30);
            }

            .contracts-names-sort-toggle {
                width: 34px;
                height: 30px;
                display: inline-flex;
                align-items: center;
                justify-content: center;
                flex-shrink: 0;
                border-radius: 10px;
                color: #FFFFFF !important;
                text-decoration: none !important;
                font-size: 1.12rem;
                font-weight: 950;
                line-height: 1;
                background: linear-gradient(135deg, rgba(255,75,170,0.92), rgba(169,28,255,0.92));
                border: 1px solid rgba(255,255,255,0.24);
                box-shadow: 0 8px 18px rgba(169,28,255,0.18);
                transition: transform 0.18s ease, filter 0.18s ease, box-shadow 0.18s ease;
            }

            .contracts-names-sort-toggle:hover {
                transform: scale(1.08);
                filter: brightness(1.08);
                box-shadow: 0 10px 22px rgba(169,28,255,0.28);
            }

            /* Página de visualização do cadastro preenchido */
            .contract-detail-shell {
                padding: 20px !important;
                border-radius: 24px !important;
                background:
                    radial-gradient(circle at top left, rgba(255,255,255,0.92) 0%, rgba(255,255,255,0.00) 32%),
                    radial-gradient(circle at bottom right, rgba(208,212,223,0.72) 0%, rgba(208,212,223,0.00) 34%),
                    linear-gradient(180deg, #F7F8FC 0%, #ECEEF4 38%, #DCE0E9 72%, #CED3DE 100%) !important;
                border: 1px solid rgba(63,53,83,0.14) !important;
                box-shadow:
                    0 18px 46px rgba(0,0,0,0.18),
                    inset 0 1px 0 rgba(255,255,255,0.80) !important;
            }

            .contract-detail-grid {
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 14px 16px;
                margin-bottom: 16px;
            }

            .contract-detail-grid.three-columns {
                grid-template-columns: repeat(3, minmax(0, 1fr));
            }

            .contract-detail-field {
                min-width: 0;
            }

            .contract-detail-field.full-width {
                grid-column: 1 / -1;
            }

            .contract-detail-label {
                margin-bottom: 7px;
                color: #1E1828;
                font-size: 0.84rem;
                font-weight: 780;
            }

            .contract-detail-value {
                min-height: 48px;
                display: flex;
                align-items: center;
                padding: 12px 14px;
                border-radius: 13px;
                border: 1px solid rgba(255,75,170,0.58);
                background: #FFFFFF;
                color: #1E1828;
                font-size: 0.91rem;
                line-height: 1.35;
                word-break: break-word;
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.03);
                transition: transform 0.18s ease, border-color 0.18s ease, box-shadow 0.18s ease;
            }

            .contract-detail-value:hover {
                transform: scale(1.012);
                border-color: rgba(255,75,170,0.82);
                box-shadow:
                    0 0 0 1px rgba(255,75,170,0.12),
                    0 0 14px rgba(169,28,255,0.12),
                    0 8px 18px rgba(169,28,255,0.10),
                    inset 0 1px 0 rgba(255,255,255,0.03);
            }

            .contract-detail-value.long-text {
                min-height: 94px;
                align-items: flex-start;
                white-space: pre-wrap;
            }

            .st-key-contract_detail_back .stButton > button {
                width: auto !important;
                min-height: 42px !important;
                margin-bottom: 14px !important;
                padding: 0 18px !important;
                border-radius: 13px !important;
            }

            /* Botão Editar dados posicionado dentro do bloco DADOS DA EMPRESA */
            .st-key-contract_detail_edit_inline {
                position: relative !important;
                z-index: 30 !important;
                height: 0 !important;
                min-height: 0 !important;
                margin: 0 !important;
                padding: 0 36px 0 0 !important;
                overflow: visible !important;
            }

            .st-key-contract_detail_edit_inline div[data-testid="stHorizontalBlock"] {
                position: relative !important;
                z-index: 31 !important;
                height: 0 !important;
                min-height: 0 !important;
                margin: 0 !important;
                padding: 0 !important;
                overflow: visible !important;
            }

            .st-key-contract_detail_edit_inline .stButton {
                position: relative !important;
                z-index: 32 !important;
                height: 0 !important;
                min-height: 0 !important;
                margin: 0 !important;
                padding: 0 !important;
                overflow: visible !important;
            }

            .st-key-contract_detail_edit_inline .stButton > button {
                width: 100% !important;
                min-height: 42px !important;
                height: 42px !important;
                margin: 0 !important;
                padding: 0 18px !important;
                border-radius: 13px !important;
                transform: translateY(62px) !important;
                box-shadow: 0 10px 22px rgba(169,28,255,0.18) !important;
            }

            .st-key-contract_detail_edit_inline .stButton > button:hover {
                transform: translateY(62px) scale(1.035) !important;
            }

            @media (max-width: 900px) {
                .st-key-contract_detail_edit_inline {
                    padding-right: 24px !important;
                }

                .st-key-contract_detail_edit_inline .stButton > button {
                    transform: translateY(62px) !important;
                    min-height: 38px !important;
                    height: 38px !important;
                    padding: 0 10px !important;
                    font-size: 0.80rem !important;
                }

                .st-key-contract_detail_edit_inline .stButton > button:hover {
                    transform: translateY(62px) scale(1.025) !important;
                }
            }

            @media (max-width: 900px) {
                .contract-detail-grid,
                .contract-detail-grid.three-columns {
                    grid-template-columns: 1fr;
                }
            }

            [data-testid="stForm"] input,
            [data-testid="stForm"] textarea,
            [data-testid="stForm"] div[data-baseweb="select"] * {
                color: #FFFFFF !important;
                -webkit-text-fill-color: #FFFFFF !important;
            }

            [data-testid="stForm"] input::placeholder,
            [data-testid="stForm"] textarea::placeholder {
                color: rgba(255,255,255,0.44) !important;
                -webkit-text-fill-color: rgba(255,255,255,0.44) !important;
            }

            [data-testid="stForm"] .stButton > button,
            [data-testid="stForm"] button[kind="secondaryFormSubmit"] {
                width: 100% !important;
                min-height: 52px !important;
                margin-top: 10px !important;
                border: none !important;
                border-radius: 14px !important;
                color: #FFFFFF !important;
                font-size: 0.96rem !important;
                font-weight: 900 !important;
                background: linear-gradient(90deg, #FF4BAA 0%, #D73AFF 54%, #8C2BFF 100%) !important;
                box-shadow: 0 14px 30px rgba(169,28,255,0.26) !important;
            }

            .registration-required {
                color: #FF79C4;
                font-weight: 900;
            }

            /* CORREÇÃO FINAL: campos do Novo cadastro brancos com texto preto.
               Este bloco fica por último para não ser sobrescrito pelo CSS anterior. */
            [data-testid="stForm"] label,
            [data-testid="stForm"] label p,
            [data-testid="stForm"] label span,
            [data-testid="stForm"] [data-testid="stWidgetLabel"],
            [data-testid="stForm"] [data-testid="stWidgetLabel"] p,
            [data-testid="stForm"] [data-testid="stWidgetLabel"] span {
                color: #1E1828 !important;
                -webkit-text-fill-color: #1E1828 !important;
            }

            [data-testid="stForm"] div[data-baseweb="input"],
            [data-testid="stForm"] div[data-baseweb="input"] > div,
            [data-testid="stForm"] div[data-baseweb="base-input"],
            [data-testid="stForm"] div[data-testid="stDateInput"] div[data-baseweb="input"],
            [data-testid="stForm"] div[data-testid="stDateInput"] div[data-baseweb="input"] > div,
            [data-testid="stForm"] div[data-baseweb="select"] > div,
            [data-testid="stForm"] div[data-baseweb="textarea"],
            [data-testid="stForm"] div[data-baseweb="textarea"] > div {
                background: #FFFFFF !important;
                color: #1E1828 !important;
                -webkit-text-fill-color: #1E1828 !important;
            }

            [data-testid="stForm"] input,
            [data-testid="stForm"] textarea,
            [data-testid="stForm"] div[data-baseweb="select"] *,
            [data-testid="stForm"] div[data-baseweb="select"] span,
            [data-testid="stForm"] div[data-baseweb="select"] div {
                background: transparent !important;
                color: #1E1828 !important;
                -webkit-text-fill-color: #1E1828 !important;
                caret-color: #1E1828 !important;
            }

            [data-testid="stForm"] input::placeholder,
            [data-testid="stForm"] textarea::placeholder {
                color: rgba(30,24,40,0.52) !important;
                -webkit-text-fill-color: rgba(30,24,40,0.52) !important;
            }

            [data-testid="stForm"] div[data-baseweb="input"] > div,
            [data-testid="stForm"] div[data-baseweb="select"] > div,
            [data-testid="stForm"] div[data-baseweb="textarea"] > div {
                border: 1px solid rgba(255,75,170,0.58) !important;
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.88) !important;
            }

            /* CORREÇÃO DEFINITIVA DOS CAMPOS DO FORMULÁRIO:
               caixas totalmente brancas e conteúdo em preto. */
            [data-testid="stForm"] div[data-testid="stTextInput"],
            [data-testid="stForm"] div[data-testid="stDateInput"],
            [data-testid="stForm"] div[data-testid="stSelectbox"],
            [data-testid="stForm"] div[data-testid="stTextArea"] {
                color-scheme: light !important;
            }

            [data-testid="stForm"] div[data-testid="stTextInput"] div[data-baseweb="input"],
            [data-testid="stForm"] div[data-testid="stTextInput"] div[data-baseweb="input"] > div,
            [data-testid="stForm"] div[data-testid="stTextInput"] div[data-baseweb="base-input"],
            [data-testid="stForm"] div[data-testid="stDateInput"] div[data-baseweb="input"],
            [data-testid="stForm"] div[data-testid="stDateInput"] div[data-baseweb="input"] > div,
            [data-testid="stForm"] div[data-testid="stDateInput"] div[data-baseweb="base-input"],
            [data-testid="stForm"] div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
            [data-testid="stForm"] div[data-testid="stTextArea"] div[data-baseweb="textarea"],
            [data-testid="stForm"] div[data-testid="stTextArea"] div[data-baseweb="textarea"] > div {
                background: #FFFFFF !important;
                background-color: #FFFFFF !important;
                color: #111111 !important;
                -webkit-text-fill-color: #111111 !important;
            }

            [data-testid="stForm"] div[data-testid="stTextInput"] input,
            [data-testid="stForm"] div[data-testid="stDateInput"] input,
            [data-testid="stForm"] div[data-testid="stTextArea"] textarea {
                background: #FFFFFF !important;
                background-color: #FFFFFF !important;
                color: #111111 !important;
                -webkit-text-fill-color: #111111 !important;
                caret-color: #111111 !important;
            }

            [data-testid="stForm"] div[data-testid="stSelectbox"] div[role="combobox"],
            [data-testid="stForm"] div[data-testid="stSelectbox"] div[role="combobox"] *,
            [data-testid="stForm"] div[data-testid="stSelectbox"] span,
            [data-testid="stForm"] div[data-testid="stSelectbox"] svg,
            [data-testid="stForm"] div[data-testid="stDateInput"] svg {
                color: #111111 !important;
                fill: #111111 !important;
                -webkit-text-fill-color: #111111 !important;
            }

            [data-testid="stForm"] div[data-testid="stTextInput"] input::placeholder,
            [data-testid="stForm"] div[data-testid="stDateInput"] input::placeholder,
            [data-testid="stForm"] div[data-testid="stTextArea"] textarea::placeholder {
                color: rgba(17,17,17,0.48) !important;
                -webkit-text-fill-color: rgba(17,17,17,0.48) !important;
                opacity: 1 !important;
            }

            [data-testid="stForm"] div[data-testid="stTextInput"] input:-webkit-autofill,
            [data-testid="stForm"] div[data-testid="stTextInput"] input:-webkit-autofill:hover,
            [data-testid="stForm"] div[data-testid="stTextInput"] input:-webkit-autofill:focus,
            [data-testid="stForm"] div[data-testid="stDateInput"] input:-webkit-autofill,
            [data-testid="stForm"] div[data-testid="stDateInput"] input:-webkit-autofill:hover,
            [data-testid="stForm"] div[data-testid="stDateInput"] input:-webkit-autofill:focus {
                -webkit-box-shadow: 0 0 0 1000px #FFFFFF inset !important;
                -webkit-text-fill-color: #111111 !important;
                caret-color: #111111 !important;
            }

            [data-testid="stForm"] div[data-testid="stTextInput"] div[data-baseweb="input"] > div,
            [data-testid="stForm"] div[data-testid="stDateInput"] div[data-baseweb="input"] > div,
            [data-testid="stForm"] div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
            [data-testid="stForm"] div[data-testid="stTextArea"] div[data-baseweb="textarea"] > div {
                border: 1px solid rgba(255,75,170,0.72) !important;
                border-radius: 20px !important;
                box-shadow: inset 0 1px 0 rgba(255,255,255,0.96) !important;
                overflow: hidden !important;
            }

            [data-testid="stForm"] div[data-testid="stTextInput"] div[data-baseweb="input"],
            [data-testid="stForm"] div[data-testid="stDateInput"] div[data-baseweb="input"],
            [data-testid="stForm"] div[data-testid="stSelectbox"] div[data-baseweb="select"],
            [data-testid="stForm"] div[data-testid="stTextArea"] div[data-baseweb="textarea"],
            [data-testid="stForm"] div[data-testid="stTextInput"] div[data-baseweb="base-input"],
            [data-testid="stForm"] div[data-testid="stDateInput"] div[data-baseweb="base-input"] {
                border-radius: 20px !important;
                overflow: hidden !important;
            }

            [data-testid="stForm"] div[data-testid="stTextInput"] input,
            [data-testid="stForm"] div[data-testid="stDateInput"] input,
            [data-testid="stForm"] div[data-testid="stTextArea"] textarea,
            [data-testid="stForm"] div[data-testid="stSelectbox"] div[role="combobox"] {
                border-radius: 20px !important;
            }


            /* Todos os cadastros: mantém os seis filtros rigorosamente alinhados na mesma linha. */
            .st-key-contracts_filters_aligned div[data-testid="stHorizontalBlock"] {
                align-items: flex-start !important;
            }

            .st-key-contracts_filters_aligned div[data-testid="stTextInput"],
            .st-key-contracts_filters_aligned div[data-testid="stDateInput"],
            .st-key-contracts_filters_aligned div[data-testid="stSelectbox"] {
                padding-top: 0 !important;
                margin-top: 0 !important;
            }

            .st-key-contracts_filters_aligned div[data-testid="stTextInput"] > div,
            .st-key-contracts_filters_aligned div[data-testid="stDateInput"] > div,
            .st-key-contracts_filters_aligned div[data-testid="stSelectbox"] > div {
                margin-top: 0 !important;
            }

            .st-key-contracts_filters_aligned div[data-testid="stTextInput"] div[data-baseweb="input"] > div,
            .st-key-contracts_filters_aligned div[data-testid="stDateInput"] div[data-baseweb="input"] > div,
            .st-key-contracts_filters_aligned div[data-testid="stSelectbox"] > div[data-baseweb="select"] > div {
                min-height: 54px !important;
                height: 54px !important;
                box-sizing: border-box !important;
            }

            /* Todos os cadastros: devolve a respiração visual sem alterar o alinhamento. */
            .contracts-page-header-gap {
                height: 18px;
            }

            .st-key-contracts_filters_aligned {
                margin-bottom: 22px !important;
            }

            .contracts-names-count-card {
                margin: 0 0 18px 0 !important;
            }

            .st-key-contracts_names_list {
                margin-top: 0 !important;
            }

            /* Visão Geral: seis filtros alinhados e com respiro entre filtros e cards. */
            .st-key-overview_filters_aligned {
                margin-bottom: 22px !important;
            }

            .st-key-overview_filters_aligned div[data-testid="stHorizontalBlock"] {
                align-items: flex-start !important;
                gap: 0.72rem !important;
            }

            .st-key-overview_filters_aligned div[data-testid="stTextInput"],
            .st-key-overview_filters_aligned div[data-testid="stDateInput"],
            .st-key-overview_filters_aligned div[data-testid="stSelectbox"] {
                padding-top: 0 !important;
                margin-top: 0 !important;
            }

            .st-key-overview_filters_aligned div[data-testid="stTextInput"] > div,
            .st-key-overview_filters_aligned div[data-testid="stDateInput"] > div,
            .st-key-overview_filters_aligned div[data-testid="stSelectbox"] > div {
                margin-top: 0 !important;
            }

            .st-key-overview_filters_aligned div[data-testid="stTextInput"] div[data-baseweb="input"] > div,
            .st-key-overview_filters_aligned div[data-testid="stDateInput"] div[data-baseweb="input"] > div,
            .st-key-overview_filters_aligned div[data-testid="stSelectbox"] > div[data-baseweb="select"] > div {
                min-height: 54px !important;
                height: 54px !important;
                box-sizing: border-box !important;
            }

            /* Visão Geral: reserva a mesma altura para todos os títulos dos filtros.
               Assim, mesmo o título maior da busca não empurra somente o último campo para baixo. */
            .st-key-overview_filters_aligned [data-testid="stWidgetLabel"],
            .st-key-overview_filters_aligned label {
                min-height: 42px !important;
                height: 42px !important;
                display: flex !important;
                align-items: flex-start !important;
                margin: 0 !important;
                padding: 0 !important;
                line-height: 1.18 !important;
                box-sizing: border-box !important;
            }

            .st-key-overview_filters_aligned [data-testid="stWidgetLabel"] p,
            .st-key-overview_filters_aligned label p {
                margin: 0 !important;
                line-height: 1.18 !important;
            }


            /* CORREÇÃO REAL DO ALINHAMENTO DOS FILTROS DA VISÃO GERAL:
               todos os labels com a mesma altura e todos os campos iniciando no mesmo topo. */
            .st-key-overview_filters_aligned div[data-testid="stHorizontalBlock"] {
                align-items: flex-start !important;
            }

            .st-key-overview_filters_aligned div[data-testid="column"] {
                padding-top: 0 !important;
                margin-top: 0 !important;
            }

            .st-key-overview_filters_aligned [data-testid="stWidgetLabel"],
            .st-key-overview_filters_aligned label {
                height: 24px !important;
                min-height: 24px !important;
                max-height: 24px !important;
                margin: 0 0 6px 0 !important;
                padding: 0 !important;
                display: flex !important;
                align-items: flex-start !important;
                overflow: visible !important;
            }

            .st-key-overview_filters_aligned [data-testid="stWidgetLabel"] p,
            .st-key-overview_filters_aligned label p,
            .st-key-overview_filters_aligned label span {
                margin: 0 !important;
                padding: 0 !important;
                line-height: 18px !important;
                height: 18px !important;
                white-space: nowrap !important;
                overflow: hidden !important;
                text-overflow: ellipsis !important;
            }

            .st-key-overview_filters_aligned div[data-testid="stSelectbox"],
            .st-key-overview_filters_aligned div[data-testid="stDateInput"],
            .st-key-overview_filters_aligned div[data-testid="stTextInput"] {
                padding-top: 0 !important;
                margin-top: 0 !important;
            }

            .st-key-overview_filters_aligned div[data-testid="stSelectbox"] > div,
            .st-key-overview_filters_aligned div[data-testid="stDateInput"] > div,
            .st-key-overview_filters_aligned div[data-testid="stTextInput"] > div {
                padding-top: 0 !important;
                margin-top: 0 !important;
            }

            .st-key-overview_filters_aligned div[data-baseweb="select"] > div,
            .st-key-overview_filters_aligned div[data-baseweb="input"] > div {
                height: 54px !important;
                min-height: 54px !important;
                margin-top: 0 !important;
                transform: none !important;
                box-sizing: border-box !important;
            }

            /* Ajuste final: os componentes nativos de data e texto possuem um recuo
               interno diferente do selectbox. Sobe somente as caixas, preservando
               a posição dos títulos e alinhando visualmente todos os filtros. */
            .st-key-overview_filters_aligned div[data-testid="stDateInput"] [data-baseweb="input"],
            .st-key-overview_filters_aligned div[data-testid="stTextInput"] [data-baseweb="input"] {
                position: relative !important;
                top: -4px !important;
            }

            .st-key-overview_filters_aligned div[data-testid="stDateInput"] [data-baseweb="input"] > div,
            .st-key-overview_filters_aligned div[data-testid="stTextInput"] [data-baseweb="input"] > div {
                margin-top: 0 !important;
                top: 0 !important;
            }

            /* ALINHAMENTO DEFINITIVO DA VISÃO GERAL:
               labels próprios em uma linha e todos os controles em outra linha. */
            .st-key-overview_filter_labels {
                margin-bottom: 6px !important;
            }

            .st-key-overview_filter_labels div[data-testid="stHorizontalBlock"],
            .st-key-overview_filter_controls div[data-testid="stHorizontalBlock"] {
                align-items: flex-start !important;
                gap: 0.72rem !important;
            }

            .overview-filter-custom-label {
                min-height: 20px !important;
                height: 20px !important;
                margin: 0 !important;
                padding: 0 !important;
                color: rgba(255,255,255,0.92) !important;
                font-size: 0.86rem !important;
                font-weight: 700 !important;
                line-height: 20px !important;
                white-space: nowrap !important;
                overflow: hidden !important;
                text-overflow: ellipsis !important;
            }

            .st-key-overview_filter_controls div[data-testid="stSelectbox"],
            .st-key-overview_filter_controls div[data-testid="stDateInput"],
            .st-key-overview_filter_controls div[data-testid="stTextInput"],
            .st-key-overview_filter_controls div[data-testid="stElementContainer"],
            .st-key-overview_filter_controls div[data-testid="stVerticalBlock"] {
                margin-top: 0 !important;
                padding-top: 0 !important;
            }

            .st-key-overview_filter_controls div[data-testid="stDateInput"] [data-baseweb="input"],
            .st-key-overview_filter_controls div[data-testid="stTextInput"] [data-baseweb="input"],
            .st-key-overview_filter_controls div[data-testid="stSelectbox"] [data-baseweb="select"] {
                position: static !important;
                top: auto !important;
                margin-top: 0 !important;
                padding-top: 0 !important;
                transform: none !important;
            }

            .st-key-overview_filter_controls div[data-testid="stDateInput"] [data-baseweb="input"] > div,
            .st-key-overview_filter_controls div[data-testid="stTextInput"] [data-baseweb="input"] > div,
            .st-key-overview_filter_controls div[data-testid="stSelectbox"] [data-baseweb="select"] > div {
                min-height: 54px !important;
                height: 54px !important;
                margin-top: 0 !important;
                padding-top: 0 !important;
                padding-bottom: 0 !important;
                transform: none !important;
                box-sizing: border-box !important;
            }

            .st-key-overview_filter_controls [data-testid="stWidgetLabel"] {
                display: none !important;
            }

            /* Visão Geral: mantém os títulos dos filtros sempre brancos. */
            .st-key-overview_filter_labels .overview-filter-custom-label,
            .st-key-overview_filter_labels .overview-filter-custom-label *,
            .st-key-overview_filter_labels .overview-filter-custom-label p,
            .st-key-overview_filter_labels .overview-filter-custom-label span,
            .st-key-overview_filter_labels [data-testid="stMarkdownContainer"],
            .st-key-overview_filter_labels [data-testid="stMarkdownContainer"] *,
            .st-key-overview_filter_labels p,
            .st-key-overview_filter_labels span,
            .st-key-overview_filter_labels div {
                color: #FFFFFF !important;
                -webkit-text-fill-color: #FFFFFF !important;
                opacity: 1 !important;
            }
        </style>
        """
    )


# =========================================================
# SESSÃO DE NAVEGAÇÃO PARA LINKS INTERNOS
# =========================================================
@st.cache_resource
def get_navigation_session_registry() -> set[str]:
    """Mantém tokens temporários para preservar o login ao usar os links do submenu."""
    return set()


def create_navigation_session_token() -> str:
    token = uuid.uuid4().hex
    get_navigation_session_registry().add(token)
    st.session_state.navigation_session_token = token
    st.query_params["session"] = token
    return token


def restore_navigation_session_from_url() -> None:
    """Restaura o login quando um link interno recarrega a página com token válido."""
    if st.session_state.get("authenticated", False):
        return

    token = normalize_text(st.query_params.get("session", ""))

    if token and token in get_navigation_session_registry():
        st.session_state.authenticated = True
        st.session_state.navigation_session_token = token
        st.session_state.auth_error = ""


def revoke_navigation_session_token() -> None:
    token = normalize_text(st.session_state.get("navigation_session_token", ""))

    if token:
        get_navigation_session_registry().discard(token)

    st.session_state.navigation_session_token = ""


# =========================================================
# LOGIN
# =========================================================
def check_login(username: str, password: str) -> bool:
    expected_user = get_runtime_setting("APP_USERNAME", "oppitech")
    expected_password = get_runtime_setting("APP_PASSWORD", "100316Rahi*")

    return username == expected_user and password == expected_password


def render_login_page() -> None:
    apply_login_css()
    logo_data_uri = get_logo_data_uri()

    left_column, right_column = st.columns([0.86, 1.14], gap="large")

    with left_column:
        logo_html = (
            f'<div class="login-logo-wrap"><img src="{logo_data_uri}" class="login-logo-img" alt="Oppi Tech"></div>'
            if logo_data_uri
            else '<div class="login-logo-wrap"><div class="login-logo-fallback"></div></div>'
        )

        render_html(
            f"""
            <div class="login-brand-panel">
                {logo_html}
                <div class="login-brand-title">
                    Dashboard
                    <span class="login-brand-highlight">Oppi Comercial</span>
                </div>
                <div class="login-brand-subtitle">Painel de gestão comercial</div>
                <div class="login-accent-line"></div>
                <div class="login-benefit">
                    <div class="login-benefit-icon">🛡️</div>
                    <div>Segurança, performance e inteligência para impulsionar seus resultados.</div>
                </div>
            </div>
            """
        )

    with right_column:
        render_html('<div class="login-right-spacer"></div>')

        with st.form("login_form", clear_on_submit=False):
            render_html(
                """
                <div class="login-top-icon">🛡️</div>
                <div class="login-card-title">Acesse o painel comercial da Oppi Tech</div>
                <div class="login-card-subtitle">Faça login para continuar</div>
                """
            )

            username = st.text_input(
                "Usuário",
                placeholder="Digite seu usuário",
            )

            password = st.text_input(
                "Senha",
                type="password",
                placeholder="Digite sua senha",
            )

            submitted = st.form_submit_button(
                "Entrar",
                use_container_width=True,
            )

            render_html(
                """
                <div class="login-forgot-row">
                    <div class="login-forgot-line"></div>
                    <div class="login-forgot-text">Esqueceu sua senha?</div>
                    <div class="login-forgot-line"></div>
                </div>
                """
            )

        if submitted:
            if check_login(username, password):
                st.session_state.authenticated = True
                st.session_state.auth_error = ""
                create_navigation_session_token()
                st.rerun()
            else:
                st.session_state.auth_error = "Usuário ou senha inválidos."

        if st.session_state.auth_error:
            render_html(
                f'<div class="login-error">{html.escape(st.session_state.auth_error)}</div>'
            )


# =========================================================
# SIDEBAR
# =========================================================
def _query_param_value(name: str) -> str:
    value = st.query_params.get(name, "")

    if isinstance(value, list):
        return normalize_text(value[-1] if value else "")

    return normalize_text(value)


def _sync_navigation_from_query_params() -> None:
    requested_page = normalize_search_text(_query_param_value("page"))
    requested_contracts_page = normalize_search_text(_query_param_value("contracts"))

    if requested_page == "visao-geral":
        st.session_state.selected_page = "Visão Geral"
    elif requested_page == "pesos-e-medidas":
        st.session_state.selected_page = "Pesos e Medidas"
    elif requested_page == "cadastro":
        st.session_state.selected_page = "Cadastro"

        if requested_contracts_page == "todos":
            st.session_state.selected_cadastro_subpage = "Todos os cadastros"
        elif requested_contracts_page == "novo":
            st.session_state.selected_cadastro_subpage = "Novo cadastro"



def install_sidebar_navigation_persistence() -> None:
    """
    Mantém o menu lateral aberto ao navegar entre as páginas internas.

    O menu só permanece recolhido quando o próprio usuário clica na seta
    nativa do Streamlit. Ao clicar em Visão Geral, Cadastro ou Pesos e
    Medidas, a navegação preserva o menu aberto em vez de recolhê-lo.
    """
    components.html(
        """
        <script>
            (function () {
                function getHostWindow() {
                    try {
                        return window.parent || window;
                    } catch (error) {
                        return window;
                    }
                }

                function getHostDocument() {
                    try {
                        if (window.frameElement && window.frameElement.ownerDocument) {
                            return window.frameElement.ownerDocument;
                        }
                    } catch (error) {}

                    try {
                        return window.parent.document;
                    } catch (error) {
                        return document;
                    }
                }

                const hostWindow = getHostWindow();
                const hostDocument = getHostDocument();
                const storageKey = "__oppi_keep_sidebar_expanded_after_navigation__";
                const installedKey = "__oppi_sidebar_navigation_persistence_installed__";

                function isElementVisible(element) {
                    if (!element) {
                        return false;
                    }

                    const style = hostWindow.getComputedStyle(element);
                    const rect = element.getBoundingClientRect();

                    return (
                        style.display !== "none" &&
                        style.visibility !== "hidden" &&
                        Number(style.opacity || "1") > 0 &&
                        rect.width > 4 &&
                        rect.height > 4
                    );
                }

                function sidebarIsExpanded() {
                    const sidebar = hostDocument.querySelector('section[data-testid="stSidebar"]');

                    if (!sidebar) {
                        return false;
                    }

                    const style = hostWindow.getComputedStyle(sidebar);
                    const rect = sidebar.getBoundingClientRect();

                    return (
                        style.display !== "none" &&
                        style.visibility !== "hidden" &&
                        rect.width > 120 &&
                        rect.right > 40
                    );
                }

                function getVisibleExpandControl() {
                    const selectors = [
                        '[data-testid="collapsedControl"] button',
                        '[data-testid="stSidebarCollapsedControl"] button',
                        'button[data-testid="collapsedControl"]',
                        'button[data-testid="stSidebarCollapsedControl"]',
                        '[data-testid="collapsedControl"]',
                        '[data-testid="stSidebarCollapsedControl"]'
                    ];

                    for (const selector of selectors) {
                        const element = hostDocument.querySelector(selector);

                        if (isElementVisible(element)) {
                            return element;
                        }
                    }

                    return null;
                }

                function ensureSidebarExpanded(attempt) {
                    if (sidebarIsExpanded()) {
                        try {
                            hostWindow.sessionStorage.removeItem(storageKey);
                        } catch (error) {}
                        return;
                    }

                    const control = getVisibleExpandControl();

                    if (control) {
                        control.click();
                    }

                    if ((attempt || 0) < 24) {
                        hostWindow.setTimeout(function () {
                            ensureSidebarExpanded((attempt || 0) + 1);
                        }, 120);
                    }
                }

                function shouldPreserveSidebarForLink(link) {
                    if (!link || !link.matches) {
                        return false;
                    }

                    if (!link.matches('.oppi-side-nav a')) {
                        return false;
                    }

                    const href = String(link.getAttribute('href') || '');
                    return href.includes('page=');
                }

                if (!hostWindow[installedKey]) {
                    hostWindow[installedKey] = true;

                    hostDocument.addEventListener(
                        'pointerdown',
                        function (event) {
                            const link = event.target && event.target.closest
                                ? event.target.closest('a')
                                : null;

                            if (!shouldPreserveSidebarForLink(link)) {
                                return;
                            }

                            if (sidebarIsExpanded()) {
                                try {
                                    hostWindow.sessionStorage.setItem(storageKey, '1');
                                } catch (error) {}
                            }
                        },
                        true
                    );
                }

                try {
                    if (hostWindow.sessionStorage.getItem(storageKey) === '1') {
                        hostWindow.setTimeout(function () {
                            ensureSidebarExpanded(0);
                        }, 80);
                    }
                } catch (error) {}
            })();
        </script>
        """,
        height=0,
        scrolling=False,
    )

def render_sidebar() -> str:
    _sync_navigation_from_query_params()

    with st.sidebar:
        logo_data_uri = get_logo_data_uri()
        logo_html = (
            f'<div class="side-logo-wrap"><img src="{logo_data_uri}" class="side-logo-img" alt="Oppi Tech"></div>'
            if logo_data_uri
            else '<div class="side-logo-wrap"><div class="side-logo-fallback"></div></div>'
        )

        render_html(
            f"""
            {logo_html}
            <div class="side-title">
                Dashboard
                <span class="side-highlight">Oppi Comercial</span>
            </div>
            <div class="side-subtitle">Painel de gestão comercial</div>
            <div class="side-line"></div>
            """
        )

        if st.session_state.selected_page == "Propostas":
            st.session_state.selected_page = "Cadastro"

        if st.session_state.selected_page not in ["Visão Geral", "Cadastro", "Pesos e Medidas"]:
            st.session_state.selected_page = "Visão Geral"

        overview_active = "active" if st.session_state.selected_page == "Visão Geral" else ""
        cadastro_active = "active" if st.session_state.selected_page == "Cadastro" else ""
        scores_active = "active" if st.session_state.selected_page == "Pesos e Medidas" else ""
        details_open = "open" if st.session_state.selected_page == "Cadastro" else ""
        novo_active = "active" if st.session_state.get("selected_cadastro_subpage", "Novo cadastro") == "Novo cadastro" else ""
        todos_active = "active" if st.session_state.get("selected_cadastro_subpage", "Novo cadastro") == "Todos os cadastros" else ""
        navigation_token = normalize_text(st.session_state.get("navigation_session_token", ""))
        session_query = f"&session={navigation_token}" if navigation_token else ""

        render_html(
            f"""
            <nav class="oppi-side-nav">
                <a class="oppi-nav-link {overview_active}" href="?page=visao-geral{session_query}" target="_self">
                    <span class="oppi-nav-dot"></span>
                    <span>Visão Geral</span>
                </a>

                <details class="oppi-cadastro-details {cadastro_active}" {details_open}>
                    <summary class="oppi-nav-summary">
                        <span class="oppi-nav-dot"></span>
                        <span>Cadastro</span>
                        <span class="oppi-nav-arrow">›</span>
                    </summary>

                    <div class="oppi-cadastro-flyout">
                        <div class="oppi-flyout-title">Cadastro</div>
                        <a class="oppi-flyout-link {novo_active}" href="?page=cadastro&contracts=novo{session_query}" target="_self">Novo cadastro</a>
                        <a class="oppi-flyout-link {todos_active}" href="?page=cadastro&contracts=todos{session_query}" target="_self">Todos os cadastros</a>
                    </div>
                </details>

                <a class="oppi-nav-link {scores_active}" href="?page=pesos-e-medidas{session_query}" target="_self">
                    <span class="oppi-nav-dot"></span>
                    <span>Pesos e Medidas</span>
                </a>
            </nav>
            """
        )

        # Fecha somente a caixinha lateral do Cadastro ao clicar fora dela.
        # O design aprovado do menu e do submenu permanece intacto.
        components.html(
            """
            <script>
                (function () {
                    function getParentDocument() {
                        try {
                            if (window.frameElement && window.frameElement.ownerDocument) {
                                return window.frameElement.ownerDocument;
                            }
                        } catch (error) {}

                        try {
                            return window.parent.document;
                        } catch (error) {
                            return null;
                        }
                    }

                    function installOutsideClickHandler() {
                        const parentDocument = getParentDocument();

                        if (!parentDocument) {
                            window.setTimeout(installOutsideClickHandler, 250);
                            return;
                        }

                        const handlerKey = "__oppiCadastroFlyoutOutsideClickHandler__";

                        if (window.parent[handlerKey]) {
                            return;
                        }

                        window.parent[handlerKey] = true;

                        parentDocument.addEventListener(
                            "pointerdown",
                            function (event) {
                                const openDetails = parentDocument.querySelector(
                                    ".oppi-cadastro-details[open]"
                                );

                                if (!openDetails) {
                                    return;
                                }

                                if (openDetails.contains(event.target)) {
                                    return;
                                }

                                openDetails.removeAttribute("open");
                            },
                            true
                        );

                        parentDocument.addEventListener(
                            "keydown",
                            function (event) {
                                if (event.key !== "Escape") {
                                    return;
                                }

                                const openDetails = parentDocument.querySelector(
                                    ".oppi-cadastro-details[open]"
                                );

                                if (openDetails) {
                                    openDetails.removeAttribute("open");
                                }
                            },
                            true
                        );
                    }

                    installOutsideClickHandler();
                })();
            </script>
            """,
            height=0,
            scrolling=False,
        )

        render_html(
            """
            <div class="side-tip">
                <div class="side-tip-icon">🛡️</div>
                <div class="side-tip-text">Segurança, performance e inteligência para impulsionar seus resultados.</div>
            </div>
            """
        )

        if st.button("Sair", use_container_width=True, key="sidebar_logout"):
            revoke_navigation_session_token()
            st.session_state.authenticated = False
            st.session_state.auth_error = ""
            st.query_params.clear()
            st.rerun()

    return st.session_state.selected_page


# =========================================================
# COMPONENTES DO DASHBOARD
# =========================================================
def render_metric_card(
    title: str,
    value: str,
    note: str,
    icon: str,
    background: str,
) -> None:
    render_html(
        f"""
        <div class="metric-card">
            <div class="metric-icon" style="background:{background};">{icon}</div>
            <div class="metric-label">{html.escape(title)}</div>
            <div class="metric-value">{html.escape(value)}</div>
            <div class="metric-note">{html.escape(note)}</div>
        </div>
        """
    )


def render_status_summary(filtered_df: pd.DataFrame) -> None:
    statuses = [
        (status_name, STATUS_COLORS.get(status_name, ("#EAF2FF", "#5C9DFF"))[1])
        for status_name in DASHBOARD_STATUS_OPTIONS
    ]

    total = max(len(filtered_df), 1)
    rows_html = ""

    for status_name, color in statuses:
        count = count_dashboard_status(filtered_df, status_name)
        percent = round((count / total) * 100)

        rows_html += (
            f'<div class="status-row">'
            f'<div class="status-left"><span style="color:{color};">●</span>&nbsp;&nbsp;{status_name}</div>'
            f'<div class="status-count">{count}</div>'
            f'<div class="status-percent">{percent}%</div>'
            f'</div>'
        )

    render_html(
        f"""
        <div class="section-heading">Resumo por status</div>
        <div class="section-subtitle">Distribuição atual dos leads no comercial.</div>
        <div class="status-wrap">{rows_html}</div>
        """
    )



def render_phone_copy_button(phone: str, row_key: str) -> None:
    """Renderiza um botão gradiente que copia o telefone no navegador."""
    safe_phone = normalize_text(phone)
    phone_json = json.dumps(safe_phone, ensure_ascii=False)

    components.html(
        f"""
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8" />
            <style>
                * {{
                    box-sizing: border-box;
                }}

                html, body {{
                    margin: 0;
                    padding: 0;
                    width: 100%;
                    height: 30px;
                    overflow: hidden;
                    background: transparent;
                    font-family: Arial, sans-serif;
                }}

                button {{
                    width: 100%;
                    height: 29px;
                    border: none;
                    border-radius: 7px;
                    cursor: pointer;
                    color: #FFFFFF;
                    font-size: 11px;
                    font-weight: 800;
                    letter-spacing: 0.01em;
                    background: linear-gradient(90deg, #FF4BAA 0%, #A91CFF 100%);
                    box-shadow: 0 8px 18px rgba(169, 28, 255, 0.22);
                    transition:
                        filter 0.16s ease,
                        box-shadow 0.16s ease;
                }}

                button:hover {{
                    filter: brightness(1.08);
                    box-shadow: 0 10px 20px rgba(169, 28, 255, 0.30);
                }}

                button:active {{
                    filter: brightness(0.96);
                }}

                button.copied {{
                    background: linear-gradient(90deg, #20B56B 0%, #55DF7D 100%);
                    box-shadow: 0 8px 18px rgba(32, 181, 107, 0.20);
                }}
            </style>
        </head>
        <body>
            <button id="copy-{html.escape(row_key)}" type="button" onclick="copyPhone()">
                Copiar
            </button>

            <script>
                const phoneValue = {phone_json};
                const button = document.getElementById("copy-{html.escape(row_key)}");

                function showCopied() {{
                    button.textContent = "Copiado!";
                    button.classList.add("copied");

                    setTimeout(() => {{
                        button.textContent = "Copiar";
                        button.classList.remove("copied");
                    }}, 1400);
                }}

                function fallbackCopy(value) {{
                    const textarea = document.createElement("textarea");
                    textarea.value = value;
                    textarea.setAttribute("readonly", "");
                    textarea.style.position = "fixed";
                    textarea.style.opacity = "0";
                    document.body.appendChild(textarea);
                    textarea.select();
                    textarea.setSelectionRange(0, textarea.value.length);
                    document.execCommand("copy");
                    document.body.removeChild(textarea);
                    showCopied();
                }}

                async function copyPhone() {{
                    if (!phoneValue) {{
                        button.textContent = "Sem número";
                        setTimeout(() => {{
                            button.textContent = "Copiar";
                        }}, 1400);
                        return;
                    }}

                    try {{
                        if (navigator.clipboard && window.isSecureContext) {{
                            await navigator.clipboard.writeText(phoneValue);
                            showCopied();
                        }} else {{
                            fallbackCopy(phoneValue);
                        }}
                    }} catch (error) {{
                        fallbackCopy(phoneValue);
                    }}
                }}
            </script>
        </body>
        </html>
        """,
        height=30,
        scrolling=False,
    )


def render_latest_calls_section(
    filtered_df: pd.DataFrame,
    columns: dict,
    source_df: pd.DataFrame,
) -> None:
    status_icons = {
        "Novo Lead": "✦",
        "Chamado Whats": "☘",
        "Conversando": "•",
        "Reunião": "◉",
        "Proposta": "▤",
        "Sem interesse": "⊘",
        "Fechado": "✓",
        "Sem Resposta": "⚑",
        "Sem Whatsapp": "–",
        "Retornar": "↩",
        "Ligação - Conversando Whats": "☎",
        "Ligação não atende/cx": "☎",
        "Ligação Numero errado": "!",
        "Ligação retornar": "↩",
    }

    statuses = [
        (
            status_name,
            status_icons.get(status_name, "•"),
            STATUS_COLORS.get(status_name, ("#EAF2FF", "#5C9DFF"))[0],
            STATUS_COLORS.get(status_name, ("#EAF2FF", "#5C9DFF"))[1],
        )
        for status_name in DASHBOARD_STATUS_OPTIONS
    ]

    selected_card_key = "ultimos_chamados_status_selecionado"

    if selected_card_key not in st.session_state:
        st.session_state[selected_card_key] = None

    valid_dates = source_df["_data_chamado"].dropna()

    if valid_dates.empty:
        date_max = date.today()
        date_min = date_max - timedelta(days=30)
    else:
        date_min = valid_dates.min().date()
        date_max = valid_dates.max().date()

    seller_options = sorted(
        [
            seller
            for seller in source_df["_vendedor"].dropna().astype(str).unique().tolist()
            if normalize_text(seller)
        ]
    )

    niche_options = sorted(
        [
            niche
            for niche in source_df["_nicho"].dropna().astype(str).unique().tolist()
            if normalize_text(niche)
        ],
        key=normalize_search_text,
    )

    state_options = sorted(
        [
            state
            for state in source_df["_estado"].dropna().astype(str).unique().tolist()
            if normalize_text(state)
        ],
        key=lambda value: (value == "Não identificado", value),
    )

    render_html(
        """
        <div class="latest-calls-shell">
            <div class="latest-calls-head">
                <div>
                    <div class="latest-filter-title">Filtros dos chamados</div>
                    <div class="latest-filter-subtitle">Refine os resultados por vendedor, status, período, nicho, estado ou empresa.</div>
                </div>
                <div class="latest-calls-chip">Filtros</div>
            </div>
        </div>
        """
    )

    # Renderiza os títulos em uma linha separada e os campos em outra linha.
    # Assim os componentes nativos do Streamlit (selectbox, date_input e text_input)
    # começam exatamente na mesma altura, independentemente do tipo do campo.
    filter_widths = [1.05, 1.0, 1.12, 1.0, 1.0, 1.32]

    with st.container(key="overview_filter_labels"):
        label_columns = st.columns(filter_widths, gap="small")
        filter_labels = [
            "Vendedor",
            "Status",
            "Período",
            "Nichos",
            "Estados",
            "Buscar empresa ou telefone",
        ]

        for label_column, filter_label in zip(label_columns, filter_labels):
            with label_column:
                render_html(
                    f'<div class="overview-filter-custom-label">{html.escape(filter_label)}</div>'
                )

    with st.container(key="overview_filter_controls"):
        filter_1, filter_2, filter_3, filter_4, filter_5, filter_6 = st.columns(
            filter_widths,
            gap="small",
        )

        with filter_1:
            st.selectbox(
                "Vendedor",
                ["Todos os vendedores"] + seller_options,
                key="dashboard_filter_seller",
                label_visibility="collapsed",
            )

        with filter_2:
            st.selectbox(
                "Status",
                ["Todos os status"] + STATUS_OPTIONS,
                key="dashboard_filter_status",
                label_visibility="collapsed",
            )

        with filter_3:
            st.date_input(
                "Período",
                value=st.session_state.dashboard_filter_period,
                min_value=date_min,
                max_value=max(date_max, date.today()),
                key="dashboard_filter_period",
                label_visibility="collapsed",
            )

        with filter_4:
            st.selectbox(
                "Nichos",
                ["Todos os nichos"] + niche_options,
                key="dashboard_filter_niche",
                label_visibility="collapsed",
            )

        with filter_5:
            st.selectbox(
                "Estados",
                ["Todos os estados"] + state_options,
                key="dashboard_filter_state",
                label_visibility="collapsed",
            )

        with filter_6:
            st.text_input(
                "Buscar empresa ou telefone",
                placeholder="Digite para buscar...",
                key="dashboard_filter_search",
                label_visibility="collapsed",
            )

    st.write("")

    def choose_status(status_name: str) -> None:
        st.session_state[selected_card_key] = status_name

    status_filter_value = normalize_text(st.session_state.get("dashboard_filter_status", "Todos os status"))

    if status_filter_value != "Todos os status":
        st.session_state[selected_card_key] = status_filter_value

    selected_status = st.session_state.get(selected_card_key)

    # Cards em duas fileiras para não espremer os textos quando houver muitos status.
    first_row_count = (len(statuses) + 1) // 2
    status_rows = [statuses[:first_row_count], statuses[first_row_count:]]

    for row_index, status_row in enumerate(status_rows, start=1):
        if not status_row:
            continue

        card_columns = st.columns(len(status_row), gap="small")

        for column, (status_name, icon, bg_color, icon_color) in zip(card_columns, status_row):
            count = count_dashboard_status(filtered_df, status_name)
            active = selected_status == status_name

            border = "1px solid rgba(255, 75, 170, 0.85)" if active else "1px solid rgba(255,255,255,0.06)"
            shadow = "0 0 0 1px rgba(169, 28, 255, 0.18), 0 18px 46px rgba(0,0,0,0.28), 0 0 22px rgba(255, 75, 170, 0.14)" if active else "0 18px 46px rgba(0,0,0,0.22)"

            with column:
                render_html(
                    f"""
                    <div class="latest-status-card" style="border:{border}; box-shadow:{shadow};">
                        <div class="latest-status-top">
                            <div class="latest-status-icon" style="background:{bg_color}; color:{icon_color};">{icon}</div>
                            <div class="latest-status-name">{html.escape(status_name)}</div>
                        </div>
                        <div class="latest-status-number">{count}</div>
                        <div class="latest-status-caption">registros nesta sessão</div>
                    </div>
                    """
                )

                st.button(
                    "Ver nomes",
                    key=f"btn_ultimos_linha_{row_index}_{normalize_search_text(status_name)}",
                    use_container_width=True,
                    on_click=choose_status,
                    args=(status_name,),
                )

        if row_index == 1:
            st.write("")

    selected_status = st.session_state.get(selected_card_key)
    search_term = normalize_text(st.session_state.get("dashboard_filter_search", ""))
    status_filter_value = normalize_text(st.session_state.get("dashboard_filter_status", "Todos os status"))

    # Quando o usuário combina filtros, a tabela respeita todos eles.
    # Se o filtro de Status já estiver selecionado, não aplica o card novamente
    # para evitar conflito entre filtros.
    if search_term or status_filter_value != "Todos os status":
        selected_df = filtered_df.copy()
    else:
        if not selected_status:
            render_html(
                """
                <div class="latest-placeholder-card">
                    Selecione um status clicando em “Ver nomes” para visualizar os registros.
                </div>
                """
            )
            return

        selected_df = filtered_df[filtered_df.apply(lambda row: row_matches_dashboard_card(row, selected_status), axis=1)].copy()

    selected_df = selected_df.sort_values(
        ["_data_chamado", "_empresa"],
        ascending=[False, True],
    )

    display_df = pd.DataFrame(
        {
            "Empresa": selected_df["_empresa"],
            "Telefone": selected_df["_telefone"],
            "E-mail": safe_series(selected_df, columns.get("email")),
            "CNPJ": safe_series(selected_df, columns.get("cnpj")),
            "Status WhatsApp": selected_df["_status_whatsapp_original"],
            "Status Ligação": selected_df["_status_ligacao_original"],
            "Status Geral": selected_df["_status_grupo"],
            "Vendedor": selected_df["_vendedor"],
            "Data": selected_df["_data_chamado"].dt.strftime("%d/%m/%Y").fillna(""),
        }
    )

    if display_df.empty:
        st.info("Nenhum chamado encontrado para este status no período selecionado.")
        return

    editor_df = display_df.copy()
    editor_df["_sheet_row"] = selected_df["_sheet_row"].astype(int).values

    flash_message = st.session_state.pop("status_auto_save_success", None)
    if flash_message:
        st.success(flash_message)

    flash_error = st.session_state.pop("status_auto_save_error", None)
    if flash_error:
        st.error(flash_error)

    render_html(
        """
        <div class="premium-inline-hint">
            Altere o <strong>Status WhatsApp</strong> ou o <strong>Status Ligação</strong> pelo seletor. Use <strong>Copiar</strong> para copiar o telefone.
        </div>
        """
    )

    with st.container(key="compact_inline_table"):
        header_columns = st.columns(
            [2.75, 1.25, 0.78, 1.45, 1.55, 1.15, 0.80],
            gap="small",
        )

        header_labels = [
            "Empresa",
            "Telefone",
            "Copiar",
            "Status WhatsApp",
            "Status Ligação",
            "Vendedor",
            "Data",
        ]

        for column, label in zip(header_columns, header_labels):
            with column:
                render_html(f'<div class="premium-inline-table-header">{html.escape(label)}</div>')

        status_whatsapp_column_name = columns.get("status_whatsapp") or columns.get("status")
        status_ligacao_column_name = columns.get("status_ligacao")

        for _, row in editor_df.iterrows():
            sheet_row = int(row["_sheet_row"])
            original_status_whatsapp = normalize_text(row["Status WhatsApp"]) or "Sem status"
            original_status_ligacao = normalize_text(row["Status Ligação"]) or "Sem status"

            if original_status_whatsapp not in STATUS_WHATSAPP_SELECT_OPTIONS:
                original_status_whatsapp = "Sem status"

            if original_status_ligacao not in STATUS_LIGACAO_SELECT_OPTIONS:
                original_status_ligacao = "Sem status"

            row_columns = st.columns(
                [2.75, 1.25, 0.78, 1.45, 1.55, 1.15, 0.80],
                gap="small",
            )

            with row_columns[0]:
                render_html(
                    f'<div class="premium-inline-cell">{html.escape(normalize_text(row["Empresa"]) or "Sem empresa")}</div>'
                )

            with row_columns[1]:
                render_html(
                    f'<div class="premium-inline-cell phone">{html.escape(normalize_text(row["Telefone"]) or "Sem número")}</div>'
                )

            with row_columns[2]:
                render_phone_copy_button(
                    normalize_text(row["Telefone"]),
                    row_key=f"phone-{sheet_row}",
                )

            whatsapp_widget_key = f"inline_status_whatsapp_{sheet_row}_{normalize_search_text(original_status_whatsapp).replace(' ', '_')}"
            ligacao_widget_key = f"inline_status_ligacao_{sheet_row}_{normalize_search_text(original_status_ligacao).replace(' ', '_')}"

            def save_inline_status_whatsapp(
                sheet_row_value: int = sheet_row,
                widget_key: str = whatsapp_widget_key,
                previous_status: str = original_status_whatsapp,
            ) -> None:
                new_status = normalize_text(st.session_state.get(widget_key, previous_status))

                if new_status == previous_status:
                    return

                if not status_whatsapp_column_name:
                    st.session_state["status_auto_save_error"] = "Não encontrei a coluna Status WhatsApp na planilha."
                    return

                try:
                    update_statuses_in_sheet(
                        changes=[{"sheet_row": sheet_row_value, "status": new_status}],
                        status_column_name=status_whatsapp_column_name,
                        updated_at_column_name=columns.get("ultima_atualizacao"),
                    )
                    display_status = "Sem status" if new_status == "Sem status" else new_status
                    st.session_state["status_auto_save_success"] = (
                        f"Status WhatsApp alterado para “{display_status}” e salvo diretamente na planilha."
                    )
                except Exception as error:
                    st.session_state["status_auto_save_error"] = f"Não consegui atualizar o Status WhatsApp: {error}"
                    st.session_state[widget_key] = previous_status

            def save_inline_status_ligacao(
                sheet_row_value: int = sheet_row,
                widget_key: str = ligacao_widget_key,
                previous_status: str = original_status_ligacao,
            ) -> None:
                new_status = normalize_text(st.session_state.get(widget_key, previous_status))

                if new_status == previous_status:
                    return

                if not status_ligacao_column_name:
                    st.session_state["status_auto_save_error"] = "Não encontrei a coluna Status Ligação na planilha."
                    return

                try:
                    update_statuses_in_sheet(
                        changes=[{"sheet_row": sheet_row_value, "status": new_status}],
                        status_column_name=status_ligacao_column_name,
                        updated_at_column_name=columns.get("ultima_atualizacao"),
                    )
                    display_status = "Sem status" if new_status == "Sem status" else new_status
                    st.session_state["status_auto_save_success"] = (
                        f"Status Ligação alterado para “{display_status}” e salvo diretamente na planilha."
                    )
                except Exception as error:
                    st.session_state["status_auto_save_error"] = f"Não consegui atualizar o Status Ligação: {error}"
                    st.session_state[widget_key] = previous_status

            with row_columns[3]:
                st.selectbox(
                    "Status WhatsApp",
                    STATUS_WHATSAPP_SELECT_OPTIONS,
                    index=STATUS_WHATSAPP_SELECT_OPTIONS.index(original_status_whatsapp),
                    key=whatsapp_widget_key,
                    label_visibility="collapsed",
                    on_change=save_inline_status_whatsapp,
                )

            with row_columns[4]:
                st.selectbox(
                    "Status Ligação",
                    STATUS_LIGACAO_SELECT_OPTIONS,
                    index=STATUS_LIGACAO_SELECT_OPTIONS.index(original_status_ligacao),
                    key=ligacao_widget_key,
                    label_visibility="collapsed",
                    on_change=save_inline_status_ligacao,
                )

            with row_columns[5]:
                render_html(
                    f'<div class="premium-inline-cell muted">{html.escape(normalize_text(row["Vendedor"]) or "Sem vendedor")}</div>'
                )

            with row_columns[6]:
                render_html(
                    f'<div class="premium-inline-cell date">{html.escape(normalize_text(row["Data"]))}</div>'
                )

def prepare_filters(df: pd.DataFrame, columns: dict) -> pd.DataFrame:
    title_column, refresh_column = st.columns([3.8, 1.0], gap="large")

    with title_column:
        render_html(
            """
            <div class="page-title">Visão Geral</div>
            <div class="page-subtitle">Acompanhe o desempenho da operação comercial em tempo real.</div>
            """
        )

    with refresh_column:
        st.write("")
        if st.button("🔄 Atualizar dados", use_container_width=True):
            st.cache_data.clear()
            st.cache_resource.clear()
            st.rerun()

    valid_dates = df["_data_chamado"].dropna()

    if valid_dates.empty:
        date_max = date.today()
        date_min = date_max - timedelta(days=30)
    else:
        date_min = valid_dates.min().date()
        date_max = valid_dates.max().date()

    seller_options = sorted(
        [
            seller
            for seller in df["_vendedor"].dropna().astype(str).unique().tolist()
            if normalize_text(seller)
        ]
    )

    if "dashboard_filter_seller" not in st.session_state:
        st.session_state.dashboard_filter_seller = "Todos os vendedores"

    if st.session_state.dashboard_filter_seller not in ["Todos os vendedores"] + seller_options:
        st.session_state.dashboard_filter_seller = "Todos os vendedores"

    if "dashboard_filter_status" not in st.session_state:
        st.session_state.dashboard_filter_status = "Todos os status"

    if st.session_state.dashboard_filter_status not in ["Todos os status"] + STATUS_OPTIONS:
        st.session_state.dashboard_filter_status = "Todos os status"

    if "dashboard_filter_period" not in st.session_state:
        st.session_state.dashboard_filter_period = (date_min, date_max)

    if "dashboard_filter_search" not in st.session_state:
        st.session_state.dashboard_filter_search = ""

    niche_options = sorted(
        [
            niche
            for niche in df["_nicho"].dropna().astype(str).unique().tolist()
            if normalize_text(niche)
        ],
        key=normalize_search_text,
    )

    state_options = sorted(
        [
            state
            for state in df["_estado"].dropna().astype(str).unique().tolist()
            if normalize_text(state)
        ],
        key=lambda value: (value == "Não identificado", value),
    )

    if "dashboard_filter_niche" not in st.session_state:
        st.session_state.dashboard_filter_niche = "Todos os nichos"

    if st.session_state.dashboard_filter_niche not in ["Todos os nichos"] + niche_options:
        st.session_state.dashboard_filter_niche = "Todos os nichos"

    if "dashboard_filter_state" not in st.session_state:
        st.session_state.dashboard_filter_state = "Todos os estados"

    if st.session_state.dashboard_filter_state not in ["Todos os estados"] + state_options:
        st.session_state.dashboard_filter_state = "Todos os estados"

    selected_seller = st.session_state.dashboard_filter_seller
    selected_status = st.session_state.dashboard_filter_status
    selected_range = st.session_state.dashboard_filter_period
    selected_niche = st.session_state.dashboard_filter_niche
    selected_state = st.session_state.dashboard_filter_state
    search_term = st.session_state.dashboard_filter_search

    # Quando existe busca digitada, ela procura na base inteira, mas somente em
    # campos reais do cadastro. Não usamos nicho/status/vendedor como alvo da busca,
    # porque isso fazia aparecer linhas sem nome de empresa quando o termo digitado
    # continha palavras genéricas como "marmoraria" ou "granitos".
    if normalize_text(search_term):
        searchable_column_keys = [
            "empresa",
            "telefone_b2b",
            "telefone_fixo",
            "telefone_alternativo",
            "cnpj",
            "endereco",
            "email",
            "site",
            "socio_1",
            "socio_2",
            "socio_3",
        ]

        def row_matches_search(row) -> bool:
            searchable_values = [
                normalize_text(row.get("_empresa", "")),
                normalize_text(row.get("_telefone", "")),
            ]

            for column_key in searchable_column_keys:
                column_name = columns.get(column_key)

                if column_name and column_name in row.index:
                    searchable_values.append(normalize_text(row.get(column_name, "")))

            search_target = " | ".join(searchable_values)
            return flexible_search_match(search_term, search_target)

        searched_df = df[df.apply(row_matches_search, axis=1)].copy()

        # Evita exibir linhas sem nome quando o usuário está procurando empresa.
        searched_df = searched_df[
            searched_df["_empresa"].apply(lambda value: normalize_text(value) != "")
        ].copy()

        return searched_df

    # Sem busca digitada, os filtros seguem cumulativos normalmente.
    filtered_df = df.copy()

    if selected_seller != "Todos os vendedores":
        filtered_df = filtered_df[filtered_df["_vendedor"] == selected_seller].copy()

    if selected_status != "Todos os status":
        filtered_df = filtered_df[filtered_df.apply(lambda row: row_matches_status_filter(row, selected_status), axis=1)].copy()

    if selected_niche != "Todos os nichos":
        filtered_df = filtered_df[filtered_df["_nicho"] == selected_niche].copy()

    if selected_state != "Todos os estados":
        filtered_df = filtered_df[filtered_df["_estado"] == selected_state].copy()

    filtered_df = apply_period_filter(
        filtered_df,
        "_data_chamado",
        selected_range,
    )

    return filtered_df


# =========================================================
# PÁGINA: VISÃO GERAL
# =========================================================
def render_overview_page(df: pd.DataFrame, columns: dict) -> None:
    registration_success = st.session_state.pop("company_registration_success", None)

    if registration_success:
        st.success(registration_success)

    filtered_df = prepare_filters(df, columns)

    today = pd.Timestamp.now(tz="America/Sao_Paulo").normalize().tz_localize(None)
    start_week = today - pd.Timedelta(days=today.weekday())
    start_month = today.replace(day=1)

    called_dates = pd.to_datetime(filtered_df["_data_chamado"], errors="coerce")

    called_today = int((called_dates.dt.normalize() == today).sum())
    called_week = int((called_dates >= start_week).sum())
    called_month = int((called_dates >= start_month).sum())
    companies = int(filtered_df["_empresa"].replace("", pd.NA).dropna().nunique())

    card_1, card_2, card_3, card_4 = st.columns(4, gap="medium")

    with card_1:
        render_metric_card("Chamados hoje", str(called_today), "Base atual", "☎", "linear-gradient(135deg,#FF4BAA,#C223FF)")

    with card_2:
        render_metric_card("Chamados na semana", str(called_week), "Base atual", "🗓", "linear-gradient(135deg,#AE4BFF,#6E23FF)")

    with card_3:
        render_metric_card("Chamados no mês", str(called_month), "Base atual", "📊", "linear-gradient(135deg,#FF4BAA,#8F2BFF)")

    with card_4:
        render_metric_card("Empresas cadastradas no mês", str(companies), "Base atual filtrada", "🏢", "linear-gradient(135deg,#8F2BFF,#C94AFF)")

    st.write("")

    chart_column, status_column = st.columns([2.1, 1.0], gap="large")

    with chart_column:
        render_html(
            """
            <div class="section-heading">Chamados por semana</div>
            <div class="section-subtitle">Volume de chamados agrupado por semana conforme o período selecionado.</div>
            """
        )

        chart_df = filtered_df.copy()
        chart_df["_data_chamado"] = pd.to_datetime(chart_df["_data_chamado"], errors="coerce")
        chart_df = chart_df.dropna(subset=["_data_chamado"]).copy()

        if chart_df.empty:
            current_week_start = (
                pd.Timestamp.today().normalize()
                - pd.to_timedelta(pd.Timestamp.today().weekday(), unit="D")
            )
            week_starts = pd.date_range(
                end=current_week_start,
                periods=4,
                freq="7D",
            )
            chart_df = pd.DataFrame({"InicioSemana": week_starts})
            chart_df["Quantidade"] = 0
        else:
            chart_df["InicioSemana"] = (
                chart_df["_data_chamado"].dt.normalize()
                - pd.to_timedelta(chart_df["_data_chamado"].dt.weekday, unit="D")
            )
            chart_df = (
                chart_df.groupby("InicioSemana")
                .size()
                .reset_index(name="Quantidade")
                .sort_values("InicioSemana")
            )

        chart_df["FimSemana"] = chart_df["InicioSemana"] + pd.Timedelta(days=6)
        chart_df["Semana"] = (
            chart_df["InicioSemana"].dt.strftime("%d/%m")
            + " – "
            + chart_df["FimSemana"].dt.strftime("%d/%m")
        )

        # Mantém o preenchimento roxo mesmo quando existe apenas uma semana.
        # O ponto auxiliar serve apenas para desenhar a área visualmente e não
        # altera a contagem dos chamados.
        plot_df = chart_df.copy()

        if len(plot_df) == 1:
            support_point = plot_df.iloc[0].copy()
            support_point["InicioSemana"] = support_point["FimSemana"]
            plot_df = pd.concat(
                [plot_df, pd.DataFrame([support_point])],
                ignore_index=True,
            )

        figure = px.area(
            plot_df,
            x="InicioSemana",
            y="Quantidade",
            markers=True,
            custom_data=["Semana"],
        )

        figure.update_traces(
            line=dict(
                color="#E14BFF",
                width=4,
                shape="spline",
            ),
            marker=dict(
                size=9,
                color="#FFFFFF",
                line=dict(width=3, color="#D74BFF"),
            ),
            fill="tozeroy",
            fillcolor="rgba(224,67,255,0.34)",
            hovertemplate="Semana: %{customdata[0]}<br>Chamados: %{y}<extra></extra>",
        )

        figure.update_layout(
            height=370,
            margin=dict(l=20, r=20, t=8, b=8),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#FFFFFF"),
            xaxis_title="",
            yaxis_title="",
        )

        figure.update_xaxes(
            showgrid=False,
            tickmode="array",
            tickvals=chart_df["InicioSemana"].tolist(),
            ticktext=chart_df["Semana"].tolist(),
        )
        figure.update_yaxes(gridcolor="rgba(255,255,255,0.08)")

        st.plotly_chart(figure, use_container_width=True)

    with status_column:
        render_status_summary(filtered_df)

    st.write("")
    render_latest_calls_section(filtered_df, columns, df)


# =========================================================
# PÁGINA: CADASTRO
# =========================================================
def render_proposals_page(df: pd.DataFrame, columns: dict) -> None:
    apply_registration_css()

    render_html(
        """
        <div class="registration-header-card">
            <div class="registration-kicker">OPPI COMERCIAL • NOVO CADASTRO</div>
            <div class="registration-title">Novo cadastro</div>
            <div class="registration-subtitle">
                Registre uma nova empresa e envie os dados diretamente para a planilha comercial.
            </div>
        </div>
        """
    )

    seller_options = sorted(
        {
            normalize_text(value)
            for value in df["_vendedor"].tolist()
            if normalize_text(value) and normalize_text(value) != "Sem vendedor"
        }
    )

    if not seller_options:
        seller_options = ["Sem vendedor"]

    with st.form("company_registration_form", clear_on_submit=True):
        render_html(
            """
            <div class="registration-note">
                Preencha os dados abaixo. Os campos com <strong>*</strong> são obrigatórios.
                Ao finalizar, a empresa será adicionada automaticamente à aba Folha1.
            </div>
            <div class="registration-section">
                <div class="registration-section-title">DADOS DA EMPRESA</div>
                <div class="registration-section-text">Informações principais da empresa e dados institucionais.</div>
            </div>
            """
        )

        company_col, opening_col = st.columns([1.65, 0.75], gap="medium")

        with company_col:
            empresa = st.text_input(
                "Nome da empresa *",
                placeholder="Digite o nome da empresa",
            )

        with opening_col:
            data_abertura = st.text_input(
                "Data de abertura",
                placeholder="DD/MM/AAAA",
            )

        cnpj_col, capital_col = st.columns(2, gap="medium")

        with cnpj_col:
            cnpj = st.text_input(
                "CNPJ *",
                placeholder="00.000.000/0000-00",
            )

        with capital_col:
            capital = st.text_input(
                "Capital social",
                placeholder="R$ 0,00",
            )

        endereco = st.text_input(
            "Endereço",
            placeholder="Rua, avenida, número, bairro, cidade, estado e CEP",
        )

        email_col, site_col = st.columns(2, gap="medium")

        with email_col:
            email_empresa = st.text_input(
                "E-mail da empresa",
                placeholder="contato@empresa.com.br",
            )

        with site_col:
            site = st.text_input(
                "Site da empresa",
                placeholder="www.empresa.com.br",
            )

        render_html(
            """
            <div class="registration-section">
                <div class="registration-section-title">TELEFONES DA EMPRESA</div>
                <div class="registration-section-text">Contatos principais utilizados no acompanhamento comercial.</div>
            </div>
            """
        )

        phone_1, phone_2, phone_3 = st.columns(3, gap="medium")

        with phone_1:
            telefone_b2b = st.text_input(
                "Telefone B2B *",
                placeholder="(00) 00000-0000",
            )

        with phone_2:
            telefone_fixo = st.text_input(
                "Telefone fixo",
                placeholder="(00) 0000-0000",
            )

        with phone_3:
            telefone_alternativo = st.text_input(
                "Telefone alternativo",
                placeholder="(00) 00000-0000",
            )

        render_html(
            """
            <div class="registration-section">
                <div class="registration-section-title">SÓCIOS E RESPONSÁVEIS</div>
                <div class="registration-section-text">Cadastre os principais responsáveis vinculados à empresa.</div>
            </div>
            """
        )

        socio_1_col, telefone_socio_1_col = st.columns([1.45, 0.95], gap="medium")

        with socio_1_col:
            socio_1 = st.text_input(
                "Sócio 1",
                placeholder="Nome completo do primeiro sócio",
            )

        with telefone_socio_1_col:
            telefone_socio_1 = st.text_input(
                "Telefone do sócio 1",
                placeholder="(00) 00000-0000",
            )

        cpf_1_col, email_socio_col = st.columns([0.95, 1.45], gap="medium")

        with cpf_1_col:
            cpf_socio_1 = st.text_input(
                "CPF do sócio 1",
                placeholder="000.000.000-00",
            )

        with email_socio_col:
            email_socio_1 = st.text_input(
                "E-mail do sócio 1",
                placeholder="socio@empresa.com.br",
            )

        socio_2_col, telefone_socio_2_col = st.columns([1.45, 0.95], gap="medium")

        with socio_2_col:
            socio_2 = st.text_input(
                "Sócio 2",
                placeholder="Nome completo do segundo sócio",
            )

        with telefone_socio_2_col:
            telefone_socio_2 = st.text_input(
                "Telefone do sócio 2",
                placeholder="(00) 00000-0000",
            )

        cpf_2_col, cpf_2_spacer = st.columns([0.95, 1.45], gap="medium")

        with cpf_2_col:
            cpf_socio_2 = st.text_input(
                "CPF do sócio 2",
                placeholder="000.000.000-00",
            )

        with cpf_2_spacer:
            st.write("")

        socio_3_col, telefone_socio_3_col = st.columns([1.45, 0.95], gap="medium")

        with socio_3_col:
            socio_3 = st.text_input(
                "Sócio 3",
                placeholder="Nome completo do terceiro sócio",
            )

        with telefone_socio_3_col:
            telefone_socio_3 = st.text_input(
                "Telefone do sócio 3",
                placeholder="(00) 00000-0000",
            )

        cpf_3_col, cpf_3_spacer = st.columns([0.95, 1.45], gap="medium")

        with cpf_3_col:
            cpf_socio_3 = st.text_input(
                "CPF do sócio 3",
                placeholder="000.000.000-00",
            )

        with cpf_3_spacer:
            st.write("")

        render_html(
            """
            <div class="registration-section">
                <div class="registration-section-title">REDES SOCIAIS E ACOMPANHAMENTO</div>
                <div class="registration-section-text">Complete os dados comerciais e defina o status inicial do atendimento.</div>
            </div>
            """
        )

        social_1, social_2 = st.columns(2, gap="medium")

        with social_1:
            instagram = st.text_input(
                "Instagram",
                placeholder="@empresa",
            )

        with social_2:
            linkedin = st.text_input(
                "LinkedIn",
                placeholder="Link ou usuário do perfil",
            )

        vendedor_col, status_col, called_at_col = st.columns([1.15, 1.15, 0.85], gap="medium")

        with vendedor_col:
            vendedor = st.selectbox(
                "Vendedor *",
                seller_options,
            )

        with status_col:
            status = st.selectbox(
                "Status comercial *",
                STATUS_OPTIONS,
                index=0,
            )

        with called_at_col:
            data_chamado = st.date_input(
                "Data do chamado *",
                value=date.today(),
                format="DD/MM/YYYY",
            )

        observacoes = st.text_area(
            "Observações",
            placeholder="Digite informações adicionais importantes sobre a empresa ou o atendimento comercial.",
            height=120,
        )

        submitted = st.form_submit_button(
            "Cadastrar empresa",
            use_container_width=True,
        )

    if submitted:
        if not normalize_text(empresa):
            st.error("Preencha o nome da empresa para concluir o cadastro.")
            return

        if not normalize_text(cnpj):
            st.error("Preencha o CNPJ para concluir o cadastro.")
            return

        if not normalize_cnpj_for_duplicate(cnpj):
            st.error("Digite um CNPJ válido com 14 números.")
            return

        if not normalize_text(telefone_b2b):
            st.error("Preencha o telefone B2B para concluir o cadastro.")
            return

        if not normalize_text(telefone_fixo):
            st.error("Preencha o telefone fixo para concluir o cadastro.")
            return

        if not normalize_text(telefone_alternativo):
            st.error("Preencha o telefone alternativo para concluir o cadastro.")
            return

        for phone_label, phone_value in [
            ("Telefone B2B", telefone_b2b),
            ("Telefone fixo", telefone_fixo),
            ("Telefone alternativo", telefone_alternativo),
        ]:
            if not normalize_phone_for_duplicate(phone_value):
                st.error(f"Digite um número válido no campo {phone_label}.")
                return

        now_text = pd.Timestamp.now(tz="America/Sao_Paulo").strftime("%d/%m/%Y %H:%M")

        try:
            append_company_to_sheet(
                {
                    "empresa": empresa,
                    "data_abertura": data_abertura,
                    "capital": capital,
                    "cnpj": cnpj,
                    "endereco": endereco,
                    "email_empresa": email_empresa,
                    "site": site,
                    "telefone_b2b": telefone_b2b,
                    "telefone_fixo": telefone_fixo,
                    "telefone_alternativo": telefone_alternativo,
                    "socio_1": socio_1,
                    "cpf_socio_1": cpf_socio_1,
                    "email_socio_1": email_socio_1,
                    "telefone_socio_1": telefone_socio_1,
                    "socio_2": socio_2,
                    "telefone_socio_2": telefone_socio_2,
                    "cpf_socio_2": cpf_socio_2,
                    "socio_3": socio_3,
                    "telefone_socio_3": telefone_socio_3,
                    "cpf_socio_3": cpf_socio_3,
                    "instagram": instagram,
                    "linkedin": linkedin,
                    "vendedor": vendedor,
                    "status": status,
                    "data_chamado": data_chamado.strftime("%d/%m/%Y"),
                    "ultima_atualizacao": now_text,
                    "observacoes": observacoes,
                }
            )
        except DuplicateRegistrationError as error:
            st.error(str(error))
            return
        except Exception as error:
            st.error("Não consegui cadastrar a empresa na planilha.")
            st.code(str(error))
            return

        # Após o cadastro, abre a Visão Geral já com os dados atualizados.
        # Também limpa filtros antigos que poderiam esconder a nova empresa.
        st.session_state.selected_page = "Visão Geral"
        st.session_state.dashboard_filter_seller = "Todos os vendedores"
        st.session_state.dashboard_filter_status = "Todos os status"
        st.session_state.dashboard_filter_search = ""
        st.session_state.pop("dashboard_filter_period", None)
        st.session_state["ultimos_chamados_status_selecionado"] = status
        st.session_state["company_registration_success"] = (
            f"Empresa “{normalize_text(empresa)}” cadastrada com sucesso na planilha e adicionada à Visão Geral com o status “{normalize_text(status)}”."
        )
        st.rerun()


# =========================================================
# PÁGINA: TODOS OS CADASTROS
# =========================================================
def _contract_detail_value(row: pd.Series, columns: dict, key: str) -> str:
    column_name = columns.get(key)

    if column_name and column_name in row.index:
        value = normalize_text(row.get(column_name, ""))

        if value:
            return value

    return "Não informado"


def _contract_detail_field(label: str, value: str, full_width: bool = False, long_text: bool = False) -> str:
    field_classes = ["contract-detail-field"]
    value_classes = ["contract-detail-value"]

    if full_width:
        field_classes.append("full-width")

    if long_text:
        value_classes.append("long-text")

    return (
        f'<div class="{" ".join(field_classes)}">'
        f'<div class="contract-detail-label">{html.escape(label)}</div>'
        f'<div class="{" ".join(value_classes)}">{html.escape(normalize_text(value) or "Não informado")}</div>'
        f'</div>'
    )


def _contract_edit_value(row: pd.Series, columns: dict, key: str) -> str:
    column_name = columns.get(key)

    if column_name and column_name in row.index:
        return normalize_text(row.get(column_name, ""))

    return ""


def _contract_edit_date(row: pd.Series, columns: dict, key: str):
    value = _contract_edit_value(row, columns, key)
    parsed_value = parse_date(value)

    if pd.isna(parsed_value):
        return date.today()

    return parsed_value.date()


def render_contract_edit_form(df: pd.DataFrame, columns: dict, row: pd.Series, sheet_row: int) -> None:
    company_name = normalize_text(row.get("_empresa", "")) or "Empresa cadastrada"
    seller_options = sorted(
        {
            normalize_text(value)
            for value in df["_vendedor"].tolist()
            if normalize_text(value) and normalize_text(value) != "Sem vendedor"
        }
    )
    current_seller = normalize_text(row.get("_vendedor", "")) or "Sem vendedor"

    if current_seller not in seller_options:
        seller_options.append(current_seller)

    seller_options = sorted(set(seller_options)) or ["Sem vendedor"]
    current_status = status_group(row.get("_status_original", row.get("_status_grupo", "Novo Lead")))

    if current_status not in STATUS_OPTIONS:
        current_status = "Novo Lead"

    render_html(
        f"""
        <div class="registration-header-card">
            <div class="registration-kicker">OPPI COMERCIAL • EDITAR CADASTRO</div>
            <div class="registration-title">Editar {html.escape(company_name)}</div>
            <div class="registration-subtitle">
                Corrija os dados abaixo e salve. As alterações serão enviadas diretamente para a planilha comercial.
            </div>
        </div>
        """
    )

    with st.form(f"edit_company_form_{sheet_row}", clear_on_submit=False):
        render_html(
            """
            <div class="registration-note">
                Revise os campos necessários. Ao clicar em <strong>Salvar alterações</strong>, os dados serão atualizados diretamente na aba Folha1.
            </div>
            <div class="registration-section">
                <div class="registration-section-title">DADOS DA EMPRESA</div>
                <div class="registration-section-text">Informações principais da empresa e dados institucionais.</div>
            </div>
            """
        )

        company_col, opening_col = st.columns([1.65, 0.75], gap="medium")

        with company_col:
            empresa = st.text_input("Nome da empresa", value=_contract_edit_value(row, columns, "empresa"))

        with opening_col:
            data_abertura = st.text_input("Data de abertura", value=_contract_edit_value(row, columns, "data_abertura"))

        cnpj_col, capital_col = st.columns(2, gap="medium")

        with cnpj_col:
            cnpj = st.text_input("CNPJ", value=_contract_edit_value(row, columns, "cnpj"))

        with capital_col:
            capital = st.text_input("Capital social", value=_contract_edit_value(row, columns, "capital"))

        endereco = st.text_input("Endereço", value=_contract_edit_value(row, columns, "endereco"))

        email_col, site_col = st.columns(2, gap="medium")

        with email_col:
            email_empresa = st.text_input("E-mail da empresa", value=_contract_edit_value(row, columns, "email"))

        with site_col:
            site = st.text_input("Site da empresa", value=_contract_edit_value(row, columns, "site"))

        render_html(
            """
            <div class="registration-section">
                <div class="registration-section-title">TELEFONES DA EMPRESA</div>
                <div class="registration-section-text">Contatos principais utilizados no acompanhamento comercial.</div>
            </div>
            """
        )

        phone_1, phone_2, phone_3 = st.columns(3, gap="medium")

        with phone_1:
            telefone_b2b = st.text_input("Telefone B2B", value=_contract_edit_value(row, columns, "telefone_b2b"))

        with phone_2:
            telefone_fixo = st.text_input("Telefone fixo", value=_contract_edit_value(row, columns, "telefone_fixo"))

        with phone_3:
            telefone_alternativo = st.text_input("Telefone alternativo", value=_contract_edit_value(row, columns, "telefone_alternativo"))

        render_html(
            """
            <div class="registration-section">
                <div class="registration-section-title">SÓCIOS E RESPONSÁVEIS</div>
                <div class="registration-section-text">Cadastre ou corrija os responsáveis vinculados à empresa.</div>
            </div>
            """
        )

        socio_1_col, telefone_socio_1_col = st.columns([1.45, 0.95], gap="medium")

        with socio_1_col:
            socio_1 = st.text_input("Sócio 1", value=_contract_edit_value(row, columns, "socio_1"))

        with telefone_socio_1_col:
            telefone_socio_1 = st.text_input("Telefone do sócio 1", value=_contract_edit_value(row, columns, "telefone_socio_1"))

        cpf_1_col, email_socio_col = st.columns([0.95, 1.45], gap="medium")

        with cpf_1_col:
            cpf_socio_1 = st.text_input("CPF do sócio 1", value=_contract_edit_value(row, columns, "cpf_socio_1"))

        with email_socio_col:
            email_socio_1 = st.text_input("E-mail do sócio 1", value=_contract_edit_value(row, columns, "email_socio_1"))

        socio_2_col, telefone_socio_2_col = st.columns([1.45, 0.95], gap="medium")

        with socio_2_col:
            socio_2 = st.text_input("Sócio 2", value=_contract_edit_value(row, columns, "socio_2"))

        with telefone_socio_2_col:
            telefone_socio_2 = st.text_input("Telefone do sócio 2", value=_contract_edit_value(row, columns, "telefone_socio_2"))

        cpf_2_col, cpf_2_spacer = st.columns([0.95, 1.45], gap="medium")

        with cpf_2_col:
            cpf_socio_2 = st.text_input("CPF do sócio 2", value=_contract_edit_value(row, columns, "cpf_socio_2"))

        with cpf_2_spacer:
            st.write("")

        socio_3_col, telefone_socio_3_col = st.columns([1.45, 0.95], gap="medium")

        with socio_3_col:
            socio_3 = st.text_input("Sócio 3", value=_contract_edit_value(row, columns, "socio_3"))

        with telefone_socio_3_col:
            telefone_socio_3 = st.text_input("Telefone do sócio 3", value=_contract_edit_value(row, columns, "telefone_socio_3"))

        cpf_3_col, cpf_3_spacer = st.columns([0.95, 1.45], gap="medium")

        with cpf_3_col:
            cpf_socio_3 = st.text_input("CPF do sócio 3", value=_contract_edit_value(row, columns, "cpf_socio_3"))

        with cpf_3_spacer:
            st.write("")

        render_html(
            """
            <div class="registration-section">
                <div class="registration-section-title">REDES SOCIAIS E ACOMPANHAMENTO</div>
                <div class="registration-section-text">Atualize os dados comerciais e o status atual do atendimento.</div>
            </div>
            """
        )

        social_1, social_2 = st.columns(2, gap="medium")

        with social_1:
            instagram = st.text_input("Instagram", value=_contract_edit_value(row, columns, "instagram"))

        with social_2:
            linkedin = st.text_input("LinkedIn", value=_contract_edit_value(row, columns, "linkedin"))

        vendedor_col, status_col, called_at_col = st.columns([1.15, 1.15, 0.85], gap="medium")

        with vendedor_col:
            vendedor = st.selectbox("Vendedor", seller_options, index=seller_options.index(current_seller))

        with status_col:
            status = st.selectbox("Status comercial", STATUS_OPTIONS, index=STATUS_OPTIONS.index(current_status))

        with called_at_col:
            data_chamado = st.date_input(
                "Data do chamado",
                value=_contract_edit_date(row, columns, "data_chamado"),
                format="DD/MM/YYYY",
            )

        observacoes = st.text_area(
            "Observações",
            value=_contract_edit_value(row, columns, "observacoes"),
            height=120,
        )

        action_col_1, action_col_2 = st.columns(2, gap="medium")

        with action_col_1:
            save_changes = st.form_submit_button("💾 Salvar alterações", use_container_width=True)

        with action_col_2:
            cancel_edit = st.form_submit_button("Cancelar edição", use_container_width=True)

    if cancel_edit:
        st.session_state.edit_contract_sheet_row = None
        st.rerun()

    if not save_changes:
        return

    if not normalize_text(empresa):
        st.error("Preencha o nome da empresa antes de salvar.")
        return

    if normalize_text(cnpj) and not normalize_cnpj_for_duplicate(cnpj):
        st.error("Digite um CNPJ válido com 14 números ou deixe o campo vazio.")
        return

    for phone_label, phone_value in [
        ("Telefone B2B", telefone_b2b),
        ("Telefone fixo", telefone_fixo),
        ("Telefone alternativo", telefone_alternativo),
        ("Telefone do sócio 1", telefone_socio_1),
        ("Telefone do sócio 2", telefone_socio_2),
        ("Telefone do sócio 3", telefone_socio_3),
    ]:
        if normalize_text(phone_value) and not normalize_phone_for_duplicate(phone_value):
            st.error(f"Digite um número válido no campo {phone_label} ou deixe o campo vazio.")
            return

    for cpf_label, cpf_value in [
        ("CPF do sócio 1", cpf_socio_1),
        ("CPF do sócio 2", cpf_socio_2),
        ("CPF do sócio 3", cpf_socio_3),
    ]:
        if normalize_text(cpf_value) and not normalize_cpf_for_duplicate(cpf_value):
            st.error(f"Digite um CPF válido no campo {cpf_label} ou deixe o campo vazio.")
            return

    now_text = pd.Timestamp.now(tz="America/Sao_Paulo").strftime("%d/%m/%Y %H:%M")

    try:
        update_company_in_sheet(
            sheet_row=int(sheet_row),
            payload={
                "empresa": empresa,
                "data_abertura": data_abertura,
                "capital": capital,
                "cnpj": cnpj,
                "endereco": endereco,
                "email_empresa": email_empresa,
                "site": site,
                "telefone_b2b": telefone_b2b,
                "telefone_fixo": telefone_fixo,
                "telefone_alternativo": telefone_alternativo,
                "socio_1": socio_1,
                "cpf_socio_1": cpf_socio_1,
                "email_socio_1": email_socio_1,
                "telefone_socio_1": telefone_socio_1,
                "socio_2": socio_2,
                "cpf_socio_2": cpf_socio_2,
                "socio_3": socio_3,
                "cpf_socio_3": cpf_socio_3,
                "instagram": instagram,
                "linkedin": linkedin,
                "vendedor": vendedor,
                "status": status,
                "data_chamado": data_chamado.strftime("%d/%m/%Y"),
                "ultima_atualizacao": now_text,
                "observacoes": observacoes,
            },
        )
    except DuplicateRegistrationError as error:
        st.error(str(error))
        return
    except Exception as error:
        st.error("Não consegui atualizar os dados na planilha.")
        st.code(str(error))
        return

    st.session_state.edit_contract_sheet_row = None
    st.session_state.contract_update_success = "Dados atualizados com sucesso diretamente na planilha."
    st.rerun()


def render_contract_detail_page(df: pd.DataFrame, columns: dict, sheet_row: int) -> None:
    selected_rows = df[df["_sheet_row"].astype(int) == int(sheet_row)].copy()

    if selected_rows.empty:
        st.session_state.selected_contract_sheet_row = None
        st.warning("Não encontrei os dados dessa empresa na planilha.")
        return

    row = selected_rows.iloc[0]
    company_name = normalize_text(row.get("_empresa", "")) or "Empresa cadastrada"

    if st.session_state.get("edit_contract_sheet_row") == int(sheet_row):
        render_contract_edit_form(df, columns, row, int(sheet_row))
        return

    flash_message = st.session_state.pop("contract_update_success", None)

    if flash_message:
        st.success(flash_message)

    with st.container(key="contract_detail_back"):
        if st.button("← Voltar para empresas cadastradas", key="back_to_contracts_names"):
            st.session_state.selected_contract_sheet_row = None
            st.session_state.edit_contract_sheet_row = None
            st.rerun()

    render_html(
        f"""
        <div class="registration-header-card">
            <div class="registration-kicker">OPPI COMERCIAL • EMPRESA CADASTRADA</div>
            <div class="registration-title">{html.escape(company_name)}</div>
            <div class="registration-subtitle">
                Visualize os dados cadastrados diretamente na planilha comercial.
            </div>
        </div>
        """
    )

    company_fields = "".join(
        [
            _contract_detail_field("Nome da empresa", _contract_detail_value(row, columns, "empresa")),
            _contract_detail_field("Data de abertura", _contract_detail_value(row, columns, "data_abertura")),
            _contract_detail_field("CNPJ", _contract_detail_value(row, columns, "cnpj")),
            _contract_detail_field("Capital social", _contract_detail_value(row, columns, "capital")),
            _contract_detail_field("Endereço", _contract_detail_value(row, columns, "endereco"), full_width=True),
            _contract_detail_field("E-mail da empresa", _contract_detail_value(row, columns, "email")),
            _contract_detail_field("Site da empresa", _contract_detail_value(row, columns, "site")),
        ]
    )

    phone_fields = "".join(
        [
            _contract_detail_field("Telefone B2B", _contract_detail_value(row, columns, "telefone_b2b")),
            _contract_detail_field("Telefone fixo", _contract_detail_value(row, columns, "telefone_fixo")),
            _contract_detail_field("Telefone alternativo", _contract_detail_value(row, columns, "telefone_alternativo")),
        ]
    )

    partner_fields = "".join(
        [
            _contract_detail_field("Sócio 1", _contract_detail_value(row, columns, "socio_1")),
            _contract_detail_field("Telefone do sócio 1", _contract_detail_value(row, columns, "telefone_socio_1")),
            _contract_detail_field("CPF do sócio 1", _contract_detail_value(row, columns, "cpf_socio_1")),
            _contract_detail_field("E-mail do sócio 1", _contract_detail_value(row, columns, "email_socio_1")),
            _contract_detail_field("Sócio 2", _contract_detail_value(row, columns, "socio_2")),
            _contract_detail_field("Telefone do sócio 2", _contract_detail_value(row, columns, "telefone_socio_2")),
            _contract_detail_field("CPF do sócio 2", _contract_detail_value(row, columns, "cpf_socio_2")),
            _contract_detail_field("Sócio 3", _contract_detail_value(row, columns, "socio_3")),
            _contract_detail_field("Telefone do sócio 3", _contract_detail_value(row, columns, "telefone_socio_3")),
            _contract_detail_field("CPF do sócio 3", _contract_detail_value(row, columns, "cpf_socio_3")),
        ]
    )

    tracking_fields = "".join(
        [
            _contract_detail_field("Instagram", _contract_detail_value(row, columns, "instagram")),
            _contract_detail_field("LinkedIn", _contract_detail_value(row, columns, "linkedin")),
            _contract_detail_field("Vendedor", normalize_text(row.get("_vendedor", "")) or "Não informado"),
            _contract_detail_field("Status comercial", normalize_text(row.get("_status_grupo", "")) or "Não informado"),
            _contract_detail_field("Data do chamado", _contract_detail_value(row, columns, "data_chamado")),
            _contract_detail_field("Última atualização", _contract_detail_value(row, columns, "ultima_atualizacao")),
            _contract_detail_field("Observações", _contract_detail_value(row, columns, "observacoes"), full_width=True, long_text=True),
        ]
    )

    with st.container(key="contract_detail_edit_inline"):
        inline_space_col, inline_edit_col = st.columns([4.2, 1.15], gap="small")

        with inline_edit_col:
            if st.button("✏️ Editar dados", key=f"edit_contract_{sheet_row}_inline", use_container_width=True):
                st.session_state.edit_contract_sheet_row = int(sheet_row)
                st.rerun()

    render_html(
        f"""
        <div class="contract-detail-shell">
            <div class="registration-section">
                <div class="registration-section-title">DADOS DA EMPRESA</div>
                <div class="registration-section-text">Informações institucionais cadastradas na planilha.</div>
            </div>
            <div class="contract-detail-grid">{company_fields}</div>

            <div class="registration-section">
                <div class="registration-section-title">TELEFONES DA EMPRESA</div>
                <div class="registration-section-text">Contatos utilizados no acompanhamento comercial.</div>
            </div>
            <div class="contract-detail-grid three-columns">{phone_fields}</div>

            <div class="registration-section">
                <div class="registration-section-title">SÓCIOS E RESPONSÁVEIS</div>
                <div class="registration-section-text">Responsáveis vinculados à empresa.</div>
            </div>
            <div class="contract-detail-grid">{partner_fields}</div>

            <div class="registration-section">
                <div class="registration-section-title">REDES SOCIAIS E ACOMPANHAMENTO</div>
                <div class="registration-section-text">Informações comerciais e status atual do atendimento.</div>
            </div>
            <div class="contract-detail-grid">{tracking_fields}</div>
        </div>
        """
    )


def render_all_contracts_page(df: pd.DataFrame, columns: dict) -> None:
    apply_registration_css()

    selected_sheet_row = st.session_state.get("selected_contract_sheet_row")

    if selected_sheet_row:
        render_contract_detail_page(df, columns, int(selected_sheet_row))
        return

    render_html(
        """
        <div class="registration-header-card">
            <div class="registration-kicker">OPPI COMERCIAL • CADASTROS</div>
            <div class="registration-title">Empresas cadastradas</div>
            <div class="registration-subtitle">
                Consulte os nomes de todas as empresas cadastradas na planilha comercial.
            </div>
        </div>
        """
    )

    render_html('<div class="contracts-page-header-gap"></div>')

    valid_dates = df["_data_chamado"].dropna()

    if valid_dates.empty:
        date_max = date.today()
        date_min = date_max - timedelta(days=30)
    else:
        date_min = valid_dates.min().date()
        date_max = valid_dates.max().date()

    seller_options = sorted(
        [
            seller
            for seller in df["_vendedor"].dropna().astype(str).unique().tolist()
            if normalize_text(seller)
        ]
    )

    if "contracts_names_filter_seller" not in st.session_state:
        st.session_state.contracts_names_filter_seller = "Todos os vendedores"

    if "contracts_names_filter_status" not in st.session_state:
        st.session_state.contracts_names_filter_status = "Todos os status"

    if "contracts_names_filter_period" not in st.session_state:
        st.session_state.contracts_names_filter_period = (date_min, date_max)

    if "contracts_names_filter_search" not in st.session_state:
        st.session_state.contracts_names_filter_search = ""

    niche_options = sorted(
        [
            niche
            for niche in df["_nicho"].dropna().astype(str).unique().tolist()
            if normalize_text(niche)
        ],
        key=normalize_search_text,
    )

    state_options = sorted(
        [
            state
            for state in df["_estado"].dropna().astype(str).unique().tolist()
            if normalize_text(state)
        ],
        key=lambda value: (value == "Não identificado", value),
    )

    if "contracts_names_filter_niche" not in st.session_state:
        st.session_state.contracts_names_filter_niche = "Todos os nichos"

    if "contracts_names_filter_state" not in st.session_state:
        st.session_state.contracts_names_filter_state = "Todos os estados"

    with st.container(key="contracts_filters_aligned"):
        (
            filter_col_1,
            filter_col_2,
            filter_col_3,
            filter_col_4,
            filter_col_5,
            filter_col_6,
        ) = st.columns([1.05, 1.0, 1.12, 1.0, 1.0, 1.25], gap="small")

        with filter_col_1:
            selected_seller = st.selectbox(
                "Vendedor",
                ["Todos os vendedores"] + seller_options,
                key="contracts_names_filter_seller",
            )

        with filter_col_2:
            selected_status = st.selectbox(
                "Status",
                ["Todos os status"] + STATUS_OPTIONS,
                key="contracts_names_filter_status",
            )

        with filter_col_3:
            selected_period = st.date_input(
                "Período",
                value=st.session_state.contracts_names_filter_period,
                min_value=date_min,
                max_value=max(date_max, date.today()),
                key="contracts_names_filter_period",
            )

        with filter_col_4:
            selected_niche = st.selectbox(
                "Nichos",
                ["Todos os nichos"] + niche_options,
                key="contracts_names_filter_niche",
            )

        with filter_col_5:
            selected_state = st.selectbox(
                "Estados",
                ["Todos os estados"] + state_options,
                key="contracts_names_filter_state",
            )

        with filter_col_6:
            search_term = st.text_input(
                "Buscar empresa",
                placeholder="Digite o nome da empresa...",
                key="contracts_names_filter_search",
            )

    filtered_df = df.copy()

    if selected_seller != "Todos os vendedores":
        filtered_df = filtered_df[filtered_df["_vendedor"] == selected_seller].copy()

    if selected_status != "Todos os status":
        filtered_df = filtered_df[filtered_df.apply(lambda row: row_matches_status_filter(row, selected_status), axis=1)].copy()

    if selected_niche != "Todos os nichos":
        filtered_df = filtered_df[filtered_df["_nicho"] == selected_niche].copy()

    if selected_state != "Todos os estados":
        filtered_df = filtered_df[filtered_df["_estado"] == selected_state].copy()

    filtered_df = apply_period_filter(
        filtered_df,
        "_data_chamado",
        selected_period,
    )

    if normalize_text(search_term):
        term = normalize_search_text(search_term)
        filtered_df = filtered_df[
            filtered_df["_empresa"].apply(
                lambda value: term in normalize_search_text(value)
            )
        ].copy()

    names_df = filtered_df[["_empresa", "_sheet_row"]].copy()
    names_df["Empresa"] = names_df["_empresa"].apply(normalize_text)
    names_df = names_df[names_df["Empresa"] != ""].copy()

    # Ao entrar em Todos os cadastros, a lista começa pelos registros mais recentes
    # da planilha. A seta no cabeçalho alterna para ordem alfabética e permite voltar.
    requested_order = normalize_search_text(_query_param_value("order"))
    sort_mode = "alfabetica" if requested_order == "alfabetica" else "recentes"

    if sort_mode == "alfabetica":
        names_df = names_df.sort_values(
            "Empresa",
            key=lambda series: series.map(normalize_search_text),
        )
    else:
        names_df = names_df.sort_values("_sheet_row", ascending=False)

    render_html(
        f"""
        <div class="contracts-names-count-card">
            Exibindo <strong>{len(names_df)}</strong> empresa(s) cadastrada(s) na planilha.
        </div>
        """
    )

    if names_df.empty:
        st.info("Nenhuma empresa cadastrada foi encontrada com os filtros informados.")
        return

    with st.container(key="contracts_names_list"):
        navigation_token = normalize_text(st.session_state.get("navigation_session_token", ""))
        next_order = "alfabetica" if sort_mode == "recentes" else "recentes"
        sort_arrow = "↓" if sort_mode == "recentes" else "↑"
        sort_title = (
            "Mais recentes primeiro — clique para ordenar em ordem alfabética"
            if sort_mode == "recentes"
            else "Ordem alfabética — clique para voltar aos cadastros mais recentes"
        )
        session_query = f"&session={html.escape(navigation_token, quote=True)}" if navigation_token else ""

        render_html(
            f"""
            <div class="contracts-names-clickable-header">
                <span>Empresas cadastradas</span>
                <a
                    class="contracts-names-sort-toggle"
                    href="?page=cadastro&contracts=todos&order={next_order}{session_query}"
                    target="_self"
                    title="{html.escape(sort_title, quote=True)}"
                    aria-label="{html.escape(sort_title, quote=True)}"
                >{sort_arrow}</a>
            </div>
            """
        )

        for _, company_row in names_df.iterrows():
            sheet_row = int(company_row["_sheet_row"])
            company_name = normalize_text(company_row["Empresa"])

            if st.button(
                company_name,
                key=f"open_contract_detail_{sheet_row}",
                use_container_width=True,
            ):
                st.session_state.selected_contract_sheet_row = sheet_row
                st.session_state.edit_contract_sheet_row = None
                st.rerun()



# =========================================================
# CORREÇÕES FINAIS: POSIÇÃO DO CHAT E SETA DO MENU RECOLHIDO
# =========================================================
def apply_final_sidebar_toggle_override_css() -> None:
    """Força uma seta cinza clara e visível quando o menu lateral estiver recolhido."""
    render_html(
        """
        <style>
            /* Compatibilidade com versões diferentes do Streamlit. */
            [data-testid="collapsedControl"],
            [data-testid="stSidebarCollapsedControl"],
            button[data-testid="collapsedControl"],
            button[data-testid="stSidebarCollapsedControl"] {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                position: fixed !important;
                top: 18px !important;
                left: 14px !important;
                z-index: 2147483646 !important;
                pointer-events: auto !important;
                align-items: center !important;
                justify-content: center !important;
                width: 42px !important;
                min-width: 42px !important;
                height: 42px !important;
                min-height: 42px !important;
                padding: 0 !important;
                border-radius: 12px !important;
                border: 1px solid rgba(107,114,128,0.34) !important;
                background: #E5E7EB !important;
                background-color: #E5E7EB !important;
                color: #6B7280 !important;
                box-shadow: 0 8px 20px rgba(0,0,0,0.18) !important;
            }

            [data-testid="collapsedControl"] > button,
            [data-testid="stSidebarCollapsedControl"] > button,
            [data-testid="collapsedControl"] button,
            [data-testid="stSidebarCollapsedControl"] button {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                align-items: center !important;
                justify-content: center !important;
                width: 100% !important;
                height: 100% !important;
                padding: 0 !important;
                border: none !important;
                border-radius: 12px !important;
                background: #E5E7EB !important;
                background-color: #E5E7EB !important;
                color: #6B7280 !important;
                box-shadow: none !important;
            }

            /* Esconde a seta escura original e desenha uma seta cinza mais legível. */
            [data-testid="collapsedControl"] svg,
            [data-testid="stSidebarCollapsedControl"] svg,
            button[data-testid="collapsedControl"] svg,
            button[data-testid="stSidebarCollapsedControl"] svg {
                display: none !important;
            }

            [data-testid="collapsedControl"] > button::after,
            [data-testid="stSidebarCollapsedControl"] > button::after,
            button[data-testid="collapsedControl"]::after,
            button[data-testid="stSidebarCollapsedControl"]::after {
                content: "›" !important;
                display: block !important;
                color: #6B7280 !important;
                font-size: 32px !important;
                font-weight: 900 !important;
                line-height: 0.88 !important;
                transform: translateY(-1px) !important;
            }

            [data-testid="collapsedControl"]:hover,
            [data-testid="stSidebarCollapsedControl"]:hover,
            button[data-testid="collapsedControl"]:hover,
            button[data-testid="stSidebarCollapsedControl"]:hover,
            [data-testid="collapsedControl"] button:hover,
            [data-testid="stSidebarCollapsedControl"] button:hover {
                background: #F3F4F6 !important;
                background-color: #F3F4F6 !important;
                transform: scale(1.06) !important;
            }
        </style>
        """
    )


def apply_final_chat_layout_override_css() -> None:
    """Abaixa o chat e garante quatro conversas completas com rolagem interna."""
    render_html(
        """
        <style>
            .diagnostic-page-top-spacer {
                display: block !important;
                width: 100% !important;
                height: 46px !important;
                min-height: 46px !important;
                flex: 0 0 46px !important;
            }

            .st-key-diagnostic_contacts_panel,
            .st-key-diagnostic_chat_panel {
                margin-top: 0 !important;
                min-height: calc(100dvh - 46px) !important;
                height: calc(100dvh - 46px) !important;
                max-height: calc(100dvh - 46px) !important;
            }

            /* Quatro cards inteiros visíveis; os demais aparecem ao rolar a lista. */
            .st-key-diagnostic_contacts_list {
                display: block !important;
                flex: 0 0 360px !important;
                width: 100% !important;
                height: 360px !important;
                min-height: 360px !important;
                max-height: 360px !important;
                overflow-y: auto !important;
                overflow-x: hidden !important;
                padding: 0 4px 0 0 !important;
                margin: 0 !important;
                scrollbar-width: thin !important;
                scrollbar-color: rgba(169,28,255,0.74) rgba(255,255,255,0.58) !important;
            }

            .st-key-diagnostic_contacts_list > div[data-testid="stVerticalBlock"],
            .st-key-diagnostic_contacts_list div[data-testid="stVerticalBlock"] {
                display: block !important;
                height: auto !important;
                min-height: 0 !important;
                max-height: none !important;
                overflow: visible !important;
                padding: 0 !important;
                margin: 0 !important;
                gap: 0 !important;
            }

            .st-key-diagnostic_contacts_list div[data-testid="stElementContainer"] {
                margin: 0 !important;
                padding: 0 !important;
            }

            .st-key-diagnostic_contacts_list .stButton > button {
                width: calc(100% - 18px) !important;
                min-height: 78px !important;
                height: 78px !important;
                margin: 5px 9px !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar {
                width: 8px !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar-track {
                background: rgba(255,255,255,0.58) !important;
                border-radius: 999px !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar-thumb {
                border-radius: 999px !important;
                background: linear-gradient(180deg, #FF4BAA 0%, #A91CFF 100%) !important;
            }
        </style>
        """
    )

# =========================================================
# PÁGINA: PESOS E MEDIDAS

# =========================================================
# AJUSTES VISUAIS GLOBAIS DO MENU RECOLHIDO
# =========================================================
def apply_global_sidebar_toggle_css() -> None:
    """Mantém a seta de reabertura do menu clara e visível em todas as páginas."""
    render_html(
        """
        <style>
            [data-testid="collapsedControl"],
            [data-testid="stSidebarCollapsedControl"],
            button[data-testid="collapsedControl"],
            button[data-testid="stSidebarCollapsedControl"],
            [data-testid="collapsedControl"] > button,
            [data-testid="stSidebarCollapsedControl"] > button,
            [data-testid="collapsedControl"] button,
            [data-testid="stSidebarCollapsedControl"] button {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                align-items: center !important;
                justify-content: center !important;
                width: 40px !important;
                min-width: 40px !important;
                height: 40px !important;
                min-height: 40px !important;
                padding: 0 !important;
                border-radius: 11px !important;
                border: 1px solid rgba(75,85,99,0.28) !important;
                background: #D1D5DB !important;
                background-color: #D1D5DB !important;
                color: #4B5563 !important;
                box-shadow: 0 8px 20px rgba(0,0,0,0.20) !important;
                z-index: 1000002 !important;
                pointer-events: auto !important;
            }

            [data-testid="collapsedControl"] svg,
            [data-testid="stSidebarCollapsedControl"] svg,
            button[data-testid="collapsedControl"] svg,
            button[data-testid="stSidebarCollapsedControl"] svg,
            [data-testid="collapsedControl"] svg path,
            [data-testid="stSidebarCollapsedControl"] svg path {
                color: #4B5563 !important;
                fill: #4B5563 !important;
                stroke: #4B5563 !important;
                opacity: 1 !important;
            }

            [data-testid="collapsedControl"]:hover,
            [data-testid="stSidebarCollapsedControl"]:hover,
            button[data-testid="collapsedControl"]:hover,
            button[data-testid="stSidebarCollapsedControl"]:hover,
            [data-testid="collapsedControl"] button:hover,
            [data-testid="stSidebarCollapsedControl"] button:hover {
                background: #E5E7EB !important;
                background-color: #E5E7EB !important;
                color: #374151 !important;
                transform: scale(1.06) !important;
            }
        </style>
        """
    )


def apply_chat_position_and_scroll_css() -> None:
    """Desce o chat alguns pixels e exibe quatro conversas completas com rolagem interna."""
    render_html(
        """
        <style>
            .st-key-diagnostic_contacts_panel,
            .st-key-diagnostic_chat_panel {
                margin-top: 42px !important;
                min-height: calc(100dvh - 42px) !important;
                height: calc(100dvh - 42px) !important;
                max-height: calc(100dvh - 42px) !important;
            }

            .st-key-diagnostic_contacts_list {
                display: block !important;
                flex: 0 0 336px !important;
                width: 100% !important;
                height: 336px !important;
                min-height: 336px !important;
                max-height: 336px !important;
                overflow-y: auto !important;
                overflow-x: hidden !important;
                padding: 0 4px 0 0 !important;
                margin: 0 !important;
                scrollbar-width: thin !important;
                scrollbar-color: rgba(169,28,255,0.72) rgba(255,255,255,0.54) !important;
            }

            .st-key-diagnostic_contacts_list > div[data-testid="stVerticalBlock"],
            .st-key-diagnostic_contacts_list div[data-testid="stVerticalBlock"] {
                display: block !important;
                height: auto !important;
                min-height: 0 !important;
                max-height: none !important;
                overflow: visible !important;
                padding: 0 !important;
                margin: 0 !important;
            }

            .st-key-diagnostic_contacts_list div[data-testid="stElementContainer"] {
                margin: 0 !important;
                padding: 0 !important;
            }

            .st-key-diagnostic_contacts_list .stButton > button {
                width: calc(100% - 18px) !important;
                min-height: 76px !important;
                height: 76px !important;
                margin: 4px 9px !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar {
                width: 8px !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar-track {
                background: rgba(255,255,255,0.54) !important;
                border-radius: 999px !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar-thumb {
                border-radius: 999px !important;
                background: linear-gradient(180deg, #FF4BAA 0%, #A91CFF 100%) !important;
            }
        </style>
        """
    )

# =========================================================
PRICING_SCRIPT_VERSION = "pricing_v4_pdf_diagnostico"

OPPI_PRICING_INTRO = (
    "Olá! Vou ajudar você, vendedor, a elaborar a faixa de preço para este cliente. "
    "As perguntas são sempre as mesmas. Nas perguntas de peso, responda somente com o número da opção correspondente. "
    "No final, descreva brevemente o cliente, os pontos discutidos na reunião, os serviços desejados e os problemas apresentados."
)

OPPI_PRICING_STEPS = [
    {
        "id": "colaboradores",
        "title": "🔵 1. Quantidade de colaboradores do cliente",
        "question": "Quantos colaboradores o cliente possui atualmente?",
        "options": [
            "1 — Pequena: 1 a 5 colaboradores",
            "2 — Média: 6 a 15 colaboradores",
            "3 — Estruturada: 16 a 30 colaboradores",
            "4 — Operação grande: acima de 30 colaboradores",
        ],
        "example": "Responda somente com: 1, 2, 3 ou 4.",
        "weighted": True,
    },
    {
        "id": "setores",
        "title": "🟣 2. Setores que o cliente deseja organizar",
        "question": "Quantos setores ou áreas o cliente deseja organizar?",
        "options": [
            "1 — Simples: apenas um fluxo ou setor",
            "2 — Média: comercial + atendimento",
            "3 — Alta: comercial + operação + pós-venda",
            "4 — Complexa: múltiplas equipes, unidades ou setores",
        ],
        "example": "Responda somente com: 1, 2, 3 ou 4.",
        "weighted": True,
    },
    {
        "id": "processos",
        "title": "🔴 3. Quantidade de processos",
        "question": "Quantos processos precisam ser organizados ou integrados?",
        "options": [
            "1 — Um processo: apenas pipeline",
            "2 — Dois processos: pipeline + propostas",
            "3 — Três processos: pipeline + operação + acompanhamento",
            "4 — Quatro ou mais processos: fluxos completos integrados",
        ],
        "example": "Responda somente com: 1, 2, 3 ou 4.",
        "weighted": True,
    },
    {
        "id": "personalizacao",
        "title": "🟢 4. Nível de personalização",
        "question": "Qual é o nível de personalização necessário para atender o cliente?",
        "options": [
            "1 — Baixa: apenas identidade visual",
            "2 — Média: ajustes de etapas e campos",
            "3 — Alta: regras específicas",
            "4 — Muito alta: fluxos únicos ou complexos",
        ],
        "example": "Responda somente com: 1, 2, 3 ou 4.",
        "weighted": True,
    },
    {
        "id": "volume",
        "title": "🟠 5. Volume operacional",
        "question": "Qual é o volume operacional do cliente?",
        "options": [
            "1 — Baixo: poucos atendimentos ou pedidos",
            "2 — Médio: fluxo diário constante",
            "3 — Alto: grande quantidade diária",
            "4 — Muito alto: operação intensa ou múltiplas equipes",
        ],
        "example": "Responda somente com: 1, 2, 3 ou 4.",
        "weighted": True,
    },
    {
        "id": "impacto",
        "title": "⚫ 6. Impacto do caos operacional",
        "question": "Qual é o impacto atual da desorganização na empresa? Esta é a pergunta mais importante.",
        "options": [
            "1 — Baixo: pequena desorganização",
            "2 — Médio: perda ocasional de clientes",
            "3 — Alto: leads ou pedidos perdidos",
            "4 — Crítico: a empresa perdeu o controle da operação",
        ],
        "example": "Responda somente com: 1, 2, 3 ou 4.",
        "weighted": True,
    },
    {
        "id": "faturamento",
        "title": "💰 7. Faturamento do cliente",
        "question": "Você sabe o faturamento aproximado do cliente?",
        "options": [
            "Informe o faturamento aproximado mensal ou anual, caso saiba.",
            "Caso ainda não tenha essa informação, responda: não informado.",
        ],
        "example": "Exemplos de resposta: R$ 80 mil/mês; R$ 1 milhão/ano; não informado.",
        "weighted": False,
    },
    {
        "id": "resumo_cliente",
        "title": "📝 8. Resumo do cliente e da reunião",
        "question": "Comente sobre o cliente: me envie um resumo da ata de reunião, informe os serviços desejados e descreva os principais problemas apresentados.",
        "options": [
            "Inclua os pontos mais importantes identificados durante a conversa.",
            "Esse resumo será utilizado para indicar a solução Oppi mais adequada e gerar o PDF do diagnóstico.",
        ],
        "example": "Exemplos: precisa organizar o comercial; deseja acompanhar a operação; perde informações entre setores; quer automatizar propostas.",
        "weighted": False,
    },
]

OPPI_PRODUCT_PRICE_TABLE = {
    "Oppi Vision": {
        "pequena": ("R$ 3.000", "R$ 5.000", "1 mês"),
        "media": ("R$ 5.000", "R$ 8.000", "1 mês"),
        "estruturada": ("R$ 8.000", "R$ 15.000", "1 mês"),
    },
    "Oppi Flow": {
        "pequena": ("R$ 4.000", "R$ 7.000", "1 mês"),
        "media": ("R$ 7.000", "R$ 12.000", "1 mês"),
        "estruturada": ("R$ 12.000", "R$ 20.000", "1 mês"),
    },
    "Oppi Track": {
        "pequena": ("R$ 5.000", "R$ 8.000", "1 mês"),
        "media": ("R$ 8.000", "R$ 15.000", "1 mês"),
        "estruturada": ("R$ 15.000", "Sob consulta", "1 mês"),
    },
}

OPPI_ADDITIONAL_SERVICES_TABLE = {
    "Pequeno": {
        "contrato": "Equipe 20 - contratos digitais: R$ 791,00 à vista ou R$ 184,78/mês em até 6x",
        "disparos": "Essencial WhatsApp - até 100 envios: R$ 149,00/mês",
    },
    "Médio": {
        "contrato": "Equipe 80 - contratos digitais: R$ 3.175,00 à vista ou R$ 645,54/mês em até 6x",
        "disparos": "Crescimento WhatsApp - até 200 envios: R$ 249,00/mês",
    },
    "Premium": {
        "contrato": "Equipe 150 - contratos digitais: R$ 3.960,00 à vista ou R$ 771,84/mês em até 6x",
        "disparos": "Profissional WhatsApp - até 300 envios: R$ 349,00/mês",
    },
    "Enterprise": {
        "contrato": "Equipe 150+ - contratos digitais: sob consulta conforme volume",
        "disparos": "Profissional WhatsApp ou pacote personalizado: sob consulta conforme volume",
    },
}


def _pricing_additional_services(profile: str) -> str:
    services = OPPI_ADDITIONAL_SERVICES_TABLE.get(profile, OPPI_ADDITIONAL_SERVICES_TABLE["Médio"])
    return f"{services['contrato']} | {services['disparos']}"


def _pricing_question_text(step: dict) -> str:
    options_text = "\n".join(step["options"])
    return (
        f"{step['title']}\n\n"
        f"{step['question']}\n\n"
        f"{options_text}\n\n"
        f"{step['example']}"
    )


def apply_chat_css() -> None:
    render_html(
        """
        <style>
            /* Remove a faixa vazia superior e inferior somente na página do chat. */
            header[data-testid="stHeader"] {
                display: block !important;
                position: fixed !important;
                top: 0 !important;
                left: 0 !important;
                right: 0 !important;
                height: 0 !important;
                min-height: 0 !important;
                background: transparent !important;
                z-index: 999998 !important;
                pointer-events: none !important;
            }

            header[data-testid="stHeader"] [data-testid="collapsedControl"],
            header[data-testid="stHeader"] [data-testid="stSidebarCollapsedControl"],
            header[data-testid="stHeader"] button {
                pointer-events: auto !important;
            }

            [data-testid="stAppViewContainer"],
            [data-testid="stMain"],
            .main {
                height: 100dvh !important;
                min-height: 100dvh !important;
                max-height: 100dvh !important;
                overflow: hidden !important;
            }

            [data-testid="stMain"] {
                padding: 0 !important;
                margin: 0 !important;
            }

            /* Tela integral: mantém somente o menu lateral e utiliza todo o restante da página. */
            .block-container,
            [data-testid="stMainBlockContainer"] {
                max-width: none !important;
                width: 100% !important;
                height: 100dvh !important;
                min-height: 100dvh !important;
                max-height: 100dvh !important;
                padding: 0 !important;
                margin: 0 !important;
                overflow: hidden !important;
            }

            [data-testid="stMain"] > div,
            [data-testid="stMainBlockContainer"] > div,
            [data-testid="stMainBlockContainer"] div[data-testid="stVerticalBlock"] {
                margin-top: 0 !important;
                margin-bottom: 0 !important;
            }

            [data-testid="stMainBlockContainer"] > div[data-testid="stVerticalBlock"] {
                gap: 0 !important;
                height: 100dvh !important;
                max-height: 100dvh !important;
                overflow: hidden !important;
            }

            [data-testid="stMainBlockContainer"] div[data-testid="stHorizontalBlock"] {
                gap: 0 !important;
                margin: 0 !important;
            }

            [data-testid="stMainBlockContainer"] div[data-testid="column"] {
                padding: 0 !important;
                margin: 0 !important;
            }

            .st-key-diagnostic_contacts_panel,
            .st-key-diagnostic_chat_panel {
                min-height: calc(100dvh - 24px) !important;
                height: calc(100dvh - 24px) !important;
                max-height: calc(100dvh - 24px) !important;
                margin: 24px 0 0 0 !important;
                border-radius: 0 !important;
                box-shadow: none !important;
                overflow: hidden !important;
            }

            /* Cantos claros e painel de conversas em cinza, seguindo o menu lateral. */
            .st-key-diagnostic_contacts_panel {
                border: none !important;
                border-right: 1px solid rgba(80,69,105,0.18) !important;
                background:
                    radial-gradient(circle at top left, rgba(255,255,255,0.98) 0%, rgba(255,255,255,0.00) 34%),
                    radial-gradient(circle at bottom right, rgba(208,212,223,0.72) 0%, rgba(208,212,223,0.00) 38%),
                    linear-gradient(180deg, #F7F8FC 0%, #ECEEF4 40%, #DCE0E9 74%, #CED3DE 100%) !important;
            }

            .st-key-diagnostic_chat_panel {
                display: flex !important;
                flex-direction: column !important;
                border: none !important;
                background: #E4E7EE !important;
            }

            .oppi-chat-contact-header,
            .oppi-chat-window-header {
                min-height: 76px;
                display: flex;
                align-items: center;
                padding: 14px 18px;
                border-bottom: 1px solid rgba(80,69,105,0.14);
                background: rgba(255,255,255,0.92);
            }

            .oppi-chat-contact-header {
                justify-content: space-between;
            }

            .oppi-chat-contact-title,
            .oppi-chat-window-name {
                color: #211A30;
                font-size: 1rem;
                font-weight: 900;
            }

            .oppi-chat-contact-subtitle {
                margin-top: 4px;
                color: rgba(33,26,48,0.64);
                font-size: 0.75rem;
            }

            .oppi-chat-avatar {
                width: 46px;
                height: 46px;
                min-width: 46px;
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
                color: #FFFFFF;
                font-size: 0.92rem;
                font-weight: 900;
                background: linear-gradient(135deg, #FF4BAA 0%, #9C19FF 100%);
                box-shadow: 0 10px 22px rgba(169,28,255,0.18);
            }

            .oppi-chat-window-person {
                display: flex;
                align-items: center;
                gap: 12px;
            }

            .oppi-chat-window-status {
                margin-top: 3px;
                color: #289B50;
                font-size: 0.78rem;
                font-weight: 800;
            }

            .oppi-chat-window-status::before {
                content: "";
                display: inline-block;
                width: 8px;
                height: 8px;
                margin-right: 6px;
                border-radius: 50%;
                background: #55DF7D;
                box-shadow: 0 0 12px rgba(85,223,125,0.56);
            }

            .oppi-chat-messages {
                min-height: calc(100dvh - 266px) !important;
                max-height: calc(100dvh - 266px) !important;
                overflow-y: auto;
                padding: 20px 22px 16px 22px;
                background:
                    radial-gradient(circle at 76% 18%, rgba(169,28,255,0.08), transparent 31%),
                    linear-gradient(180deg, #E8EAF0 0%, #DCE0E9 100%);
                scrollbar-width: thin;
                scrollbar-color: rgba(169,28,255,0.48) rgba(255,255,255,0.42);
            }

            .oppi-chat-day {
                display: flex;
                justify-content: center;
                margin-bottom: 16px;
            }

            .oppi-chat-day span {
                padding: 5px 12px;
                border-radius: 999px;
                color: rgba(33,26,48,0.66);
                font-size: 0.73rem;
                background: rgba(255,255,255,0.78);
                border: 1px solid rgba(80,69,105,0.10);
            }

            .oppi-chat-message-row {
                display: flex;
                margin: 9px 0;
            }

            .oppi-chat-message-row.assistant {
                justify-content: flex-start;
            }

            .oppi-chat-message-row.user {
                justify-content: flex-end;
            }

            .oppi-chat-bubble {
                max-width: min(78%, 760px);
                padding: 12px 14px 9px 14px;
                border-radius: 17px;
                font-size: 0.89rem;
                line-height: 1.48;
                box-shadow: 0 8px 18px rgba(45,36,70,0.08);
                word-break: break-word;
            }

            /* Perguntas do robô: caixas brancas com letras escuras. */
            .oppi-chat-message-row.assistant .oppi-chat-bubble {
                color: #211A30;
                border-bottom-left-radius: 5px;
                background: #FFFFFF;
                border: 1px solid rgba(80,69,105,0.12);
            }

            /* Respostas do vendedor: mantém o gradiente rosa e roxo. */
            .oppi-chat-message-row.user .oppi-chat-bubble {
                color: #FFFFFF;
                border-bottom-right-radius: 5px;
                background: linear-gradient(135deg, #FF4BAA 0%, #A91CFF 100%);
                border: 1px solid rgba(255,255,255,0.24);
            }

            .oppi-chat-bubble-time {
                display: block;
                text-align: right;
                margin-top: 4px;
                color: rgba(33,26,48,0.48);
                font-size: 0.67rem;
            }

            .oppi-chat-message-row.user .oppi-chat-bubble-time {
                color: rgba(255,255,255,0.72);
            }

            .oppi-chat-progress-wrap {
                margin: 0;
                padding: 10px 18px;
                border-top: 1px solid rgba(80,69,105,0.12);
                border-bottom: 1px solid rgba(80,69,105,0.12);
                background: rgba(255,255,255,0.86);
            }

            .oppi-chat-progress-label {
                color: rgba(33,26,48,0.72);
                font-size: 0.75rem;
                font-weight: 800;
                margin-bottom: 7px;
            }

            .oppi-chat-progress-bar {
                width: 100%;
                height: 7px;
                overflow: hidden;
                border-radius: 999px;
                background: rgba(80,69,105,0.12);
            }

            .oppi-chat-progress-fill {
                height: 100%;
                border-radius: 999px;
                background: linear-gradient(90deg, #FF4BAA 0%, #A91CFF 100%);
                box-shadow: 0 0 18px rgba(169,28,255,0.18);
            }

            .st-key-diagnostic_contacts_panel div[data-testid="stTextInput"] {
                padding: 0 14px !important;
                margin: 12px 0 8px 0 !important;
            }

            .st-key-diagnostic_contacts_panel div[data-testid="stTextInput"] div[data-baseweb="input"] > div {
                min-height: 44px !important;
                height: 44px !important;
                border-radius: 999px !important;
                border: 1px solid rgba(80,69,105,0.14) !important;
                background: rgba(255,255,255,0.88) !important;
                box-shadow: none !important;
            }

            .st-key-diagnostic_contacts_panel div[data-testid="stTextInput"] input {
                min-height: 44px !important;
                height: 44px !important;
                line-height: 44px !important;
                color: #211A30 !important;
                -webkit-text-fill-color: #211A30 !important;
            }

            .st-key-diagnostic_contacts_panel div[data-testid="stTextInput"] input::placeholder {
                color: rgba(33,26,48,0.48) !important;
                -webkit-text-fill-color: rgba(33,26,48,0.48) !important;
            }

            /* Exibe quatro conversas completas por vez e mantém rolagem interna para as demais. */
            .st-key-diagnostic_contacts_list {
                height: 344px !important;
                max-height: 344px !important;
                min-height: 344px !important;
                overflow: hidden !important;
                padding: 2px 0 4px 0 !important;
            }

            .st-key-diagnostic_contacts_list > div[data-testid="stVerticalBlock"],
            .st-key-diagnostic_contacts_list div[data-testid="stVerticalBlock"] {
                height: 338px !important;
                max-height: 338px !important;
                min-height: 338px !important;
                overflow-y: auto !important;
                overflow-x: hidden !important;
                padding: 2px 5px 4px 0 !important;
                gap: 0 !important;
                scrollbar-width: thin;
                scrollbar-color: rgba(169,28,255,0.48) rgba(255,255,255,0.44);
            }

            .st-key-diagnostic_contacts_list div[data-testid="stElementContainer"] {
                margin: 0 !important;
                padding: 0 !important;
            }

            .st-key-diagnostic_contacts_list div[data-testid="stVerticalBlock"]::-webkit-scrollbar {
                width: 7px;
            }

            .st-key-diagnostic_contacts_list div[data-testid="stVerticalBlock"]::-webkit-scrollbar-track {
                background: rgba(255,255,255,0.44);
            }

            .st-key-diagnostic_contacts_list div[data-testid="stVerticalBlock"]::-webkit-scrollbar-thumb {
                border-radius: 999px;
                background: linear-gradient(180deg, #FF4BAA, #A91CFF);
            }

            .st-key-diagnostic_contacts_list .stButton > button {
                width: calc(100% - 18px) !important;
                min-height: 76px !important;
                height: 76px !important;
                margin: 4px 9px !important;
                padding: 8px 12px !important;
                justify-content: flex-start !important;
                border: 1px solid rgba(80,69,105,0.08) !important;
                border-radius: 14px !important;
                text-align: left !important;
                color: #211A30 !important;
                font-size: 0.80rem !important;
                background: rgba(255,255,255,0.36) !important;
                box-shadow: none !important;
                transform-origin: center center !important;
                transition: transform 0.20s ease, box-shadow 0.20s ease, border-color 0.20s ease, background 0.20s ease !important;
            }

            .st-key-diagnostic_contacts_list .stButton > button:hover {
                transform: scale(1.025) !important;
                border-color: rgba(255,75,170,0.52) !important;
                background: rgba(255,255,255,0.82) !important;
                box-shadow: 0 10px 22px rgba(169,28,255,0.14) !important;
            }

            .st-key-diagnostic_contacts_list .stButton > button[kind="primary"] {
                color: #FFFFFF !important;
                border-color: rgba(255,75,170,0.60) !important;
                background: linear-gradient(90deg, #FF4BAA 0%, #A91CFF 100%) !important;
            }

            .st-key-diagnostic_chat_toolbar {
                padding: 8px 14px 0 14px;
                background: rgba(255,255,255,0.86);
            }

            .st-key-diagnostic_chat_toolbar .stButton > button {
                min-height: 40px !important;
                border-radius: 999px !important;
                color: #FFFFFF !important;
                font-size: 0.76rem !important;
                font-weight: 850 !important;
                background: linear-gradient(90deg, #FF4BAA 0%, #A91CFF 100%) !important;
                border: 1px solid rgba(255,255,255,0.20) !important;
                box-shadow: 0 8px 18px rgba(169,28,255,0.18) !important;
                transform-origin: center center !important;
                transition: transform 0.20s ease, filter 0.20s ease, box-shadow 0.20s ease !important;
            }

            .st-key-diagnostic_chat_toolbar .stButton > button:hover {
                transform: scale(1.035) !important;
                filter: brightness(1.07) !important;
                border-color: rgba(255,255,255,0.34) !important;
                background: linear-gradient(90deg, #FF4BAA 0%, #A91CFF 100%) !important;
                box-shadow: 0 12px 24px rgba(169,28,255,0.28) !important;
            }

            .st-key-diagnostic_chat_toolbar [data-testid="stDownloadButton"] > button,
            .st-key-diagnostic_chat_toolbar [data-testid="stDownloadButton"] > a {
                min-height: 40px !important;
                width: 100% !important;
                border-radius: 999px !important;
                color: #FFFFFF !important;
                font-size: 0.76rem !important;
                font-weight: 850 !important;
                background: linear-gradient(90deg, #FF4BAA 0%, #A91CFF 100%) !important;
                border: 1px solid rgba(255,255,255,0.20) !important;
                box-shadow: 0 8px 18px rgba(169,28,255,0.18) !important;
                transform-origin: center center !important;
                transition: transform 0.20s ease, filter 0.20s ease, box-shadow 0.20s ease !important;
            }

            .st-key-diagnostic_chat_toolbar [data-testid="stDownloadButton"] > button:hover,
            .st-key-diagnostic_chat_toolbar [data-testid="stDownloadButton"] > a:hover {
                transform: scale(1.035) !important;
                filter: brightness(1.07) !important;
                box-shadow: 0 12px 24px rgba(169,28,255,0.28) !important;
            }

            .st-key-diagnostic_chat_form {
                padding: 9px 14px 12px 14px;
                background: rgba(255,255,255,0.86);
            }

            .st-key-diagnostic_chat_form [data-testid="stForm"] {
                display: block !important;
                margin: 0 !important;
                padding: 0 !important;
                max-width: none !important;
                border: none !important;
                border-radius: 0 !important;
                background: transparent !important;
                box-shadow: none !important;
            }

            .st-key-diagnostic_chat_form [data-testid="stForm"]::before,
            .st-key-diagnostic_chat_form [data-testid="stForm"]::after {
                display: none !important;
            }

            .st-key-diagnostic_chat_form div[data-testid="stTextInput"] div[data-baseweb="input"] > div {
                min-height: 50px !important;
                height: 50px !important;
                border-radius: 999px !important;
                border: 1px solid rgba(80,69,105,0.14) !important;
                background: #FFFFFF !important;
                box-shadow: none !important;
            }

            .st-key-diagnostic_chat_form div[data-testid="stTextInput"] input {
                min-height: 50px !important;
                height: 50px !important;
                line-height: 50px !important;
                padding: 0 18px !important;
                color: #211A30 !important;
                -webkit-text-fill-color: #211A30 !important;
            }

            .st-key-diagnostic_chat_form div[data-testid="stTextInput"] input::placeholder {
                color: rgba(33,26,48,0.46) !important;
                -webkit-text-fill-color: rgba(33,26,48,0.46) !important;
            }

            .st-key-diagnostic_chat_form .stButton > button,
            .st-key-diagnostic_chat_form button[kind="secondaryFormSubmit"] {
                min-height: 50px !important;
                height: 50px !important;
                margin-top: 0 !important;
                border-radius: 999px !important;
                border: none !important;
                color: #FFFFFF !important;
                background: linear-gradient(135deg, #FF4BAA 0%, #A91CFF 100%) !important;
                box-shadow: 0 10px 22px rgba(169,28,255,0.18) !important;
            }

            .st-key-diagnostic_chat_form .stButton > button:hover,
            .st-key-diagnostic_chat_form button[kind="secondaryFormSubmit"]:hover {
                transform: scale(1.025) !important;
            }

            /* Ajuste final da lista lateral: quatro conversas completas visíveis e rolagem interna para as demais. */
            .st-key-diagnostic_contacts_list {
                height: 308px !important;
                min-height: 308px !important;
                max-height: 308px !important;
                overflow-y: auto !important;
                overflow-x: hidden !important;
                padding: 0 4px 0 0 !important;
                scrollbar-width: thin !important;
                scrollbar-color: rgba(169,28,255,0.62) rgba(255,255,255,0.50) !important;
            }

            .st-key-diagnostic_contacts_list > div[data-testid="stVerticalBlock"],
            .st-key-diagnostic_contacts_list div[data-testid="stVerticalBlock"] {
                height: auto !important;
                min-height: 0 !important;
                max-height: none !important;
                overflow: visible !important;
                padding: 0 !important;
                gap: 0 !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar {
                width: 8px !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar-track {
                background: rgba(255,255,255,0.50) !important;
                border-radius: 999px !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar-thumb {
                border-radius: 999px !important;
                background: linear-gradient(180deg, #FF4BAA 0%, #A91CFF 100%) !important;
            }

            .st-key-diagnostic_contacts_list .stButton > button {
                min-height: 69px !important;
                height: 69px !important;
                margin: 4px 9px !important;
            }

            /* Quando o menu lateral for fechado, mantém a seta de reabertura visível em cinza. */
            [data-testid="collapsedControl"],
            [data-testid="stSidebarCollapsedControl"] {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                position: fixed !important;
                top: 14px !important;
                left: 12px !important;
                z-index: 999999 !important;
                align-items: center !important;
                justify-content: center !important;
                width: 36px !important;
                height: 36px !important;
                border-radius: 10px !important;
                border: 1px solid rgba(255,255,255,0.15) !important;
                background: rgba(156,163,175,0.92) !important;
                box-shadow: 0 8px 20px rgba(0,0,0,0.22) !important;
            }

            [data-testid="collapsedControl"] button,
            [data-testid="stSidebarCollapsedControl"] button {
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
                width: 100% !important;
                height: 100% !important;
                padding: 0 !important;
                border: none !important;
                background: transparent !important;
                box-shadow: none !important;
            }

            [data-testid="collapsedControl"] svg,
            [data-testid="stSidebarCollapsedControl"] svg {
                color: #374151 !important;
                fill: #374151 !important;
                stroke: #374151 !important;
                opacity: 1 !important;
            }

            [data-testid="collapsedControl"]:hover,
            [data-testid="stSidebarCollapsedControl"]:hover {
                background: rgba(209,213,219,0.98) !important;
                transform: scale(1.06) !important;
            }

            /* AJUSTE FINAL: mantém o menu lateral funcional e mostra exatamente quatro conversas completas. */
            .st-key-diagnostic_contacts_panel {
                display: flex !important;
                flex-direction: column !important;
                min-height: calc(100dvh - 24px) !important;
                height: calc(100dvh - 24px) !important;
                max-height: calc(100dvh - 24px) !important;
                overflow: hidden !important;
            }

            .st-key-diagnostic_contacts_panel > div[data-testid="stVerticalBlock"] {
                display: flex !important;
                flex-direction: column !important;
                min-height: 0 !important;
                height: 100% !important;
                overflow: hidden !important;
            }

            .st-key-diagnostic_contacts_list {
                display: block !important;
                flex: 0 0 336px !important;
                height: 336px !important;
                min-height: 336px !important;
                max-height: 336px !important;
                overflow-y: auto !important;
                overflow-x: hidden !important;
                padding: 0 4px 0 0 !important;
                margin: 0 !important;
                scrollbar-width: thin !important;
                scrollbar-color: rgba(169,28,255,0.72) rgba(255,255,255,0.54) !important;
            }

            .st-key-diagnostic_contacts_list > div[data-testid="stVerticalBlock"],
            .st-key-diagnostic_contacts_list div[data-testid="stVerticalBlock"] {
                display: block !important;
                height: auto !important;
                min-height: 0 !important;
                max-height: none !important;
                overflow: visible !important;
                padding: 0 !important;
                margin: 0 !important;
            }

            .st-key-diagnostic_contacts_list div[data-testid="stElementContainer"] {
                margin: 0 !important;
                padding: 0 !important;
            }

            .st-key-diagnostic_contacts_list .stButton > button {
                width: calc(100% - 18px) !important;
                min-height: 76px !important;
                height: 76px !important;
                margin: 4px 9px !important;
            }

            /* A seta para reabrir o menu continua visível mesmo com o sidebar recolhido. */
            [data-testid="collapsedControl"],
            [data-testid="stSidebarCollapsedControl"] {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                position: fixed !important;
                top: 18px !important;
                left: 12px !important;
                z-index: 1000001 !important;
                pointer-events: auto !important;
                align-items: center !important;
                justify-content: center !important;
                width: 40px !important;
                height: 40px !important;
                border-radius: 11px !important;
                border: 1px solid rgba(55,65,81,0.20) !important;
                background: #D1D5DB !important;
                box-shadow: 0 8px 20px rgba(0,0,0,0.20) !important;
            }

            [data-testid="collapsedControl"] button,
            [data-testid="stSidebarCollapsedControl"] button {
                display: flex !important;
                width: 100% !important;
                height: 100% !important;
                align-items: center !important;
                justify-content: center !important;
                padding: 0 !important;
                pointer-events: auto !important;
                background: transparent !important;
                border: none !important;
                box-shadow: none !important;
            }

            [data-testid="collapsedControl"] svg,
            [data-testid="stSidebarCollapsedControl"] svg {
                color: #4B5563 !important;
                fill: #4B5563 !important;
                stroke: #4B5563 !important;
                opacity: 1 !important;
            }

            [data-testid="collapsedControl"]:hover,
            [data-testid="stSidebarCollapsedControl"]:hover {
                background: #E5E7EB !important;
                transform: scale(1.06) !important;
            }

            @media (max-width: 980px) {
                .st-key-diagnostic_contacts_panel,
                .st-key-diagnostic_chat_panel {
                    min-height: auto !important;
                    height: auto !important;
                }

                .oppi-chat-messages {
                    min-height: 420px !important;
                    max-height: 420px !important;
                }
            }
        </style>
        """
    )


def _diagnostic_initials(company_name: str) -> str:
    words = [word for word in normalize_text(company_name).split() if word]

    if not words:
        return "OP"

    if len(words) == 1:
        return words[0][:2].upper()

    return (words[0][0] + words[1][0]).upper()


def _diagnostic_now() -> str:
    return pd.Timestamp.now(tz="America/Sao_Paulo").strftime("%H:%M")


def _pricing_safe_key_fragment(value: str) -> str:
    """Cria um fragmento de key único e estável para evitar StreamlitDuplicateElementKey."""
    clean = normalize_search_text(value)
    clean = re.sub(r"[^a-z0-9]+", "_", clean).strip("_") or "empresa"
    unique = uuid.uuid5(uuid.NAMESPACE_DNS, normalize_text(value)).hex[:10]
    return f"{clean[:42]}_{unique}"


def _diagnostic_get_threads() -> dict:
    key = f"oppi_pricing_threads_{PRICING_SCRIPT_VERSION}"

    if key not in st.session_state:
        st.session_state[key] = {}

    return st.session_state[key]


def _diagnostic_get_progress() -> dict:
    key = f"oppi_pricing_progress_{PRICING_SCRIPT_VERSION}"

    if key not in st.session_state:
        st.session_state[key] = {}

    return st.session_state[key]


def _diagnostic_get_answers() -> dict:
    key = f"oppi_pricing_answers_{PRICING_SCRIPT_VERSION}"

    if key not in st.session_state:
        st.session_state[key] = {}

    return st.session_state[key]


def _pricing_extract_explicit_option(answer: str) -> Optional[int]:
    normalized = normalize_search_text(answer)
    patterns = [
        r"^(?:opcao|opção|peso)?\s*([1-4])(?:\s|$|[-—:])",
        r"\b(?:opcao|opção|peso)\s*([1-4])\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, normalized)

        if match:
            return int(match.group(1))

    return None


def _pricing_extract_numbers(answer: str) -> list[int]:
    return [int(value) for value in re.findall(r"\d+", normalize_text(answer))]


def _pricing_weight_from_answer(step_id: str, answer: str) -> Optional[int]:
    explicit_option = _pricing_extract_explicit_option(answer)

    if explicit_option:
        return explicit_option

    normalized = normalize_search_text(answer)
    numbers = _pricing_extract_numbers(answer)

    if step_id == "colaboradores" and numbers:
        collaborators = numbers[0]

        if collaborators <= 5:
            return 1
        if collaborators <= 15:
            return 2
        if collaborators <= 30:
            return 3
        return 4

    if step_id == "setores":
        if any(term in normalized for term in ["multi", "unidade", "varios setores", "vários setores", "complexa"]):
            return 4
        if "operacao" in normalized or "pos-venda" in normalized or "pós-venda" in normalized:
            return 3
        if "atendimento" in normalized and "comercial" in normalized:
            return 2
        if any(term in normalized for term in ["um setor", "1 setor", "apenas um", "simples"]):
            return 1

    if step_id == "processos":
        if numbers:
            quantity = numbers[0]
            return 4 if quantity >= 4 else max(1, quantity)
        if any(term in normalized for term in ["fluxos completos", "integrados", "quatro", "4+"]):
            return 4
        if "acompanhamento" in normalized or "operacao" in normalized:
            return 3
        if "proposta" in normalized:
            return 2
        if "pipeline" in normalized:
            return 1

    if step_id == "personalizacao":
        if "muito alta" in normalized or "complex" in normalized or "unico" in normalized or "único" in normalized:
            return 4
        if "alta" in normalized or "regra" in normalized:
            return 3
        if "media" in normalized or "média" in normalized or "etapa" in normalized or "campo" in normalized:
            return 2
        if "baixa" in normalized or "visual" in normalized:
            return 1

    if step_id == "volume":
        if "muito alto" in normalized or "intensa" in normalized or "multi" in normalized:
            return 4
        if "alto" in normalized or "grande" in normalized:
            return 3
        if "medio" in normalized or "médio" in normalized or "constante" in normalized:
            return 2
        if "baixo" in normalized or "pouco" in normalized:
            return 1

    if step_id == "impacto":
        if "critico" in normalized or "crítico" in normalized or "perdeu controle" in normalized:
            return 4
        if "alto" in normalized or "lead" in normalized or "pedido" in normalized:
            return 3
        if "medio" in normalized or "médio" in normalized or "ocasional" in normalized:
            return 2
        if "baixo" in normalized or "pequena" in normalized:
            return 1

    return None


def _pricing_profile(total_score: int) -> str:
    if total_score <= 10:
        return "Pequeno"
    if total_score <= 15:
        return "Médio"
    if total_score <= 20:
        return "Premium"
    return "Enterprise"


def _pricing_product(answer_map: dict) -> str:
    combined_text = normalize_search_text(
        " | ".join(normalize_text(item.get("answer")) for item in answer_map.values())
    )
    sectors_weight = int(answer_map.get("setores", {}).get("weight") or 0)
    processes_weight = int(answer_map.get("processos", {}).get("weight") or 0)
    volume_weight = int(answer_map.get("volume", {}).get("weight") or 0)

    operational_terms = [
        "operacao",
        "operacional",
        "pos-venda",
        "pós-venda",
        "pedido",
        "acompanhamento",
        "multi equipe",
        "multiequipe",
        "unidade",
    ]

    if any(term in combined_text for term in operational_terms) or sectors_weight >= 3 or processes_weight >= 3 or volume_weight >= 4:
        return "Oppi Track"

    commercial_terms = ["comercial", "pipeline", "proposta", "atendimento", "lead"]

    if any(term in combined_text for term in commercial_terms) or sectors_weight >= 2 or processes_weight >= 2:
        return "Oppi Flow"

    return "Oppi Vision"


def _pricing_company_size(answer_map: dict) -> str:
    collaborators_weight = int(answer_map.get("colaboradores", {}).get("weight") or 1)

    if collaborators_weight <= 1:
        return "pequena"
    if collaborators_weight == 2:
        return "media"
    return "estruturada"


def _pricing_result_message(company_name: str) -> str:
    answer_map = _diagnostic_get_answers().get(company_name, {})
    weights = [
        int(answer_map.get(step["id"], {}).get("weight") or 0)
        for step in OPPI_PRICING_STEPS
        if step["weighted"]
    ]
    total_score = sum(weights)
    profile = _pricing_profile(total_score)
    product = _pricing_product(answer_map)
    company_size = _pricing_company_size(answer_map)
    price_from, price_to, ideal_term = OPPI_PRODUCT_PRICE_TABLE[product][company_size]
    ideal_term = "1 mês"
    additional_services = _pricing_additional_services(profile)
    revenue = normalize_text(answer_map.get("faturamento", {}).get("answer")) or "Não informado"
    meeting_summary = normalize_text(answer_map.get("resumo_cliente", {}).get("answer")) or "Não informado"

    if price_to == "Sob consulta":
        pricing_text = f"a partir de {price_from}"
    else:
        pricing_text = f"entre {price_from} e {price_to}"

    weights_text = " + ".join(str(weight) for weight in weights)

    return (
        "✅ Diagnóstico concluído.\n\n"
        f"Soma dos pesos: {weights_text} = {total_score}.\n"
        f"Perfil do projeto: {profile}.\n"
        f"Solução sugerida: {product}.\n"
        f"Prazo ideal: {ideal_term}.\n"
        f"Faturamento informado: {revenue}.\n"
        f"Serviços adicionais sugeridos: {additional_services}.\n\n"
        f"Resumo registrado: {meeting_summary}\n\n"
        f"Pelo que vi aqui, o valor ficaria {pricing_text}. Quanto você deseja gerar a proposta?\n\n"
        "Exemplos de resposta: R$ 8.500; R$ 10.000; R$ 12.000; sob consulta. "
        "Depois de confirmar o valor, utilize o botão Gerar PDF do diagnóstico."
    )



def _pricing_pdf_safe_filename(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", normalize_text(value))
    normalized = "".join(character for character in normalized if not unicodedata.combining(character))
    normalized = re.sub(r"[^A-Za-z0-9_-]+", "_", normalized).strip("_")
    return normalized or "cliente"


def _pricing_get_confirmed_value(company_name: str) -> str:
    answer_map = _diagnostic_get_answers().get(company_name, {})
    return normalize_text(answer_map.get("valor_proposta", {}).get("answer"))


def _pricing_report_summary(company_name: str) -> dict:
    answer_map = _diagnostic_get_answers().get(company_name, {})
    weights = [
        int(answer_map.get(step["id"], {}).get("weight") or 0)
        for step in OPPI_PRICING_STEPS
        if step["weighted"]
    ]
    total_score = sum(weights)
    profile = _pricing_profile(total_score)
    product = _pricing_product(answer_map)
    company_size = _pricing_company_size(answer_map)
    price_from, price_to, ideal_term = OPPI_PRODUCT_PRICE_TABLE[product][company_size]
    ideal_term = "1 mês"

    if price_to == "Sob consulta":
        suggested_price = f"A partir de {price_from}"
    else:
        suggested_price = f"{price_from} a {price_to}"

    return {
        "answer_map": answer_map,
        "weights": weights,
        "total_score": total_score,
        "profile": profile,
        "product": product,
        "company_size": company_size,
        "suggested_price": suggested_price,
        "ideal_term": ideal_term,
        "additional_services": _pricing_additional_services(profile),
        "confirmed_value": _pricing_get_confirmed_value(company_name) or "Não informado",
    }


def _pricing_company_registration_data(df: pd.DataFrame, columns: dict, company_name: str) -> list[tuple[str, str]]:
    selected_rows = df[df["_empresa"].astype(str) == normalize_text(company_name)].copy()

    if selected_rows.empty:
        return [
            ("Nome da empresa", normalize_text(company_name) or "Não informado"),
            ("Telefone", "Não informado"),
            ("CNPJ", "Não informado"),
            ("Endereço", "Não informado"),
        ]

    if "_sheet_row" in selected_rows.columns:
        selected_rows = selected_rows.sort_values("_sheet_row", ascending=False)

    row = selected_rows.iloc[0]

    def value(column_key: str, fallback: str = "Não informado") -> str:
        column_name = columns.get(column_key)
        raw_value = normalize_text(row.get(column_name, "")) if column_name else ""
        return raw_value or fallback

    return [
        ("Nome da empresa", value("empresa", normalize_text(company_name) or "Não informado")),
        ("Telefone", value("telefone_b2b")),
        ("CNPJ", value("cnpj")),
        ("Endereço", value("endereco")),
    ]

def _pricing_answer_label_for_pdf(step: dict, answer_data: dict) -> str:
    """Mostra no PDF a opção completa escolhida, não apenas o número digitado."""
    answer = normalize_text(answer_data.get("answer"))

    if not step.get("weighted"):
        return answer or "Não informado"

    weight = answer_data.get("weight")

    try:
        weight_int = int(weight)
    except Exception:
        weight_int = _pricing_weight_from_answer(step.get("id", ""), answer) or 0

    if weight_int:
        for option in step.get("options", []):
            option_text = normalize_text(option)
            if re.match(rf"^\s*{weight_int}\s*[—-]", option_text):
                return option_text

        return f"Peso {weight_int}"

    return answer or "Não informado"

def _pricing_generate_pdf(company_name: str, df: pd.DataFrame, columns: dict) -> bytes:
    """Gera um PDF de diagnóstico com os dados cadastrais e todas as respostas do vendedor."""
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.platypus import (
            KeepTogether,
            Paragraph,
            SimpleDocTemplate,
            Spacer,
            Table,
            TableStyle,
        )
    except Exception as error:
        raise RuntimeError(
            "A biblioteca reportlab não está instalada. Adicione a linha reportlab no requirements.txt, salve e faça o deploy novamente."
        ) from error

    report = _pricing_report_summary(company_name)
    answer_map = report["answer_map"]
    registration_rows = _pricing_company_registration_data(df, columns, company_name)
    proposal_number = f"OPPI-{pd.Timestamp.now(tz='America/Sao_Paulo').strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"
    buffer = io.BytesIO()

    page_width, page_height = A4
    margin = 16 * mm

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=margin,
        leftMargin=margin,
        topMargin=18 * mm,
        bottomMargin=16 * mm,
        title=f"Diagnóstico comercial - {company_name}",
        author="Oppi Comercial",
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="OppiTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=20,
        leading=24,
        textColor=colors.HexColor("#FFFFFF"),
        alignment=TA_LEFT,
        spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        name="OppiSubtitle",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor("#E8DDF4"),
    ))
    styles.add(ParagraphStyle(
        name="OppiSection",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#FFFFFF"),
    ))
    styles.add(ParagraphStyle(
        name="OppiLabel",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#271B35"),
    ))
    styles.add(ParagraphStyle(
        name="OppiValue",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#2B2237"),
    ))
    styles.add(ParagraphStyle(
        name="OppiProposalText",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8.4,
        leading=12,
        textColor=colors.HexColor("#2B2237"),
        spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        name="OppiProposalBold",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=8.8,
        leading=12,
        textColor=colors.HexColor("#271B35"),
        spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        name="OppiSmall",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor("#5D5368"),
    ))
    styles.add(ParagraphStyle(
        name="OppiCenter",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=11,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#FFFFFF"),
    ))

    story = []

    header = Table([
        [Paragraph("OPPI COMERCIAL", styles["OppiSubtitle"])],
        [Paragraph("Diagnóstico de precificação", styles["OppiTitle"])],
        [Paragraph(f"Gerado em: {pd.Timestamp.now(tz='America/Sao_Paulo').strftime('%d/%m/%Y %H:%M')}", styles["OppiSubtitle"])],
        [Paragraph(f"Número da proposta: {html.escape(proposal_number)}", styles["OppiSubtitle"])],
    ], colWidths=[page_width - (2 * margin) - 10])
    header.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#160C2D")),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#FF4BAA")),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
        ("RIGHTPADDING", (0, 0), (-1, -1), 14),
        ("TOPPADDING", (0, 0), (-1, 0), 11),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 11),
    ]))
    story.append(header)
    story.append(Spacer(1, 10))

    registration_header = Table([[Paragraph("DADOS CADASTRAIS DA EMPRESA", styles["OppiSection"]) ]], colWidths=[176 * mm])
    registration_header.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#3B174D")),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#FF4BAA")),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(registration_header)

    registration_data = []
    for label, value in registration_rows:
        registration_data.append([
            Paragraph(html.escape(normalize_text(label)), styles["OppiLabel"]),
            Paragraph(html.escape(normalize_text(value)), styles["OppiValue"]),
        ])

    registration_table = Table(registration_data, colWidths=[48 * mm, 128 * mm], repeatRows=0)
    registration_table.setStyle(TableStyle([
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.HexColor("#FFFFFF"), colors.HexColor("#F4F1F8")]),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#E24AA8")),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D8CBE6")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.extend([registration_table, Spacer(1, 10)])

    story.append(Spacer(1, 10))

    def _proposal_header(title: str):
        section = Table([[Paragraph(html.escape(title), styles["OppiSection"])]], colWidths=[176 * mm])
        section.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#3B174D")),
            ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#FF4BAA")),
            ("LEFTPADDING", (0, 0), (-1, -1), 9),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ]))
        return section

    def _proposal_box(paragraphs: list[str]):
        data = [[Paragraph(paragraph, styles["OppiProposalText"])] for paragraph in paragraphs if normalize_text(paragraph)]
        box = Table(data, colWidths=[176 * mm])
        box.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFFFFF")),
            ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#E24AA8")),
            ("LEFTPADDING", (0, 0), (-1, -1), 9),
            ("RIGHTPADDING", (0, 0), (-1, -1), 9),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        return box

    def _proposal_section(title: str, paragraphs: list[str]):
        story.append(_proposal_header(title))
        story.append(_proposal_box(paragraphs))
        story.append(Spacer(1, 8))

    selected_product = normalize_text(report["product"]) or "Oppi"
    selected_profile = normalize_text(report["profile"]) or "Não informado"
    selected_price = normalize_text(report["confirmed_value"]) or "Não informado"
    selected_range = normalize_text(report["suggested_price"]) or "Não informado"
    selected_term = "1 mês"
    selected_additional_services = normalize_text(report["additional_services"]) or "Contratos digitais e disparos pelo WhatsApp, conforme necessidade."

    solution_focus = {
        "Oppi Vision": "acompanhamento estratégico da operação, gestão da equipe e dashboards de performance.",
        "Oppi Flow": "organização do fluxo comercial, pipeline, propostas e acompanhamento das etapas de atendimento.",
        "Oppi Track": "acompanhamento operacional, execução dos processos internos e controle das etapas da operação.",
    }.get(selected_product, "organização operacional, automação de processos e acompanhamento visual da operação.")

    _proposal_section("1. CONTEXTO IDENTIFICADO", [
        "Após análise inicial da operação da empresa, identificamos oportunidades relacionadas à organização operacional, acompanhamento da equipe e centralização das informações internas.",
        "Atualmente, muitos processos ainda podem depender de controles manuais, dificultando o acompanhamento em tempo real da operação e reduzindo a previsibilidade dos resultados.",
        "Nosso objetivo é estruturar uma operação mais organizada, acompanhável e automatizada, proporcionando maior controle operacional e melhor acompanhamento dos processos internos.",
        f"Com base nas respostas registradas, o perfil identificado foi <b>{html.escape(selected_profile)}</b> e a solução mais adequada inicialmente é <b>{html.escape(selected_product)}</b>, com foco em {html.escape(solution_focus)}",
    ])

    _proposal_section("2. PRINCIPAIS DESAFIOS IDENTIFICADOS", [
        "Durante a análise inicial, foram considerados desafios como:",
        "- processos realizados manualmente;",
        "- informações descentralizadas;",
        "- dificuldade no acompanhamento operacional;",
        "- falta de visibilidade da equipe;",
        "- ausência de fluxo estruturado;",
        "- retrabalho operacional;",
        "- perda de acompanhamento de clientes e processos.",
        f"<b>Resumo registrado pelo vendedor:</b> {html.escape(normalize_text(answer_map.get('resumo_cliente', {}).get('answer')) or 'Não informado')}",
    ])

    _proposal_section("3. SOLUÇÃO PROPOSTA — ECOSSISTEMA OPPI", [
        "A OPPI desenvolve soluções voltadas à organização operacional, automação de processos e acompanhamento estratégico da operação.",
        "Conforme a necessidade da empresa, a solução poderá envolver:",
        "<b>A) OPPI VISION — Gestão & Performance</b><br/>Sistema voltado ao acompanhamento estratégico da operação e gestão da equipe.<br/>Funcionalidades: dashboards estratégicos; indicadores operacionais; acompanhamento da equipe; análise de produtividade; centralização de informações; gestão visual da operação.",
        "<b>B) OPPI FLOW — Pipeline & Propostas</b><br/>Sistema voltado ao gerenciamento comercial e fluxo operacional de atendimento.<br/>Funcionalidades: pipeline operacional; geração de propostas; acompanhamento de etapas; histórico operacional; gestão visual do fluxo; acompanhamento comercial.",
        "<b>C) OPPI TRACK — Operação & Execução</b><br/>Sistema voltado ao acompanhamento operacional e execução dos processos internos.<br/>Funcionalidades: acompanhamento operacional; organização de etapas; gestão de responsáveis; controle de execução; acompanhamento de status; centralização operacional.",
        f"<b>Solução indicada para este diagnóstico:</b> {html.escape(selected_product)}.",
    ])

    _proposal_section("4. COMO FUNCIONA A OPERAÇÃO", [
        "Fluxo operacional simplificado:",
        "<b>Cliente → Atendimento → Processo → Automação → Dashboard → Acompanhamento → Gestão Operacional</b>",
        "A proposta da OPPI não é apenas implementar tecnologia, mas estruturar um fluxo operacional mais inteligente, organizado e acompanhável para a empresa.",
    ])

    _proposal_section("5. BENEFÍCIOS OPERACIONAIS", [
        "Com a implantação da solução OPPI, a empresa terá:",
        "- maior organização operacional;",
        "- centralização das informações;",
        "- redução de retrabalho;",
        "- acompanhamento da equipe em tempo real;",
        "- redução de perda de informações;",
        "- maior previsibilidade operacional;",
        "- acompanhamento estratégico da operação;",
        "- melhoria no fluxo interno de trabalho.",
    ])

    _proposal_section("6. EXEMPLO DE TRANSFORMAÇÃO OPERACIONAL", [
        "<b>Antes da OPPI</b><br/>- processos espalhados;<br/>- equipe sem acompanhamento visual;<br/>- controles manuais;<br/>- dificuldade na gestão operacional;<br/>- falta de visibilidade dos processos.",
        "<b>Depois da OPPI</b><br/>- operação centralizada;<br/>- acompanhamento em tempo real;<br/>- dashboards operacionais;<br/>- automações integradas;<br/>- maior controle e previsibilidade.",
    ])

    _proposal_section("7. IMPLANTAÇÃO", [
        "A implantação contempla:",
        "- análise operacional inicial;",
        "- estruturação do fluxo;",
        "- configuração da solução;",
        "- parametrização das etapas;",
        "- automações operacionais;",
        "- treinamento inicial da equipe;",
        "- acompanhamento de implantação;",
        "- suporte operacional inicial.",
    ])

    _proposal_section("8. PRAZO ESTIMADO", [
        f"Prazo ideal estimado: <b>{html.escape(selected_term)}</b>.",
        "O prazo poderá variar conforme: complexidade da operação; quantidade de usuários; quantidade de setores envolvidos; disponibilidade das informações; e nível de personalização definido no projeto.",
    ])

    _proposal_section("9. INVESTIMENTO", [
        f"Implantação sugerida pelo diagnóstico: <b>{html.escape(selected_range)}</b>.",
        "Forma de pagamento: a definir em proposta comercial ou contrato.",
        f"Serviços adicionais sugeridos: <b>{html.escape(selected_additional_services)}</b>.",
        "Os serviços adicionais podem envolver contratos digitais, propostas, documentos operacionais e pacotes de disparos pelo WhatsApp, conforme o cenário identificado.",
    ])

    _proposal_section("10. SUPORTE E ACOMPANHAMENTO", [
        "O suporte contempla:",
        "- acompanhamento operacional inicial;",
        "- suporte remoto;",
        "- ajustes simples;",
        "- orientações de utilização;",
        "- acompanhamento estratégico inicial.",
        "Horário de atendimento: Segunda a sexta-feira — 08h às 18h.",
    ])

    _proposal_section("11. DIFERENCIAL OPPI", [
        "A OPPI não atua apenas como fornecedora de tecnologia.",
        "Nosso foco é estruturar operações mais organizadas, previsíveis e acompanháveis, utilizando automação, gestão visual e inteligência operacional para reduzir o caos operacional e melhorar o controle interno da empresa.",
    ])

    _proposal_section("12. CONSIDERAÇÕES FINAIS", [
        "Agradecemos pela oportunidade de apresentar nossa proposta comercial.",
        "Estamos à disposição para alinhamentos, demonstrações e esclarecimentos adicionais.",
    ])

    footer = Table([[Paragraph(
        "Documento gerado automaticamente pelo Dashboard Oppi Comercial.",
        styles["OppiSmall"],
    )]], colWidths=[176 * mm])
    footer.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F5F2FA")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#D8CBE6")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(footer)

    doc.build(story)
    return buffer.getvalue()

def _diagnostic_ensure_thread(company_name: str) -> list[dict]:
    threads = _diagnostic_get_threads()
    progress = _diagnostic_get_progress()
    answers = _diagnostic_get_answers()

    if company_name not in threads:
        threads[company_name] = [
            {
                "role": "assistant",
                "content": OPPI_PRICING_INTRO,
                "time": _diagnostic_now(),
            },
            {
                "role": "assistant",
                "content": _pricing_question_text(OPPI_PRICING_STEPS[0]),
                "time": _diagnostic_now(),
            },
        ]
        progress[company_name] = 0
        answers[company_name] = {}

    return threads[company_name]


def _diagnostic_add_answer(company_name: str, answer: str) -> None:
    messages = _diagnostic_ensure_thread(company_name)
    progress = _diagnostic_get_progress()
    answers = _diagnostic_get_answers()
    clean_answer = normalize_text(answer)

    if not clean_answer:
        return

    current_index = int(progress.get(company_name, 0))

    if current_index >= len(OPPI_PRICING_STEPS):
        messages.append(
            {
                "role": "user",
                "content": clean_answer,
                "time": _diagnostic_now(),
            }
        )
        answers.setdefault(company_name, {})["valor_proposta"] = {
            "answer": clean_answer,
            "weight": None,
        }
        messages.append(
            {
                "role": "assistant",
                "content": (
                    "Perfeito. O valor desejado para a proposta foi confirmado e registrado nesta conversa.\n\n"
                    "Agora clique no botão Gerar PDF do diagnóstico para baixar o documento com os dados do cadastro, "
                    "as respostas da precificação, o resumo da reunião, a solução indicada e o valor confirmado."
                ),
                "time": _diagnostic_now(),
            }
        )
        return

    step = OPPI_PRICING_STEPS[current_index]
    weight = None

    if step["weighted"]:
        weight = _pricing_weight_from_answer(step["id"], clean_answer)

        if weight is None:
            messages.append(
                {
                    "role": "assistant",
                    "content": (
                        "Não consegui identificar o peso dessa resposta. Responda com uma opção de 1 a 4 para continuarmos.\n\n"
                        f"{step['example']}"
                    ),
                    "time": _diagnostic_now(),
                }
            )
            return

    messages.append(
        {
            "role": "user",
            "content": clean_answer,
            "time": _diagnostic_now(),
        }
    )

    answers.setdefault(company_name, {})[step["id"]] = {
        "answer": clean_answer,
        "weight": weight,
    }

    next_index = current_index + 1
    progress[company_name] = next_index

    if next_index < len(OPPI_PRICING_STEPS):
        messages.append(
            {
                "role": "assistant",
                "content": _pricing_question_text(OPPI_PRICING_STEPS[next_index]),
                "time": _diagnostic_now(),
            }
        )
        return

    messages.append(
        {
            "role": "assistant",
            "content": _pricing_result_message(company_name),
            "time": _diagnostic_now(),
        }
    )


def _diagnostic_reset(company_name: str) -> None:
    threads = _diagnostic_get_threads()
    progress = _diagnostic_get_progress()
    answers = _diagnostic_get_answers()
    threads.pop(company_name, None)
    progress.pop(company_name, None)
    answers.pop(company_name, None)
    _diagnostic_ensure_thread(company_name)


def _diagnostic_render_messages(messages: list[dict]) -> str:
    rows = ['<div class="oppi-chat-messages">', '<div class="oppi-chat-day"><span>Hoje</span></div>']

    for message in messages:
        role = "user" if message.get("role") == "user" else "assistant"
        safe_content = html.escape(normalize_text(message.get("content"))).replace("\n", "<br>")
        safe_time = html.escape(normalize_text(message.get("time")))
        check = " ✓✓" if role == "user" else ""

        rows.append(
            f'<div class="oppi-chat-message-row {role}">'
            f'<div class="oppi-chat-bubble">{safe_content}'
            f'<span class="oppi-chat-bubble-time">{safe_time}{check}</span>'
            f'</div></div>'
        )

    rows.append("</div>")
    return "".join(rows)



def apply_pesos_chat_visibility_fix() -> None:
    """Garante que o painel da consultoria em Pesos e Medidas fique visível."""
    render_html(
        """
        <style>
            /* Força o painel da direita a ocupar a tela e aparecer claro. */
            .st-key-diagnostic_chat_panel {
                display: flex !important;
                flex-direction: column !important;
                visibility: visible !important;
                opacity: 1 !important;
                background: #E4E7EE !important;
                background-color: #E4E7EE !important;
                min-height: 100dvh !important;
                height: 100dvh !important;
                max-height: 100dvh !important;
                overflow: hidden !important;
                position: relative !important;
                z-index: 1 !important;
            }

            .st-key-diagnostic_chat_panel > div[data-testid="stVerticalBlock"],
            .st-key-diagnostic_chat_panel div[data-testid="stVerticalBlock"] {
                display: flex !important;
                flex-direction: column !important;
                visibility: visible !important;
                opacity: 1 !important;
                height: 100% !important;
                max-height: 100% !important;
                min-height: 0 !important;
                background: transparent !important;
                overflow: hidden !important;
            }

            .oppi-chat-window-header {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                min-height: 82px !important;
                background: rgba(255,255,255,0.96) !important;
                color: #111827 !important;
                border-bottom: 1px solid rgba(80,69,105,0.16) !important;
                position: relative !important;
                z-index: 3 !important;
            }

            .oppi-chat-window-name {
                color: #111827 !important;
                -webkit-text-fill-color: #111827 !important;
            }

            .oppi-chat-window-status {
                color: #16A34A !important;
                -webkit-text-fill-color: #16A34A !important;
            }

            .oppi-chat-messages {
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
                flex: 1 1 auto !important;
                min-height: 260px !important;
                height: auto !important;
                max-height: none !important;
                overflow-y: auto !important;
                background:
                    radial-gradient(circle at 78% 8%, rgba(190, 72, 255, 0.17), transparent 24%),
                    linear-gradient(180deg, #ECEEF4 0%, #DCE0E9 100%) !important;
                padding: 22px 26px !important;
                position: relative !important;
                z-index: 2 !important;
            }

            .oppi-chat-bubble {
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
            }

            .oppi-chat-progress-wrap,
            .st-key-diagnostic_chat_toolbar,
            .st-key-diagnostic_chat_form {
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
                position: relative !important;
                z-index: 4 !important;
                background: #F7F8FC !important;
            }
        </style>
        """
    )



# =========================================================
# CORREÇÃO FINAL: RESTAURAR CHAT DO PAINEL DIREITO
# =========================================================
def apply_restore_right_chat_only_css() -> None:
    """Restaura somente o painel direito da página Pesos e Medidas, sem mexer na lista de empresas."""
    render_html(
        """
        <style>
            /* Painel direito: precisa existir, ficar acima do fundo roxo/preto e ocupar a altura toda. */
            .st-key-diagnostic_chat_panel {
                display: flex !important;
                flex-direction: column !important;
                visibility: visible !important;
                opacity: 1 !important;
                position: relative !important;
                z-index: 999 !important;
                width: 100% !important;
                min-width: 100% !important;
                min-height: 100dvh !important;
                height: 100dvh !important;
                max-height: 100dvh !important;
                margin: 0 !important;
                padding: 0 !important;
                overflow: hidden !important;
                background: #E4E7EE !important;
                background-color: #E4E7EE !important;
                border-left: 1px solid rgba(80,69,105,0.16) !important;
                border-radius: 0 !important;
                box-shadow: none !important;
            }

            .st-key-diagnostic_chat_panel > div[data-testid="stVerticalBlock"],
            .st-key-diagnostic_chat_panel div[data-testid="stVerticalBlock"] {
                display: flex !important;
                flex-direction: column !important;
                visibility: visible !important;
                opacity: 1 !important;
                width: 100% !important;
                height: 100% !important;
                min-height: 100% !important;
                max-height: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
                gap: 0 !important;
                overflow: hidden !important;
                background: #E4E7EE !important;
            }

            .st-key-diagnostic_chat_panel .oppi-chat-window-header {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                flex: 0 0 76px !important;
                min-height: 76px !important;
                background: rgba(255,255,255,0.96) !important;
                color: #111827 !important;
                border-bottom: 1px solid rgba(80,69,105,0.14) !important;
            }

            .st-key-diagnostic_chat_panel .oppi-chat-window-name {
                color: #111827 !important;
                -webkit-text-fill-color: #111827 !important;
            }

            .st-key-diagnostic_chat_panel .oppi-chat-window-status {
                color: #16A34A !important;
                -webkit-text-fill-color: #16A34A !important;
            }

            .st-key-diagnostic_chat_panel .oppi-chat-messages {
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
                flex: 1 1 auto !important;
                min-height: 260px !important;
                height: auto !important;
                max-height: none !important;
                overflow-y: auto !important;
                padding: 20px 22px 16px 22px !important;
                background: linear-gradient(180deg, #E8EAF0 0%, #DCE0E9 100%) !important;
            }

            .st-key-diagnostic_chat_panel .oppi-chat-progress-wrap,
            .st-key-diagnostic_chat_toolbar,
            .st-key-diagnostic_chat_form {
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
                flex: 0 0 auto !important;
                background: #F8F8FB !important;
                position: relative !important;
                z-index: 1000 !important;
            }

            .st-key-diagnostic_chat_form {
                padding: 10px 14px 12px 14px !important;
                border-top: 1px solid rgba(80,69,105,0.12) !important;
            }

            .st-key-diagnostic_chat_form input {
                color: #211A30 !important;
                -webkit-text-fill-color: #211A30 !important;
            }
        </style>
        """
    )


def apply_chat_controls_and_pdf_visibility_fix() -> None:
    """Mantém o campo de resposta, botões e PDF visíveis no painel direito do chat."""
    render_html(
        """
        <style>
            /* Correção aplicada SOMENTE no lado direito do Pesos e Medidas. */
            .st-key-diagnostic_chat_panel {
                display: flex !important;
                flex-direction: column !important;
                height: 100dvh !important;
                min-height: 100dvh !important;
                max-height: 100dvh !important;
                overflow: hidden !important;
                background: #E4E7EE !important;
            }

            .st-key-diagnostic_chat_panel > div[data-testid="stVerticalBlock"] {
                display: flex !important;
                flex-direction: column !important;
                height: 100% !important;
                min-height: 100% !important;
                max-height: 100% !important;
                overflow: hidden !important;
                gap: 0 !important;
                background: #E4E7EE !important;
            }

            /* O histórico do chat rola internamente e deixa espaço para barra, botões e campo de digitação. */
            .st-key-diagnostic_chat_panel .oppi-chat-messages {
                flex: 1 1 auto !important;
                height: auto !important;
                min-height: 230px !important;
                max-height: calc(100dvh - 280px) !important;
                overflow-y: auto !important;
                overflow-x: hidden !important;
                padding-bottom: 16px !important;
            }

            .st-key-diagnostic_chat_panel .oppi-chat-progress-wrap {
                flex: 0 0 auto !important;
                min-height: 42px !important;
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
                background: #F7F7FB !important;
                position: relative !important;
                z-index: 20 !important;
            }

            .st-key-diagnostic_chat_toolbar {
                flex: 0 0 auto !important;
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
                background: #F7F0FA !important;
                border-top: 1px solid rgba(80,69,105,0.10) !important;
                border-bottom: 1px solid rgba(80,69,105,0.10) !important;
                padding: 8px 14px !important;
                position: relative !important;
                z-index: 30 !important;
            }

            .st-key-diagnostic_chat_toolbar div[data-testid="stHorizontalBlock"] {
                gap: 8px !important;
            }

            .st-key-diagnostic_chat_toolbar button,
            .st-key-diagnostic_chat_toolbar [data-testid="stDownloadButton"] button {
                min-height: 42px !important;
                border-radius: 999px !important;
                border: 0 !important;
                color: #FFFFFF !important;
                font-weight: 850 !important;
                background: linear-gradient(90deg, #FF4BAA 0%, #A91CFF 100%) !important;
                box-shadow: 0 8px 18px rgba(169,28,255,0.18) !important;
            }

            .st-key-diagnostic_chat_form {
                flex: 0 0 auto !important;
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
                padding: 10px 14px 0 14px !important;
                background: #FFFFFF !important;
                border-top: 1px solid rgba(80,69,105,0.14) !important;
                position: relative !important;
                z-index: 40 !important;
            }

            .st-key-diagnostic_chat_form [data-testid="stForm"] {
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
                max-width: none !important;
                width: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
                border: 0 !important;
                background: transparent !important;
                box-shadow: none !important;
            }

            .st-key-diagnostic_chat_form [data-testid="stForm"]::before,
            .st-key-diagnostic_chat_form [data-testid="stForm"]::after {
                display: none !important;
            }

            .st-key-diagnostic_chat_form div[data-testid="stHorizontalBlock"] {
                gap: 8px !important;
                align-items: center !important;
            }

            .st-key-diagnostic_chat_form div[data-baseweb="input"] > div {
                min-height: 48px !important;
                height: 48px !important;
                border-radius: 999px !important;
                background: #FFFFFF !important;
                border: 1px solid rgba(80,69,105,0.18) !important;
            }

            .st-key-diagnostic_chat_form input {
                min-height: 48px !important;
                color: #211A30 !important;
                -webkit-text-fill-color: #211A30 !important;
            }

            .st-key-diagnostic_chat_form input::placeholder {
                color: rgba(33,26,48,0.46) !important;
                -webkit-text-fill-color: rgba(33,26,48,0.46) !important;
            }

            .st-key-diagnostic_chat_form button[kind="secondaryFormSubmit"] {
                width: 48px !important;
                min-width: 48px !important;
                height: 48px !important;
                min-height: 48px !important;
                border-radius: 50% !important;
                border: 0 !important;
                color: #FFFFFF !important;
                background: linear-gradient(135deg, #FF4BAA 0%, #A91CFF 100%) !important;
                box-shadow: 0 8px 18px rgba(169,28,255,0.22) !important;
            }
        </style>
        """
    )




def apply_pesos_bottom_alignment_fix() -> None:
    """Remove a faixa preta inferior e alinha o painel do chat até a base da tela."""
    render_html(
        """
        <style>
            html, body,
            .stApp,
            [data-testid="stAppViewContainer"],
            [data-testid="stMain"],
            [data-testid="stMainBlockContainer"] {
                background: #E4E7EE !important;
            }

            [data-testid="stMainBlockContainer"] {
                padding-top: 0 !important;
                padding-bottom: 0 !important;
                margin-top: 0 !important;
                margin-bottom: 0 !important;
                min-height: 100dvh !important;
                height: 100dvh !important;
                max-height: 100dvh !important;
                overflow: hidden !important;
            }

            [data-testid="stMainBlockContainer"] > div[data-testid="stVerticalBlock"] {
                min-height: 100dvh !important;
                height: 100dvh !important;
                max-height: 100dvh !important;
                gap: 0 !important;
                margin: 0 !important;
                padding: 0 !important;
                overflow: hidden !important;
            }

            .st-key-diagnostic_contacts_panel,
            .st-key-diagnostic_chat_panel {
                min-height: 100dvh !important;
                height: 100dvh !important;
                max-height: 100dvh !important;
                margin-bottom: 0 !important;
                padding-bottom: 0 !important;
                overflow: hidden !important;
            }

            .st-key-diagnostic_chat_form {
                margin-bottom: 0 !important;
                padding-bottom: 0 !important;
                border-bottom: 0 !important;
            }

            .st-key-diagnostic_chat_form [data-testid="stForm"],
            .st-key-diagnostic_chat_form div[data-testid="stElementContainer"] {
                margin-bottom: 0 !important;
                padding-bottom: 0 !important;
            }
        </style>
        """
    )

def render_scoring_page(df: pd.DataFrame, columns: dict) -> None:
    apply_chat_css()
    apply_global_sidebar_toggle_css()
    apply_pesos_chat_visibility_fix()
    apply_restore_right_chat_only_css()
    apply_chat_controls_and_pdf_visibility_fix()
    apply_pesos_bottom_alignment_fix()

    companies = sorted(
        {
            normalize_text(company)
            for company in df["_empresa"].tolist()
            if normalize_text(company)
        }
    )

    if not companies:
        st.info("Nenhuma empresa cadastrada para iniciar uma precificação.")
        return

    selected_company_key = f"oppi_pricing_selected_company_{PRICING_SCRIPT_VERSION}"

    if selected_company_key not in st.session_state:
        st.session_state[selected_company_key] = companies[0]

    if st.session_state[selected_company_key] not in companies:
        st.session_state[selected_company_key] = companies[0]

    left_column, right_column = st.columns([0.92, 1.78], gap=None)

    with left_column:
        with st.container(key="diagnostic_contacts_panel"):
            render_html(
                """
                <div class="oppi-chat-contact-header">
                    <div>
                        <div class="oppi-chat-contact-title">Empresas</div>
                        <div class="oppi-chat-contact-subtitle">Assistente de precificação para o vendedor</div>
                    </div>
                    <div class="oppi-chat-avatar">OP</div>
                </div>
                """
            )

            search_term = st.text_input(
                "Buscar empresas",
                placeholder="🔍  Buscar empresa...",
                label_visibility="collapsed",
                key=f"oppi_pricing_search_{PRICING_SCRIPT_VERSION}",
            )

            normalized_search = normalize_search_text(search_term)
            visible_companies = [
                company
                for company in companies
                if not normalized_search or normalized_search in normalize_search_text(company)
            ]

            with st.container(key="diagnostic_contacts_list"):
                for company in visible_companies:
                    messages = _diagnostic_ensure_thread(company)
                    last_message = normalize_text(messages[-1].get("content")) if messages else ""
                    snippet = last_message[:48] + ("..." if len(last_message) > 48 else "")
                    initials = _diagnostic_initials(company)
                    label = f"{initials}   {company}\n{snippet}"
                    selected = company == st.session_state[selected_company_key]

                    if st.button(
                        label,
                        key=f"pricing_contact_{_pricing_safe_key_fragment(company)}",
                        use_container_width=True,
                        type="primary" if selected else "secondary",
                    ):
                        st.session_state[selected_company_key] = company
                        st.rerun()

    with right_column:
        with st.container(key="diagnostic_chat_panel"):
            selected_company = st.session_state[selected_company_key]
            messages = _diagnostic_ensure_thread(selected_company)
            progress = int(_diagnostic_get_progress().get(selected_company, 0))
            answered = min(progress, len(OPPI_PRICING_STEPS))
            progress_percent = round((answered / len(OPPI_PRICING_STEPS)) * 100)
            initials = _diagnostic_initials(selected_company)

            render_html(
                f"""
                <div class="oppi-chat-window-header">
                    <div class="oppi-chat-window-person">
                        <div class="oppi-chat-avatar">{html.escape(initials)}</div>
                        <div>
                            <div class="oppi-chat-window-name">{html.escape(selected_company)}</div>
                            <div class="oppi-chat-window-status">Precificação ativa para o vendedor</div>
                        </div>
                    </div>
                </div>
                {_diagnostic_render_messages(messages)}
                <div class="oppi-chat-progress-wrap">
                    <div class="oppi-chat-progress-label">Tabela de elaboração de preço: {answered} de {len(OPPI_PRICING_STEPS)} respostas registradas</div>
                    <div class="oppi-chat-progress-bar"><div class="oppi-chat-progress-fill" style="width:{progress_percent}%;"></div></div>
                </div>
                """
            )

            with st.container(key="diagnostic_chat_toolbar"):
                confirmed_value = _pricing_get_confirmed_value(selected_company)

                if confirmed_value:
                    toolbar_left, toolbar_middle, toolbar_right = st.columns([1.0, 1.0, 1.0], gap="small")
                else:
                    toolbar_left, toolbar_right = st.columns([1.0, 1.0], gap="small")
                    toolbar_middle = None

                with toolbar_left:
                    if st.button(
                        "↻ Reiniciar precificação",
                        use_container_width=True,
                        key=f"reset_pricing_{_pricing_safe_key_fragment(selected_company)}",
                    ):
                        _diagnostic_reset(selected_company)
                        st.rerun()

                if toolbar_middle is not None:
                    with toolbar_middle:
                        try:
                            pdf_bytes = _pricing_generate_pdf(selected_company, df, columns)
                            st.download_button(
                                "📄 Gerar PDF do diagnóstico",
                                data=pdf_bytes,
                                file_name=f"diagnostico_oppi_{_pricing_pdf_safe_filename(selected_company)}.pdf",
                                mime="application/pdf",
                                use_container_width=True,
                                key=f"download_pricing_pdf_{_pricing_safe_key_fragment(selected_company)}",
                            )
                        except Exception as error:
                            st.error(f"Não consegui gerar o PDF: {error}")

                with toolbar_right:
                    st.button(
                        "✓ Perguntas fixas da Oppi",
                        use_container_width=True,
                        disabled=True,
                        key=f"fixed_pricing_script_{_pricing_safe_key_fragment(selected_company)}",
                    )

            with st.container(key="diagnostic_chat_form"):
                with st.form(
                    f"pricing_form_{_pricing_safe_key_fragment(selected_company)}",
                    clear_on_submit=True,
                ):
                    input_column, send_column = st.columns([8.4, 1.0], gap="small")

                    with input_column:
                        answer = st.text_input(
                            "Mensagem",
                            placeholder="Digite a resposta do vendedor...",
                            label_visibility="collapsed",
                            key=f"pricing_answer_{_pricing_safe_key_fragment(selected_company)}",
                        )

                    with send_column:
                        submitted = st.form_submit_button("➤", use_container_width=True)

                if submitted and normalize_text(answer):
                    _diagnostic_add_answer(selected_company, answer)
                    st.rerun()



# =========================================================
# AJUSTE DEFINITIVO SOLICITADO: SETA CINZA + LISTA COM SCROLL
# =========================================================
def apply_requested_sidebar_and_chat_fix_css() -> None:
    """Ajuste final: remove faixas vazias, centraliza a seta cinza e mostra 4 conversas com scroll."""
    render_html(
        """
        <style>
            /* Remove definitivamente qualquer faixa vazia criada acima do chat. */
            .diagnostic-page-top-spacer {
                display: none !important;
                width: 0 !important;
                height: 0 !important;
                min-height: 0 !important;
                max-height: 0 !important;
                margin: 0 !important;
                padding: 0 !important;
            }

            /* O chat encosta no topo e na base da área útil, sem faixas pretas. */
            .st-key-diagnostic_contacts_panel,
            .st-key-diagnostic_chat_panel {
                margin: 0 !important;
                padding: 0 !important;
                min-height: 100dvh !important;
                height: 100dvh !important;
                max-height: 100dvh !important;
                border-radius: 0 !important;
                box-shadow: none !important;
                overflow: hidden !important;
            }

            [data-testid="stMainBlockContainer"],
            [data-testid="stMainBlockContainer"] > div[data-testid="stVerticalBlock"] {
                padding-top: 0 !important;
                padding-bottom: 0 !important;
                margin-top: 0 !important;
                margin-bottom: 0 !important;
                gap: 0 !important;
            }

            /* Quando o menu estiver recolhido, a seta aparece no meio da lateral em cinza. */
            [data-testid="collapsedControl"],
            [data-testid="stSidebarCollapsedControl"],
            button[data-testid="collapsedControl"],
            button[data-testid="stSidebarCollapsedControl"] {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                position: fixed !important;
                top: 50vh !important;
                left: 12px !important;
                width: 42px !important;
                min-width: 42px !important;
                height: 42px !important;
                min-height: 42px !important;
                padding: 0 !important;
                align-items: center !important;
                justify-content: center !important;
                border-radius: 12px !important;
                border: 1px solid rgba(75, 85, 99, 0.34) !important;
                background: #D1D5DB !important;
                background-color: #D1D5DB !important;
                color: #4B5563 !important;
                box-shadow: 0 8px 20px rgba(0, 0, 0, 0.20) !important;
                transform: translateY(-50%) !important;
                pointer-events: auto !important;
                z-index: 2147483647 !important;
                overflow: visible !important;
            }

            [data-testid="collapsedControl"] button,
            [data-testid="stSidebarCollapsedControl"] button {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                width: 100% !important;
                height: 100% !important;
                padding: 0 !important;
                align-items: center !important;
                justify-content: center !important;
                border: none !important;
                border-radius: 12px !important;
                background: transparent !important;
                box-shadow: none !important;
                pointer-events: auto !important;
            }

            [data-testid="collapsedControl"] svg,
            [data-testid="stSidebarCollapsedControl"] svg {
                display: none !important;
                visibility: hidden !important;
                opacity: 0 !important;
            }

            [data-testid="collapsedControl"]::after,
            [data-testid="stSidebarCollapsedControl"]::after,
            button[data-testid="collapsedControl"]::after,
            button[data-testid="stSidebarCollapsedControl"]::after {
                content: "›" !important;
                display: block !important;
                position: absolute !important;
                top: 50% !important;
                left: 50% !important;
                transform: translate(-50%, -54%) !important;
                color: #4B5563 !important;
                font-family: Arial, sans-serif !important;
                font-size: 33px !important;
                font-weight: 900 !important;
                line-height: 1 !important;
                pointer-events: none !important;
            }

            [data-testid="collapsedControl"]:hover,
            [data-testid="stSidebarCollapsedControl"]:hover,
            button[data-testid="collapsedControl"]:hover,
            button[data-testid="stSidebarCollapsedControl"]:hover {
                background: #E5E7EB !important;
                background-color: #E5E7EB !important;
                transform: translateY(-50%) scale(1.06) !important;
            }

            /* Mostra exatamente 4 conversas completas inicialmente e permite rolar todas as demais. */
            .st-key-diagnostic_contacts_list {
                display: block !important;
                flex: 0 0 344px !important;
                width: 100% !important;
                height: 344px !important;
                min-height: 344px !important;
                max-height: 344px !important;
                margin: 0 !important;
                padding: 0 4px 0 0 !important;
                overflow-y: auto !important;
                overflow-x: hidden !important;
                scrollbar-width: thin !important;
                scrollbar-color: rgba(169, 28, 255, 0.76) rgba(255, 255, 255, 0.58) !important;
            }

            .st-key-diagnostic_contacts_list > div[data-testid="stVerticalBlock"],
            .st-key-diagnostic_contacts_list div[data-testid="stVerticalBlock"] {
                display: block !important;
                width: 100% !important;
                height: auto !important;
                min-height: 0 !important;
                max-height: none !important;
                margin: 0 !important;
                padding: 0 !important;
                overflow: visible !important;
                gap: 0 !important;
            }

            .st-key-diagnostic_contacts_list div[data-testid="stElementContainer"],
            .st-key-diagnostic_contacts_list .stButton {
                margin: 0 !important;
                padding: 0 !important;
            }

            .st-key-diagnostic_contacts_list .stButton > button {
                width: calc(100% - 16px) !important;
                min-height: 76px !important;
                height: 76px !important;
                max-height: 76px !important;
                margin: 5px 8px !important;
                padding: 8px 12px !important;
                overflow: hidden !important;
                transition: transform 0.18s ease, box-shadow 0.18s ease, filter 0.18s ease !important;
            }

            .st-key-diagnostic_contacts_list .stButton > button:hover {
                transform: scale(1.018) !important;
                filter: brightness(1.03) !important;
                box-shadow: 0 10px 22px rgba(169, 28, 255, 0.16) !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar {
                width: 8px !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar-track {
                background: rgba(255, 255, 255, 0.58) !important;
                border-radius: 999px !important;
            }

            .st-key-diagnostic_contacts_list::-webkit-scrollbar-thumb {
                border-radius: 999px !important;
                background: linear-gradient(180deg, #FF4BAA 0%, #A91CFF 100%) !important;
            }
        </style>
        """
    )

# =========================================================
# AJUSTE FINAL DO CHAT: ESPAÇO SUPERIOR PARA A SETA DO MENU
# =========================================================
def apply_chat_sidebar_toggle_slot_css() -> None:
    """Reserva espaço no topo da coluna de empresas e posiciona a seta do menu recolhido nesse espaço."""
    render_html(
        """
        <style>
            /*
               Na tela Pesos e Medidas, a coluna de empresas tinha sobra na parte inferior,
               mas o topo ficava ocupado pelo título. Reservamos um pequeno espaço no topo
               para a seta que reabre o menu quando o sidebar estiver recolhido.
            */
            .st-key-diagnostic_contacts_panel > div[data-testid="stVerticalBlock"] {
                box-sizing: border-box !important;
                padding-top: 46px !important;
            }

            /* A seta fica dentro da área clara liberada no topo esquerdo do chat. */
            [data-testid="collapsedControl"],
            [data-testid="stSidebarCollapsedControl"],
            button[data-testid="collapsedControl"],
            button[data-testid="stSidebarCollapsedControl"] {
                top: 12px !important;
                left: 14px !important;
                transform: none !important;
            }

            [data-testid="collapsedControl"]:hover,
            [data-testid="stSidebarCollapsedControl"]:hover,
            button[data-testid="collapsedControl"]:hover,
            button[data-testid="stSidebarCollapsedControl"]:hover {
                transform: scale(1.06) !important;
            }

            /* Mantém exatamente quatro conversas visíveis e rolagem interna para as demais. */
            .st-key-diagnostic_contacts_list {
                height: 344px !important;
                min-height: 344px !important;
                max-height: 344px !important;
                overflow-y: auto !important;
                overflow-x: hidden !important;
            }
        </style>
        """
    )


# =========================================================
# AJUSTE FINAL REAL: ESPAÇO SUPERIOR NA COLUNA DE EMPRESAS
# =========================================================
def apply_chat_contacts_top_slot_real_css() -> None:
    """
    Reserva uma faixa clara acima de "Empresas" para a seta do menu recolhido.
    A lista continua exibindo quatro conversas completas e aproveita a sobra inferior.
    """
    render_html(
        """
        <style>
            /* Move somente o conteúdo da coluna esquerda para baixo. */
            .st-key-diagnostic_contacts_panel {
                box-sizing: border-box !important;
                padding-top: 54px !important;
                position: relative !important;
            }

            /* Neutraliza o padding antigo aplicado no bloco interno para não duplicar o espaço. */
            .st-key-diagnostic_contacts_panel > div[data-testid="stVerticalBlock"] {
                box-sizing: border-box !important;
                padding-top: 0 !important;
            }

            /* Mantém quatro conversas completas e rolagem interna para todas as demais. */
            .st-key-diagnostic_contacts_list {
                height: 344px !important;
                min-height: 344px !important;
                max-height: 344px !important;
                overflow-y: auto !important;
                overflow-x: hidden !important;
            }

            /* Seta nativa visível dentro da faixa livre quando o menu estiver fechado. */
            [data-testid="collapsedControl"],
            [data-testid="stSidebarCollapsedControl"],
            button[data-testid="collapsedControl"],
            button[data-testid="stSidebarCollapsedControl"] {
                position: fixed !important;
                top: 14px !important;
                left: 14px !important;
                width: 40px !important;
                min-width: 40px !important;
                height: 40px !important;
                min-height: 40px !important;
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                align-items: center !important;
                justify-content: center !important;
                padding: 0 !important;
                border-radius: 12px !important;
                border: 1px solid rgba(75,85,99,0.30) !important;
                background: #D1D5DB !important;
                background-color: #D1D5DB !important;
                color: #4B5563 !important;
                box-shadow: 0 8px 18px rgba(0,0,0,0.16) !important;
                transform: none !important;
                pointer-events: auto !important;
                z-index: 2147483647 !important;
            }

            [data-testid="collapsedControl"] button,
            [data-testid="stSidebarCollapsedControl"] button {
                width: 100% !important;
                height: 100% !important;
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
                padding: 0 !important;
                border: none !important;
                border-radius: 12px !important;
                background: transparent !important;
                box-shadow: none !important;
            }

            /* Usa a seta nativa do Streamlit em cinza; remove o símbolo artificial antigo. */
            [data-testid="collapsedControl"]::after,
            [data-testid="stSidebarCollapsedControl"]::after,
            button[data-testid="collapsedControl"]::after,
            button[data-testid="stSidebarCollapsedControl"]::after {
                content: none !important;
                display: none !important;
            }

            [data-testid="collapsedControl"] svg,
            [data-testid="stSidebarCollapsedControl"] svg,
            button[data-testid="collapsedControl"] svg,
            button[data-testid="stSidebarCollapsedControl"] svg {
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
                width: 22px !important;
                height: 22px !important;
                color: #4B5563 !important;
                fill: #4B5563 !important;
                stroke: #4B5563 !important;
            }

            [data-testid="collapsedControl"]:hover,
            [data-testid="stSidebarCollapsedControl"]:hover,
            button[data-testid="collapsedControl"]:hover,
            button[data-testid="stSidebarCollapsedControl"]:hover {
                background: #E5E7EB !important;
                background-color: #E5E7EB !important;
                transform: scale(1.06) !important;
            }
        </style>
        """
    )


def install_chat_contacts_top_slot_runtime_fix() -> None:
    """Reaplica o posicionamento após os reruns do Streamlit."""
    components.html(
        """
        <script>
            (function () {
                function getHostDocument() {
                    try {
                        if (window.frameElement && window.frameElement.ownerDocument) {
                            return window.frameElement.ownerDocument;
                        }
                    } catch (error) {}

                    try {
                        return window.parent.document;
                    } catch (error) {
                        return document;
                    }
                }

                const hostDocument = getHostDocument();
                const hostWindow = window.parent || window;

                function forceStyle(element, property, value) {
                    if (element && element.style) {
                        element.style.setProperty(property, value, 'important');
                    }
                }

                function applyFix() {
                    const panel = hostDocument.querySelector('.st-key-diagnostic_contacts_panel');

                    if (panel) {
                        forceStyle(panel, 'box-sizing', 'border-box');
                        forceStyle(panel, 'padding-top', '54px');
                        forceStyle(panel, 'position', 'relative');
                    }

                    const directBlock = panel
                        ? panel.querySelector(':scope > div[data-testid="stVerticalBlock"]')
                        : null;

                    if (directBlock) {
                        forceStyle(directBlock, 'padding-top', '0px');
                    }

                    const selectors = [
                        '[data-testid="collapsedControl"]',
                        '[data-testid="stSidebarCollapsedControl"]',
                        'button[data-testid="collapsedControl"]',
                        'button[data-testid="stSidebarCollapsedControl"]'
                    ];

                    selectors.forEach(function (selector) {
                        hostDocument.querySelectorAll(selector).forEach(function (control) {
                            forceStyle(control, 'top', '14px');
                            forceStyle(control, 'left', '14px');
                            forceStyle(control, 'background', '#D1D5DB');
                            forceStyle(control, 'background-color', '#D1D5DB');
                            forceStyle(control, 'color', '#4B5563');
                            forceStyle(control, 'visibility', 'visible');
                            forceStyle(control, 'opacity', '1');
                            forceStyle(control, 'z-index', '2147483647');

                            control.querySelectorAll('svg').forEach(function (svg) {
                                forceStyle(svg, 'display', 'block');
                                forceStyle(svg, 'visibility', 'visible');
                                forceStyle(svg, 'opacity', '1');
                                forceStyle(svg, 'color', '#4B5563');
                                forceStyle(svg, 'fill', '#4B5563');
                                forceStyle(svg, 'stroke', '#4B5563');
                            });
                        });
                    });
                }

                applyFix();
                hostWindow.setTimeout(applyFix, 80);
                hostWindow.setTimeout(applyFix, 240);
                hostWindow.setTimeout(applyFix, 700);

                const observer = new MutationObserver(function () {
                    applyFix();
                });

                observer.observe(hostDocument.body, { childList: true, subtree: true });
                hostWindow.setTimeout(function () { observer.disconnect(); }, 6000);
            })();
        </script>
        """,
        height=0,
        scrolling=False,
    )


# =========================================================
# CSS FINAL: TÍTULOS DOS FILTROS DA VISÃO GERAL EM BRANCO
# =========================================================
def apply_overview_filter_labels_white_css() -> None:
    """Força os títulos customizados dos filtros da Visão Geral a permanecerem brancos."""
    render_html(
        """
        <style>
            /* Este bloco é aplicado globalmente no dashboard, inclusive na Visão Geral.
               O ajuste anterior estava dentro do CSS exclusivo das páginas de cadastro. */
            .overview-filter-custom-label,
            .overview-filter-custom-label *,
            .st-key-overview_filter_labels .overview-filter-custom-label,
            .st-key-overview_filter_labels .overview-filter-custom-label *,
            .st-key-overview_filter_labels [data-testid="stMarkdownContainer"],
            .st-key-overview_filter_labels [data-testid="stMarkdownContainer"] *,
            .st-key-overview_filter_labels p,
            .st-key-overview_filter_labels span {
                color: #FFFFFF !important;
                -webkit-text-fill-color: #FFFFFF !important;
                opacity: 1 !important;
            }
        </style>
        """
    )


# =========================================================
# TRATAMENTO DE ERROS
# =========================================================
def render_connection_error(error: Exception) -> None:
    apply_dashboard_css()

    render_html(
        """
        <div class="page-title">Dashboard Oppi Comercial</div>
        <div class="page-subtitle">Erro ao conectar com a planilha.</div>
        """
    )

    if isinstance(error, SpreadsheetNotFound):
        st.error(
            "A credencial foi aceita, mas a planilha não foi localizada. "
            "Confirme se o SHEET_ID está correto e se a planilha foi compartilhada "
            "diretamente com o e-mail da conta de serviço."
        )
        st.code(SHEET_ID)
        return

    if isinstance(error, WorksheetNotFound):
        st.error(
            f"A planilha foi localizada, mas não encontrei a aba '{WORKSHEET_NAME}'."
        )
        return

    st.error("Não consegui carregar os dados da planilha.")
    st.code(str(error))


# =========================================================
# APLICAÇÃO PRINCIPAL
# =========================================================
def main() -> None:
    restore_navigation_session_from_url()

    if not st.session_state.authenticated:
        render_login_page()
        return

    apply_dashboard_css()
    apply_overview_filter_labels_white_css()
    apply_global_sidebar_toggle_css()
    apply_final_sidebar_toggle_override_css()
    page = render_sidebar()
    install_sidebar_navigation_persistence()

    try:
        df = load_sheet_data()
    except Exception as error:
        render_connection_error(error)
        return

    if df.empty:
        render_html(
            """
            <div class="page-title">Dashboard Oppi Comercial</div>
            <div class="page-subtitle">A conexão foi realizada, mas a planilha está vazia.</div>
            """
        )
        st.warning("A planilha foi encontrada, mas não possui registros preenchidos.")
        return

    columns = identify_columns(df)
    prepared_df = prepare_data(df, columns)

    if page == "Visão Geral":
        render_overview_page(prepared_df, columns)
    elif page == "Cadastro":
        cadastro_subpage = st.session_state.get("selected_cadastro_subpage", "Novo cadastro")

        if cadastro_subpage == "Todos os cadastros":
            render_all_contracts_page(prepared_df, columns)
        else:
            render_proposals_page(prepared_df, columns)
    elif page == "Pesos e Medidas":
        render_scoring_page(prepared_df, columns)


if __name__ == "__main__":
    main()
