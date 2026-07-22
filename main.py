import io
import os
import re
import sys
import threading
import time
import unicodedata
from datetime import datetime
from typing import Any, Optional, Union
from urllib.parse import quote
from zipfile import BadZipFile

import msal
import pandas as pd
import requests
from dotenv import load_dotenv
from fastapi import BackgroundTasks, FastAPI
from pydantic import BaseModel


load_dotenv()

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
SUPPORTED_EXTENSIONS = (".xls", ".xlsx", ".csv")
PAD_COLUMNS = ["EGR", "CODIGO_PROVEEDOR", "DETALLE", "DETALLE_COMPLETO"]
CONTROL_CSV_PREFIX = "CSV_"
CONTROL_LOG_PREFIX = "LOG_"
DEFAULT_FACTURAS_BASE_PATH = "SPORTEK/2026/05 COMISIONES/5. XML FACTURAS/7. JULIO"
FACTURA_SEARCH_MAX_ATTEMPTS = 25
FACTURA_SEARCH_RETRY_DELAY_SECONDS = int(
    os.environ.get("GRAPH_FACTURA_SEARCH_RETRY_SECONDS", "10")
)
_process_lock = threading.Lock()
_current_logs: Optional[list[str]] = None


def log_event(message: str) -> None:
    """Escribe en consola y, si hay un buffer activo, guarda el evento para el LOG."""
    sys.stdout.write(f"{message}\n")
    sys.stdout.flush()
    if _current_logs is not None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        _current_logs.append(f"[{timestamp}] {message}")

app = FastAPI(
    title="Ecuabet RPA SharePoint Service",
    version="1.0.0",
    description="Procesa egresos desde SharePoint y genera el CSV de control para PAD.",
)


class EgresosRequest(BaseModel):
    ruta_completa: str


class IgnoredResponse(BaseModel):
    status: str
    message: str


class AcceptedResponse(BaseModel):
    status: str
    message: str


def get_required_env(variable_name: str) -> str:
    value = os.environ.get(variable_name)
    if not value:
        raise RuntimeError(f"Falta la variable de entorno requerida: {variable_name}")

    return value


def get_config() -> dict[str, str]:
    return {
        "tenant_id": get_required_env("GRAPH_TENANT_ID"),
        "client_id": get_required_env("GRAPH_CLIENT_ID"),
        "client_secret": get_required_env("GRAPH_CLIENT_SECRET"),
        "site_id": os.environ.get("GRAPH_SITE_ID") or os.environ.get("SITE_ID") or "",
        "drive_id": get_required_env("GRAPH_DRIVE_ID"),
        "facturas_base_path": os.environ.get(
            "GRAPH_FACTURAS_BASE_PATH",
            DEFAULT_FACTURAS_BASE_PATH,
        ),
    }


def get_access_token(config: dict[str, str]) -> str:
    authority = f"https://login.microsoftonline.com/{config['tenant_id']}"
    app_client = msal.ConfidentialClientApplication(
        client_id=config["client_id"],
        client_credential=config["client_secret"],
        authority=authority,
    )

    result = app_client.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in result:
        error = result.get("error_description") or result.get("error") or "Error desconocido"
        raise RuntimeError(f"No se pudo obtener el Access Token: {error}")

    return result["access_token"]


def build_headers(access_token: str, content_type: Optional[str] = None) -> dict[str, str]:
    headers = {"Authorization": f"Bearer {access_token}"}
    if content_type:
        headers["Content-Type"] = content_type

    return headers


def build_content_url(config: dict[str, str], clean_path: str) -> str:
    drive_id = config["drive_id"]
    site_id = config.get("site_id")
    if not site_id:
        raise RuntimeError("Falta GRAPH_SITE_ID o SITE_ID para descargar por ruta")

    encoded_path = quote(clean_path, safe="/")
    return f"{GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/content"


def download_file_bytes_by_path(
    access_token: str,
    config: dict[str, str],
    clean_path: str,
) -> bytes:
    url = build_content_url(config, clean_path)
    response = requests.get(url, headers=build_headers(access_token), timeout=120)
    response.raise_for_status()
    return response.content


def read_report_dataframe(
    file_bytes: bytes,
    file_name: str,
    *,
    nrows: Optional[int] = None,
    header: Optional[int] = 0,
) -> pd.DataFrame:
    stream = io.BytesIO(file_bytes)
    normalized_name = file_name.lower()

    if normalized_name.endswith(".csv"):
        return pd.read_csv(stream, nrows=nrows, header=header)

    if normalized_name.endswith(".xlsx"):
        return pd.read_excel(
            stream,
            nrows=nrows,
            header=header,
            engine="openpyxl",
        )

    if normalized_name.endswith(".xls"):
        try:
            return pd.read_excel(stream, nrows=nrows, header=header)
        except (ValueError, BadZipFile, ImportError, OSError) as excel_error:
            log_event(
                f"Lectura Excel fallo para {file_name}; "
                f"intentando contingencia CSV: {excel_error}"
            )
            stream.seek(0)
            return pd.read_csv(
                stream,
                sep=",",
                nrows=nrows,
                header=header,
                on_bad_lines="skip",
            )

    raise ValueError(f"Extension no soportada para el archivo: {file_name}")


def read_stream_dataframe(
    stream: io.BytesIO,
    file_name: str,
    *,
    nrows: Optional[int] = None,
    header: Optional[int] = 0,
) -> pd.DataFrame:
    stream.seek(0)
    normalized_name = file_name.lower()

    if normalized_name.endswith(".csv"):
        return pd.read_csv(stream, nrows=nrows, header=header, on_bad_lines="skip")

    if normalized_name.endswith(".xlsx"):
        return pd.read_excel(stream, nrows=nrows, header=header, engine="openpyxl")

    if normalized_name.endswith(".xls"):
        try:
            return pd.read_excel(stream, nrows=nrows, header=header)
        except (ValueError, BadZipFile, ImportError, OSError) as excel_error:
            log_event(
                f"Lectura Excel fallo para {file_name}; "
                f"intentando contingencia CSV: {excel_error}"
            )
            stream.seek(0)
            return pd.read_csv(
                stream,
                sep=",",
                nrows=nrows,
                header=header,
                on_bad_lines="skip",
            )

    raise ValueError(f"Extension no soportada para el archivo: {file_name}")


def leer_archivo_inteligente(stream: io.BytesIO, file_name: str) -> pd.DataFrame:
    preview = read_stream_dataframe(stream, file_name, nrows=20, header=None)
    expected_headers = {"DOCUMENTO", "CUENTA"}

    for row_index, row in preview.iterrows():
        normalized_values = {
            str(value).strip().upper()
            for value in row.dropna().tolist()
        }
        if normalized_values.intersection(expected_headers):
            header_row = int(row_index)
            log_event(f"Fila de encabezados detectada en indice {header_row} para {file_name}")
            dataframe = read_stream_dataframe(stream, file_name, header=header_row)
            dataframe.columns = dataframe.columns.astype(str).str.strip()
            return dataframe

    log_event(
        f"No se encontro encabezado DOCUMENTO/CUENTA en {file_name}; "
        "aplicando lectura posicional"
    )
    return leer_archivo_posicional(stream, file_name)


def clean_numeric_text(value: Any) -> str:
    text = str(value).strip()
    text = re.sub(r"\.0$", "", text)
    return text


def leer_archivo_posicional(stream: io.BytesIO, file_name: str) -> pd.DataFrame:
    raw_dataframe = read_stream_dataframe(stream, file_name, header=None)
    raw_dataframe = raw_dataframe.dropna(how="all").copy()

    if raw_dataframe.shape[1] < 3:
        raise ValueError(
            f"No se encontro fila de encabezados ni estructura posicional valida en {file_name}"
        )

    provider_codes = raw_dataframe.iloc[:, 1].apply(clean_numeric_text)
    provider_names = raw_dataframe.iloc[:, 2].fillna("").astype(str).str.strip()
    valid_rows = provider_codes.str.fullmatch(r"\d+", na=False) & (provider_names != "")

    positional = pd.DataFrame(
        {
            "CUENTA": "Proveedores Locales",
            "CODIGO_PROVEEDOR": provider_codes[valid_rows],
            "DETALLE": provider_names[valid_rows],
        }
    )
    positional["DETALLE_COMPLETO"] = (
        positional["CODIGO_PROVEEDOR"] + " " + positional["DETALLE"]
    )

    log_event(f"Filas posicionales validas detectadas en {file_name}: {len(positional)}")
    return positional.reset_index(drop=True)


def filtrar_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    required_columns = {"CUENTA", "DETALLE"}
    missing_columns = required_columns - set(dataframe.columns)
    if missing_columns:
        missing_text = ", ".join(sorted(missing_columns))
        raise ValueError(f"Faltan columnas requeridas: {missing_text}")

    account_column = "NOMBRE CTA." if "NOMBRE CTA." in dataframe.columns else "CUENTA"
    log_event(f"Filtrando Proveedores Locales usando columna: {account_column}")

    account_series = dataframe[account_column].astype(str)
    return dataframe[
        account_series.str.contains("Proveedores Locales", case=False, na=False)
    ].copy()


def extract_provider_from_detail(detail_value: Any) -> tuple[str, str]:
    detail_text = str(detail_value).strip()
    match = re.search(r"\bFC\b\s*([0-9]+)\s+(.+)$", detail_text, flags=re.IGNORECASE)
    if not match:
        match = re.search(r"^\s*([0-9]+)\s+(.+)$", detail_text)
        if not match:
            return "", ""

    provider_code = match.group(1).strip()
    provider_name = match.group(2).strip()
    return provider_code, provider_name


def extract_commissioner_transactions(
    file_bytes: bytes,
    file_name: str,
    egr: str,
) -> pd.DataFrame:
    stream = io.BytesIO(file_bytes)
    dataframe = leer_archivo_inteligente(stream, file_name)
    filtered = filtrar_dataframe(dataframe)
    del dataframe

    if filtered.empty:
        log_event(f"No hay comisionistas validos en {file_name}")
        return pd.DataFrame(columns=PAD_COLUMNS)

    filtered["EGR"] = egr
    if {"CODIGO_PROVEEDOR", "DETALLE_COMPLETO"}.issubset(filtered.columns):
        filtered["CODIGO_PROVEEDOR"] = filtered["CODIGO_PROVEEDOR"].astype(str).str.strip()
        filtered["DETALLE"] = filtered["DETALLE"].fillna("").astype(str).str.strip()
        filtered["DETALLE_COMPLETO"] = (
            filtered["DETALLE_COMPLETO"].fillna("").astype(str).str.strip()
        )
    else:
        detail_complete = filtered["DETALLE"].fillna("").astype(str).str.strip()
        provider_data = detail_complete.apply(extract_provider_from_detail)
        filtered["CODIGO_PROVEEDOR"] = provider_data.apply(lambda value: value[0])
        filtered["DETALLE"] = provider_data.apply(lambda value: value[1])
        filtered["DETALLE_COMPLETO"] = detail_complete

    transactions = filtered[
        (filtered["CODIGO_PROVEEDOR"] != "") & (filtered["DETALLE"] != "")
    ][PAD_COLUMNS].copy()

    log_event(f"{file_name}: transacciones validas filtradas: {len(transactions)}")
    return transactions


def sanitize_folder_name(folder_name: str) -> str:
    invalid_chars = '"*:<>?/\\|'
    sanitized_name = "".join("-" if char in invalid_chars else char for char in folder_name)
    sanitized_name = sanitized_name.strip().rstrip(".")
    if not sanitized_name:
        raise ValueError("El nombre de carpeta quedo vacio despues de sanitizar")

    return sanitized_name


def normalize_graph_path(path_value: str) -> str:
    clean_path = path_value.strip().replace("\\", "/").strip("/")
    prefixes = ("Documentos compartidos/", "Shared Documents/")

    for prefix in prefixes:
        if clean_path.casefold().startswith(prefix.casefold()):
            clean_path = clean_path[len(prefix) :]
            break

    if not clean_path or "/" not in clean_path:
        raise ValueError(
            "ruta_completa debe incluir carpeta padre y nombre de archivo"
        )

    return clean_path


def extract_file_name_from_path(clean_path: str) -> str:
    file_name = clean_path.rstrip("/").split("/")[-1].strip()
    if not file_name:
        raise ValueError("No se pudo extraer el nombre de archivo desde ruta_completa")

    return file_name


def should_process_excel_path(path_value: str) -> bool:
    file_name = path_value.strip().replace("\\", "/").rstrip("/").split("/")[-1]
    return file_name.lower().endswith((".xls", ".xlsx"))


def extract_parent_folder_from_path(clean_path: str) -> str:
    path_parts = [part.strip() for part in clean_path.split("/") if part.strip()]
    if len(path_parts) < 2:
        raise ValueError("No se pudo extraer la carpeta padre desde ruta_completa")

    return path_parts[-2]


def extract_parent_path_from_path(clean_path: str) -> str:
    path_parts = [part.strip() for part in clean_path.split("/") if part.strip()]
    if len(path_parts) < 2:
        raise ValueError("No se pudo extraer la ruta padre desde ruta_completa")

    return "/".join(path_parts[:-1])


def get_drive_item_by_path(
    access_token: str,
    config: dict[str, str],
    clean_path: str,
) -> dict[str, Any]:
    drive_id = config["drive_id"]
    site_id = config.get("site_id")
    if not site_id:
        raise RuntimeError("Falta GRAPH_SITE_ID o SITE_ID para resolver ruta en Graph")

    encoded_path = quote(clean_path, safe="/")
    url = f"{GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}"
    response = requests.get(url, headers=build_headers(access_token), timeout=60)
    response.raise_for_status()
    return response.json()


def list_child_folders(
    access_token: str,
    drive_id: str,
    parent_folder_id: str,
) -> list[dict[str, Any]]:
    url = f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{parent_folder_id}/children"
    folders: list[dict[str, Any]] = []

    while url:
        response = requests.get(url, headers=build_headers(access_token), timeout=60)
        response.raise_for_status()

        payload = response.json()
        folders.extend(item for item in payload.get("value", []) if "folder" in item)
        url = payload.get("@odata.nextLink")

    return folders


def list_drive_children_by_path(
    access_token: str,
    config: dict[str, str],
    folder_path: str,
) -> list[dict[str, Any]]:
    drive_id = config["drive_id"]
    site_id = config.get("site_id")
    if not site_id:
        raise RuntimeError("Falta GRAPH_SITE_ID o SITE_ID para listar facturas")

    encoded_path = quote(folder_path.strip("/"), safe="/")
    url = f"{GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
    items: list[dict[str, Any]] = []

    while url:
        response = requests.get(url, headers=build_headers(access_token), timeout=60)
        response.raise_for_status()
        payload = response.json()
        items.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")

    return items


def list_pdf_items_recursive(
    access_token: str,
    config: dict[str, str],
    folder_path: str,
) -> list[dict[str, Any]]:
    pdf_items: list[dict[str, Any]] = []

    try:
        children = list_drive_children_by_path(access_token, config, folder_path)
    except requests.HTTPError:
        return pdf_items

    for item in children:
        item_name = item.get("name", "")
        if "file" in item and item_name.lower().endswith(".pdf"):
            pdf_items.append(item)
            continue

        if "folder" in item:
            child_path = f"{folder_path.rstrip('/')}/{item_name}"
            pdf_items.extend(list_pdf_items_recursive(access_token, config, child_path))

    return pdf_items


def download_drive_item_bytes(
    access_token: str,
    config: dict[str, str],
    item_id: str,
) -> bytes:
    drive_id = config["drive_id"]
    url = f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}/content"
    response = requests.get(url, headers=build_headers(access_token), timeout=120)
    response.raise_for_status()
    return response.content


def upload_bytes_to_folder(
    access_token: str,
    config: dict[str, str],
    folder_id: str,
    file_name: str,
    file_bytes: bytes,
    content_type: str,
) -> str:
    drive_id = config["drive_id"]
    encoded_name = quote(file_name, safe="")
    url = f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{folder_id}:/{encoded_name}:/content"
    response = requests.put(
        url,
        headers=build_headers(access_token, content_type),
        data=file_bytes,
        timeout=120,
    )
    response.raise_for_status()
    uploaded_file = response.json()
    return uploaded_file.get("webUrl", file_name)


def normalize_match_text(value: Any) -> str:
    text = str(value).lower().replace("_", " ")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def invoice_matches_provider(
    pdf_name: str,
    provider_code: str,
    provider_name: str,
) -> bool:
    pdf_digits = re.sub(r"\D+", "", pdf_name)
    provider_digits = re.sub(r"\D+", "", str(provider_code))
    if len(provider_digits) >= 10 and provider_digits in pdf_digits:
        return True

    normalized_pdf = normalize_match_text(pdf_name)
    normalized_provider = normalize_match_text(provider_name)
    if not normalized_provider:
        return False

    return (
        normalized_provider in normalized_pdf
        or normalized_provider.replace(" ", "") in normalized_pdf.replace(" ", "")
    )


def find_matching_invoice_item(
    invoice_items: list[dict[str, Any]],
    provider_code: str,
    provider_name: str,
) -> Optional[dict[str, Any]]:
    for item in invoice_items:
        pdf_name = item.get("name", "")
        if "file" not in item or not pdf_name.lower().endswith(".pdf"):
            continue

        if invoice_matches_provider(pdf_name, provider_code, provider_name):
            return item

    return None


def copiar_factura_pdf_proveedor(
    access_token: str,
    config: dict[str, str],
    transaction: dict[str, Any],
    destination_folder_id: str,
) -> bool:
    provider_code = str(transaction.get("CODIGO_PROVEEDOR", "")).strip()
    provider_name = str(transaction.get("DETALLE", "")).strip()
    provider_label = f"{provider_code} {provider_name}".strip()
    invoice_month_path = config["facturas_base_path"].rstrip("/")

    for attempt in range(1, FACTURA_SEARCH_MAX_ATTEMPTS + 1):
        invoice_items = list_pdf_items_recursive(access_token, config, invoice_month_path)
        matched_item = find_matching_invoice_item(
            invoice_items,
            provider_code,
            provider_name,
        )
        if matched_item:
            pdf_name = matched_item.get("name", "")
            pdf_bytes = download_drive_item_bytes(access_token, config, matched_item["id"])
            uploaded_url = upload_bytes_to_folder(
                access_token,
                config,
                destination_folder_id,
                pdf_name,
                pdf_bytes,
                "application/pdf",
            )
            log_event(
                f"Factura copiada para {provider_label} "
                f"(intento {attempt}/{FACTURA_SEARCH_MAX_ATTEMPTS}): {uploaded_url}"
            )
            return True

        if attempt < FACTURA_SEARCH_MAX_ATTEMPTS:
            log_event(
                f"Factura no encontrada para {provider_label} "
                f"(intento {attempt}/{FACTURA_SEARCH_MAX_ATTEMPTS}); "
                f"reintentando en {FACTURA_SEARCH_RETRY_DELAY_SECONDS}s..."
            )
            time.sleep(FACTURA_SEARCH_RETRY_DELAY_SECONDS)

    log_event(
        f"Factura no encontrada para {provider_label} "
        f"tras {FACTURA_SEARCH_MAX_ATTEMPTS} intentos; se continua con la siguiente"
    )
    return False


def find_child_folder(
    access_token: str,
    drive_id: str,
    parent_folder_id: str,
    folder_name: str,
) -> Optional[dict[str, Any]]:
    normalized_name = folder_name.casefold()
    for folder in list_child_folders(access_token, drive_id, parent_folder_id):
        if folder.get("name", "").casefold() == normalized_name:
            return folder

    return None


def ensure_folder(
    access_token: str,
    drive_id: str,
    parent_folder_id: str,
    folder_name: str,
    level_label: str,
) -> dict[str, Any]:
    safe_folder_name = sanitize_folder_name(folder_name)
    log_event(f"Validando carpeta {level_label}: {safe_folder_name}")

    existing_folder = find_child_folder(
        access_token,
        drive_id,
        parent_folder_id,
        safe_folder_name,
    )
    if existing_folder:
        log_event(f"Carpeta {level_label} existente: {existing_folder['name']}")
        return existing_folder

    log_event(f"Creando carpeta {level_label}: {safe_folder_name}")
    url = f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{parent_folder_id}/children"
    payload = {
        "name": safe_folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "replace",
    }

    response = requests.post(
        url,
        headers=build_headers(access_token, "application/json"),
        json=payload,
        timeout=60,
    )
    if response.status_code == 409:
        existing_folder = find_child_folder(
            access_token,
            drive_id,
            parent_folder_id,
            safe_folder_name,
        )
        if existing_folder:
            log_event(f"Carpeta {level_label} encontrada tras conflicto")
            return existing_folder

    response.raise_for_status()
    created_folder = response.json()
    log_event(f"Carpeta {level_label} creada: {created_folder.get('name', safe_folder_name)}")
    return created_folder


def create_destination_hierarchy(
    access_token: str,
    config: dict[str, str],
    parent_folder_id: str,
    egreso_base_name: str,
    carpeta_fecha: str,
    transactions: pd.DataFrame,
) -> None:
    drive_id = config["drive_id"]
    egreso_folder = ensure_folder(
        access_token,
        drive_id,
        parent_folder_id,
        egreso_base_name,
        "Nivel 2 Egreso",
    )

    unique_transactions = transactions.drop_duplicates(
        subset=["CODIGO_PROVEEDOR", "DETALLE"]
    )
    transaction_records = unique_transactions.to_dict("records")
    log_event(f"Creando/validando carpetas FC unicas: {len(transaction_records)}")

    fc_destinations: list[tuple[dict[str, Any], dict[str, Any]]] = []
    for transaction in transaction_records:
        folder_name = f"FC {transaction['CODIGO_PROVEEDOR']} {transaction['DETALLE']}"
        fc_folder = ensure_folder(
            access_token,
            drive_id,
            egreso_folder["id"],
            folder_name,
            "Nivel 3 Factura Comisionista",
        )
        fc_destinations.append((transaction, fc_folder))

    log_event(
        f"Carpetas FC listas. Iniciando copia de facturas con hasta "
        f"{FACTURA_SEARCH_MAX_ATTEMPTS} intentos por proveedor"
    )
    for transaction, fc_folder in fc_destinations:
        copiar_factura_pdf_proveedor(
            access_token,
            config,
            transaction,
            fc_folder["id"],
        )


def build_control_csv_bytes(transactions: list[pd.DataFrame]) -> bytes:
    if transactions:
        consolidated = pd.concat(transactions, ignore_index=True)
    else:
        consolidated = pd.DataFrame(columns=PAD_COLUMNS)

    consolidated = consolidated[PAD_COLUMNS]
    csv_stream = io.StringIO()
    consolidated.to_csv(csv_stream, index=False)
    return csv_stream.getvalue().encode("utf-8-sig")


def get_file_base_name(file_name: str) -> str:
    base_name = os.path.basename(file_name)
    file_base_name, _ = os.path.splitext(base_name)
    if not file_base_name:
        raise ValueError(f"No se pudo obtener el nombre base de {file_name}")

    return file_base_name


def build_control_csv_name(file_name: str) -> str:
    file_base_name = sanitize_folder_name(get_file_base_name(file_name))
    return f"{CONTROL_CSV_PREFIX}{file_base_name}.csv"


def build_control_log_name(file_name: str) -> str:
    file_base_name = sanitize_folder_name(get_file_base_name(file_name))
    return f"{CONTROL_LOG_PREFIX}{file_base_name}.txt"


def upload_text_file(
    access_token: str,
    config: dict[str, str],
    parent_path: str,
    file_bytes: bytes,
    file_name: str,
    content_type: str,
) -> str:
    drive_id = config["drive_id"]
    site_id = config.get("site_id")
    if not site_id:
        raise RuntimeError("Falta GRAPH_SITE_ID o SITE_ID para subir archivo por ruta")

    upload_path = f"{parent_path.rstrip('/')}/{file_name}"
    encoded_upload_path = quote(upload_path, safe="/")
    url = f"{GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{encoded_upload_path}:/content"
    response = requests.put(
        url,
        headers=build_headers(access_token, content_type),
        data=file_bytes,
        timeout=120,
    )
    response.raise_for_status()

    uploaded_file = response.json()
    return uploaded_file.get("webUrl", file_name)


def upload_control_csv(
    access_token: str,
    config: dict[str, str],
    parent_path: str,
    csv_bytes: bytes,
    csv_name: str,
) -> str:
    drive_id = config["drive_id"]
    site_id = config.get("site_id")
    if not site_id:
        raise RuntimeError("Falta GRAPH_SITE_ID o SITE_ID para subir CSV por ruta")

    upload_path = f"{parent_path.rstrip('/')}/{csv_name}"
    encoded_upload_path = quote(upload_path, safe="/")
    url = f"{GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{encoded_upload_path}:/content"
    response = requests.put(
        url,
        headers=build_headers(access_token, "text/csv"),
        data=csv_bytes,
        timeout=120,
    )
    response.raise_for_status()

    uploaded_file = response.json()
    web_url = uploaded_file.get("webUrl", csv_name)
    log_event(f"CSV de control subido a SharePoint: {web_url}")
    return web_url


def process_file(
    access_token: str,
    config: dict[str, str],
    clean_path: str,
    nombre_archivo: str,
    parent_folder_id: str,
    egreso_base_name: str,
    carpeta_fecha: str,
) -> pd.DataFrame:
    log_event(f"Procesando archivo: {nombre_archivo}")
    file_bytes = download_file_bytes_by_path(access_token, config, clean_path)

    egr = egreso_base_name
    log_event(f"EGR definido desde nombre de archivo: {egr}")

    transactions = extract_commissioner_transactions(file_bytes, nombre_archivo, egr)
    create_destination_hierarchy(
        access_token,
        config,
        parent_folder_id,
        egreso_base_name,
        carpeta_fecha,
        transactions,
    )
    return transactions


def upload_process_log(
    access_token: Optional[str],
    config: Optional[dict[str, str]],
    parent_path: Optional[str],
    nombre_archivo: Optional[str],
    log_lines: list[str],
) -> None:
    """Sube el buffer de logs como LOG_<egreso>.txt junto al Excel (best-effort)."""
    if not (access_token and config and parent_path and nombre_archivo):
        sys.stdout.write(
            "No se pudo subir el LOG a SharePoint: falta contexto "
            "(token, config, ruta o nombre de archivo)\n"
        )
        return

    try:
        log_name = build_control_log_name(nombre_archivo)
        log_bytes = ("\n".join(log_lines) + "\n").encode("utf-8-sig")
        log_url = upload_text_file(
            access_token,
            config,
            parent_path,
            log_bytes,
            log_name,
            "text/plain; charset=utf-8",
        )
        sys.stdout.write(f"LOG de proceso subido a SharePoint: {log_url}\n")
        sys.stdout.flush()
    except Exception as log_error:
        sys.stdout.write(f"No se pudo subir el LOG a SharePoint: {log_error}\n")
        sys.stdout.flush()


def procesar_egreso_background(ruta_completa: str) -> None:
    """Ejecuta todo el flujo pesado en segundo plano y sube CSV + LOG a SharePoint."""
    global _current_logs

    log_event("Solicitud en cola. Esperando turno de procesamiento...")
    with _process_lock:
        _current_logs = []
        access_token: Optional[str] = None
        config: Optional[dict[str, str]] = None
        parent_path: Optional[str] = None
        nombre_archivo: Optional[str] = None

        try:
            log_event(
                "Turno adquirido. Iniciando procesamiento en segundo plano: "
                f"{ruta_completa}"
            )
            config = get_config()
            access_token = get_access_token(config)
            clean_path = normalize_graph_path(ruta_completa)
            nombre_archivo = extract_file_name_from_path(clean_path)
            carpeta_fecha = extract_parent_folder_from_path(clean_path)
            parent_path = extract_parent_path_from_path(clean_path)
            parent_folder = get_drive_item_by_path(access_token, config, parent_path)
            csv_name = build_control_csv_name(nombre_archivo)
            egreso_base_name = sanitize_folder_name(get_file_base_name(nombre_archivo))
            transactions = pd.DataFrame(columns=PAD_COLUMNS)

            try:
                transactions = process_file(
                    access_token,
                    config,
                    clean_path,
                    nombre_archivo,
                    parent_folder["id"],
                    egreso_base_name,
                    carpeta_fecha,
                )
            except Exception as file_error:
                log_event(f"Error procesando {nombre_archivo}: {file_error}")

            csv_bytes = build_control_csv_bytes(
                [transactions] if not transactions.empty else []
            )
            uploaded_url = upload_control_csv(
                access_token, config, parent_path, csv_bytes, csv_name
            )
            log_event(
                "Proceso finalizado. "
                f"Transacciones exportadas: {len(transactions)}; "
                f"CSV de control: {uploaded_url}"
            )
        except Exception as error:
            log_event(f"Error general procesando egresos: {error}")
        finally:
            log_lines = list(_current_logs or [])
            upload_process_log(
                access_token,
                config,
                parent_path,
                nombre_archivo,
                log_lines,
            )
            _current_logs = None


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post(
    "/api/v1/procesar-egresos",
    status_code=202,
    response_model=Union[AcceptedResponse, IgnoredResponse],
)
def procesar_egresos(
    request: EgresosRequest,
    background_tasks: BackgroundTasks,
) -> Union[AcceptedResponse, IgnoredResponse]:
    if not should_process_excel_path(request.ruta_completa):
        log_event(f"Archivo ignorado para evitar bucle: {request.ruta_completa}")
        return IgnoredResponse(
            status="ignored",
            message="El archivo no es un Excel, se ignora para evitar bucles",
        )

    log_event(
        "Solicitud recibida para procesar egreso: "
        f"{request.ruta_completa}"
    )
    background_tasks.add_task(procesar_egreso_background, request.ruta_completa)
    return AcceptedResponse(
        status="accepted",
        message="Procesamiento en segundo plano iniciado",
    )
