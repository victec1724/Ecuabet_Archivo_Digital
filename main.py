import io
import os
import re
from typing import Any, Optional, Union
from urllib.parse import quote
from zipfile import BadZipFile

import msal
import pandas as pd
import requests
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel


load_dotenv()

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
SUPPORTED_EXTENSIONS = (".xls", ".xlsx", ".csv")
PAD_COLUMNS = ["EGR", "CODIGO_PROVEEDOR", "DETALLE", "DETALLE_COMPLETO"]
CONTROL_CSV_PREFIX = "CSV_"

app = FastAPI(
    title="Ecuabet RPA SharePoint Service",
    version="1.0.0",
    description="Procesa egresos desde SharePoint y genera el CSV de control para PAD.",
)


class EgresosRequest(BaseModel):
    ruta_completa: str


class FileError(BaseModel):
    archivo: str
    error: str


class IgnoredResponse(BaseModel):
    status: str
    message: str


class ProcessResponse(BaseModel):
    status: str
    ruta_completa: str
    ruta_limpia: str
    nombre_archivo: str
    carpeta_fecha: str
    csv_generado: str
    transacciones_exportadas: int
    archivo_control: str
    errores: list[FileError]


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
            print(
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
            print(
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
            print(f"Fila de encabezados detectada en indice {header_row} para {file_name}")
            dataframe = read_stream_dataframe(stream, file_name, header=header_row)
            dataframe.columns = dataframe.columns.astype(str).str.strip()
            return dataframe

    raise ValueError(
        f"No se encontro fila de encabezados con DOCUMENTO o CUENTA en {file_name}"
    )


def filtrar_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    required_columns = {"CUENTA", "DETALLE"}
    missing_columns = required_columns - set(dataframe.columns)
    if missing_columns:
        missing_text = ", ".join(sorted(missing_columns))
        raise ValueError(f"Faltan columnas requeridas: {missing_text}")

    account_column = "NOMBRE CTA." if "NOMBRE CTA." in dataframe.columns else "CUENTA"
    print(f"Filtrando Proveedores Locales usando columna: {account_column}")

    account_series = dataframe[account_column].astype(str)
    return dataframe[
        account_series.str.contains("Proveedores Locales", case=False, na=False)
    ].copy()


def extract_provider_from_detail(detail_value: Any) -> tuple[str, str]:
    detail_text = str(detail_value).strip()
    match = re.search(r"\bFC\b\s*([0-9]+)\s+(.+)$", detail_text, flags=re.IGNORECASE)
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
        print(f"No hay comisionistas validos en {file_name}")
        return pd.DataFrame(columns=PAD_COLUMNS)

    detail_complete = filtered["DETALLE"].fillna("").astype(str).str.strip()
    provider_data = detail_complete.apply(extract_provider_from_detail)
    filtered["EGR"] = egr
    filtered["CODIGO_PROVEEDOR"] = provider_data.apply(lambda value: value[0])
    filtered["DETALLE"] = provider_data.apply(lambda value: value[1])
    filtered["DETALLE_COMPLETO"] = detail_complete

    transactions = filtered[
        (filtered["CODIGO_PROVEEDOR"] != "") & (filtered["DETALLE"] != "")
    ][PAD_COLUMNS].copy()

    print(f"{file_name}: transacciones validas filtradas: {len(transactions)}")
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
    print(f"Validando carpeta {level_label}: {safe_folder_name}")

    existing_folder = find_child_folder(
        access_token,
        drive_id,
        parent_folder_id,
        safe_folder_name,
    )
    if existing_folder:
        print(f"Carpeta {level_label} existente: {existing_folder['name']}")
        return existing_folder

    print(f"Creando carpeta {level_label}: {safe_folder_name}")
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
            print(f"Carpeta {level_label} encontrada tras conflicto")
            return existing_folder

    response.raise_for_status()
    created_folder = response.json()
    print(f"Carpeta {level_label} creada: {created_folder.get('name', safe_folder_name)}")
    return created_folder


def create_destination_hierarchy(
    access_token: str,
    config: dict[str, str],
    parent_folder_id: str,
    egreso_base_name: str,
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
    print(f"Creando/validando carpetas FC unicas: {len(unique_transactions)}")

    for transaction in unique_transactions.to_dict("records"):
        folder_name = f"FC {transaction['CODIGO_PROVEEDOR']} {transaction['DETALLE']}"
        ensure_folder(
            access_token,
            drive_id,
            egreso_folder["id"],
            folder_name,
            "Nivel 3 Factura Comisionista",
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
    print(f"CSV de control subido a SharePoint: {web_url}")
    return web_url


def process_file(
    access_token: str,
    config: dict[str, str],
    clean_path: str,
    nombre_archivo: str,
    parent_folder_id: str,
    egreso_base_name: str,
) -> pd.DataFrame:
    print(f"Procesando archivo: {nombre_archivo}")
    file_bytes = download_file_bytes_by_path(access_token, config, clean_path)

    egr = egreso_base_name
    print(f"EGR definido desde nombre de archivo: {egr}")

    transactions = extract_commissioner_transactions(file_bytes, nombre_archivo, egr)
    create_destination_hierarchy(
        access_token,
        config,
        parent_folder_id,
        egreso_base_name,
        transactions,
    )
    return transactions


def run_egresos_process(ruta_completa: str) -> ProcessResponse:
    config = get_config()
    access_token = get_access_token(config)
    clean_path = normalize_graph_path(ruta_completa)
    nombre_archivo = extract_file_name_from_path(clean_path)
    carpeta_fecha = extract_parent_folder_from_path(clean_path)
    parent_path = extract_parent_path_from_path(clean_path)
    parent_folder = get_drive_item_by_path(access_token, config, parent_path)
    errors: list[FileError] = []
    transactions = pd.DataFrame(columns=PAD_COLUMNS)
    csv_name = build_control_csv_name(nombre_archivo)
    egreso_base_name = sanitize_folder_name(get_file_base_name(nombre_archivo))

    try:
        transactions = process_file(
            access_token,
            config,
            clean_path,
            nombre_archivo,
            parent_folder["id"],
            egreso_base_name,
        )
    except requests.HTTPError:
        raise
    except Exception as file_error:
        print(f"Error procesando {nombre_archivo}: {file_error}")
        errors.append(FileError(archivo=nombre_archivo, error=str(file_error)))

    csv_bytes = build_control_csv_bytes([transactions] if not transactions.empty else [])
    uploaded_url = upload_control_csv(access_token, config, parent_path, csv_bytes, csv_name)

    return ProcessResponse(
        status="ok" if not errors else "ok_con_errores",
        ruta_completa=ruta_completa,
        ruta_limpia=clean_path,
        nombre_archivo=nombre_archivo,
        carpeta_fecha=carpeta_fecha,
        csv_generado=csv_name,
        transacciones_exportadas=len(transactions),
        archivo_control=uploaded_url,
        errores=errors,
    )


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/v1/procesar-egresos", response_model=Union[ProcessResponse, IgnoredResponse])
def procesar_egresos(request: EgresosRequest) -> Union[ProcessResponse, IgnoredResponse]:
    try:
        if not should_process_excel_path(request.ruta_completa):
            print(f"Archivo ignorado para evitar bucle: {request.ruta_completa}")
            return IgnoredResponse(
                status="ignored",
                message="El archivo no es un Excel, se ignora para evitar bucles",
            )

        print(
            "Iniciando procesamiento por ruta de egreso: "
            f"{request.ruta_completa}"
        )
        return run_egresos_process(request.ruta_completa)
    except requests.HTTPError as http_error:
        graph_response = http_error.response
        status_code = graph_response.status_code if graph_response is not None else 500
        detail = graph_response.text if graph_response is not None else str(http_error)
        print(f"Error HTTP Graph: {detail}")
        raise HTTPException(status_code=status_code, detail=detail) from http_error
    except Exception as error:
        print(f"Error general procesando egresos: {error}")
        raise HTTPException(status_code=500, detail=str(error)) from error
