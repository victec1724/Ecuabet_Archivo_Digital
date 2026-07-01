import io
import os
from typing import Any, Optional
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


class ProcessRequest(BaseModel):
    archivo_id: str
    nombre_archivo: str


class FileError(BaseModel):
    archivo: str
    error: str


class ProcessResponse(BaseModel):
    status: str
    archivo_id: str
    nombre_archivo: str
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
        "source_folder_id": get_required_env("SOURCE_FOLDER_ID"),
        "destination_folder_id": get_required_env("DESTINATION_FOLDER_ID"),
        "rpa_control_folder_id": get_required_env("RPA_CONTROL_FOLDER_ID"),
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


def list_source_files(
    access_token: str,
    drive_id: str,
    source_folder_id: str,
) -> list[dict[str, Any]]:
    print("Listando archivos origen en SharePoint...")
    url = f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{source_folder_id}/children"
    files: list[dict[str, Any]] = []

    while url:
        response = requests.get(url, headers=build_headers(access_token), timeout=60)
        response.raise_for_status()

        payload = response.json()
        for item in payload.get("value", []):
            file_name = item.get("name", "")
            is_control_csv = file_name.casefold().startswith(CONTROL_CSV_PREFIX.casefold())
            if (
                "file" in item
                and not is_control_csv
                and file_name.lower().endswith(SUPPORTED_EXTENSIONS)
            ):
                files.append(item)

        url = payload.get("@odata.nextLink")

    print(f"Archivos soportados encontrados: {len(files)}")
    return files


def build_content_url(config: dict[str, str], file_id: str) -> str:
    drive_id = config["drive_id"]
    site_id = config.get("site_id")
    if site_id:
        return f"{GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/items/{file_id}/content"

    return f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{file_id}/content"


def download_file_bytes_by_id(
    access_token: str,
    config: dict[str, str],
    file_id: str,
) -> bytes:
    url = build_content_url(config, file_id)
    response = requests.get(url, headers=build_headers(access_token), timeout=120)
    response.raise_for_status()
    return response.content


def read_report_dataframe(
    file_bytes: bytes,
    file_name: str,
    *,
    nrows: Optional[int] = None,
    header: Optional[int] = 0,
    skiprows: Optional[int] = None,
) -> pd.DataFrame:
    stream = io.BytesIO(file_bytes)
    normalized_name = file_name.lower()

    if normalized_name.endswith(".csv"):
        return pd.read_csv(stream, nrows=nrows, header=header, skiprows=skiprows)

    if normalized_name.endswith(".xlsx"):
        return pd.read_excel(
            stream,
            nrows=nrows,
            header=header,
            skiprows=skiprows,
            engine="openpyxl",
        )

    if normalized_name.endswith(".xls"):
        try:
            return pd.read_excel(stream, nrows=nrows, header=header, skiprows=skiprows)
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
                skiprows=skiprows,
                on_bad_lines="skip",
            )

    raise ValueError(f"Extension no soportada para el archivo: {file_name}")


def extract_egr(file_bytes: bytes, file_name: str) -> str:
    dataframe = read_report_dataframe(file_bytes, file_name, nrows=2, header=None)
    if dataframe.shape[0] < 2 or dataframe.shape[1] < 1:
        raise ValueError(f"El archivo {file_name} no contiene la celda A2 esperada")

    raw_value = str(dataframe.iloc[1, 0])
    if ":" not in raw_value:
        raise ValueError(f"La celda A2 del archivo {file_name} no contiene ':'")

    egr = raw_value.split(":", maxsplit=1)[1].strip()
    if not egr:
        raise ValueError(f"No se encontro EGR valido en el archivo {file_name}")

    return egr


def extract_commissioner_transactions(
    file_bytes: bytes,
    file_name: str,
    egr: str,
) -> pd.DataFrame:
    dataframe = read_report_dataframe(file_bytes, file_name, skiprows=2)
    dataframe.columns = dataframe.columns.astype(str).str.strip()

    required_columns = {"CUENTA", "DETALLE"}
    missing_columns = required_columns - set(dataframe.columns)
    if missing_columns:
        missing_text = ", ".join(sorted(missing_columns))
        raise ValueError(f"Faltan columnas requeridas en {file_name}: {missing_text}")

    account_series = dataframe["CUENTA"].astype(str)
    filtered = dataframe[
        account_series.str.contains("Proveedores Locales", case=False, na=False)
    ].copy()

    if filtered.empty:
        print(f"No hay comisionistas validos en {file_name}")
        return pd.DataFrame(columns=PAD_COLUMNS)

    detail_complete = filtered["DETALLE"].fillna("").astype(str).str.strip()
    filtered["EGR"] = egr
    filtered["CODIGO_PROVEEDOR"] = detail_complete.str[:9].str.strip()
    filtered["DETALLE"] = detail_complete.str[9:].str.strip()
    filtered["DETALLE_COMPLETO"] = detail_complete

    transactions = filtered[
        (filtered["CODIGO_PROVEEDOR"] != "") & (filtered["DETALLE"] != "")
    ][PAD_COLUMNS].copy()

    print(f"{file_name}: transacciones validas filtradas: {len(transactions)}")
    return transactions


def sanitize_folder_name(folder_name: str) -> str:
    invalid_chars = '"*:<>?/\\|'
    sanitized_name = "".join("_" if char in invalid_chars else char for char in folder_name)
    sanitized_name = sanitized_name.strip().rstrip(".")
    if not sanitized_name:
        raise ValueError("El nombre de carpeta quedo vacio despues de sanitizar")

    return sanitized_name


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
    file_base_name: str,
    egr: str,
    transactions: pd.DataFrame,
) -> None:
    drive_id = config["drive_id"]
    file_folder = ensure_folder(
        access_token,
        drive_id,
        config["destination_folder_id"],
        file_base_name,
        "Nivel 1 Archivo",
    )
    egr_folder = ensure_folder(
        access_token,
        drive_id,
        file_folder["id"],
        egr,
        "Nivel 2 EGR",
    )

    for transaction in transactions.to_dict("records"):
        folder_name = f"FC {transaction['CODIGO_PROVEEDOR']} {transaction['DETALLE']}"
        ensure_folder(
            access_token,
            drive_id,
            egr_folder["id"],
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
    csv_bytes: bytes,
    csv_name: str,
) -> str:
    drive_id = config["drive_id"]
    control_folder_id = config["rpa_control_folder_id"]
    url = (
        f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{control_folder_id}:/"
        f"{csv_name}:/content"
    )
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
    archivo_id: str,
    nombre_archivo: str,
) -> pd.DataFrame:
    print(f"Procesando archivo: {nombre_archivo}")
    file_bytes = download_file_bytes_by_id(access_token, config, archivo_id)

    egr = extract_egr(file_bytes, nombre_archivo)
    print(f"EGR extraido: {egr}")

    transactions = extract_commissioner_transactions(file_bytes, nombre_archivo, egr)
    file_base_name = sanitize_folder_name(get_file_base_name(nombre_archivo))
    create_destination_hierarchy(access_token, config, file_base_name, egr, transactions)
    return transactions


def run_egresos_process(archivo_id: str, nombre_archivo: str) -> ProcessResponse:
    config = get_config()
    access_token = get_access_token(config)
    errors: list[FileError] = []
    transactions = pd.DataFrame(columns=PAD_COLUMNS)

    try:
        transactions = process_file(access_token, config, archivo_id, nombre_archivo)
    except requests.HTTPError:
        raise
    except Exception as file_error:
        print(f"Error procesando {nombre_archivo}: {file_error}")
        errors.append(FileError(archivo=nombre_archivo, error=str(file_error)))

    csv_name = build_control_csv_name(nombre_archivo)
    csv_bytes = build_control_csv_bytes([transactions] if not transactions.empty else [])
    uploaded_url = upload_control_csv(access_token, config, csv_bytes, csv_name)

    return ProcessResponse(
        status="ok" if not errors else "ok_con_errores",
        archivo_id=archivo_id,
        nombre_archivo=nombre_archivo,
        csv_generado=csv_name,
        transacciones_exportadas=len(transactions),
        archivo_control=uploaded_url,
        errores=errors,
    )


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/v1/procesar-egresos", response_model=ProcessResponse)
def procesar_egresos(request: ProcessRequest) -> ProcessResponse:
    try:
        print(
            "Iniciando procesamiento stateless de egreso: "
            f"{request.nombre_archivo} ({request.archivo_id})"
        )
        return run_egresos_process(request.archivo_id, request.nombre_archivo)
    except requests.HTTPError as http_error:
        graph_response = http_error.response
        status_code = graph_response.status_code if graph_response is not None else 500
        detail = graph_response.text if graph_response is not None else str(http_error)
        print(f"Error HTTP Graph: {detail}")
        raise HTTPException(status_code=status_code, detail=detail) from http_error
    except Exception as error:
        print(f"Error general procesando egresos: {error}")
        raise HTTPException(status_code=500, detail=str(error)) from error
