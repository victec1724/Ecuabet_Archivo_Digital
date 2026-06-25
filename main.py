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
DEFAULT_BATCH_DATE = "18-06-2026"
CONTROL_CSV_NAME = "egresos_a_procesar.csv"

app = FastAPI(
    title="Ecuabet RPA SharePoint Service",
    version="1.0.0",
    description="Procesa egresos desde SharePoint y genera el CSV de control para PAD.",
)


class ProcessRequest(BaseModel):
    fecha_lote: Optional[str] = DEFAULT_BATCH_DATE


class FileError(BaseModel):
    archivo: str
    error: str


class ProcessResponse(BaseModel):
    status: str
    fecha_lote: str
    archivos_encontrados: int
    archivos_procesados: int
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
            is_control_csv = file_name.casefold() == CONTROL_CSV_NAME.casefold()
            if (
                "file" in item
                and not is_control_csv
                and file_name.lower().endswith(SUPPORTED_EXTENSIONS)
            ):
                files.append(item)

        url = payload.get("@odata.nextLink")

    print(f"Archivos soportados encontrados: {len(files)}")
    return files


def download_file_bytes(access_token: str, drive_id: str, file_id: str) -> bytes:
    url = f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{file_id}/content"
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
    batch_date: str,
    egr: str,
    transactions: pd.DataFrame,
) -> None:
    drive_id = config["drive_id"]
    date_folder = ensure_folder(
        access_token,
        drive_id,
        config["destination_folder_id"],
        batch_date,
        "Nivel 1 Fecha",
    )
    egr_folder = ensure_folder(
        access_token,
        drive_id,
        date_folder["id"],
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


def upload_control_csv(
    access_token: str,
    config: dict[str, str],
    csv_bytes: bytes,
) -> str:
    drive_id = config["drive_id"]
    control_folder_id = config["rpa_control_folder_id"]
    url = (
        f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{control_folder_id}:/"
        f"{CONTROL_CSV_NAME}:/content"
    )
    response = requests.put(
        url,
        headers=build_headers(access_token, "text/csv"),
        data=csv_bytes,
        timeout=120,
    )
    response.raise_for_status()

    uploaded_file = response.json()
    web_url = uploaded_file.get("webUrl", CONTROL_CSV_NAME)
    print(f"CSV de control subido a SharePoint: {web_url}")
    return web_url


def process_file(
    access_token: str,
    config: dict[str, str],
    file_item: dict[str, Any],
    batch_date: str,
) -> pd.DataFrame:
    file_name = file_item["name"]
    file_id = file_item["id"]

    print(f"Procesando archivo: {file_name}")
    file_bytes = download_file_bytes(access_token, config["drive_id"], file_id)

    egr = extract_egr(file_bytes, file_name)
    print(f"EGR extraido: {egr}")

    transactions = extract_commissioner_transactions(file_bytes, file_name, egr)
    create_destination_hierarchy(access_token, config, batch_date, egr, transactions)
    return transactions


def run_egresos_process(batch_date: str) -> ProcessResponse:
    config = get_config()
    access_token = get_access_token(config)
    files = list_source_files(
        access_token,
        config["drive_id"],
        config["source_folder_id"],
    )

    processed_transactions: list[pd.DataFrame] = []
    errors: list[FileError] = []
    processed_files = 0

    for file_item in files:
        try:
            transactions = process_file(access_token, config, file_item, batch_date)
            processed_files += 1
            if not transactions.empty:
                processed_transactions.append(transactions)
        except Exception as file_error:
            file_name = file_item.get("name", "archivo desconocido")
            print(f"Error procesando {file_name}: {file_error}")
            errors.append(FileError(archivo=file_name, error=str(file_error)))

    csv_bytes = build_control_csv_bytes(processed_transactions)
    uploaded_url = upload_control_csv(access_token, config, csv_bytes)
    total_transactions = (
        sum(len(dataframe) for dataframe in processed_transactions)
        if processed_transactions
        else 0
    )

    return ProcessResponse(
        status="ok" if not errors else "ok_con_errores",
        fecha_lote=batch_date,
        archivos_encontrados=len(files),
        archivos_procesados=processed_files,
        transacciones_exportadas=total_transactions,
        archivo_control=uploaded_url,
        errores=errors,
    )


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/v1/procesar-egresos", response_model=ProcessResponse)
def procesar_egresos(request: ProcessRequest) -> ProcessResponse:
    try:
        batch_date = request.fecha_lote or DEFAULT_BATCH_DATE
        print(f"Iniciando procesamiento de egresos para lote: {batch_date}")
        return run_egresos_process(batch_date)
    except requests.HTTPError as http_error:
        graph_response = http_error.response
        status_code = graph_response.status_code if graph_response is not None else 500
        detail = graph_response.text if graph_response is not None else str(http_error)
        print(f"Error HTTP Graph: {detail}")
        raise HTTPException(status_code=status_code, detail=detail) from http_error
    except Exception as error:
        print(f"Error general procesando egresos: {error}")
        raise HTTPException(status_code=500, detail=str(error)) from error
