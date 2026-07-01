import os

from datetime import datetime
from typing import Any, Optional

import msal
import pandas as pd
import requests

from main import extract_commissioner_transactions as extract_transactions_intelligently


CLIENT_ID = os.getenv("GRAPH_CLIENT_ID")
CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET")
TENANT_ID = os.getenv("GRAPH_TENANT_ID")

DRIVE_ID = os.getenv("GRAPH_DRIVE_ID", "TU_DRIVE_ID")
SOURCE_FOLDER_ID = os.getenv("GRAPH_SOURCE_FOLDER_ID", "TU_SOURCE_FOLDER_ID")
DESTINATION_FOLDER_ID = os.getenv("GRAPH_DESTINATION_FOLDER_ID", "TU_DESTINATION_FOLDER_ID")
BATCH_DATE = os.getenv("GRAPH_BATCH_DATE") or datetime.now().strftime("%d-%m-%Y")
PAD_OUTPUT_CSV_PATH = r"C:\Temp\egresos_a_procesar.csv"

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
SUPPORTED_EXTENSIONS = (".xls", ".xlsx", ".csv")
PAD_COLUMNS = ["EGR", "CODIGO_PROVEEDOR", "DETALLE", "DETALLE_COMPLETO"]


def validate_credentials() -> None:
    missing_values = [
        variable_name
        for variable_name, variable_value in {
            "GRAPH_CLIENT_ID": CLIENT_ID,
            "GRAPH_CLIENT_SECRET": CLIENT_SECRET,
            "GRAPH_TENANT_ID": TENANT_ID,
        }.items()
        if not variable_value
    ]

    if missing_values:
        missing_text = ", ".join(missing_values)
        raise RuntimeError(f"Faltan variables de entorno requeridas: {missing_text}")


def get_access_token() -> str:
    """Obtiene un token OAuth 2.0 usando Client Credentials Flow."""
    validate_credentials()
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=authority,
    )

    result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in result:
        error = result.get("error_description") or result.get("error") or "Error desconocido"
        raise RuntimeError(f"No se pudo obtener el Access Token: {error}")

    return result["access_token"]


def build_headers(access_token: str) -> dict[str, str]:
    return {"Authorization": f"Bearer {access_token}"}


def list_source_files(access_token: str) -> list[dict[str, Any]]:
    """Lista archivos soportados dentro de SOURCE_FOLDER_ID."""
    headers = build_headers(access_token)
    url = f"{GRAPH_BASE_URL}/drives/{DRIVE_ID}/items/{SOURCE_FOLDER_ID}/children"
    files: list[dict[str, Any]] = []

    while url:
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()

        payload = response.json()
        for item in payload.get("value", []):
            file_name = item.get("name", "")
            is_file = "file" in item
            if is_file and file_name.lower().endswith(SUPPORTED_EXTENSIONS):
                files.append(item)

        url = payload.get("@odata.nextLink")

    return files


def download_file_bytes(access_token: str, file_id: str) -> bytes:
    """Descarga el contenido del archivo desde Graph sin escribir en disco."""
    headers = build_headers(access_token)
    url = f"{GRAPH_BASE_URL}/drives/{DRIVE_ID}/items/{file_id}/content"

    response = requests.get(url, headers=headers, timeout=120)
    response.raise_for_status()
    return response.content


def extract_code_from_file(file_bytes: bytes, file_name: str) -> str:
    """Obtiene el codigo EGR desde el nombre del archivo."""
    del file_bytes
    code = os.path.splitext(os.path.basename(file_name))[0].strip()
    if not code:
        raise ValueError(f"No se pudo obtener el codigo desde el nombre: {file_name}")

    return code


def sanitize_folder_name(folder_name: str) -> str:
    """Evita caracteres no permitidos por SharePoint en nombres de carpeta."""
    invalid_chars = '"*:<>?/\\|'
    sanitized_name = "".join("_" if char in invalid_chars else char for char in folder_name)
    sanitized_name = sanitized_name.strip().rstrip(".")
    if not sanitized_name:
        raise ValueError("El nombre de carpeta quedo vacio despues de sanitizar")

    return sanitized_name


def list_child_folders(access_token: str, parent_folder_id: str) -> list[dict[str, Any]]:
    headers = build_headers(access_token)
    url = f"{GRAPH_BASE_URL}/drives/{DRIVE_ID}/items/{parent_folder_id}/children"
    folders: list[dict[str, Any]] = []

    while url:
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()

        payload = response.json()
        folders.extend(item for item in payload.get("value", []) if "folder" in item)
        url = payload.get("@odata.nextLink")

    return folders


def find_child_folder(
    access_token: str, parent_folder_id: str, folder_name: str
) -> Optional[dict[str, Any]]:
    normalized_name = folder_name.casefold()
    for folder in list_child_folders(access_token, parent_folder_id):
        if folder.get("name", "").casefold() == normalized_name:
            return folder

    return None


def ensure_folder(access_token: str, parent_folder_id: str, folder_name: str, level_label: str) -> dict[str, Any]:
    """Obtiene una carpeta existente o la crea con conflictBehavior replace."""
    safe_folder_name = sanitize_folder_name(folder_name)
    print(f"Validando carpeta {level_label}: {safe_folder_name}")

    existing_folder = find_child_folder(access_token, parent_folder_id, safe_folder_name)
    if existing_folder:
        print(f"Carpeta {level_label} existente: {existing_folder['name']}")
        return existing_folder

    print(f"Creando carpeta {level_label}: {safe_folder_name}")
    headers = {
        **build_headers(access_token),
        "Content-Type": "application/json",
    }
    url = f"{GRAPH_BASE_URL}/drives/{DRIVE_ID}/items/{parent_folder_id}/children"
    payload = {
        "name": safe_folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "replace",
    }

    response = requests.post(url, headers=headers, json=payload, timeout=60)
    if response.status_code == 409:
        existing_folder = find_child_folder(access_token, parent_folder_id, safe_folder_name)
        if existing_folder:
            print(f"Carpeta {level_label} detectada tras conflicto: {existing_folder['name']}")
            return existing_folder

    response.raise_for_status()
    created_folder = response.json()
    print(f"Carpeta {level_label} creada: {created_folder.get('name', safe_folder_name)}")
    return created_folder


def create_destination_hierarchy(
    access_token: str,
    batch_date: str,
    egr_code: str,
    transactions: pd.DataFrame,
) -> None:
    date_folder = ensure_folder(
        access_token,
        DESTINATION_FOLDER_ID,
        batch_date,
        "Nivel 1 Fecha",
    )
    egr_folder = ensure_folder(
        access_token,
        date_folder["id"],
        egr_code,
        "Nivel 2 Egreso",
    )

    for transaction in transactions.to_dict("records"):
        folder_name = (
            f"FC {transaction['CODIGO_PROVEEDOR']} {transaction['DETALLE']}"
        )
        ensure_folder(
            access_token,
            egr_folder["id"],
            folder_name,
            "Nivel 3 Factura del Comisionista",
        )


def export_pad_exchange_csv(transactions: list[pd.DataFrame]) -> None:
    os.makedirs(os.path.dirname(PAD_OUTPUT_CSV_PATH), exist_ok=True)

    if transactions:
        consolidated = pd.concat(transactions, ignore_index=True)
    else:
        consolidated = pd.DataFrame(columns=PAD_COLUMNS)

    consolidated = consolidated[PAD_COLUMNS]
    consolidated.to_csv(PAD_OUTPUT_CSV_PATH, index=False, encoding="utf-8-sig")
    print(f"Archivo de intercambio PAD generado: {PAD_OUTPUT_CSV_PATH}")
    print(f"Total transacciones exportadas: {len(consolidated)}")


def process_file(access_token: str, file_item: dict[str, Any], batch_date: str) -> pd.DataFrame:
    file_name = file_item["name"]
    file_id = file_item["id"]

    print(f"Procesando archivo {file_name}...")
    file_bytes = download_file_bytes(access_token, file_id)

    code = extract_code_from_file(file_bytes, file_name)
    print(f"Codigo {code} extraido")

    transactions = extract_transactions_intelligently(file_bytes, file_name, code)
    create_destination_hierarchy(access_token, batch_date, code, transactions)
    return transactions


def main() -> None:
    try:
        print("Iniciando autenticacion con Microsoft Graph...")
        access_token = get_access_token()
        print("Autenticacion completada")
        print(f"Fecha del lote de procesamiento: {BATCH_DATE}")

        print("Listando archivos del directorio origen...")
        files = list_source_files(access_token)
        print(f"Archivos encontrados: {len(files)}")

        processed_transactions: list[pd.DataFrame] = []
        for file_item in files:
            try:
                transactions = process_file(access_token, file_item, BATCH_DATE)
                if not transactions.empty:
                    processed_transactions.append(transactions)
            except Exception as file_error:
                file_name = file_item.get("name", "archivo desconocido")
                print(f"Error procesando {file_name}: {file_error}")

        export_pad_exchange_csv(processed_transactions)
        print("Proceso finalizado")
    except Exception as error:
        print(f"Error general del orquestador: {error}")
        raise


if __name__ == "__main__":
    main()
