import os
from typing import Any, Optional
from urllib.parse import parse_qs, unquote, urlparse

import msal
import requests


CLIENT_ID = os.getenv("GRAPH_CLIENT_ID")
CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET")
TENANT_ID = os.getenv("GRAPH_TENANT_ID")

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


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


def get_user_drive_id(access_token: str, user_principal_name: str) -> str:
    """Consulta el OneDrive de un usuario e imprime su DRIVE_ID."""
    headers = build_headers(access_token)
    url = f"{GRAPH_BASE_URL}/users/{user_principal_name}/drive"

    response = requests.get(url, headers=headers, timeout=60)
    response.raise_for_status()

    drive = response.json()
    drive_id = drive["id"]

    print("\nDRIVE encontrado")
    print(f"Usuario: {user_principal_name}")
    print(f"Nombre : {drive.get('name', 'Sin nombre')}")
    print(f"DRIVE_ID: {drive_id}\n")

    return drive_id


def parse_sharepoint_url(sharepoint_url: str) -> tuple[str, str, Optional[str]]:
    """Extrae hostname, ruta del sitio y biblioteca desde una URL de SharePoint."""
    parsed_url = urlparse(sharepoint_url)
    query_params = parse_qs(parsed_url.query)
    list_url = query_params.get("listurl", [sharepoint_url])[0]

    parsed_list_url = urlparse(unquote(list_url))
    hostname = parsed_list_url.netloc
    path_parts = [part for part in parsed_list_url.path.split("/") if part]

    if len(path_parts) < 2 or path_parts[0].lower() != "sites":
        raise ValueError(
            "No se pudo identificar la ruta del sitio. "
            "Use una URL tipo https://tenant.sharepoint.com/sites/NOMBRE_SITIO"
        )

    site_path = f"/sites/{path_parts[1]}"
    library_name = unquote(path_parts[2]) if len(path_parts) > 2 else None

    return hostname, site_path, library_name


def get_sharepoint_site_id(access_token: str, sharepoint_url: str) -> tuple[str, Optional[str]]:
    """Consulta Graph para obtener el SITE_ID desde una URL de SharePoint."""
    hostname, site_path, library_name = parse_sharepoint_url(sharepoint_url)
    headers = build_headers(access_token)
    url = f"{GRAPH_BASE_URL}/sites/{hostname}:{site_path}"

    response = requests.get(url, headers=headers, timeout=60)
    response.raise_for_status()

    site = response.json()
    site_id = site["id"]

    print("\nSitio de SharePoint encontrado")
    print(f"Nombre : {site.get('displayName', 'Sin nombre')}")
    print(f"Web URL: {site.get('webUrl', 'Sin URL')}")
    print(f"SITE_ID: {site_id}\n")

    return site_id, library_name


def list_sharepoint_drives(access_token: str, site_id: str) -> list[dict[str, Any]]:
    """Lista las bibliotecas de documentos de un sitio SharePoint."""
    headers = build_headers(access_token)
    url = f"{GRAPH_BASE_URL}/sites/{site_id}/drives"
    drives: list[dict[str, Any]] = []

    while url:
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()

        payload = response.json()
        drives.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")

    print_drives_table(drives)
    return drives


def find_drive_by_name(
    drives: list[dict[str, Any]], library_name: Optional[str]
) -> Optional[dict[str, Any]]:
    if not library_name:
        return None

    normalized_library_name = library_name.lower()
    for drive in drives:
        drive_name = drive.get("name", "").lower()
        if (
            drive_name == normalized_library_name
            or drive_name in normalized_library_name
            or normalized_library_name in drive_name
        ):
            return drive

    return None


def print_drives_table(drives: list[dict[str, Any]]) -> None:
    """Muestra una tabla simple con bibliotecas y DRIVE_ID."""
    if not drives:
        print("No se encontraron bibliotecas de documentos en este sitio.")
        return

    name_width = min(max(len(drive.get("name", "")) for drive in drives), 60)

    print(f"{'Biblioteca':<{name_width}}  DRIVE_ID")
    print(f"{'-' * name_width}  {'-' * 80}")

    for drive in drives:
        drive_name = drive.get("name", "Sin nombre")[:name_width]
        drive_id = drive.get("id", "Sin ID")
        print(f"{drive_name:<{name_width}}  {drive_id}")


def list_root_children(access_token: str, drive_id: str) -> list[dict[str, Any]]:
    """Lista los elementos de la raiz del drive."""
    url = f"{GRAPH_BASE_URL}/drives/{drive_id}/root/children"
    return list_folder_children_by_url(access_token, url)


def list_folder_children(access_token: str, drive_id: str, folder_id: str) -> list[dict[str, Any]]:
    """Lista los elementos dentro de una carpeta especifica."""
    url = f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{folder_id}/children"
    return list_folder_children_by_url(access_token, url)


def list_folder_children_by_url(access_token: str, url: str) -> list[dict[str, Any]]:
    """Ejecuta consultas paginadas para obtener hijos de una carpeta."""
    headers = build_headers(access_token)
    children: list[dict[str, Any]] = []

    while url:
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()

        payload = response.json()
        children.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")

    print_children_table(children)
    return children


def print_children_table(children: list[dict[str, Any]]) -> None:
    """Muestra una tabla simple con nombre, tipo e ID."""
    if not children:
        print("No se encontraron elementos en esta ubicacion.")
        return

    name_width = min(max(len(item.get("name", "")) for item in children), 60)
    type_width = 10

    print(f"{'Tipo':<{type_width}}  {'Nombre':<{name_width}}  ID")
    print(f"{'-' * type_width}  {'-' * name_width}  {'-' * 80}")

    for item in children:
        item_type = "Carpeta" if "folder" in item else "Archivo"
        item_name = item.get("name", "Sin nombre")
        item_id = item.get("id", "Sin ID")
        trimmed_name = item_name[:name_width]
        print(f"{item_type:<{type_width}}  {trimmed_name:<{name_width}}  {item_id}")


def navigate_drive(access_token: str, drive_id: str) -> None:
    print("\nElementos en la raiz del drive:")
    list_root_children(access_token, drive_id)

    while True:
        folder_id = input(
            "\nIngrese un FOLDER_ID para navegar o presione Enter para salir: "
        ).strip()
        if not folder_id:
            print("Consulta finalizada.")
            break

        print(f"\nElementos dentro de la carpeta {folder_id}:")
        list_folder_children(access_token, drive_id, folder_id)


def run_onedrive_flow(access_token: str) -> None:
    user_principal_name = input("Correo/User Principal Name del usuario: ").strip()
    if not user_principal_name:
        raise ValueError("Debe ingresar un correo valido.")

    drive_id = get_user_drive_id(access_token, user_principal_name)
    navigate_drive(access_token, drive_id)


def run_sharepoint_flow(access_token: str) -> None:
    sharepoint_url = input("URL del sitio o enlace compartido de SharePoint: ").strip()
    if not sharepoint_url:
        raise ValueError("Debe ingresar una URL de SharePoint valida.")

    site_id, library_name = get_sharepoint_site_id(access_token, sharepoint_url)

    print("Bibliotecas de documentos del sitio:")
    drives = list_sharepoint_drives(access_token, site_id)

    detected_drive = find_drive_by_name(drives, library_name)
    if detected_drive:
        drive_id = detected_drive["id"]
        print(f"\nBiblioteca detectada desde la URL: {detected_drive.get('name')}")
        print(f"DRIVE_ID: {drive_id}")
    else:
        drive_id = input("\nIngrese el DRIVE_ID de la biblioteca a explorar: ").strip()
        if not drive_id:
            raise ValueError("Debe ingresar un DRIVE_ID valido.")

    navigate_drive(access_token, drive_id)


def main() -> None:
    try:
        print("Autenticando con Microsoft Graph...")
        access_token = get_access_token()

        print("\nSeleccione el origen:")
        print("1. OneDrive de usuario")
        print("2. Sitio de SharePoint")
        option = input("Opcion: ").strip()

        if option == "1":
            run_onedrive_flow(access_token)
        elif option == "2":
            run_sharepoint_flow(access_token)
        else:
            raise ValueError("Opcion invalida. Use 1 o 2.")
    except Exception as error:
        print(f"Error en graph_id_finder: {error}")
        raise


if __name__ == "__main__":
    main()
