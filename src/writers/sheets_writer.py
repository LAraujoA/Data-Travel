"""
sheets_writer.py — Fase 2: Escritura de datos mensuales en Google Sheets.

Responsabilidad:
- Recibir un diccionario {numero_indicador: valor_real} y volcarlo via la
  API de Google Sheets (gspread) en la columna REAL del mes indicado.
- Usar batch_update para minimizar llamadas a la API y no consumir cuota
  innecesariamente.
- El archivo de credenciales de service account por defecto es credentials.json
  en la raiz del proyecto; tambien acepta credenciales OAuth2 del usuario.

Prerequisitos:
  pip install gspread google-auth
  Compartir la hoja con el email del service account.

Nota de diseno:
  Este modulo es OPCIONAL en la Fase 2 de pruebas locales.
  La prueba de integracion (test_phase2.py) solo valida excel_writer.
  sheets_writer se activa cuando se provee un spreadsheet_id/url valido.
"""

from __future__ import annotations

import logging
import re
import unicodedata
from typing import Union

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Importacion diferida de gspread para no romper si no esta instalado
# ---------------------------------------------------------------------------
try:
    import gspread
    from google.oauth2.service_account import Credentials as SACredentials
    from google.oauth2.credentials import Credentials as OAuthCredentials
    _GSPREAD_AVAILABLE = True
except ImportError:
    _GSPREAD_AVAILABLE = False


_SHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Extrae el spreadsheet ID de una URL completa o lo devuelve tal cual
_SHEET_ID_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9_-]+)")


def _extract_spreadsheet_id(id_or_url: str) -> str:
    """Extrae el ID de una URL de Google Sheets o retorna el string tal cual."""
    match = _SHEET_ID_RE.search(id_or_url)
    return match.group(1) if match else id_or_url.strip()


def _strip_accents(text: str) -> str:
    nfd = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in nfd if unicodedata.category(ch) != "Mn")


def _col_index_to_a1(col: int) -> str:
    """Convierte indice de columna 1-based a notacion A1 (A, B, ..., Z, AA, ...)."""
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _find_real_column_sheets(worksheet, target_month: str) -> int:
    """
    Localiza la columna REAL para el mes dado en una hoja de Google Sheets.

    Usa los valores de las filas 4 y 5 (indices 3 y 4 en 0-based de gspread
    get_all_values) para localizar la columna, igual que en excel_writer.

    Parameters
    ----------
    worksheet : gspread.Worksheet
        Hoja de la hoja de calculo.
    target_month : str
        Nombre del mes objetivo.

    Returns
    -------
    int
        Indice de columna 1-based.
    """
    target = _strip_accents(target_month.strip().upper())
    all_rows = worksheet.get_all_values()

    if len(all_rows) < 5:
        raise ValueError("La hoja no tiene suficientes filas (se esperan al menos 5).")

    row4 = all_rows[3]  # fila 4 (0-indexed: 3)
    row5 = all_rows[4]  # fila 5 (0-indexed: 4)

    month_col_start: int | None = None
    for idx, val in enumerate(row4):
        if val and _strip_accents(val.strip().upper()) == target:
            month_col_start = idx  # 0-based
            break

    if month_col_start is None:
        raise ValueError(f"Mes '{target_month}' no encontrado en la fila 4.")

    real_col_0: int | None = None
    for idx in range(month_col_start, len(row5)):
        # Nuevo mes -> salir
        if idx > month_col_start and row4[idx]:
            break
        if row5[idx].strip().upper() == "REAL":
            real_col_0 = idx
            break

    if real_col_0 is None:
        raise ValueError(f"Sub-columna 'REAL' no encontrada para el mes '{target_month}'.")

    return real_col_0 + 1  # convertir a 1-based


def write_month_data_to_sheets(
    spreadsheet_id_or_url: str,
    sheet_name: str,
    month: str,
    data_dict: dict,
    credentials_path: str = "credentials.json",
    data_start_row: int = 6,
    indicator_col: int = 1,
) -> bool:
    """
    Escribe los valores REAL del diccionario en una hoja de Google Sheets.

    Usa gspread con service account credentials. Aplica batch_update para
    minimizar el consumo de cuota de la API.

    Parameters
    ----------
    spreadsheet_id_or_url : str
        ID del spreadsheet o URL completa de Google Sheets.
    sheet_name : str
        Nombre de la pestana destino.
    month : str
        Nombre del mes objetivo, ej: ``"ENERO"``.
    data_dict : dict
        Diccionario ``{int_indicador: valor_real}``.
    credentials_path : str
        Ruta al JSON de credenciales de service account. Default: "credentials.json".
    data_start_row : int
        Primera fila de datos (1-based). Default: 6.
    indicator_col : int
        Columna del indicador (1-based). Default: 1.

    Returns
    -------
    bool
        True si la escritura fue exitosa, False en caso de error.

    Raises
    ------
    ImportError
        Si gspread o google-auth no estan instalados.
    FileNotFoundError
        Si el archivo de credenciales no existe.
    gspread.exceptions.WorksheetNotFound
        Si la pestana no existe en el spreadsheet.
    """
    if not _GSPREAD_AVAILABLE:
        raise ImportError(
            "gspread y google-auth son necesarios para usar sheets_writer.\n"
            "Instalalos con: pip install gspread google-auth"
        )

    import os
    creds_path_obj = __import__("pathlib").Path(credentials_path)
    if not creds_path_obj.exists():
        raise FileNotFoundError(
            f"Archivo de credenciales no encontrado: {creds_path_obj.resolve()}\n"
            "Descarga el JSON de tu service account desde Google Cloud Console."
        )

    try:
        creds = SACredentials.from_service_account_file(
            credentials_path, scopes=_SHEETS_SCOPES
        )
        gc = gspread.authorize(creds)
    except Exception as exc:
        log.error("Error autenticando con Google Sheets: %s", exc)
        raise

    spreadsheet_id = _extract_spreadsheet_id(spreadsheet_id_or_url)

    try:
        sh = gc.open_by_key(spreadsheet_id)
    except gspread.exceptions.SpreadsheetNotFound:
        log.error("Spreadsheet no encontrado: %s", spreadsheet_id)
        raise

    try:
        ws = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        available = [w.title for w in sh.worksheets()]
        raise gspread.exceptions.WorksheetNotFound(
            f"Pestana '{sheet_name}' no encontrada. "
            f"Disponibles: {available}"
        )

    # Localizar columna REAL
    real_col_1based = _find_real_column_sheets(ws, month)
    real_col_letter = _col_index_to_a1(real_col_1based)
    log.debug("Columna REAL para '%s': %s (col %d)", month, real_col_letter, real_col_1based)

    # Leer columna de indicadores para mapear indicador -> fila
    ind_col_letter = _col_index_to_a1(indicator_col)
    ind_values = ws.col_values(indicator_col)  # lista 1-indexed desde fila 1

    # Construir batch de actualizaciones
    updates: list[dict] = []
    not_found: list = list(data_dict.keys())

    for row_0, cell_val in enumerate(ind_values):
        row_1based = row_0 + 1
        if row_1based < data_start_row:
            continue
        if not cell_val:
            continue
        try:
            ind_key = int(str(cell_val).strip())
        except (TypeError, ValueError):
            continue

        if ind_key in data_dict:
            cell_a1 = f"{real_col_letter}{row_1based}"
            updates.append({
                "range": cell_a1,
                "values": [[data_dict[ind_key]]],
            })
            if ind_key in not_found:
                not_found.remove(ind_key)

    if not updates:
        log.warning("No se generaron actualizaciones para la pestana '%s'.", sheet_name)
        return False

    # Batch update — una sola llamada a la API
    try:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        log.info(
            "Google Sheets actualizado | Pestana: '%s' | Mes: '%s' | "
            "Celdas escritas: %d | Indicadores no encontrados en hoja: %s",
            sheet_name, month, len(updates), not_found,
        )
        return True
    except Exception as exc:
        log.error("Error en batch_update: %s", exc)
        raise
