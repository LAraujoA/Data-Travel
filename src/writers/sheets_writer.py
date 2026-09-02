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


# ---------------------------------------------------------------------------
#  get_sheet_tabs — lista de pestanas de un Spreadsheet
# ---------------------------------------------------------------------------

def get_sheet_tabs(
    spreadsheet_id_or_url: str,
    credentials_path: str = "credentials.json",
) -> list[str]:
    """
    Retorna la lista de nombres de pestanas del Spreadsheet.

    Parameters
    ----------
    spreadsheet_id_or_url : str
        URL completa o ID del spreadsheet.
    credentials_path : str
        Ruta al JSON de service account.

    Returns
    -------
    list[str]
        Nombres de las hojas disponibles.

    Raises
    ------
    ImportError   : gspread / google-auth no instalados.
    FileNotFoundError : credentials.json no encontrado.
    gspread.exceptions.APIError : credenciales invalidas o sin permiso.
    """
    if not _GSPREAD_AVAILABLE:
        raise ImportError(
            "gspread y google-auth son necesarios.\n"
            "Instala con: pip install gspread google-auth"
        )
    from pathlib import Path as _Path
    creds_p = _Path(credentials_path)
    if not creds_p.exists():
        raise FileNotFoundError(
            f"Credenciales no encontradas: {creds_p.resolve()}\n"
            "Descarga el JSON de service account desde Google Cloud Console."
        )
    try:
        creds = SACredentials.from_service_account_file(
            credentials_path, scopes=_SHEETS_SCOPES
        )
        gc = gspread.authorize(creds)
        sid = _extract_spreadsheet_id(spreadsheet_id_or_url)
        sh  = gc.open_by_key(sid)
        return [ws.title for ws in sh.worksheets()]
    except gspread.exceptions.SpreadsheetNotFound:
        raise ValueError(
            f"Spreadsheet no encontrado. Verifica el ID/URL y que la hoja "
            f"este compartida con el service account."
        )
    except Exception as exc:
        log.error("get_sheet_tabs error: %s", exc)
        raise


# ---------------------------------------------------------------------------
#  write_range_to_sheets — escribe lista plana de valores via batch_update
# ---------------------------------------------------------------------------

def write_range_to_sheets(
    spreadsheet_id_or_url: str,
    sheet_name: str,
    values: list,
    mode: str,
    credentials_path: str = "credentials.json",
    start_cell: str = "A1",
    direction: str = "Vertical",
    stride: int = 1,
    cell_list_str: str = "",
) -> dict:
    """
    Escribe una lista plana de valores en una hoja de Google Sheets.

    Soporta las mismas 3 modalidades que los writers de Excel:
      - "Bloque Continuo" : valores en columna a partir de start_cell.
      - "Salto"           : distribuye con salto N horizontal o vertical.
      - "Lista de Celdas" : asigna cada valor a la celda explicita (puede
                            contener rangos: "C20:C22, G20:G22").

    Usa una sola llamada batch_update para minimizar el consumo de cuota
    y evitar errores 429 (rate limit).

    Parameters
    ----------
    spreadsheet_id_or_url : str
        URL o ID del spreadsheet.
    sheet_name : str
        Nombre de la pestana destino.
    values : list
        Lista plana de valores a escribir.
    mode : str
        Una de: "Bloque Continuo", "Salto", "Lista de Celdas".
    credentials_path : str
        Ruta al JSON de service account.
    start_cell : str
        Celda de inicio para modos Bloque y Salto. Ej: "C3".
    direction : str
        "Horizontal" o "Vertical" (solo modo Salto).
    stride : int
        Salto entre celdas (solo modo Salto).
    cell_list_str : str
        Expresion de celdas/rangos destino (solo modo Lista).
        Ej: "C20:C22, G20:G22, K20:K22".

    Returns
    -------
    dict
        {
          "written" : int,        # cantidad de valores escritos
          "cells"   : list[str],  # coordenadas A1 destino
          "detail"  : list[str],  # ["C3=4", "C4=12", ...]
        }

    Raises
    ------
    ImportError      : gspread / google-auth no instalados.
    FileNotFoundError: credenciales no encontradas.
    ValueError       : hoja no encontrada o celda invalida.
    gspread.exceptions.APIError : error de cuota o autenticacion.
    """
    if not _GSPREAD_AVAILABLE:
        raise ImportError(
            "gspread y google-auth son necesarios.\n"
            "Instala con: pip install gspread google-auth"
        )
    from pathlib import Path as _Path
    from openpyxl.utils import column_index_from_string, get_column_letter

    creds_p = _Path(credentials_path)
    if not creds_p.exists():
        raise FileNotFoundError(
            f"Credenciales no encontradas: {creds_p.resolve()}"
        )

    # --- Autenticar y abrir hoja -------------------------------------------
    try:
        creds = SACredentials.from_service_account_file(
            credentials_path, scopes=_SHEETS_SCOPES
        )
        gc = gspread.authorize(creds)
        sid = _extract_spreadsheet_id(spreadsheet_id_or_url)
        sh  = gc.open_by_key(sid)
    except gspread.exceptions.SpreadsheetNotFound:
        raise ValueError(
            "Spreadsheet no encontrado. Verifica el ID/URL y los permisos."
        )
    except Exception as exc:
        log.error("Error autenticando con Google Sheets: %s", exc)
        raise

    try:
        ws = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        available = [w.title for w in sh.worksheets()]
        raise ValueError(
            f"Pestana '{sheet_name}' no encontrada. "
            f"Disponibles: {available}"
        )

    # --- Calcular coordenadas destino --------------------------------------
    def _parse_a1(cell: str):
        """Retorna (row_1based, col_1based) de una celda A1."""
        import re
        cell = cell.strip().upper()
        m = re.match(r"([A-Z]+)(\d+)", cell)
        if not m:
            raise ValueError(f"Celda invalida: '{cell}'")
        col = column_index_from_string(m.group(1))
        row = int(m.group(2))
        return row, col

    def _to_a1(row: int, col: int) -> str:
        return f"{get_column_letter(col)}{row}"

    def _expand_tokens(expr: str) -> list[str]:
        """Expande expresion mixta de rangos a lista de coordenadas A1."""
        import re
        tokens = [t.strip().upper() for t in expr.split(",") if t.strip()]
        cells = []
        for tok in tokens:
            if ":" in tok:
                a, b = tok.split(":", 1)
                r1, c1 = _parse_a1(a)
                r2, c2 = _parse_a1(b)
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        cells.append(_to_a1(r, c))
            else:
                _parse_a1(tok)  # valida
                cells.append(tok)
        return cells

    dest_coords: list[str] = []

    if mode == "Bloque Continuo":
        r0, c0 = _parse_a1(start_cell)
        for i in range(len(values)):
            dest_coords.append(_to_a1(r0 + i, c0))

    elif mode == "Salto":
        r0, c0 = _parse_a1(start_cell)
        for i in range(len(values)):
            if direction == "Horizontal":
                dest_coords.append(_to_a1(r0, c0 + i * stride))
            else:
                dest_coords.append(_to_a1(r0 + i * stride, c0))

    else:  # Lista de Celdas
        dest_coords = _expand_tokens(cell_list_str)

    # Recortar si hay mas celdas que valores (o viceversa)
    n = min(len(values), len(dest_coords))
    values    = values[:n]
    dest_coords = dest_coords[:n]

    if not values:
        return {"written": 0, "cells": [], "detail": []}

    # --- Construir batch y enviar ------------------------------------------
    updates = [
        {"range": coord, "values": [[val]]}
        for coord, val in zip(dest_coords, values)
    ]

    try:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
    except Exception as exc:
        # Manejo especifico de error 429 (quota)
        err_str = str(exc)
        if "429" in err_str or "Quota" in err_str or "quota" in err_str:
            raise RuntimeError(
                "Limite de cuota de Google Sheets API (error 429).\n"
                "Espera unos segundos e intenta de nuevo, o reduce la frecuencia "
                "de solicitudes."
            ) from exc
        log.error("batch_update error: %s", exc)
        raise

    detail = [f"{c}={v}" for c, v in zip(dest_coords, values)]
    log.info(
        "Sheets batch_update | Hoja: '%s' | Celdas: %d | Modo: %s",
        sheet_name, n, mode,
    )
    return {
        "written": n,
        "cells":   dest_coords,
        "detail":  detail,
    }
