"""
matcher.py — Fase 1 Core: Normalizacion y emparejamiento difuso de nombres.

Responsabilidades:
- Normalizar nombres de archivos origen (quitar prefijos/tokens institucionales
  en cualquier posicion, extensiones no estandar, acentos, espacios extra)
  a una forma canonica lowercase-sin-espacios.
- Emparejar (fuzzy match) cada nombre de archivo con la lista de pestanas del
  libro destino, devolviendo la mejor coincidencia y su puntuacion de confianza.
"""

from __future__ import annotations

import os
import re
import unicodedata
from typing import Optional

from rapidfuzz import process, fuzz


# ---------------------------------------------------------------------------
# Tokens institucionales que se eliminan donde quiera que aparezcan
# Orden importa: los mas largos van primero para evitar coincidencias parciales
# ---------------------------------------------------------------------------
_INSTITUTIONAL_TOKENS: list[str] = [
    r"us-[a-z]",          # US-B, US-C, US-A (sin espacio posterior)
    r"us\b",
    r"ucsf\b",
    r"cs\b",
    r"\bsm\b",            # SM como palabra completa
    r"\bcentro\s+de\s+salud\b",
    r"\bunidad\s+de\s+salud\b",
    r"\bcarolina\b",      # nombre de municipio padre, no de la unidad destino
    r"\bchirilagua\b",    # nombre de municipio padre
]

# Patron global (aplica en cualquier posicion del string)
_TOKENS_RE = re.compile(
    r"(?:" + "|".join(_INSTITUTIONAL_TOKENS) + r")",
    flags=re.IGNORECASE,
)

# Fragmentos de anio en el nombre (ej: xls2026, .xls2026., 2026)
_YEAR_RE = re.compile(r"(?:\.xls)?\d{4}", flags=re.IGNORECASE)


def normalize_name(raw: str) -> str:
    """
    Convierte un nombre de archivo (o pestana) a forma canonica para comparar.

    Pasos aplicados:
    1. Extraer solo el nombre base (sin ruta).
    2. Eliminar extension(es): quita todo desde el ultimo punto .xlsx/.xls, pero
       tambien fragmentos embebidos como xls2026 antes del .xlsx.
    3. Quitar acentos y caracteres diacriticos -> ASCII puro.
    4. Eliminar tokens institucionales globalmente (US-B, SM, Carolina, etc.).
    5. Eliminar caracteres no alfanumericos (guiones, puntos, parentesis...).
    6. Colapsar espacios y pasar a minusculas sin espacios internos.

    Parameters
    ----------
    raw : str
        Nombre de archivo completo (puede incluir ruta) o nombre de pestana.

    Returns
    -------
    str
        Cadena normalizada, ej: 'laceibita', 'sanpedro', 'elcuco'.

    Examples
    --------
    >>> normalize_name("US-B Carolina SM La Ceibita.xlsx")
    'laceibita'
    >>> normalize_name("EL CUCO.xlsx")
    'elcuco'
    >>> normalize_name("US-B Chirilagua SM San Pedro.xls2026.xlsx")
    'sanpedro'
    >>> normalize_name("LaCeibita")
    'laceibita'
    """
    # 1. Solo nombre base sin ruta
    stem = os.path.basename(raw)

    # 2. Quitar extension(es) — incluyendo compuestas como .xls2026.xlsx
    #    Primero quitamos el .xlsx / .xls final, luego xls2026
    stem = re.sub(r"\.xlsx?$", "", stem, flags=re.IGNORECASE)
    stem = _YEAR_RE.sub("", stem)

    # 3. Quitar acentos: NFD -> filtrar combining marks -> ASCII
    nfd = unicodedata.normalize("NFD", stem)
    stem = "".join(ch for ch in nfd if unicodedata.category(ch) != "Mn")

    # 4. Eliminar tokens institucionales (en cualquier posicion)
    stem = _TOKENS_RE.sub(" ", stem)

    # 5. Quitar caracteres que no sean letras/digitos/espacios
    stem = re.sub(r"[^a-zA-Z0-9 ]", " ", stem)

    # 6. Colapsar espacios, minusculas, sin espacios internos
    stem = re.sub(r"\s+", "", stem).lower()

    return stem


def match_file_to_sheet(
    filename: str,
    sheet_names: list,
    score_cutoff: float = 60.0,
):
    """
    Encuentra la pestana destino que mejor coincide con el nombre del archivo.

    Utiliza rapidfuzz.fuzz.token_sort_ratio para tolerar diferencias en el
    orden de palabras.

    Parameters
    ----------
    filename : str
        Nombre del archivo origen.
    sheet_names : list[str]
        Lista de nombres de pestanas del libro destino.
    score_cutoff : float
        Umbral minimo de similitud (0-100).

    Returns
    -------
    tuple[str | None, float]
        (nombre_pestana_ganadora, puntuacion) o (None, 0.0) si no coincide.
    """
    norm_filename = normalize_name(filename)

    # Mapa: nombre_normalizado -> nombre_original
    norm_to_original = {normalize_name(s): s for s in sheet_names}

    result = process.extractOne(
        query=norm_filename,
        choices=list(norm_to_original.keys()),
        scorer=fuzz.token_sort_ratio,
        score_cutoff=score_cutoff,
    )

    if result is None:
        return None, 0.0

    matched_norm, score, _ = result
    return norm_to_original[matched_norm], float(score)


def build_mapping(
    source_filenames: list,
    sheet_names: list,
    score_cutoff: float = 60.0,
) -> list:
    """
    Construye el mapeo completo entre archivos origen y pestanas destino.

    Parameters
    ----------
    source_filenames : list[str]
        Lista de nombres de archivo de la carpeta origen.
    sheet_names : list[str]
        Nombres de pestanas del libro destino.
    score_cutoff : float
        Umbral minimo de similitud.

    Returns
    -------
    list[dict]
        Lista de diccionarios con claves:
        - file       : nombre de archivo original
        - norm_file  : forma normalizada del archivo
        - sheet      : pestana destino emparejada (o None)
        - norm_sheet : forma normalizada de la pestana
        - score      : puntuacion de similitud (0-100)
        - matched    : True si se encontro coincidencia
    """
    mapping = []
    for fname in source_filenames:
        sheet, score = match_file_to_sheet(fname, sheet_names, score_cutoff)
        mapping.append(
            {
                "file": fname,
                "norm_file": normalize_name(fname),
                "sheet": sheet,
                "norm_sheet": normalize_name(sheet) if sheet else None,
                "score": score,
                "matched": sheet is not None,
            }
        )
    return mapping
