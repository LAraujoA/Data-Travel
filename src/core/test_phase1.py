"""
test_phase1.py — Script de prueba para la Fase 1 del sistema de migracion.

Ejecuta el flujo completo:
  1. Lee los archivos de DATATEST/origen_unidades/
  2. Lee las pestanas de DATATEST/destino/POA_Destino_Test.xlsx
  3. Empareja archivos con pestanas usando fuzzy matching
  4. Extrae datos de ENERO de cada archivo origen
  5. Imprime un reporte detallado en consola

Uso:
    python -m src.core.test_phase1
    # o desde la raiz del proyecto:
    python src/core/test_phase1.py
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

# Permitir ejecucion directa sin instalar el paquete
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl

from src.core.matcher import normalize_name, build_mapping
from src.core.extractor import extract_month_data, extract_month_data_full

# ---------------------------------------------------------------------------
# Configuracion de rutas (relativas a la raiz del proyecto)
# ---------------------------------------------------------------------------
DATATEST_ROOT = ROOT / "ARCHIVOSPRUEBAS" / "DATATEST"
ORIGEN_DIR = DATATEST_ROOT / "origen_unidades"
DESTINO_FILE = DATATEST_ROOT / "destino" / "POA_Destino_Test.xlsx"
TARGET_MONTH = "ENERO"

# Separador visual
SEP = "=" * 70
SEP2 = "-" * 70


def get_sheet_names(destino_path: Path) -> list:
    """Lee los nombres de pestanas del libro destino."""
    wb = openpyxl.load_workbook(destino_path, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def get_source_files(origen_dir: Path) -> list:
    """Lista los archivos Excel de la carpeta origen."""
    extensions = {".xlsx", ".xls"}
    return sorted(
        f.name
        for f in origen_dir.iterdir()
        if f.is_file() and f.suffix.lower() in extensions
    )


def print_mapping_report(mapping: list) -> None:
    """Imprime el reporte de mapeo archivo -> pestana."""
    print(SEP)
    print("  REPORTE DE MAPEO: Archivos Origen -> Pestanas Destino")
    print(SEP)
    print(f"  {'ARCHIVO ORIGEN':<42} {'NORM':<22} {'PESTANA':<14} {'SCORE':>5}  {'OK?'}")
    print(SEP2)

    matched_count = sum(1 for m in mapping if m["matched"])
    for m in mapping:
        status = "OK" if m["matched"] else "FALLO"
        sheet_str = m["sheet"] if m["sheet"] else "(sin coincidencia)"
        print(
            f"  {m['file']:<42} {m['norm_file']:<22} "
            f"{sheet_str:<14} {m['score']:>5.1f}  {status}"
        )

    print(SEP2)
    print(f"  Total archivos: {len(mapping)} | Emparejados: {matched_count} | Sin mapeo: {len(mapping) - matched_count}")
    print()


def print_extraction_report(mapping: list, origen_dir: Path, month: str) -> None:
    """Extrae y muestra los datos del mes para cada archivo mapeado."""
    print(SEP)
    print(f"  REPORTE DE EXTRACCION — Mes: {month}")
    print(SEP)

    for m in mapping:
        print(f"\n  Archivo : {m['file']}")
        print(f"  Pestana : {m['sheet'] or '(sin mapeo)'}")
        print(f"  Norm    : {m['norm_file']} -> {m['norm_sheet'] or 'N/A'}")
        print(f"  Score   : {m['score']:.1f}")

        if not m["matched"]:
            print("  [OMITIDO] No se encontro pestana destino.")
            continue

        filepath = origen_dir / m["file"]
        try:
            rows = extract_month_data_full(filepath, month)
            simple = extract_month_data(filepath, month)

            print(f"  Datos extraidos ({len(rows)} indicadores):")
            print(f"  {'Ind':>4}  {'PROG':>6}  {'REAL':>6}")
            print(f"  {'-'*4}  {'-'*6}  {'-'*6}")
            for r in rows:
                prog_str = str(r["prog"]) if r["prog"] is not None else "—"
                real_str = str(r["real"]) if r["real"] is not None else "—"
                print(f"  {r['indicador']:>4}  {prog_str:>6}  {real_str:>6}")

            print(f"\n  Diccionario simple {{ind: real}}: {simple}")

        except Exception as exc:
            print(f"  [ERROR] {type(exc).__name__}: {exc}")

        print(SEP2)


def run() -> None:
    """Punto de entrada principal del script de prueba."""
    print()
    print(SEP)
    print("  DATA-TRAVEL — Fase 1: Core Logico — Prueba de integracion")
    print(SEP)
    print(f"  Origen   : {ORIGEN_DIR}")
    print(f"  Destino  : {DESTINO_FILE}")
    print(f"  Mes      : {TARGET_MONTH}")
    print()

    # 1. Leer nombres de pestanas del destino
    sheet_names = get_sheet_names(DESTINO_FILE)
    print(f"  Pestanas en destino ({len(sheet_names)}): {sheet_names}")
    print()

    # 2. Leer archivos origen
    source_files = get_source_files(ORIGEN_DIR)
    print(f"  Archivos en origen ({len(source_files)}):")
    for f in source_files:
        print(f"    {f}  ->  norm: {normalize_name(f)}")
    print()

    # 3. Construir mapeo
    mapping = build_mapping(source_files, sheet_names, score_cutoff=60.0)
    print_mapping_report(mapping)

    # 4. Extraer datos
    print_extraction_report(mapping, ORIGEN_DIR, TARGET_MONTH)

    # 5. Resumen final
    total = len(mapping)
    ok = sum(1 for m in mapping if m["matched"])
    print()
    print(SEP)
    print(f"  RESUMEN FINAL")
    print(SEP)
    print(f"  Archivos procesados  : {total}")
    print(f"  Emparejados con exito: {ok}")
    print(f"  Sin coincidencia     : {total - ok}")
    if ok == total:
        print("  [PASS] Todos los archivos fueron emparejados correctamente.")
    else:
        print("  [WARN] Algunos archivos no encontraron pestana destino.")
    print(SEP)
    print()


if __name__ == "__main__":
    run()
