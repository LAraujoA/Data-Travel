import re
import openpyxl
from pathlib import Path
from src.core.range_migrator import match_dest_sheet

def scan_horizontal_blocks(fp: Path, src_sheet: str | None, id_row: int, col_start: str, width: int, stride: int, dest_sheets: list[str]) -> list[dict]:
    """
    Escanea la fila 'id_row' de 'src_sheet' empezando en 'col_start'.
    Avanza en pasos de 'stride'.
    Cada bloque detectado se le asigna su 'width'.
    Retorna lista para MappingDialog:
    [
        {"label": str, "fp": Path, "src_sheet": str, "dest_sh": str, "score": float, 
         "col_start": int, "col_end": int, "id_row": int}
    ]
    """
    wb = openpyxl.load_workbook(fp, data_only=True, read_only=True)
    if src_sheet and src_sheet in wb.sheetnames:
        ws = wb[src_sheet]
    else:
        ws = wb.active
        src_sheet = ws.title
        
    start_col_idx = openpyxl.utils.column_index_from_string(col_start.upper())
    
    max_col = ws.max_column
    if max_col is None:
        max_col = 1000 # Fallback for read-only if max_col fails
        
    blocks = []
    
    curr_col = start_col_idx
    while curr_col <= max_col:
        cell_val = ws.cell(row=id_row, column=curr_col).value
        
        if not cell_val:
            curr_col += stride
            continue
            
        label = str(cell_val).strip()
        if not label:
            curr_col += stride
            continue
            
        # Ignore total/consolidado
        label_up = label.upper()
        if "TOTAL" in label_up or "CONSOLIDADO" in label_up or "PROMEDIO" in label_up:
            curr_col += stride
            continue
            
        # Match
        dest_sh, score = match_dest_sheet(label, dest_sheets)
        
        blocks.append({
            "label": f"{label} [Col {openpyxl.utils.get_column_letter(curr_col)}]",
            "fp": fp,
            "src_sheet": src_sheet,
            "dest_sh": dest_sh,
            "score": score,
            "col_start": curr_col,
            "col_end": curr_col + width - 1,
            "id_row": id_row
        })
        
        curr_col += stride
        
    wb.close()
    return blocks

def parse_row_mapping(mapping_str: str) -> dict[int, int]:
    """
    Parse '2:16, 3:18' into {2: 16, 3: 18}
    """
    mapping = {}
    tokens = mapping_str.split(",")
    for t in tokens:
        t = t.strip()
        if not t: continue
        parts = t.split(":")
        if len(parts) == 2:
            try:
                mapping[int(parts[0])] = int(parts[1])
            except ValueError:
                raise ValueError(f"Fila invalida en mapeo: {t}")
        else:
            raise ValueError(f"Formato invalido en mapeo: {t}")
    return mapping

def parse_col_dest(cols_str: str) -> list[str]:
    """
    Parse 'S, V, Y, AE, AH' into ['S', 'V', 'Y', 'AE', 'AH']
    """
    cols = [c.strip().upper() for c in cols_str.split(",")]
    cols = [c for c in cols if c]
    return cols

def extract_horizontal_blocks(fp: Path, src_sheet: str, blocks: list[dict], row_mapping: dict[int, int]) -> dict[str, list[list]]:
    """
    Extrae los datos de los bloques configurados.
    Retorna: { "dest_sh": [ [val1, val2, ...], ... ] } (ordenado por fila destino)
    Donde la clave es la hoja destino y el valor son las filas a insertar/actualizar.
    
    Nota: La estructura de retorno depende del Writer. Para batch_update en sheets,
    podriamos necesitar emitir comandos precisos por celda.
    Mejor retorno:
    list de updates: [{"dest_sh": str, "dest_row": int, "values": [v1, v2, ...]}]
    """
    wb = openpyxl.load_workbook(fp, data_only=True)
    ws = wb[src_sheet]
    
    updates = []
    
    for b in blocks:
        dest_sh = b["dest_sh"]
        col_start = b["col_start"]
        col_end = b["col_end"]
        
        for src_r, dest_r in row_mapping.items():
            row_vals = []
            for c in range(col_start +1, col_end + 1):
                row_vals.append(ws.cell(row=src_r, column=c).value)
            
            updates.append({
                "dest_sh": dest_sh,
                "dest_row": dest_r,
                "values": row_vals
            })
            
    wb.close()
    return updates
