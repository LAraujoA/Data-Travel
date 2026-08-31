
"""
app_ui.py - Fase 3: Interfaz Grafica de Usuario - Data-Travel POA Migrador.

Arquitectura de la ventana:
  Header con logo/titulo
  Seccion 1: Seleccion de Mes
  Seccion 2: Carpeta Origen  (campo + botones Seleccionar / Escanear)
  Seccion 3: Tabla de Mapeo (Treeview con scroll)
  Seccion 4: Destino         (SegmentedButton Excel / Sheets + campos dinamicos)
  Seccion 5: Ejecucion       (boton principal, barra de progreso, log)
  Modal de resumen al finalizar

Todos los procesos pesados (scan + transfer) corren en threading.Thread
para que la UI nunca se congele.
"""

from __future__ import annotations

import os
import sys
import queue
import threading
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
from typing import Optional

import customtkinter as ctk

# ---------------------------------------------------------------------------
# Ajuste de PATH para ejecucion directa o como modulo
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.core.matcher import build_mapping
from src.core.extractor import extract_month_data
from src.writers.excel_writer import write_month_data_to_excel

# ---------------------------------------------------------------------------
# Paleta y constantes
# ---------------------------------------------------------------------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

ACCENT     = "#3B8ED0"
ACCENT_DIM = "#1F538A"
SUCCESS    = "#27AE60"
WARNING    = "#F39C12"
DANGER     = "#E74C3C"
BG_CARD    = "#1E2430"
BG_DARK    = "#161B27"
TEXT_MUTED = "#8B9DC3"
TEXT_DIM   = "#5A6480"

MONTHS = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL",
    "MAYO", "JUNIO", "JULIO", "AGOSTO",
    "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
]

WINDOW_W = 840
WINDOW_H = 780
CARD_PAD  = 16
CORNER_R  = 10


# ---------------------------------------------------------------------------
# Helpers de UI
# ---------------------------------------------------------------------------
def _card(parent, **kw) -> ctk.CTkFrame:
    defaults = dict(fg_color=BG_CARD, corner_radius=CORNER_R)
    defaults.update(kw)
    return ctk.CTkFrame(parent, **defaults)


def _section_label(parent, text: str) -> ctk.CTkLabel:
    return ctk.CTkLabel(
        parent, text=text.upper(),
        font=ctk.CTkFont(size=10, weight="bold"),
        text_color=ACCENT, anchor="w",
    )


def _muted(parent, text: str) -> ctk.CTkLabel:
    return ctk.CTkLabel(
        parent, text=text,
        font=ctk.CTkFont(size=11),
        text_color=TEXT_MUTED, anchor="w",
    )


# ---------------------------------------------------------------------------
# Clase principal
# ---------------------------------------------------------------------------
class DataTravelApp(ctk.CTk):
    """Ventana principal de Data-Travel POA Migrador."""

    def __init__(self):
        super().__init__()
        self.title("Data-Travel  —  Migrador de Reportes POA")
        self.geometry(f"{WINDOW_W}x{WINDOW_H}")
        self.minsize(720, 680)
        self.configure(fg_color=BG_DARK)
        self.resizable(True, True)

        # Estado interno
        self._mapping_data: list[dict] = []
        self._origin_dir: Optional[Path] = None
        self._dest_file: Optional[Path] = None
        self._creds_file: Optional[Path] = None
        self._log_queue: queue.Queue = queue.Queue()
        self._transfer_running = False
        self._scan_running = False

        self._build_header()
        self._build_scroll()
        self._build_footer()

        # Iniciar polling de la cola de mensajes
        self.after(100, self._poll_queue)

    # ===== CONSTRUCCION =================================================

    def _build_header(self):
        h = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=66)
        h.pack(fill="x", side="top")
        h.pack_propagate(False)

        ctk.CTkLabel(h, text="🏥", font=ctk.CTkFont(size=28)).pack(
            side="left", padx=(18, 6), pady=12)

        tf = ctk.CTkFrame(h, fg_color="transparent")
        tf.pack(side="left", pady=10)
        ctk.CTkLabel(tf, text="Data-Travel",
                     font=ctk.CTkFont(size=19, weight="bold"),
                     text_color="white").pack(anchor="w")
        ctk.CTkLabel(tf, text="Migrador de Reportes POA  •  MINSAL El Salvador",
                     font=ctk.CTkFont(size=11),
                     text_color=TEXT_MUTED).pack(anchor="w")

        self._theme_btn = ctk.CTkButton(
            h, text="☀", width=36, height=36,
            fg_color="transparent", hover_color=BG_DARK,
            font=ctk.CTkFont(size=17),
            command=self._toggle_theme,
        )
        self._theme_btn.pack(side="right", padx=16)

    def _build_scroll(self):
        self._scroll = ctk.CTkScrollableFrame(
            self, fg_color="transparent", corner_radius=0)
        self._scroll.pack(fill="both", expand=True)
        self._scroll.columnconfigure(0, weight=1)

        self._s1_month(self._scroll)
        self._s2_origen(self._scroll)
        self._s3_mapping(self._scroll)
        self._s4_destino(self._scroll)
        self._s5_exec(self._scroll)

    def _build_footer(self):
        f = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=28)
        f.pack(fill="x", side="bottom")
        f.pack_propagate(False)
        ctk.CTkLabel(f, text="Data-Travel v0.3  •  Fase 3 GUI  •  MINSAL",
                     font=ctk.CTkFont(size=10),
                     text_color=TEXT_DIM).pack(side="left", padx=16, pady=4)

    # ===== SECCION 1: MES ===============================================

    def _s1_month(self, parent):
        c = _card(parent)
        c.pack(fill="x", padx=18, pady=(14, 6))
        c.columnconfigure(2, weight=1)

        _section_label(c, "01  ·  Mes de reporte").grid(
            row=0, column=0, columnspan=3, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 4))

        _muted(c, "Mes objetivo:").grid(
            row=1, column=0, padx=(CARD_PAD, 8), pady=(0, CARD_PAD), sticky="w")

        self._month_var = ctk.StringVar(value="ENERO")
        ctk.CTkOptionMenu(
            c, variable=self._month_var, values=MONTHS,
            width=175, height=34,
            fg_color=ACCENT_DIM, button_color=ACCENT,
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=1, column=1, padx=(0, 6), pady=(0, CARD_PAD), sticky="w")

        _muted(c, "Selecciona el mes a migrar desde los reportes de origen.").grid(
            row=1, column=2, padx=(10, CARD_PAD), pady=(0, CARD_PAD), sticky="w")

    # ===== SECCION 2: ORIGEN ============================================

    def _s2_origen(self, parent):
        c = _card(parent)
        c.pack(fill="x", padx=18, pady=6)
        c.columnconfigure(1, weight=1)

        _section_label(c, "02  ·  Archivos de origen").grid(
            row=0, column=0, columnspan=3, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 4))

        _muted(c, "Carpeta:").grid(
            row=1, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")

        self._origen_entry = ctk.CTkEntry(
            c, placeholder_text="Selecciona la carpeta con los .xlsx de las unidades...",
            height=36, font=ctk.CTkFont(size=12))
        self._origen_entry.grid(row=1, column=1, sticky="ew", pady=(0, 6))

        bf = ctk.CTkFrame(c, fg_color="transparent")
        bf.grid(row=1, column=2, padx=(8, CARD_PAD), pady=(0, 6))

        ctk.CTkButton(
            bf, text="📂  Seleccionar", width=135, height=36,
            fg_color=ACCENT_DIM, hover_color=ACCENT,
            font=ctk.CTkFont(size=12),
            command=self._select_origen,
        ).pack(side="left", padx=(0, 6))

        self._scan_btn = ctk.CTkButton(
            bf, text="🔍  Escanear", width=115, height=36,
            fg_color="#2D4A6E", hover_color=ACCENT_DIM,
            font=ctk.CTkFont(size=12),
            command=self._scan_files,
        )
        self._scan_btn.pack(side="left")

        self._file_count_lbl = _muted(c, "")
        self._file_count_lbl.grid(
            row=2, column=0, columnspan=3, sticky="w",
            padx=CARD_PAD, pady=(0, CARD_PAD))

    # ===== SECCION 3: TABLA DE MAPEO ====================================

    def _s3_mapping(self, parent):
        c = _card(parent)
        c.pack(fill="x", padx=18, pady=6)
        c.columnconfigure(0, weight=1)

        _section_label(c, "03  ·  Previsualizacion del mapeo").grid(
            row=0, column=0, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 6))

        # Estilo personalizado para Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("DT.Treeview",
            background="#1A2035", foreground="#D0D8F0",
            fieldbackground="#1A2035", borderwidth=0,
            font=("Segoe UI", 11), rowheight=30)
        style.configure("DT.Treeview.Heading",
            background="#0F1420", foreground=ACCENT,
            relief="flat", font=("Segoe UI", 10, "bold"))
        style.map("DT.Treeview",
            background=[("selected", ACCENT_DIM)],
            foreground=[("selected", "white")])

        tf = ctk.CTkFrame(c, fg_color="#1A2035", corner_radius=8)
        tf.grid(row=1, column=0, sticky="ew", padx=CARD_PAD, pady=(0, CARD_PAD))

        cols = ("archivo", "pestana", "score", "estado")
        self._tree = ttk.Treeview(
            tf, columns=cols, show="headings",
            height=7, style="DT.Treeview")

        self._tree.heading("archivo",  text="  Archivo Origen")
        self._tree.heading("pestana",  text="Pestaña Destino")
        self._tree.heading("score",    text="Coincidencia")
        self._tree.heading("estado",   text="Estado")

        self._tree.column("archivo",  width=315, anchor="w", stretch=True)
        self._tree.column("pestana",  width=165, anchor="w")
        self._tree.column("score",    width=105, anchor="center")
        self._tree.column("estado",   width=125, anchor="center")

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        vsb.pack(side="right", fill="y", pady=4)

        self._tree.tag_configure("ok",    foreground="#5BEB8A")
        self._tree.tag_configure("warn",  foreground=WARNING)
        self._tree.tag_configure("error", foreground=DANGER)

    # ===== SECCION 4: DESTINO ===========================================

    def _s4_destino(self, parent):
        c = _card(parent)
        c.pack(fill="x", padx=18, pady=6)
        c.columnconfigure(1, weight=1)

        _section_label(c, "04  ·  Destino de escritura").grid(
            row=0, column=0, columnspan=3, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 6))

        self._dest_mode = ctk.StringVar(value="Excel Local")
        ctk.CTkSegmentedButton(
            c,
            values=["Excel Local", "Google Sheets"],
            variable=self._dest_mode,
            command=self._on_dest_change,
            width=295, height=36,
            font=ctk.CTkFont(size=13),
            selected_color=ACCENT, selected_hover_color=ACCENT_DIM,
            unselected_color="#2D3A55",
        ).grid(row=1, column=0, columnspan=3, sticky="w",
               padx=CARD_PAD, pady=(0, 10))

        # --- Panel Excel ---
        self._excel_pnl = ctk.CTkFrame(c, fg_color="transparent")
        self._excel_pnl.grid(row=2, column=0, columnspan=3,
                              sticky="ew", padx=0, pady=0)
        self._excel_pnl.columnconfigure(1, weight=1)

        _muted(self._excel_pnl, "Archivo POA:").grid(
            row=0, column=0, padx=(CARD_PAD, 8), pady=(0, CARD_PAD), sticky="w")
        self._dest_entry = ctk.CTkEntry(
            self._excel_pnl,
            placeholder_text="Ruta al archivo Excel destino (.xlsx)...",
            height=36, font=ctk.CTkFont(size=12))
        self._dest_entry.grid(row=0, column=1, sticky="ew", pady=(0, CARD_PAD))
        ctk.CTkButton(
            self._excel_pnl, text="📄  Seleccionar",
            width=135, height=36,
            fg_color=ACCENT_DIM, hover_color=ACCENT,
            font=ctk.CTkFont(size=12),
            command=self._select_dest,
        ).grid(row=0, column=2, padx=(8, CARD_PAD), pady=(0, CARD_PAD))

        # --- Panel Sheets (oculto) ---
        self._sheets_pnl = ctk.CTkFrame(c, fg_color="transparent")
        self._sheets_pnl.columnconfigure(1, weight=1)

        _muted(self._sheets_pnl, "URL / ID Sheet:").grid(
            row=0, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")
        self._sheets_url = ctk.CTkEntry(
            self._sheets_pnl,
            placeholder_text="https://docs.google.com/spreadsheets/d/...",
            height=36, font=ctk.CTkFont(size=12))
        self._sheets_url.grid(
            row=0, column=1, columnspan=2, sticky="ew",
            padx=(0, CARD_PAD), pady=(0, 6))

        _muted(self._sheets_pnl, "Credenciales:").grid(
            row=1, column=0, padx=(CARD_PAD, 8), pady=(0, CARD_PAD), sticky="w")
        self._creds_entry = ctk.CTkEntry(
            self._sheets_pnl,
            placeholder_text="credentials.json  (service account)",
            height=36, font=ctk.CTkFont(size=12))
        self._creds_entry.grid(row=1, column=1, sticky="ew", pady=(0, CARD_PAD))
        ctk.CTkButton(
            self._sheets_pnl, text="🔑  Seleccionar",
            width=135, height=36,
            fg_color=ACCENT_DIM, hover_color=ACCENT,
            font=ctk.CTkFont(size=12),
            command=self._select_creds,
        ).grid(row=1, column=2, padx=(8, CARD_PAD), pady=(0, CARD_PAD))

    # ===== SECCION 5: EJECUCION =========================================

    def _s5_exec(self, parent):
        c = _card(parent)
        c.pack(fill="x", padx=18, pady=(6, 20))
        c.columnconfigure(0, weight=1)

        _section_label(c, "05  ·  Ejecucion").grid(
            row=0, column=0, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 6))

        self._transfer_btn = ctk.CTkButton(
            c, text="⚡  Transferir Datos",
            height=46, font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=SUCCESS, hover_color="#1E8449",
            corner_radius=CORNER_R,
            command=self._on_transfer,
        )
        self._transfer_btn.grid(
            row=1, column=0, sticky="ew", padx=CARD_PAD, pady=(0, 10))

        self._progress = ctk.CTkProgressBar(
            c, height=8, corner_radius=4,
            progress_color=ACCENT)
        self._progress.set(0)
        self._progress.grid(
            row=2, column=0, sticky="ew", padx=CARD_PAD, pady=(0, 6))

        self._status_lbl = ctk.CTkLabel(
            c, text="Listo para comenzar.",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_MUTED, anchor="w")
        self._status_lbl.grid(
            row=3, column=0, sticky="w", padx=CARD_PAD, pady=(0, 6))

        self._log_box = ctk.CTkTextbox(
            c, height=170,
            font=ctk.CTkFont(family="Cascadia Code", size=11),
            fg_color="#0D1117", text_color="#8FD3A7",
            corner_radius=8, wrap="word", state="disabled")
        self._log_box.grid(
            row=4, column=0, sticky="ew", padx=CARD_PAD, pady=(0, 6))

        ctk.CTkButton(
            c, text="🗑  Limpiar log",
            width=120, height=28,
            fg_color="transparent", hover_color="#2D3A55",
            border_width=1, border_color=TEXT_DIM,
            font=ctk.CTkFont(size=11), text_color=TEXT_MUTED,
            command=self._clear_log,
        ).grid(row=5, column=0, sticky="e", padx=CARD_PAD, pady=(0, CARD_PAD))

    # ===== LOGICA DE INTERACCION ========================================

    def _toggle_theme(self):
        current = ctk.get_appearance_mode()
        new = "light" if current == "Dark" else "dark"
        ctk.set_appearance_mode(new)
        self._theme_btn.configure(text="🌙" if new == "dark" else "☀")

    def _on_dest_change(self, val):
        if val == "Excel Local":
            self._sheets_pnl.grid_forget()
            self._excel_pnl.grid(row=2, column=0, columnspan=3,
                                  sticky="ew", padx=0, pady=0)
        else:
            self._excel_pnl.grid_forget()
            self._sheets_pnl.grid(row=2, column=0, columnspan=3,
                                   sticky="ew", padx=0, pady=0)

    def _select_origen(self):
        folder = filedialog.askdirectory(
            title="Selecciona la carpeta con los archivos .xlsx de las unidades")
        if folder:
            self._origin_dir = Path(folder)
            self._origen_entry.delete(0, "end")
            self._origen_entry.insert(0, str(self._origin_dir))
            self._log(f"📂 Carpeta seleccionada: {self._origin_dir.name}")

    def _select_dest(self):
        path = filedialog.askopenfilename(
            title="Selecciona el archivo Excel destino (POA consolidado)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if path:
            self._dest_file = Path(path)
            self._dest_entry.delete(0, "end")
            self._dest_entry.insert(0, str(self._dest_file))
            self._log(f"📄 Destino Excel: {self._dest_file.name}")

    def _select_creds(self):
        path = filedialog.askopenfilename(
            title="Selecciona el archivo credentials.json",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")])
        if path:
            self._creds_file = Path(path)
            self._creds_entry.delete(0, "end")
            self._creds_entry.insert(0, str(self._creds_file))
            self._log(f"🔑 Credenciales: {self._creds_file.name}")

    def _get_dest_path(self) -> Optional[Path]:
        if self._dest_mode.get() == "Excel Local":
            val = self._dest_entry.get().strip()
            if val:
                self._dest_file = Path(val)
            if self._dest_file and self._dest_file.exists():
                return self._dest_file
        return None

    # ===== ESCANEO ======================================================

    def _scan_files(self):
        val = self._origen_entry.get().strip()
        if val:
            self._origin_dir = Path(val)

        if not self._origin_dir or not self._origin_dir.is_dir():
            messagebox.showwarning(
                "Carpeta no valida",
                "Por favor selecciona una carpeta de origen valida.")
            return

        dest = self._get_dest_path()
        if dest is None:
            messagebox.showwarning(
                "Destino requerido",
                "Selecciona el archivo destino (.xlsx) antes de escanear\n"
                "para poder previsualizar el mapeo de pestanas.")
            return

        if self._scan_running:
            return
        self._scan_running = True
        self._scan_btn.configure(state="disabled", text="⏳  Escaneando...")
        self._clear_tree()
        self._status_lbl.configure(text="Escaneando archivos...", text_color=ACCENT)
        self._log("🔍 Escaneando archivos de origen...")
        threading.Thread(target=self._scan_worker, args=(dest,), daemon=True).start()

    def _scan_worker(self, dest_path: Path):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(dest_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()

            files = sorted(
                f.name for f in self._origin_dir.iterdir()
                if f.suffix.lower() in {".xlsx", ".xls"})

            if not files:
                self._qlog("⚠  No se encontraron archivos .xlsx en la carpeta.", "warn")
                self._log_queue.put(("scan_done", []))
                return

            self._qlog(f"  Encontrados {len(files)} archivo(s) Excel.")
            mapping = build_mapping(files, sheet_names, score_cutoff=60.0)
            self._mapping_data = mapping
            self._log_queue.put(("scan_done", mapping))

        except PermissionError:
            self._qlog("❌  Sin permiso para leer el archivo destino.", "error")
            self._log_queue.put(("scan_done", []))
        except Exception as exc:
            self._qlog(f"❌  Error en escaneo: {exc}", "error")
            self._log_queue.put(("scan_done", []))
        finally:
            self._scan_running = False

    def _populate_tree(self, mapping: list):
        self._clear_tree()
        matched = 0
        for m in mapping:
            score_s = f"{m['score']:.1f} %" if m["matched"] else "—"
            if m["matched"] and m["score"] >= 90:
                estado, tag = "✅ Listo",    "ok"
                matched += 1
            elif m["matched"]:
                estado, tag = "⚠ Revisar",  "warn"
                matched += 1
            else:
                estado, tag = "❌ Sin mapeo", "error"

            self._tree.insert("", "end",
                values=("  " + m["file"], m["sheet"] or "—", score_s, estado),
                tags=(tag,))

        total = len(mapping)
        color = SUCCESS if matched == total else WARNING
        self._file_count_lbl.configure(
            text=(f"  {total} archivo(s)  ·  "
                  f"{matched} mapeado(s)  ·  "
                  f"{total - matched} sin coincidencia"),
            text_color=color)
        self._status_lbl.configure(
            text=f"Escaneo completado: {matched}/{total} archivos mapeados.",
            text_color=SUCCESS if matched == total else WARNING)

    def _clear_tree(self):
        for row in self._tree.get_children():
            self._tree.delete(row)

    # ===== TRANSFERENCIA ================================================

    def _on_transfer(self):
        if self._transfer_running:
            return

        val = self._origen_entry.get().strip()
        if val:
            self._origin_dir = Path(val)

        if not self._origin_dir or not self._origin_dir.is_dir():
            messagebox.showwarning("Origen requerido",
                "Selecciona la carpeta de archivos de origen.")
            return

        dest = self._get_dest_path()
        if dest is None and self._dest_mode.get() == "Excel Local":
            messagebox.showwarning("Destino requerido",
                "Selecciona un archivo Excel destino valido (.xlsx).")
            return

        if not self._mapping_data:
            if messagebox.askyesno(
                "Sin escaneo previo",
                "No has escaneado los archivos aun.\n"
                "¿Deseas escanear y transferir automaticamente?",
            ):
                dest_copy = dest
                self._scan_files()
                self.after(2000, lambda: self._start_transfer(dest_copy))
            return

        self._start_transfer(dest)

    def _start_transfer(self, dest_path: Optional[Path]):
        matched = [m for m in self._mapping_data if m["matched"]]
        if not matched:
            messagebox.showwarning("Sin mapeo",
                "No hay archivos mapeados. Escanea primero.")
            return

        self._transfer_running = True
        self._transfer_btn.configure(state="disabled", text="⏳  Transfiriendo...")
        self._progress.set(0)
        self._status_lbl.configure(
            text="Iniciando transferencia...", text_color=ACCENT)
        self._log("─" * 54)
        self._log(f"🚀 Mes: {self._month_var.get()}  |  "
                  f"Modo: {self._dest_mode.get()}  |  "
                  f"Unidades: {len(matched)}")

        threading.Thread(
            target=self._transfer_worker,
            args=(dest_path,),
            daemon=True).start()

    def _transfer_worker(self, dest_path: Optional[Path]):
        month = self._month_var.get()
        mode  = self._dest_mode.get()
        matched = [m for m in self._mapping_data if m["matched"]]
        total = len(matched)
        cells_written = 0
        failed: list[str] = []

        for i, m in enumerate(matched):
            sheet = m["sheet"]
            fpath = self._origin_dir / m["file"]
            self._qlog(f"  [{i+1}/{total}]  {m['file']}  →  '{sheet}'")
            self._log_queue.put(("progress", i / total))

            # --- Extraccion ---
            try:
                data = extract_month_data(fpath, month)
                self._qlog(f"    Extraidos {len(data)} indicadores.")
            except Exception as exc:
                self._qlog(f"    ❌ Error extrayendo: {exc}", "error")
                failed.append(sheet)
                continue

            # --- Escritura ---
            if mode == "Excel Local" and dest_path:
                try:
                    ok = write_month_data_to_excel(
                        dest_path, sheet, month, data,
                        create_backup=(i == 0))
                    if ok:
                        cells_written += len(data)
                        self._qlog(f"    ✅  {len(data)} celdas REAL escritas.")
                    else:
                        failed.append(sheet)
                        self._qlog(f"    ❌  write_month_data_to_excel retorno False.", "error")
                except PermissionError:
                    self._qlog(
                        "    ❌  ARCHIVO ABIERTO EN EXCEL.\n"
                        "       Cierralo y vuelve a intentarlo.", "error")
                    failed.append(sheet)
                    break
                except Exception as exc:
                    self._qlog(f"    ❌  Error escritura: {exc}", "error")
                    failed.append(sheet)

            elif mode == "Google Sheets":
                url = self._sheets_url.get().strip()
                creds = str(self._creds_file) if self._creds_file else "credentials.json"
                try:
                    from src.writers.sheets_writer import write_month_data_to_sheets
                    ok = write_month_data_to_sheets(
                        url, sheet, month, data, credentials_path=creds)
                    if ok:
                        cells_written += len(data)
                        self._qlog(f"    ✅  Google Sheets actualizado.")
                    else:
                        failed.append(sheet)
                except ImportError:
                    self._qlog("    ❌  gspread no instalado. Ejecuta: pip install gspread google-auth", "error")
                    break
                except Exception as exc:
                    self._qlog(f"    ❌  Error Sheets: {exc}", "error")
                    failed.append(sheet)

            self._log_queue.put(("progress", (i + 1) / total))

        self._log_queue.put(("transfer_done", {
            "total": total, "failed": failed,
            "cells": cells_written, "month": month, "mode": mode,
        }))

    # ===== COLA DE MENSAJES (hilo principal) ============================

    def _poll_queue(self):
        try:
            while True:
                kind, data = self._log_queue.get_nowait()

                if kind == "log":
                    self._append_log(*data)

                elif kind == "progress":
                    self._progress.set(data)

                elif kind == "scan_done":
                    self._populate_tree(data)
                    self._scan_btn.configure(state="normal", text="🔍  Escanear")

                elif kind == "transfer_done":
                    self._on_done(data)

        except queue.Empty:
            pass
        finally:
            self.after(100, self._poll_queue)

    def _on_done(self, r: dict):
        self._transfer_running = False
        ok_count = r["total"] - len(r["failed"])
        all_ok = not r["failed"]

        self._progress.set(1.0)
        self._transfer_btn.configure(state="normal", text="⚡  Transferir Datos")
        self._status_lbl.configure(
            text=f"{'✅' if all_ok else '⚠'}  {ok_count}/{r['total']} unidades "
                 f"· {r['cells']} celdas actualizadas",
            text_color=SUCCESS if all_ok else WARNING)
        self._log("─" * 54)
        self._log(f"{'✅' if all_ok else '⚠'}  Transferencia finalizada.")

        lines = [
            f"Mes procesado   :  {r['month']}",
            f"Modo destino    :  {r['mode']}",
            f"Unidades OK     :  {ok_count} / {r['total']}",
            f"Celdas REAL     :  {r['cells']} actualizadas",
        ]
        if r["failed"]:
            lines.append(f"Con errores     :  {', '.join(r['failed'])}")

        summary = "\n".join(lines)
        title  = "✅  Transferencia Exitosa" if all_ok else "⚠  Transferencia Parcial"
        msg    = (
            ("¡La migración se completó con éxito!\n\n" if all_ok
             else "La migración terminó con algunos errores.\n\n")
            + summary
            + ("" if all_ok else "\n\nRevisa el log para más detalles.")
        )
        (messagebox.showinfo if all_ok else messagebox.showwarning)(title, msg)

    # ===== LOG HELPERS ==================================================

    def _qlog(self, msg: str, level: str = "info"):
        """Hilo seguro: encola el mensaje para pintarlo en el hilo principal."""
        self._log_queue.put(("log", (msg, level)))

    def _log(self, msg: str, level: str = "info"):
        """Alias de _qlog para llamadas desde el hilo principal."""
        self._log_queue.put(("log", (msg, level)))

    def _append_log(self, msg: str, level: str = "info"):
        """Pinta un mensaje en el log. DEBE llamarse desde el hilo principal."""
        color = {"info": "#8FD3A7", "warn": "#F5C842", "error": "#FF6B6B"}.get(level, "#8FD3A7")
        self._log_box.configure(state="normal", text_color=color)
        self._log_box.insert("end", msg + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")
        self._progress.set(0)
        self._status_lbl.configure(text="Log limpiado.", text_color=TEXT_MUTED)
