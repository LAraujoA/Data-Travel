
"""
app_ui.py — Fase 4: GUI con dos pestanas (CTkTabview).

Pestana 1: Reportes POA  (lote automatico — Fases 1+2)
Pestana 2: Migrador Universal (rangos libres — range_migrator)

Todos los procesos pesados corren en threading.Thread.
La comunicacion con el hilo principal se hace via queue.Queue + polling 100 ms.
"""

from __future__ import annotations

import queue
import re
import sys
import threading
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
from typing import Optional

import customtkinter as ctk

# ── PATH ──────────────────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.core.matcher import build_mapping
from src.core.extractor import extract_month_data
from src.writers.excel_writer import write_month_data_to_excel
from src.core.range_migrator import (
    extract_range, get_sheet_names, migrate_range, MODES,
)

# ── TEMA ──────────────────────────────────────────────────────────────────────
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
PURPLE     = "#8B5CF6"
PURPLE_DIM = "#5B21B6"

MONTHS = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL",
    "MAYO", "JUNIO", "JULIO", "AGOSTO",
    "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
]

WINDOW_W = 880
WINDOW_H = 820
CARD_PAD  = 16
CORNER_R  = 10


# ── HELPERS ───────────────────────────────────────────────────────────────────
def _card(parent, **kw) -> ctk.CTkFrame:
    d = dict(fg_color=BG_CARD, corner_radius=CORNER_R)
    d.update(kw)
    return ctk.CTkFrame(parent, **d)


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


def _entry(parent, placeholder: str = "", width: int = 0,
           font_size: int = 12) -> ctk.CTkEntry:
    kw = dict(placeholder_text=placeholder, height=34,
              font=ctk.CTkFont(size=font_size))
    if width:
        kw["width"] = width
    return ctk.CTkEntry(parent, **kw)


def _btn(parent, text: str, command=None, color=ACCENT_DIM,
         hover=ACCENT, width: int = 0, height: int = 34) -> ctk.CTkButton:
    kw = dict(text=text, fg_color=color, hover_color=hover,
              height=height, font=ctk.CTkFont(size=12), command=command)
    if width:
        kw["width"] = width
    return ctk.CTkButton(parent, **kw)


# ── CLASE PRINCIPAL ───────────────────────────────────────────────────────────
class DataTravelApp(ctk.CTk):
    """Ventana principal de Data-Travel."""

    def __init__(self):
        super().__init__()
        self.title("Data-Travel  —  Migrador de Reportes POA")
        self.geometry(f"{WINDOW_W}x{WINDOW_H}")
        self.minsize(800, 720)
        self.configure(fg_color=BG_DARK)
        self.resizable(True, True)

        # Estado POA
        self._poa_mapping:   list[dict] = []
        self._poa_origin:    Optional[Path] = None
        self._poa_dest:      Optional[Path] = None
        self._poa_creds:     Optional[Path] = None
        self._poa_scanning   = False
        self._poa_running    = False

        # Estado Migrador Universal
        self._uni_src_file:  Optional[Path] = None
        self._uni_dest_file: Optional[Path] = None
        self._uni_running    = False

        # Cola de mensajes compartida (se identifica con "tab")
        self._q: queue.Queue = queue.Queue()

        self._build_header()
        self._build_tabs()
        self._build_footer()

        self.after(100, self._poll)

    # ── FRAME GLOBAL ──────────────────────────────────────────────────────────

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
                     font=ctk.CTkFont(size=11), text_color=TEXT_MUTED).pack(anchor="w")
        self._theme_btn = ctk.CTkButton(
            h, text="☀", width=36, height=36,
            fg_color="transparent", hover_color=BG_DARK,
            font=ctk.CTkFont(size=17), command=self._toggle_theme)
        self._theme_btn.pack(side="right", padx=16)

    def _build_footer(self):
        f = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=28)
        f.pack(fill="x", side="bottom")
        f.pack_propagate(False)
        ctk.CTkLabel(f, text="Data-Travel v0.4  •  Fase 4  •  MINSAL",
                     font=ctk.CTkFont(size=10),
                     text_color=TEXT_DIM).pack(side="left", padx=16, pady=4)

    def _build_tabs(self):
        self._tabs = ctk.CTkTabview(
            self,
            fg_color=BG_DARK,
            segmented_button_fg_color=BG_CARD,
            segmented_button_selected_color=ACCENT,
            segmented_button_selected_hover_color=ACCENT_DIM,
            segmented_button_unselected_color=BG_CARD,
            segmented_button_unselected_hover_color="#2D3A55",
            text_color="white",
            text_color_disabled=TEXT_DIM,
        )
        self._tabs.pack(fill="both", expand=True, padx=0, pady=0)

        self._tabs.add("🏥  Reportes POA (Lote Automático)")
        self._tabs.add("🎯  Migrador Universal")

        self._build_poa_tab(self._tabs.tab("🏥  Reportes POA (Lote Automático)"))
        self._build_uni_tab(self._tabs.tab("🎯  Migrador Universal"))

    def _toggle_theme(self):
        new = "light" if ctk.get_appearance_mode() == "Dark" else "dark"
        ctk.set_appearance_mode(new)
        self._theme_btn.configure(text="🌙" if new == "dark" else "☀")

    # ══════════════════════════════════════════════════════════════════════════
    #  PESTANA 1 — REPORTES POA
    # ══════════════════════════════════════════════════════════════════════════

    def _build_poa_tab(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True)
        scroll.columnconfigure(0, weight=1)

        self._poa_s1_month(scroll)
        self._poa_s2_origen(scroll)
        self._poa_s3_mapping(scroll)
        self._poa_s4_destino(scroll)
        self._poa_s5_exec(scroll)

    # ── S1 Mes ────────────────────────────────────────────────────────────────
    def _poa_s1_month(self, p):
        c = _card(p); c.pack(fill="x", padx=14, pady=(12, 5))
        c.columnconfigure(2, weight=1)
        _section_label(c, "01  ·  Mes de reporte").grid(
            row=0, column=0, columnspan=3, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 4))
        _muted(c, "Mes objetivo:").grid(
            row=1, column=0, padx=(CARD_PAD, 8), pady=(0, CARD_PAD), sticky="w")
        self._month_var = ctk.StringVar(value="ENERO")
        ctk.CTkOptionMenu(c, variable=self._month_var, values=MONTHS,
                          width=175, height=34,
                          fg_color=ACCENT_DIM, button_color=ACCENT,
                          font=ctk.CTkFont(size=13, weight="bold"),
                          ).grid(row=1, column=1, padx=(0, 6),
                                 pady=(0, CARD_PAD), sticky="w")
        _muted(c, "Selecciona el mes a migrar desde los reportes de origen.").grid(
            row=1, column=2, padx=(10, CARD_PAD), pady=(0, CARD_PAD), sticky="w")

    # ── S2 Origen ─────────────────────────────────────────────────────────────
    def _poa_s2_origen(self, p):
        c = _card(p); c.pack(fill="x", padx=14, pady=5)
        c.columnconfigure(1, weight=1)
        _section_label(c, "02  ·  Archivos de origen").grid(
            row=0, column=0, columnspan=3, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 4))
        _muted(c, "Carpeta:").grid(
            row=1, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")
        self._poa_origen_entry = _entry(
            c, "Selecciona la carpeta con los .xlsx de las unidades...")
        self._poa_origen_entry.grid(row=1, column=1, sticky="ew", pady=(0, 6))
        bf = ctk.CTkFrame(c, fg_color="transparent")
        bf.grid(row=1, column=2, padx=(8, CARD_PAD), pady=(0, 6))
        _btn(bf, "📂  Seleccionar", self._poa_select_origen,
             width=130).pack(side="left", padx=(0, 6))
        self._poa_scan_btn = _btn(bf, "🔍  Escanear", self._poa_scan,
                                  color="#2D4A6E", width=115)
        self._poa_scan_btn.pack(side="left")
        self._poa_file_lbl = _muted(c, "")
        self._poa_file_lbl.grid(
            row=2, column=0, columnspan=3, sticky="w",
            padx=CARD_PAD, pady=(0, CARD_PAD))

    # ── S3 Mapeo ──────────────────────────────────────────────────────────────
    def _poa_s3_mapping(self, p):
        c = _card(p); c.pack(fill="x", padx=14, pady=5)
        c.columnconfigure(0, weight=1)
        _section_label(c, "03  ·  Previsualizacion del mapeo").grid(
            row=0, column=0, sticky="w", padx=CARD_PAD, pady=(CARD_PAD, 6))

        style = ttk.Style(); style.theme_use("clam")
        style.configure("DT.Treeview", background="#1A2035", foreground="#D0D8F0",
            fieldbackground="#1A2035", borderwidth=0,
            font=("Segoe UI", 11), rowheight=28)
        style.configure("DT.Treeview.Heading", background="#0F1420", foreground=ACCENT,
            relief="flat", font=("Segoe UI", 10, "bold"))
        style.map("DT.Treeview",
            background=[("selected", ACCENT_DIM)],
            foreground=[("selected", "white")])

        tf = ctk.CTkFrame(c, fg_color="#1A2035", corner_radius=8)
        tf.grid(row=1, column=0, sticky="ew", padx=CARD_PAD, pady=(0, CARD_PAD))
        cols = ("archivo", "pestana", "score", "estado")
        self._poa_tree = ttk.Treeview(tf, columns=cols, show="headings",
                                      height=6, style="DT.Treeview")
        for col, txt, w in [
            ("archivo", "  Archivo Origen", 310),
            ("pestana", "Pestaña Destino",  165),
            ("score",   "Coincidencia",     105),
            ("estado",  "Estado",           120),
        ]:
            self._poa_tree.heading(col, text=txt)
            self._poa_tree.column(col, width=w,
                anchor="w" if col in ("archivo", "pestana") else "center")
        vsb = ttk.Scrollbar(tf, orient="vertical", command=self._poa_tree.yview)
        self._poa_tree.configure(yscrollcommand=vsb.set)
        self._poa_tree.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        vsb.pack(side="right", fill="y", pady=4)
        self._poa_tree.tag_configure("ok",    foreground="#5BEB8A")
        self._poa_tree.tag_configure("warn",  foreground=WARNING)
        self._poa_tree.tag_configure("error", foreground=DANGER)

    # ── S4 Destino ────────────────────────────────────────────────────────────
    def _poa_s4_destino(self, p):
        c = _card(p); c.pack(fill="x", padx=14, pady=5)
        c.columnconfigure(1, weight=1)
        _section_label(c, "04  ·  Destino de escritura").grid(
            row=0, column=0, columnspan=3, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 6))
        self._poa_dest_mode = ctk.StringVar(value="Excel Local")
        ctk.CTkSegmentedButton(c, values=["Excel Local", "Google Sheets"],
            variable=self._poa_dest_mode, command=self._poa_dest_change,
            width=295, height=34, font=ctk.CTkFont(size=13),
            selected_color=ACCENT, selected_hover_color=ACCENT_DIM,
            unselected_color="#2D3A55",
        ).grid(row=1, column=0, columnspan=3, sticky="w",
               padx=CARD_PAD, pady=(0, 10))

        # Excel panel
        self._poa_excel_pnl = ctk.CTkFrame(c, fg_color="transparent")
        self._poa_excel_pnl.grid(row=2, column=0, columnspan=3,
                                  sticky="ew", padx=0, pady=0)
        self._poa_excel_pnl.columnconfigure(1, weight=1)
        _muted(self._poa_excel_pnl, "Archivo POA:").grid(
            row=0, column=0, padx=(CARD_PAD, 8), pady=(0, CARD_PAD), sticky="w")
        self._poa_dest_entry = _entry(
            self._poa_excel_pnl, "Ruta al archivo Excel destino (.xlsx)...")
        self._poa_dest_entry.grid(row=0, column=1, sticky="ew", pady=(0, CARD_PAD))
        _btn(self._poa_excel_pnl, "📄  Seleccionar", self._poa_select_dest,
             width=130).grid(row=0, column=2, padx=(8, CARD_PAD), pady=(0, CARD_PAD))

        # Sheets panel (oculto)
        self._poa_sheets_pnl = ctk.CTkFrame(c, fg_color="transparent")
        self._poa_sheets_pnl.columnconfigure(1, weight=1)
        _muted(self._poa_sheets_pnl, "URL / ID Sheet:").grid(
            row=0, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")
        self._poa_sheets_url = _entry(
            self._poa_sheets_pnl,
            "https://docs.google.com/spreadsheets/d/...")
        self._poa_sheets_url.grid(
            row=0, column=1, columnspan=2, sticky="ew",
            padx=(0, CARD_PAD), pady=(0, 6))
        _muted(self._poa_sheets_pnl, "Credenciales:").grid(
            row=1, column=0, padx=(CARD_PAD, 8), pady=(0, CARD_PAD), sticky="w")
        self._poa_creds_entry = _entry(
            self._poa_sheets_pnl, "credentials.json")
        self._poa_creds_entry.grid(
            row=1, column=1, sticky="ew", pady=(0, CARD_PAD))
        _btn(self._poa_sheets_pnl, "🔑  Seleccionar", self._poa_select_creds,
             width=130).grid(row=1, column=2, padx=(8, CARD_PAD), pady=(0, CARD_PAD))

    # ── S5 Ejecucion ──────────────────────────────────────────────────────────
    def _poa_s5_exec(self, p):
        c = _card(p); c.pack(fill="x", padx=14, pady=(5, 16))
        c.columnconfigure(0, weight=1)
        _section_label(c, "05  ·  Ejecucion").grid(
            row=0, column=0, sticky="w", padx=CARD_PAD, pady=(CARD_PAD, 6))
        self._poa_run_btn = ctk.CTkButton(
            c, text="⚡  Transferir Datos", height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=SUCCESS, hover_color="#1E8449",
            corner_radius=CORNER_R, command=self._poa_on_transfer)
        self._poa_run_btn.grid(row=1, column=0, sticky="ew",
                               padx=CARD_PAD, pady=(0, 8))
        self._poa_bar = ctk.CTkProgressBar(c, height=8, corner_radius=4,
                                            progress_color=ACCENT)
        self._poa_bar.set(0)
        self._poa_bar.grid(row=2, column=0, sticky="ew",
                           padx=CARD_PAD, pady=(0, 6))
        self._poa_status = ctk.CTkLabel(
            c, text="Listo.", font=ctk.CTkFont(size=12),
            text_color=TEXT_MUTED, anchor="w")
        self._poa_status.grid(row=3, column=0, sticky="w",
                              padx=CARD_PAD, pady=(0, 6))
        self._poa_log = ctk.CTkTextbox(
            c, height=150,
            font=ctk.CTkFont(family="Cascadia Code", size=11),
            fg_color="#0D1117", text_color="#8FD3A7",
            corner_radius=8, wrap="word", state="disabled")
        self._poa_log.grid(row=4, column=0, sticky="ew",
                           padx=CARD_PAD, pady=(0, 6))
        _btn(c, "🗑  Limpiar log", self._poa_clear_log,
             color="transparent", hover="#2D3A55",
             width=120, height=28).grid(
             row=5, column=0, sticky="e", padx=CARD_PAD, pady=(0, CARD_PAD))

    # ── Logica POA ────────────────────────────────────────────────────────────
    def _poa_dest_change(self, val):
        if val == "Excel Local":
            self._poa_sheets_pnl.grid_forget()
            self._poa_excel_pnl.grid(row=2, column=0, columnspan=3,
                                      sticky="ew")
        else:
            self._poa_excel_pnl.grid_forget()
            self._poa_sheets_pnl.grid(row=2, column=0, columnspan=3,
                                       sticky="ew")

    def _poa_select_origen(self):
        d = filedialog.askdirectory(
            title="Selecciona carpeta con .xlsx de las unidades")
        if d:
            self._poa_origin = Path(d)
            self._poa_origen_entry.delete(0, "end")
            self._poa_origen_entry.insert(0, str(self._poa_origin))
            self._poa_log_msg(f"📂 Carpeta: {self._poa_origin.name}")

    def _poa_select_dest(self):
        p = filedialog.askopenfilename(
            title="Selecciona el Excel destino",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if p:
            self._poa_dest = Path(p)
            self._poa_dest_entry.delete(0, "end")
            self._poa_dest_entry.insert(0, str(self._poa_dest))
            self._poa_log_msg(f"📄 Destino: {self._poa_dest.name}")

    def _poa_select_creds(self):
        p = filedialog.askopenfilename(
            title="Selecciona credentials.json",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")])
        if p:
            self._poa_creds = Path(p)
            self._poa_creds_entry.delete(0, "end")
            self._poa_creds_entry.insert(0, str(self._poa_creds))

    def _poa_get_dest(self) -> Optional[Path]:
        if self._poa_dest_mode.get() == "Excel Local":
            v = self._poa_dest_entry.get().strip()
            if v:
                self._poa_dest = Path(v)
            return self._poa_dest if self._poa_dest and self._poa_dest.exists() else None
        return None

    def _poa_scan(self):
        v = self._poa_origen_entry.get().strip()
        if v:
            self._poa_origin = Path(v)
        if not self._poa_origin or not self._poa_origin.is_dir():
            messagebox.showwarning("Carpeta no valida",
                "Selecciona una carpeta de origen valida.")
            return
        dest = self._poa_get_dest()
        if dest is None:
            messagebox.showwarning("Destino requerido",
                "Selecciona el archivo destino antes de escanear.")
            return
        if self._poa_scanning:
            return
        self._poa_scanning = True
        self._poa_scan_btn.configure(state="disabled", text="⏳  Escaneando...")
        self._poa_clear_tree()
        self._poa_log_msg("🔍 Escaneando...")
        threading.Thread(target=self._poa_scan_worker,
                         args=(dest,), daemon=True).start()

    def _poa_scan_worker(self, dest: Path):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(dest, read_only=True)
            sheets = wb.sheetnames; wb.close()
            files = sorted(f.name for f in self._poa_origin.iterdir()
                           if f.suffix.lower() in {".xlsx", ".xls"})
            if not files:
                self._qlog("poa", "⚠ No se encontraron .xlsx.", "warn")
                self._q.put(("poa_scan_done", []))
                return
            self._qlog("poa", f"  {len(files)} archivo(s) encontrado(s).")
            mapping = build_mapping(files, sheets, score_cutoff=60.0)
            self._poa_mapping = mapping
            self._q.put(("poa_scan_done", mapping))
        except Exception as exc:
            self._qlog("poa", f"❌ Error: {exc}", "error")
            self._q.put(("poa_scan_done", []))
        finally:
            self._poa_scanning = False

    def _poa_populate_tree(self, mapping):
        self._poa_clear_tree()
        matched = 0
        for m in mapping:
            ss = f"{m['score']:.1f} %" if m["matched"] else "—"
            if m["matched"] and m["score"] >= 90:
                est, tag = "✅ Listo", "ok"; matched += 1
            elif m["matched"]:
                est, tag = "⚠ Revisar", "warn"; matched += 1
            else:
                est, tag = "❌ Sin mapeo", "error"
            self._poa_tree.insert("", "end",
                values=("  " + m["file"], m["sheet"] or "—", ss, est),
                tags=(tag,))
        total = len(mapping)
        self._poa_file_lbl.configure(
            text=f"  {total} archivo(s) · {matched} mapeado(s) · {total-matched} sin coincidencia",
            text_color=SUCCESS if matched == total else WARNING)
        self._poa_status.configure(
            text=f"Escaneo: {matched}/{total} archivos mapeados.",
            text_color=SUCCESS if matched == total else WARNING)

    def _poa_clear_tree(self):
        for r in self._poa_tree.get_children():
            self._poa_tree.delete(r)

    def _poa_on_transfer(self):
        if self._poa_running:
            return
        v = self._poa_origen_entry.get().strip()
        if v:
            self._poa_origin = Path(v)
        if not self._poa_origin or not self._poa_origin.is_dir():
            messagebox.showwarning("Origen requerido",
                "Selecciona la carpeta de origen.")
            return
        dest = self._poa_get_dest()
        if dest is None and self._poa_dest_mode.get() == "Excel Local":
            messagebox.showwarning("Destino requerido",
                "Selecciona un Excel destino valido.")
            return
        if not self._poa_mapping:
            if messagebox.askyesno("Sin escaneo",
                "No escaneaste aun. ¿Escanear y transferir ahora?"):
                self._poa_scan()
                self.after(2000, lambda: self._poa_start(dest))
            return
        self._poa_start(dest)

    def _poa_start(self, dest):
        matched = [m for m in self._poa_mapping if m["matched"]]
        if not matched:
            messagebox.showwarning("Sin mapeo", "Escanea primero.")
            return
        self._poa_running = True
        self._poa_run_btn.configure(state="disabled", text="⏳  Transfiriendo...")
        self._poa_bar.set(0)
        self._poa_status.configure(text="Transfiriendo...", text_color=ACCENT)
        self._poa_log_msg("─" * 50)
        self._poa_log_msg(f"🚀 Mes: {self._month_var.get()} | {len(matched)} unidades")
        threading.Thread(target=self._poa_worker, args=(dest,), daemon=True).start()

    def _poa_worker(self, dest):
        month = self._month_var.get()
        mode  = self._poa_dest_mode.get()
        matched = [m for m in self._poa_mapping if m["matched"]]
        total = len(matched); cells = 0; failed = []
        for i, m in enumerate(matched):
            sheet = m["sheet"]
            fp    = self._poa_origin / m["file"]
            self._qlog("poa", f"  [{i+1}/{total}] {m['file']} → '{sheet}'")
            self._q.put(("poa_progress", i / total))
            try:
                data = extract_month_data(fp, month)
            except Exception as exc:
                self._qlog("poa", f"    ❌ Extraccion: {exc}", "error")
                failed.append(sheet); continue
            if mode == "Excel Local" and dest:
                try:
                    ok = write_month_data_to_excel(
                        dest, sheet, month, data,
                        create_backup=(i == 0))
                    if ok:
                        cells += len(data)
                        self._qlog("poa", f"    ✅ {len(data)} celdas escritas.")
                    else:
                        failed.append(sheet)
                except PermissionError:
                    self._qlog("poa",
                        "    ❌ ARCHIVO ABIERTO. Cierralo e intenta de nuevo.", "error")
                    failed.append(sheet); break
                except Exception as exc:
                    self._qlog("poa", f"    ❌ {exc}", "error")
                    failed.append(sheet)
            elif mode == "Google Sheets":
                url = self._poa_sheets_url.get().strip()
                creds = str(self._poa_creds) if self._poa_creds else "credentials.json"
                try:
                    from src.writers.sheets_writer import write_month_data_to_sheets
                    ok = write_month_data_to_sheets(url, sheet, month, data, creds)
                    if ok:
                        cells += len(data)
                        self._qlog("poa", "    ✅ Sheets actualizado.")
                    else:
                        failed.append(sheet)
                except Exception as exc:
                    self._qlog("poa", f"    ❌ {exc}", "error")
                    failed.append(sheet)
            self._q.put(("poa_progress", (i + 1) / total))
        self._q.put(("poa_done", {"total": total, "failed": failed,
                                   "cells": cells, "month": month, "mode": mode}))

    def _poa_on_done(self, r):
        self._poa_running = False
        ok = r["total"] - len(r["failed"])
        all_ok = not r["failed"]
        self._poa_bar.set(1.0)
        self._poa_run_btn.configure(state="normal", text="⚡  Transferir Datos")
        self._poa_status.configure(
            text=f"{'✅' if all_ok else '⚠'} {ok}/{r['total']} · {r['cells']} celdas",
            text_color=SUCCESS if all_ok else WARNING)
        self._poa_log_msg(f"{'✅' if all_ok else '⚠'} Transferencia finalizada.")
        lines = [
            f"Mes    : {r['month']}",
            f"Modo   : {r['mode']}",
            f"OK     : {ok}/{r['total']}",
            f"Celdas : {r['cells']}",
        ]
        if r["failed"]:
            lines.append(f"Errors : {', '.join(r['failed'])}")
        fn = messagebox.showinfo if all_ok else messagebox.showwarning
        fn("✅ Exitoso" if all_ok else "⚠ Parcial",
           ("¡Completado!\n\n" if all_ok else "Finalizado con errores.\n\n")
           + "\n".join(lines))

    def _poa_log_msg(self, msg: str, level: str = "info"):
        self._q.put(("poa_log", (msg, level)))

    def _poa_clear_log(self):
        self._poa_log.configure(state="normal")
        self._poa_log.delete("1.0", "end")
        self._poa_log.configure(state="disabled")
        self._poa_bar.set(0)
        self._poa_status.configure(text="Log limpiado.", text_color=TEXT_MUTED)

    def _poa_append_log(self, msg: str, level: str = "info"):
        color = {"info": "#8FD3A7", "warn": "#F5C842",
                 "error": "#FF6B6B"}.get(level, "#8FD3A7")
        self._poa_log.configure(state="normal", text_color=color)
        self._poa_log.insert("end", msg + "\n")
        self._poa_log.see("end")
        self._poa_log.configure(state="disabled")

    # ══════════════════════════════════════════════════════════════════════════
    #  PESTANA 2 — MIGRADOR UNIVERSAL
    # ══════════════════════════════════════════════════════════════════════════

    def _build_uni_tab(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True)
        scroll.columnconfigure(0, weight=1)

        self._uni_s1_origen(scroll)
        self._uni_s2_destino(scroll)
        self._uni_s3_modalidad(scroll)
        self._uni_s4_exec(scroll)

    # ── U1 Origen ─────────────────────────────────────────────────────────────
    def _uni_s1_origen(self, p):
        c = _card(p); c.pack(fill="x", padx=14, pady=(12, 5))
        c.columnconfigure(1, weight=1)
        _section_label(c, "01  ·  Origen — Archivo y Rango").grid(
            row=0, column=0, columnspan=3, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 4))

        # Fila archivo
        _muted(c, "Archivo:").grid(
            row=1, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")
        self._uni_src_entry = _entry(
            c, "Selecciona el archivo Excel origen (.xlsx)...")
        self._uni_src_entry.grid(row=1, column=1, sticky="ew", pady=(0, 6))
        _btn(c, "📂  Abrir", self._uni_select_src,
             width=110).grid(row=1, column=2, padx=(8, CARD_PAD), pady=(0, 6))

        # Fila hoja origen
        _muted(c, "Hoja:").grid(
            row=2, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")
        self._uni_src_sheet_var = ctk.StringVar(value="(sin archivo)")
        self._uni_src_sheet_menu = ctk.CTkOptionMenu(
            c, variable=self._uni_src_sheet_var, values=["(sin archivo)"],
            width=200, height=34,
            fg_color=ACCENT_DIM, button_color=ACCENT,
            font=ctk.CTkFont(size=12))
        self._uni_src_sheet_menu.grid(
            row=2, column=1, padx=(0, 6), pady=(0, 6), sticky="w")
        _btn(c, "🔄  Cargar hojas", self._uni_load_src_sheets,
             color="#2D4A6E", width=130).grid(
             row=2, column=2, padx=(8, CARD_PAD), pady=(0, 6))

        # Fila rango
        _muted(c, "Rango:").grid(
            row=3, column=0, padx=(CARD_PAD, 8), pady=(0, CARD_PAD), sticky="w")
        rf = ctk.CTkFrame(c, fg_color="transparent")
        rf.grid(row=3, column=1, columnspan=2, sticky="ew",
                pady=(0, CARD_PAD), padx=(0, CARD_PAD))
        rf.columnconfigure(0, weight=0)
        rf.columnconfigure(2, weight=1)
        self._uni_range_entry = _entry(rf, "ej: C3:C13", width=130)
        self._uni_range_entry.pack(side="left")
        self._uni_preview_btn = _btn(
            rf, "👁  Previsualizar", self._uni_preview_range,
            color="#374151", hover="#4B5563", width=140)
        self._uni_preview_btn.pack(side="left", padx=(10, 0))
        self._uni_range_info = _muted(rf, "")
        self._uni_range_info.pack(side="left", padx=(12, 0))

    # ── U2 Destino ────────────────────────────────────────────────────────────
    def _uni_s2_destino(self, p):
        c = _card(p); c.pack(fill="x", padx=14, pady=5)
        c.columnconfigure(1, weight=1)
        _section_label(c, "02  ·  Destino").grid(
            row=0, column=0, columnspan=3, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 6))

        self._uni_dest_mode = ctk.StringVar(value="Excel Local")
        ctk.CTkSegmentedButton(
            c, values=["Excel Local", "Google Sheets"],
            variable=self._uni_dest_mode,
            command=self._uni_dest_change,
            width=280, height=32, font=ctk.CTkFont(size=13),
            selected_color=PURPLE, selected_hover_color=PURPLE_DIM,
            unselected_color="#2D3A55",
        ).grid(row=1, column=0, columnspan=3, sticky="w",
               padx=CARD_PAD, pady=(0, 8))

        # Panel Excel destino
        self._uni_excel_pnl = ctk.CTkFrame(c, fg_color="transparent")
        self._uni_excel_pnl.grid(row=2, column=0, columnspan=3,
                                  sticky="ew", padx=0, pady=0)
        self._uni_excel_pnl.columnconfigure(1, weight=1)

        _muted(self._uni_excel_pnl, "Archivo:").grid(
            row=0, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")
        self._uni_dest_entry = _entry(
            self._uni_excel_pnl, "Archivo Excel destino (.xlsx)...")
        self._uni_dest_entry.grid(row=0, column=1, sticky="ew", pady=(0, 6))
        _btn(self._uni_excel_pnl, "📄  Seleccionar",
             self._uni_select_dest, width=130).grid(
             row=0, column=2, padx=(8, CARD_PAD), pady=(0, 6))

        _muted(self._uni_excel_pnl, "Hoja destino:").grid(
            row=1, column=0, padx=(CARD_PAD, 8), pady=(0, CARD_PAD), sticky="w")
        self._uni_dest_sheet_var = ctk.StringVar(value="(sin archivo)")
        self._uni_dest_sheet_menu = ctk.CTkOptionMenu(
            self._uni_excel_pnl,
            variable=self._uni_dest_sheet_var,
            values=["(sin archivo)"],
            width=200, height=32,
            fg_color=PURPLE_DIM, button_color=PURPLE,
            font=ctk.CTkFont(size=12))
        self._uni_dest_sheet_menu.grid(
            row=1, column=1, padx=(0, 6), pady=(0, CARD_PAD), sticky="w")
        _btn(self._uni_excel_pnl, "🔄  Cargar hojas",
             self._uni_load_dest_sheets,
             color="#2D4A6E", width=130).grid(
             row=1, column=2, padx=(8, CARD_PAD), pady=(0, CARD_PAD))

        # Panel Sheets destino (oculto)
        self._uni_sheets_pnl = ctk.CTkFrame(c, fg_color="transparent")
        self._uni_sheets_pnl.columnconfigure(1, weight=1)
        _muted(self._uni_sheets_pnl, "URL / ID:").grid(
            row=0, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")
        self._uni_sheets_url = _entry(
            self._uni_sheets_pnl, "https://docs.google.com/...")
        self._uni_sheets_url.grid(
            row=0, column=1, columnspan=2, sticky="ew",
            padx=(0, CARD_PAD), pady=(0, 6))
        _muted(self._uni_sheets_pnl, "Hoja destino:").grid(
            row=1, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")
        self._uni_sheets_tab = _entry(
            self._uni_sheets_pnl, "Nombre de la hoja", width=200)
        self._uni_sheets_tab.grid(
            row=1, column=1, sticky="w", pady=(0, 6))
        _muted(self._uni_sheets_pnl, "Credenciales:").grid(
            row=2, column=0, padx=(CARD_PAD, 8), pady=(0, CARD_PAD), sticky="w")
        self._uni_creds_entry = _entry(
            self._uni_sheets_pnl, "credentials.json")
        self._uni_creds_entry.grid(
            row=2, column=1, sticky="ew", pady=(0, CARD_PAD))
        _btn(self._uni_sheets_pnl, "🔑  Seleccionar",
             self._uni_select_creds, width=130).grid(
             row=2, column=2, padx=(8, CARD_PAD), pady=(0, CARD_PAD))

    # ── U3 Modalidad ──────────────────────────────────────────────────────────
    def _uni_s3_modalidad(self, p):
        c = _card(p); c.pack(fill="x", padx=14, pady=5)
        c.columnconfigure(0, weight=1)
        _section_label(c, "03  ·  Modalidad de pegado").grid(
            row=0, column=0, columnspan=4, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 6))

        self._uni_mode_var = ctk.StringVar(value=MODES[0])
        ctk.CTkSegmentedButton(
            c, values=list(MODES),
            variable=self._uni_mode_var,
            command=self._uni_mode_change,
            height=34, font=ctk.CTkFont(size=12),
            selected_color=PURPLE, selected_hover_color=PURPLE_DIM,
            unselected_color="#2D3A55",
        ).grid(row=1, column=0, columnspan=4, sticky="ew",
               padx=CARD_PAD, pady=(0, 10))

        # Panel A — Bloque Continuo
        self._uni_pnl_block = ctk.CTkFrame(c, fg_color="transparent")
        self._uni_pnl_block.grid(row=2, column=0, columnspan=4,
                                  sticky="ew", padx=CARD_PAD, pady=(0, CARD_PAD))
        _muted(self._uni_pnl_block, "Celda de inicio:").pack(side="left")
        self._uni_start_cell = _entry(
            self._uni_pnl_block, "ej: B5", width=100)
        self._uni_start_cell.pack(side="left", padx=(8, 0))
        self._uni_start_cell.insert(0, "A1")
        _muted(self._uni_pnl_block, "  — La matriz se pega a partir de esta celda.").pack(
            side="left", padx=(12, 0))

        # Panel B — Salto
        self._uni_pnl_stride = ctk.CTkFrame(c, fg_color="transparent")
        self._uni_pnl_stride.columnconfigure(1, weight=0)
        row_s = self._uni_pnl_stride

        _muted(row_s, "Celda inicio:").grid(
            row=0, column=0, padx=(0, 8), pady=(0, 4), sticky="w")
        self._uni_stride_start = _entry(row_s, "ej: C20", width=100)
        self._uni_stride_start.grid(row=0, column=1, padx=(0, 16), pady=(0, 4))
        self._uni_stride_start.insert(0, "A1")

        _muted(row_s, "Dirección:").grid(
            row=0, column=2, padx=(0, 8), pady=(0, 4), sticky="w")
        self._uni_dir_var = ctk.StringVar(value="Horizontal")
        ctk.CTkSegmentedButton(
            row_s, values=["Horizontal", "Vertical"],
            variable=self._uni_dir_var,
            width=200, height=30, font=ctk.CTkFont(size=12),
            selected_color=PURPLE, selected_hover_color=PURPLE_DIM,
            unselected_color="#2D3A55",
        ).grid(row=0, column=3, padx=(0, 16), pady=(0, 4))

        _muted(row_s, "Salto (N):").grid(
            row=0, column=4, padx=(0, 8), pady=(0, 4), sticky="w")
        self._uni_stride_n = _entry(row_s, "ej: 7", width=70)
        self._uni_stride_n.grid(row=0, column=5, pady=(0, 4))
        self._uni_stride_n.insert(0, "1")

        _muted(row_s,
               "\n  Ejemplo: inicio=C20, Horizontal, salto=7 → C20 → J20 → Q20 ...").grid(
            row=1, column=0, columnspan=6, sticky="w", pady=(2, 6))

        # Panel C — Lista
        self._uni_pnl_list = ctk.CTkFrame(c, fg_color="transparent")
        self._uni_pnl_list.columnconfigure(1, weight=1)
        _muted(self._uni_pnl_list, "Celdas:").grid(
            row=0, column=0, padx=(0, 8), pady=(0, 4), sticky="w")
        self._uni_cell_list = _entry(
            self._uni_pnl_list, "ej: C20, G20, K20, Q20")
        self._uni_cell_list.grid(row=0, column=1, sticky="ew", pady=(0, 4))
        _muted(self._uni_pnl_list,
               "Separa con comas. Cada valor del rango va a la celda correspondiente.").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 4))

        # Mostrar panel activo
        self._uni_mode_change(MODES[0])

    # ── U4 Ejecucion ──────────────────────────────────────────────────────────
    def _uni_s4_exec(self, p):
        c = _card(p); c.pack(fill="x", padx=14, pady=(5, 16))
        c.columnconfigure(0, weight=1)
        _section_label(c, "04  ·  Ejecucion").grid(
            row=0, column=0, sticky="w", padx=CARD_PAD, pady=(CARD_PAD, 6))
        self._uni_run_btn = ctk.CTkButton(
            c, text="🎯  Transferir Rango", height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=PURPLE, hover_color=PURPLE_DIM,
            corner_radius=CORNER_R, command=self._uni_on_transfer)
        self._uni_run_btn.grid(row=1, column=0, sticky="ew",
                               padx=CARD_PAD, pady=(0, 8))
        self._uni_bar = ctk.CTkProgressBar(c, height=8, corner_radius=4,
                                            progress_color=PURPLE)
        self._uni_bar.set(0)
        self._uni_bar.grid(row=2, column=0, sticky="ew",
                           padx=CARD_PAD, pady=(0, 6))
        self._uni_status = ctk.CTkLabel(
            c, text="Listo.", font=ctk.CTkFont(size=12),
            text_color=TEXT_MUTED, anchor="w")
        self._uni_status.grid(row=3, column=0, sticky="w",
                              padx=CARD_PAD, pady=(0, 6))
        self._uni_log = ctk.CTkTextbox(
            c, height=140,
            font=ctk.CTkFont(family="Cascadia Code", size=11),
            fg_color="#0D1117", text_color="#C4B5FD",
            corner_radius=8, wrap="word", state="disabled")
        self._uni_log.grid(row=4, column=0, sticky="ew",
                           padx=CARD_PAD, pady=(0, 6))
        _btn(c, "🗑  Limpiar log", self._uni_clear_log,
             color="transparent", hover="#2D3A55",
             width=120, height=28).grid(
             row=5, column=0, sticky="e", padx=CARD_PAD, pady=(0, CARD_PAD))

    # ── Logica Universal ──────────────────────────────────────────────────────
    def _uni_mode_change(self, val):
        for pnl in (self._uni_pnl_block,
                    self._uni_pnl_stride,
                    self._uni_pnl_list):
            pnl.grid_forget()
        if val == MODES[0]:       # Bloque Continuo
            self._uni_pnl_block.grid(row=2, column=0, columnspan=4,
                                      sticky="ew", padx=CARD_PAD,
                                      pady=(0, CARD_PAD))
        elif val == MODES[1]:     # Salto
            self._uni_pnl_stride.grid(row=2, column=0, columnspan=4,
                                       sticky="ew", padx=CARD_PAD,
                                       pady=(0, CARD_PAD))
        else:                     # Lista
            self._uni_pnl_list.grid(row=2, column=0, columnspan=4,
                                     sticky="ew", padx=CARD_PAD,
                                     pady=(0, CARD_PAD))

    def _uni_dest_change(self, val):
        if val == "Excel Local":
            self._uni_sheets_pnl.grid_forget()
            self._uni_excel_pnl.grid(row=2, column=0, columnspan=3,
                                      sticky="ew")
        else:
            self._uni_excel_pnl.grid_forget()
            self._uni_sheets_pnl.grid(row=2, column=0, columnspan=3,
                                       sticky="ew")

    def _uni_select_src(self):
        p = filedialog.askopenfilename(
            title="Selecciona el archivo origen",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if p:
            self._uni_src_file = Path(p)
            self._uni_src_entry.delete(0, "end")
            self._uni_src_entry.insert(0, str(self._uni_src_file))
            self._uni_load_src_sheets()

    def _uni_load_src_sheets(self):
        v = self._uni_src_entry.get().strip()
        if v:
            self._uni_src_file = Path(v)
        if not self._uni_src_file or not self._uni_src_file.exists():
            messagebox.showwarning("Archivo no encontrado",
                "Selecciona primero un archivo origen valido.")
            return
        try:
            names = get_sheet_names(self._uni_src_file)
            self._uni_src_sheet_menu.configure(values=names)
            self._uni_src_sheet_var.set(names[0])
            self._uni_log_msg(f"📋 Hojas origen: {names}")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _uni_select_dest(self):
        p = filedialog.askopenfilename(
            title="Selecciona el archivo destino",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if p:
            self._uni_dest_file = Path(p)
            self._uni_dest_entry.delete(0, "end")
            self._uni_dest_entry.insert(0, str(self._uni_dest_file))
            self._uni_load_dest_sheets()

    def _uni_load_dest_sheets(self):
        v = self._uni_dest_entry.get().strip()
        if v:
            self._uni_dest_file = Path(v)
        if not self._uni_dest_file or not self._uni_dest_file.exists():
            messagebox.showwarning("Archivo no encontrado",
                "Selecciona primero un archivo destino valido.")
            return
        try:
            names = get_sheet_names(self._uni_dest_file)
            self._uni_dest_sheet_menu.configure(values=names)
            self._uni_dest_sheet_var.set(names[0])
            self._uni_log_msg(f"📋 Hojas destino: {names}")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _uni_select_creds(self):
        p = filedialog.askopenfilename(
            title="credentials.json",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")])
        if p:
            self._uni_creds_entry.delete(0, "end")
            self._uni_creds_entry.insert(0, p)

    def _uni_preview_range(self):
        v_src  = self._uni_src_entry.get().strip()
        v_rng  = self._uni_range_entry.get().strip()
        v_shee = self._uni_src_sheet_var.get()
        if not v_src or not v_rng:
            messagebox.showwarning("Datos incompletos",
                "Selecciona archivo origen y escribe el rango.")
            return
        try:
            src = Path(v_src)
            sheet = v_shee if v_shee not in ("(sin archivo)", "") else None
            matrix = extract_range(src, v_rng, sheet)
            rows = len(matrix)
            cols = len(matrix[0]) if matrix else 0
            flat = [v for row in matrix for v in row]
            self._uni_range_info.configure(
                text=f"{rows}×{cols}  ({len(flat)} valores)",
                text_color="#C4B5FD")
            preview = str(flat[:8])
            if len(flat) > 8:
                preview = preview[:-1] + ", ...]"
            self._uni_log_msg(
                f"👁 Rango '{v_rng}': {rows} filas × {cols} cols  "
                f"→ primeros valores: {preview}")
        except Exception as exc:
            messagebox.showerror("Error al leer rango", str(exc))

    def _uni_on_transfer(self):
        if self._uni_running:
            return
        mode = self._uni_dest_mode.get()

        # Validar origen
        v_src = self._uni_src_entry.get().strip()
        if v_src:
            self._uni_src_file = Path(v_src)
        if not self._uni_src_file or not self._uni_src_file.exists():
            messagebox.showwarning("Origen requerido",
                "Selecciona un archivo Excel origen valido.")
            return

        v_rng = self._uni_range_entry.get().strip()
        if not v_rng:
            messagebox.showwarning("Rango requerido",
                "Escribe el rango de origen (ej: C3:C13).")
            return

        if mode == "Excel Local":
            v_dest = self._uni_dest_entry.get().strip()
            if v_dest:
                self._uni_dest_file = Path(v_dest)
            if not self._uni_dest_file or not self._uni_dest_file.exists():
                messagebox.showwarning("Destino requerido",
                    "Selecciona un archivo Excel destino valido.")
                return
            dest_sheet = self._uni_dest_sheet_var.get()
            if dest_sheet in ("(sin archivo)", ""):
                messagebox.showwarning("Hoja requerida",
                    "Selecciona la hoja destino.")
                return

        self._uni_running = True
        self._uni_run_btn.configure(state="disabled", text="⏳  Procesando...")
        self._uni_bar.set(0)
        self._uni_status.configure(text="Migrando rango...", text_color=PURPLE)

        params = {
            "src_file":   self._uni_src_file,
            "src_range":  v_rng,
            "src_sheet":  (self._uni_src_sheet_var.get()
                           if self._uni_src_sheet_var.get()
                           not in ("(sin archivo)", "") else None),
            "dest_mode":  mode,
            "paste_mode": self._uni_mode_var.get(),
            "start_cell": self._uni_start_cell.get().strip() or "A1",
            "stride_start": self._uni_stride_start.get().strip() or "A1",
            "direction":  self._uni_dir_var.get(),
            "stride":     self._uni_stride_n.get().strip(),
            "cell_list":  self._uni_cell_list.get().strip(),
        }
        if mode == "Excel Local":
            params["dest_file"]  = self._uni_dest_file
            params["dest_sheet"] = self._uni_dest_sheet_var.get()
        else:
            params["sheets_url"]   = self._uni_sheets_url.get().strip()
            params["sheets_tab"]   = self._uni_sheets_tab.get().strip()
            params["sheets_creds"] = self._uni_creds_entry.get().strip()

        threading.Thread(target=self._uni_worker,
                         args=(params,), daemon=True).start()

    def _uni_worker(self, p: dict):
        try:
            # Parsear stride
            try:
                stride = max(1, int(p["stride"]))
            except (ValueError, TypeError):
                stride = 1

            # Celda de inicio segun modo
            start_cell = (p["stride_start"] if p["paste_mode"] == MODES[1]
                          else p["start_cell"])

            self._qlog("uni",
                f"📐 Rango origen : {p['src_range']}")
            self._qlog("uni",
                f"🎯 Modalidad    : {p['paste_mode']}")

            if p["dest_mode"] == "Excel Local":
                dest_sheet = p["dest_sheet"]
                self._qlog("uni",
                    f"📄 Destino      : {Path(p['dest_file']).name} → '{dest_sheet}'")

                # Pre-visualizacion de celdas destino (usa el mismo parser que el backend)
                if p["paste_mode"] == MODES[2]:  # Lista
                    try:
                        from src.core.range_migrator import _expand_cell_tokens
                        cells_preview = _expand_cell_tokens(p["cell_list"])
                        preview_str = ", ".join(cells_preview[:12])
                        if len(cells_preview) > 12:
                            preview_str += f" ... (+{len(cells_preview)-12} mas)"
                        self._qlog("uni",
                            f"📌 Pegando en   : {preview_str}  "
                            f"({len(cells_preview)} celdas)")
                    except ValueError as e:
                        self._qlog("uni", f"⚠ Lista invalida: {e}", "warn")
                elif p["paste_mode"] == MODES[1]:  # Stride
                    self._qlog("uni",
                        f"📌 Inicio={start_cell} dir={p['direction']} salto={stride}")
                else:  # Bloque
                    self._qlog("uni",
                        f"📌 Bloque desde : {start_cell}")

                self._q.put(("uni_progress", 0.3))
                result = migrate_range(
                    src_file   = p["src_file"],
                    src_range  = p["src_range"],
                    dest_file  = p["dest_file"],
                    dest_sheet = dest_sheet,
                    mode       = p["paste_mode"],
                    src_sheet  = p["src_sheet"],
                    start_cell = start_cell,
                    direction  = p["direction"],
                    stride     = stride,
                    cell_list  = p["cell_list"],
                    create_backup = True,
                )
                self._q.put(("uni_progress", 1.0))

                # Log detallado: coord=valor para cada celda escrita
                detail = result.get("detail", [])
                dest_name = Path(p["dest_file"]).name
                preview = ", ".join(detail[:8])
                if len(detail) > 8:
                    preview += f" ... (+{len(detail)-8} mas)"
                self._qlog("uni",
                    f"💾 Guardado en {dest_name} → '{dest_sheet}'")
                self._qlog("uni",
                    f"   Valores: {preview}")

                self._q.put(("uni_done", result))

            else:  # Google Sheets
                self._qlog("uni", "⚠ Google Sheets: extrayendo datos...", "warn")
                matrix = extract_range(
                    p["src_file"], p["src_range"], p["src_sheet"])
                flat = [v for row in matrix for v in row]
                self._qlog("uni", f"  {len(flat)} valores extraidos.")
                self._q.put(("uni_progress", 0.5))
                try:
                    from src.writers.sheets_writer import (
                        write_month_data_to_sheets)
                    self._qlog("uni",
                        "⚠ Google Sheets: usa write_month_data_to_sheets (modo POA).\n"
                        "  Para rangos universales en Sheets implementa un writer dedicado.",
                        "warn")
                except ImportError:
                    self._qlog("uni",
                        "❌ gspread no instalado. Ejecuta: pip install gspread google-auth",
                        "error")
                self._q.put(("uni_done", {"written": 0, "cells": [],
                                           "backup": None, "matrix": matrix}))

        except Exception as exc:
            self._qlog("uni", f"❌ Error: {exc}", "error")
            self._q.put(("uni_done", None))
        finally:
            self._uni_running = False

    def _uni_on_done(self, result: dict | None):
        self._uni_run_btn.configure(state="normal", text="🎯  Transferir Rango")
        if result is None:
            self._uni_status.configure(
                text="❌ Error en la transferencia.", text_color=DANGER)
            return
        n = result.get("written", 0)
        detail = result.get("detail", result.get("cells", []))
        self._uni_bar.set(1.0)
        self._uni_status.configure(
            text=f"✅ {n} celdas escritas.", text_color=SUCCESS)
        # Mostrar pares coord=valor reales (Bug 2: confirmacion explicita)
        if detail:
            for chunk_start in range(0, min(len(detail), 24), 6):
                chunk = ", ".join(detail[chunk_start:chunk_start + 6])
                self._uni_log_msg(f"   {chunk}")
            if len(detail) > 24:
                self._uni_log_msg(
                    f"   ... (+{len(detail)-24} celdas mas)")
        messagebox.showinfo(
            "✅ Transferencia completada",
            f"Rango migrado correctamente.\n\n"
            f"Celdas escritas : {n}\n"
            f"Modalidad       : {self._uni_mode_var.get()}\n"
            f"Backup          : {Path(result['backup']).name if result.get('backup') else 'N/A'}",
        )

    def _uni_log_msg(self, msg: str, level: str = "info"):
        self._q.put(("uni_log", (msg, level)))

    def _uni_clear_log(self):
        self._uni_log.configure(state="normal")
        self._uni_log.delete("1.0", "end")
        self._uni_log.configure(state="disabled")
        self._uni_bar.set(0)
        self._uni_status.configure(text="Log limpiado.", text_color=TEXT_MUTED)

    def _uni_append_log(self, msg: str, level: str = "info"):
        color = {"info": "#C4B5FD", "warn": "#F5C842",
                 "error": "#FF6B6B"}.get(level, "#C4B5FD")
        self._uni_log.configure(state="normal", text_color=color)
        self._uni_log.insert("end", msg + "\n")
        self._uni_log.see("end")
        self._uni_log.configure(state="disabled")

    # ══════════════════════════════════════════════════════════════════════════
    #  COLA DE MENSAJES COMPARTIDA
    # ══════════════════════════════════════════════════════════════════════════

    def _qlog(self, tab: str, msg: str, level: str = "info"):
        """Encola un mensaje de log desde cualquier hilo."""
        self._q.put((f"{tab}_log", (msg, level)))

    def _poll(self):
        try:
            while True:
                kind, data = self._q.get_nowait()

                # ── POA ──────────────────────────────────────────
                if kind == "poa_log":
                    self._poa_append_log(*data)
                elif kind == "poa_progress":
                    self._poa_bar.set(data)
                elif kind == "poa_scan_done":
                    self._poa_populate_tree(data)
                    self._poa_scan_btn.configure(
                        state="normal", text="🔍  Escanear")
                elif kind == "poa_done":
                    self._poa_on_done(data)

                # ── Universal ────────────────────────────────────
                elif kind == "uni_log":
                    self._uni_append_log(*data)
                elif kind == "uni_progress":
                    self._uni_bar.set(data)
                elif kind == "uni_done":
                    self._uni_on_done(data)

        except queue.Empty:
            pass
        finally:
            self.after(100, self._poll)
