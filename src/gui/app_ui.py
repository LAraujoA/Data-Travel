
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
from src.writers.sheets_writer import (
    write_range_to_sheets, get_sheet_tabs,
    _GSPREAD_AVAILABLE as GSPREAD_OK,
)
from src.core.range_migrator import (
    extract_range, extract_multi_range, get_sheet_names, match_dest_sheet,
    migrate_range, _expand_cell_tokens, MODES,
    write_block, write_stride, write_cell_list,
)

# ── TEMA ──────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

ACCENT     = "#3B8ED0"
ACCENT_DIM = "#1F538A"
SUCCESS    = "#27AE60"
WARNING    = "#F39C12"
DANGER     = "#E74C3C"

BG_WINDOW   = ("#F1F5F9", "#0B1120")
BG_CARD     = ("#FFFFFF", "#1E293B")
BG_DARK     = ("#F8FAFC", "#0F172A")  # Alias for compatibility if missed
CARD_BORDER = ("#CBD5E1", "#334155")
TEXT_MAIN   = ("#0F172A", "#F8FAFC")
TEXT_MUTED  = ("#64748B", "#94A3B8")
ENTRY_BG    = ("#F1F5F9", "#0F172A")
ENTRY_TEXT  = ("#0F172A", "#FFFFFF")
TEXT_DIM    = ("#94A3B8", "#5A6480")
TAB_BG      = ("#F8FAFC", "#0F172A")
TAB_UNSEL   = ("#FFFFFF", "#1E293B")
TAB_HOVER   = ("#F1F5F9", "#2D3A55")

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
CORNER_R  = 8


# ── HELPERS ───────────────────────────────────────────────────────────────────
def _card(parent, **kw) -> ctk.CTkFrame:
    d = dict(bg_color="transparent", fg_color=BG_CARD, border_color=CARD_BORDER, border_width=1, corner_radius=CORNER_R)
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
              font=ctk.CTkFont(size=font_size),
              fg_color=ENTRY_BG, text_color=ENTRY_TEXT)
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



# ══════════════════════════════════════════════════════════════════════════════
#  DIALOGO DE CONFIRMACION DE MAPEO
# ══════════════════════════════════════════════════════════════════════════════

OMIT_LABEL = "-- Omitir --"


class MappingDialog(ctk.CTkToplevel):
    """
    Modal de confirmacion de mapeo origen -> hoja destino.

    Muestra una fila por cada origen con:
      - Checkbox (habilitado por defecto si score >= 85, deshabilitado si sin match)
      - Nombre del origen
      - CTkOptionMenu con todas las hojas + OMIT_LABEL, preseleccionando el match
      - Etiqueta de confianza coloreada
    Retorna self.result: list[dict] | None
    """

    def __init__(self, master, rows: list[dict], dest_sheets: list[str]):
        """
        Parameters
        ----------
        rows : list[dict]
            Cada dict: {"label": str, "fp": Path, "src_sheet": str|None,
                        "dest_sh": str|None, "score": float}
        dest_sheets : list[str]
            Hojas del libro destino.
        """
        super().__init__(master)
        self.result: list[dict] | None = None
        self._rows      = rows
        self._dest_opts = [OMIT_LABEL] + list(dest_sheets)

        self.title("Confirmar Mapeo  —  Data-Travel")
        self.geometry("860x520")
        self.minsize(700, 380)
        self.configure(fg_color=BG_WINDOW)
        self.resizable(True, True)

        # Modal: captura el foco
        self.transient(master)
        self.grab_set()

        self._check_vars: list[ctk.BooleanVar]  = []
        self._dest_vars:  list[ctk.StringVar]   = []

        self._build()

    # ── Construccion ──────────────────────────────────────────────────────────
    def _build(self):
        # Titulo
        hdr = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=54)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        ctk.CTkLabel(
            hdr, text="Revision y Confirmacion de Mapeo",
            font=ctk.CTkFont(size=16, weight="bold"), text_color=TEXT_MAIN,
        ).pack(side="left", padx=18, pady=12)
        ctk.CTkLabel(
            hdr,
            text="Ajusta la hoja destino de cada origen antes de transferir.",
            font=ctk.CTkFont(size=11), text_color=TEXT_MUTED,
        ).pack(side="left", padx=(0, 16))

        # Cabecera de columnas
        col_hdr = ctk.CTkFrame(self, fg_color="#1A2035", corner_radius=0, height=30)
        col_hdr.pack(fill="x", padx=0)
        col_hdr.pack_propagate(False)
        for txt, w in [("", 40), ("Origen", 260), ("Hoja destino", 230), ("Confianza", 120)]:
            ctk.CTkLabel(
                col_hdr, text=txt.upper(),
                font=ctk.CTkFont(size=9, weight="bold"),
                text_color=ACCENT, width=w, anchor="w",
            ).pack(side="left", padx=(8, 0))

        # Filas desplazables
        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        for i, row in enumerate(self._rows):
            self._add_row(scroll, i, row)

        # Pie con botones
        foot = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=54)
        foot.pack(fill="x", side="bottom")
        foot.pack_propagate(False)
        info_lbl = ctk.CTkLabel(
            foot,
            text=f"{len(self._rows)} origen(s)  |  Solo se transfieren las filas activadas.",
            font=ctk.CTkFont(size=11), text_color=TEXT_MUTED,
        )
        info_lbl.pack(side="left", padx=16)
        _btn(foot, "Cancelar", self._cancel,
             color="#374151", hover="#4B5563", width=110, height=36).pack(
             side="right", padx=(8, 16), pady=9)
        _btn(foot, "Confirmar y Transferir", self._confirm,
             color=PURPLE, hover=PURPLE_DIM, width=190, height=36).pack(
             side="right", padx=(0, 4), pady=9)

    def _add_row(self, parent, i: int, row: dict):
        score     = row["score"]
        matched   = row["dest_sh"] is not None
        bg_color  = "#161C2E" if i % 2 == 0 else "#1A2238"

        f = ctk.CTkFrame(parent, bg_color="transparent", fg_color=bg_color, corner_radius=6, height=42)
        f.pack(fill="x", padx=6, pady=2)
        f.pack_propagate(False)

        # Checkbox
        chk_var = ctk.BooleanVar(value=matched)
        self._check_vars.append(chk_var)
        ctk.CTkCheckBox(
            f, text="", variable=chk_var, width=28,
            fg_color=PURPLE, hover_color=PURPLE_DIM, checkmark_color="white",
        ).pack(side="left", padx=(8, 0))

        # Nombre de origen
        ctk.CTkLabel(
            f, text=row["label"],
            font=ctk.CTkFont(size=12), text_color=TEXT_MAIN,
            width=255, anchor="w",
        ).pack(side="left", padx=(8, 0))

        # Dropdown de hoja destino
        dest_var = ctk.StringVar(value=row["dest_sh"] if matched else OMIT_LABEL)
        self._dest_vars.append(dest_var)
        ctk.CTkOptionMenu(
            f,
            variable=dest_var,
            values=self._dest_opts,
            width=220, height=30,
            fg_color=ACCENT_DIM if matched else "#374151",
            button_color=ACCENT if matched else "#4B5563",
            font=ctk.CTkFont(size=11),
        ).pack(side="left", padx=(8, 0))

        # Etiqueta de confianza
        if matched:
            if score >= 90:
                conf_txt   = f"Score {score:.0f}%"
                conf_color = "#27AE60"
            elif score >= 75:
                conf_txt   = f"Score {score:.0f}% (revisar)"
                conf_color = "#F39C12"
            else:
                conf_txt   = f"Score {score:.0f}% (bajo)"
                conf_color = "#E74C3C"
        else:
            conf_txt   = "Sin match"
            conf_color = "#5A6480"

        ctk.CTkLabel(
            f, text=conf_txt,
            font=ctk.CTkFont(size=10), text_color=conf_color,
            width=115, anchor="w",
        ).pack(side="left", padx=(10, 0))

    # ── Acciones ──────────────────────────────────────────────────────────────
    def _confirm(self):
        confirmed = []
        for i, row in enumerate(self._rows):
            if not self._check_vars[i].get():
                continue
            dest_sel = self._dest_vars[i].get()
            if dest_sel == OMIT_LABEL:
                continue
            confirmed.append({
                "fp":        row["fp"],
                "src_sheet": row["src_sheet"],
                "dest_sh":   dest_sel,
                "label":     row["label"],
            })
        self.result = confirmed
        self.grab_release()
        self.destroy()

    def _cancel(self):
        self.result = None
        self.grab_release()
        self.destroy()


# ── CLASE PRINCIPAL ───────────────────────────────────────────────────────────
class DataTravelApp(ctk.CTk):
    """Ventana principal de Data-Travel."""

    def __init__(self):
        super().__init__()
        self.title("Data-Travel  —  Migrador de Reportes POA")
        self.geometry(f"{WINDOW_W}x{WINDOW_H}")
        self.minsize(800, 720)
        self.configure(fg_color=BG_WINDOW)
        self.resizable(True, True)

        # Estado POA
        self._poa_mapping:   list[dict] = []
        self._poa_origin:    Optional[Path] = None
        self._poa_dest:      Optional[Path] = None
        self._poa_creds:     Optional[Path] = None
        self._poa_scanning   = False
        self._poa_running    = False

        # Estado Migrador Universal
        self._uni_src_files: list[Path] = []     # uno o varios archivos origen
        self._uni_dest_file: Optional[Path] = None
        self._uni_sheet_vars: dict = {}           # {sheet_name: BooleanVar}
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
                     text_color=TEXT_MAIN).pack(anchor="w")
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
            fg_color=TAB_BG,
            segmented_button_fg_color=BG_CARD,
            segmented_button_selected_color=ACCENT,
            segmented_button_selected_hover_color=ACCENT_DIM,
            segmented_button_unselected_color=TAB_UNSEL,
            segmented_button_unselected_hover_color=TAB_HOVER,
            text_color=TEXT_MAIN,
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

        tf = ctk.CTkFrame(c, bg_color="transparent", fg_color="#1A2035", corner_radius=8)
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
        _section_label(c, "01  ·  Origen — Archivos y Rango").grid(
            row=0, column=0, columnspan=3, sticky="w",
            padx=CARD_PAD, pady=(CARD_PAD, 4))

        # ── Modo de seleccion (1 archivo / multiples) ──────────────────────
        self._uni_src_mode = ctk.StringVar(value="Archivo Unico")
        ctk.CTkSegmentedButton(
            c, values=["Archivo Unico", "Multiples Archivos"],
            variable=self._uni_src_mode,
            command=self._uni_src_mode_change,
            height=32, font=ctk.CTkFont(size=12),
            selected_color=PURPLE, selected_hover_color=PURPLE_DIM,
            unselected_color="#2D3A55",
        ).grid(row=1, column=0, columnspan=3, sticky="w",
               padx=CARD_PAD, pady=(0, 8))

        # ── Panel Archivo Unico ────────────────────────────────────────────
        self._uni_single_pnl = ctk.CTkFrame(c, fg_color="transparent")
        self._uni_single_pnl.columnconfigure(1, weight=1)

        _muted(self._uni_single_pnl, "Archivo:").grid(
            row=0, column=0, padx=(0, 8), pady=(0, 6), sticky="w")
        self._uni_src_entry = _entry(
            self._uni_single_pnl,
            "Selecciona el archivo Excel origen (.xlsx)...")
        self._uni_src_entry.grid(row=0, column=1, sticky="ew", pady=(0, 6))
        bf1 = ctk.CTkFrame(self._uni_single_pnl, fg_color="transparent")
        bf1.grid(row=0, column=2, padx=(8, 0), pady=(0, 6))
        _btn(bf1, "📂  Abrir", self._uni_select_single,
             width=110).pack(side="left", padx=(0, 4))
        _btn(bf1, "🔄  Hojas", self._uni_load_src_sheets,
             color="#2D4A6E", width=90).pack(side="left")

        # Sub-panel de hojas con checkboxes (ocupa fila 1)
        _muted(self._uni_single_pnl, "Hojas:").grid(
            row=1, column=0, padx=(0, 8), pady=(0, 4), sticky="nw")
        self._uni_sheets_frame_outer = ctk.CTkFrame(
            self._uni_single_pnl, fg_color="#12192A",
            corner_radius=8, height=90)
        self._uni_sheets_frame_outer.grid(
            row=1, column=1, columnspan=2, sticky="ew",
            pady=(0, 6))
        self._uni_sheets_frame_outer.grid_propagate(False)
        self._uni_sheets_scroll = ctk.CTkScrollableFrame(
            self._uni_sheets_frame_outer, bg_color="transparent",
            fg_color=("#F8FAFC", "#0F172A"), border_width=1,
            border_color=("#E2E8F0", "#334155"), corner_radius=6,
            orientation="horizontal", height=70)
        self._uni_sheets_scroll.pack(fill="both", expand=True, padx=6, pady=4)

        sf_btn = ctk.CTkFrame(self._uni_single_pnl, fg_color="transparent")
        sf_btn.grid(row=2, column=1, columnspan=2, sticky="w", pady=(0, 6))
        _btn(sf_btn, "Todas", lambda: self._uni_check_all(True),
             color="#374151", hover="#4B5563", width=70, height=28).pack(
             side="left", padx=(0, 4))
        _btn(sf_btn, "Ninguna", lambda: self._uni_check_all(False),
             color="#374151", hover="#4B5563", width=80, height=28).pack(
             side="left")
        self._uni_sheets_hint = _muted(sf_btn, "  (carga hojas primero)")
        self._uni_sheets_hint.pack(side="left", padx=(8, 0))

        self._uni_single_pnl.grid(
            row=2, column=0, columnspan=3, sticky="ew",
            padx=CARD_PAD, pady=(0, 6))

        # ── Panel Multiples Archivos ───────────────────────────────────────
        self._uni_multi_pnl = ctk.CTkFrame(c, fg_color="transparent")
        self._uni_multi_pnl.columnconfigure(1, weight=1)

        _muted(self._uni_multi_pnl, "Archivos:").grid(
            row=0, column=0, padx=(0, 8), pady=(0, 6), sticky="nw")
        self._uni_multi_list = ctk.CTkTextbox(
            self._uni_multi_pnl, height=70,
            font=ctk.CTkFont(family="Cascadia Code", size=10),
            fg_color="#12192A", text_color="#8FD3A7",
            corner_radius=8, state="disabled")
        self._uni_multi_list.grid(
            row=0, column=1, sticky="ew", pady=(0, 6))
        _btn(self._uni_multi_pnl, "📂  Seleccionar archivos",
             self._uni_select_multi,
             width=175).grid(row=0, column=2, padx=(8, 0), pady=(0, 6), sticky="n")
        self._uni_multi_count = _muted(self._uni_multi_pnl, "0 archivos seleccionados")
        self._uni_multi_count.grid(
            row=1, column=1, sticky="w", pady=(0, 6))

        # (no se muestra hasta que se active "Multiples Archivos")

        # ── Rango de origen ────────────────────────────────────────────────
        _muted(c, "Rango:").grid(
            row=3, column=0, padx=(CARD_PAD, 8), pady=(0, CARD_PAD), sticky="w")
        rf = ctk.CTkFrame(c, fg_color="transparent")
        rf.grid(row=3, column=1, columnspan=2, sticky="ew",
                pady=(0, CARD_PAD), padx=(0, CARD_PAD))
        self._uni_range_entry = _entry(
            rf, "ej: D6:D15  o  D6:D10, G6:G10  o  D6, D7, D8", width=270)
        self._uni_range_entry.pack(side="left")
        self._uni_preview_btn = _btn(
            rf, "👁  Vista previa", self._uni_preview_range,
            color="#374151", hover="#4B5563", width=135)
        self._uni_preview_btn.pack(side="left", padx=(10, 0))
        self._uni_range_info = _muted(rf, "")
        self._uni_range_info.pack(side="left", padx=(10, 0))

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

        # ── Panel Excel ────────────────────────────────────────────────────
        self._uni_excel_pnl = ctk.CTkFrame(c, fg_color="transparent")
        self._uni_excel_pnl.columnconfigure(1, weight=1)
        self._uni_excel_pnl.grid(
            row=2, column=0, columnspan=3, sticky="ew", padx=0, pady=0)

        _muted(self._uni_excel_pnl, "Archivo:").grid(
            row=0, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")
        self._uni_dest_entry = _entry(
            self._uni_excel_pnl, "Archivo Excel destino (.xlsx)...")
        self._uni_dest_entry.grid(row=0, column=1, sticky="ew", pady=(0, 6))
        bf_d = ctk.CTkFrame(self._uni_excel_pnl, fg_color="transparent")
        bf_d.grid(row=0, column=2, padx=(8, CARD_PAD), pady=(0, 6))
        _btn(bf_d, "📄  Seleccionar",
             self._uni_select_dest, width=130).pack(side="left", padx=(0, 4))
        _btn(bf_d, "🔄  Hojas",
             self._uni_load_dest_sheets, color="#2D4A6E", width=90).pack(side="left")

        # Fila: Logica de hoja destino (nueva)
        _muted(self._uni_excel_pnl, "Mapeo:").grid(
            row=1, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="nw")

        self._uni_dest_map_mode = ctk.StringVar(value="Mismo nombre")
        map_frame = ctk.CTkFrame(self._uni_excel_pnl, fg_color="transparent")
        map_frame.grid(row=1, column=1, columnspan=2, sticky="ew",
                       padx=(0, CARD_PAD), pady=(0, 6))

        ctk.CTkRadioButton(
            map_frame,
            text="Mapear a pestana con el mismo nombre",
            variable=self._uni_dest_map_mode,
            value="Mismo nombre",
            command=self._uni_map_mode_change,
            font=ctk.CTkFont(size=12),
            fg_color=PURPLE, hover_color=PURPLE_DIM,
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        ctk.CTkRadioButton(
            map_frame,
            text="Consolidar en una sola pestana:",
            variable=self._uni_dest_map_mode,
            value="Consolidar",
            command=self._uni_map_mode_change,
            font=ctk.CTkFont(size=12),
            fg_color=PURPLE, hover_color=PURPLE_DIM,
        ).grid(row=1, column=0, sticky="w")

        self._uni_dest_sheet_var = ctk.StringVar(value="")
        self._uni_dest_sheet_menu = ctk.CTkOptionMenu(
            map_frame,
            variable=self._uni_dest_sheet_var,
            values=["(carga el destino primero)"],
            width=210, height=30,
            fg_color=PURPLE_DIM, button_color=PURPLE,
            font=ctk.CTkFont(size=11),
            state="disabled",
        )
        self._uni_dest_sheet_menu.grid(
            row=1, column=1, padx=(10, 0), sticky="w")

        # ── Panel Sheets ───────────────────────────────────────────────────
        self._uni_sheets_pnl = ctk.CTkFrame(c, fg_color="transparent")
        self._uni_sheets_pnl.columnconfigure(1, weight=1)

        # Fila 0: URL / ID
        _muted(self._uni_sheets_pnl, "URL / ID:").grid(
            row=0, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")
        self._uni_sheets_url = _entry(
            self._uni_sheets_pnl, "https://docs.google.com/spreadsheets/d/...")
        self._uni_sheets_url.grid(
            row=0, column=1, columnspan=2, sticky="ew",
            padx=(0, CARD_PAD), pady=(0, 6))

        # Fila 1: Credenciales + botones
        _muted(self._uni_sheets_pnl, "Credenciales:").grid(
            row=1, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="w")
        self._uni_creds_entry = _entry(
            self._uni_sheets_pnl, "credentials.json")
        self._uni_creds_entry.grid(
            row=1, column=1, sticky="ew", pady=(0, 6))
        _bf_gs = ctk.CTkFrame(self._uni_sheets_pnl, fg_color="transparent")
        _bf_gs.grid(row=1, column=2, padx=(8, CARD_PAD), pady=(0, 6))
        _btn(_bf_gs, "Cargar hojas",
             self._uni_load_sheets_tabs, color="#2D4A6E",
             width=120, height=30).pack(side="left", padx=(0, 4))
        _btn(_bf_gs, "Seleccionar",
             self._uni_select_creds, color="#374151", hover="#4B5563",
             width=100, height=30).pack(side="left")

        # Fila 2: Modo de mapeo (mismo esquema que Excel Local)
        _muted(self._uni_sheets_pnl, "Mapeo:").grid(
            row=2, column=0, padx=(CARD_PAD, 8), pady=(0, 6), sticky="nw")
        self._uni_gs_map_mode = ctk.StringVar(value="Mismo nombre")
        _gs_map_frame = ctk.CTkFrame(self._uni_sheets_pnl, fg_color="transparent")
        _gs_map_frame.grid(row=2, column=1, columnspan=2, sticky="ew",
                           padx=(0, CARD_PAD), pady=(0, 6))
        ctk.CTkRadioButton(
            _gs_map_frame,
            text="Mapear a pestana con el mismo nombre (abre confirmacion)",
            variable=self._uni_gs_map_mode,
            value="Mismo nombre",
            command=self._uni_gs_map_mode_change,
            font=ctk.CTkFont(size=12),
            fg_color=PURPLE, hover_color=PURPLE_DIM,
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))
        ctk.CTkRadioButton(
            _gs_map_frame,
            text="Consolidar en una sola hoja:",
            variable=self._uni_gs_map_mode,
            value="Consolidar",
            command=self._uni_gs_map_mode_change,
            font=ctk.CTkFont(size=12),
            fg_color=PURPLE, hover_color=PURPLE_DIM,
        ).grid(row=1, column=0, sticky="w")
        self._uni_sheets_tab_var = ctk.StringVar(value="")
        self._uni_sheets_tab_menu = ctk.CTkOptionMenu(
            _gs_map_frame,
            variable=self._uni_sheets_tab_var,
            values=["(carga hojas primero)"],
            width=210, height=30,
            fg_color=PURPLE_DIM, button_color=PURPLE,
            font=ctk.CTkFont(size=11),
            state="disabled",
        )
        self._uni_sheets_tab_menu.grid(row=1, column=1, padx=(10, 0), sticky="w")
        self._uni_sheets_tabs_hint = _muted(_gs_map_frame, "  (cargar primero)")
        self._uni_sheets_tabs_hint.grid(row=2, column=0, columnspan=2,
                                        sticky="w", pady=(2, 0))

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

        # ── Panel A: Bloque Continuo ───────────────────────────────────────
        self._uni_pnl_block = ctk.CTkFrame(c, fg_color="transparent")
        self._uni_pnl_block.grid(
            row=2, column=0, columnspan=4, sticky="ew",
            padx=CARD_PAD, pady=(0, CARD_PAD))
        _muted(self._uni_pnl_block, "Celda de inicio:").pack(side="left")
        self._uni_start_cell = _entry(
            self._uni_pnl_block,
            placeholder="ej: B5",
            width=100)
        self._uni_start_cell.pack(side="left", padx=(8, 0))
        _muted(self._uni_pnl_block,
               "  — La matriz se pega a partir de esta celda.").pack(
               side="left", padx=(12, 0))

        # ── Panel B: Salto ─────────────────────────────────────────────────
        self._uni_pnl_stride = ctk.CTkFrame(c, fg_color="transparent")

        _muted(self._uni_pnl_stride, "Celda inicio:").grid(
            row=0, column=0, padx=(0, 8), pady=(0, 4), sticky="w")
        self._uni_stride_start = _entry(
            self._uni_pnl_stride, placeholder="ej: C20", width=100)
        self._uni_stride_start.grid(row=0, column=1, padx=(0, 16), pady=(0, 4))

        _muted(self._uni_pnl_stride, "Direccion:").grid(
            row=0, column=2, padx=(0, 8), pady=(0, 4), sticky="w")
        self._uni_dir_var = ctk.StringVar(value="Horizontal")
        ctk.CTkSegmentedButton(
            self._uni_pnl_stride,
            values=["Horizontal", "Vertical"],
            variable=self._uni_dir_var,
            width=200, height=30, font=ctk.CTkFont(size=12),
            selected_color=PURPLE, selected_hover_color=PURPLE_DIM,
            unselected_color="#2D3A55",
        ).grid(row=0, column=3, padx=(0, 16), pady=(0, 4))

        _muted(self._uni_pnl_stride, "Salto (N):").grid(
            row=0, column=4, padx=(0, 8), pady=(0, 4), sticky="w")
        self._uni_stride_n = _entry(
            self._uni_pnl_stride, placeholder="ej: 7", width=70)
        self._uni_stride_n.grid(row=0, column=5, pady=(0, 4))

        _muted(self._uni_pnl_stride,
               "Ejemplo: inicio=C20, Horizontal, salto=7  ->  C20 -> J20 -> Q20 ...").grid(
            row=1, column=0, columnspan=6, sticky="w", pady=(2, 6))

        # ── Panel C: Lista ─────────────────────────────────────────────────
        self._uni_pnl_list = ctk.CTkFrame(c, fg_color="transparent")
        self._uni_pnl_list.columnconfigure(1, weight=1)
        _muted(self._uni_pnl_list, "Celdas:").grid(
            row=0, column=0, padx=(0, 8), pady=(0, 4), sticky="w")
        self._uni_cell_list = _entry(
            self._uni_pnl_list,
            placeholder="ej: C20, G20, K20  o  C20:C22, G20:G22")
        self._uni_cell_list.grid(row=0, column=1, sticky="ew", pady=(0, 4))
        _muted(self._uni_pnl_list,
               "Separa con comas. Acepta celdas simples y rangos mezclados.").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 4))

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
            c, height=155,
            font=ctk.CTkFont(family="Cascadia Code", size=11),
            fg_color="#0D1117", text_color="#C4B5FD",
            corner_radius=8, wrap="word", state="disabled")
        self._uni_log.grid(row=4, column=0, sticky="ew",
                           padx=CARD_PAD, pady=(0, 6))
        _btn(c, "🗑  Limpiar log", self._uni_clear_log,
             color="transparent", hover="#2D3A55",
             width=120, height=28).grid(
             row=5, column=0, sticky="e", padx=CARD_PAD, pady=(0, CARD_PAD))

    # ── Logica: modo fuente ───────────────────────────────────────────────────
    def _uni_src_mode_change(self, val):
        if val == "Archivo Unico":
            self._uni_multi_pnl.grid_forget()
            self._uni_single_pnl.grid(
                row=2, column=0, columnspan=3, sticky="ew",
                padx=CARD_PAD, pady=(0, 6))
        else:
            self._uni_single_pnl.grid_forget()
            self._uni_multi_pnl.grid(
                row=2, column=0, columnspan=3, sticky="ew",
                padx=CARD_PAD, pady=(0, 6))

    # ── Logica: selector archivos ─────────────────────────────────────────────
    def _uni_select_single(self):
        p = filedialog.askopenfilename(
            title="Selecciona el archivo origen",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if p:
            self._uni_src_files = [Path(p)]
            self._uni_src_entry.delete(0, "end")
            self._uni_src_entry.insert(0, str(self._uni_src_files[0]))
            self._uni_load_src_sheets()

    def _uni_select_multi(self):
        paths = filedialog.askopenfilenames(
            title="Selecciona los archivos Excel origen",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if paths:
            self._uni_src_files = [Path(p) for p in paths]
            names = [p.name for p in self._uni_src_files]
            self._uni_multi_list.configure(state="normal")
            self._uni_multi_list.delete("1.0", "end")
            self._uni_multi_list.insert("end", "\n".join(names))
            self._uni_multi_list.configure(state="disabled")
            self._uni_multi_count.configure(
                text=f"{len(self._uni_src_files)} archivo(s) seleccionado(s)",
                text_color=SUCCESS)
            self._uni_log_msg(f"Multi-archivo: {len(self._uni_src_files)} archivos.")

    # ── Logica: hojas con checkboxes ──────────────────────────────────────────
    def _uni_load_src_sheets(self):
        v = self._uni_src_entry.get().strip()
        if v:
            f = Path(v)
            if f.exists():
                self._uni_src_files = [f]
        if not self._uni_src_files or not self._uni_src_files[0].exists():
            messagebox.showwarning("Archivo no encontrado",
                "Selecciona primero un archivo origen valido.")
            return
        try:
            names = get_sheet_names(self._uni_src_files[0])
            self._uni_build_sheet_checkboxes(names)
            self._uni_sheets_hint.configure(
                text=f"  {len(names)} hoja(s) encontrada(s)",
                text_color=SUCCESS)
            self._uni_log_msg(f"Hojas cargadas: {names}")
        except Exception as exc:
            messagebox.showerror("Error al leer hojas", str(exc))

    def _uni_build_sheet_checkboxes(self, sheet_names: list[str]):
        # Limpiar checkboxes anteriores
        for w in self._uni_sheets_scroll.winfo_children():
            w.destroy()
        self._uni_sheet_vars = {}
        for name in sheet_names:
            var = ctk.BooleanVar(value=True)
            self._uni_sheet_vars[name] = var
            ctk.CTkCheckBox(
                self._uni_sheets_scroll,
                text=name,
                variable=var,
                font=ctk.CTkFont(size=11),
                fg_color=PURPLE, hover_color=PURPLE_DIM,
                checkmark_color="white",
                width=15, height=15,
            ).pack(side="left", padx=(0, 10), pady=4)

    def _uni_check_all(self, state: bool):
        for var in self._uni_sheet_vars.values():
            var.set(state)

    def _uni_get_selected_sheets(self) -> list[str]:
        return [name for name, var in self._uni_sheet_vars.items() if var.get()]

    # ── Logica: destino ───────────────────────────────────────────────────────
    def _uni_dest_change(self, val):
        if val == "Excel Local":
            self._uni_sheets_pnl.grid_forget()
            self._uni_excel_pnl.grid(row=2, column=0, columnspan=3,
                                      sticky="ew")
        else:
            self._uni_excel_pnl.grid_forget()
            self._uni_sheets_pnl.grid(row=2, column=0, columnspan=3,
                                       sticky="ew")

    def _uni_map_mode_change(self):
        if self._uni_dest_map_mode.get() == "Consolidar":
            self._uni_dest_sheet_menu.configure(state="normal")
        else:
            self._uni_dest_sheet_menu.configure(state="disabled")

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
            self._uni_log_msg(f"Hojas destino: {names}")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    # ── Google Sheets: cargar hojas y modo de mapeo ──────────────────────────
    def _uni_load_sheets_tabs(self):
        url    = self._uni_sheets_url.get().strip()
        creds  = self._uni_creds_entry.get().strip() or "credentials.json"
        if not url:
            messagebox.showwarning("URL requerida",
                "Ingresa la URL o ID del Spreadsheet.")
            return
        if not GSPREAD_OK:
            messagebox.showerror("gspread no instalado",
                "Instala gspread con: pip install gspread google-auth")
            return
        try:
            tabs = get_sheet_tabs(url, creds)
            self._uni_sheets_tab_menu.configure(values=tabs, state="normal")
            self._uni_sheets_tab_var.set(tabs[0])
            self._uni_sheets_tabs_hint.configure(
                text=f"  {len(tabs)} hoja(s) encontrada(s)",
                text_color=SUCCESS)
            # Guardar para el MappingDialog
            self._uni_gs_tabs_cache = tabs
            self._uni_log_msg(f"Hojas Sheets cargadas: {tabs}")
        except Exception as exc:
            messagebox.showerror("Error al cargar hojas", str(exc))

    def _uni_gs_map_mode_change(self):
        if self._uni_gs_map_mode.get() == "Consolidar":
            self._uni_sheets_tab_menu.configure(state="normal")
        else:
            self._uni_sheets_tab_menu.configure(state="disabled")

    def _uni_select_creds(self):
        p = filedialog.askopenfilename(
            title="credentials.json",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")])
        if p:
            self._uni_creds_entry.delete(0, "end")
            self._uni_creds_entry.insert(0, p)

    # ── Logica: modalidad ─────────────────────────────────────────────────────
    def _uni_mode_change(self, val):
        for pnl in (self._uni_pnl_block,
                    self._uni_pnl_stride,
                    self._uni_pnl_list):
            pnl.grid_forget()
        if val == MODES[0]:
            self._uni_pnl_block.grid(
                row=2, column=0, columnspan=4, sticky="ew",
                padx=CARD_PAD, pady=(0, CARD_PAD))
        elif val == MODES[1]:
            self._uni_pnl_stride.grid(
                row=2, column=0, columnspan=4, sticky="ew",
                padx=CARD_PAD, pady=(0, CARD_PAD))
        else:
            self._uni_pnl_list.grid(
                row=2, column=0, columnspan=4, sticky="ew",
                padx=CARD_PAD, pady=(0, CARD_PAD))

    # ── Vista previa del rango ────────────────────────────────────────────────
    def _uni_preview_range(self):
        src_files = self._get_uni_src_files()
        v_rng = self._uni_range_entry.get().strip()
        if not src_files:
            messagebox.showwarning("Sin archivo origen",
                "Selecciona al menos un archivo origen.")
            return
        if not v_rng:
            messagebox.showwarning("Sin rango",
                "Escribe el rango de origen.")
            return
        try:
            sheet = self._get_uni_src_sheet_for_preview(src_files[0])
            values = extract_multi_range(src_files[0], v_rng, sheet)
            n = len(values)
            preview = str(values[:8])
            if n > 8:
                preview = preview[:-1] + ", ...]"
            self._uni_range_info.configure(
                text=f"{n} valores",
                text_color="#C4B5FD")
            self._uni_log_msg(
                f"Vista previa '{v_rng}': {n} valores  ->  {preview}")
        except Exception as exc:
            messagebox.showerror("Error al leer rango", str(exc))

    def _get_uni_src_files(self) -> list[Path]:
        if self._uni_src_mode.get() == "Archivo Unico":
            v = self._uni_src_entry.get().strip()
            if v:
                f = Path(v)
                if f.exists():
                    self._uni_src_files = [f]
            return self._uni_src_files
        else:
            return self._uni_src_files

    def _get_uni_src_sheet_for_preview(self, fp: Path):
        """Retorna la primera hoja seleccionada (o None)."""
        if self._uni_sheet_vars:
            sel = self._uni_get_selected_sheets()
            return sel[0] if sel else None
        return None

    # ── Validacion de campos obligatorios ─────────────────────────────────────
    def _uni_validate(self) -> tuple[bool, str]:
        """Retorna (valido, mensaje_error)."""
        mode = self._uni_mode_var.get()

        if mode == MODES[0]:   # Bloque Continuo
            v = self._uni_start_cell.get().strip()
            if not v:
                return False, (
                    "Modalidad Bloque Continuo requiere una celda de inicio. "
                    "Por favor ingresa la celda de inicio (ej: B5).")
        elif mode == MODES[1]:  # Salto
            v1 = self._uni_stride_start.get().strip()
            v2 = self._uni_stride_n.get().strip()
            if not v1:
                return False, (
                    "Modalidad Salto requiere la celda de inicio. "
                    "Por favor ingresa la celda de inicio (ej: C20).")
            if not v2:
                return False, (
                    "Modalidad Salto requiere el numero de salto (N). "
                    "Por favor ingresa el salto (ej: 7).")
            try:
                int(v2)
            except ValueError:
                return False, f"El salto debe ser un numero entero, no '{v2}'."
        else:  # Lista de Celdas
            v = self._uni_cell_list.get().strip()
            if not v:
                return False, (
                    "Modalidad Lista de Celdas requiere al menos una celda. "
                    "Por favor ingresa las celdas destino (ej: C20, G20, K20).")
            try:
                _expand_cell_tokens(v)
            except ValueError as e:
                return False, f"Lista de celdas invalida: {e}"

        # Validar rango origen
        v_rng = self._uni_range_entry.get().strip()
        if not v_rng:
            return False, "Por favor ingresa el rango de origen (ej: D6:D15)."

        return True, ""

    # ── Transferencia ─────────────────────────────────────────────────────────
    def _uni_on_transfer(self):
        if self._uni_running:
            return

        # ── Validaciones ──────────────────────────────────────────────────────
        src_files = self._get_uni_src_files()
        if not src_files:
            messagebox.showwarning("Origen requerido",
                "Selecciona al menos un archivo Excel origen.")
            return

        v_rng = self._uni_range_entry.get().strip()
        if not v_rng:
            messagebox.showwarning("Rango requerido",
                "Escribe el rango de origen (ej: D6:D15).")
            return

        ok, msg = self._uni_validate()
        if not ok:
            messagebox.showwarning("Configuracion incompleta", msg)
            return

        dest_mode = self._uni_dest_mode.get()
        if dest_mode == "Excel Local":
            v_dest = self._uni_dest_entry.get().strip()
            if v_dest:
                self._uni_dest_file = Path(v_dest)
            if not self._uni_dest_file or not self._uni_dest_file.exists():
                messagebox.showwarning("Destino requerido",
                    "Selecciona un archivo Excel destino valido.")
                return

        # ── Parametros de escritura ───────────────────────────────────────────
        mode_paste = self._uni_mode_var.get()
        start_cell = (self._uni_stride_start.get().strip()
                      if mode_paste == MODES[1]
                      else self._uni_start_cell.get().strip())
        try:
            stride = max(1, int(self._uni_stride_n.get().strip() or "1"))
        except ValueError:
            stride = 1

        if self._uni_src_mode.get() == "Archivo Unico":
            selected_sheets = self._uni_get_selected_sheets() or [None]
        else:
            selected_sheets = [None]

        map_mode = self._uni_dest_map_mode.get()

        # ── Calcular jobs ─────────────────────────────────────────────────────
        jobs_raw: list[tuple[Path, str | None]] = []
        if len(src_files) == 1:
            for sh in selected_sheets:
                jobs_raw.append((src_files[0], sh if sh else None))
        else:
            for fp in src_files:
                jobs_raw.append((fp, None))

        # ── Si modo "Mismo nombre" en Excel: mostrar dialogo de confirmacion ──
        confirmed_jobs: list[dict] | None = None

        # Determinar hojas destino disponibles para el MappingDialog
        dest_sheets_avail: list[str] = []
        need_dialog = False

        if dest_mode == "Excel Local" and map_mode == "Mismo nombre":
            if self._uni_dest_file and self._uni_dest_file.exists():
                try:
                    dest_sheets_avail = get_sheet_names(self._uni_dest_file)
                except Exception:
                    pass
            need_dialog = True

        elif dest_mode == "Google Sheets":
            gs_map = self._uni_gs_map_mode.get()
            if gs_map == "Mismo nombre":
                tabs_cache = getattr(self, "_uni_gs_tabs_cache", [])
                if not tabs_cache:
                    # Intentar cargar en el momento
                    url   = self._uni_sheets_url.get().strip()
                    creds = self._uni_creds_entry.get().strip() or "credentials.json"
                    if url and GSPREAD_OK:
                        try:
                            tabs_cache = get_sheet_tabs(url, creds)
                            self._uni_gs_tabs_cache = tabs_cache
                        except Exception as exc:
                            messagebox.showerror("Error Google Sheets", str(exc))
                            return
                    else:
                        messagebox.showwarning("Hojas no cargadas",
                            "Presiona 'Cargar hojas' en la seccion de destino.")
                        return
                dest_sheets_avail = tabs_cache
                need_dialog = True

        if need_dialog:
            # Calcular matches previos (sin escribir)
            dialog_rows: list[dict] = []
            for fp, src_sheet in jobs_raw:
                origin_label = src_sheet if src_sheet else fp.name
                lbl          = fp.name + (f"/{src_sheet}" if src_sheet else "")
                if dest_sheets_avail:
                    dest_sh, score = match_dest_sheet(origin_label, dest_sheets_avail)
                else:
                    dest_sh, score = None, 0.0
                dialog_rows.append({
                    "label":      lbl,
                    "fp":         fp,
                    "src_sheet":  src_sheet,
                    "dest_sh":    dest_sh,
                    "score":      score,
                })

            # Abrir MappingDialog modal (igual para Excel y Sheets)
            dlg = MappingDialog(self, dialog_rows, dest_sheets_avail)
            self.wait_window(dlg)

            if dlg.result is None:
                return
            if not dlg.result:
                messagebox.showinfo("Sin filas activas",
                    "No hay filas activadas para transferir.")
                return
            confirmed_jobs = dlg.result

        # ── Construir params y lanzar worker ──────────────────────────────────
        # Calcular hoja Sheets fija (modo Consolidar)
        gs_tab_fixed = ""
        if dest_mode == "Google Sheets":
            gs_mode = self._uni_gs_map_mode.get()
            if gs_mode == "Consolidar":
                gs_tab_fixed = self._uni_sheets_tab_var.get()
            else:
                gs_tab_fixed = ""  # el dialogo asigna por job

        params = {
            "src_files":       src_files,
            "selected_sheets": selected_sheets,
            "src_range":       v_rng,
            "dest_mode":       dest_mode,
            "paste_mode":      mode_paste,
            "start_cell":      start_cell or "A1",
            "direction":       self._uni_dir_var.get(),
            "stride":          stride,
            "cell_list":       self._uni_cell_list.get().strip(),
            "dest_file":       self._uni_dest_file,
            "dest_sheet":      self._uni_dest_sheet_var.get(),
            "map_mode":        map_mode,
            "dest_file_name":  (self._uni_dest_file.name
                                if self._uni_dest_file else ""),
            "sheets_url":      self._uni_sheets_url.get().strip(),
            "sheets_creds":    (self._uni_creds_entry.get().strip()
                                or "credentials.json"),
            "sheets_tab_fixed": gs_tab_fixed,
            "gs_map_mode":     getattr(self, "_uni_gs_map_mode",
                               ctk.StringVar(value="Mismo nombre")).get(),
            # Jobs pre-confirmados (None = el worker calcula su propio mapeo)
            "confirmed_jobs":  confirmed_jobs,
        }

        self._uni_running = True
        self._uni_run_btn.configure(state="disabled", text="⏳  Procesando...")
        self._uni_bar.set(0)
        self._uni_status.configure(text="Migrando rango...", text_color=PURPLE)
        self._uni_log_msg("─" * 50)

        threading.Thread(target=self._uni_worker,
                         args=(params,), daemon=True).start()

    def _uni_worker(self, p: dict):
        try:
            src_files:      list[Path]       = p["src_files"]
            sel_sheets:     list             = p["selected_sheets"]
            v_rng:          str              = p["src_range"]
            paste_mode:     str              = p["paste_mode"]
            start_cell:     str              = p["start_cell"]
            direction:      str              = p["direction"]
            stride:         int              = p["stride"]
            cell_list:      str              = p["cell_list"]
            dest_mode:      str              = p["dest_mode"]
            map_mode:       str              = p["map_mode"]
            dest_file:      Path             = p["dest_file"]
            dest_sheet_fixed: str            = p["dest_sheet"]
            dest_name:      str              = p["dest_file_name"]
            confirmed_jobs: list[dict] | None = p.get("confirmed_jobs")

            # ── Construir lista de trabajos ───────────────────────────────────
            # Si hay jobs pre-confirmados (via MappingDialog), usarlos directo.
            # De lo contrario, calcular desde src_files + sel_sheets.
            cells_total = 0
            failed:    list[str] = []
            unmatched: list[str] = []
            backup_done = False

            if confirmed_jobs is not None:
                # Modo con dialogo: jobs ya tienen (fp, src_sheet, dest_sh, label)
                total = len(confirmed_jobs)
                self._qlog("uni", f"Iniciando transferencia de {total} origen(s) confirmado(s).")

                for idx, job in enumerate(confirmed_jobs):
                    fp:       Path      = job["fp"]
                    src_sheet: str|None = job["src_sheet"]
                    dest_sh:  str       = job["dest_sh"]
                    label:    str       = job["label"]

                    self._qlog("uni", f"  [{idx+1}/{total}]  {label}  | rango: {v_rng}")
                    self._qlog("uni", f"    Destino: {dest_name} -> '{dest_sh}'")
                    self._q.put(("uni_progress", idx / total))

                    try:
                        values = extract_multi_range(fp, v_rng, src_sheet)
                        self._qlog("uni", f"    Extraidos: {len(values)} valores")
                    except Exception as exc:
                        self._qlog("uni", f"    Extraccion fallida: {exc}", "error")
                        failed.append(label); continue

                    self._q.put(("uni_progress", (idx + 0.5) / total))

                    if dest_mode == "Excel Local":
                        try:
                            if paste_mode == MODES[0]:
                                result = write_block(
                                    dest_file, dest_sh, [[v] for v in values],
                                    start_cell=start_cell, create_backup=not backup_done)
                            elif paste_mode == MODES[1]:
                                result = write_stride(
                                    dest_file, dest_sh, values,
                                    start_cell=start_cell, direction=direction,
                                    stride=stride, create_backup=not backup_done)
                            else:
                                result = write_cell_list(
                                    dest_file, dest_sh, values,
                                    cell_list_str=cell_list,
                                    create_backup=not backup_done)
                            backup_done = True
                            cells_total += result["written"]
                            det = result.get("detail", [])
                            prev = ", ".join(det[:6])
                            if len(det) > 6:
                                prev += f" ... (+{len(det)-6})"
                            self._qlog("uni", f"    Guardado: {prev}")
                        except PermissionError:
                            self._qlog("uni",
                                "    ARCHIVO ABIERTO EN EXCEL. Cierralo e intenta de nuevo.",
                                "error")
                            failed.append(label); break
                        except Exception as exc:
                            self._qlog("uni", f"    Error escritura: {exc}", "error")
                            failed.append(label)
                    else:  # Google Sheets
                        sheets_url    = p["sheets_url"]
                        sheets_creds  = p.get("sheets_creds", "credentials.json")
                        try:
                            result = write_range_to_sheets(
                                spreadsheet_id_or_url = sheets_url,
                                sheet_name  = dest_sh,
                                values      = values,
                                mode        = paste_mode,
                                credentials_path = sheets_creds,
                                start_cell  = start_cell,
                                direction   = direction,
                                stride      = stride,
                                cell_list_str = cell_list,
                            )
                            cells_total += result["written"]
                            det = result.get("detail", [])
                            prev = ", ".join(det[:6])
                            if len(det) > 6:
                                prev += f" ... (+{len(det)-6})"
                            self._qlog("uni", f"    Sheets guardado: {prev}")
                        except RuntimeError as exc:
                            self._qlog("uni", f"    {exc}", "warn")
                            failed.append(label)
                        except Exception as exc:
                            self._qlog("uni", f"    Error Sheets: {exc}", "error")
                            failed.append(label)

                    self._q.put(("uni_progress", (idx + 1) / total))

                self._q.put(("uni_done", {
                    "total": total, "failed": failed,
                    "unmatched": unmatched, "cells": cells_total,
                }))
                return

            # ── Modo sin dialogo (Consolidar / Google Sheets) ─────────────────
            jobs: list[tuple[Path, str | None]] = []
            if len(src_files) == 1:
                for sh in sel_sheets:
                    jobs.append((src_files[0], sh if sh else None))
            else:
                for fp in src_files:
                    jobs.append((fp, None))

            total = len(jobs)
            dest_sheets_available: list[str] = []
            if dest_mode == "Excel Local" and dest_file and dest_file.exists():
                try:
                    dest_sheets_available = get_sheet_names(dest_file)
                except Exception:
                    dest_sheets_available = []

            for idx, (fp, src_sheet) in enumerate(jobs):
                label = f"{fp.name}" + (f"/{src_sheet}" if src_sheet else "")
                self._qlog("uni", f"  [{idx+1}/{total}]  {label}  | rango: {v_rng}")
                self._q.put(("uni_progress", (idx) / total))

                try:
                    values = extract_multi_range(fp, v_rng, src_sheet)
                    self._qlog("uni", f"    Extraidos: {len(values)} valores")
                except Exception as exc:
                    self._qlog("uni", f"    Extraccion fallida: {exc}", "error")
                    failed.append(label); continue

                if dest_mode == "Excel Local":
                    if map_mode == "Consolidar":
                        dest_sh = dest_sheet_fixed
                    else:
                        origin_label = src_sheet if src_sheet else fp.name
                        if dest_sheets_available:
                            dest_sh, score = match_dest_sheet(
                                origin_label, dest_sheets_available)
                            if dest_sh is None:
                                self._qlog(
                                    "uni",
                                    f"    Omitido: no se encontro pestana destino "
                                    f"para '{origin_label}' "
                                    f"(Disponibles: {dest_sheets_available})",
                                    "warn",
                                )
                                unmatched.append(origin_label)
                                self._q.put(("uni_progress", (idx + 1) / total))
                                continue
                            self._qlog(
                                "uni",
                                f"    Match: '{origin_label}' -> '{dest_sh}' "
                                f"(score={score:.0f})",
                            )
                        else:
                            dest_sh = src_sheet if src_sheet else fp.stem

                    # Pre-visualizacion
                    if paste_mode == MODES[2]:
                        try:
                            cp = _expand_cell_tokens(cell_list)
                            ps = ", ".join(cp[:10])
                            if len(cp) > 10:
                                ps += f" ... (+{len(cp)-10})"
                            self._qlog("uni", f"    Pegando en: {ps}")
                        except ValueError as e:
                            self._qlog("uni", f"    Lista invalida: {e}", "error")
                            failed.append(label); continue
                    elif paste_mode == MODES[1]:
                        self._qlog("uni",
                            f"    Inicio={start_cell} dir={direction} salto={stride}")
                    else:
                        self._qlog("uni", f"    Bloque desde: {start_cell}")

                    self._qlog("uni",
                        f"    Destino: {dest_name} -> '{dest_sh}'")
                    self._q.put(("uni_progress", (idx + 0.5) / total))

                    # ── Escritura directa con valores ya extraidos ──────────
                    # No volver a parsear v_rng; usar 'values' directamente.
                    try:
                        if paste_mode == MODES[0]:        # Bloque Continuo
                            # write_block espera list[list]; envolvemos cada
                            # valor en su propia fila [[v1],[v2],...]
                            matrix_2d = [[v] for v in values]
                            result = write_block(
                                dest_file, dest_sh, matrix_2d,
                                start_cell    = start_cell,
                                create_backup = not backup_done,
                            )
                        elif paste_mode == MODES[1]:       # Salto
                            result = write_stride(
                                dest_file, dest_sh, values,
                                start_cell    = start_cell,
                                direction     = direction,
                                stride        = stride,
                                create_backup = not backup_done,
                            )
                        else:                              # Lista de Celdas
                            result = write_cell_list(
                                dest_file, dest_sh, values,
                                cell_list_str = cell_list,
                                create_backup = not backup_done,
                            )
                        backup_done = True
                        cells_total += result["written"]
                        det = result.get("detail", [])
                        preview = ", ".join(det[:6])
                        if len(det) > 6:
                            preview += f" ... (+{len(det)-6})"
                        self._qlog("uni", f"    Guardado: {preview}")
                    except PermissionError:
                        self._qlog("uni",
                            "    ARCHIVO ABIERTO EN EXCEL. Cierralo e intenta de nuevo.",
                            "error")
                        failed.append(label); break
                    except Exception as exc:
                        self._qlog("uni", f"    Error escritura: {exc}", "error")
                        failed.append(label)

                else:  # Google Sheets
                    sheets_url    = p["sheets_url"]
                    sheets_creds  = p.get("sheets_creds", "credentials.json")
                    # Si confirmed_jobs ya asigno dest_sh, usarlo;
                    # de lo contrario usar gs_tab_fixed o label como fallback
                    if "dest_sh" not in locals():
                        gs_mode = p.get("gs_map_mode", "Consolidar")
                        if gs_mode == "Consolidar":
                            dest_sh = p.get("sheets_tab_fixed", "")
                        else:
                            dest_sh = src_sheet if src_sheet else fp.stem
                    if not sheets_url:
                        self._qlog("uni",
                            "    Google Sheets: URL no configurada.", "error")
                        failed.append(label)
                    elif not dest_sh:
                        self._qlog("uni",
                            "    Google Sheets: hoja destino no especificada.", "error")
                        failed.append(label)
                    else:
                        self._qlog("uni",
                            f"    Google Sheets -> hoja '{dest_sh}'")
                        try:
                            result = write_range_to_sheets(
                                spreadsheet_id_or_url = sheets_url,
                                sheet_name  = dest_sh,
                                values      = values,
                                mode        = paste_mode,
                                credentials_path = sheets_creds,
                                start_cell  = start_cell,
                                direction   = direction,
                                stride      = stride,
                                cell_list_str = cell_list,
                            )
                            cells_total += result["written"]
                            det = result.get("detail", [])
                            prev = ", ".join(det[:6])
                            if len(det) > 6:
                                prev += f" ... (+{len(det)-6})"
                            self._qlog("uni", f"    Sheets guardado: {prev}")
                        except RuntimeError as exc:
                            # Error 429 / cuota
                            self._qlog("uni", f"    {exc}", "warn")
                            failed.append(label)
                        except Exception as exc:
                            self._qlog("uni",
                                f"    Error Sheets: {exc}", "error")
                            failed.append(label)

                self._q.put(("uni_progress", (idx + 1) / total))

            self._q.put(("uni_done", {
                "total":     total,
                "failed":    failed,
                "unmatched": unmatched,
                "cells":     cells_total,
            }))

        except Exception as exc:
            self._qlog("uni", f"Error critico: {exc}", "error")
            self._q.put(("uni_done", None))
        finally:
            self._uni_running = False

    def _uni_on_done(self, result):
        self._uni_run_btn.configure(state="normal", text="🎯  Transferir Rango")
        if result is None:
            self._uni_status.configure(
                text="Error en la transferencia.", text_color=DANGER)
            return
        n         = result.get("cells", 0)
        total     = result.get("total", 0)
        failed    = result.get("failed", [])
        unmatched = result.get("unmatched", [])
        any_issue = bool(failed or unmatched)
        self._uni_bar.set(1.0)

        ok_count = total - len(failed) - len(unmatched)
        self._uni_status.configure(
            text=f"{'OK' if not any_issue else 'Parcial'}: "
                 f"{ok_count}/{total} origenes  |  {n} celdas"
                 + (f"  |  {len(unmatched)} sin match" if unmatched else ""),
            text_color=SUCCESS if not any_issue else WARNING)

        self._uni_log_msg("─" * 50)

        # Log de no-matches en el panel
        for u in unmatched:
            self._uni_log_msg(
                f"  Omitido (sin match): '{u}'", "warn")

        summary = (
            f"Origenes procesados : {ok_count} / {total}\n"
            f"Celdas escritas     : {n}\n"
            f"Modalidad           : {self._uni_mode_var.get()}"
        )
        if failed:
            summary += f"\nErrores de escritura: {', '.join(failed)}"
        if unmatched:
            names = "\n  ".join(unmatched)
            summary += (
                f"\n\nHojas sin coincidencia omitidas ({len(unmatched)}):\n"
                f"  {names}"
            )

        if not any_issue:
            messagebox.showinfo("Transferencia completada",
                                "Completado!\n\n" + summary)
        else:
            messagebox.showwarning(
                "Transferencia " + ("con omisiones" if unmatched and not failed
                                    else "parcial"),
                ("Finalizado con advertencias.\n\n"
                 if not failed else "Finalizado con errores.\n\n") + summary)

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
