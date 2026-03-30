"""
Unified Measurement Tool - Main Application
Combines vertical and horizontal measurement capabilities with customizable color filtering
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser, ttk
from typing import Dict, List, Optional, Tuple
import cv2
import numpy as np
from PIL import Image, ImageTk

from utils import (
    select_folder,
    select_file,
    find_images_recursively,
    process_images,
    rgb_to_hex,
    categorize_data_files,
    categorize_mask_files,
    build_pixel_size_map,
)
from image_processor import (
    draw_segments_on_image,
    analyze_all_vertical_segments,
    analyze_all_horizontal_segments,
    measure_mask_image,
)
from measurement_engine import (
    save_vertical_segments_to_csv,
    save_vertical_segments_to_excel,
    save_horizontal_segments_to_csv,
    save_horizontal_segments_to_excel,
    calculate_segment_statistics,
    write_measurements_csv,
    write_measurements_excel,
    write_interface_distances_csv,
    write_interface_distances_excel,
)
from update_manager import UpdateManager
from version import __version__

# ── Palette ───────────────────────────────────────────────────────────────────
ACCENT      = "#CC0000"   # NCSU red
SUCCESS     = "#111827"   # black accent for success states
DANGER      = "#7A0019"   # deep red for failures/errors
BG          = "#F7F3F3"   # warm off-white background
SURFACE     = "#FFF9F9"   # light surface color
PANEL_BG    = "#F6E3E3"   # light red-tinted panel background
HEADER_BG   = "#990000"   # dark NCSU red header
BORDER      = "#D9B5B5"   # soft border tone
SUBTAB_BG   = "#EFD9D9"   # muted sub-tab surface
SUBTAB_ON   = "#FFF4F4"   # active sub-tab surface
THUMB_COLS  = 4           # columns in the preview/segments thumbnail grids
BTN_BG      = "#F0DEDE"   # soft button background
BTN_HOVER   = "#FFDADA"   # light red hover/active background
BTN_TEXT    = "#111827"   # near-black button text


class UnifiedMeasurementApp:
    """Main application class for the Unified Measurement Tool."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Unified Measurement Tool")
        self.root.geometry("1400x900")
        self.root.minsize(900, 600)
        self.root.configure(bg=BG)

        # ── State ──────────────────────────────────────────────────────────
        self.folder_path: str = ""
        self.images_list: list = []
        self.processed_images: list = []
        self.thumbnail_refs: list = []
        self.selected_color: tuple = (128, 0, 128)      # default purple
        self.measurement_mode = tk.StringVar(value="vertical")

        # Ellipsoidal workflow state
        self.ellipsoidal_data_folder_path: str = ""
        self.ellipsoidal_mask_folder_path: str = ""
        self.ellipsoidal_txt_rel_paths: List[str] = []
        self.ellipsoidal_pixel_size_map: Dict[str, Tuple[Optional[float], Optional[float]]] = {}
        self.ellipsoidal_mask_files: List[str] = []
        self.ellipsoidal_processed_rows: List[Dict[str, object]] = []
        self.ellipsoidal_interface_rows: List[Dict[str, object]] = []
        self.ellipsoidal_last_preview_refs: List[ImageTk.PhotoImage] = []
        self.update_manager = UpdateManager()
        self.startup_update_check_started = False

        self._configure_styles()
        self._build_header()
        self._build_notebook()

        # Global mouse-wheel scrolling for the active tab's canvas
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        self.root.bind_all("<Button-4>",   self._on_mousewheel)  # Linux scroll up
        self.root.bind_all("<Button-5>",   self._on_mousewheel)  # Linux scroll down
        self.root.after(1200, self._auto_check_for_updates)

    def _bind_button_hover(self, button: tk.Button) -> None:
        """Give tk.Button a consistent light-blue hover state with stable text color."""
        button.bind("<Enter>", lambda _e: button.configure(bg=BTN_HOVER, fg=BTN_TEXT))
        button.bind("<Leave>", lambda _e: button.configure(bg=BTN_BG, fg=BTN_TEXT))

    def _add_section_divider(
        self,
        parent,
        row: int,
        column: int = 0,
        columnspan: int = 1,
        padx: int | Tuple[int, int] = 0,
        pady: Tuple[int, int] = (0, 8),
    ) -> None:
        """Add a subtle divider to reinforce section hierarchy."""
        tk.Frame(parent, bg=BORDER, height=1).grid(
            row=row,
            column=column,
            columnspan=columnspan,
            sticky="ew",
            padx=padx,
            pady=pady,
        )
    
    # ── Styles ────────────────────────────────────────────────────────────────
    def _configure_styles(self) -> None:
        s = ttk.Style(self.root)
        s.theme_use("clam")
        s.configure("TNotebook", background=BG, borderwidth=0)
        s.configure("TNotebook.Tab", padding=[14, 6], font=("Arial", 10, "bold"))
        s.map("TNotebook.Tab",
              background=[("selected", ACCENT),  ("!selected", "#EBCACA")],
              foreground=[("selected", "white"),  ("!selected", BTN_TEXT)])
        s.configure("Layer.TNotebook", background=BG, borderwidth=0, tabmargins=[12, 10, 12, 0])
        s.configure(
            "Layer.TNotebook.Tab",
            padding=[16, 7],
            font=("Arial", 10, "bold"),
            background=SUBTAB_BG,
            foreground=HEADER_BG,
            borderwidth=0,
        )
        s.map(
            "Layer.TNotebook.Tab",
            background=[("selected", SUBTAB_ON), ("active", BTN_HOVER), ("!selected", SUBTAB_BG)],
            foreground=[("selected", HEADER_BG), ("active", HEADER_BG), ("!selected", "#6B1B1B")],
        )
        s.configure("TFrame", background=BG)
        s.configure("TLabel", background=BG, font=("Arial", 10))
        s.configure("TLabelframe", background=PANEL_BG, bordercolor=BORDER, relief="solid")
        s.configure(
            "TLabelframe.Label",
            background=PANEL_BG,
            font=("Arial", 10, "bold"),
            foreground=HEADER_BG,
        )
        s.configure("TRadiobutton", background=BG, font=("Arial", 10))
        s.configure("TCheckbutton", background=BG, font=("Arial", 10))

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg=HEADER_BG, pady=8)
        header.pack(side=tk.TOP, fill=tk.X)

        tk.Label(header, text="Unified Measurement Tool",
                 bg=HEADER_BG, fg="white",
                 font=("Arial", 15, "bold")).pack(side=tk.LEFT, padx=16)

        right = tk.Frame(header, bg=HEADER_BG)
        right.pack(side=tk.RIGHT, padx=16)

        choose_folder_btn = tk.Button(right, text="Choose Image Folder",
                          command=self.choose_folder,
                          bg=BTN_BG, fg=BTN_TEXT,
                          activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
                          font=("Arial", 10, "bold"),
                          relief=tk.FLAT, padx=12, pady=5,
                          cursor="hand2")
        choose_folder_btn.pack(side=tk.LEFT, padx=(0, 14))
        self._bind_button_hover(choose_folder_btn)

        update_btn = tk.Button(right, text=f"Check for Updates ({__version__})",
                  command=self.check_for_updates,
                  bg=BTN_BG, fg=BTN_TEXT,
                  activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
                  font=("Arial", 10, "bold"),
                  relief=tk.FLAT, padx=12, pady=5,
                  cursor="hand2")
        update_btn.pack(side=tk.LEFT, padx=(0, 14))
        self._bind_button_hover(update_btn)

        info = tk.Frame(right, bg=HEADER_BG)
        info.pack(side=tk.LEFT)

        self.folder_var = tk.StringVar(value="No folder selected")
        self.images_var = tk.StringVar(value="")

        tk.Label(info, textvariable=self.folder_var,
                 bg=HEADER_BG, fg="#CBD5E1",
                 font=("Arial", 9), wraplength=480, anchor="w").pack(anchor="w")
        tk.Label(info, textvariable=self.images_var,
                 bg=HEADER_BG, fg="#86EFAC",
                 font=("Arial", 9, "bold"), anchor="w").pack(anchor="w")

    # ── Notebook ──────────────────────────────────────────────────────────────
    def _build_notebook(self) -> None:
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_layer_measurements = ttk.Frame(self.notebook)
        self.tab_ellipsoidal  = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_layer_measurements, text="  Layer Measurements  ")
        self.notebook.add(self.tab_ellipsoidal,  text="  Ellipsoidal Features  ")

        self._build_tab_layer_measurements(self.tab_layer_measurements)
        self._build_tab_ellipsoidal(self.tab_ellipsoidal)

    def _build_tab_layer_measurements(self, parent: ttk.Frame) -> None:
        """Create nested sub-tabs for layer-measurement workflows."""
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        self.layer_notebook = ttk.Notebook(parent, style="Layer.TNotebook")
        self.layer_notebook.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        self.tab_layer_setup_preview = ttk.Frame(self.layer_notebook)
        self.tab_layer_segments_export = ttk.Frame(self.layer_notebook)

        self.layer_notebook.add(self.tab_layer_setup_preview, text="  Setup + Preview  ")
        self.layer_notebook.add(self.tab_layer_segments_export, text="  Segments + Export  ")

        # Setup + Preview split
        self.tab_layer_setup_preview.columnconfigure(0, weight=3)
        self.tab_layer_setup_preview.columnconfigure(1, weight=2)
        self.tab_layer_setup_preview.rowconfigure(0, weight=1)

        layer_setup_frame = ttk.Frame(self.tab_layer_setup_preview)
        layer_setup_frame.grid(row=0, column=0, sticky="nsew")

        layer_preview_frame = ttk.Frame(self.tab_layer_setup_preview)
        layer_preview_frame.grid(row=0, column=1, sticky="nsew")

        self._build_tab_setup(layer_setup_frame)
        self._build_tab_preview(layer_preview_frame)

        # Segments + Export split
        self.tab_layer_segments_export.columnconfigure(0, weight=3)
        self.tab_layer_segments_export.columnconfigure(1, weight=2)
        self.tab_layer_segments_export.rowconfigure(0, weight=1)

        layer_segments_frame = ttk.Frame(self.tab_layer_segments_export)
        layer_segments_frame.grid(row=0, column=0, sticky="nsew")

        layer_export_frame = ttk.Frame(self.tab_layer_segments_export)
        layer_export_frame.grid(row=0, column=1, sticky="nsew")

        self._build_tab_segments(layer_segments_frame)
        self._build_tab_export(layer_export_frame)

    # =========================================================================
    # TAB 1 – Setup
    # =========================================================================
    def _build_tab_setup(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        mode_lf = ttk.LabelFrame(parent, text="Measurement Mode", padding=10)
        mode_lf.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 0))

        ttk.Radiobutton(mode_lf, text="Vertical Measurements",
                        variable=self.measurement_mode,
                        value="vertical").pack(side=tk.LEFT, padx=12)
        ttk.Radiobutton(mode_lf, text="Horizontal Measurements",
                        variable=self.measurement_mode,
                        value="horizontal").pack(side=tk.LEFT, padx=12)

        controls = ttk.Frame(parent, padding=(8, 6, 8, 6))
        controls.grid(row=1, column=0, sticky="ew")
        controls.columnconfigure(0, weight=3)
        controls.columnconfigure(1, weight=2)

        color_lf = ttk.LabelFrame(controls, text="Color Filtering", padding=10)
        color_lf.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.color_display = tk.Frame(
            color_lf,
            width=34,
            height=34,
            bg=rgb_to_hex(self.selected_color),
            relief=tk.RIDGE,
            bd=2,
        )
        self.color_display.pack(side=tk.LEFT, padx=(0, 8))
        self.color_display.pack_propagate(False)

        pick_color_btn = tk.Button(
            color_lf,
            text="Pick Color",
            command=self.choose_color,
            bg=BTN_BG,
            fg=BTN_TEXT,
            activebackground=BTN_HOVER,
            activeforeground=BTN_TEXT,
            font=("Arial", 9),
            relief=tk.FLAT,
            padx=8,
            cursor="hand2",
        )
        pick_color_btn.pack(side=tk.LEFT, padx=(0, 10))
        self._bind_button_hover(pick_color_btn)

        self.color_label = ttk.Label(color_lf, text=f"RGB{self.selected_color}")
        self.color_label.pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(color_lf, text="Tolerance:").pack(side=tk.LEFT)
        self.tolerance_var = tk.IntVar(value=20)
        tk.Scale(
            color_lf,
            from_=5,
            to=50,
            orient=tk.HORIZONTAL,
            variable=self.tolerance_var,
            length=150,
            bg=BG,
            highlightthickness=0,
            troughcolor="#EBCACA",
        ).pack(side=tk.LEFT, padx=6)
        self.tol_label = ttk.Label(color_lf, text="20", width=3)
        self.tol_label.pack(side=tk.LEFT)
        self.tolerance_var.trace_add(
            "write",
            lambda *_: self.tol_label.config(text=str(self.tolerance_var.get()))
        )

        action_panel = ttk.Frame(controls)
        action_panel.grid(row=0, column=1, sticky="nsew")
        action_panel.columnconfigure(0, weight=1)

        opts_lf = ttk.LabelFrame(action_panel, text="Processing Options", padding=10)
        opts_lf.grid(row=0, column=0, sticky="ew")

        self.crop_mode_var = tk.StringVar(value="none")
        ttk.Label(opts_lf, text="Crop mode:").pack(side=tk.LEFT, padx=(6, 8))
        crop_mode_menu = ttk.Combobox(
            opts_lf,
            textvariable=self.crop_mode_var,
            values=["none", "right", "left"],
            state="readonly",
            width=12,
        )
        crop_mode_menu.pack(side=tk.LEFT, padx=(0, 6))

        process_btn = tk.Button(
            action_panel,
            text="▶  Process Images",
            command=self.process_all,
            bg=BTN_BG,
            fg=BTN_TEXT,
            activebackground=BTN_HOVER,
            activeforeground=BTN_TEXT,
            font=("Arial", 12, "bold"),
            relief=tk.FLAT,
            padx=20,
            pady=9,
            cursor="hand2",
        )
        process_btn.grid(row=1, column=0, sticky="w", pady=(8, 10))
        self._bind_button_hover(process_btn)

        ttk.Label(
            action_panel,
            text="After processing, inspect previews in Setup + Preview.\n"
                 "Use Segments + Export for overlays, statistics, and file export.",
            foreground="#5A3131",
            font=("Arial", 9),
        ).grid(row=2, column=0, sticky="w")

        lists = ttk.Frame(parent, padding=(8, 0, 8, 8))
        lists.grid(row=2, column=0, sticky="nsew")
        lists.columnconfigure(0, weight=1)
        lists.columnconfigure(1, weight=1)
        lists.rowconfigure(2, weight=1)

        ttk.Label(lists, text="Image Files",
              font=("Arial", 11, "bold"), foreground=ACCENT).grid(
              row=0, column=0, sticky="w", pady=(0, 2))
        ttk.Label(lists, text="Processing Status",
              font=("Arial", 11, "bold"), foreground=ACCENT).grid(
              row=0, column=1, sticky="w", padx=(8, 0), pady=(0, 2))
        self._add_section_divider(lists, row=1, column=0, padx=(0, 4), pady=(0, 5))
        self._add_section_divider(lists, row=1, column=1, padx=(8, 0), pady=(0, 5))

        lf = ttk.Frame(lists)
        lf.grid(row=2, column=0, sticky="nsew", padx=(0, 4))
        lf.rowconfigure(0, weight=1)
        lf.columnconfigure(0, weight=1)

        sb = ttk.Scrollbar(lf, orient="vertical")
        self.listbox = tk.Listbox(
            lf,
            yscrollcommand=sb.set,
            font=("Courier", 9),
            bg="white",
            selectbackground=ACCENT,
            selectforeground="white",
            activestyle="none",
            relief=tk.FLAT,
            bd=1,
            exportselection=False,
        )
        sb.config(command=self.listbox.yview)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.listbox.bind("<Double-Button-1>", self.open_image)

        sf = ttk.Frame(lists)
        sf.grid(row=2, column=1, sticky="nsew", padx=(8, 0))
        sf.rowconfigure(0, weight=1)
        sf.columnconfigure(0, weight=1)

        sb2 = ttk.Scrollbar(sf, orient="vertical")
        self.processed_listbox = tk.Listbox(
            sf,
            yscrollcommand=sb2.set,
            font=("Courier", 9),
            bg="white",
            selectbackground=ACCENT,
            selectforeground="white",
            activestyle="none",
            relief=tk.FLAT,
            bd=1,
            exportselection=False,
        )
        sb2.config(command=self.processed_listbox.yview)
        self.processed_listbox.grid(row=0, column=0, sticky="nsew")
        sb2.grid(row=0, column=1, sticky="ns")

    # =========================================================================
    # TAB 2 – Preview
    # =========================================================================
    def _build_tab_preview(self, parent: ttk.Frame) -> None:
        parent.rowconfigure(2, weight=1)
        parent.columnconfigure(0, weight=1)

        ttk.Label(parent, text="Processed Image Preview",
                  font=("Arial", 12, "bold"),
                  foreground=ACCENT).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 4))
        self._add_section_divider(parent, row=1, padx=14, pady=(0, 8))

        cf = ttk.Frame(parent)
        cf.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        cf.rowconfigure(0, weight=1)
        cf.columnconfigure(0, weight=1)

        self.preview_canvas = tk.Canvas(cf, bg=SURFACE, relief=tk.FLAT, bd=0, highlightthickness=1, highlightbackground=BORDER)
        vsb = ttk.Scrollbar(cf, orient="vertical",   command=self.preview_canvas.yview)
        hsb = ttk.Scrollbar(cf, orient="horizontal", command=self.preview_canvas.xview)
        self.preview_canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.preview_frame = ttk.Frame(self.preview_canvas)
        self.preview_canvas_window = self.preview_canvas.create_window(
            (0, 0), window=self.preview_frame, anchor="nw")

        self.preview_frame.bind("<Configure>",  self._on_preview_configure)
        self.preview_canvas.bind("<Configure>", self._on_preview_canvas_resize)

    # =========================================================================
    # TAB 3 – Segments
    # =========================================================================
    def _build_tab_segments(self, parent: ttk.Frame) -> None:
        parent.rowconfigure(2, weight=1)
        parent.columnconfigure(0, weight=1)

        ttk.Label(parent, text="Segmented Regions",
                  font=("Arial", 12, "bold"),
                  foreground=ACCENT).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 4))
        self._add_section_divider(parent, row=1, padx=14, pady=(0, 8))

        cf = ttk.Frame(parent)
        cf.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        cf.rowconfigure(0, weight=1)
        cf.columnconfigure(0, weight=1)

        self.segment_canvas = tk.Canvas(cf, bg=SURFACE, relief=tk.FLAT, bd=0, highlightthickness=1, highlightbackground=BORDER)
        vsb = ttk.Scrollbar(cf, orient="vertical",   command=self.segment_canvas.yview)
        hsb = ttk.Scrollbar(cf, orient="horizontal", command=self.segment_canvas.xview)
        self.segment_canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.segment_canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.segment_frame = ttk.Frame(self.segment_canvas)
        self.segment_canvas_window = self.segment_canvas.create_window(
            (0, 0), window=self.segment_frame, anchor="nw")

        self.segment_frame.bind("<Configure>",  self._on_segment_configure)
        self.segment_canvas.bind("<Configure>", self._on_segment_canvas_resize)

    # =========================================================================
    # TAB 4 – Export
    # =========================================================================
    def _build_tab_export(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(3, weight=1)

        ttk.Label(parent, text="Export & Statistics",
                  font=("Arial", 12, "bold"),
                  foreground=ACCENT).grid(row=0, column=0, sticky="w", padx=14, pady=(14, 0))
        self._add_section_divider(parent, row=1, padx=14, pady=(6, 10))

        # Save section
        save_lf = ttk.LabelFrame(parent, text="Save Results", padding=16)
        save_lf.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 10))

        save_csv_btn = tk.Button(save_lf, text="Save to CSV",
                     command=self.save_to_csv,
                     bg=BTN_BG, fg=BTN_TEXT,
                     activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
                                         font=("Arial", 11), relief=tk.FLAT,
                                         padx=16, pady=9, cursor="hand2")
        save_csv_btn.grid(row=0, column=0, padx=(0, 10), pady=4, sticky="w")
        self._bind_button_hover(save_csv_btn)

        save_excel_btn = tk.Button(save_lf, text="Save to Excel",
                       command=self.save_to_excel,
                       bg=BTN_BG, fg=BTN_TEXT,
                       activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
                                             font=("Arial", 11), relief=tk.FLAT,
                                             padx=16, pady=9, cursor="hand2")
        save_excel_btn.grid(row=0, column=1, padx=(0, 10), pady=4, sticky="w")
        self._bind_button_hover(save_excel_btn)
        ttk.Label(save_lf,
                  text="Saves measurements for the currently selected mode (Vertical / Horizontal).",
                  foreground="#64748B").grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # Statistics section
        stats_lf = ttk.LabelFrame(parent, text="Statistics", padding=16)
        stats_lf.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 10))
        stats_lf.columnconfigure(0, weight=1)
        stats_lf.rowconfigure(1, weight=1)

        stats_btn = tk.Button(stats_lf, text="Calculate Statistics",
                      command=self.show_statistics,
                      bg=BTN_BG, fg=BTN_TEXT,
                      activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
                      font=("Arial", 11), relief=tk.FLAT,
                      padx=16, pady=9, cursor="hand2")
        stats_btn.grid(row=0, column=0, sticky="w", pady=(0, 10))
        self._bind_button_hover(stats_btn)

        self.stats_text = tk.Text(stats_lf, height=12,
                                  font=("Courier", 10),
                                  state=tk.DISABLED,
                      bg=SURFACE, relief=tk.FLAT, bd=1,
                                  wrap=tk.WORD)
        stats_sb = ttk.Scrollbar(stats_lf, orient="vertical",
                                 command=self.stats_text.yview)
        self.stats_text.configure(yscrollcommand=stats_sb.set)
        self.stats_text.grid(row=1, column=0, sticky="nsew")
        stats_sb.grid(row=1, column=1, sticky="ns")

        # Close button
        close_btn = tk.Button(parent, text="Close Application",
                      command=self.root.destroy,
                      bg=BTN_BG, fg=BTN_TEXT,
                      activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
                      font=("Arial", 10), relief=tk.FLAT,
                      padx=12, pady=6,
                      cursor="hand2")
        close_btn.grid(row=4, column=0, sticky="e", padx=20, pady=(0, 12))
        self._bind_button_hover(close_btn)

    # =========================================================================
    # TAB 5 – Ellipsoidal Features
    # =========================================================================
    def _build_tab_ellipsoidal(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=2, minsize=460)
        parent.columnconfigure(1, weight=3)
        parent.rowconfigure(0, weight=1)

        left = ttk.Frame(parent, padding=(12, 12, 8, 12))
        left.grid(row=0, column=0, sticky="nsew")
        left.columnconfigure(0, weight=1)

        right = ttk.Frame(parent, padding=(8, 12, 12, 12))
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        # Color and detection controls
        control_lf = ttk.LabelFrame(left, text="Detection Controls", padding=10)
        control_lf.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(control_lf, text="Ellipsoidal Features:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.ellipsoidal_feature_color_var = tk.StringVar(value="Green")
        feature_menu = ttk.Combobox(
            control_lf,
            textvariable=self.ellipsoidal_feature_color_var,
            values=["Green", "Red", "Blue", "Yellow", "Cyan", "Magenta", "Orange", "Purple"],
            state="readonly",
            width=18,
        )
        feature_menu.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(control_lf, text="Baseline Interface:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.ellipsoidal_baseline_color_var = tk.StringVar(value="Auto")
        baseline_menu = ttk.Combobox(
            control_lf,
            textvariable=self.ellipsoidal_baseline_color_var,
            values=["Auto", "Red", "Blue", "Yellow", "Cyan", "Magenta", "Green", "Orange", "Purple"],
            state="readonly",
            width=18,
        )
        baseline_menu.grid(row=1, column=1, sticky="w", pady=4)

        self.ellipsoidal_measure_interface_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            control_lf,
            text="Measure Interface Distances",
            variable=self.ellipsoidal_measure_interface_var,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=6)

        ttk.Label(control_lf, text="Distance Direction:").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        self.ellipsoidal_distance_direction_var = tk.StringVar(value="Bottom")
        distance_menu = ttk.Combobox(
            control_lf,
            textvariable=self.ellipsoidal_distance_direction_var,
            values=["Top", "Bottom"],
            state="readonly",
            width=18,
        )
        distance_menu.grid(row=3, column=1, sticky="w", pady=4)

        # Data source
        data_lf = ttk.LabelFrame(left, text="Data Source (TIFF + TXT)", padding=10)
        data_lf.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        data_lf.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)

        data_btns = ttk.Frame(data_lf)
        data_btns.grid(row=0, column=0, sticky="ew")
        data_folder_btn = tk.Button(
            data_btns, text="Folder", command=self.choose_ellipsoidal_data_folder,
            bg=BTN_BG, fg=BTN_TEXT, activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
            relief=tk.FLAT, padx=10, pady=4, cursor="hand2",
        )
        data_folder_btn.pack(side=tk.LEFT, padx=(0, 6))
        self._bind_button_hover(data_folder_btn)

        data_file_btn = tk.Button(
            data_btns, text="File", command=self.choose_ellipsoidal_data_file,
            bg=BTN_BG, fg=BTN_TEXT, activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
            relief=tk.FLAT, padx=10, pady=4, cursor="hand2",
        )
        data_file_btn.pack(side=tk.LEFT, padx=(0, 6))
        self._bind_button_hover(data_file_btn)

        self.ellipsoidal_data_summary_var = tk.StringVar(value="No data selected")
        ttk.Label(data_lf, textvariable=self.ellipsoidal_data_summary_var, foreground="#64748B").grid(row=1, column=0, sticky="w", pady=(6, 4))

        data_lists = ttk.Frame(data_lf)
        data_lists.grid(row=2, column=0, sticky="nsew")
        data_lists.columnconfigure(0, weight=1)
        data_lists.columnconfigure(1, weight=1)
        data_lf.rowconfigure(2, weight=1)

        tiff_frame = ttk.Frame(data_lists)
        tiff_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        ttk.Label(tiff_frame, text="TIFF Files", font=("Arial", 9, "bold")).pack(anchor="w")
        tiff_scroll = ttk.Scrollbar(tiff_frame, orient="vertical")
        self.ellipsoidal_tiff_listbox = tk.Listbox(tiff_frame, yscrollcommand=tiff_scroll.set, height=7, exportselection=False)
        tiff_scroll.config(command=self.ellipsoidal_tiff_listbox.yview)
        tiff_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.ellipsoidal_tiff_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        txt_frame = ttk.Frame(data_lists)
        txt_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        ttk.Label(txt_frame, text="TXT PixelSize", font=("Arial", 9, "bold")).pack(anchor="w")
        txt_scroll = ttk.Scrollbar(txt_frame, orient="vertical")
        self.ellipsoidal_txt_listbox = tk.Listbox(txt_frame, yscrollcommand=txt_scroll.set, height=7, exportselection=False)
        txt_scroll.config(command=self.ellipsoidal_txt_listbox.yview)
        txt_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.ellipsoidal_txt_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Mask source
        mask_lf = ttk.LabelFrame(left, text="Mask Source", padding=10)
        mask_lf.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        mask_lf.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)

        mask_btns = ttk.Frame(mask_lf)
        mask_btns.grid(row=0, column=0, sticky="ew")
        mask_folder_btn = tk.Button(
            mask_btns, text="Folder", command=self.choose_ellipsoidal_mask_folder,
            bg=BTN_BG, fg=BTN_TEXT, activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
            relief=tk.FLAT, padx=10, pady=4, cursor="hand2",
        )
        mask_folder_btn.pack(side=tk.LEFT, padx=(0, 6))
        self._bind_button_hover(mask_folder_btn)

        mask_file_btn = tk.Button(
            mask_btns, text="File", command=self.choose_ellipsoidal_mask_file,
            bg=BTN_BG, fg=BTN_TEXT, activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
            relief=tk.FLAT, padx=10, pady=4, cursor="hand2",
        )
        mask_file_btn.pack(side=tk.LEFT, padx=(0, 6))
        self._bind_button_hover(mask_file_btn)

        self.ellipsoidal_mask_summary_var = tk.StringVar(value="No mask selected")
        ttk.Label(mask_lf, textvariable=self.ellipsoidal_mask_summary_var, foreground="#64748B").grid(row=1, column=0, sticky="w", pady=(6, 4))

        mask_list_frame = ttk.Frame(mask_lf)
        mask_list_frame.grid(row=2, column=0, sticky="nsew")
        mask_lf.rowconfigure(2, weight=1)
        mask_scroll = ttk.Scrollbar(mask_list_frame, orient="vertical")
        self.ellipsoidal_mask_listbox = tk.Listbox(mask_list_frame, yscrollcommand=mask_scroll.set, height=7, exportselection=False)
        mask_scroll.config(command=self.ellipsoidal_mask_listbox.yview)
        mask_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.ellipsoidal_mask_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Actions
        action_frame = ttk.Frame(left)
        action_frame.grid(row=3, column=0, sticky="ew", pady=(0, 8))

        process_btn = tk.Button(
            action_frame, text="Process Ellipsoidal Features", command=self.process_ellipsoidal,
            bg=BTN_BG, fg=BTN_TEXT, activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
            relief=tk.FLAT, padx=14, pady=7, cursor="hand2", font=("Arial", 10, "bold"),
        )
        process_btn.pack(side=tk.LEFT, padx=(0, 6))
        self._bind_button_hover(process_btn)

        export_csv_btn = tk.Button(
            action_frame, text="Export CSV", command=self.export_ellipsoidal_csv,
            bg=BTN_BG, fg=BTN_TEXT, activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
            relief=tk.FLAT, padx=14, pady=7, cursor="hand2",
        )
        export_csv_btn.pack(side=tk.LEFT, padx=(0, 6))
        self._bind_button_hover(export_csv_btn)

        export_excel_btn = tk.Button(
            action_frame, text="Export Excel", command=self.export_ellipsoidal_excel,
            bg=BTN_BG, fg=BTN_TEXT, activebackground=BTN_HOVER, activeforeground=BTN_TEXT,
            relief=tk.FLAT, padx=14, pady=7, cursor="hand2",
        )
        export_excel_btn.pack(side=tk.LEFT)
        self._bind_button_hover(export_excel_btn)

        self.ellipsoidal_status_var = tk.StringVar(value="Ready")
        ttk.Label(left, textvariable=self.ellipsoidal_status_var, foreground="#0F766E").grid(row=4, column=0, sticky="w")

        # Right preview panel
        ttk.Label(
            right,
            text="Ellipsoidal Features Preview (Original, Baseline, Features)",
            font=("Arial", 11, "bold"),
            foreground=ACCENT,
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))
        self._add_section_divider(right, row=1, pady=(0, 8))

        preview_wrap = ttk.Frame(right)
        preview_wrap.grid(row=2, column=0, sticky="nsew")
        preview_wrap.rowconfigure(0, weight=1)
        preview_wrap.columnconfigure(0, weight=1)

        self.ellipsoidal_preview_canvas = tk.Canvas(preview_wrap, bg=SURFACE, relief=tk.FLAT, bd=0, highlightthickness=1, highlightbackground=BORDER)
        ellipsoidal_scroll = ttk.Scrollbar(preview_wrap, orient="vertical", command=self.ellipsoidal_preview_canvas.yview)
        self.ellipsoidal_preview_canvas.configure(yscrollcommand=ellipsoidal_scroll.set)
        self.ellipsoidal_preview_canvas.grid(row=0, column=0, sticky="nsew")
        ellipsoidal_scroll.grid(row=0, column=1, sticky="ns")

        self.ellipsoidal_preview_frame = ttk.Frame(self.ellipsoidal_preview_canvas)
        self.ellipsoidal_preview_window = self.ellipsoidal_preview_canvas.create_window((0, 0), window=self.ellipsoidal_preview_frame, anchor="nw")
        self.ellipsoidal_preview_frame.bind("<Configure>", self._on_ellipsoidal_preview_configure)
        self.ellipsoidal_preview_canvas.bind("<Configure>", self._on_ellipsoidal_preview_canvas_resize)

    def _set_listbox_items(self, listbox: tk.Listbox, items: List[str]) -> None:
        listbox.delete(0, tk.END)
        for item in items:
            listbox.insert(tk.END, item)

    def choose_ellipsoidal_data_folder(self) -> None:
        path = select_folder("Select folder containing TIFF/TXT")
        if not path:
            return
        self.ellipsoidal_data_folder_path = path
        tiffs, txts = categorize_data_files(path, recursive=True)
        self.ellipsoidal_txt_rel_paths = txts
        self.ellipsoidal_pixel_size_map = build_pixel_size_map(path, txts)

        self._set_listbox_items(self.ellipsoidal_tiff_listbox, tiffs)
        txt_display: List[str] = []
        found_count = 0
        for rel in txts:
            nm, _um = self.ellipsoidal_pixel_size_map.get(rel, (None, None))
            if nm is not None:
                txt_display.append(f"{rel} - {nm:.3f} nm")
                found_count += 1
            else:
                txt_display.append(f"{rel} - N/A")
        self._set_listbox_items(self.ellipsoidal_txt_listbox, txt_display)
        self.ellipsoidal_data_summary_var.set(f"TIFFs: {len(tiffs)} | TXTs: {len(txts)} (Found PixelSize: {found_count})")

    def choose_ellipsoidal_data_file(self) -> None:
        file_path = select_file(
            title="Select TIFF or TXT file",
            filetypes=[("TIFF/TXT", "*.tif *.tiff *.txt"), ("All files", "*.*")],
        )
        if not file_path:
            return

        self.ellipsoidal_data_folder_path = os.path.dirname(file_path)
        name = os.path.basename(file_path)
        if name.lower().endswith((".tif", ".tiff")):
            self._set_listbox_items(self.ellipsoidal_tiff_listbox, [name])
            self._set_listbox_items(self.ellipsoidal_txt_listbox, [])
            self.ellipsoidal_txt_rel_paths = []
            self.ellipsoidal_pixel_size_map = {}
            self.ellipsoidal_data_summary_var.set(f"Single TIFF: {name}")
        else:
            self._set_listbox_items(self.ellipsoidal_tiff_listbox, [])
            self.ellipsoidal_txt_rel_paths = [name]
            self.ellipsoidal_pixel_size_map = build_pixel_size_map(self.ellipsoidal_data_folder_path, self.ellipsoidal_txt_rel_paths)
            nm, _um = self.ellipsoidal_pixel_size_map.get(name, (None, None))
            label = f"{name} - {nm:.3f} nm" if nm is not None else f"{name} - N/A"
            self._set_listbox_items(self.ellipsoidal_txt_listbox, [label])
            self.ellipsoidal_data_summary_var.set(f"Single TXT: {name}")

    def choose_ellipsoidal_mask_folder(self) -> None:
        path = select_folder("Select folder containing mask images")
        if not path:
            return
        self.ellipsoidal_mask_folder_path = path
        self.ellipsoidal_mask_files = categorize_mask_files(path)
        self._set_listbox_items(self.ellipsoidal_mask_listbox, self.ellipsoidal_mask_files)
        self.ellipsoidal_mask_summary_var.set(f"Masks: {len(self.ellipsoidal_mask_files)}")

    def choose_ellipsoidal_mask_file(self) -> None:
        file_path = select_file(
            title="Select mask image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp"), ("All files", "*.*")],
        )
        if not file_path:
            return
        self.ellipsoidal_mask_folder_path = os.path.dirname(file_path)
        self.ellipsoidal_mask_files = [os.path.basename(file_path)]
        self._set_listbox_items(self.ellipsoidal_mask_listbox, self.ellipsoidal_mask_files)
        self.ellipsoidal_mask_summary_var.set(f"Single mask: {self.ellipsoidal_mask_files[0]}")

    def _lookup_default_pixel_size(self) -> Optional[float]:
        if len(self.ellipsoidal_pixel_size_map) == 1:
            for _rel, (nm, _um) in self.ellipsoidal_pixel_size_map.items():
                if nm is not None:
                    return nm
        return None

    def process_ellipsoidal(self) -> None:
        if not self.ellipsoidal_mask_folder_path or not self.ellipsoidal_mask_files:
            messagebox.showwarning("Missing Selection", "Please select mask source.")
            return

        feature_color = self.ellipsoidal_feature_color_var.get().strip().lower()
        baseline_color = self.ellipsoidal_baseline_color_var.get().strip().lower()
        measure_interface = self.ellipsoidal_measure_interface_var.get()
        distance_direction = self.ellipsoidal_distance_direction_var.get().strip().lower()

        basename_to_pix_nm: Dict[str, Optional[float]] = {}
        for rel, (nm, _um) in self.ellipsoidal_pixel_size_map.items():
            base = os.path.splitext(os.path.basename(rel))[0]
            basename_to_pix_nm.setdefault(base, nm)
        default_pix_nm = self._lookup_default_pixel_size()

        rows: List[Dict[str, object]] = []
        interface_rows: List[Dict[str, object]] = []
        total_features = 0

        last_original = None
        last_baseline = None
        last_features = None
        last_name = ""

        for mask_name in self.ellipsoidal_mask_files:
            full_mask_path = os.path.join(self.ellipsoidal_mask_folder_path, mask_name)
            base = os.path.splitext(mask_name)[0]
            pix_nm = basename_to_pix_nm.get(base, default_pix_nm)

            mask_rows, original, baseline_viz, features_viz, interface_distance_rows = measure_mask_image(
                full_mask_path,
                pix_nm,
                feature_color=feature_color,
                baseline_color=baseline_color,
                measure_interface_distances=measure_interface,
                distance_direction=distance_direction,
            )

            for row in mask_rows:
                row["mask_file"] = mask_name
                row["PixelSize (nm/px)"] = pix_nm if pix_nm is not None else ""

            for row in interface_distance_rows:
                row["mask_file"] = mask_name
                row["PixelSize (nm/px)"] = pix_nm if pix_nm is not None else ""

            rows.extend(mask_rows)
            interface_rows.extend(interface_distance_rows)
            total_features += len(mask_rows)

            if original is not None and baseline_viz is not None and features_viz is not None:
                last_original = original
                last_baseline = baseline_viz
                last_features = features_viz
                last_name = mask_name

        self.ellipsoidal_processed_rows = rows
        self.ellipsoidal_interface_rows = interface_rows

        if last_original is not None and last_baseline is not None and last_features is not None:
            self._update_ellipsoidal_preview(last_name, last_original, last_baseline, last_features)

        msg = f"Processed {len(self.ellipsoidal_mask_files)} mask(s) with {total_features} features."
        if measure_interface:
            msg += f" Interface points: {len(interface_rows)}."
        self.ellipsoidal_status_var.set(msg)
        self.notebook.select(self.tab_ellipsoidal)

    def _update_ellipsoidal_preview(self, mask_name: str, original: np.ndarray, baseline_viz: np.ndarray, features_viz: np.ndarray) -> None:
        for w in self.ellipsoidal_preview_frame.winfo_children():
            w.destroy()
        self.ellipsoidal_last_preview_refs.clear()

        ttk.Label(self.ellipsoidal_preview_frame, text=f"File: {mask_name}", font=("Arial", 11, "bold")).pack(anchor="w", padx=8, pady=(8, 6))

        def _render(label_text: str, image_bgr: np.ndarray) -> None:
            ttk.Label(self.ellipsoidal_preview_frame, text=label_text, font=("Arial", 10, "bold")).pack(anchor="w", padx=8, pady=(6, 4))
            max_width = 700
            h, w = image_bgr.shape[:2]
            if w > max_width:
                scale = max_width / w
                resized = cv2.resize(image_bgr, (int(w * scale), int(h * scale)))
            else:
                resized = image_bgr
            rgb = cv2.cvtColor(resized, cv2.COLOR_BGR2RGB)
            photo = ImageTk.PhotoImage(Image.fromarray(rgb))
            panel = tk.Label(self.ellipsoidal_preview_frame, image=photo, bg="white")
            panel.image = photo  # type: ignore[attr-defined]
            panel.pack(anchor="w", padx=8, pady=(0, 6))
            self.ellipsoidal_last_preview_refs.append(photo)

        _render("Original", original)
        _render("Baseline Detection", baseline_viz)
        _render("Ellipsoidal Feature Detection", features_viz)

    def export_ellipsoidal_csv(self) -> None:
        if not self.ellipsoidal_processed_rows:
            messagebox.showwarning("No Data", "Process ellipsoidal features first.")
            return

        out_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Save Ellipsoidal Measurements CSV",
        )
        if not out_path:
            return

        try:
            write_measurements_csv(out_path, self.ellipsoidal_processed_rows)
            if self.ellipsoidal_interface_rows:
                iface_path = filedialog.asksaveasfilename(
                    defaultextension=".csv",
                    filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                    initialfile=f"{os.path.splitext(os.path.basename(out_path))[0]}_interface_distances.csv",
                    title="Save Interface Distances CSV",
                )
                if not iface_path:
                    messagebox.showinfo("Saved", f"Feature measurements:\n{out_path}\n\nInterface distances were not exported.")
                    return
                write_interface_distances_csv(iface_path, self.ellipsoidal_interface_rows)
                messagebox.showinfo("Saved", f"Feature measurements:\n{out_path}\n\nInterface distances:\n{iface_path}")
            else:
                messagebox.showinfo("Saved", f"Feature measurements:\n{out_path}")
        except Exception as exc:
            messagebox.showerror("Export Error", f"Failed to export CSV:\n{exc}")

    def export_ellipsoidal_excel(self) -> None:
        if not self.ellipsoidal_processed_rows:
            messagebox.showwarning("No Data", "Process ellipsoidal features first.")
            return

        out_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Ellipsoidal Measurements Excel",
        )
        if not out_path:
            return

        try:
            write_measurements_excel(out_path, self.ellipsoidal_processed_rows)
            if self.ellipsoidal_interface_rows:
                iface_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    initialfile=f"{os.path.splitext(os.path.basename(out_path))[0]}_interface_distances.xlsx",
                    title="Save Interface Distances Excel",
                )
                if not iface_path:
                    messagebox.showinfo("Saved", f"Feature measurements:\n{out_path}\n\nInterface distances were not exported.")
                    return
                write_interface_distances_excel(iface_path, self.ellipsoidal_interface_rows)
                messagebox.showinfo("Saved", f"Feature measurements:\n{out_path}\n\nInterface distances:\n{iface_path}")
            else:
                messagebox.showinfo("Saved", f"Feature measurements:\n{out_path}")
        except Exception as exc:
            messagebox.showerror("Export Error", f"Failed to export Excel:\n{exc}")

    # =========================================================================
    # Business logic
    # =========================================================================
    def choose_folder(self) -> None:
        path = select_folder()
        if not path:
            return
        self.folder_path = path
        self.folder_var.set(path)
        self.images_list = find_images_recursively(path)
        self.images_var.set(f"Indexed {len(self.images_list)} images")

        self.listbox.delete(0, tk.END)
        for img in self.images_list:
            self.listbox.insert(tk.END, img)

        self.processed_listbox.delete(0, tk.END)
        self.clear_previews()

    def choose_color(self) -> None:
        color = colorchooser.askcolor(
            title="Choose target color to filter",
            initialcolor=rgb_to_hex(self.selected_color),
        )
        if color[0]:
            self.selected_color = tuple(int(c) for c in color[0])
            self.color_display.config(bg=rgb_to_hex(self.selected_color))
            self.color_label.config(text=f"RGB{self.selected_color}")

    def open_image(self, event) -> None:
        sel = self.listbox.curselection()
        if not sel:
            return
        img_name = self.listbox.get(sel[0])
        if not os.path.isdir(self.folder_path):
            return
        img_path = os.path.join(self.folder_path, img_name)
        if os.path.exists(img_path):
            os.system(f'open "{img_path}"')

    def process_all(self) -> None:
        if not os.path.isdir(self.folder_path):
            messagebox.showerror("Error", "No valid folder selected.")
            return

        crop_mode = self.crop_mode_var.get()

        self.processed_images = process_images(
            self.folder_path,
            self.images_list,
            crop_right_half=(crop_mode == "right"),
            crop_left_half=(crop_mode == "left"),
            target_color=self.selected_color,
            tolerance=self.tolerance_var.get(),
        )

        self.processed_listbox.delete(0, tk.END)
        success = 0
        for i, (name, bw) in enumerate(zip(self.images_list, self.processed_images)):
            if bw is not None:
                self.processed_listbox.insert(tk.END, f"✓  {name}")
                self.processed_listbox.itemconfig(i, fg=SUCCESS)
                success += 1
            else:
                self.processed_listbox.insert(tk.END, f"✗  {name}")
                self.processed_listbox.itemconfig(i, fg=DANGER)

        self.images_var.set(
            f"✓ {success}/{len(self.processed_images)} images processed successfully"
        )
        self.generate_previews()
        # Jump to Layer Measurements -> Setup + Preview automatically
        self.notebook.select(self.tab_layer_measurements)
        self.layer_notebook.select(self.tab_layer_setup_preview)

    def generate_previews(self) -> None:
        self.clear_previews()
        orientation = self.measurement_mode.get()
        p_col = p_row = s_col = s_row = 0

        for name, bw in zip(self.images_list, self.processed_images):
            if bw is None:
                continue

            # Processed thumbnail
            tf = ttk.Frame(self.preview_frame, relief="ridge")
            tf.grid(row=p_row, column=p_col, padx=8, pady=8, sticky="n")
            p_col += 1
            if p_col >= THUMB_COLS:
                p_col = 0
                p_row += 1

            ttk.Label(tf, text=name, font=("Arial", 8),
                      wraplength=200).pack(pady=4, padx=4)
            pil = Image.fromarray(bw)
            pil.thumbnail((200, 200))
            tkimg = ImageTk.PhotoImage(pil)
            lbl = tk.Label(tf, image=tkimg, bg="white")
            lbl.image = tkimg  # type: ignore[attr-defined]
            lbl.pack(padx=4, pady=(0, 4))
            self.thumbnail_refs.append(tkimg)

            # Segmented thumbnail
            sf = ttk.Frame(self.segment_frame, relief="ridge")
            sf.grid(row=s_row, column=s_col, padx=8, pady=8, sticky="n")
            s_col += 1
            if s_col >= THUMB_COLS:
                s_col = 0
                s_row += 1

            ttk.Label(sf, text=name, font=("Arial", 8),
                      wraplength=200).pack(pady=4, padx=4)
            seg_rgb = draw_segments_on_image(bw, orientation=orientation)
            seg_pil = Image.fromarray(seg_rgb)
            seg_pil.thumbnail((200, 200))
            seg_tkimg = ImageTk.PhotoImage(seg_pil)
            seg_lbl = tk.Label(sf, image=seg_tkimg, bg="white")
            seg_lbl.image = seg_tkimg  # type: ignore[attr-defined]
            seg_lbl.pack(padx=4, pady=(0, 4))
            self.thumbnail_refs.append(seg_tkimg)

    def clear_previews(self) -> None:
        for w in self.preview_frame.winfo_children():
            w.destroy()
        for w in self.segment_frame.winfo_children():
            w.destroy()
        self.thumbnail_refs.clear()

    def save_to_csv(self) -> None:
        if not self.processed_images:
            messagebox.showwarning("No Data", "Please process images first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Save Measurements as CSV",
        )
        if not path:
            return
        try:
            if self.measurement_mode.get() == "vertical":
                save_vertical_segments_to_csv(self.images_list, self.processed_images, path)
            else:
                save_horizontal_segments_to_csv(self.images_list, self.processed_images, path)
            messagebox.showinfo("Saved", f"CSV saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save CSV:\n{e}")

    def save_to_excel(self) -> None:
        if not self.processed_images:
            messagebox.showwarning("No Data", "Please process images first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Measurements as Excel",
        )
        if not path:
            return
        try:
            if self.measurement_mode.get() == "vertical":
                save_vertical_segments_to_excel(self.images_list, self.processed_images, path)
            else:
                save_horizontal_segments_to_excel(self.images_list, self.processed_images, path)
            messagebox.showinfo("Saved", f"Excel saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel:\n{e}")

    def show_statistics(self) -> None:
        if not self.processed_images:
            messagebox.showwarning("No Data", "Please process images first.")
            return

        all_lengths: list = []
        for bw in self.processed_images:
            if bw is None:
                continue
            if self.measurement_mode.get() == "vertical":
                segments = analyze_all_vertical_segments(bw)
            else:
                segments = analyze_all_horizontal_segments(bw)
            for seg_list in segments.values():
                for start, end in seg_list:
                    all_lengths.append(end - start + 1)

        if not all_lengths:
            text = "No segments found in the processed images."
        else:
            mode = self.measurement_mode.get().capitalize()
            text = (
                f"Measurement Statistics — {mode}\n"
                f"{'─' * 42}\n"
                f"Total Segments :  {len(all_lengths)}\n"
                f"Min Length     :  {min(all_lengths):.2f} px\n"
                f"Max Length     :  {max(all_lengths):.2f} px\n"
                f"Mean Length    :  {np.mean(all_lengths):.2f} px\n"
                f"Median Length  :  {np.median(all_lengths):.2f} px\n"
                f"Std Deviation  :  {np.std(all_lengths):.2f} px\n"
            )

        self.stats_text.config(state=tk.NORMAL)
        self.stats_text.delete("1.0", tk.END)
        self.stats_text.insert(tk.END, text)
        self.stats_text.config(state=tk.DISABLED)

    def check_for_updates(self) -> None:
        """Check GitHub releases and offer to open the latest download."""
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            result = self.update_manager.check_for_updates()
        finally:
            self.root.config(cursor="")

        error = result.get("error")
        if error:
            messagebox.showerror("Update Check Failed", str(error))
            return

        if result.get("update_available"):
            latest_version = result.get("latest_version")
            current_version = result.get("current_version")
            platform_name = result.get("platform")
            download_name = result.get("download_name")
            release_notes = str(result.get("release_notes") or "No release notes available.")
            if len(release_notes) > 1200:
                release_notes = f"{release_notes[:1200]}\n\n..."
            installer_line = ""
            if isinstance(download_name, str) and download_name:
                installer_line = f"Installer for {platform_name}: {download_name}\n\n"
            should_open = messagebox.askyesno(
                "Update Available",
                f"Current version: {current_version}\n"
                f"Latest version: {latest_version}\n\n"
                f"{installer_line}"
                f"Release notes:\n{release_notes}\n\n"
                "Open the download page now?",
            )
            if should_open:
                download_url = result.get("download_url")
                html_url = result.get("html_url")
                target_url = download_url if isinstance(download_url, str) else None
                if target_url is None and isinstance(html_url, str):
                    target_url = html_url
                self.update_manager.open_download_page(target_url)
            return

        messagebox.showinfo(
            "No Updates Available",
            f"You are running the latest version ({result.get('current_version')}).",
        )

    def _auto_check_for_updates(self) -> None:
        """Start a silent, non-blocking update check after the UI has loaded."""
        if self.startup_update_check_started:
            return
        self.startup_update_check_started = True

        worker = threading.Thread(target=self._run_startup_update_check, daemon=True)
        worker.start()

    def _run_startup_update_check(self) -> None:
        """Perform the startup update request off the Tkinter UI thread."""
        result = self.update_manager.check_for_updates()
        self.root.after(0, lambda: self._handle_startup_update_result(result))

    def _handle_startup_update_result(self, result: Dict[str, object]) -> None:
        """Show an update prompt only when a newer release is available."""
        if result.get("error"):
            return
        if not result.get("update_available"):
            return

        latest_version = result.get("latest_version")
        current_version = result.get("current_version")
        platform_name = result.get("platform")
        download_name = result.get("download_name")
        release_notes = str(result.get("release_notes") or "No release notes available.")
        if len(release_notes) > 1200:
            release_notes = f"{release_notes[:1200]}\n\n..."
        installer_line = ""
        if isinstance(download_name, str) and download_name:
            installer_line = f"Installer for {platform_name}: {download_name}\n\n"

        should_open = messagebox.askyesno(
            "Update Available",
            f"Current version: {current_version}\n"
            f"Latest version: {latest_version}\n\n"
            f"{installer_line}"
            f"Release notes:\n{release_notes}\n\n"
            "Open the download page now?",
        )
        if should_open:
            download_url = result.get("download_url")
            html_url = result.get("html_url")
            target_url = download_url if isinstance(download_url, str) else None
            if target_url is None and isinstance(html_url, str):
                target_url = html_url
            self.update_manager.open_download_page(target_url)

    # ── Scroll helpers ────────────────────────────────────────────────────────
    def _on_preview_configure(self, event) -> None:
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))

    def _on_preview_canvas_resize(self, event) -> None:
        self.preview_canvas.itemconfig(self.preview_canvas_window, width=event.width)

    def _on_segment_configure(self, event) -> None:
        self.segment_canvas.configure(scrollregion=self.segment_canvas.bbox("all"))

    def _on_segment_canvas_resize(self, event) -> None:
        self.segment_canvas.itemconfig(self.segment_canvas_window, width=event.width)

    def _on_ellipsoidal_preview_configure(self, event) -> None:
        self.ellipsoidal_preview_canvas.configure(scrollregion=self.ellipsoidal_preview_canvas.bbox("all"))

    def _on_ellipsoidal_preview_canvas_resize(self, event) -> None:
        self.ellipsoidal_preview_canvas.itemconfig(self.ellipsoidal_preview_window, width=event.width)

    def _on_mousewheel(self, event) -> None:
        """Route mouse-wheel to whichever scrollable tab is active."""
        tab = self.notebook.select()
        if event.num == 4:
            delta = -1
        elif event.num == 5:
            delta = 1
        else:
            raw = event.delta
            delta = int(-1 * raw / 120) if abs(raw) >= 120 else int(-raw)

        if tab == str(self.tab_layer_measurements):
            layer_tab = self.layer_notebook.select()
            if layer_tab == str(self.tab_layer_setup_preview):
                self.preview_canvas.yview_scroll(delta, "units")
            elif layer_tab == str(self.tab_layer_segments_export):
                self.segment_canvas.yview_scroll(delta, "units")
        elif tab == str(self.tab_ellipsoidal):
            self.ellipsoidal_preview_canvas.yview_scroll(delta, "units")


# ── Entry point ───────────────────────────────────────────────────────────────
def main() -> None:
    root = tk.Tk()
    UnifiedMeasurementApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
