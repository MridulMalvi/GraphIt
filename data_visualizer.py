"""
GraphIt
=======
A modern, feature-rich desktop data visualization application built with
Tkinter, Pandas, and Matplotlib.

Features:
    - CSV / Excel (.xlsx) file upload with validation
    - Scrollable data preview (first 10 rows) and dataset info
    - Dynamic X/Y column selection from uploaded data
    - Six chart types: Line, Bar, Scatter, Histogram, Pie, Box Plot
    - Column-based data filtering with row-count feedback
    - Interactive Matplotlib charts embedded in the GUI (mplcursors tooltips)
    - Save charts as PNG/JPG and export filtered data as CSV
    - Dark mode toggle, reset, status bar, column search, summary statistics
    - Comprehensive error handling with user-friendly popups

Installation:
    pip install pandas matplotlib openpyxl mplcursors

Usage:
    python data_visualizer_pro.py
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional

import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure

try:
    import mplcursors
    HAS_MPLCURSORS = True
except ImportError:
    HAS_MPLCURSORS = False


# ---------------------------------------------------------------------------
# Color Palettes
# ---------------------------------------------------------------------------

LIGHT_THEME = {
    "bg": "#F0F2F5",
    "fg": "#1A1A2E",
    "frame_bg": "#FFFFFF",
    "accent": "#4361EE",
    "accent_hover": "#3A56D4",
    "accent_fg": "#FFFFFF",
    "entry_bg": "#F8F9FA",
    "entry_fg": "#1A1A2E",
    "entry_border": "#DEE2E6",
    "table_bg": "#FFFFFF",
    "table_fg": "#1A1A2E",
    "table_selected": "#4361EE",
    "table_heading_bg": "#E9ECEF",
    "table_heading_fg": "#495057",
    "status_bg": "#E9ECEF",
    "status_fg": "#495057",
    "separator": "#DEE2E6",
    "chart_bg": "#FFFFFF",
    "chart_face": "#FAFAFA",
    "label_fg": "#495057",
    "danger": "#E63946",
    "success": "#2A9D8F",
    "warning": "#E9C46A",
}

DARK_THEME = {
    "bg": "#0F0F1A",
    "fg": "#E0E0E0",
    "frame_bg": "#1A1A2E",
    "accent": "#7B68EE",
    "accent_hover": "#6A5ACD",
    "accent_fg": "#FFFFFF",
    "entry_bg": "#16213E",
    "entry_fg": "#E0E0E0",
    "entry_border": "#2A2A4A",
    "table_bg": "#16213E",
    "table_fg": "#E0E0E0",
    "table_selected": "#7B68EE",
    "table_heading_bg": "#1A1A2E",
    "table_heading_fg": "#A0A0C0",
    "status_bg": "#16213E",
    "status_fg": "#8888AA",
    "separator": "#2A2A4A",
    "chart_bg": "#1A1A2E",
    "chart_face": "#0F0F1A",
    "label_fg": "#8888AA",
    "danger": "#FF6B6B",
    "success": "#4ECDC4",
    "warning": "#FFE66D",
}

# Matplotlib chart color cycle (vibrant palette)
CHART_COLORS = [
    "#4361EE", "#F72585", "#4CC9F0", "#7209B7",
    "#3A0CA3", "#F94144", "#F8961E", "#90BE6D",
    "#43AA8B", "#577590",
]


class DataVisualizerApp:
    """Main application class for GraphIt."""

    # Supported chart types
    CHART_TYPES = ["Line Chart", "Bar Chart", "Scatter Plot", "Histogram", "Pie Chart", "Box Plot"]

    def __init__(self, root: tk.Tk) -> None:
        """Initialise the application, build the GUI, and bind events."""
        self.root = root
        self.root.title("GraphIt")
        self.root.geometry("1000x700")
        self.root.minsize(900, 650)

        # State
        self.df: Optional[pd.DataFrame] = None          # Original dataframe
        self.filtered_df: Optional[pd.DataFrame] = None  # Filtered view
        self.current_file: Optional[str] = None
        self.dark_mode: bool = False
        self.theme = LIGHT_THEME.copy()

        # Matplotlib figure (created once, reused)
        self.figure: Optional[Figure] = None
        self.canvas: Optional[FigureCanvasTkAgg] = None
        self.toolbar: Optional[NavigationToolbar2Tk] = None

        # Build UI
        self._configure_styles()
        self._build_menu()
        self._build_toolbar_frame()
        self._build_main_area()
        self._build_status_bar()

        self._apply_theme()
        self._set_status("Ready — upload a CSV or Excel file to begin.")

    # ------------------------------------------------------------------
    # Theme & Style helpers
    # ------------------------------------------------------------------

    def _configure_styles(self) -> None:
        """Configure ttk styles."""
        self.style = ttk.Style()
        self.style.theme_use("clam")

    def _apply_theme(self) -> None:
        """Apply the current colour theme to every widget."""
        t = self.theme

        # Root
        self.root.configure(bg=t["bg"])

        # ttk styles
        self.style.configure("TFrame", background=t["bg"])
        self.style.configure("Card.TFrame", background=t["frame_bg"])
        self.style.configure("TLabel", background=t["bg"], foreground=t["fg"])
        self.style.configure("Card.TLabel", background=t["frame_bg"], foreground=t["fg"])
        self.style.configure("Heading.TLabel", background=t["bg"], foreground=t["accent"],
                             font=("Segoe UI", 14, "bold"))
        self.style.configure("SubHeading.TLabel", background=t["frame_bg"], foreground=t["label_fg"],
                             font=("Segoe UI", 9))
        self.style.configure("Info.TLabel", background=t["frame_bg"], foreground=t["label_fg"],
                             font=("Segoe UI", 9))
        self.style.configure("Status.TLabel", background=t["status_bg"], foreground=t["status_fg"],
                             font=("Segoe UI", 9))

        # Buttons
        self.style.configure("Accent.TButton", background=t["accent"], foreground=t["accent_fg"],
                             font=("Segoe UI", 9, "bold"), padding=(12, 4))
        self.style.map("Accent.TButton",
                       background=[("active", t["accent_hover"]), ("pressed", t["accent_hover"])])

        self.style.configure("TButton", background=t["frame_bg"], foreground=t["fg"],
                             font=("Segoe UI", 9), padding=(10, 4))
        self.style.map("TButton",
                       background=[("active", t["entry_bg"])])

        self.style.configure("Danger.TButton", background=t["danger"], foreground="#FFFFFF",
                             font=("Segoe UI", 9, "bold"), padding=(10, 4))
        self.style.map("Danger.TButton",
                       background=[("active", t["danger"])])

        # Combobox
        self.style.configure("TCombobox", fieldbackground=t["entry_bg"], background=t["entry_bg"],
                             foreground=t["entry_fg"], selectbackground=t["accent"],
                             selectforeground=t["accent_fg"])
        self.style.map("TCombobox", fieldbackground=[("readonly", t["entry_bg"])])

        # Entry
        self.style.configure("TEntry", fieldbackground=t["entry_bg"], foreground=t["entry_fg"])

        # Treeview (data table)
        self.style.configure("Treeview",
                             background=t["table_bg"],
                             foreground=t["table_fg"],
                             fieldbackground=t["table_bg"],
                             rowheight=24,
                             font=("Segoe UI", 9))
        self.style.configure("Treeview.Heading",
                             background=t["table_heading_bg"],
                             foreground=t["table_heading_fg"],
                             font=("Segoe UI", 9, "bold"))
        self.style.map("Treeview",
                       background=[("selected", t["table_selected"])],
                       foreground=[("selected", t["accent_fg"])])

        # Separator
        self.style.configure("TSeparator", background=t["separator"])

        # Notebook
        self.style.configure("TNotebook", background=t["bg"])
        self.style.configure("TNotebook.Tab", background=t["frame_bg"], foreground=t["fg"],
                             padding=(14, 6), font=("Segoe UI", 9))
        self.style.map("TNotebook.Tab",
                       background=[("selected", t["accent"])],
                       foreground=[("selected", t["accent_fg"])])

        # Status bar
        self.style.configure("StatusBar.TFrame", background=t["status_bg"])

        # Chart colours
        if self.figure is not None:
            self.figure.patch.set_facecolor(t["chart_bg"])
            for ax in self.figure.axes:
                ax.set_facecolor(t["chart_face"])
            if self.canvas:
                self.canvas.draw_idle()

    def _toggle_dark_mode(self) -> None:
        """Switch between light and dark themes."""
        self.dark_mode = not self.dark_mode
        self.theme = DARK_THEME.copy() if self.dark_mode else LIGHT_THEME.copy()
        self._apply_theme()
        mode = "Dark" if self.dark_mode else "Light"
        self._set_status(f"Switched to {mode} mode.")

    # ------------------------------------------------------------------
    # UI Construction
    # ------------------------------------------------------------------

    def _build_menu(self) -> None:
        """Build the top-level menu bar."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open File…", command=self._open_file, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Export Filtered Data…", command=self._export_csv)
        file_menu.add_command(label="Save Chart…", command=self._save_chart)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Toggle Dark Mode", command=self._toggle_dark_mode)
        view_menu.add_command(label="Summary Statistics", command=self._show_summary)
        menubar.add_cascade(label="View", menu=view_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        # Keyboard shortcut
        self.root.bind("<Control-o>", lambda e: self._open_file())

    def _build_toolbar_frame(self) -> None:
        """Build the top toolbar with action buttons."""
        toolbar = ttk.Frame(self.root, style="TFrame")
        toolbar.pack(fill=tk.X, padx=10, pady=(8, 0))

        # Title
        ttk.Label(toolbar, text="📊  GraphIt", style="Heading.TLabel").pack(side=tk.LEFT)

        # Right-side buttons
        btn_frame = ttk.Frame(toolbar, style="TFrame")
        btn_frame.pack(side=tk.RIGHT)

        ttk.Button(btn_frame, text="🌙 Dark Mode", command=self._toggle_dark_mode,
                   style="TButton").pack(side=tk.RIGHT, padx=3)
        ttk.Button(btn_frame, text="🔄 Reset", command=self._reset, style="Danger.TButton").pack(
            side=tk.RIGHT, padx=3)
        ttk.Button(btn_frame, text="📂 Open File", command=self._open_file,
                   style="Accent.TButton").pack(side=tk.RIGHT, padx=3)

    def _build_main_area(self) -> None:
        """Build the notebook (tabs) area."""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

        # Tab 1 — Data Preview
        self._build_data_tab()
        # Tab 2 — Visualization
        self._build_viz_tab()

    # -- Data Tab --

    def _build_data_tab(self) -> None:
        """Build the 'Data Preview' tab."""
        tab = ttk.Frame(self.notebook, style="TFrame")
        self.notebook.add(tab, text="  📋 Data Preview  ")

        # Info panel (top)
        info_frame = ttk.Frame(tab, style="Card.TFrame")
        info_frame.pack(fill=tk.X, padx=6, pady=(6, 2))

        self.info_label = ttk.Label(info_frame, text="No dataset loaded.", style="Info.TLabel",
                                    wraplength=950, justify=tk.LEFT)
        self.info_label.pack(anchor=tk.W, padx=10, pady=6)

        # Search bar
        search_frame = ttk.Frame(tab, style="TFrame")
        search_frame.pack(fill=tk.X, padx=6, pady=2)
        ttk.Label(search_frame, text="🔍 Search columns:", style="TLabel").pack(side=tk.LEFT, padx=(4, 4))
        self.col_search_var = tk.StringVar()
        self.col_search_var.trace_add("write", self._on_column_search)
        self.col_search_entry = ttk.Entry(search_frame, textvariable=self.col_search_var, width=30)
        self.col_search_entry.pack(side=tk.LEFT, padx=4)

        self.col_search_result = ttk.Label(search_frame, text="", style="TLabel")
        self.col_search_result.pack(side=tk.LEFT, padx=8)

        # Summary stats button
        ttk.Button(search_frame, text="📈 Summary Stats", command=self._show_summary,
                   style="TButton").pack(side=tk.RIGHT, padx=4)

        # Treeview table
        table_frame = ttk.Frame(tab, style="TFrame")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        self.tree_scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        self.tree_scroll_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)

        self.tree = ttk.Treeview(table_frame,
                                 yscrollcommand=self.tree_scroll_y.set,
                                 xscrollcommand=self.tree_scroll_x.set,
                                 show="headings")
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

    # -- Visualization Tab --

    def _build_viz_tab(self) -> None:
        """Build the 'Visualization' tab with controls and chart canvas."""
        tab = ttk.Frame(self.notebook, style="TFrame")
        self.notebook.add(tab, text="  📈 Visualization  ")

        # Controls panel (left)
        controls = ttk.Frame(tab, style="Card.TFrame", width=260)
        controls.pack(side=tk.LEFT, fill=tk.Y, padx=(6, 2), pady=6)
        controls.pack_propagate(False)

        self._add_control_section(controls, "Chart Type")
        self.chart_type_var = tk.StringVar(value=self.CHART_TYPES[0])
        ttk.Combobox(controls, textvariable=self.chart_type_var, values=self.CHART_TYPES,
                     state="readonly", width=28).pack(padx=10, pady=(0, 8))

        self._add_control_section(controls, "X-Axis Column")
        self.x_col_var = tk.StringVar()
        self.x_combo = ttk.Combobox(controls, textvariable=self.x_col_var, state="readonly", width=28)
        self.x_combo.pack(padx=10, pady=(0, 8))

        self._add_control_section(controls, "Y-Axis Column")
        self.y_col_var = tk.StringVar()
        self.y_combo = ttk.Combobox(controls, textvariable=self.y_col_var, state="readonly", width=28)
        self.y_combo.pack(padx=10, pady=(0, 8))

        ttk.Separator(controls, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=6)

        # Filter section
        self._add_control_section(controls, "Filter Column")
        self.filter_col_var = tk.StringVar()
        self.filter_col_combo = ttk.Combobox(controls, textvariable=self.filter_col_var,
                                             state="readonly", width=28)
        self.filter_col_combo.pack(padx=10, pady=(0, 4))

        self._add_control_section(controls, "Filter Value")
        self.filter_val_var = tk.StringVar()
        ttk.Entry(controls, textvariable=self.filter_val_var, width=30).pack(padx=10, pady=(0, 4))

        btn_row = ttk.Frame(controls, style="Card.TFrame")
        btn_row.pack(fill=tk.X, padx=10, pady=4)
        ttk.Button(btn_row, text="Apply Filter", command=self._apply_filter,
                   style="Accent.TButton").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        ttk.Button(btn_row, text="Clear", command=self._clear_filter,
                   style="TButton").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))

        self.filter_info_label = ttk.Label(controls, text="", style="Info.TLabel")
        self.filter_info_label.pack(padx=10, pady=(0, 6))

        ttk.Separator(controls, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=6)

        # Action buttons
        ttk.Button(controls, text="📊  Plot Chart", command=self._plot_chart,
                   style="Accent.TButton").pack(fill=tk.X, padx=10, pady=4)
        ttk.Button(controls, text="💾  Save Chart", command=self._save_chart,
                   style="TButton").pack(fill=tk.X, padx=10, pady=2)
        ttk.Button(controls, text="📤  Export CSV", command=self._export_csv,
                   style="TButton").pack(fill=tk.X, padx=10, pady=2)

        # Chart area (right)
        chart_frame = ttk.Frame(tab, style="TFrame")
        chart_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(2, 6), pady=6)

        self.figure = Figure(figsize=(6, 4), dpi=100,
                             facecolor=self.theme["chart_bg"])
        self.canvas = FigureCanvasTkAgg(self.figure, master=chart_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Navigation toolbar
        toolbar_frame = ttk.Frame(chart_frame, style="TFrame")
        toolbar_frame.pack(fill=tk.X)
        self.toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame)
        self.toolbar.update()

    @staticmethod
    def _add_control_section(parent: ttk.Frame, title: str) -> None:
        """Add a small section label inside a controls panel."""
        ttk.Label(parent, text=title, style="SubHeading.TLabel").pack(
            anchor=tk.W, padx=10, pady=(8, 2))

    # -- Status bar --

    def _build_status_bar(self) -> None:
        """Build the bottom status bar."""
        self.status_bar = ttk.Frame(self.root, style="StatusBar.TFrame")
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_label = ttk.Label(self.status_bar, text="Ready", style="Status.TLabel",
                                      anchor=tk.W)
        self.status_label.pack(fill=tk.X, padx=10, pady=3)

    def _set_status(self, message: str) -> None:
        """Update the status bar text."""
        self.status_label.config(text=message)

    # ------------------------------------------------------------------
    # File operations
    # ------------------------------------------------------------------

    def _open_file(self) -> None:
        """Open a file dialog and load a CSV or Excel file."""
        filepath = filedialog.askopenfilename(
            title="Open Data File",
            filetypes=[
                ("All Supported", "*.csv *.xlsx"),
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx"),
                ("All Files", "*.*"),
            ],
        )
        if not filepath:
            return

        ext = os.path.splitext(filepath)[1].lower()
        if ext not in (".csv", ".xlsx"):
            messagebox.showerror("Unsupported File",
                                 f"File type '{ext}' is not supported.\nPlease upload a .csv or .xlsx file.")
            return

        try:
            if ext == ".csv":
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath, engine="openpyxl")
        except Exception as exc:
            messagebox.showerror("File Error", f"Failed to read file:\n{exc}")
            return

        if df.empty:
            messagebox.showwarning("Empty File", "The selected file contains no data.")
            return

        self.df = df
        self.filtered_df = df.copy()
        self.current_file = filepath

        self._populate_table()
        self._update_info()
        self._populate_column_combos()
        self._set_status(f"Loaded: {os.path.basename(filepath)}  |  {len(df)} rows × {len(df.columns)} cols")

    # ------------------------------------------------------------------
    # Data preview helpers
    # ------------------------------------------------------------------

    def _populate_table(self) -> None:
        """Display the first 10 rows of the dataframe in the Treeview."""
        # Clear existing
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []

        if self.df is None:
            return

        cols = list(self.df.columns)
        self.tree["columns"] = cols

        for col in cols:
            self.tree.heading(col, text=col)
            # Auto-width: use header length or sample data length
            max_w = max(len(str(col)), 8)
            for val in self.df[col].head(10):
                max_w = max(max_w, len(str(val)))
            self.tree.column(col, width=min(max_w * 9, 200), anchor=tk.W)

        for _, row in self.df.head(10).iterrows():
            self.tree.insert("", tk.END, values=list(row))

    def _update_info(self) -> None:
        """Update the info label with dataset metadata."""
        if self.df is None:
            self.info_label.config(text="No dataset loaded.")
            return

        rows, cols = self.df.shape
        missing = self.df.isnull().sum().sum()
        col_names = ", ".join(self.df.columns[:20])
        if len(self.df.columns) > 20:
            col_names += f"  … (+{len(self.df.columns) - 20} more)"

        info = (f"Rows: {rows}   |   Columns: {cols}   |   Missing values: {missing}\n"
                f"Columns: {col_names}")
        self.info_label.config(text=info)

    def _populate_column_combos(self) -> None:
        """Fill every Combobox with column names from the current dataframe."""
        if self.df is None:
            return

        cols = list(self.df.columns)
        self.x_combo["values"] = cols
        self.y_combo["values"] = cols
        self.filter_col_combo["values"] = cols

        if cols:
            self.x_col_var.set(cols[0])
            self.y_col_var.set(cols[1] if len(cols) > 1 else cols[0])

    # ------------------------------------------------------------------
    # Column search
    # ------------------------------------------------------------------

    def _on_column_search(self, *_args) -> None:
        """React to changes in the column search entry."""
        if self.df is None:
            self.col_search_result.config(text="")
            return

        query = self.col_search_var.get().strip().lower()
        if not query:
            self.col_search_result.config(text="")
            return

        matches = [c for c in self.df.columns if query in c.lower()]
        if matches:
            self.col_search_result.config(text=f"Matches: {', '.join(matches[:8])}")
        else:
            self.col_search_result.config(text="No matching columns.")

    # ------------------------------------------------------------------
    # Data filtering
    # ------------------------------------------------------------------

    def _apply_filter(self) -> None:
        """Filter the dataframe by the chosen column and value."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load a dataset first.")
            return

        col = self.filter_col_var.get()
        val = self.filter_val_var.get().strip()

        # If no filter column or value is set, silently reset to full dataset
        if not col or not val:
            self._clear_filter()
            return

        try:
            # Attempt numeric comparison first
            if pd.api.types.is_numeric_dtype(self.df[col]):
                numeric_val = float(val)
                mask = self.df[col] == numeric_val
            else:
                # Case-insensitive string matching (partial)
                mask = self.df[col].astype(str).str.contains(val, case=False, na=False)

            self.filtered_df = self.df[mask].copy()
            n = len(self.filtered_df)
            self.filter_info_label.config(text=f"Showing {n} of {len(self.df)} rows")
            self._set_status(f"Filter applied on '{col}' → {n} rows matched.")
        except Exception as exc:
            messagebox.showerror("Filter Error", f"Could not apply filter:\n{exc}")

    def _clear_filter(self) -> None:
        """Remove the current filter and restore the full dataset."""
        if self.df is not None:
            self.filtered_df = self.df.copy()
        self.filter_val_var.set("")
        self.filter_col_var.set("")
        self.filter_info_label.config(text="")
        self._set_status("Filter cleared.")

    # ------------------------------------------------------------------
    # Plotting
    # ------------------------------------------------------------------

    def _plot_chart(self) -> None:
        """Draw the selected chart type using the current (filtered) data."""
        if self.filtered_df is None or self.filtered_df.empty:
            messagebox.showwarning("No Data", "Please load a dataset first.")
            return

        chart_type = self.chart_type_var.get()
        x_col = self.x_col_var.get()
        y_col = self.y_col_var.get()
        df = self.filtered_df

        # Validate column selection
        if chart_type in ("Line Chart", "Bar Chart", "Scatter Plot"):
            if not x_col or not y_col:
                messagebox.showwarning("Missing Columns", "Please select both X and Y columns.")
                return
            if x_col not in df.columns or y_col not in df.columns:
                messagebox.showerror("Invalid Column", "Selected column not found in the dataset.")
                return

        if chart_type in ("Histogram", "Box Plot"):
            if not x_col:
                messagebox.showwarning("Missing Column", "Please select at least the X-axis column.")
                return

        if chart_type == "Pie Chart":
            if not x_col:
                messagebox.showwarning("Missing Column", "Please select the X-axis column for Pie labels.")
                return

        # Clear figure
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_facecolor(self.theme["chart_face"])
        self.figure.patch.set_facecolor(self.theme["chart_bg"])

        text_color = self.theme["fg"]
        ax.tick_params(colors=text_color)
        ax.xaxis.label.set_color(text_color)
        ax.yaxis.label.set_color(text_color)
        ax.title.set_color(text_color)
        for spine in ax.spines.values():
            spine.set_color(self.theme["separator"])

        try:
            artist = None  # For mplcursors

            if chart_type == "Line Chart":
                self._assert_numeric(df, y_col)
                artist, = ax.plot(df[x_col], df[y_col], color=CHART_COLORS[0],
                                  linewidth=2, marker="o", markersize=4)
                ax.set_xlabel(x_col)
                ax.set_ylabel(y_col)
                ax.set_title(f"Line Chart — {y_col} vs {x_col}")

            elif chart_type == "Bar Chart":
                self._assert_numeric(df, y_col)
                data = df.head(30)  # limit bars for readability
                bars = ax.bar(data[x_col].astype(str), data[y_col], color=CHART_COLORS[:len(data)],
                              edgecolor="none")
                ax.set_xlabel(x_col)
                ax.set_ylabel(y_col)
                ax.set_title(f"Bar Chart — {y_col} by {x_col}")
                plt.setp(ax.get_xticklabels(), rotation=45, ha="right", fontsize=8)
                artist = bars

            elif chart_type == "Scatter Plot":
                self._assert_numeric(df, x_col)
                self._assert_numeric(df, y_col)
                artist = ax.scatter(df[x_col], df[y_col], c=CHART_COLORS[0], alpha=0.7,
                                    edgecolors=CHART_COLORS[1], linewidths=0.5, s=40)
                ax.set_xlabel(x_col)
                ax.set_ylabel(y_col)
                ax.set_title(f"Scatter Plot — {y_col} vs {x_col}")

            elif chart_type == "Histogram":
                self._assert_numeric(df, x_col)
                _, _, patches = ax.hist(df[x_col].dropna(), bins="auto", color=CHART_COLORS[0],
                                        edgecolor=self.theme["chart_bg"], alpha=0.85)
                ax.set_xlabel(x_col)
                ax.set_ylabel("Frequency")
                ax.set_title(f"Histogram — {x_col}")

            elif chart_type == "Pie Chart":
                counts = df[x_col].value_counts().head(10)
                wedges, texts, autotexts = ax.pie(
                    counts, labels=counts.index.astype(str), autopct="%1.1f%%",
                    colors=CHART_COLORS[:len(counts)], startangle=140,
                    textprops={"fontsize": 8, "color": text_color})
                ax.set_title(f"Pie Chart — {x_col}")

            elif chart_type == "Box Plot":
                self._assert_numeric(df, x_col)
                bp = ax.boxplot(df[x_col].dropna(), patch_artist=True,
                                boxprops=dict(facecolor=CHART_COLORS[0], color=CHART_COLORS[1]),
                                medianprops=dict(color=CHART_COLORS[3]),
                                whiskerprops=dict(color=text_color),
                                capprops=dict(color=text_color),
                                flierprops=dict(markerfacecolor=CHART_COLORS[5], marker="o",
                                                markersize=5, alpha=0.6))
                ax.set_ylabel(x_col)
                ax.set_title(f"Box Plot — {x_col}")

            # Add hover tooltips if mplcursors available
            if HAS_MPLCURSORS and artist is not None:
                try:
                    mplcursors.cursor(artist, hover=True)
                except Exception:
                    pass  # Some artist types may not support cursor

            self.figure.tight_layout()
            self.canvas.draw()
            self._set_status(f"Chart plotted: {chart_type}")

        except ValueError as ve:
            messagebox.showerror("Plot Error", str(ve))
        except Exception as exc:
            messagebox.showerror("Plot Error", f"An error occurred while plotting:\n{exc}")

    @staticmethod
    def _assert_numeric(df: pd.DataFrame, col: str) -> None:
        """Raise ValueError if the column is not numeric."""
        if not pd.api.types.is_numeric_dtype(df[col]):
            raise ValueError(
                f"Column '{col}' is not numeric.\n"
                "Please select a numeric column for this chart type."
            )

    # ------------------------------------------------------------------
    # Save / Export
    # ------------------------------------------------------------------

    def _save_chart(self) -> None:
        """Save the current Matplotlib chart to a PNG or JPG file."""
        if self.figure is None or not self.figure.axes:
            messagebox.showwarning("No Chart", "There is no chart to save. Plot one first.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Save Chart",
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("All Files", "*.*")],
        )
        if filepath:
            try:
                self.figure.savefig(filepath, dpi=150, bbox_inches="tight",
                                    facecolor=self.figure.get_facecolor())
                self._set_status(f"Chart saved to {os.path.basename(filepath)}")
                messagebox.showinfo("Saved", f"Chart saved successfully:\n{filepath}")
            except Exception as exc:
                messagebox.showerror("Save Error", f"Could not save chart:\n{exc}")

    def _export_csv(self) -> None:
        """Export the current (filtered) dataframe to a CSV file."""
        if self.filtered_df is None:
            messagebox.showwarning("No Data", "No dataset loaded to export.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Export Data as CSV",
            defaultextension=".csv",
            filetypes=[("CSV File", "*.csv"), ("All Files", "*.*")],
        )
        if filepath:
            try:
                self.filtered_df.to_csv(filepath, index=False)
                self._set_status(f"Exported {len(self.filtered_df)} rows to {os.path.basename(filepath)}")
                messagebox.showinfo("Exported", f"Data exported successfully:\n{filepath}")
            except Exception as exc:
                messagebox.showerror("Export Error", f"Could not export data:\n{exc}")

    # ------------------------------------------------------------------
    # Summary statistics
    # ------------------------------------------------------------------

    def _show_summary(self) -> None:
        """Show summary statistics for the loaded dataset in a popup."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load a dataset first.")
            return

        win = tk.Toplevel(self.root)
        win.title("Summary Statistics")
        win.geometry("700x500")
        win.configure(bg=self.theme["bg"])

        ttk.Label(win, text="📈  Summary Statistics", style="Heading.TLabel").pack(
            anchor=tk.W, padx=12, pady=(10, 4))

        text = tk.Text(win, wrap=tk.WORD, font=("Consolas", 10),
                       bg=self.theme["frame_bg"], fg=self.theme["fg"],
                       insertbackground=self.theme["fg"], relief=tk.FLAT, padx=10, pady=10)
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        desc = self.df.describe(include="all").to_string()
        dtypes = self.df.dtypes.to_string()
        missing = self.df.isnull().sum().to_string()

        text.insert(tk.END, "=== Data Types ===\n")
        text.insert(tk.END, dtypes + "\n\n")
        text.insert(tk.END, "=== Missing Values ===\n")
        text.insert(tk.END, missing + "\n\n")
        text.insert(tk.END, "=== Descriptive Statistics ===\n")
        text.insert(tk.END, desc + "\n")
        text.config(state=tk.DISABLED)

    # ------------------------------------------------------------------
    # Reset
    # ------------------------------------------------------------------

    def _reset(self) -> None:
        """Reset the application to its initial state."""
        self.df = None
        self.filtered_df = None
        self.current_file = None

        # Clear table
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []
        self.info_label.config(text="No dataset loaded.")

        # Clear combos
        for combo in (self.x_combo, self.y_combo, self.filter_col_combo):
            combo["values"] = []
            combo.set("")
        self.filter_val_var.set("")
        self.filter_info_label.config(text="")
        self.col_search_var.set("")

        # Clear chart
        if self.figure:
            self.figure.clear()
            self.canvas.draw()

        self._set_status("Application reset. Ready.")

    # ------------------------------------------------------------------
    # About dialog
    # ------------------------------------------------------------------

    @staticmethod
    def _show_about() -> None:
        """Show a simple About dialog."""
        messagebox.showinfo(
            "About GraphIt",
            "GraphIt v1.0\n\n"
            "A modern desktop data visualization tool.\n\n"
            "Built with:\n"
            "  • Python & Tkinter\n"
            "  • Pandas\n"
            "  • Matplotlib\n"
            "  • mplcursors\n\n"
            "© 2026"
        )


# -----------------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------------

def main() -> None:
    """Launch the GraphIt application."""
    root = tk.Tk()

    # Try to set a window icon (silently skip if unavailable)
    try:
        root.iconbitmap(default="")
    except Exception:
        pass

    _app = DataVisualizerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
