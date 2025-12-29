import os
import sys
import threading
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


# =========================
# Helpers
# =========================
def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def normalize_ani(series: pd.Series) -> pd.Series:
    """
    Normaliza ANI:
    - convierte a string
    - recorta espacios
    - quita decimales tipo '123.0'
    - deja solo dígitos (si viene con espacios/guiones)
    """
    s = series.astype(str).str.strip()
    # Quitar .0 típico de Excel
    s = s.str.replace(r"\.0$", "", regex=True)
    # Dejar solo dígitos
    s = s.str.replace(r"\D+", "", regex=True)
    # Vacíos -> NaN para filtrar
    s = s.replace("", pd.NA)
    return s


def list_excel_files(folder: str):
    exts = (".xlsx", ".xls", ".xlsm", ".xlsb", ".ods")
    files = []
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if os.path.isfile(path) and name.lower().endswith(exts):
            files.append(path)
    return sorted(files)


def read_portin_anis_from_file(path: str) -> set[str]:
    """
    Lee un archivo PORTIN:
    - intenta encontrar columna 'ani' (case-insensitive)
    - si no, toma primera columna
    """
    # engine = openpyxl para xlsx/xlsm; pandas elige solo; para xls puede requerir xlrd (depende)
    df = pd.read_excel(path, dtype=str)

    if df.empty:
        return set()

    cols_lower = {c.lower(): c for c in df.columns}
    if "ani" in cols_lower:
        col = cols_lower["ani"]
    else:
        col = df.columns[0]

    anis = normalize_ani(df[col]).dropna().unique().tolist()
    return set(anis)


def read_portout(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str)
    if df.empty:
        return df

    # Buscar columna ani
    cols_lower = {c.lower(): c for c in df.columns}
    if "ani" not in cols_lower:
        raise ValueError("El archivo PortOut no tiene una columna 'ani'.")

    ani_col = cols_lower["ani"]
    df[ani_col] = normalize_ani(df[ani_col])
    return df


# =========================
# UI (OSAR style)
# =========================
OSAR_BG = "#0B1220"        # azul oscuro
OSAR_CARD = "#111A2E"      # tarjeta
OSAR_ACCENT = "#18C29C"    # verde OSAR
OSAR_ACCENT_2 = "#2D6CDF"  # azul OSAR
OSAR_TEXT = "#E8EEF9"
OSAR_MUTED = "#9FB0D0"
OSAR_DANGER = "#FF5A5F"


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("OSAR · Filtrar PortOut vs PortIn")
        self.geometry("920x560")
        self.minsize(900, 520)
        self.configure(bg=OSAR_BG)

        self.portin_folder = tk.StringVar()
        self.portout_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.status = tk.StringVar(value="Listo.")
        self.is_running = False

        self._build_styles()
        self._build_layout()

    def _build_styles(self):
        style = ttk.Style(self)
        # Usar tema base
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("TFrame", background=OSAR_BG)
        style.configure("Card.TFrame", background=OSAR_CARD)

        style.configure("TLabel", background=OSAR_BG, foreground=OSAR_TEXT, font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=OSAR_BG, foreground=OSAR_TEXT, font=("Segoe UI", 18, "bold"))
        style.configure("Sub.TLabel", background=OSAR_BG, foreground=OSAR_MUTED, font=("Segoe UI", 10))

        style.configure("CardTitle.TLabel", background=OSAR_CARD, foreground=OSAR_TEXT, font=("Segoe UI", 12, "bold"))
        style.configure("CardText.TLabel", background=OSAR_CARD, foreground=OSAR_MUTED, font=("Segoe UI", 10))

        style.configure("TEntry", fieldbackground="#0E1730", foreground=OSAR_TEXT, bordercolor="#223055")
        style.map("TEntry", fieldbackground=[("readonly", "#0E1730")])

        style.configure("Accent.TButton",
                        background=OSAR_ACCENT_2,
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        padding=(14, 10),
                        borderwidth=0)
        style.map("Accent.TButton", background=[("active", "#3A7AF0")])

        style.configure("Green.TButton",
                        background=OSAR_ACCENT,
                        foreground="#06111A",
                        font=("Segoe UI", 10, "bold"),
                        padding=(14, 10),
                        borderwidth=0)
        style.map("Green.TButton", background=[("active", "#2BE0B7")])

        style.configure("Danger.TButton",
                        background=OSAR_DANGER,
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        padding=(14, 10),
                        borderwidth=0)
        style.map("Danger.TButton", background=[("active", "#FF777B")])

        style.configure("TProgressbar", troughcolor="#0E1730", background=OSAR_ACCENT, bordercolor="#223055")

    def _build_layout(self):
        header = ttk.Frame(self, style="TFrame")
        header.pack(fill="x", padx=18, pady=(16, 10))

        ttk.Label(header, text="OSAR · Filtrar PortOut vs PortIn", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Carga una carpeta de PORTIN (ANIs) + un Excel PORTOUT, y exporta solo los ANIs que NO están en PortIn.",
            style="Sub.TLabel"
        ).pack(anchor="w", pady=(6, 0))

        body = ttk.Frame(self, style="TFrame")
        body.pack(fill="both", expand=True, padx=18, pady=10)

        # Left card (inputs)
        left = ttk.Frame(body, style="Card.TFrame")
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))

        ttk.Label(left, text="Entradas", style="CardTitle.TLabel").pack(anchor="w", padx=14, pady=(14, 6))
        ttk.Label(left, text="Seleccioná carpeta PORTIN y el archivo PORTOUT.", style="CardText.TLabel").pack(
            anchor="w", padx=14, pady=(0, 10)
        )

        self._row_picker(left, "Carpeta PortIn (muchos Excel):", self.portin_folder, self.pick_portin_folder)
        self._row_picker(left, "Archivo PortOut (Excel):", self.portout_file, self.pick_portout_file)
        self._row_picker(left, "Carpeta destino (salida):", self.output_folder, self.pick_output_folder)

        # Options
        opts = ttk.Frame(left, style="Card.TFrame")
        opts.pack(fill="x", padx=14, pady=(10, 12))

        self.chk_keep_empty_ani = tk.BooleanVar(value=False)
        cb = tk.Checkbutton(
            opts,
            text="Mantener filas con ANI vacío (NO recomendado)",
            variable=self.chk_keep_empty_ani,
            bg=OSAR_CARD,
            fg=OSAR_TEXT,
            activebackground=OSAR_CARD,
            activeforeground=OSAR_TEXT,
            selectcolor=OSAR_CARD,
            highlightthickness=0,
            font=("Segoe UI", 10),
        )
        cb.pack(anchor="w", pady=8)

        # Buttons
        btns = ttk.Frame(left, style="Card.TFrame")
        btns.pack(fill="x", padx=14, pady=(0, 14))

        self.btn_run = ttk.Button(btns, text="Procesar y Exportar", style="Green.TButton", command=self.on_run)
        self.btn_run.pack(side="left", padx=(0, 10), pady=10)

        self.btn_clear = ttk.Button(btns, text="Limpiar", style="Accent.TButton", command=self.on_clear)
        self.btn_clear.pack(side="left", pady=10)

        self.btn_stop = ttk.Button(btns, text="Cancelar", style="Danger.TButton", command=self.on_cancel, state="disabled")
        self.btn_stop.pack(side="right", pady=10)

        # Right card (log)
        right = ttk.Frame(body, style="Card.TFrame")
        right.pack(side="right", fill="both", expand=True)

        ttk.Label(right, text="Progreso / Log", style="CardTitle.TLabel").pack(anchor="w", padx=14, pady=(14, 6))
        self.progress = ttk.Progressbar(right, mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=14, pady=(0, 10))

        self.txt = tk.Text(
            right,
            height=18,
            bg="#0E1730",
            fg=OSAR_TEXT,
            insertbackground=OSAR_TEXT,
            bd=0,
            highlightthickness=1,
            highlightbackground="#223055",
            font=("Consolas", 10),
        )
        self.txt.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        footer = ttk.Frame(self, style="TFrame")
        footer.pack(fill="x", padx=18, pady=(0, 14))
        ttk.Label(footer, textvariable=self.status, style="Sub.TLabel").pack(anchor="w")

        self._log("Listo. Elegí carpeta PORTIN y archivo PORTOUT.")

        # cancel flag
        self.cancel_requested = False

    def _row_picker(self, parent, label, var, cmd):
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x", padx=14, pady=6)

        ttk.Label(row, text=label, style="CardText.TLabel").pack(anchor="w")
        inner = ttk.Frame(row, style="Card.TFrame")
        inner.pack(fill="x", pady=(6, 0))

        ent = ttk.Entry(inner, textvariable=var)
        ent.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ent.configure(state="readonly")

        ttk.Button(inner, text="Elegir...", style="Accent.TButton", command=cmd).pack(side="right")

    def _log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.txt.insert("end", f"[{ts}] {msg}\n")
        self.txt.see("end")
        self.update_idletasks()

    def pick_portin_folder(self):
        folder = filedialog.askdirectory(title="Seleccionar carpeta PORTIN")
        if folder:
            self.portin_folder.set(folder)
            self._log(f"Carpeta PORTIN: {folder}")

    def pick_portout_file(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo PORTOUT",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm *.xlsb *.ods"), ("Todos", "*.*")]
        )
        if path:
            self.portout_file.set(path)
            self._log(f"Archivo PORTOUT: {path}")

    def pick_output_folder(self):
        folder = filedialog.askdirectory(title="Seleccionar carpeta destino")
        if folder:
            self.output_folder.set(folder)
            self._log(f"Destino: {folder}")

    def on_clear(self):
        if self.is_running:
            return
        self.portin_folder.set("")
        self.portout_file.set("")
        self.output_folder.set("")
        self.progress["value"] = 0
        self.txt.delete("1.0", "end")
        self.status.set("Listo.")
        self._log("Listo. Elegí carpeta PORTIN y archivo PORTOUT.")

    def on_cancel(self):
        if self.is_running:
            self.cancel_requested = True
            self._log("Cancelación solicitada...")

    def _set_running(self, running: bool):
        self.is_running = running
        self.btn_run.configure(state="disabled" if running else "normal")
        self.btn_clear.configure(state="disabled" if running else "normal")
        self.btn_stop.configure(state="normal" if running else "disabled")

    def on_run(self):
        if self.is_running:
            return

        portin_folder = self.portin_folder.get().strip()
        portout_file = self.portout_file.get().strip()
        out_folder = self.output_folder.get().strip()

        if not portin_folder or not os.path.isdir(portin_folder):
            messagebox.showerror("Falta dato", "Seleccioná una carpeta válida de PORTIN.")
            return
        if not portout_file or not os.path.isfile(portout_file):
            messagebox.showerror("Falta dato", "Seleccioná un archivo válido de PORTOUT.")
            return
        if not out_folder or not os.path.isdir(out_folder):
            messagebox.showerror("Falta dato", "Seleccioná una carpeta válida de destino.")
            return

        self.cancel_requested = False
        self.progress["value"] = 0
        self._set_running(True)
        self.status.set("Procesando...")

        th = threading.Thread(target=self._process, daemon=True)
        th.start()

    def _process(self):
        try:
            portin_folder = self.portin_folder.get().strip()
            portout_file = self.portout_file.get().strip()
            out_folder = self.output_folder.get().strip()

            self._log("Buscando archivos Excel en PORTIN...")
            portin_files = list_excel_files(portin_folder)
            if not portin_files:
                raise ValueError("No se encontraron archivos Excel en la carpeta PORTIN.")

            self._log(f"Encontrados {len(portin_files)} archivos PORTIN.")
            self.progress["value"] = 5

            # Leer PORTIN anis
            anis_portin = set()
            total_files = len(portin_files)
            for i, f in enumerate(portin_files, start=1):
                if self.cancel_requested:
                    raise RuntimeError("Proceso cancelado por el usuario.")

                try:
                    s = read_portin_anis_from_file(f)
                    anis_portin |= s
                    self._log(f"PORTIN {i}/{total_files}: +{len(s)} ANIs · total={len(anis_portin)}")
                except Exception as e:
                    self._log(f"⚠️ Error leyendo PORTIN '{os.path.basename(f)}': {e}")

                # Progreso 5% a 55%
                self.progress["value"] = 5 + (50 * i / max(1, total_files))
                self.update_idletasks()

            if not anis_portin:
                self._log("⚠️ No se detectaron ANIs en PORTIN. Se exportará PORTOUT completo (según configuración).")

            # Leer PORTOUT
            self._log("Leyendo PORTOUT...")
            df_out = read_portout(portout_file)
            self.progress["value"] = 65

            if df_out.empty:
                raise ValueError("PORTOUT está vacío.")

            cols_lower = {c.lower(): c for c in df_out.columns}
            ani_col = cols_lower["ani"]

            # Filtrar
            self._log("Filtrando filas de PORTOUT (ANIs que NO están en PORTIN)...")
            before = len(df_out)

            if self.chk_keep_empty_ani.get():
                mask_keep = df_out[ani_col].isna() | (~df_out[ani_col].isin(anis_portin))
            else:
                # descartar ANI vacío
                mask_keep = df_out[ani_col].notna() & (~df_out[ani_col].isin(anis_portin))

            df_res = df_out.loc[mask_keep].copy()
            after = len(df_res)

            self.progress["value"] = 85
            self._log(f"PORTOUT filas: {before} · Resultado: {after} · Quitadas: {before - after}")

            # Guardar
            base_name = os.path.splitext(os.path.basename(portout_file))[0]
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = os.path.join(out_folder, f"{base_name}__NO_EN_PORTIN__{ts}.xlsx")

            self._log(f"Exportando: {out_path}")
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                df_res.to_excel(writer, index=False, sheet_name="Resultado")

            self.progress["value"] = 100
            self.status.set("Finalizado.")
            self._log("✅ Listo. Exportación completada.")
            self._log(f"Archivo generado: {out_path}")

            messagebox.showinfo("OK", f"Proceso finalizado.\n\nResultado:\n{out_path}")

        except Exception as e:
            self.status.set("Error.")
            self._log(f"❌ Error: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self._set_running(False)


if __name__ == "__main__":
    app = App()
    app.mainloop()
