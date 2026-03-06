"""
app.py - GUI fuer den Synthetischen Datengenerator
Tkinter-basiert, kein Browser noetig, laeuft komplett lokal.
"""

import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from audit import generate_audit_report
from synthesizer import TYPE_LABELS, LABEL_TO_TYPE, ALL_TYPES, analyze_excel, synthesize_excel

# ---------------------------------------------------------------------------
# Farben & Fonts
# ---------------------------------------------------------------------------
COLOR_BG        = "#f5f6fa"
COLOR_HEADER    = "#1a3a5c"
COLOR_BTN_BLUE  = "#1a3a5c"
COLOR_BTN_GREEN = "#27ae60"
COLOR_ANON_BG   = "#eaf4ea"
COLOR_ANON_FG   = "#1e6b1e"
COLOR_SKIP_FG   = "#999999"
COLOR_ERROR     = "#c0392b"
COLOR_MUTED     = "#7f8c8d"
COLOR_ROW_ALT   = "#f0f4fa"

FONT_TITLE  = ("Segoe UI", 15, "bold")
FONT_LABEL  = ("Segoe UI", 9, "bold")
FONT_SMALL  = ("Segoe UI", 8)
FONT_BTN    = ("Segoe UI", 10, "bold")
FONT_HEADER = ("Segoe UI", 8, "bold")
FONT_MONO   = ("Consolas", 8)

DROPDOWN_VALUES = [TYPE_LABELS[t] for t in ALL_TYPES]


# ---------------------------------------------------------------------------
# Vorschau-Dialog
# ---------------------------------------------------------------------------

class PreviewDialog(tk.Toplevel):
    ROW_H = 34

    def __init__(self, parent, input_path: str, output_path: str, on_generate):
        super().__init__(parent)
        self.title("Spalten-Vorschau & Konfiguration")
        self.resizable(True, True)
        self.configure(bg=COLOR_BG)
        self.grab_set()

        self._on_generate  = on_generate
        self._input_path   = input_path
        self._output_path  = output_path
        self._type_vars: dict[str, tk.StringVar] = {}

        self._build_loading_screen()
        self.update()
        threading.Thread(target=self._load_analysis, daemon=True).start()

    def _build_loading_screen(self):
        self.geometry("780x200")
        self._center()
        tk.Label(self, text="Datei wird analysiert...",
                 font=FONT_TITLE, bg=COLOR_BG, fg=COLOR_HEADER).pack(expand=True)

    def _load_analysis(self):
        try:
            analysis = analyze_excel(self._input_path)
            self.after(0, self._build_ui, analysis)
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror("Fehler", str(exc), parent=self))
            self.after(0, self.destroy)

    def _build_ui(self, analysis: dict):
        for w in self.winfo_children():
            w.destroy()

        all_cols: list[tuple[str, str, dict]] = []
        for sheet, cols in analysis.items():
            for col, info in cols.items():
                all_cols.append((sheet, col, info))

        height = min(max(len(all_cols) * self.ROW_H + 160, 300), 650)
        self.geometry(f"860x{height}")
        self._center()

        # Header
        header = tk.Frame(self, bg=COLOR_HEADER, pady=12)
        header.pack(fill=tk.X)
        tk.Label(header, text="Spalten-Vorschau & Konfiguration",
                 font=FONT_TITLE, bg=COLOR_HEADER, fg="white").pack()
        tk.Label(header,
                 text="Bitte pruefen Sie die erkannten Typen und korrigieren Sie diese falls noetig.",
                 font=FONT_SMALL, bg=COLOR_HEADER, fg="#a8c8e8").pack(pady=(2, 0))

        # Scrollbarer Bereich
        container = tk.Frame(self, bg=COLOR_BG)
        container.pack(fill=tk.BOTH, expand=True)

        canvas    = tk.Canvas(container, bg=COLOR_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        inner = tk.Frame(canvas, bg=COLOR_BG)
        win   = canvas.create_window((0, 0), window=inner, anchor="nw")

        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), "units"))

        # Spaltenkoepfe
        hdr = tk.Frame(inner, bg=COLOR_HEADER)
        hdr.pack(fill=tk.X)
        for text, w in [("Spalte", 200), ("Erkannter Typ", 240), ("Beispielwerte (Original)", 380)]:
            tk.Label(hdr, text=text, font=FONT_HEADER, bg=COLOR_HEADER, fg="white",
                     width=w // 7, anchor="w", padx=8, pady=6).pack(side=tk.LEFT)

        # Zeilen
        for i, (sheet, col, info) in enumerate(all_cols):
            detected  = info["type"]
            samples   = info["samples"]
            row_bg    = COLOR_ROW_ALT if i % 2 == 0 else "white"

            row = tk.Frame(inner, bg=row_bg)
            row.pack(fill=tk.X)

            col_display = col if len(analysis) == 1 else f"[{sheet}] {col}"
            tk.Label(row, text=col_display, font=FONT_SMALL,
                     bg=row_bg, anchor="w", padx=10, width=28).pack(side=tk.LEFT, ipady=6)

            var = tk.StringVar(value=TYPE_LABELS.get(detected, TYPE_LABELS["text"]))
            self._type_vars[(sheet, col)] = var

            combo = ttk.Combobox(row, textvariable=var, values=DROPDOWN_VALUES,
                                  state="readonly", width=30, font=FONT_SMALL)
            combo.pack(side=tk.LEFT, padx=(0, 8), pady=4)
            self._color_combo(combo, var)
            var.trace_add("write", lambda *_, c=combo, v=var: self._color_combo(c, v))

            sample_text = "  |  ".join(str(s)[:24] for s in samples[:3]) if samples else "—"
            tk.Label(row, text=sample_text, font=FONT_MONO,
                     bg=row_bg, fg="#555", anchor="w", padx=6).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Footer
        footer = tk.Frame(self, bg=COLOR_BG, pady=12, padx=20)
        footer.pack(fill=tk.X)

        tk.Label(footer, text="Gruen = wird anonymisiert   |   Grau = unveraendert",
                 font=FONT_SMALL, fg=COLOR_MUTED, bg=COLOR_BG).pack(side=tk.LEFT)

        btn_frame = tk.Frame(footer, bg=COLOR_BG)
        btn_frame.pack(side=tk.RIGHT)

        tk.Button(btn_frame, text="Abbrechen", command=self.destroy,
                  bg="#ccc", fg="#333", font=FONT_SMALL, relief=tk.FLAT,
                  padx=14, pady=6, cursor="hand2").pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(btn_frame, text="Synthetische Datei erstellen",
                  command=self._on_confirm,
                  bg=COLOR_BTN_GREEN, fg="white", font=FONT_BTN,
                  relief=tk.FLAT, padx=20, pady=8, cursor="hand2").pack(side=tk.LEFT)

    def _color_combo(self, combo: ttk.Combobox, var: tk.StringVar):
        ct = LABEL_TO_TYPE.get(var.get(), "text")
        combo.configure(foreground=COLOR_ANON_FG if ct != "text" else COLOR_SKIP_FG)

    def _on_confirm(self):
        overrides = {col: LABEL_TO_TYPE.get(var.get(), "text")
                     for (_, col), var in self._type_vars.items()}
        self.destroy()
        self._on_generate(overrides)

    def _center(self):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w,  h  = self.winfo_width(),       self.winfo_height()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")


# ---------------------------------------------------------------------------
# Haupt-App
# ---------------------------------------------------------------------------

class SynthesizerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Synthetische Daten Generator")
        self.root.geometry("620x440")
        self.root.resizable(False, False)
        self.root.configure(bg=COLOR_BG)

        self.input_path  = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Bereit")
        self.progress_var = tk.DoubleVar(value=0.0)
        self._last_audit_path: str | None = None

        self._build_ui()

    def _build_ui(self):
        # Header
        header = tk.Frame(self.root, bg=COLOR_HEADER, pady=18)
        header.pack(fill=tk.X)
        tk.Label(header, text="Synthetische Daten Generator",
                 font=FONT_TITLE, bg=COLOR_HEADER, fg="white").pack()
        tk.Label(header,
                 text="Excel-Datei anonymisieren - 100% lokal, keine Daten verlassen Ihren Computer",
                 font=FONT_SMALL, bg=COLOR_HEADER, fg="#a8c8e8").pack(pady=(2, 0))

        main = tk.Frame(self.root, bg=COLOR_BG, padx=35, pady=22)
        main.pack(fill=tk.BOTH, expand=True)

        # Input
        tk.Label(main, text="Original Excel-Datei:", font=FONT_LABEL, bg=COLOR_BG).pack(anchor=tk.W)
        inp = tk.Frame(main, bg=COLOR_BG)
        inp.pack(fill=tk.X, pady=(3, 12))
        tk.Entry(inp, textvariable=self.input_path, font=("Segoe UI", 9),
                 state="readonly", width=58, relief=tk.SOLID, bd=1
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        tk.Button(inp, text="Durchsuchen", command=self._browse_input,
                  bg=COLOR_BTN_BLUE, fg="white", font=FONT_SMALL,
                  relief=tk.FLAT, padx=12, pady=4, cursor="hand2"
                  ).pack(side=tk.RIGHT, padx=(6, 0))

        # Output
        tk.Label(main, text="Ausgabe-Datei (synthetisch):", font=FONT_LABEL, bg=COLOR_BG).pack(anchor=tk.W)
        out = tk.Frame(main, bg=COLOR_BG)
        out.pack(fill=tk.X, pady=(3, 16))
        tk.Entry(out, textvariable=self.output_path, font=("Segoe UI", 9),
                 state="readonly", width=58, relief=tk.SOLID, bd=1
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        tk.Button(out, text="Speichern als", command=self._browse_output,
                  bg=COLOR_BTN_BLUE, fg="white", font=FONT_SMALL,
                  relief=tk.FLAT, padx=12, pady=4, cursor="hand2"
                  ).pack(side=tk.RIGHT, padx=(6, 0))

        # Info-Box
        info_box = tk.Frame(main, bg=COLOR_ANON_BG, relief=tk.GROOVE, bd=1, padx=14, pady=9)
        info_box.pack(fill=tk.X, pady=(0, 16))
        tk.Label(info_box,
                 text=("Nach der Dateiauswahl oeffnet sich eine Vorschau aller Spalten.\n"
                       "Pruefen und korrigieren Sie die erkannten Typen - dann 'Erstellen' klicken.\n"
                       "Excel-Formatierung bleibt erhalten. Auditbericht wird automatisch erzeugt."),
                 font=FONT_SMALL, bg=COLOR_ANON_BG, fg=COLOR_ANON_FG, justify=tk.LEFT
                 ).pack(anchor=tk.W)

        # Fortschritt
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("green.Horizontal.TProgressbar",
                        troughcolor="#dde", background=COLOR_BTN_GREEN)
        ttk.Progressbar(main, variable=self.progress_var,
                        maximum=1.0, style="green.Horizontal.TProgressbar"
                        ).pack(fill=tk.X, pady=(0, 4))
        tk.Label(main, textvariable=self.status_text,
                 font=FONT_SMALL, fg=COLOR_MUTED, bg=COLOR_BG).pack(anchor=tk.W)

        # Hauptbutton
        self.btn_open = tk.Button(
            main, text="Datei auswaehlen & Vorschau",
            command=self._browse_input,
            bg=COLOR_BTN_BLUE, fg="white", font=FONT_BTN,
            relief=tk.FLAT, padx=24, pady=10, cursor="hand2"
        )
        self.btn_open.pack(pady=(12, 4))

        self.result_label = tk.Label(main, text="", font=FONT_SMALL, bg=COLOR_BG)
        self.result_label.pack()

    # ---- Dateidialoge ----

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Excel-Datei auswaehlen",
            filetypes=[("Excel Dateien", "*.xlsx *.xls"), ("Alle Dateien", "*.*")],
        )
        if not path:
            return
        self.input_path.set(path)
        p = Path(path)
        self.output_path.set(str(p.parent / f"{p.stem}_synthetisch{p.suffix}"))
        self.result_label.config(text="")
        self.progress_var.set(0.0)
        self.status_text.set("Bereit")
        self._last_audit_path = None

        PreviewDialog(self.root, input_path=path,
                      output_path=self.output_path.get(),
                      on_generate=self._generate)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Ausgabedatei speichern",
            defaultextension=".xlsx",
            filetypes=[("Excel Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if path:
            self.output_path.set(path)

    # ---- Generierung ----

    def _generate(self, overrides: dict):
        self.btn_open.config(state=tk.DISABLED, bg="#aaa")
        self.result_label.config(text="")
        self.progress_var.set(0.0)
        self.status_text.set("Wird verarbeitet...")

        def run():
            try:
                info = synthesize_excel(
                    self.input_path.get(),
                    self.output_path.get(),
                    overrides=overrides,
                    progress_callback=self._update_progress,
                )
                # Auditbericht erstellen
                audit_path = generate_audit_report(
                    self.input_path.get(),
                    self.output_path.get(),
                    info["sheets"],
                    info.get("row_counts", {}),
                )
                self.root.after(0, self._on_success, info, audit_path)
            except Exception as exc:
                self.root.after(0, self._on_error, str(exc))

        threading.Thread(target=run, daemon=True).start()

    def _update_progress(self, value: float, status: str):
        self.root.after(0, lambda: self.progress_var.set(value))
        self.root.after(0, lambda: self.status_text.set(status))

    def _on_success(self, info: dict, audit_path: str):
        self.btn_open.config(state=tk.NORMAL, bg=COLOR_BTN_BLUE)
        self.progress_var.set(1.0)
        self.status_text.set("Fertig!")
        self._last_audit_path = audit_path

        all_cols   = {col: ct for sh in info["sheets"].values() for col, ct in sh.items()}
        total      = len(all_cols)
        anonymized = sum(1 for ct in all_cols.values() if ct not in ("text", "beschreibung"))

        self.result_label.config(
            text=f"Fertig! {anonymized}/{total} Spalten anonymisiert.  Auditbericht erstellt.",
            fg=COLOR_BTN_GREEN,
        )

        col_lines = "\n".join(
            f"  {'[OK]' if ct not in ('text','beschreibung') else '[--]'}  {col}"
            for col, ct in all_cols.items()
        )
        msg = (
            f"Synthetische Datei erstellt!\n\n"
            f"Anonymisiert: {anonymized}  |  Unveraendert: {total - anonymized}\n\n"
            f"{col_lines}\n\n"
            f"Auditbericht: {Path(audit_path).name}\n\n"
            f"Was moechten Sie oeffnen?"
        )

        dlg = _SuccessDialog(self.root, msg,
                              excel_path=self.output_path.get(),
                              audit_path=audit_path)
        dlg.wait_window()

    def _on_error(self, error_msg: str):
        self.btn_open.config(state=tk.NORMAL, bg=COLOR_BTN_BLUE)
        self.status_text.set("Fehler aufgetreten")
        self.result_label.config(text=f"Fehler: {error_msg}", fg=COLOR_ERROR)
        messagebox.showerror("Fehler", f"Fehler:\n\n{error_msg}")

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Erfolgs-Dialog mit zwei Buttons
# ---------------------------------------------------------------------------

class _SuccessDialog(tk.Toplevel):
    def __init__(self, parent, message: str, excel_path: str, audit_path: str):
        super().__init__(parent)
        self.title("Fertig!")
        self.resizable(False, False)
        self.configure(bg=COLOR_BG)
        self.grab_set()

        tk.Label(self, text="Synthetische Datei erstellt!", font=FONT_BTN,
                 bg=COLOR_BG, fg=COLOR_BTN_GREEN).pack(pady=(20, 4))
        tk.Label(self, text=message, font=FONT_SMALL, bg=COLOR_BG,
                 justify=tk.LEFT, wraplength=420).pack(padx=24, pady=4)

        btn_frame = tk.Frame(self, bg=COLOR_BG, pady=16)
        btn_frame.pack()

        tk.Button(btn_frame, text="Excel oeffnen",
                  command=lambda: (os.startfile(excel_path), self.destroy()),
                  bg=COLOR_BTN_GREEN, fg="white", font=FONT_SMALL,
                  relief=tk.FLAT, padx=16, pady=7, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=6)

        tk.Button(btn_frame, text="Auditbericht oeffnen",
                  command=lambda: (os.startfile(audit_path), self.destroy()),
                  bg=COLOR_BTN_BLUE, fg="white", font=FONT_SMALL,
                  relief=tk.FLAT, padx=16, pady=7, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=6)

        tk.Button(btn_frame, text="Schliessen", command=self.destroy,
                  bg="#ccc", fg="#333", font=FONT_SMALL,
                  relief=tk.FLAT, padx=16, pady=7, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=6)

        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w,  h  = self.winfo_width(),       self.winfo_height()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")


if __name__ == "__main__":
    SynthesizerApp().run()
