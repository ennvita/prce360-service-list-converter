#!/usr/bin/env python3
"""PRCe360 Service List Converter — NMPRC branded GUI."""

import sys
import io
import contextlib
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ── NMPRC brand palette ───────────────────────────────────────────────────────
DARK_SLATE  = '#2C4043'
TEAL        = '#008585'
TURQUOISE   = '#29A39B'
WHITE_SMOKE = '#F3F3F3'
SOFT_BLACK  = '#242424'
WHITE       = '#FFFFFF'
TEAL_DIM    = '#006666'   # Convert button hover / disabled state

# ── Typography — Segoe UI is the closest Windows system font to Montserrat ────
FONT_BODY   = ('Segoe UI', 10)
FONT_BOLD   = ('Segoe UI', 10, 'bold')
FONT_TITLE  = ('Segoe UI', 13, 'bold')
FONT_SUB    = ('Segoe UI',  9)
FONT_LOG    = ('Consolas',  9)


def _asset(relative: str) -> Path:
    """Resolve a bundled asset path for both frozen and source runs."""
    base = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    return base / relative


class _WidgetStream(io.TextIOBase):
    """Thread-safe stream that forwards writes to a tk.Text widget."""

    def __init__(self, widget: tk.Text):
        self._widget = widget

    def write(self, text: str) -> int:
        self._widget.after(0, self._append, text)
        return len(text)

    def _append(self, text: str) -> None:
        self._widget.configure(state='normal')
        self._widget.insert('end', text)
        self._widget.see('end')
        self._widget.configure(state='disabled')


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PRCe360 Service List Converter")
        self.configure(bg=WHITE_SMOKE)
        self.minsize(660, 500)
        self._logo_img = None
        self._build_ui()
        self._set_defaults()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_header()
        tk.Frame(self, bg=TEAL, height=3).pack(fill='x')
        self._build_body()

    def _build_header(self):
        hdr = tk.Frame(self, bg=DARK_SLATE)
        hdr.pack(fill='x')

        logo_path = _asset('brand/NMPRC_logo-1.png')
        if logo_path.exists():
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_path).convert('RGBA')
                h = 48
                w = int(img.width * h / img.height)
                img = img.resize((w, h), Image.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img)
                tk.Label(hdr, image=self._logo_img, bg=DARK_SLATE,
                         padx=20, pady=12).pack(side='left')
            except Exception:
                pass

        titles = tk.Frame(hdr, bg=DARK_SLATE)
        titles.pack(side='left', pady=14, padx=(0, 20))
        tk.Label(titles, text="PRCe360 Service List Converter",
                 font=FONT_TITLE, fg=WHITE, bg=DARK_SLATE,
                 anchor='w').pack(anchor='w')
        tk.Label(titles, text="New Mexico Public Regulation Commission",
                 font=FONT_SUB, fg=TURQUOISE, bg=DARK_SLATE,
                 anchor='w').pack(anchor='w')

    def _build_body(self):
        body = tk.Frame(self, bg=WHITE_SMOKE, padx=28, pady=22)
        body.pack(fill='both', expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(3, weight=1)

        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar()

        self._file_row(body, "Input file  (.xlsx or .docx):",
                       self.source_var, self._browse_source, row=0)
        self._file_row(body, "Output file  (.xlsx):",
                       self.output_var, self._browse_output, row=1)

        tk.Frame(body, bg=TEAL, height=1).grid(
            row=2, column=0, columnspan=3, sticky='ew', pady=(8, 0))

        btn_wrap = tk.Frame(body, bg=WHITE_SMOKE)
        btn_wrap.grid(row=2, column=0, columnspan=3, pady=(14, 6))
        self.convert_btn = tk.Button(
            btn_wrap, text="Convert",
            font=FONT_BOLD, fg=WHITE, bg=TEAL,
            activebackground=TURQUOISE, activeforeground=WHITE,
            disabledforeground='#aaaaaa',
            relief='flat', padx=36, pady=9, cursor='hand2',
            command=self._run,
        )
        self.convert_btn.pack()

        log_outer = tk.Frame(body, bg=WHITE_SMOKE)
        log_outer.grid(row=3, column=0, columnspan=3, sticky='nsew', pady=(4, 0))
        log_outer.columnconfigure(0, weight=1)
        log_outer.rowconfigure(2, weight=1)

        tk.Label(log_outer, text="Activity Log",
                 font=FONT_BOLD, fg=DARK_SLATE, bg=WHITE_SMOKE,
                 anchor='w').grid(row=0, column=0, columnspan=2, sticky='w')
        tk.Frame(log_outer, bg=TEAL, height=2).grid(
            row=1, column=0, columnspan=2, sticky='ew', pady=(2, 6))

        self.log = tk.Text(
            log_outer, height=8, state='disabled', wrap='word',
            font=FONT_LOG, bg=WHITE, fg=SOFT_BLACK,
            relief='flat', highlightbackground='#c0c8c9',
            highlightthickness=1,
        )
        self.log.grid(row=2, column=0, sticky='nsew')
        sb = ttk.Scrollbar(log_outer, command=self.log.yview)
        sb.grid(row=2, column=1, sticky='ns')
        self.log['yscrollcommand'] = sb.set

    def _file_row(self, parent, label_text: str,
                  var: tk.StringVar, cmd, *, row: int):
        tk.Label(parent, text=label_text,
                 font=FONT_BODY, fg=DARK_SLATE, bg=WHITE_SMOKE,
                 anchor='w').grid(row=row, column=0, sticky='w',
                                  padx=(0, 12), pady=6)

        entry_wrap = tk.Frame(parent, bg='#c0c8c9', padx=1, pady=1)
        entry_wrap.grid(row=row, column=1, sticky='ew', padx=(0, 8), pady=6)
        tk.Entry(entry_wrap, textvariable=var,
                 font=FONT_BODY, fg=SOFT_BLACK, bg=WHITE,
                 relief='flat', bd=0).pack(fill='x', ipady=4, padx=1, pady=1)

        tk.Button(parent, text="Browse…",
                  font=FONT_SUB, fg=WHITE, bg=DARK_SLATE,
                  activebackground=TEAL, activeforeground=WHITE,
                  relief='flat', padx=10, pady=5, cursor='hand2',
                  command=cmd).grid(row=row, column=2, pady=6)

    # ── Defaults ──────────────────────────────────────────────────────────────

    def _set_defaults(self):
        here = (Path(sys.executable).parent if getattr(sys, 'frozen', False)
                else Path(__file__).parent)
        for candidate in ('26-00007-UT-PRC e360.docx',
                          'UPDATED CURRENT Master Telecom List EE 10-6-25.xlsx'):
            if (here / candidate).exists():
                self.source_var.set(str(here / candidate))
                break
        self.output_var.set(str(here / 'output.xlsx'))

    # ── File pickers ──────────────────────────────────────────────────────────

    def _browse_source(self):
        path = filedialog.askopenfilename(
            title="Select input file",
            filetypes=[("Supported files", "*.xlsx *.docx"),
                       ("Excel workbook", "*.xlsx"),
                       ("Word document", "*.docx"),
                       ("All files", "*.*")],
        )
        if path:
            self.source_var.set(path)
            p = Path(path)
            self.output_var.set(str(p.with_name(f"output_{p.stem}.xlsx")))

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save output as",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.output_var.set(path)

    # ── Conversion ────────────────────────────────────────────────────────────

    def _run(self):
        source = self.source_var.get().strip()
        output = self.output_var.get().strip()

        if not source or not output:
            messagebox.showerror("Missing input",
                                 "Please select an input file and output path.")
            return
        if not Path(source).exists():
            messagebox.showerror("File not found",
                                 f"Input file not found:\n{source}")
            return

        template = str(_asset('GridTemplate.xlsx'))
        if not Path(template).exists():
            messagebox.showerror("Template missing",
                                 "GridTemplate.xlsx was not found in the install folder.\n"
                                 "Please reinstall the application.")
            return

        self.convert_btn.configure(state='disabled', text="Converting…",
                                   bg=TEAL_DIM)
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')

        threading.Thread(target=self._do_convert,
                         args=(source, template, output), daemon=True).start()


    def _do_convert(self, source: str, template: str, output: str):
        stream = _WidgetStream(self.log)
        try:
            with contextlib.redirect_stdout(stream):
                ext = Path(source).suffix.lower()
                if ext == '.docx':
                    from convert import convert_docx
                    convert_docx(source, template, output)
                elif ext == '.xlsx':
                    from convert import convert
                    convert(source, template, output)
                else:
                    print(f"Unsupported file type '{ext}'. Use .xlsx or .docx.")
        except Exception as exc:
            stream.write(f"\nError: {exc}\n")
        finally:
            self.after(0, self._done)

    def _done(self):
        self.convert_btn.configure(state='normal', text="Convert", bg=TEAL)


if __name__ == '__main__':
    app = App()
    app.mainloop()
