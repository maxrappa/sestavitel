"""
Main application window.

Kept deliberately simple: this file only builds and wires up widgets and
delegates all real work to `excel_processor.process_file`. Processing runs
in a background thread so the window stays responsive, and progress
messages are relayed back to the UI thread through a queue.
"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import constants as C
from process import ProcessingError, process_file


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(C.APP_TITLE)
        self.geometry('760x680')
        self.minsize(640, 480)
        self.resizable(True, True)

        self.selected_file: str | None = None
        self._queue: queue.Queue[tuple[str, object]] = queue.Queue()

        self._build_widgets()
        self._load_config()
        self.after(100, self._drain_queue)

    # ------------------------------------------------------------------
    # Widget layout
    # ------------------------------------------------------------------
    # Root window uses grid() with row weights (instead of pack()) so that
    # resizing/maximizing the window actually grows the text areas: rows
    # holding fixed-height widgets (file picker, buttons, progress bar,
    # status) get weight=0 and stay their natural size, while the top-30
    # box and the log get weight and share any extra space between them.

    def _build_widgets(self) -> None:

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=0)  # file picker
        self.rowconfigure(1, weight=1)  # top-30 codes input
        self.rowconfigure(2, weight=0)  # action buttons
        self.rowconfigure(3, weight=0)  # progress bar
        self.rowconfigure(4, weight=2)  # log (gets the most extra space)
        self.rowconfigure(5, weight=0)  # status line

        padx = 12
        pady = 12

        file_frame = ttk.LabelFrame(self, text='1. Zdrojový soubor')
        file_frame.grid(row=0, column=0, sticky='nsew', padx=padx, pady=pady)
        file_frame.columnconfigure(0, weight=1)

        self.file_label = ttk.Label(file_frame, text='Nevybrán žádný soubor', foreground='#666')
        self.file_label.grid(row=0, column=0, padx=10, pady=10, sticky='ew')

        ttk.Button(file_frame, text='Vybrat soubor…', command=self._choose_file).grid(
            row=0, column=1, padx=10, pady=10
        )

        top_frame = ttk.LabelFrame(
            self,
            text='2. Top 30 nejprodávanějších položek (kódy, jeden na řádek nebo oddělené čárkou)',
        )
        top_frame.grid(row=1, column=0, sticky='nsew', padx=padx, pady=pady)
        top_frame.columnconfigure(0, weight=1)
        top_frame.rowconfigure(0, weight=1)

        self.top_codes_text = scrolledtext.ScrolledText(top_frame, height=8, wrap='word')
        self.top_codes_text.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)

        action_frame = ttk.Frame(self)
        action_frame.grid(row=2, column=0, sticky='ew', padx=padx, pady=pady)

        self.process_button = ttk.Button(
            action_frame, text='Zpracovat', command=self._start_processing
        )
        self.process_button.pack(side='left')

        self.open_output_button = ttk.Button(
            action_frame,
            text='Otevřít výstupní složku',
            command=self._open_output_folder,
            state='disabled',
        )
        self.open_output_button.pack(side='left', padx=10)

        self.progress = ttk.Progressbar(self, mode='indeterminate')
        self.progress.grid(row=3, column=0, sticky='ew', padx=padx, pady=pady)

        log_frame = ttk.LabelFrame(self, text='3. Průběh a chyby')
        log_frame.grid(row=4, column=0, sticky='nsew', padx=padx, pady=pady)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', wrap='word')
        self.log_text.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
        self.log_text.tag_config('error', foreground='#c0392b')
        self.log_text.tag_config('success', foreground='#1e8449')

        self.status_label = ttk.Label(self, text='Připraveno', foreground='#666')
        self.status_label.grid(row=5, column=0, sticky='w', padx=16, pady=(0, 10))

    # ------------------------------------------------------------------
    # File selection
    # ------------------------------------------------------------------

    def _choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title='Vyberte Excel soubor',
            filetypes=[('Excel files', '*.xlsx')],
        )
        if not path:
            return
        self.selected_file = path
        self.file_label.config(text=path, foreground='#000')
        self._log(f'Vybrán soubor: {os.path.basename(path)}')

    # ------------------------------------------------------------------
    # Processing
    # ------------------------------------------------------------------

    def _start_processing(self) -> None:
        if not self.selected_file:
            messagebox.showwarning(C.APP_TITLE, 'Nejprve vyberte zdrojový soubor.')
            return

        top_codes = self._read_top_codes()
        self._save_config(top_codes)

        self.process_button.config(state='disabled')
        self.open_output_button.config(state='disabled')
        self.progress.start(12)
        self.status_label.config(text='Zpracovávám…', foreground='#666')

        thread = threading.Thread(
            target=self._run_processing, args=(self.selected_file, top_codes), daemon=True
        )
        thread.start()

    def _run_processing(self, path: str, top_codes: list[str]) -> None:
        try:
            result = process_file(path, top_codes, progress=self._progress_callback)
            self._queue.put(('done', result))
        except ProcessingError as exc:
            self._queue.put(('error', str(exc)))
        except Exception as exc:  # noqa: BLE001 - last-resort safety net for the UI
            self._queue.put(('error', f'Neočekávaná chyba: {exc}'))

    def _progress_callback(self, message: str) -> None:
        self._queue.put(('progress', message))

    def _drain_queue(self) -> None:
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == 'progress':
                    self._log(payload)  # pyright: ignore[reportArgumentType]
                elif kind == 'done':
                    self._on_success(payload)
                elif kind == 'error':
                    self._on_error(payload)  # pyright: ignore[reportArgumentType]
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    def _on_success(self, result) -> None:
        self.progress.stop()
        self.process_button.config(state='normal')
        self.open_output_button.config(state='normal')
        self._last_output_folder = os.path.dirname(result.output_path)

        self._log(f'Hotovo — {result.row_count} řádků zpracováno.', tag='success')
        self._log(f'Uloženo do: {result.output_path}', tag='success')
        if result.bestseller_matches:
            self._log(f'Zvýrazněno {result.bestseller_matches} top-prodejních položek.')
        if result.unmatched_codes:
            self._log(
                'Nenalezené kódy z top 30 seznamu: ' + ', '.join(result.unmatched_codes),
                tag='error',
            )
        self.status_label.config(text='Hotovo ✓', foreground='#1e8449')

    def _on_error(self, message: str) -> None:
        self.progress.stop()
        self.process_button.config(state='normal')
        self._log(message, tag='error')
        self.status_label.config(text='Chyba ✗', foreground='#c0392b')
        messagebox.showerror(C.APP_TITLE, message)

    def _open_output_folder(self) -> None:
        folder = getattr(self, '_last_output_folder', None)
        if not folder:
            return
        if sys.platform == 'win32':
            os.startfile(folder)  # type: ignore[attr-defined]
        elif sys.platform == 'darwin':
            subprocess.run(['open', folder], check=False)
        else:
            subprocess.run(['xdg-open', folder], check=False)

    # ------------------------------------------------------------------
    # Logging helper
    # ------------------------------------------------------------------

    def _log(self, message: str, tag: str | None = None) -> None:
        self.log_text.config(state='normal')
        self.log_text.insert('end', message + '\n', tag or ())
        self.log_text.see('end')
        self.log_text.config(state='disabled')

    # ------------------------------------------------------------------
    # Top-30 input + persistence
    # ------------------------------------------------------------------

    def _read_top_codes(self) -> list[str]:
        raw = self.top_codes_text.get('1.0', 'end').strip()
        if not raw:
            return []
        parts = raw.replace(',', '\n').splitlines()
        return [p.strip() for p in parts if p.strip()]

    def _load_config(self) -> None:
        if not C.CONFIG_FILE.exists():
            return
        try:
            data = json.loads(C.CONFIG_FILE.read_text(encoding='utf-8'))
            codes = data.get('top_codes', [])
            if codes:
                self.top_codes_text.insert('1.0', '\n'.join(codes))
        except (json.JSONDecodeError, OSError):
            pass  # a corrupt/missing config file should never block startup

    def _save_config(self, top_codes: list[str]) -> None:
        try:
            C.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            C.CONFIG_FILE.write_text(
                json.dumps({'top_codes': top_codes}, ensure_ascii=False, indent=2),
                encoding='utf-8',
            )
        except OSError:
            pass  # persistence is a convenience, not critical to the run


def run() -> None:
    app = App()
    app.mainloop()
