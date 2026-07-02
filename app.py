"""
Coor-Abgleich-Tool - grafische Oberflaeche.

Zwei Drop-Zonen (C16-Export links, eigene Eingangsrechnungen-Excel
rechts), Knopf "Abgleichen", Ausgabe von zwei Dateien:
  <name>_markiert.xlsx   - Original mit gelben Markierungen bei Abweichungen
  <name>_korrigiert.xlsx - wie oben, zusaetzlich eindeutige Faelle automatisch korrigiert

Startet als eigenstaendiges Fenster. Fuer die Windows-.exe wird dieses
Skript mit PyInstaller gepackt (siehe .github/workflows/build-windows-exe.yml).
"""

import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

import core

APP_TITLE = "Coor-Abgleich-Tool"


class DropZone(tk.Frame):
    def __init__(self, master, label_text, on_file, **kw):
        super().__init__(master, relief="groove", borderwidth=2, bg="#f5f5f5", **kw)
        self.on_file = on_file
        self.path = None

        self.title_label = tk.Label(self, text=label_text, font=("Segoe UI", 11, "bold"), bg="#f5f5f5")
        self.title_label.pack(pady=(14, 4))

        self.hint_label = tk.Label(
            self,
            text="Datei hierher ziehen\n(oder klicken zum Auswaehlen)",
            fg="#666666",
            bg="#f5f5f5",
            justify="center",
        )
        self.hint_label.pack(pady=4)

        self.file_label = tk.Label(self, text="", fg="#0a7a0a", bg="#f5f5f5", wraplength=220, justify="center")
        self.file_label.pack(pady=(4, 14))

        for widget in (self, self.title_label, self.hint_label, self.file_label):
            widget.bind("<Button-1>", self._browse)

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._drop)

    def _browse(self, _event=None):
        path = filedialog.askopenfilename(
            title="Excel-Datei waehlen", filetypes=[("Excel-Dateien", "*.xlsx *.xlsm")]
        )
        if path:
            self._set_path(path)

    def _drop(self, event):
        raw = event.data
        # tkinterdnd2 liefert Pfade ggf. mit geschweiften Klammern bei Leerzeichen
        path = raw.strip("{}")
        if path.lower().endswith((".xlsx", ".xlsm")):
            self._set_path(path)
        else:
            messagebox.showwarning(APP_TITLE, "Bitte eine .xlsx- oder .xlsm-Datei ablegen.")

    def _set_path(self, path):
        self.path = path
        self.file_label.config(text=os.path.basename(path))
        self.config(bg="#eaf6ea")
        for w in (self.title_label, self.hint_label, self.file_label):
            w.config(bg="#eaf6ea")
        self.on_file()


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("640x480")
        root.minsize(560, 440)

        header = tk.Label(
            root,
            text="C16-Buchungen mit eigener Eingangsrechnungen-Liste abgleichen",
            font=("Segoe UI", 12, "bold"),
        )
        header.pack(pady=(14, 6))

        if not HAS_DND:
            warn = tk.Label(
                root,
                text="Hinweis: Drag & Drop ist in dieser Installation nicht aktiv - bitte Dateien per Klick auswaehlen.",
                fg="#a05a00",
            )
            warn.pack(pady=(0, 6))

        zones = tk.Frame(root)
        zones.pack(fill="both", expand=False, padx=16, pady=8)
        zones.columnconfigure(0, weight=1)
        zones.columnconfigure(1, weight=1)

        self.c16_zone = DropZone(zones, "1. Coor / C16-Export", self._check_ready, width=260, height=140)
        self.c16_zone.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self.own_zone = DropZone(zones, "2. Eigene Eingangsrechnungen-Excel", self._check_ready, width=260, height=140)
        self.own_zone.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self.run_btn = tk.Button(
            root, text="Abgleichen", state="disabled", font=("Segoe UI", 11, "bold"),
            command=self._run, bg="#2e7d32", fg="white", padx=16, pady=8,
        )
        self.run_btn.pack(pady=12)

        self.progress = ttk.Progressbar(root, mode="indeterminate")

        log_frame = tk.Frame(root)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        self.log = tk.Text(log_frame, height=10, state="disabled", bg="#111111", fg="#dddddd", font=("Consolas", 10))
        self.log.pack(fill="both", expand=True)

        self._log("Bereit. Bitte beide Dateien auswaehlen bzw. hineinziehen.")

    def _check_ready(self):
        if self.c16_zone.path and self.own_zone.path:
            self.run_btn.config(state="normal")

    def _log(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _run(self):
        out_dir = filedialog.askdirectory(
            title="Ordner fuer die Ergebnis-Dateien waehlen",
            initialdir=os.path.dirname(self.own_zone.path),
        )
        if not out_dir:
            return
        self.run_btn.config(state="disabled")
        self.progress.pack(fill="x", padx=16, pady=(0, 8))
        self.progress.start(10)
        self._log("\nStarte Abgleich ...")
        thread = threading.Thread(target=self._run_worker, args=(out_dir,), daemon=True)
        thread.start()

    def _run_worker(self, out_dir):
        try:
            summary, marked_path, corrected_path, log_lines = core.run_abgleich(
                self.c16_zone.path, self.own_zone.path, out_dir
            )
            self.root.after(0, self._on_done, log_lines, marked_path, corrected_path)
        except Exception as exc:
            tb = traceback.format_exc()
            self.root.after(0, self._on_error, str(exc), tb)

    def _on_done(self, log_lines, marked_path, corrected_path):
        self.progress.stop()
        self.progress.pack_forget()
        for line in log_lines:
            self._log(line)
        self._log(f"\nMarkierte Datei: {marked_path}")
        self._log(f"Korrigierte Datei: {corrected_path}")
        self.run_btn.config(state="normal")
        messagebox.showinfo(
            APP_TITLE,
            "Abgleich fertig!\n\n"
            f"Markierte Datei:\n{marked_path}\n\n"
            f"Korrigierte Datei:\n{corrected_path}",
        )

    def _on_error(self, message, tb):
        self.progress.stop()
        self.progress.pack_forget()
        self.run_btn.config(state="normal")
        self._log(f"\nFEHLER: {message}")
        self._log(tb)
        messagebox.showerror(APP_TITLE, f"Fehler beim Abgleich:\n\n{message}")


def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
