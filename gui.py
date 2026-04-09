import logging
import threading
import tkinter as tk
from tkinter import scrolledtext, ttk

import assistant
import user_settings


class _TextHandler(logging.Handler):
    def __init__(self, widget: tk.Text):
        super().__init__()
        self.widget = widget

    def emit(self, record):
        msg = self.format(record) + "\n"
        self.widget.after(0, self._append, msg)

    def _append(self, msg):
        self.widget.configure(state="normal")
        self.widget.insert(tk.END, msg)
        self.widget.see(tk.END)
        self.widget.configure(state="disabled")


class SettingsDialog:
    def __init__(self, parent: tk.Tk):
        self._win = tk.Toplevel(parent)
        self._win.title("Settings")
        self._win.configure(bg="#1e1e1e")
        self._win.resizable(False, False)
        self._win.transient(parent)
        self._win.grab_set()

        self._model_var = tk.StringVar()
        self._ptt_var = tk.StringVar()
        self._build_ui()

        s = user_settings.load()
        self._model_var.set(s.get("model", "full"))
        ptt_key = s.get("ptt_key", "ctrl")
        self._ptt_var.set(self._ptt_key_to_label.get(ptt_key, "Either Ctrl"))

        self._win.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width()  - self._win.winfo_width())  // 2
        py = parent.winfo_y() + (parent.winfo_height() - self._win.winfo_height()) // 2
        self._win.geometry(f"+{px}+{py}")

    def _build_ui(self):
        PAD = dict(padx=16, pady=6)
        tk.Label(self._win, text="Speech Recognition Model",
                 bg="#1e1e1e", fg="#c9d1d9",
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", **PAD)

        for value, label in [
            ("small", "Low-end  —  small & fast  (vosk-model-small-en-us-0.15)"),
            ("full",  "Mid/High-end  —  full accuracy  (vosk-model-en-us-0.22)"),
        ]:
            tk.Radiobutton(
                self._win, text=label, variable=self._model_var, value=value,
                bg="#1e1e1e", fg="#c9d1d9", selectcolor="#0d1117",
                activebackground="#1e1e1e", activeforeground="#ffffff",
                font=("Segoe UI", 10),
            ).pack(anchor="w", padx=24, pady=2)

        tk.Frame(self._win, bg="#444", height=1).pack(fill=tk.X, padx=16, pady=(10, 0))

        tk.Label(self._win, text="Push to Talk Key",
                 bg="#1e1e1e", fg="#c9d1d9",
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=16, pady=(10, 4))

        ptt_options = list(user_settings.PTT_KEY_OPTIONS.keys())
        ptt_labels  = [user_settings.PTT_KEY_OPTIONS[k][0] for k in ptt_options]
        ptt_combo = ttk.Combobox(
            self._win, textvariable=self._ptt_var,
            values=ptt_labels, state="readonly", width=20,
            font=("Segoe UI", 10),
        )
        self._ptt_label_to_key = {v[0]: k for k, v in user_settings.PTT_KEY_OPTIONS.items()}
        self._ptt_key_to_label = {k: v[0] for k, v in user_settings.PTT_KEY_OPTIONS.items()}
        ptt_combo.pack(anchor="w", padx=24, pady=(0, 4))
        self._ptt_combo = ptt_combo

        tk.Label(self._win, text="Changes take effect on next start.",
                 bg="#1e1e1e", fg="#888",
                 font=("Segoe UI", 8, "italic")).pack(anchor="w", padx=16, pady=(8, 4))

        btn_row = tk.Frame(self._win, bg="#1e1e1e")
        btn_row.pack(fill=tk.X, padx=16, pady=(4, 12))

        btn_cfg = dict(font=("Segoe UI", 10, "bold"), width=8,
                       relief=tk.FLAT, cursor="hand2", pady=4)
        tk.Button(btn_row, text="Save", bg="#2ea043", fg="white",
                  activebackground="#3fb950",
                  command=self._save, **btn_cfg).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_row, text="Cancel", bg="#444", fg="#c9d1d9",
                  activebackground="#555",
                  command=self._win.destroy, **btn_cfg).pack(side=tk.LEFT)

    def _save(self):
        ptt_key = self._ptt_label_to_key.get(self._ptt_var.get(), "ctrl")
        user_settings.save({
            "model": self._model_var.get(),
            "ptt_key": ptt_key,
        })
        self._win.destroy()


class VAVIApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.thread: threading.Thread | None = None

        root.title("VAVI")
        root.resizable(False, False)
        root.configure(bg="#1e1e1e")

        btn_frame = tk.Frame(root, bg="#1e1e1e", pady=10, padx=10)
        btn_frame.pack(fill=tk.X)

        btn_cfg = dict(font=("Segoe UI", 11, "bold"), width=10,
                       relief=tk.FLAT, cursor="hand2", pady=6)

        self.start_btn = tk.Button(
            btn_frame, text="START", bg="#2ea043", fg="white",
            activebackground="#3fb950", command=self._start, **btn_cfg)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = tk.Button(
            btn_frame, text="STOP", bg="#444", fg="#888",
            activebackground="#da3633", state=tk.DISABLED,
            command=self._stop, **btn_cfg)
        self.stop_btn.pack(side=tk.LEFT)

        self.settings_btn = tk.Button(
            btn_frame, text="Settings", bg="#30363d", fg="#c9d1d9",
            activebackground="#484f58", command=self._open_settings,
            font=("Segoe UI", 11, "bold"), width=10,
            relief=tk.FLAT, cursor="hand2", pady=6)
        self.settings_btn.pack(side=tk.LEFT, padx=(8, 0))

        self.status = tk.Label(btn_frame, text="Idle",
                               bg="#1e1e1e", fg="#888", font=("Segoe UI", 9))
        self.status.pack(side=tk.RIGHT, padx=4)

        log_frame = tk.Frame(root, bg="#1e1e1e")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.log_box = scrolledtext.ScrolledText(
            log_frame, state="disabled", width=70, height=20,
            bg="#0d1117", fg="#c9d1d9", font=("Consolas", 9),
            relief=tk.FLAT, wrap=tk.WORD)
        self.log_box.pack(fill=tk.BOTH, expand=True)

        handler = _TextHandler(self.log_box)
        handler.setFormatter(logging.Formatter(
            "%(asctime)s  %(levelname)-7s  %(message)s", "%H:%M:%S"))
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)

        root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _start(self):
        if self.thread and self.thread.is_alive():
            return
        self.start_btn.config(state=tk.DISABLED, bg="#444", fg="#888")
        self.stop_btn.config(state=tk.NORMAL, bg="#da3633", fg="white",
                             activebackground="#f85149")
        self.status.config(text="Running", fg="#3fb950")
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()
        self.root.after(500, self._poll)

    def _stop(self):
        assistant.stop_event.set()
        self.stop_btn.config(state=tk.DISABLED)
        self.status.config(text="Stopping", fg="#e3b341")

    def _run(self):
        try:
            assistant.main()
        except Exception as exc:
            logging.getLogger("vavi.gui").error("Assistant crashed: %s", exc)
        finally:
            assistant.stop_event.set()

    def _poll(self):
        if self.thread and self.thread.is_alive():
            self.root.after(500, self._poll)
        else:
            self.start_btn.config(state=tk.NORMAL, bg="#2ea043", fg="white")
            self.stop_btn.config(state=tk.DISABLED, bg="#444", fg="#888")
            self.status.config(text="● Idle", fg="#888")

    def _open_settings(self):
        SettingsDialog(self.root)

    def _on_close(self):
        assistant.stop_event.set()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    VAVIApp(root)
    root.mainloop()
