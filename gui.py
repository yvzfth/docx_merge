from __future__ import annotations

import threading
import subprocess
from pathlib import Path
import os
import platform
import logging
import traceback
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from main import collect_summaries


class SummaryQuickMergeApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("SummaryQuickMerge")
        self.geometry("640x260")
        self.minsize(580, 240)

        self._init_logging()

        self.parent_dir_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.include_root_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Pick a parent folder and click Collect")

        self._build_ui()
        self._bind_events()

    def _build_ui(self) -> None:
        padding = {"padx": 12, "pady": 8}

        frm = ttk.Frame(self)
        frm.pack(fill=tk.BOTH, expand=True)

        # Parent folder
        ttk.Label(frm, text="Parent folder:").grid(row=0, column=0, sticky=tk.W, **padding)
        self.parent_entry = ttk.Entry(frm, textvariable=self.parent_dir_var)
        self.parent_entry.grid(row=0, column=1, sticky=tk.EW, **padding)
        self.browse_parent_btn = ttk.Button(frm, text="Browse…", command=self._browse_parent)
        self.browse_parent_btn.grid(row=0, column=2, sticky=tk.W, **padding)

        # Output file
        ttk.Label(frm, text="Output file:").grid(row=1, column=0, sticky=tk.W, **padding)
        self.output_entry = ttk.Entry(frm, textvariable=self.output_file_var)
        self.output_entry.grid(row=1, column=1, sticky=tk.EW, **padding)
        self.browse_output_btn = ttk.Button(frm, text="Save as…", command=self._browse_output)
        self.browse_output_btn.grid(row=1, column=2, sticky=tk.W, **padding)

        # Options
        self.include_root_chk = ttk.Checkbutton(
            frm,
            text="Include .docx files directly in the parent folder",
            variable=self.include_root_var,
        )
        self.include_root_chk.grid(row=2, column=1, sticky=tk.W, **padding)

        # Action row
        self.collect_btn = ttk.Button(frm, text="Collect Summaries", command=self._on_collect)
        self.collect_btn.grid(row=3, column=1, sticky=tk.E, **padding)

        # Status + progress
        self.status_lbl = ttk.Label(frm, textvariable=self.status_var, foreground="#444")
        self.status_lbl.grid(row=4, column=0, columnspan=3, sticky=tk.W, **padding)

        self.progress = ttk.Progressbar(frm, mode="determinate")
        self.progress.grid(row=5, column=0, columnspan=3, sticky=tk.EW, **padding)
        self.progress.configure(value=0, maximum=100)

        # Clear (×) buttons next to pickers
        self.clear_parent_btn = ttk.Button(frm, text="×", width=2, command=lambda: self.parent_dir_var.set(""))
        self.clear_parent_btn.grid(row=0, column=3, sticky=tk.W, **padding)
        self.clear_output_btn = ttk.Button(frm, text="×", width=2, command=lambda: self.output_file_var.set(""))
        self.clear_output_btn.grid(row=1, column=3, sticky=tk.W, **padding)

        frm.columnconfigure(1, weight=1)

    def _bind_events(self) -> None:
        self.parent_dir_var.trace_add("write", lambda *_: self._maybe_set_default_output())

    # Logging setup
    def _log_dir(self) -> Path:
        # Write logs to Desktop for easy access across platforms
        return Path.home() / "Desktop" / "SummaryQuickMergeLogs"

    def _log_path(self) -> Path:
        return self._log_dir() / "app.log"

    def _init_logging(self) -> None:
        try:
            log_dir = self._log_dir()
            log_dir.mkdir(parents=True, exist_ok=True)
            log_path = self._log_path()
            if not logging.getLogger().handlers:
                logging.basicConfig(
                    filename=str(log_path),
                    level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s",
                )
            logging.info("SummaryQuickMerge started")
        except Exception:
            pass

    def _browse_parent(self) -> None:
        initialdir = self._safe_parent_dir()
        chosen = filedialog.askdirectory(initialdir=initialdir or None, title="Select parent folder")
        if chosen:
            self.parent_dir_var.set(chosen)

    def _browse_output(self) -> None:
        parent = self._safe_parent_dir()
        default_name = "collected_summaries.docx"
        initialfile = Path(self.output_file_var.get()).name if self.output_file_var.get() else default_name
        chosen = filedialog.asksaveasfilename(
            title="Save output DOCX as…",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            initialdir=parent or None,
            initialfile=initialfile,
        )
        if chosen:
            self.output_file_var.set(chosen)

    def _safe_parent_dir(self) -> str | None:
        val = self.parent_dir_var.get().strip()
        if not val:
            return None
        try:
            p = Path(val).expanduser().resolve()
            return str(p)
        except Exception:
            return None

    def _maybe_set_default_output(self) -> None:
        parent = self._safe_parent_dir()
        if not parent:
            return
        # Only set default if output is empty or points under old parent default
        out = self.output_file_var.get().strip()
        if not out:
            self.output_file_var.set(str(Path(parent) / "collected_summaries.docx"))

    def _set_running(self, running: bool) -> None:
        widgets = [
            self.parent_entry,
            self.browse_parent_btn,
            self.output_entry,
            self.browse_output_btn,
            self.include_root_chk,
            self.collect_btn,
        ]
        for w in widgets:
            state = tk.DISABLED if running else tk.NORMAL
            w.configure(state=state)
        # Progress bar is determinate; just keep it enabled

    def _on_collect(self) -> None:
        parent = self._safe_parent_dir()
        if not parent:
            messagebox.showwarning("Missing folder", "Please select a valid parent folder.")
            return

        output = self.output_file_var.get().strip()
        if not output:
            output = str(Path(parent) / "collected_summaries.docx")
            self.output_file_var.set(output)

        self._set_running(True)
        self.status_var.set("Collecting summaries…")
        self.progress.configure(value=0, maximum=100)

        include_root = bool(self.include_root_var.get())

        def worker() -> None:
            errors: list[str] = []
            start_ts = time.monotonic()
            timed_out = False

            def should_cancel() -> bool:
                nonlocal timed_out
                if time.monotonic() - start_ts > 60:
                    timed_out = True
                    return True
                return False

            def on_progress(current: int, total: int, path: Path | None) -> None:
                # Marshal to UI thread
                def update() -> None:
                    if total > 0:
                        self.progress.configure(maximum=total)
                        self.progress.configure(value=current)
                    if path is not None:
                        self.status_var.set(f"Processing: {path.name} ({current}/{total})")
                    else:
                        self.status_var.set(f"Found {total} file(s). Starting…")
                self.after(0, update)

            try:
                logging.info("Starting collection: parent=%s output=%s include_root=%s", parent, output, include_root)
                count = collect_summaries(
                    target_root=Path(parent),
                    output_path=Path(output),
                    include_root_files=include_root,
                    on_progress=on_progress,
                    should_cancel=should_cancel,
                    errors=errors,
                )
                if errors:
                    for e in errors[:50]:
                        logging.warning("File error: %s", e)
                def done_ok() -> None:
                    self._set_running(False)
                    # Build detailed completion message
                    if count == 0:
                        self.status_var.set("No Summary sections found. Output created with a notice.")
                        details = "No Summary sections were found. A notice document was created."
                        if errors:
                            details += f"\n\n{len(errors)} file(s) had errors. First errors:\n- " + "\n- ".join(errors[:5])
                        messagebox.showinfo("Done", details)
                    else:
                        self.status_var.set(f"Collected {count} Summary section(s) → {output}")
                        if errors:
                            messagebox.showwarning(
                                "Completed with errors",
                                (
                                    f"Collected {count} Summary section(s).\n"
                                    f"{len(errors)} file(s) failed. First errors:\n- "
                                    + "\n- ".join(errors[:5])
                                ),
                            )
                        if messagebox.askyesno("Open Output", "Open the output file in Finder?"):
                            try:
                                subprocess.run(["open", "-R", output], check=False)
                            except Exception:
                                pass
                self.after(0, done_ok)
            except TimeoutError:
                def done_timeout() -> None:
                    self._set_running(False)
                    self.progress.configure(value=0)
                    self.status_var.set("Timed out after 60s.")
                    messagebox.showerror("Timeout", "The operation timed out after 60 seconds and was stopped.")
                self.after(0, done_timeout)
            except Exception as exc:
                # Log full traceback for packaged app diagnostics
                try:
                    logging.exception("Unhandled error during collection: %s", exc)
                except Exception:
                    pass
                def done_err() -> None:
                    self._set_running(False)
                    self.status_var.set("Failed. See details.")
                    # Show full traceback directly in the popup
                    tb = traceback.format_exc()
                    # Trim very long traces to avoid UI issues
                    if len(tb) > 8000:
                        tb = tb[:7800] + "\n... (truncated)"
                    details = (
                        "An unexpected error occurred.\n\n"
                        f"Type: {type(exc).__name__}\nMessage: {exc}\n\nTraceback:\n{tb}"
                    )
                    messagebox.showerror("Error", details)
                self.after(0, done_err)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = SummaryQuickMergeApp()
    app.mainloop()


