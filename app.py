#!/usr/bin/env python3
"""會議錄音轉文字：圖形介面版。

雙擊執行（或 `python app.py`），按「開始錄音」→ 開會 → 「停止並轉文字」，
完成後逐字稿會顯示在視窗中，並自動存成 .txt / .srt。
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from meeting_transcriber import (DEFAULT_PROMPT, MODEL_CHOICES, Recorder,
                                 Transcriber, _fmt_clock,
                                 default_recording_path, write_outputs)

AUDIO_TYPES = [("音訊/影片", "*.wav *.mp3 *.m4a *.aac *.flac *.ogg *.wma *.mp4 *.mov *.webm"),
               ("所有檔案", "*.*")]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("會議錄音轉文字")
        self.geometry("760x560")
        self.minsize(560, 420)

        self.recorder: Recorder | None = None
        self.transcriber: Transcriber | None = None
        self.loaded_model: str | None = None
        self.busy = False
        self.last_txt: Path | None = None

        self._build()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------ 介面 --
    def _build(self):
        pad = {"padx": 8, "pady": 4}

        top = ttk.Frame(self)
        top.pack(fill="x", **pad)
        self.rec_btn = ttk.Button(top, text="● 開始錄音", command=self.toggle_record)
        self.rec_btn.pack(side="left")
        self.file_btn = ttk.Button(top, text="📂 選擇錄音檔轉文字", command=self.pick_file)
        self.file_btn.pack(side="left", padx=6)
        self.open_btn = ttk.Button(top, text="開啟逐字稿", command=self.open_result,
                                   state="disabled")
        self.open_btn.pack(side="left")

        ttk.Label(top, text="模型").pack(side="left", padx=(16, 2))
        self.model_var = tk.StringVar(value="small")
        ttk.Combobox(top, textvariable=self.model_var, values=MODEL_CHOICES,
                     width=14, state="readonly").pack(side="left")

        pf = ttk.Frame(self)
        pf.pack(fill="x", **pad)
        ttk.Label(pf, text="專有名詞提示").pack(side="left")
        self.prompt_var = tk.StringVar(value=DEFAULT_PROMPT)
        ttk.Entry(pf, textvariable=self.prompt_var).pack(side="left", fill="x",
                                                         expand=True, padx=4)

        sf = ttk.Frame(self)
        sf.pack(fill="x", **pad)
        self.status_var = tk.StringVar(value="準備就緒")
        ttk.Label(sf, textvariable=self.status_var).pack(side="left")
        self.level = ttk.Progressbar(sf, length=120, maximum=1.0)
        self.level.pack(side="right")
        self.progress = ttk.Progressbar(self, maximum=100)
        self.progress.pack(fill="x", **pad)

        self.text = scrolledtext.ScrolledText(self, wrap="word", font=("Microsoft JhengHei", 12))
        self.text.pack(fill="both", expand=True, **pad)

    def _set_busy(self, busy: bool):
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.rec_btn.config(state=state)
        self.file_btn.config(state=state)

    # ------------------------------------------------------------ 錄音 --
    def toggle_record(self):
        if self.recorder is None:
            path = default_recording_path(Path.home() / "會議錄音")
            try:
                self.recorder = Recorder(path)
                self.recorder.start()
            except Exception as e:  # 沒有麥克風等
                self.recorder = None
                messagebox.showerror("無法錄音", str(e))
                return
            self.rec_btn.config(text="■ 停止並轉文字")
            self.file_btn.config(state="disabled")
            self._tick()
        else:
            path = self.recorder.stop()
            self.recorder = None
            self.rec_btn.config(text="● 開始錄音")
            self.level["value"] = 0
            self.start_transcribe(path)

    def _tick(self):
        if self.recorder is None:
            return
        self.status_var.set(f"● 錄音中 {_fmt_clock(self.recorder.seconds)}"
                            f"  （存到 {self.recorder.path}）")
        self.level["value"] = self.recorder.level
        self.after(200, self._tick)

    # ------------------------------------------------------------ 辨識 --
    def pick_file(self):
        path = filedialog.askopenfilename(title="選擇錄音檔", filetypes=AUDIO_TYPES)
        if path:
            self.start_transcribe(Path(path))

    def start_transcribe(self, path: Path):
        self._set_busy(True)
        self.text.delete("1.0", "end")
        self.progress["value"] = 0
        threading.Thread(target=self._worker, args=(path,), daemon=True).start()

    def _ui(self, fn, *a):
        self.after(0, fn, *a)

    def _worker(self, path: Path):
        try:
            model = self.model_var.get()
            if self.transcriber is None or self.loaded_model != model:
                self._ui(self.status_var.set, f"載入模型 {model}（第一次會自動下載，請稍候）…")
                self.transcriber = Transcriber(model)
                self.loaded_model = model
            self._ui(self.status_var.set, f"辨識中：{path.name}")

            def on_seg(seg, total):
                self._ui(self.text.insert, "end", f"[{_fmt_clock(seg.start)}] {seg.text}\n")
                self._ui(self.text.see, "end")
                if total:
                    self._ui(self.progress.config, {"value": min(100, seg.end / total * 100)})

            segs = self.transcriber.transcribe(path, prompt=self.prompt_var.get(),
                                               on_segment=on_seg)
        except Exception as e:
            self._ui(messagebox.showerror, "辨識失敗", str(e))
            self._ui(self.status_var.set, "發生錯誤")
            self._ui(self._set_busy, False)
            return
        self._ui(self._done, path, segs)

    def _done(self, path: Path, segs):
        txt, _ = write_outputs(segs, path)
        self.last_txt = txt
        self.progress["value"] = 100
        self.open_btn.config(state="normal")
        self.status_var.set(f"✔ 完成，已存檔：{txt}")
        self._set_busy(False)

    def open_result(self):
        if not self.last_txt:
            return
        if sys.platform.startswith("win"):
            os.startfile(self.last_txt)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(self.last_txt)])
        else:
            subprocess.Popen(["xdg-open", str(self.last_txt)])

    def _on_close(self):
        if self.recorder is not None:
            if not messagebox.askyesno("還在錄音", "正在錄音中，確定要結束？（已錄部分會保留）"):
                return
            self.recorder.stop()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
