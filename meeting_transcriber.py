#!/usr/bin/env python3
"""會議錄音轉文字工具（核心 + 指令列）。

使用 faster-whisper 在本機離線辨識語音，不會把錄音上傳到任何雲端。

用法：
    # 把現有錄音檔轉成文字（支援 mp3 / m4a / wav / mp4 等）
    python meeting_transcriber.py file 會議錄音.m4a

    # 直接用麥克風錄音，按 Ctrl+C 停止後自動轉文字
    python meeting_transcriber.py record

輸出（與錄音檔放在同一個資料夾）：
    xxx.txt  每句前面附時間戳記的逐字稿
    xxx.srt  字幕檔，可搭配錄音/影片播放對照
"""

from __future__ import annotations

import argparse
import datetime as _dt
import queue
import sys
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

# 提示模型使用繁體中文與標點，可大幅減少簡體字輸出
DEFAULT_PROMPT = "以下是一場會議的逐字稿，使用繁體中文，並加上標點符號。"

MODEL_CHOICES = ["tiny", "base", "small", "medium", "large-v3", "large-v3-turbo"]


@dataclass
class Segment:
    start: float
    end: float
    text: str


# ---------------------------------------------------------------- 時間格式 --

def _fmt_clock(seconds: float) -> str:
    seconds = int(seconds)
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def _fmt_srt(seconds: float) -> str:
    ms = int(round(seconds * 1000))
    h, rem = divmod(ms, 3_600_000)
    m, rem = divmod(rem, 60_000)
    s, ms = divmod(rem, 1000)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


# ---------------------------------------------------------------- 簡轉繁 --

def _make_converter() -> Callable[[str], str]:
    """回傳簡體→台灣繁體的轉換函式；沒裝 opencc 時原樣輸出。"""
    try:
        from opencc import OpenCC  # type: ignore
    except ImportError:
        return lambda s: s
    try:
        cc = OpenCC("s2twp")
    except Exception:  # 不同 opencc 套件的設定檔名稱略有差異
        cc = OpenCC("s2twp.json")
    return cc.convert


# ---------------------------------------------------------------- 辨識 --

class Transcriber:
    """包裝 faster-whisper，模型只載入一次可重複使用。"""

    def __init__(self, model_size: str = "small", device: str = "auto",
                 compute_type: str = "auto"):
        from faster_whisper import WhisperModel

        if compute_type == "auto":
            compute_type = "int8" if device in ("auto", "cpu") else "float16"
        self.model = WhisperModel(model_size, device=device,
                                  compute_type=compute_type)
        self._to_trad = _make_converter()

    def transcribe(
        self,
        audio_path: str | Path,
        language: str | None = "zh",
        prompt: str | None = DEFAULT_PROMPT,
        on_segment: Callable[[Segment, float], None] | None = None,
    ) -> list[Segment]:
        """辨識整個檔案。on_segment(seg, 總長秒數) 可用來顯示進度。"""
        segments, info = self.model.transcribe(
            str(audio_path),
            language=language or None,
            initial_prompt=prompt or None,
            vad_filter=True,  # 略過靜音段，速度更快、幻覺更少
            vad_parameters={"min_silence_duration_ms": 500},
            beam_size=5,
        )
        result: list[Segment] = []
        for s in segments:
            text = s.text.strip()
            if language in (None, "zh"):
                text = self._to_trad(text)
            seg = Segment(s.start, s.end, text)
            result.append(seg)
            if on_segment:
                on_segment(seg, info.duration)
        return result


# ---------------------------------------------------------------- 輸出 --

def write_outputs(segments: Iterable[Segment], audio_path: str | Path,
                  out_dir: str | Path | None = None) -> tuple[Path, Path]:
    audio_path = Path(audio_path)
    out_dir = Path(out_dir) if out_dir else audio_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    base = out_dir / audio_path.stem
    segments = list(segments)

    txt_path = base.with_suffix(".txt")
    with txt_path.open("w", encoding="utf-8") as f:
        f.write(f"會議逐字稿：{audio_path.name}\n")
        f.write(f"產生時間：{_dt.datetime.now():%Y-%m-%d %H:%M}\n\n")
        for seg in segments:
            f.write(f"[{_fmt_clock(seg.start)}] {seg.text}\n")

    srt_path = base.with_suffix(".srt")
    with srt_path.open("w", encoding="utf-8") as f:
        for i, seg in enumerate(segments, 1):
            f.write(f"{i}\n{_fmt_srt(seg.start)} --> {_fmt_srt(seg.end)}\n"
                    f"{seg.text}\n\n")

    return txt_path, srt_path


# ---------------------------------------------------------------- 錄音 --

class Recorder:
    """從麥克風錄音並即時寫入 WAV 檔（16kHz 單聲道，語音辨識最適合）。

    邊錄邊寫檔，就算程式中途當掉，已錄的部分也不會遺失。
    """

    SAMPLE_RATE = 16000

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self._q: queue.Queue = queue.Queue()
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self.seconds = 0.0
        self.level = 0.0  # 目前音量 0~1，可做音量條

    def _callback(self, indata, frames, time_info, status):
        self._q.put(indata.copy())
        self.level = float(min(1.0, abs(indata).max()))

    def _run(self):
        import sounddevice as sd
        import soundfile as sf

        self.path.parent.mkdir(parents=True, exist_ok=True)
        with sf.SoundFile(self.path, mode="w", samplerate=self.SAMPLE_RATE,
                          channels=1, subtype="PCM_16") as f, \
             sd.InputStream(samplerate=self.SAMPLE_RATE, channels=1,
                            dtype="float32", callback=self._callback):
            while not self._stop.is_set() or not self._q.empty():
                try:
                    block = self._q.get(timeout=0.2)
                except queue.Empty:
                    continue
                f.write(block)
                self.seconds += len(block) / self.SAMPLE_RATE

    def start(self):
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> Path:
        self._stop.set()
        if self._thread:
            self._thread.join()
        return self.path


def default_recording_path(folder: str | Path = "recordings") -> Path:
    return Path(folder) / f"會議_{_dt.datetime.now():%Y%m%d_%H%M%S}.wav"


# ---------------------------------------------------------------- 指令列 --

def _progress_printer(seg: Segment, total: float):
    pct = f"{seg.end / total * 100:5.1f}%" if total else ""
    print(f"{pct} [{_fmt_clock(seg.start)}] {seg.text}", flush=True)


def _transcribe_and_save(args, audio_paths: list[Path]):
    print(f"載入模型 {args.model} 中（第一次使用會自動下載）…", flush=True)
    tr = Transcriber(args.model, device=args.device)
    for path in audio_paths:
        print(f"\n▶ 開始辨識：{path}", flush=True)
        segs = tr.transcribe(path, language=args.language or None,
                             prompt=args.prompt, on_segment=_progress_printer)
        txt, srt = write_outputs(segs, path, args.out)
        print(f"✔ 完成：{txt}\n         {srt}")


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="會議錄音轉文字（本機離線辨識）")
    p.add_argument("--model", default="small", choices=MODEL_CHOICES,
                   help="模型大小：越大越準但越慢（預設 small；"
                        "有獨立顯卡建議 large-v3-turbo）")
    p.add_argument("--language", default="zh",
                   help="語言代碼，預設 zh；中英夾雜也用 zh；設為空字串自動偵測")
    p.add_argument("--prompt", default=DEFAULT_PROMPT,
                   help="提示詞，可加入公司名稱、專有名詞提高辨識率")
    p.add_argument("--device", default="auto", choices=["auto", "cpu", "cuda"])
    p.add_argument("--out", help="輸出資料夾（預設與錄音檔相同）")
    sub = p.add_subparsers(dest="cmd", required=True)

    f = sub.add_parser("file", help="把現有錄音檔轉成文字")
    f.add_argument("audio", nargs="+", type=Path, help="錄音檔（可多個）")

    r = sub.add_parser("record", help="用麥克風錄音，Ctrl+C 停止後轉文字")
    r.add_argument("--save-to", type=Path, help="錄音檔路徑（預設 recordings/會議_時間.wav）")
    r.add_argument("--no-transcribe", action="store_true", help="只錄音，不轉文字")

    args = p.parse_args(argv)

    if args.cmd == "file":
        missing = [a for a in args.audio if not a.exists()]
        if missing:
            print("找不到檔案：" + ", ".join(map(str, missing)), file=sys.stderr)
            return 1
        _transcribe_and_save(args, args.audio)
        return 0

    path = args.save_to or default_recording_path()
    rec = Recorder(path)
    rec.start()
    print(f"● 錄音中… 存到 {path}\n  按 Ctrl+C 停止", flush=True)
    try:
        while True:
            threading.Event().wait(1)
            bar = "█" * int(rec.level * 20)
            print(f"\r  {_fmt_clock(rec.seconds)}  {bar:<20}", end="", flush=True)
    except KeyboardInterrupt:
        pass
    rec.stop()
    print(f"\n■ 錄音結束，共 {_fmt_clock(rec.seconds)}：{path}")
    if not args.no_transcribe:
        _transcribe_and_save(args, [path])
    return 0


if __name__ == "__main__":
    sys.exit(main())
