#!/usr/bin/env python3
"""
Local long-form transcription using faster-whisper (Whisper weights, chunked decode).
Requires ffmpeg installed on your system for non-WAV formats.
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Transcribe long audio locally with Whisper (faster-whisper)."
    )
    p.add_argument(
        "audio",
        type=Path,
        help="Path to audio file (wav, mp3, m4a, flac, etc.; needs ffmpeg for most codecs).",
    )
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Write transcript here (default: <audio_stem>.txt next to file).",
    )
    p.add_argument(
        "-m",
        "--model",
        default="medium",
        choices=[
            "tiny",
            "tiny.en",
            "base",
            "base.en",
            "small",
            "small.en",
            "medium",
            "medium.en",
            "large-v1",
            "large-v2",
            "large-v3",
        ],
        help="Model size (default: medium). Larger = more accurate, slower, more RAM/VRAM.",
    )
    p.add_argument(
        "--device",
        default="auto",
        choices=["auto", "cpu", "cuda"],
        help="Compute device (default: auto).",
    )
    p.add_argument(
        "--compute-type",
        default=None,
        help="Override faster-whisper compute type, e.g. int8, float16. "
        "Default: sensible per device (int8 on CPU, float16 on CUDA).",
    )
    p.add_argument(
        "--language",
        default=None,
        metavar="CODE",
        help="ISO language code (e.g. en). Omit for auto-detect.",
    )
    p.add_argument(
        "--beam-size",
        type=int,
        default=5,
        help="Beam size for decoding (default: 5).",
    )
    p.add_argument(
        "--no-vad",
        action="store_true",
        help="Disable VAD filtering (not recommended for long noisy recordings).",
    )
    p.add_argument(
        "--timestamps",
        action="store_true",
        help="Prefix each line with [start-end] seconds.",
    )
    return p


def pick_compute_type(device: str, explicit: str | None) -> str:
    if explicit:
        return explicit
    if device == "cuda":
        return "float16"
    return "int8"


def main() -> int:
    args = build_parser().parse_args()
    audio: Path = args.audio.expanduser().resolve()

    if not audio.is_file():
        print(f"Not a file: {audio}", file=sys.stderr)
        return 1

    out_path = args.output
    if out_path is None:
        out_path = audio.with_suffix(".txt")
    else:
        out_path = out_path.expanduser().resolve()

    try:
        from faster_whisper import WhisperModel
    except ImportError:
        print(
            "Install dependencies: pip install -r requirements.txt",
            file=sys.stderr,
        )
        return 1

    device = args.device
    if device == "auto":
        import ctranslate2

        device = "cuda" if ctranslate2.get_cuda_device_count() > 0 else "cpu"

    compute_type = pick_compute_type(device, args.compute_type)

    print(
        f"Loading model {args.model!r} ({device}, {compute_type}) — first run downloads weights.",
        flush=True,
    )
    t0 = time.perf_counter()
    model = WhisperModel(
        args.model,
        device=device,
        compute_type=compute_type,
    )
    print(f"Model ready in {time.perf_counter() - t0:.1f}s", flush=True)

    language = args.language
    task_kwargs = {
        "beam_size": args.beam_size,
        "vad_filter": not args.no_vad,
    }
    if language:
        task_kwargs["language"] = language

    print(f"Transcribing {audio} ...", flush=True)
    t1 = time.perf_counter()
    segments, info = model.transcribe(str(audio), **task_kwargs)

    lines: list[str] = []
    if info.language:
        det = f"detected_language={info.language}"
        if getattr(info, "language_probability", None) is not None:
            det += f" (p={info.language_probability:.2f})"
        lines.append(f"# {det}")
        lines.append("")

    dur_hint = getattr(info, "duration", None)
    if dur_hint:
        lines.append(f"# duration≈{dur_hint:.1f}s")
        lines.append("")

    char_count = 0
    for seg in segments:
        text = (seg.text or "").strip()
        if not text:
            continue
        if args.timestamps:
            line = f"[{seg.start:.2f}-{seg.end:.2f}] {text}"
        else:
            line = text
        lines.append(line)
        char_count += len(text)
        print(line, flush=True)

    body = "\n".join(lines).rstrip() + "\n"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(body, encoding="utf-8")

    elapsed = time.perf_counter() - t1
    rtf = elapsed / float(dur_hint) if dur_hint and dur_hint > 0 else 0.0
    print(
        f"\nWrote {out_path} ({char_count} chars) in {elapsed:.1f}s"
        + (f" ~ {rtf:.2f}× realtime" if rtf else ""),
        flush=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
