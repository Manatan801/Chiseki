#!/usr/bin/env python3
from __future__ import annotations

import argparse
import importlib.util
import os
import shutil
import subprocess
from pathlib import Path

from build_setsumeikai_video import (
    DEFAULT_DATE,
    DEFAULT_VIDEO_ID,
    ProjectConfig,
    default_asset_dir,
    default_source_dir,
    ensure_layout,
    extract_docx_text,
    normalize_tts_text,
    parse_segments,
    source_paths,
)


TTS_SCRIPT = Path("/home/maaatan/my_udemy/tts_test/scripts/generate_tts.py")


def load_tts_module():
    spec = importlib.util.spec_from_file_location("generate_tts", TTS_SCRIPT)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load {TTS_SCRIPT}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run(cmd: list[str]) -> None:
    subprocess.run(cmd, check=True)


def make_config(args: argparse.Namespace) -> ProjectConfig:
    asset_dir = args.asset_dir.resolve() if args.asset_dir else default_asset_dir(args.video_id)
    if args.source_dir:
        source_dir = args.source_dir.resolve()
    elif args.asset_dir or args.video_id != DEFAULT_VIDEO_ID:
        source_dir = asset_dir / "source"
    else:
        source_dir = default_source_dir(asset_dir)
    return ProjectConfig(video_id=args.video_id, date=args.date, asset_dir=asset_dir, source_dir=source_dir)


def gemini_voice_name(label: str) -> str:
    if not label:
        return "Sulafat"
    return label if any(char.isupper() for char in label) else label.capitalize()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", default=DEFAULT_DATE)
    parser.add_argument("--video-id", default=DEFAULT_VIDEO_ID)
    parser.add_argument("--asset-dir", type=Path)
    parser.add_argument("--source-dir", type=Path)
    parser.add_argument("--voice", default="sulafat")
    parser.add_argument("--only-slide", type=int, help="Generate only one slide segment without rebuilding the full concat wav")
    args = parser.parse_args()

    config = make_config(args)
    ensure_layout(config.asset_dir)
    _slide_sources, narration_docx, _telop_docx = source_paths(config)

    module = load_tts_module()
    module.load_env(module.ENV_PATH)
    module.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    os.environ["GEMINI_TTS_VOICE"] = gemini_voice_name(args.voice)

    segments = parse_segments(extract_docx_text(narration_docx))
    if args.only_slide is not None:
        segments = [segment for segment in segments if segment.slide == args.only_slide]
        if not segments:
            raise SystemExit(f"No narration segment found for slide {args.only_slide}")
    raw_dir = config.asset_dir / "audio" / "tts_raw"
    final_dir = config.asset_dir / "audio" / "tts_final"
    manifest_dir = config.manifest_dir
    raw_dir.mkdir(parents=True, exist_ok=True)
    final_dir.mkdir(parents=True, exist_ok=True)
    manifest_dir.mkdir(parents=True, exist_ok=True)

    generated: list[Path] = []
    voice_label = args.voice.lower()
    for segment in segments:
        text = normalize_tts_text([segment])
        (manifest_dir / f"{config.video_id}_tts_segment_{segment.slide:03d}_{config.date}.txt").write_text(text, encoding="utf-8")
        print(f"Generating slide {segment.slide:03d}...")
        out = module.generate_gemini(text)
        segment_out = raw_dir / f"{config.video_id}_tts_{voice_label}_segment_{segment.slide:03d}_{config.date}.wav"
        shutil.copy2(out, segment_out)
        generated.append(segment_out)

    if args.only_slide is not None:
        for wav in generated:
            print(wav)
        return 0

    concat_list = manifest_dir / f"{config.video_id}_tts_concat_{config.date}.txt"
    with concat_list.open("w", encoding="utf-8") as f:
        for wav in generated:
            f.write(f"file '{wav}'\n")

    final_wav = final_dir / f"{config.video_id}_tts_{voice_label}_{config.date}.wav"
    run([
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error",
        "-f", "concat", "-safe", "0", "-i", str(concat_list),
        "-c", "copy", str(final_wav),
    ])
    print(final_wav)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
