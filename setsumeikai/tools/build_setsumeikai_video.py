#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import math
import re
import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path("/home/maaatan/Chiseki/setsumeikai")
DEPS = ROOT / ".deps"
if str(DEPS) not in sys.path:
    sys.path.insert(0, str(DEPS))

import fitz  # type: ignore  # noqa: E402


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
DEFAULT_DATE = "2026-05-07"
DEFAULT_VIDEO_ID = "setsumeikai_chiseki_overview"
SLIDE_SOURCE_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}

# Backward-compatible aliases for callers that imported the original constants.
DATE = DEFAULT_DATE
VIDEO_ID = DEFAULT_VIDEO_ID


@dataclass(frozen=True)
class ProjectConfig:
    video_id: str
    date: str
    asset_dir: Path
    source_dir: Path

    @property
    def manifest_dir(self) -> Path:
        return self.asset_dir / "manifests"

    @property
    def build_dir(self) -> Path:
        return self.asset_dir / "build"

    @property
    def slide_export_dir(self) -> Path:
        return self.asset_dir / "slides" / "exports"


@dataclass
class Segment:
    slide: int
    title: str
    body: str
    chars: int
    start: float = 0
    end: float = 0
    pre_read_silence: float = 0
    audio_start: float = 0
    audio_end: float = 0
    audio_duration: float = 0
    tail_hold: float = 0
    audio_path: str = ""


def run(cmd: list[str]) -> str:
    completed = subprocess.run(cmd, check=True, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    return completed.stdout.strip()


def default_asset_dir(video_id: str) -> Path:
    return ROOT / "video_assets" / "setsumeikai" / video_id


def default_source_dir(asset_dir: Path) -> Path:
    source_dir = asset_dir / "source"
    if source_dir.exists():
        return source_dir
    return ROOT


def make_config(args: argparse.Namespace) -> ProjectConfig:
    video_id = args.video_id
    date = args.date
    asset_dir = args.asset_dir.resolve() if args.asset_dir else default_asset_dir(video_id)
    if args.source_dir:
        source_dir = args.source_dir.resolve()
    elif args.asset_dir or video_id != DEFAULT_VIDEO_ID:
        source_dir = asset_dir / "source"
    else:
        source_dir = default_source_dir(asset_dir)
    return ProjectConfig(video_id=video_id, date=date, asset_dir=asset_dir, source_dir=source_dir)


def slides_source(config: ProjectConfig) -> Path:
    if config.source_dir == ROOT:
        return ROOT / "スライド"
    return config.source_dir / "slides_pdf"


def ensure_layout(asset_dir: Path) -> None:
    for rel in [
        "source/slides_pdf", "source/narration", "source/telops",
        "slides/exports", "audio/music_raw", "audio/music_clipped", "audio/tts_raw",
        "audio/tts_final", "audio/mixed", "images/raw", "images/final", "manifests",
        "build", "final_refs",
    ]:
        (asset_dir / rel).mkdir(parents=True, exist_ok=True)


def first_docx(folder: Path, *, fallback: Path | None = None) -> Path:
    if folder.exists():
        docs = sorted(path for path in folder.glob("*.docx") if ":Zone.Identifier" not in path.name)
        if docs:
            return docs[0]
    if fallback and fallback.exists():
        return fallback
    raise FileNotFoundError(f"No .docx found in {folder}")


def optional_docx(folder: Path, *, fallback: Path | None = None) -> Path | None:
    try:
        return first_docx(folder, fallback=fallback)
    except FileNotFoundError:
        return None


def source_paths(config: ProjectConfig) -> tuple[list[Path], Path, Path | None]:
    if config.source_dir == ROOT:
        slide_source_dir = ROOT / "スライド"
        narration_docx = ROOT / "読み原稿" / "AI音声　読み原稿　説明会.docx"
        telop_docx = optional_docx(ROOT / "読み原稿", fallback=ROOT / "読み原稿" / "テロップ原稿　説明会.docx")
    else:
        slide_source_dir = config.source_dir / "slides_pdf"
        narration_docx = first_docx(config.source_dir / "narration")
        telop_docx = optional_docx(config.source_dir / "telops")

    slide_sources = sorted(
        path for path in slide_source_dir.iterdir()
        if path.is_file()
        and ":Zone.Identifier" not in path.name
        and path.suffix.lower() in SLIDE_SOURCE_EXTENSIONS
    )
    if not slide_sources:
        raise FileNotFoundError(f"No slide source found in {slide_source_dir}")
    if not narration_docx.exists():
        raise FileNotFoundError(f"Missing narration docx: {narration_docx}")
    return slide_sources, narration_docx, telop_docx


def extract_docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    paragraphs: list[str] = []
    for paragraph in root.findall(".//w:p", NS):
        parts = [node.text or "" for node in paragraph.findall(".//w:t", NS)]
        text = "".join(parts).strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs)


def parse_segments(text: str) -> list[Segment]:
    slide_re = re.compile(r"^スライド\s*([０-９0-9一二三四五六七八九十じゅういちにさんよんごろくななはちきゅう]+)[：、,:]?\s*(.*)$")
    number_map = {
        "１": 1, "２": 2, "３": 3, "４": 4, "５": 5, "６": 6, "７": 7, "８": 8, "９": 9,
        "１０": 10, "１１": 11, "１２": 12, "いち": 1, "に": 2, "さん": 3, "よん": 4,
        "ご": 5, "ろく": 6, "なな": 7, "はち": 8, "きゅう": 9, "じゅう": 10,
        "じゅういち": 11, "じゅうに": 12,
    }
    segments: list[Segment] = []
    current_slide = 1
    current_title = "地籍調査とは？"
    current_lines: list[str] = []

    def flush() -> None:
        nonlocal current_lines
        body = "\n".join(line for line in current_lines if line and "スライドをきりかえ" not in line and not line.startswith("【指示】"))
        body = re.sub(r"^\s*ついかスライド.*$", "", body, flags=re.MULTILINE).strip()
        if body:
            chars = len(re.sub(r"\s+", "", body))
            segments.append(Segment(current_slide, current_title, body, max(chars, 20)))
        current_lines = []

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        match = slide_re.match(line)
        if match:
            flush()
            token = match.group(1)
            parsed_slide = number_map.get(token)
            if parsed_slide is None and (token[:1].isdigit() or token[:1] in "０１２３４５６７８９"):
                parsed_slide = int(token.translate(str.maketrans("０１２３４５６７８９", "0123456789")))
            if parsed_slide is None:
                parsed_slide = current_slide + 1
            current_slide = parsed_slide
            current_title = match.group(2).strip() or f"スライド{current_slide}"
            continue
        current_lines.append(line)
    flush()
    return sorted(segments, key=lambda item: item.slide)


def normalize_tts_text(segments: list[Segment]) -> str:
    return "\n\n".join(segment.body for segment in segments)


def render_pdf(pdf: Path, output_dir: Path, start_index: int, video_id: str) -> int:
    doc = fitz.open(pdf)
    page_no = start_index
    for page in doc:
        rect = page.rect
        scale = min(1920 / rect.width, 1080 / rect.height) * 2
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        raw = output_dir / f"{video_id}_raw_{page_no:03d}.png"
        pix.save(raw)
        out = output_dir / f"{video_id}_slide_{page_no:03d}.png"
        vf = (
            "scale=1920:1080:force_original_aspect_ratio=decrease,"
            "pad=1920:1080:(ow-iw)/2:(oh-ih)/2:white,"
            "setsar=1"
        )
        run(["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", str(raw), "-vf", vf, str(out)])
        raw.unlink()
        page_no += 1
    return page_no


def render_image(image: Path, output_dir: Path, slide_index: int, video_id: str) -> int:
    out = output_dir / f"{video_id}_slide_{slide_index:03d}.png"
    vf = (
        "scale=1920:1080:force_original_aspect_ratio=decrease,"
        "pad=1920:1080:(ow-iw)/2:(oh-ih)/2:white,"
        "setsar=1"
    )
    run(["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", str(image), "-vf", vf, str(out)])
    return slide_index + 1


def render_slide_source(source: Path, output_dir: Path, start_index: int, video_id: str) -> int:
    if source.suffix.lower() == ".pdf":
        return render_pdf(source, output_dir, start_index, video_id)
    return render_image(source, output_dir, start_index, video_id)


def audio_duration(path: Path) -> float:
    value = run([
        "ffprobe", "-v", "error", "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1", str(path)
    ])
    return float(value)


def assign_timings(segments: list[Segment], total_duration: float) -> None:
    intro_pad = 0.4
    usable = max(total_duration - intro_pad, 1)
    total_chars = sum(segment.chars for segment in segments)
    cursor = 0.0
    for index, segment in enumerate(segments):
        segment.start = cursor
        if index == len(segments) - 1:
            segment.end = total_duration
        else:
            raw = usable * (segment.chars / total_chars)
            segment.end = min(total_duration, cursor + max(raw, 6.0))
        cursor = segment.end


def find_segment_audio(config: ProjectConfig, voice: str, slide: int) -> Path:
    filename = f"{config.video_id}_tts_{voice}_segment_{slide:03d}_{config.date}.wav"
    for folder in ["tts_final", "tts_raw"]:
        candidate = config.asset_dir / "audio" / folder / filename
        if candidate.exists():
            return candidate
    raise FileNotFoundError(f"Missing segment audio: {filename}")


def assign_segment_timings(segments: list[Segment], audio_paths: dict[int, Path], tail_hold: float) -> None:
    cursor = 0.0
    for segment in segments:
        path = audio_paths[segment.slide]
        segment.pre_read_silence = 3.0 if segment.slide == 1 else 2.0
        segment.audio_duration = audio_duration(path)
        segment.tail_hold = tail_hold
        segment.audio_path = str(path)
        segment.start = cursor
        segment.audio_start = cursor + segment.pre_read_silence
        segment.audio_end = segment.audio_start + segment.audio_duration
        segment.end = segment.audio_end + segment.tail_hold
        cursor = segment.end


def make_silence(build_dir: Path, duration: float) -> Path:
    label = f"{duration:.3f}".replace(".", "p")
    path = build_dir / f"silence_{label}.wav"
    if not path.exists():
        run([
            "ffmpeg", "-y", "-hide_banner", "-loglevel", "error",
            "-f", "lavfi", "-i", "anullsrc=r=24000:cl=mono",
            "-t", f"{duration:.3f}", "-c:a", "pcm_s16le", str(path),
        ])
    return path


def write_segment_audio(config: ProjectConfig, segments: list[Segment], output: Path) -> None:
    build_dir = config.build_dir
    concat_path = build_dir / f"{config.video_id}_tts_segment_sync_concat_{config.date}.txt"
    with concat_path.open("w", encoding="utf-8") as f:
        for segment in segments:
            for path in [
                make_silence(build_dir, segment.pre_read_silence),
                Path(segment.audio_path),
                make_silence(build_dir, segment.tail_hold),
            ]:
                f.write(f"file '{path}'\n")
    run([
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error",
        "-f", "concat", "-safe", "0", "-i", str(concat_path),
        "-c", "copy", str(output),
    ])


def write_slide_sync_csv(segments: list[Segment], path: Path) -> None:
    fields = [
        "slide", "title", "start_seconds", "end_seconds", "duration_seconds",
        "pre_read_silence_seconds", "audio_start_seconds", "audio_end_seconds",
        "audio_duration_seconds", "tail_hold_seconds", "audio_path",
    ]
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for segment in segments:
            writer.writerow({
                "slide": segment.slide,
                "title": segment.title,
                "start_seconds": f"{segment.start:.3f}",
                "end_seconds": f"{segment.end:.3f}",
                "duration_seconds": f"{segment.end - segment.start:.3f}",
                "pre_read_silence_seconds": f"{segment.pre_read_silence:.3f}",
                "audio_start_seconds": f"{segment.audio_start:.3f}",
                "audio_end_seconds": f"{segment.audio_end:.3f}",
                "audio_duration_seconds": f"{segment.audio_duration:.3f}",
                "tail_hold_seconds": f"{segment.tail_hold:.3f}",
                "audio_path": segment.audio_path,
            })


def segment_sync_json(segments: list[Segment]) -> list[dict]:
    return [{
        "slide": s.slide,
        "title": s.title,
        "start_seconds": round(s.start, 3),
        "end_seconds": round(s.end, 3),
        "duration_seconds": round(s.end - s.start, 3),
        "pre_read_silence_seconds": round(s.pre_read_silence, 3),
        "audio_start_seconds": round(s.audio_start, 3),
        "audio_end_seconds": round(s.audio_end, 3),
        "audio_duration_seconds": round(s.audio_duration, 3),
        "tail_hold_seconds": round(s.tail_hold, 3),
        "audio_path": s.audio_path,
    } for s in segments]


def make_telops(segments: list[Segment]) -> list[dict]:
    text_by_slide = {
        1: "地籍調査は、土地の最新記録を作り直す事業です",
        2: "古い記録のズレを、正確な座標で直します",
        3: "土地にも「戸籍」にあたる記録があります",
        4: "境界明確化・無料測量・災害復旧・公共事業効率化",
        5: "地権者の作業は、大きく4ステップです",
        6: "書類提出と、道路・水路境界の確認をお願いします",
        7: "個人同士の境界は、地権者の皆様で確認します",
        8: "合意できない場合、筆界未定地になるリスクがあります",
        9: "成果閲覧で、境界と面積を最終確認します",
        10: "立入・杭管理・固定資産税への影響をご確認ください",
        11: "現地に行けない場合は、代理人選任届を使えます",
        12: "一生に一度の機会として、ご協力をお願いします",
    }
    telops = []
    for segment in segments:
        base = segment.audio_start if segment.audio_start else segment.start
        limit = segment.audio_end if segment.audio_end else segment.end
        start = base + min(1.2, max((limit - base) * 0.12, 0.6))
        end = min(limit - 0.4, start + 7.0)
        if end > start + 2.0:
            telops.append({
                "slide": segment.slide,
                "start_seconds": round(start, 2),
                "end_seconds": round(end, 2),
                "text": text_by_slide.get(segment.slide, segment.title),
                "position": "bottom",
            })
    return telops


def write_ass(telops: list[dict], ass_path: Path) -> None:
    def ts(seconds: float) -> str:
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        s = seconds % 60
        return f"{h}:{m:02d}:{s:05.2f}"

    header = """[Script Info]
ScriptType: v4.00+
PlayResX: 1920
PlayResY: 1080

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Default, IPAexMincho, 54, &H00FFFFFF, &H000000FF, &H80222222, &H99000000, 1, 0, 0, 0, 100, 100, 0, 0, 3, 2, 0, 2, 130, 130, 70, 1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"""
    lines = [header]
    for telop in telops:
        text = str(telop["text"]).replace("\n", "\\N")
        lines.append(f"Dialogue: 0,{ts(telop['start_seconds'])},{ts(telop['end_seconds'])},Default,,0,0,0,,{text}\n")
    ass_path.write_text("".join(lines), encoding="utf-8")


def build_video(
    config: ProjectConfig,
    segments: list[Segment],
    audio: Path,
    final_mp4: Path,
    ass_path: Path | None = None,
) -> None:
    slide_dir = config.slide_export_dir
    build_dir = config.build_dir
    concat_list = build_dir / f"{config.video_id}_concat.txt"
    with concat_list.open("w", encoding="utf-8") as f:
        for segment in segments:
            duration = max(segment.end - segment.start, 1.0)
            slide = slide_dir / f"{config.video_id}_slide_{segment.slide:03d}.png"
            f.write(f"file '{slide}'\n")
            f.write(f"duration {duration:.3f}\n")
        last = slide_dir / f"{config.video_id}_slide_{segments[-1].slide:03d}.png"
        f.write(f"file '{last}'\n")

    silent_video = build_dir / f"{config.video_id}_silent_{config.date}.mp4"
    run([
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error",
        "-f", "concat", "-safe", "0", "-i", str(concat_list),
        "-vsync", "vfr", "-pix_fmt", "yuv420p", str(silent_video),
    ])
    cmd = [
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error",
        "-i", str(silent_video), "-i", str(audio),
    ]
    if ass_path is not None:
        cmd.extend(["-vf", f"subtitles={ass_path}:fontsdir=/usr/share/fonts/truetype"])
    cmd.extend([
        "-c:v", "libx264", "-preset", "medium", "-crf", "20",
        "-c:a", "aac", "-b:a", "160k", "-shortest", str(final_mp4),
    ])
    run(cmd)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--prepare", action="store_true")
    parser.add_argument("--assemble", action="store_true")
    parser.add_argument("--assemble-segments", action="store_true")
    parser.add_argument("--date", default=DEFAULT_DATE)
    parser.add_argument("--video-id", default=DEFAULT_VIDEO_ID)
    parser.add_argument("--asset-dir", type=Path)
    parser.add_argument("--source-dir", type=Path)
    parser.add_argument("--audio", type=Path)
    parser.add_argument("--voice", default="sulafat")
    parser.add_argument("--tail-hold", type=float, default=0.5)
    parser.add_argument("--draft-output", type=Path)
    parser.add_argument("--final-output", type=Path)
    args = parser.parse_args()

    config = make_config(args)
    ensure_layout(config.asset_dir)
    slide_sources, narration_docx, telop_docx = source_paths(config)
    narration_text = extract_docx_text(narration_docx)
    segments = parse_segments(narration_text)

    if args.prepare:
        tts_text = normalize_tts_text(segments)
        (config.manifest_dir / f"{config.video_id}_tts_input_{config.date}.txt").write_text(tts_text, encoding="utf-8")
        (config.manifest_dir / f"{config.video_id}_narration_extracted_{config.date}.txt").write_text(narration_text, encoding="utf-8")
        if telop_docx is not None:
            (config.manifest_dir / f"{config.video_id}_telop_source_{config.date}.txt").write_text(extract_docx_text(telop_docx), encoding="utf-8")
        slide_dir = config.slide_export_dir
        next_index = 1
        for source in slide_sources:
            next_index = render_slide_source(source, slide_dir, next_index, config.video_id)
        manifest = {
            "video_id": config.video_id,
            "status": "assets_ready",
            "date": config.date,
            "asset_dir": str(config.asset_dir),
            "source_dir": str(config.source_dir),
            "slide_count": next_index - 1,
            "slide_sources": [str(source) for source in slide_sources],
            "narration_docx": str(narration_docx),
            "telop_docx": str(telop_docx) if telop_docx is not None else None,
            "telops_enabled": telop_docx is not None,
            "segments": [{"slide": s.slide, "title": s.title, "chars": s.chars} for s in segments],
        }
        (config.manifest_dir / "production_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        return 0

    if args.assemble:
        if not args.audio:
            raise SystemExit("--audio is required with --assemble")
        audio = args.audio.resolve()
        target_audio = config.asset_dir / "audio" / "tts_final" / f"{config.video_id}_tts_{args.voice}_{config.date}{audio.suffix}"
        if audio != target_audio:
            shutil.copy2(audio, target_audio)
        duration = audio_duration(target_audio)
        assign_timings(segments, duration)
        telops = make_telops(segments) if telop_docx is not None else []
        manifest_dir = config.manifest_dir
        slide_sync = [{
            "slide": s.slide,
            "title": s.title,
            "start_seconds": round(s.start, 2),
            "end_seconds": round(s.end, 2),
        } for s in segments]
        (manifest_dir / f"{config.video_id}_slide_sync_{config.date}.json").write_text(json.dumps(slide_sync, ensure_ascii=False, indent=2), encoding="utf-8")
        (manifest_dir / f"{config.video_id}_telop_sync_{config.date}.json").write_text(json.dumps(telops, ensure_ascii=False, indent=2), encoding="utf-8")
        ass_path = None
        if telops:
            ass_path = manifest_dir / f"{config.video_id}_telops_{config.date}.ass"
            write_ass(telops, ass_path)
        final_dir = Path("/mnt/f/udemy_videos")
        final_dir.mkdir(parents=True, exist_ok=True)
        final_mp4 = args.final_output or final_dir / f"{config.video_id}_{config.date}.mp4"
        final_mp4.parent.mkdir(parents=True, exist_ok=True)
        build_video(config, segments, target_audio, final_mp4, ass_path)
        final_duration = audio_duration(final_mp4)
        ref = (
            f"# Final Video Reference\n\n"
            f"- video_id: {config.video_id}\n"
            f"- final_mp4: {final_mp4}\n"
            f"- duration_seconds: {final_duration:.2f}\n"
            f"- confirmed_date: {config.date}\n"
            f"- slides: {slides_source(config)}\n"
            f"- narration: {narration_docx}\n"
            f"- telops: {manifest_dir / f'{config.video_id}_telop_sync_{config.date}.json'}\n"
        )
        (config.asset_dir / "final_refs" / f"{config.video_id}_{config.date}.md").write_text(ref, encoding="utf-8")
        production_md = (
            "# Production Manifest\n\n"
            f"- video_id: {config.video_id}\n"
            "- status: final_exported\n"
            "- 対象: 地籍調査 地元説明会動画\n"
            f"- 参照スライド素材: {len(slide_sources)}ファイル、全{len(segments)}セグメント\n"
            f"- TTS本文: {manifest_dir / f'{config.video_id}_tts_input_{config.date}.txt'}\n"
            f"- 採用音声: {target_audio}\n"
            "- 採用BGM/イントロ/アウトロ: なし\n"
            f"- テロップ: {'あり' if telops else 'なし'}\n"
            f"- スライド書き出し物: {config.slide_export_dir}\n"
            f"- 同期マニフェスト: {manifest_dir / f'{config.video_id}_slide_sync_{config.date}.json'}\n"
            f"- 最終mp4保存先: {final_mp4}\n"
            f"- 尺: {final_duration:.2f}秒\n"
            f"- 確認日: {config.date}\n"
            "- 未完了項目: 目視・聴感による最終レビュー\n"
        )
        (manifest_dir / "production_manifest.md").write_text(production_md, encoding="utf-8")
        print(final_mp4)
        return 0

    if args.assemble_segments:
        if args.tail_hold < 0:
            raise SystemExit("--tail-hold must be zero or greater")
        manifest_dir = config.manifest_dir
        build_dir = config.build_dir
        audio_paths = {segment.slide: find_segment_audio(config, args.voice, segment.slide) for segment in segments}
        assign_segment_timings(segments, audio_paths, args.tail_hold)
        review_audio = config.asset_dir / "audio" / "tts_final" / f"{config.video_id}_tts_{args.voice}_segment_sync_review_{config.date}.wav"
        write_segment_audio(config, segments, review_audio)

        telops = make_telops(segments) if telop_docx is not None else []
        sync_csv = manifest_dir / f"{config.video_id}_slide_sync_review_{config.date}.csv"
        sync_json = manifest_dir / f"{config.video_id}_slide_sync_review_{config.date}.json"
        write_slide_sync_csv(segments, sync_csv)
        sync_json.write_text(json.dumps(segment_sync_json(segments), ensure_ascii=False, indent=2), encoding="utf-8")
        telop_json = manifest_dir / f"{config.video_id}_telop_sync_review_{config.date}.json"
        telop_json.write_text(json.dumps(telops, ensure_ascii=False, indent=2), encoding="utf-8")
        ass_path = None
        if telops:
            ass_path = manifest_dir / f"{config.video_id}_telops_{config.date}.ass"
            write_ass(telops, ass_path)

        draft_mp4 = args.draft_output or build_dir / f"{config.video_id}_review_{config.date}.mp4"
        build_video(config, segments, review_audio, draft_mp4, ass_path)

        final_mp4 = args.final_output or Path("/mnt/f/udemy_videos") / f"{config.video_id}_{config.date}.mp4"
        final_mp4.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(draft_mp4, final_mp4)
        final_duration = audio_duration(final_mp4)
        audio_total = audio_duration(review_audio)
        ref = (
            f"# Final Video Reference\n\n"
            f"- video_id: {config.video_id}\n"
            f"- final_mp4: {final_mp4}\n"
            f"- draft_mp4: {draft_mp4}\n"
            f"- duration_seconds: {final_duration:.2f}\n"
            f"- sync_method: per-slide actual wav duration + pre-read silence + tail hold\n"
            f"- sync_csv: {sync_csv}\n"
            f"- sync_json: {sync_json}\n"
            f"- tts_segment_audio: {review_audio}\n"
            f"- confirmed_date: {config.date}\n"
            f"- slides: {slides_source(config)}\n"
            f"- narration: {narration_docx}\n"
            f"- telops: {telop_json}\n"
        )
        (config.asset_dir / "final_refs" / f"{config.video_id}_{config.date}.md").write_text(ref, encoding="utf-8")
        production_md = (
            "# Production Manifest\n\n"
            f"- video_id: {config.video_id}\n"
            "- status: final_exported_v2_pending_human_review\n"
            "- 対象: 地籍調査 地元説明会動画\n"
            f"- 参照スライド素材: {len(slide_sources)}ファイル、全{len(segments)}セグメント\n"
            f"- TTS本文: {manifest_dir / f'{config.video_id}_tts_input_{config.date}.txt'}\n"
            f"- 採用音声: {review_audio}\n"
            "- 採用BGM/イントロ/アウトロ: なし\n"
            f"- テロップ: {'あり' if telops else 'なし'}\n"
            f"- スライド書き出し物: {config.slide_export_dir}\n"
            f"- 同期CSV: {sync_csv}\n"
            f"- 同期JSON: {sync_json}\n"
            f"- 仮動画保存先: {draft_mp4}\n"
            f"- 最終mp4保存先: {final_mp4}\n"
            f"- 尺: {final_duration:.2f}秒\n"
            f"- 音声尺: {audio_total:.2f}秒\n"
            "- 同期方式: スライド別wav実尺 + 読み出し前無音 + 末尾保持\n"
            "- 読み出し前無音: スライド1は3.0秒、スライド2以降は2.0秒\n"
            f"- 末尾保持: {args.tail_hold:.1f}秒\n"
            "- 確認工程:\n"
            "  - 資料確認: ページ数、順番、原稿対応、欠落、不一致を確認\n"
            "  - スライド確認: PDF書き出しPNGを目視確認\n"
            "  - TTS本文確認: 読み原稿をスライド単位で確認\n"
            "  - 音声確認: 各スライドwavを聴感確認し、承認済みだけtts_finalに置く\n"
            "  - 同期確認: 実尺同期CSVで切替、無音、読み出し開始位置を確認\n"
            "  - 仮動画確認: スライド切替、音声、テロップ被りを確認\n"
            "  - 最終書き出し: 承認後のみ最終mp4を保存\n"
            f"- 確認日: {config.date}\n"
            "- 未完了項目: 目視・聴感による最終レビュー\n"
        )
        (manifest_dir / "production_manifest.md").write_text(production_md, encoding="utf-8")
        print(final_mp4)
        return 0

    raise SystemExit("Use --prepare, --assemble, or --assemble-segments")


if __name__ == "__main__":
    raise SystemExit(main())
