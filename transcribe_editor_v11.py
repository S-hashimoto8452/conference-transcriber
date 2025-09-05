# transcribe_editor_v8.py
# -------------------------------------------------------------
# 機能:
# 1) 音声/動画アップロード → Whisper系(faster-whisper)で文字起こし
# 2) 出力選択: 逐語(タイムスタンプ) / 直訳（日本語化のみ）/ 議事録 / 要旨 / 記事 / ガイドライン解説
# 3) 目的選択: 学会発表 / ガイドライン解説 / ディスカッション（LLM整形に反映）
# 4) 動画オプション: スライドOCR(キーフレーム抽出 + OCR) 併用の可否
# 5) 生成AIは任意。APIキー未入力でもヒューリスティック整形で動作
# 6) TXT/DOCXでダウンロード可能
# 7) 生成AIで整形（日本語）→ 上から「原文英語／直訳／整形結果」の三段表示
# -------------------------------------------------------------

import os
import io
import time
import glob
import shutil
import subprocess
import mimetypes
from datetime import timedelta
import re
from typing import List, Tuple, Dict, Any

import streamlit as st
from pydub import AudioSegment
from pydub.utils import which
from docx import Document
from docx.shared import Pt

from faster_whisper import WhisperModel

# 任意: OpenAI
try:
    import openai as openai_mod  # pip install openai
except Exception:
    openai_mod = None

# 任意: スライドOCR（easyocr）
try:
    import easyocr  # pip install easyocr
except Exception:
    easyocr = None

import math
from pathlib import Path

# 🔐 パスワード認証用関数
def require_password():
    """共通パスワードで簡易ログイン。通過しない限りアプリ本体を表示しない。"""
    if st.session_state.get("auth_ok", False):
        # ログアウトボタン（任意）
        with st.sidebar:
            if st.button("🔒 ログアウト"):
                st.session_state.clear()
                st.experimental_rerun()
        return

    with st.sidebar:
        st.header("社内ログイン")
        pw = st.text_input("パスワード", type="password", help="社内共通パスワードを入力")
        submitted = st.button("ログイン")

    correct = st.secrets.get("APP_PASSWORD", None)

    if submitted:
        if correct is None:
            st.error("APP_PASSWORD が未設定です。.streamlit/secrets.toml を確認してください。")
        elif pw == correct:
            st.session_state["auth_ok"] = True
            st.success("ログイン成功")
            st.experimental_rerun()
        else:
            st.error("パスワードが違います。")

    if not st.session_state.get("auth_ok", False):
        st.stop()

# ---------------------- ユーティリティ ----------------------

def split_text_by_chars(text: str, chunk_size: int = 6000, overlap: int = 300) -> list[str]:
    text = text.strip()
    if len(text) <= chunk_size:
        return [text]
    chunks = []
    start = 0
    while start < len(text):
        end = min(len(text), start + chunk_size)
        cut = end
        for p in ("。", "！", "？", "\n"):
            idx = text.rfind(p, start, end)
            if idx != -1 and idx > start + 1000:
                cut = idx + 1
                break
        chunks.append(text[start:cut].strip())
        if cut >= len(text):
            break
        start = max(cut - overlap, 0)
    return [c for c in chunks if c]


def strip_timestamps(text: str) -> str:
    pattern = re.compile(
        r"^\s*\[\d{2}:\d{2}:\d{2}(?:\.\d{3})?\s*(?:→|->|-|－|—)\s*\d{2}:\d{2}:\d{2}(?:\.\d{3})?\]\s*",
        re.MULTILINE,
    )
    return pattern.sub("", text).strip()


# ====================== FFmpeg/ffprobe の明示設定 ======================
PROJECT_DIR = Path(__file__).parent
FFBIN_CANDIDATES = [
    PROJECT_DIR / "ffmpeg-7.0.2-essentials_build" / "bin",
    Path(r"C:\\Users\\s-has\\Desktop\\動画音声原稿作成082025\\ffmpeg-7.0.2-essentials_build\\bin"),
    Path(r"C:\\Users\\s-has\\Desktop\\ffmpeg-7.0.2-essentials_build\\bin"),
]

FFMPEG_EXE = None
FFPROBE_EXE = None
for _bin in FFBIN_CANDIDATES:
    ff = _bin / "ffmpeg.exe"
    fp = _bin / "ffprobe.exe"
    if ff.exists():
        FFMPEG_EXE, FFPROBE_EXE = ff, (fp if fp.exists() else None)
        os.environ["PATH"] = str(_bin) + os.pathsep + os.environ.get("PATH", "")
        os.environ["FFMPEG_BINARY"] = str(ff)
        os.environ["IMAGEIO_FFMPEG_EXE"] = str(ff)
        AudioSegment.converter = str(ff)
        AudioSegment.ffmpeg = str(ff)
        if FFPROBE_EXE:
            AudioSegment.ffprobe = str(FFPROBE_EXE)
        break
else:
    ffmpeg_found = which("ffmpeg")
    ffprobe_found = which("ffprobe")
    if ffmpeg_found:
        FFMPEG_EXE = Path(ffmpeg_found)
        AudioSegment.converter = ffmpeg_found
        AudioSegment.ffmpeg = ffmpeg_found
    if ffprobe_found:
        FFPROBE_EXE = Path(ffprobe_found)
        AudioSegment.ffprobe = ffprobe_found
# ======================================================================


def save_uploaded_file_to_temp(uploaded_file) -> str:
    suffix = os.path.splitext(uploaded_file.name)[1]
    tmp_path = os.path.join(st.session_state["workdir"], f"upload_{int(time.time())}{suffix}")
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return tmp_path


def ensure_wav(input_path: str) -> str:
    try:
        audio = AudioSegment.from_file(input_path)
    except Exception as e:
        st.error(
            "音声/動画の読み込みに失敗しました。\n"
            "ffmpeg/ffprobe が実行可能か、PATH/コード設定が正しいか確認してください。\n\n"
            f"詳細: {e}"
        )
        st.stop()
    audio = audio.set_channels(1).set_frame_rate(16000)
    wav_path = os.path.splitext(input_path)[0] + "_16k.wav"
    audio.export(wav_path, format="wav")
    return wav_path


def format_timestamp(seconds: float) -> str:
    td = timedelta(seconds=float(seconds))
    total_seconds = int(td.total_seconds())
    ms = int((td.total_seconds() - total_seconds) * 1000)
    return f"{total_seconds//3600:02d}:{(total_seconds%3600)//60:02d}:{total_seconds%60:02d}.{ms:03d}"


def fmt_ts(x: float) -> str:
    return format_timestamp(x) if math.isfinite(x) else "…"


# ---------------------- スライドと発話の対応付け ----------------------

def group_segments_by_slides(
    segments: List[Tuple[str, float, float]],
    slide_change_times: List[float]
) -> List[Dict[str, Any]]:
    last_end = max((e for _, _, e in segments), default=0.0)
    bounds = [0.0] + [t for t in slide_change_times if t < last_end] + [last_end]
    grouped = []
    for i in range(len(bounds)-1):
        start, end = bounds[i], bounds[i+1]
        bucket = []
        for t, s, e in segments:
            if e > start and s < end:
                bucket.append((t, max(s, start), min(e, end)))
        grouped.append({"index": i+1, "start": start, "end": end, "segments": bucket})
    return grouped


# ---------------------- 文字起こし本体 ----------------------

def transcribe_faster_whisper(
    wav_path: str,
    model_size: str = "small",
    language: str | None = "auto",
    compute_type: str = "auto",
    beam_size: int = 5,
) -> tuple[list[tuple[str, float, float]], str | None]:
    lang_arg = None if (language is None or str(language).lower() == "auto") else language
    model = WhisperModel(model_size, device="auto", compute_type=compute_type)
    segments_gen, info = model.transcribe(
        wav_path,
        language=lang_arg,
        beam_size=beam_size,
        vad_filter=True,
        vad_parameters=dict(min_silence_duration_ms=500),
    )
    results: list[tuple[str, float, float]] = []
    for seg in segments_gen:
        results.append((seg.text.strip(), seg.start, seg.end))
    detected = getattr(info, "language", None)
    return results, detected


# ---------------------- スライド抽出 & OCR ----------------------
def extract_slide_keyframes_with_times(video_path: str, out_dir: str, scene_thr: float=0.35) -> tuple[list[str], list[float]]:
    os.makedirs(out_dir, exist_ok=True)
    for p in glob.glob(os.path.join(out_dir, "*.jpg")):
        try:
            os.remove(p)
        except:
            pass

    ff_cmd = str(FFMPEG_EXE) if FFMPEG_EXE else (shutil.which("ffmpeg") or "ffmpeg")

    # 1) シーン変化抽出
    cmd = [
        ff_cmd, "-y", "-i", video_path,
        "-vf", f"select='gt(scene,{scene_thr})',showinfo",
        "-vsync", "vfr",
        os.path.join(out_dir, "%04d.jpg"),
    ]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="ignore")
    stderr = proc.stderr or ""

    times = []
    for m in re.finditer(r"pts_time:([0-9]+\.[0-9]+)", stderr):
        try:
            times.append(float(m.group(1)))
        except:
            pass

    image_paths = sorted(glob.glob(os.path.join(out_dir, "*.jpg")))
    n = min(len(image_paths), len(times))
    if n > 0:
        return image_paths[:n], times[:n]

    # 2) フォールバック：3秒間隔で抽出
    for p in glob.glob(os.path.join(out_dir, "*.jpg")):
        try: os.remove(p)
        except: pass
    cmd_fb = [
        ff_cmd, "-y", "-i", video_path,
        "-vf", "fps=1/3,showinfo",
        "-vsync", "vfr",
        os.path.join(out_dir, "%04d.jpg"),
    ]
    proc_fb = subprocess.run(cmd_fb, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="ignore")

    image_paths = sorted(glob.glob(os.path.join(out_dir, "*.jpg")))
    if image_paths:
        approx_times = [i * 3.0 for i in range(len(image_paths))]
        return image_paths, approx_times

    # 3) それでも0なら先頭1枚だけ確保
    one_path = os.path.join(out_dir, "0001.jpg")
    subprocess.run([ff_cmd, "-y", "-ss", "00:00:01", "-i", video_path, "-vframes", "1", one_path],
                   stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="ignore")
    image_paths = sorted(glob.glob(os.path.join(out_dir, "*.jpg")))
    if image_paths:
        return image_paths, [0.0]

    return [], []

# 先頭付近（import の近く）に必要なら追加
import numpy as np
import cv2
from PIL import Image

def _to_cv2_bgr(image_like):
    try:
        if isinstance(image_like, (bytes, bytearray)):
            arr = np.frombuffer(image_like, np.uint8)
            img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
            return img

        if isinstance(image_like, str):
            img = cv2.imread(image_like, cv2.IMREAD_COLOR)
            if img is None:  # ← 日本語パス等で失敗するケースに対応
                try:
                    pil = Image.open(image_like).convert("RGB")
                    return cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)
                except Exception:
                    return None
            return img

        if isinstance(image_like, Image.Image):
            return cv2.cvtColor(np.array(image_like), cv2.COLOR_RGB2BGR)

        if isinstance(image_like, np.ndarray):
            if image_like.ndim == 2:
                return cv2.cvtColor(image_like, cv2.COLOR_GRAY2BGR)
            return image_like
    except Exception:
        return None
    return None

def _get_reader():
    """EasyOCR Reader を（あれば）キャッシュして使い回し。"""
    try:
        import streamlit as st
        @st.cache_resource(show_spinner=False)
        def _cached_reader():
            return easyocr.Reader(['ja', 'en'], gpu=False)
        return _cached_reader()
    except Exception:
        return easyocr.Reader(['ja', 'en'], gpu=False)

def ocr_slides(image_paths: list) -> list[dict]:
    """
    image_paths: 画像パス/bytes/PIL/ndarray が混在していてもOK
    return: [{"index": i, "path": 元の参照, "text": 認識文字列}, ...]
    """
    if not image_paths:
        return []

    if easyocr is None:
        # EasyOCR が無い環境でも落ちないように空文字で返す
        return [{"index": i+1, "path": p, "text": ""} for i, p in enumerate(image_paths)]

    reader = _get_reader()
    results = []
    valid_found = False

    for idx, src in enumerate(image_paths, start=1):
        img = _to_cv2_bgr(src)
        if img is None or getattr(img, "size", 0) == 0:
            # 画像化できなかった場合でも空文字でレコードを返す（落とさない）
            results.append({"index": idx, "path": src, "text": ""})
            continue

        valid_found = True
        try:
            # detail=0 だとテキストのみ返る。detail=1 でも可。
            lines = reader.readtext(img, detail=0)
            text = "\n".join(lines) if lines else ""
            results.append({"index": idx, "path": src, "text": text})
        except Exception:
            # 1枚NGでも全体は継続
            results.append({"index": idx, "path": src, "text": ""})

    # 1枚も有効画像が作れなかったときはユーザに通知（Streamlitがある場合）
    if not valid_found:
        try:
            import streamlit as st
            st.error("OCR用の画像を正しく読み込めませんでした（パス・形式・抽出処理をご確認ください）。")
        except Exception:
            pass

    return results

# ---------------------- 整形(生成AIなし) ----------------------

def to_verbatim_with_timestamps(segments: List[Tuple[str, float, float]]) -> str:
    lines: List[str] = []
    for t, s, e in segments:
        start_disp = format_timestamp(s) if math.isfinite(s) else "…"
        end_disp   = format_timestamp(e) if math.isfinite(e) else "…"
        lines.append(f"[{start_disp} → {end_disp}] {t}")
    return "\n".join(lines)


def heuristic_minutes(segments: List[Tuple[str, float, float]]) -> str:
    block, blocks, char_limit = [], [], 300
    for t, s, e in segments:
        if sum(len(x[0]) for x in block) + len(t) > char_limit and block:
            blocks.append(block); block = []
        block.append((t, s, e))
    if block: blocks.append(block)
    out = ["【議事録（自動整形・要点）】\n"]
    for i, b in enumerate(blocks, 1):
        out.append(f"■ トピック{i}（{format_timestamp(b[0][1])}–{format_timestamp(b[-1][2])}）")
        for t, _, _ in b: out.append(f"・{t}")
        out.append("")
    return "\n".join(out).strip()


def heuristic_abstract(segments: List[Tuple[str, float, float]]) -> str:
    text = " ".join(t for t, _, _ in segments)
    sentences = [s.strip() for s in text.replace("。", "。\n").splitlines() if s.strip()]
    return "【要旨（自動抽出）】\n" + "\n".join(sentences[:6])


def heuristic_article_academic(segments: List[Tuple[str, float, float]]) -> str:
    body = " ".join(t for t, _, _ in segments)
    lines = [
        "【学会報告記事（自動整形・AI不使用）】",
        "",
        "■ リード",
        "本講演では、演者が提示した主要ポイントを抜粋し、内容を簡潔に整理する。本文は自動整形のため、要点レベルの抜粋である。",
        "",
        "■ 背景・目的",
        "講演の背景、臨床上の意義、目的を本文から機械的に抽出・再構成。",
        "",
        "■ 方法・資料",
        "使用データ、対象、手法、評価指標などの記載を要点として抽出。",
        "",
        "■ 結果・所見",
        "本文から結果に相当する文を優先的に拾い上げ反映。",
        "",
        "■ 考察・結論",
        "臨床現場への示唆、限界、今後の展望を簡潔にまとめる。",
        "",
        "— 以下は逐語ベース本文（機械抽出） —",
        body,
    ]
    return "\n".join(lines)


def heuristic_guideline_commentary(slide_groups: List[Dict[str, Any]], ocr_notes: List[dict]) -> str:
    ocr_map = {o.get("index"): (o.get("text") or "").strip() for o in (ocr_notes or [])}
    lines = [
        "【ガイドライン解説（自動整形・AI不使用）】\n",
        "■ 背景",
        "・本解説は演者スライドとスピーチ内容を対応付けて再構成したもの。",
        "",
    ]
    for g in slide_groups:
        idx, ocr = g["index"], ocr_map.get(g["index"], "")
        lines.append(f"▼ Slide {idx}（{format_timestamp(g['start'])}–{fmt_ts(g['end'])}）")
        if ocr:
            title = ocr.splitlines()[0][:50]
            lines.append(f"【スライド要旨】{title}")
        for t, s, e in g["segments"][:6]:
            lines.append(f"・{t}")
        lines.append("")
    lines += ["■ 臨床への含意", "・本改訂により想定される診療上の影響点を要点化。", "", "■ 今後の課題", "・エビデンス強化が必要な論点、運用時の留意点。"]
    return "\n".join(lines).strip()


# ---------------------- 目的別プロンプト（記事化/要旨/議事録 用） ----------------------

PURPOSE_PROMPTS = {
    "学会発表": (
        "以下の素材（音声逐語と任意のスライドOCR要約）から、学会報告記事を作成してください。"
        "見出し（導入/背景/目的/方法/結果/考察/結語）を付け、固有名詞と数値は改変せず、"
        "誇張や創作は避けてください。専門読者向けに簡潔で正確に。"
    ),
    "ガイドライン解説": (
        "以下の素材から、日本語のガイドライン改訂解説記事を作成してください。"
        "背景/改訂ポイント/推奨度・エビデンス/臨床への影響/課題/今後、の順に一度だけ骨組みを提示してください。"
        "テキストが複数パートに分割される場合でも、見出し・導入の再掲はしないでください。"
        "既出内容の再掲を避け、新規情報のみ追記する形で連続性を保ってください。"
        "英語は正確に日本語化し、引用は要旨化して書き直してください。"
    ),
    "ディスカッション": (
        "以下の素材から、ディスカッション記事を作成してください。"
        "論点整理/賛否の主張/根拠/一致点と相違点/結論と今後の検討課題、の順で、中立・簡潔にまとめてください。"
        "冗長な口語表現は削除し、方言は標準語に直してください。"
    ),
}


# ---------------------- LLM 出力（記事化/要旨/議事録） ----------------------

def llm_rewrite(kind: str, text: str, api_key: str | None,
                purpose: str | None = None,
                source_lang: str | None = None,
                target_lang: str | None = "ja") -> str:
    """
    記事化・要旨・議事録・逐語（軽整形）用。
    ※ 直訳は使わず、別関数 llm_translate_only() を使うこと！
    """
    if openai_mod is None:
        return "[LLM未インストール] `pip install -U openai` を実行してください。"

    if not api_key:
        api_key = os.getenv("OPENAI_API_KEY", "")
    if not api_key:
        return "[APIキー未入力] サイドバーでAPIキーを入れるか、環境変数 OPENAI_API_KEY を設定してください。"

    sys_prompt = (
        "あなたは医学・医療系の日本語編集者です。臨床・学術文脈に沿って、"
        "読みやすく事実関係を保ったまま整文します。数値や引用は改変しません。"
    )
    pre = PURPOSE_PROMPTS.get(purpose or "学会発表", "")

    if (target_lang or "ja").lower() == "ja":
        lang_policy = (
            "最終出力は必ず日本語で書いてください。音声/スライドが日本語でない場合は正確に日本語へ翻訳し、"
            "専門用語は適切な日本語訳を用い、固有名詞・数値・単位は保持してください。"
        )
        if source_lang and str(source_lang).lower() != "ja":
            lang_policy += "（入力は日本語以外と検出されたため翻訳が必要です）"
    else:
        lang_policy = f"最終出力は必ず {target_lang} で書いてください。固有名詞・数値・単位は保持してください。"

    user_prompt_map = {
        "verbatim": "逐語記録（軽微な句読点整形のみ、意味改変禁止）：\n\n" + text,
        "minutes":  "議事録（見出し＋箇条書き、時系列）：\n\n" + text,
        "abstract": "学会抄録（目的/方法/結果/結論、600-900字）：\n\n" + text,
        "article":  "記事化（導入/背景/目的/方法/結果/考察/結語）：\n\n" + text,
    }
    if kind not in user_prompt_map:
        kind = "article"

    prompt = (pre + "\n\n" + lang_policy + "\n\n" + user_prompt_map[kind]).strip()

    client = openai_mod.OpenAI(api_key=api_key) if hasattr(openai_mod, "OpenAI") else None
    try:
        if client:
            resp = client.chat.completions.create(
                model="gpt-4o-mini-2024-07-18",
                messages=[{"role": "system", "content": sys_prompt},
                          {"role": "user", "content": prompt}],
                temperature=0.1,
            )
            result = resp.choices[0].message.content
        else:
            openai_mod.api_key = api_key
            resp = openai_mod.ChatCompletion.create(
                model="gpt-4o-mini-2024-07-18",
                messages=[{"role": "system", "content": sys_prompt},
                          {"role": "user", "content": prompt}],
                temperature=0.1,
            )
            result = resp["choices"][0]["message"]["content"]
    except Exception as e:
        return f"[LLMエラー] {e}"

    if kind != "verbatim":
        result = "【AI整形】\n" + result
    return result


# ---------------------- LLM 直訳（翻訳専用・整形一切なし） ----------------------

def llm_translate_only(text: str, api_key: str | None,
                       source_lang: str | None = None,
                       target_lang: str = "ja") -> str:
    """
    逐語的な翻訳に特化。記事化プロンプトを一切付けない。
    - 要約・見出し・箇条書きの追加は禁止
    - 「【AI整形】」等のヘッダーも付けない
    """
    if openai_mod is None:
        return "[LLM未インストール] `pip install -U openai` を実行してください。"
    if not api_key:
        api_key = os.getenv("OPENAI_API_KEY", "")
    if not api_key:
        return "[APIキー未入力] サイドバーでAPIキーを入れるか、環境変数 OPENAI_API_KEY を設定してください。"

    sys_prompt = (
        "あなたは忠実な専門翻訳者です。以下のテキストを逐語的に日本語へ翻訳してください。"
        "要約・意訳・見出し付け・箇条書き化・体裁変更は行わないでください。"
        "段落や改行等の構造は可能な限り保持し、固有名詞・数値・単位は維持してください。"
    )
    if (target_lang or "ja").lower() != "ja":
        sys_prompt = sys_prompt.replace("日本語", target_lang)

    prompt = "【翻訳対象】\n" + text

    client = openai_mod.OpenAI(api_key=api_key) if hasattr(openai_mod, "OpenAI") else None
    try:
        if client:
            resp = client.chat.completions.create(
                model="gpt-4o-mini-2024-07-18",
                messages=[{"role": "system", "content": sys_prompt},
                          {"role": "user", "content": prompt}],
                temperature=0.0,
            )
            return resp.choices[0].message.content
        else:
            openai_mod.api_key = api_key
            resp = openai_mod.ChatCompletion.create(
                model="gpt-4o-mini-2024-07-18",
                messages=[{"role": "system", "content": sys_prompt},
                          {"role": "user", "content": prompt}],
                temperature=0.0,
            )
            return resp["choices"][0]["message"]["content"]
    except Exception as e:
        return f"[LLMエラー] {e}"

# ---------------------- LLM: 直訳日本語 → 記事調（重複以外を削らない） ----------------------
def llm_article_from_literal(literal_ja: str,
                             api_key: str | None,
                             purpose: str | None = "学会発表") -> str:
    """
    逐語直訳（日本語）を素材に、削りすぎを避けつつ記事調（常体）へ整文。
    - 重複以外は削らない（=意味の落ちを防ぐ）
    - 数値・試験名・薬剤名は保持
    - 語順やつなぎの調整は可（読みやすさ確保）
    """
    if openai_mod is None:
        return "[LLM未インストール] `pip install -U openai` を実行してください。"
    if not api_key:
        api_key = os.getenv("OPENAI_API_KEY", "")
    if not api_key:
        return "[APIキー未入力] サイドバーでAPIキーを入れるか、OPENAI_API_KEY を設定してください。"

    sys_prompt = (
        "あなたは医療・医学分野の編集者。入力は既に日本語へ逐語直訳された原稿。"
        "重複・言い換えの冗長だけを整理し、意味・事実は落とさず記事調（常体）に整える。"
        "【厳守】重複以外の削除禁止／数値・試験名・薬剤名・用量・単位は保持。"
        "見出しは『導入/背景/目的/方法/結果/考察/結語』の順で一度だけ。"
        "脚色・新情報の追加は禁止。"
        "文末は常体（〜だ／〜である）に統一し、です・ます調は禁止。"  # ← 追加
    )
    preface = {
        "学会発表": "学会報告の速報トーンで、専門読者向けに簡潔で正確に。",
        "ガイドライン解説": "解説記事の文体で、背景→要点→臨床的含意を明確に。",
        "ディスカッション": "論点を明確化しつつ中立に記述。"
    }.get(purpose or "学会発表", "専門読者向けに簡潔で正確に。")

    user_prompt = (
        f"{preface}\n\n"
        "【入力（逐語直訳・日本語）】\n"
        + literal_ja.strip()
        + "\n\n【出力仕様】\n"
          "- TCROSS NEWS 学会発表記事のフォーマットに整形すること。\n"
          "- タイトルは「対象/疾患・介入: 試験名」とする。\n"
          "- 第1段落は「△△試験より、□□ことが、国、所属、演者名により、学会名とセッション名で発表された。」という形で書く（Conclusionの冒頭文を反映）。\n"
          "- 第2段落は試験デザインを記載（試験名、登録期間、国・施設数、患者数、群割付け、割付け数）。\n"
          "- 第3段落は患者背景を詳細に記載（差がなければ平均値で、年齢・性別・併存症・薬剤処方率を含める）。\n"
          "- 第4段落は主要評価項目の結果を記載（追跡期間、イベント率、HR、95%CI、p値を保持）。\n"
          "- 第5段落以降にサブ解析結果があれば記載。\n"
          "- 最終段落は演者のラストネームから始め、「…と、まとめた。」で必ず締める。\n"
          "- 同時掲載があれば「尚、△△試験は○○誌に掲載された。」と加える。\n"
          "- 記事調（常体）。\n"
          "- 見出しは『導入/背景/目的/方法/結果/考察/結語』。\n"
          "- 冗長な重複は統合。その他の内容は残す（削りすぎ禁止）。\n"
          "- 数値・用語はそのまま保持。\n"
          "- 箇条書きではなく段落ごとにまとめ、論理的な流れを持たせる。\n"
          "- 試験背景→方法→患者背景→結果→解釈→制限→結論の流れを基本とする。\n"
          "- 全体のボリュームを落とさず、原文と同等の情報量を保つ。\n"
          "- 要約ではなく、構成整理と文体変換を主目的とする。\n"
          "- 医学系ニュース記事や学術誌速報記事にふさわしい文体を用いる。\n"
          "- 演者が提示した患者背景、手技特徴、薬物療法の詳細は必ず含める（数値・割合・レジメンを省略しない）。\n"
          "- 結果部分は逐語性を優先し、省略や要約は一切禁止する（統計値・HR・P値・イベント率などを必ず残す）。\n"
          "- 月数の表記は「か月」ではなく「ヶ月」を用いる（例：6ヶ月、12ヶ月）。\n"
          "- 英語スクリプトに含まれる「結果」の逐語内容は削らず、全て日本語に反映する。\n"
          "- 結果に関する逐語的な統計数値・発言内容（イベント率・ハザード比・p値・サブ解析など）は省略禁止。\n"
          "- 結果は逐語スクリプトの情報量を保持したまま記事調に整えること。\n"
          "- 出力は入力と同等の情報量を保持し、文字数はおおむね {int(target_chars*0.95)}〜{int(target_chars*1.05)} 文字（±5%）に収める。\n"
          "- 短縮・要約を禁止。段落整理・文体整形・フォーマット化のみ行う。\n"
          "- 結果・患者背景・手技・薬物療法・限界・考察の逐語情報は省略禁止（統計値・イベント率・HR・95%CI・p値・サブ解析を含む）。\n"
    )

    client = openai_mod.OpenAI(api_key=api_key) if hasattr(openai_mod, "OpenAI") else None
    try:
        if client:
            resp = client.chat.completions.create(
                model="gpt-4o-mini-2024-07-18",
                messages=[{"role": "system", "content": sys_prompt},
                          {"role": "user", "content": user_prompt}],
                temperature=0.15,
            )
            return "【AI整形（直訳→記事調）】\n" + resp.choices[0].message.content
        else:
            openai_mod.api_key = api_key
            resp = openai_mod.ChatCompletion.create(
                model="gpt-4o-mini-2024-07-18",
                messages=[{"role": "system", "content": sys_prompt},
                          {"role": "user", "content": user_prompt}],
                temperature=0.15,
            )
            return "【AI整形（直訳→記事調）】\n" + resp["choices"][0]["message"]["content"]
    except Exception as e:
        return f"[LLMエラー] {e}"

# ---------------------- DOCX出力 ----------------------

def make_docx(title: str, content: str) -> bytes:
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Yu Gothic'
    font.size = Pt(11)

    doc.add_heading(title or "出力", level=1)
    for line in content.splitlines():
        if line.strip() == "":
            doc.add_paragraph("")
        else:
            doc.add_paragraph(line)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ---------------------- Streamlit UI ----------------------

def main():
    st.set_page_config(page_title="InsighTCROSS® Smart Writer v11", layout="wide")
    # 🔐 最初に認証をかける
    require_password()
        
    if "workdir" not in st.session_state:
        st.session_state["workdir"] = os.path.abspath("./.work")
        os.makedirs(st.session_state["workdir"], exist_ok=True)

    st.title("InsighTCROSS® Smart Writer v11")
    if "transcript_text" not in st.session_state:
        st.session_state["transcript_text"] = ""
    if "generated_text" not in st.session_state:
        st.session_state["generated_text"] = ""
    st.write("音声/動画をアップロードして、逐語・直訳・議事録・要旨・記事に整形。動画はスライドOCR併用も可能。生成AIは任意です。")

    with st.sidebar:
        st.header("設定")
        file_type = st.radio("ファイルタイプ", ["自動判定", "音声", "動画"], index=0)
        use_slide_ocr = st.toggle("スライドOCRも併用（動画時）", value=False,
                                  help="スライドのキーフレームを抽出しOCRで文字も取り込みます")
        scene_sensitivity = st.slider("シーン変化感度", 0.10, 0.60, 0.35, 0.01)

        model_size = st.selectbox("Whisperモデル", ["tiny","base","small","medium"], index=2)
        compute_type = st.selectbox("Compute type", ["auto","int8","int8_float16","float16","float32"], index=0)
        language = st.text_input("言語コード", value="auto")

        output_lang_label = st.selectbox("出力言語", ["日本語 (JPN)", "English (EN)"], index=0)
        output_lang = "ja" if "JPN" in output_lang_label else "en"

        out_kind = st.selectbox(
            "出力タイプ",
            ["逐語(タイムスタンプ)", "直訳（日本語化のみ）", "議事録", "要旨", "記事", "ガイドライン解説"]
        )
        purpose = st.selectbox("記事化の目的", ["学会発表", "ガイドライン解説", "ディスカッション"], index=0)
        attach_verbatim = st.toggle("末尾に逐語原文を添付", value=False,
                                    help="原文言語の逐語テキストを末尾に付けます（通常はOFF推奨）")
        use_llm = st.toggle("生成AIで整形（任意）", value=False)
        api_key_input = ""
        if use_llm:
            api_key_input = st.text_input("OpenAI API Key", type="password",
                                          help="未入力なら環境変数 OPENAI_API_KEY を参照")
            if not api_key_input:
                st.warning("生成AIで整形がONですが APIキーが未入力です。AI整形は実行されず、元言語のまま出力されます。")

    uploaded = st.file_uploader(
        "音声/動画ファイルをアップロード (mp3, m4a, wav, mp4, mov など)",
        type=["mp3","m4a","wav","mp4","mov","mkv","aac","flac"]
    )

    if not uploaded:
        return

    st.info(f"受信: {uploaded.name} / {uploaded.size/1024:.1f} KB")
    temp_path = save_uploaded_file_to_temp(uploaded)
    guessed = (uploaded.type or mimetypes.guess_type(uploaded.name)[0] or "")
    is_video = (file_type == "動画") or (file_type == "自動判定" and guessed.startswith("video/"))

    with st.spinner("変換中（WAV 16kHz mono）..."):
        wav_path = ensure_wav(temp_path)
    segments, detected_lang = transcribe_faster_whisper(
        wav_path=wav_path,
        model_size=model_size,
        language="en",
        compute_type=compute_type,
    )
    st.success(f"文字起こし完了。セグメント数: {len(segments)} / 言語検出: {detected_lang}")

    # 逐語（タイムスタンプ付き）原稿
    verbatim_text = to_verbatim_with_timestamps(segments)
    st.session_state["transcript_text"] = verbatim_text

    st.subheader("✍️ 逐語テキスト（編集可）")
    st.session_state["transcript_text"] = st.text_area(
        "逐語（必要に応じて修正してください）",
        value=st.session_state["transcript_text"],
        height=300
    )

    # スライドOCR（任意）
    slide_groups, slide_notes, slide_digest = [], [], ""
    if is_video and use_slide_ocr:
        with st.spinner("スライド抽出（キーフレーム+時刻）→ OCR 中..."):
            frames, slide_times = extract_slide_keyframes_with_times(
                video_path=temp_path,
                out_dir=os.path.join(st.session_state["workdir"], "slides"),
                scene_thr=scene_sensitivity,
            )

            # === ここから追加の見える化デバッグ ===
            st.write(f"抽出フレーム枚数: {len(frames)} / 切替検出: {len(slide_times)}")
            if frames:
                st.write("先頭3枚のパス:", frames[:3])
                try:
                    st.image(frames[0], caption="スライド抽出プレビュー（先頭）", use_container_width=True)
                except Exception as e:
                    st.warning(f"プレビュー表示に失敗: {e}")
            else:
                st.warning("抽出された画像が0枚です。フォールバックが効いていない可能性があります。")
            # === ここまで追加 ===

            slide_notes = ocr_slides(frames)
            slide_groups = group_segments_by_slides(segments, slide_times)
            slide_digest = "\n\n".join(
                [f"[Slide {s['index']}]\n{s.get('text','')}" for s in slide_notes if s.get('text','').strip()]
            )
        st.success(f"スライド抽出: {len(slide_notes)} 枚 / 切替: {len(slide_times)} 点")

    edited_transcript = st.session_state["transcript_text"]
    cleaned_for_llm = strip_timestamps(edited_transcript)

    if out_kind == "ガイドライン解説" and slide_groups:
        chunks = []
        for g in slide_groups:
            idx = g["index"]
            ocr_text = ""
            for s in (slide_notes or []):
                if s.get("index") == idx:
                    ocr_text = (s.get("text") or "").strip()
                    break
            speech_text = "\n".join([t for (t, _, _) in g["segments"]])
            chunks.append(
                f"[Slide {idx} {format_timestamp(g['start'])}–{fmt_ts(g['end'])}]\n"
                f"<OCR>\n{ocr_text}\n</OCR>\n<SPEECH>\n{speech_text}\n</SPEECH>"
            )
        llm_source = "【スライド別素材】\n" + "\n\n".join(chunks)
    else:
        llm_source = cleaned_for_llm if not slide_digest else (
            f"【音声逐語】\n{cleaned_for_llm}\n\n【スライドOCR】\n{slide_digest}"
        )

    # ---- 既定（ヒューリスティック）出力
    if out_kind == "逐語(タイムスタンプ)":
        base_out = to_verbatim_with_timestamps(segments); kind_key = "verbatim"
    elif out_kind == "議事録":
        base_out = heuristic_minutes(segments); kind_key = "minutes"
    elif out_kind == "要旨":
        base_out = heuristic_abstract(segments); kind_key = "abstract"
    elif out_kind == "ガイドライン解説":
        base_out = heuristic_guideline_commentary(slide_groups, slide_notes) if slide_groups else \
                   "【ガイドライン解説（簡易）】\n" + heuristic_article_academic(segments)
        kind_key = "article"
    else:
        base_out = heuristic_article_academic(segments); kind_key = "article"

    final_out = base_out

    # ---------------- 生成AIで整形ボタン ----------------
    st.markdown("---")
    st.subheader("🧠 生成AIで整形する")
    label_lang = "日本語" if output_lang == "ja" else "English"
    do_generate = st.button(f"✨ 生成AIで整形（{label_lang}で出力）")

    if not do_generate:
        st.text_area("結果テキスト", value=final_out or "", height=400)
        return

    # 押下後
    if use_llm and not api_key_input:
        st.error("生成AIで整形がONですが APIキーが未入力です。AI整形は実行できません。")
        st.stop()
    
    # （任意）古い直訳をクリアしておく
    st.session_state.pop("ja_literal_for_article", None)

    final_out = base_out
    try:
        if use_llm and api_key_input:
            if out_kind == "逐語(タイムスタンプ)":
                with st.spinner("生成AIで整形中..."):
                    final_out = llm_rewrite(
                        kind="verbatim",
                        text="【出力は必ず日本語】\n" + st.session_state["transcript_text"],
                        api_key=api_key_input,
                        purpose=purpose,
                        source_lang=detected_lang,
                        target_lang=output_lang,
                    )
            elif out_kind == "直訳（日本語化のみ）":
                with st.spinner("英語→日本語 直訳中..."):
                    final_out = llm_translate_only(
                        text=cleaned_for_llm,              # タイムスタンプ除去版を翻訳
                        api_key=api_key_input,
                        source_lang=detected_lang,
                        target_lang="ja",
                    )
            else:
                # ★★★ ここから追加：記事だけ“直訳→記事調”に切り替える ★★★
                if out_kind == "記事" and (output_lang == "ja"):
                    with st.spinner("英語→日本語 直訳 → 記事調 へ整形中..."):
                        # 1) まずタイムスタンプ除去版を“直訳（日本語）”
                        ja_literal_for_article = llm_translate_only(
                            text=cleaned_for_llm,
                            api_key=api_key_input,
                            source_lang=detected_lang,
                            target_lang="ja",
                        )
                        # 2) 直訳を素材に、重複だけを整理し“記事調（常体）”に
                        final_out = llm_article_from_literal(
                            literal_ja=ja_literal_for_article,
                            api_key=api_key_input,
                            purpose=purpose,
                        )
                        # ★ この直後に置く
                        st.caption("route: ARTICLE_FROM_LITERAL (ja) ✓ 直訳→記事調ルートを通過")
                        
                        # ★ 追加：プレビュー用にも同じ直訳を使えるよう保存
                        st.session_state["ja_literal_for_article"] = ja_literal_for_article
                else:
                    llm_kind_call = {"議事録": "minutes", "要旨": "abstract"}.get(out_kind, "article")
                    parts = split_text_by_chars(llm_source, chunk_size=6000, overlap=300)
                    outs = []
                    N = len(parts)
                    for i, part in enumerate(parts, start=1):
                        meta = (
                            f"【分割パート {i}/{N}】\n"
                            "このパートでは新規情報のみを反映し、既出の見出しや導入は再掲しないでください。"
                        )
                        out_i = llm_rewrite(
                            kind=llm_kind_call,
                            text="【出力は必ず日本語】\n" + meta + "\n\n" + part,
                            api_key=api_key_input,
                            purpose=purpose,
                            source_lang=detected_lang,
                            target_lang=output_lang,
                        )
                        outs.append(out_i.strip())
                    final_out = "\n\n".join(outs)
            st.success("生成AIでの整形が完了しました。")
        else:
            st.info("生成AIがOFFのため、ヒューリスティック整形で出力しました。")
    except Exception as e:
        st.error(f"整形に失敗しました: {e}")

    # ===== 三段表示 =====
    st.subheader("📝 原文（変更前・英語／タイムスタンプ除去）")
    st.text_area("原文", value=cleaned_for_llm, height=260)

    st.subheader("🇯🇵 英語→日本語（直訳・整形なし）")
    if use_llm and api_key_input:
        cached_literal = st.session_state.get("ja_literal_for_article")
        if cached_literal:
            ja_literal = cached_literal
        else:
            with st.spinner("英語→日本語 直訳（プレビュー用）..."):
                ja_literal = llm_translate_only(
                    text=cleaned_for_llm,
                    api_key=api_key_input,
                    source_lang=detected_lang,
                    target_lang="ja",
                )
        st.text_area("直訳", value=ja_literal, height=260)
    else:
        st.text_area("直訳", value="(APIキー未入力または生成AIがOFFのため直訳は表示できません)", height=260)

    st.subheader("📄 整形結果プレビュー")
    if out_kind == "ガイドライン解説" and output_lang == "ja" and final_out:
        for _p in ["背景", "改訂ポイント", "推奨度・エビデンス", "臨床への影響", "課題", "今後"]:
            final_out = re.sub(rf"(#+\s*{_p}\s*\n)(\s*\1)+", r"\1", final_out)
    st.text_area("整形結果", value=final_out, height=380)

    st.download_button("TXTダウンロード", data=final_out.encode("utf-8"), file_name="output.txt")
    docx_bytes = make_docx(title=f"{out_kind}（{purpose}）", content=final_out)
    st.download_button("DOCXダウンロード", data=docx_bytes, file_name="output.docx")


if __name__ == "__main__":
    main()
