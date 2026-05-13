"""
voice.py — local speech-to-text for tsifl using OpenAI Whisper.

Records from the default microphone, runs Whisper locally, returns the
transcribed text. No API calls, no per-second cost, no network round-trip.

First-press cost: ~140MB one-time download of the 'base' model (cached at
~/.cache/whisper/). Subsequent presses are instant model load + a few
seconds of inference per ~10s of audio.

Dependencies (lazy-imported):
- openai-whisper  (the model + inference)
- sounddevice     (record from default mic)
- numpy           (audio buffer math; whisper already needs it)

Public API:
- start_recording()  → begin capturing; returns True if started
- stop_recording()   → stop + transcribe; returns (text, error)
- is_recording()     → bool
- ensure_dependencies() → bool; True if everything imports
"""

from __future__ import annotations
import os
import sys
import threading
import time
from pathlib import Path
from typing import Optional


# ── Module state ────────────────────────────────────────────────────────────
_recording = False
_recording_thread: Optional[threading.Thread] = None
_audio_buffer: list = []  # list of numpy arrays
_recording_started_at: float = 0.0
_sample_rate = 16000  # Whisper expects 16kHz mono

# Lazy-loaded references
_whisper_model = None
_sd_module = None
_numpy_module = None


# Path to cache the loaded Whisper model — we keep it in memory after
# first load so subsequent transcriptions are fast.
_WHISPER_MODEL_SIZE = os.environ.get("TSIFL_WHISPER_MODEL", "base")
# Sizes: tiny (39MB), base (140MB), small (470MB), medium (1.5GB), large (3GB)
# `base` is the sweet spot for English on Apple Silicon: accurate enough
# for natural speech, fast enough that 10s of audio transcribes in ~3s.


# ── Dependency loading ─────────────────────────────────────────────────────

def ensure_dependencies() -> tuple[bool, str]:
    """Try to import whisper + sounddevice. Returns (ok, error_msg)."""
    global _sd_module, _numpy_module
    if _sd_module is not None and _numpy_module is not None:
        return True, ""

    try:
        import sounddevice as _sd
        import numpy as _np
        _sd_module = _sd
        _numpy_module = _np
    except ImportError as e:
        return False, (
            f"Voice input needs `sounddevice` and `numpy`. Install with:\n"
            f"  pip install sounddevice numpy openai-whisper\n"
            f"Error: {e}"
        )

    try:
        import whisper  # noqa: F401
    except ImportError as e:
        return False, (
            f"Voice input needs `openai-whisper`. Install with:\n"
            f"  pip install openai-whisper\n"
            f"Error: {e}"
        )

    return True, ""


def _ensure_model():
    """Lazy-load the Whisper model.

    First call downloads ~140MB to ~/.cache/whisper/. On networks with SSL
    inspection (corp/edu MITM), the default urlopen rejects the connection
    with CERTIFICATE_VERIFY_FAILED. We monkey-patch ssl to unverified for
    the download only — Whisper checks the file SHA after download anyway,
    so the integrity guarantee is preserved.
    """
    global _whisper_model
    if _whisper_model is not None:
        return _whisper_model

    import whisper
    from pathlib import Path as _P

    # Check if the model file already exists locally — skip the SSL dance
    # if we're loading from disk. Whisper caches at ~/.cache/whisper/<name>.pt
    cache_root = _P.home() / ".cache" / "whisper"
    cached = cache_root / f"{_WHISPER_MODEL_SIZE}.pt"
    needs_download = not cached.exists()

    if needs_download:
        sys.stderr.write(
            f"[voice] downloading whisper '{_WHISPER_MODEL_SIZE}' model "
            f"(~140MB, one-time)…\n"
        )

    # Monkey-patch SSL verification for the duration of the load call.
    # Whisper hashes the file after download so a MITM substitution would
    # still fail integrity checks — bypassing SSL here is safe.
    import ssl
    _orig_https_ctx = ssl._create_default_https_context
    try:
        ssl._create_default_https_context = ssl._create_unverified_context
        _whisper_model = whisper.load_model(_WHISPER_MODEL_SIZE)
    finally:
        ssl._create_default_https_context = _orig_https_ctx

    sys.stderr.write(f"[voice] model '{_WHISPER_MODEL_SIZE}' loaded\n")
    return _whisper_model


# ── Recording lifecycle ────────────────────────────────────────────────────

def is_recording() -> bool:
    return _recording


def start_recording() -> tuple[bool, str]:
    """Begin capturing audio from the default mic. Non-blocking.

    Returns (True, "") on success, (False, error_msg) otherwise.
    """
    global _recording, _recording_thread, _audio_buffer, _recording_started_at

    if _recording:
        return False, "Already recording"

    ok, err = ensure_dependencies()
    if not ok:
        return False, err

    _audio_buffer = []
    _recording = True
    _recording_started_at = time.time()

    def _record_loop():
        sd = _sd_module
        np = _numpy_module
        try:
            # Start an input stream and append blocks to the buffer until
            # _recording becomes False.
            def _callback(indata, frames, time_info, status):
                if status:
                    sys.stderr.write(f"[voice] sounddevice status: {status}\n")
                # Copy because sounddevice reuses the buffer
                _audio_buffer.append(indata.copy())

            with sd.InputStream(
                samplerate=_sample_rate,
                channels=1,
                dtype="float32",
                callback=_callback,
            ):
                # Cap recording length to 60s as a safety net
                while _recording and (time.time() - _recording_started_at) < 60:
                    time.sleep(0.05)
        except Exception as e:
            sys.stderr.write(f"[voice] recording crashed: {e}\n")

    _recording_thread = threading.Thread(target=_record_loop, daemon=True)
    _recording_thread.start()
    sys.stderr.write("[voice] recording started\n")
    return True, ""


def stop_recording() -> tuple[Optional[str], str]:
    """Stop recording, transcribe with Whisper, return (text, error).

    On error, text is None and error is populated.
    """
    global _recording, _audio_buffer
    if not _recording:
        return None, "Not recording"

    _recording = False
    duration = time.time() - _recording_started_at
    sys.stderr.write(f"[voice] recording stopped after {duration:.1f}s\n")

    # Wait briefly for the recording thread to drain its last callback
    if _recording_thread is not None:
        _recording_thread.join(timeout=2.0)

    if not _audio_buffer:
        return None, "No audio captured. Try again — and check System Settings → Privacy → Microphone."

    np = _numpy_module
    audio = np.concatenate(_audio_buffer, axis=0).flatten()

    # Sanity check: at least 0.5s of audio
    if len(audio) < _sample_rate * 0.5:
        return None, "Too short. Hold for at least half a second."

    # Transcribe directly from the numpy buffer. This sidesteps ffmpeg
    # entirely (Whisper only needs ffmpeg when given a file path / bytes;
    # numpy float32 at 16kHz mono is its native input format).
    try:
        model = _ensure_model()
        sys.stderr.write(f"[voice] transcribing {duration:.1f}s of audio…\n")
        t0 = time.time()
        # Whisper expects mono float32 at 16kHz. Our recording already is.
        result = model.transcribe(audio, language="en", fp16=False)
        elapsed = time.time() - t0
        text = (result.get("text") or "").strip()
        sys.stderr.write(f"[voice] transcribed in {elapsed:.1f}s: {text!r}\n")
    except Exception as e:
        return None, f"Transcription failed: {e}"

    if not text:
        return None, "Whisper couldn't hear anything. Speak a bit louder."
    return text, ""


def cancel_recording():
    """Drop the current recording without transcribing."""
    global _recording, _audio_buffer
    _recording = False
    _audio_buffer = []
