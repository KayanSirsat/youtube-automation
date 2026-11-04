"""Clone each reference clip in assets/voice/ with Chatterbox-Turbo and read the test paragraph.

Usage: .venv-tts/bin/python clone_samples.py   -> assets/voice/sample_<name>.wav for each ref_<name>.wav
Reference clips: 10-15 s of clean speech as WAV. Convert phone recordings first:
    .venv/bin/ffmpeg -i ref_natural.m4a -ac 1 -ar 24000 assets/voice/ref_natural.wav
Runs in .venv-tts (Chatterbox pins torch 2.6, which would break the main .venv).
"""
from pathlib import Path

import torchaudio as ta
from chatterbox.tts_turbo import ChatterboxTurboTTS

TEXT = ("Welcome back. Today, we are analyzing the underlying shifts in the global economy, "
        "tracking key data points that reveal where the markets are moving next. Let's look at the numbers.")

refs = sorted(Path("assets/voice").glob("ref_*.wav"))
if not refs:
    raise SystemExit("No assets/voice/ref_*.wav clips found")

model = ChatterboxTurboTTS.from_pretrained(device="cpu")  # GTX 1650's 4 GB is below Turbo's VRAM need
for ref in refs:
    wav = model.generate(TEXT, audio_prompt_path=str(ref))
    out = ref.with_name(f"sample_{ref.stem.removeprefix('ref_')}.wav")
    ta.save(str(out), wav, model.sr)
    print(f"{out}  <- cloned from {ref.name}")
