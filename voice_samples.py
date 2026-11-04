"""Render the same narration paragraph in 5 Kokoro voices so the channel narrator can be picked by ear.

Usage: python voice_samples.py   -> voice_sample_<FIRST>.wav onwards
"""
import numpy as np
import soundfile as sf
from kokoro import KPipeline

TEXT = ("Welcome back. Today, we are analyzing the underlying shifts in the global economy, "
        "tracking key data points that reveal where the markets are moving next. Let's look at the numbers.")

# The first letter of a Kokoro voice id is its accent: a = American, b = British.
# Round 2: male only. Round 1 (af_heart, af_bella, am_michael, bf_emma, bm_george) was rejected.
VOICES = [
    ("am_fenrir", "American male, deep"),
    ("am_puck", "American male, brighter"),
    ("am_onyx", "American male, low and calm"),
    ("bm_fable", "British male"),
    ("bm_lewis", "British male, deeper"),
]
FIRST = 6  # number files after round 1 so nothing is overwritten

pipelines = {}
rows = []
for i, (voice, desc) in enumerate(VOICES, FIRST):
    lang = voice[0]
    pipe = pipelines.setdefault(lang, KPipeline(lang_code=lang, repo_id="hexgrad/Kokoro-82M"))
    # Kokoro yields one chunk per text segment; join them so each voice is one file.
    audio = np.concatenate([a for _, _, a in pipe(TEXT, voice=voice, speed=1.0)])
    name = f"voice_sample_{i}.wav"
    sf.write(name, audio, 24000)  # Kokoro outputs 24 kHz
    rows.append((name, voice, desc, len(audio) / 24000))

print("\nFile                 Voice        Seconds  Profile")
for name, voice, desc, secs in rows:
    print(f"{name:<20} {voice:<12} {secs:>7.1f}  {desc}")
