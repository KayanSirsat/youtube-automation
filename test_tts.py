import soundfile as sf
from kokoro import KPipeline

# 'a' stands for American English. Use 'b' for British, etc.
pipeline = KPipeline(lang_code='a') 
text = "Hello! Kokoro is now running completely locally on my Linux machine."

# Generate audio (using the built-in 'af_heart' voice)
generator = pipeline(text, voice='af_heart', speed=1.0)

for i, (gs, ps, audio) in enumerate(generator):
    sf.write(f"output_{i}.wav", audio, 24000) # Kokoro outputs 24kHz audio
    print(f"Generated output_{i}.wav")
