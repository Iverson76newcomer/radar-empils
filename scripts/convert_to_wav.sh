#!/bin/bash
# Convert m4a segments to WAV format (16kHz mono, PCM 16-bit)
# This format is required by the ASR service

INDIR="/home/z/my-project/upload/audio/segments"
OUTDIR="/home/z/my-project/upload/audio/wav_segments"
mkdir -p "$OUTDIR"

# Process each mp3 segment
for mp3 in "$INDIR"/segment_*.mp3; do
  base=$(basename "$mp3" .mp3)
  wav="$OUTDIR/${base}.wav"
  if [ ! -f "$wav" ]; then
    ffmpeg -i "$mp3" -ac 1 -ar 16000 -c:a pcm_s16le "$wav" -y -loglevel error 2>&1 | tail -3
    echo "Converted: $base.wav"
  fi
done

echo ""
echo "=== WAV segments ==="
ls -la "$OUTDIR/" | head -25
echo ""
echo "Total WAV files: $(ls $OUTDIR/*.wav 2>/dev/null | wc -l)"
echo "Total size: $(du -sh $OUTDIR | cut -f1)"
