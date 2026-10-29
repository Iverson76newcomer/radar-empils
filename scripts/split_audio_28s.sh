#!/bin/bash
# Split audio into 28-second segments (small margin under 30s limit)
# Convert to WAV 16kHz mono PCM

INPUT="/home/z/my-project/upload/audio/Центр_Красок_Краснодар_01.07.2026.m4a"
OUTDIR="/home/z/my-project/upload/audio/wav_28s"
mkdir -p "$OUTDIR"

# Split into 28-second segments
ffmpeg -i "$INPUT" \
  -f segment -segment_time 28 \
  -ac 1 -ar 16000 -b:a 32k \
  -c:a pcm_s16le \
  "$OUTDIR/seg_%04d.wav" -y -loglevel error 2>&1 | tail -3

echo ""
echo "=== 28s WAV segments ==="
echo "Total files: $(ls $OUTDIR/*.wav 2>/dev/null | wc -l)"
echo "Total size: $(du -sh $OUTDIR | cut -f1)"
echo "Sample file size: $(ls -la $OUTDIR/seg_0000.wav | awk '{print $5}') bytes"
