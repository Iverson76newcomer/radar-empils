#!/bin/bash
# Split audio into 5-minute segments and convert to mp3 16kHz mono
# (ASR works best with smaller files)

INPUT="/home/z/my-project/upload/audio/Центр_Красок_Краснодар_01.07.2026.m4a"
OUTDIR="/home/z/my-project/upload/audio/segments"
mkdir -p "$OUTDIR"

# Split into 5-minute (300 sec) segments, convert to mp3 16kHz mono 32kbps
ffmpeg -i "$INPUT" \
  -f segment -segment_time 300 \
  -ac 1 -ar 16000 -b:a 32k \
  -c:a libmp3lame \
  "$OUTDIR/segment_%03d.mp3" 2>&1 | tail -5

echo ""
echo "=== Segments created ==="
ls -la "$OUTDIR/"
echo ""
echo "Total segments: $(ls $OUTDIR/*.mp3 2>/dev/null | wc -l)"
echo "Total size: $(du -sh $OUTDIR | cut -f1)"
