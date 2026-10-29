#!/bin/bash
# Self-respawning runner - restarts the Python script if it dies
while true; do
  python3 /home/z/my-project/scripts/robust_transcribe.py >> /tmp/asr_robust.log 2>&1
  EXIT_CODE=$?
  echo "[$(date '+%H:%M:%S')] Process exited with code $EXIT_CODE. Restarting in 10s..." >> /tmp/asr_robust.log
  sleep 10
  # Check if completed
  if [ -f /home/z/my-project/upload/audio/transcripts/COMPLETED.marker ]; then
    echo "[$(date '+%H:%M:%S')] COMPLETED! Exiting." >> /tmp/asr_robust.log
    break
  fi
done
