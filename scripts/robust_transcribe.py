#!/usr/bin/env python3
"""
Robust background transcription runner.
Uses 30s delay between requests to avoid 429 rate limit.
Saves progress after each segment. Auto-restarts on failure.
Runs forever until all 186 segments are done.
"""

import os
import subprocess
import json
import time
import sys
from pathlib import Path

SEGMENTS_DIR = '/home/z/my-project/upload/audio/wav_28s'
OUTPUT_DIR = '/home/z/my-project/upload/audio/transcripts'
PROGRESS_FILE = '/home/z/my-project/upload/audio/transcripts/progress.json'
LOG_FILE = '/tmp/asr_robust.log'

os.makedirs(OUTPUT_DIR, exist_ok=True)


def log(msg):
    """Log with timestamp, flush immediately."""
    from datetime import datetime
    ts = datetime.now().strftime('%H:%M:%S')
    line = f'[{ts}] {msg}'
    print(line, flush=True)
    with open(LOG_FILE, 'a') as f:
        f.write(line + '\n')


def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE) as f:
            return json.load(f)
    return {'completed': [], 'failed': []}


def save_progress(progress):
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, indent=2)


def transcribe_one(wav_path, output_json_path):
    """Single attempt, 90s timeout."""
    try:
        result = subprocess.run(
            ['z-ai', 'asr', '-f', str(wav_path), '-o', output_json_path],
            capture_output=True, text=True, timeout=90
        )
        if result.returncode == 0 and os.path.exists(output_json_path):
            with open(output_json_path) as f:
                data = json.load(f)
            return data.get('text', '').strip()
        # Check for 429 in stderr
        if '429' in result.stderr or 'Too many requests' in result.stderr:
            return 'RATE_LIMITED'
    except subprocess.TimeoutExpired:
        return 'TIMEOUT'
    except Exception as e:
        log(f'  Exception: {e}')
    return None


def main():
    log('=== Starting robust transcription runner ===')

    while True:
        # Load progress
        progress = load_progress()
        completed_set = set(progress['completed'])

        # Get all segments
        segments = sorted(Path(SEGMENTS_DIR).glob('seg_*.wav'))
        total = len(segments)

        # Find next to do
        todo = [s for s in segments if s.name not in completed_set]

        if not todo:
            log(f'🎉 ALL DONE! {len(completed_set)}/{total} segments completed.')
            # Touch a completion marker file
            with open('/home/z/my-project/upload/audio/transcripts/COMPLETED.marker', 'w') as f:
                f.write(f'Completed at {time.strftime("%Y-%m-%d %H:%M:%S")}\n')
                f.write(f'Total segments: {total}\n')
            return

        # Process ONE segment at a time, then loop
        seg_path = todo[0]
        seg_name = seg_path.name
        seg_idx = int(seg_name.split('_')[1].split('.')[0])
        start_time = seg_idx * 28
        mins, secs = divmod(start_time, 60)
        timestamp = f'{mins:02d}:{secs:02d}'

        output_json = os.path.join(OUTPUT_DIR, seg_name.replace('.wav', '.json'))

        log(f'[{len(completed_set)+1}/{total}] {seg_name} ({timestamp})...')

        result = transcribe_one(seg_path, output_json)

        if result and result not in ('RATE_LIMITED', 'TIMEOUT', None):
            log(f'  ✅ OK ({len(result)} chars)')
            if seg_name not in progress['completed']:
                progress['completed'].append(seg_name)
            save_progress(progress)
            # Small delay between successful requests
            time.sleep(5)
        elif result == 'RATE_LIMITED':
            log(f'  ⏳ Rate limited. Waiting 120s...')
            time.sleep(120)
        elif result == 'TIMEOUT':
            log(f'  ⏰ Timeout. Will retry. Waiting 30s...')
            time.sleep(30)
        else:
            log(f'  ❌ Failed. Will retry. Waiting 60s...')
            time.sleep(60)

        # Continue loop - will pick up next segment


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        log('Stopped by user')
    except Exception as e:
        log(f'Fatal error: {e}')
        # Auto-restart after 60s
        log('Auto-restarting in 60s...')
        time.sleep(60)
        os.execv(sys.executable, [sys.executable] + sys.argv)
