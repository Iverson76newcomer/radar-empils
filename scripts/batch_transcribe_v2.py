#!/usr/bin/env python3
"""
Improved batch transcription - shorter timeout, single attempt per segment,
saves progress after each segment. Designed to be restartable.
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

os.makedirs(OUTPUT_DIR, exist_ok=True)


def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE) as f:
            return json.load(f)
    return {'completed': [], 'failed': []}


def save_progress(progress):
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, indent=2)


def transcribe_one(wav_path, output_json_path):
    """Single attempt transcription with 120s timeout."""
    try:
        result = subprocess.run(
            ['z-ai', 'asr', '-f', str(wav_path), '-o', output_json_path],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode == 0 and os.path.exists(output_json_path):
            with open(output_json_path) as f:
                data = json.load(f)
            return data.get('text', '').strip()
    except subprocess.TimeoutExpired:
        pass
    except Exception as e:
        print(f'  Error: {e}', flush=True)
    return None


def main():
    segments = sorted(Path(SEGMENTS_DIR).glob('seg_*.wav'))
    total = len(segments)
    progress = load_progress()
    completed_set = set(progress['completed'])

    todo = [s for s in segments if s.name not in completed_set]
    print(f'Total: {total}, completed: {len(completed_set)}, remaining: {len(todo)}', flush=True)

    if not todo:
        print('All done!', flush=True)
        return

    # Process a batch of up to 5 segments with 15s delay between requests
    BATCH_SIZE = 5
    batch = todo[:BATCH_SIZE]
    print(f'Processing batch of {len(batch)} segments (with 15s delay)...', flush=True)

    success = 0
    failed = 0

    for i, seg_path in enumerate(batch):
        seg_name = seg_path.name
        seg_idx = int(seg_name.split('_')[1].split('.')[0])
        start_time = seg_idx * 28
        mins, secs = divmod(start_time, 60)
        timestamp = f'{mins:02d}:{secs:02d}'

        output_json = os.path.join(OUTPUT_DIR, seg_name.replace('.wav', '.json'))

        print(f'[{i+1}/{len(batch)}] {seg_name} ({timestamp})...', end=' ', flush=True)

        text = transcribe_one(seg_path, output_json)

        if text:
            print(f'OK ({len(text)} chars)', flush=True)
            success += 1
            if seg_name not in progress['completed']:
                progress['completed'].append(seg_name)
        else:
            print('FAIL - will retry later', flush=True)
            failed += 1

        save_progress(progress)
        # 15 second delay to avoid rate limiting (429 errors)
        time.sleep(15)

    print(f'\nBatch done: {success} success, {failed} failed', flush=True)
    print(f'Total progress: {len(progress["completed"])}/{total}', flush=True)


if __name__ == '__main__':
    main()
