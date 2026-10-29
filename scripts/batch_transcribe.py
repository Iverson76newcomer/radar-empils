#!/usr/bin/env python3
"""
Batch transcribe audio segments using z-ai ASR CLI.
Processes segments in order, with retry logic and progress tracking.
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
FINAL_TEXT_FILE = '/home/z/my-project/upload/audio/transcripts/full_transcript.txt'

os.makedirs(OUTPUT_DIR, exist_ok=True)


def load_progress():
    """Load progress to resume from where we stopped."""
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE) as f:
            return json.load(f)
    return {'completed': [], 'failed': []}


def save_progress(progress):
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, indent=2)


def transcribe_segment(wav_path, output_json_path, max_retries=3):
    """Transcribe a single WAV segment with retries."""
    for attempt in range(max_retries):
        try:
            result = subprocess.run(
                ['z-ai', 'asr', '-f', wav_path, '-o', output_json_path],
                capture_output=True, text=True, timeout=300
            )
            if result.returncode == 0 and os.path.exists(output_json_path):
                with open(output_json_path) as f:
                    data = json.load(f)
                return data.get('text', '').strip()
            else:
                print(f'    Attempt {attempt+1} failed: {result.stderr[:200]}')
                time.sleep(3)
        except subprocess.TimeoutExpired:
            print(f'    Attempt {attempt+1} timed out (300s)')
            time.sleep(5)
        except Exception as e:
            print(f'    Attempt {attempt+1} error: {e}')
            time.sleep(3)
    return None


def main():
    # Get all segments sorted
    segments = sorted(Path(SEGMENTS_DIR).glob('seg_*.wav'))
    total = len(segments)
    print(f'Total segments to process: {total}')

    progress = load_progress()
    completed_set = set(progress['completed'])

    # Check which are already done
    todo = [s for s in segments if s.name not in completed_set]
    print(f'Already completed: {len(completed_set)}')
    print(f'Remaining: {len(todo)}')
    print()

    if not todo:
        print('All segments already processed. Use --combine to build final text.')
        return

    success_count = 0
    fail_count = 0

    for i, seg_path in enumerate(todo):
        seg_name = seg_path.name
        seg_idx = int(seg_name.split('_')[1].split('.')[0])
        start_time = seg_idx * 28  # seconds

        # Format timestamp
        mins = start_time // 60
        secs = start_time % 60
        timestamp = f'{mins:02d}:{secs:02d}'

        output_json = os.path.join(OUTPUT_DIR, seg_name.replace('.wav', '.json'))

        print(f'[{i+1}/{len(todo)}] {seg_name} (timestamp {timestamp})...', end=' ', flush=True)

        text = transcribe_segment(seg_path, output_json)

        if text:
            print(f'OK ({len(text)} chars)')
            success_count += 1
            progress['completed'].append(seg_name)
        else:
            print('FAILED')
            fail_count += 1
            progress['failed'].append(seg_name)

        # Save progress every 5 segments
        if (i + 1) % 5 == 0:
            save_progress(progress)
            print(f'  [Progress saved: {success_count} success, {fail_count} failed]')

        # Small delay to avoid rate limits
        time.sleep(0.3)

    save_progress(progress)

    print()
    print(f'=== Final results ===')
    print(f'Success: {success_count}')
    print(f'Failed: {fail_count}')
    print(f'Total completed: {len(progress["completed"])}/{total}')


if __name__ == '__main__':
    main()
