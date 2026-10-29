#!/usr/bin/env python3
"""
ASR-транскрипция длинного аудиофайла.
Разбивает на 30-секундные фрагменты, распознаёт каждый, объединяет.
"""

import subprocess
import json
import os
import time
import base64

AUDIO_FILE = '/home/z/my-project/upload/Трианон_Новосибирск_compressed.mp3'
PARTS_DIR = '/tmp/asr_parts'
OUTPUT_JSON = '/home/z/my-project/scripts/trianon_transcript.json'
PROGRESS_FILE = '/home/z/my-project/scripts/asr_progress.json'

# Длительность в секундах (из ffprobe: 4173 сек)
TOTAL_DURATION = 4173
SEGMENT_DURATION = 28  # 28 сек (запас от 30)

os.makedirs(PARTS_DIR, exist_ok=True)

# Загружаем прогресс
progress = {}
if os.path.exists(PROGRESS_FILE):
    with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
        progress = json.load(f)
    print(f"Прогресс: {len(progress)} фрагментов уже распознано")

# Сколько всего фрагментов
total_segments = (TOTAL_DURATION + SEGMENT_DURATION - 1) // SEGMENT_DURATION
print(f"Всего фрагментов: {total_segments} (по {SEGMENT_DURATION} сек)")


def transcribe_segment(seg_num):
    """Распознаёт один 30-сек фрагмент."""
    wav_path = f'{PARTS_DIR}/seg_{seg_num:04d}.wav'
    start = seg_num * SEGMENT_DURATION
    
    # Конвертируем фрагмент в WAV 8kHz mono
    if not os.path.exists(wav_path):
        r = subprocess.run([
            'ffmpeg', '-nostdin', '-y', '-i', AUDIO_FILE,
            '-ss', str(start), '-t', str(SEGMENT_DURATION),
            '-ac', '1', '-ar', '8000', '-c:a', 'pcm_s16le',
            wav_path
        ], capture_output=True, text=True, timeout=30)
        if r.returncode != 0:
            return None
    
    # Проверяем размер
    size = os.path.getsize(wav_path)
    if size < 1000:
        return None
    
    # Распознаём через z-ai CLI
    result = subprocess.run([
        'z-ai', 'asr', '-f', wav_path
    ], capture_output=True, text=True, timeout=120)
    
    if result.returncode != 0:
        return None
    
    # Парсим JSON из stdout
    output = result.stdout
    idx = output.find('{')
    if idx < 0:
        return None
    
    try:
        data = json.loads(output[idx:])
        return data.get('text', '')
    except:
        return None


# Обрабатываем фрагменты
texts = []
processed = 0
failed = 0

for seg in range(total_segments):
    if str(seg) in progress:
        texts.append(progress[str(seg)])
        continue
    
    print(f'[{seg+1}/{total_segments}] Фрагмент {seg}...', flush=True)
    
    text = transcribe_segment(seg)
    
    if text is not None:
        texts.append(text)
        progress[str(seg)] = text
        processed += 1
    else:
        texts.append('')
        progress[str(seg)] = ''
        failed += 1
        print(f'  ❌ ошибка', flush=True)
    
    # Сохраняем прогресс каждые 5 фрагментов
    if processed % 5 == 0 and processed > 0:
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(progress, f, ensure_ascii=False)
        full_text = ' '.join(texts)
        with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
            json.dump({'text': full_text, 'segments': texts}, f, ensure_ascii=False, indent=2)
        print(f'  → сохранено (обработано: {processed}, ошибок: {failed})', flush=True)
    
    time.sleep(10)  # пауза 10 сек чтобы не ловить 429

# Финальное сохранение
full_text = ' '.join(texts)
with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
    json.dump({'text': full_text, 'segments': texts}, f, ensure_ascii=False, indent=2)
with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
    json.dump(progress, f, ensure_ascii=False)

print(f'\n{"="*70}')
print(f'ГОТОВО')
print(f'{"="*70}')
print(f'Обработано: {processed} фрагментов')
print(f'Ошибок: {failed}')
print(f'Полный текст: {len(full_text)} символов')
print(f'Сохранено: {OUTPUT_JSON}')
