#!/usr/bin/env python3
"""Print all VLM results from partners presentation analysis."""
import json
import os

VLM_DIR = "/home/z/my-project/extracted_images/partners_vlm"

files = sorted(os.listdir(VLM_DIR))
for f in files:
    if not f.endswith(".json"):
        continue
    print("=" * 80)
    print(f"📄 {f}")
    print("=" * 80)
    try:
        with open(os.path.join(VLM_DIR, f)) as fh:
            data = json.load(fh)
        # Try to extract content
        if "choices" in data:
            content = data["choices"][0]["message"]["content"]
            print(content)
        else:
            print(json.dumps(data, ensure_ascii=False, indent=2)[:3000])
    except Exception as e:
        print(f"ERROR: {e}")
    print()
