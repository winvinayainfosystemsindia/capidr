import json
import sys
sys.stdout.reconfigure(encoding='utf-8')

with open(r"c:\External-projects\WinVinaya\capidr\.last_response.json", "r", encoding="utf-8") as f:
    d = json.load(f)

p = d['pages'][4]
for i, e in enumerate(p['elements']):
    t = e.get('type')
    txt = repr(e.get('text', ''))[:50]
    cap = repr(e.get('caption', ''))[:30]
    img_idx = e.get('image_index')
    print(f"{i}: type={t} img_idx={img_idx} cap={cap} txt={txt}")
