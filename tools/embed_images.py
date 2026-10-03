"""Встраивает картинки из папки img/ в index.html (base64), чтобы сайт работал одним файлом.
Запуск из корня репозитория:  python3 tools/embed_images.py
"""
import base64, json, os, re
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MIME = {'.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.webp': 'image/webp'}
data = {}
img = os.path.join(ROOT, 'img')
for dp, _, files in os.walk(img):
    for f in sorted(files):
        ext = os.path.splitext(f)[1].lower()
        if ext not in MIME: continue
        path = os.path.join(dp, f)
        key = os.path.relpath(path, img).replace(os.sep, '/')
        data[key] = 'data:%s;base64,%s' % (MIME[ext], base64.b64encode(open(path, 'rb').read()).decode())
p = os.path.join(ROOT, 'index.html')
s = open(p, encoding='utf-8').read()
s, n = re.subn(r'/\*IMG\*/.*?/\*/IMG\*/', lambda m: '/*IMG*/' + json.dumps(data, separators=(',', ':')) + '/*/IMG*/', s, count=1, flags=re.S)
assert n == 1, 'метка /*IMG*/ не найдена'
open(p, 'w', encoding='utf-8').write(s)
print('встроено картинок:', len(data), ' размер index.html:', round(len(s.encode()) / 1024), 'КБ')
