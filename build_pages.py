#!/usr/bin/env python3
"""يولّد docs/index.html (صفحة ثابتة لـ GitHub Pages) من Index/Styles/Client.
التشغيل من جذر المشروع:  python3 tools/build_pages.py"""
import re, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
read = lambda n: (root / n).read_text(encoding='utf-8')
html = read('Index.html')
html = re.sub(r"<\?!=\s*include\('Styles'\);?\s*\?>", lambda m: read('Styles.html'), html)
html = re.sub(r"<\?!=\s*include\('Client'\);?\s*\?>", lambda m: '<script src="config.js"></script>\n' + read('Client.html'), html)
assert '<?' not in html, 'بقيت وسوم قالب غير مستبدلة'
(root / 'docs').mkdir(exist_ok=True)
(root / 'docs' / 'index.html').write_text(html, encoding='utf-8')
(root / 'docs' / '.nojekyll').write_text('')
print('OK -> docs/index.html', len(html), 'bytes')
