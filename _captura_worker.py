# ══════════════════════════════════════════════════════════════
# _CAPTURA_WORKER.PY  —  ejecutado por capturar_pbi.py en subprocess
#
# Uso interno:
#   python _captura_worker.py <dest_png> <report_url> <render_wait>
#                             <clip_x> <clip_y> <clip_w> <clip_h>
#
# Salida stdout (JSON):  {"ok": true, "ultimo_corte": "..."}
#                    o:  {"ok": false, "error": "..."}
# ══════════════════════════════════════════════════════════════

import sys
import os
import json
import time

# Suprimir warnings de proxy y Node.js antes de importar Playwright
for _pvar in ('HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy'):
    os.environ.pop(_pvar, None)
os.environ['NODE_NO_WARNINGS'] = '1'

from playwright.sync_api import sync_playwright


def _leer_ultimo_corte(page):
    try:
        candidates = page.locator('text=/\\d{2}\\/\\d{2}\\/\\d{4}/')
        for i in range(candidates.count()):
            txt = candidates.nth(i).inner_text().strip()
            if len(txt) >= 16 and ('AM' in txt or 'PM' in txt):
                return txt
    except Exception:
        pass
    return None


def main():
    dest, report_url, render_wait, clip_x, clip_y, clip_w, clip_h = sys.argv[1:]
    render_wait = int(render_wait)
    clip_x, clip_y, clip_w, clip_h = int(clip_x), int(clip_y), int(clip_w), int(clip_h)

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp('http://localhost:9222')
        context = browser.contexts[0]

        page = None
        for pg in context.pages:
            if 'app.powerbi.com' in pg.url:
                page = pg
                break

        if page is None:
            page = context.new_page()
            page.goto(report_url, wait_until='domcontentloaded', timeout=60000)
            time.sleep(15)
        elif report_url.split('?')[0].split('/')[-1] not in page.url:
            page.goto(report_url, wait_until='domcontentloaded', timeout=60000)
            time.sleep(10)
        else:
            page.reload(wait_until='domcontentloaded', timeout=60000)
            time.sleep(20)

        time.sleep(render_wait)

        ultimo_corte = _leer_ultimo_corte(page)

        vp = page.viewport_size or {'width': 1366, 'height': 768}
        clip = {
            'x': clip_x,
            'y': clip_y,
            'width':  min(clip_w, vp['width']  - clip_x),
            'height': min(clip_h, vp['height'] - clip_y),
        }

        session = context.new_cdp_session(page)
        result = session.send('Page.captureScreenshot', {
            'format': 'png',
            'optimizeForSpeed': True,
            'clip': {**clip, 'scale': 1},
        })
        session.detach()

        import base64
        with open(dest, 'wb') as f:
            f.write(base64.b64decode(result['data']))

        if ultimo_corte:
            txt_path = os.path.splitext(dest)[0] + '.txt'
            with open(txt_path, 'w', encoding='utf-8') as f:
                f.write(ultimo_corte)

    print(json.dumps({'ok': True, 'ultimo_corte': ultimo_corte}))


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(json.dumps({'ok': False, 'error': str(e)}))
        sys.exit(1)
