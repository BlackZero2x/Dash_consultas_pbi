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


_MESES_ES = {
    1: 'ene', 2: 'feb', 3: 'mar', 4: 'abr', 5: 'may', 6: 'jun',
    7: 'jul', 8: 'ago', 9: 'sep', 10: 'oct', 11: 'nov', 12: 'dic',
}


def _aplicar_filtro_fecha_slicer(page):
    """Hace clic en el botón de hoy en el slicer de fechas de PBI."""
    from datetime import date
    hoy = date.today()
    # Formato que usa el slicer: "12-may", "08-may", etc.
    texto_hoy = f'{hoy.day:02d}-{_MESES_ES[hoy.month]}'

    selector = '.visual-advancedSlicerVisual .small-multiples-grid-cell-content.buttonSlicerVisual'
    try:
        botones = page.locator(selector)
        count = botones.count()
        for i in range(count):
            btn = botones.nth(i)
            txt = (btn.inner_text() or '').strip()
            if texto_hoy in txt:
                # Solo clicar si no está ya seleccionado
                if btn.get_attribute('aria-pressed') != 'true':
                    btn.click()
                    time.sleep(3)
                return True
    except Exception:
        pass
    return False


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

        def _url_base(url):
            return url.split('?')[0].rstrip('/')

        if page is None:
            page = context.new_page()
            page.goto(report_url, wait_until='domcontentloaded', timeout=60000)
            time.sleep(15)
        elif _url_base(report_url) != _url_base(page.url):
            page.goto(report_url, wait_until='domcontentloaded', timeout=60000)
            time.sleep(10)
        else:
            # Misma página base — navegar igualmente para aplicar el filtro de fecha
            page.goto(report_url, wait_until='domcontentloaded', timeout=60000)
            time.sleep(10)

        # Esperar a que desaparezcan los spinners de carga de Power BI
        try:
            page.wait_for_selector(
                '[class*="spinner"], [class*="loading"], [aria-label*="cargando"], [aria-label*="Loading"]',
                state='hidden',
                timeout=60000,
            )
        except Exception:
            pass

        time.sleep(render_wait)

        # Forzar filtro de fecha a hoy via clic en slicer
        filtro_aplicado = _aplicar_filtro_fecha_slicer(page)
        if filtro_aplicado:
            # Esperar re-renderizado tras el clic
            try:
                page.wait_for_selector(
                    '[class*="spinner"], [class*="loading"]',
                    state='hidden',
                    timeout=30000,
                )
            except Exception:
                pass
            time.sleep(5)

        ultimo_corte = _leer_ultimo_corte(page)

        # Ampliar viewport para captura en alta resolución
        page.set_viewport_size({'width': 1920, 'height': 1080})
        time.sleep(2)

        vp = page.viewport_size or {'width': 1920, 'height': 1080}
        clip = {
            'x': clip_x,
            'y': clip_y,
            'width':  min(clip_w, vp['width']  - clip_x),
            'height': min(clip_h, vp['height'] - clip_y),
        }

        # Intentar capturar el contenedor del reporte PBI directamente
        _SELECTOR_PBI = (
            '[class*="reportPage"], [class*="report-page"], '
            '[class*="canvasFlexBox"], [class*="report-canvas"]'
        )
        try:
            elemento = page.locator(_SELECTOR_PBI).first
            elemento.wait_for(state='visible', timeout=10000)
            elemento.screenshot(path=dest, scale='device')
            clip = None  # indica que se usó selector
        except Exception:
            # Fallback: clip fijo calibrado
            page.screenshot(path=dest, clip=clip, scale='device')

        if ultimo_corte:
            txt_path = os.path.splitext(dest)[0] + '.txt'
            with open(txt_path, 'w', encoding='utf-8') as f:
                f.write(ultimo_corte)

    print(json.dumps({
        'ok': True,
        'ultimo_corte': ultimo_corte,
        'viewport': vp,
        'clip': clip,
        'metodo': 'selector' if clip is None else 'clip_fijo',
        'filtro_fecha': 'slicer_clic' if filtro_aplicado else 'sin_cambio',
    }))


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(json.dumps({'ok': False, 'error': str(e)}))
        sys.exit(1)
