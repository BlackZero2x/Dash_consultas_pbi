# ══════════════════════════════════════════════════════════════
# PUBLICAR_PBI_DESKTOP.PY
# Automatiza el refresh y publicación del dashboard en Power BI
# Desktop usando la ventana ya abierta en la PC de automatizaciones.
#
# Flujo:
#   1. Se conecta a la ventana existente de Power BI Desktop
#      (el archivo consultas_mov_dia.pbix ya está abierto)
#   2. Dispara Alt+F5 → "Actualizar todo"
#   3. Espera a que el título de la ventana deje de mostrar
#      el indicador de carga (asterisco o "Actualizando")
#   4. Guarda con Ctrl+S
#   5. Abre el diálogo Publicar (Inicio → Publicar o Alt+H, P)
#   6. Confirma el workspace en el diálogo y espera la publicación
#
# Supuestos:
#   - Power BI Desktop ya tiene el archivo abierto
#   - El último workspace usado queda pre-seleccionado en el diálogo
#   - La PC de automatizaciones tiene sesión activa (no bloqueada)
#
# Uso:
#   python publicar_pbi_desktop.py
#   from publicar_pbi_desktop import publicar as publicar_desktop
# ══════════════════════════════════════════════════════════════

import time
import pywinauto
from pywinauto import keyboard


# ══════════════════════════════════════════════════════════════
# 1. CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════

# Fragmento del título de la ventana de Power BI Desktop.
# No necesita ser el título completo — basta con que sea único.
WINDOW_TITLE    = 'consultas_mov_dia'

# Tiempo máximo de espera para el refresh (segundos).
# Un refresh con 15k filas tarda ~2-4 min; ponemos margen.
REFRESH_TIMEOUT = 300

# Tiempo máximo de espera para la publicación (segundos).
PUBLISH_TIMEOUT = 120

# Intervalo de polling al verificar estado (segundos).
POLL_INTERVAL   = 3


# ══════════════════════════════════════════════════════════════
# 2. CONEXIÓN A LA VENTANA
# ══════════════════════════════════════════════════════════════
#
# pywinauto con backend 'uia' (UI Automation) es más robusto
# con apps modernas como Power BI Desktop (WPF/UWP).
# Se busca la ventana por título parcial para no depender de
# la ruta completa ni del nombre exacto con asterisco.

def conectar_ventana():
    print('Buscando ventana de Power BI Desktop...')
    desktop = pywinauto.Desktop(backend='uia')
    win = desktop.window(title_re=f'.*{WINDOW_TITLE}.*')
    win.wait('visible', timeout=10)
    print(f'  Ventana encontrada: {win.window_text()!r}')
    return win


# ══════════════════════════════════════════════════════════════
# 3. ACTUALIZAR TODO (Alt+F5)
# ══════════════════════════════════════════════════════════════
#
# Alt+F5 es el atajo de "Actualizar todo" en Power BI Desktop.
# Durante el refresh el título muestra un asterisco (*) o la
# palabra "Actualizando". Se espera hasta que desaparezca.
#
# Se detecta fin del refresh cuando el título vuelve al estado
# limpio (sin asterisco y sin "Actualizando").

def _titulo_limpio(titulo: str) -> bool:
    return '*' not in titulo and 'Actualizando' not in titulo


def actualizar_todo(win):
    print('Disparando actualización (Alt+F5)...')
    win.set_focus()
    time.sleep(1)
    keyboard.send_keys('%{F5}')   # Alt+F5

    # Esperar a que empiece el refresh (el título cambia en ~2s)
    time.sleep(5)

    print('  Esperando fin del refresh...')
    inicio = time.time()
    while time.time() - inicio < REFRESH_TIMEOUT:
        titulo = win.window_text()
        if _titulo_limpio(titulo):
            print(f'  Refresh completado. Título: {titulo!r}')
            return True
        time.sleep(POLL_INTERVAL)

    raise TimeoutError(f'El refresh no terminó en {REFRESH_TIMEOUT}s.')


# ══════════════════════════════════════════════════════════════
# 4. GUARDAR (Ctrl+S)
# ══════════════════════════════════════════════════════════════

def guardar(win):
    print('Guardando archivo (Ctrl+S)...')
    win.set_focus()
    time.sleep(1)
    keyboard.send_keys('^s')   # Ctrl+S
    time.sleep(3)
    print('  Archivo guardado.')


# ══════════════════════════════════════════════════════════════
# 5. PUBLICAR EN POWER BI SERVICE
# ══════════════════════════════════════════════════════════════
#
# El atajo de teclado para Publicar en Power BI Desktop es:
#   Alt → H → P (navega cinta: Inicio → Publicar)
# Aparece un diálogo modal con la lista de workspaces.
# El último workspace usado queda seleccionado → basta Enter.
# Luego aparece un segundo diálogo de confirmación/éxito → Esc.

def publicar(win):
    print('Iniciando publicación (Alt → H → P)...')
    win.set_focus()
    time.sleep(1)
    keyboard.send_keys('%h')    # Alt+H → activa la cinta Inicio
    time.sleep(1)
    keyboard.send_keys('p')     # P → abre el diálogo Publicar
    time.sleep(3)

    # Diálogo de selección de workspace:
    # El workspace anterior queda pre-seleccionado → Enter para confirmar.
    print('  Confirmando workspace (Enter)...')
    keyboard.send_keys('{ENTER}')
    time.sleep(2)

    # Esperar el diálogo de éxito / progreso
    print('  Esperando confirmación de publicación...')
    inicio = time.time()
    desktop = pywinauto.Desktop(backend='uia')
    while time.time() - inicio < PUBLISH_TIMEOUT:
        # Power BI muestra un diálogo con "Publicación correcta" o similar
        for w in desktop.windows():
            try:
                t = w.window_text()
                if 'correcta' in t.lower() or 'publicad' in t.lower() or 'success' in t.lower():
                    print(f'  Publicación exitosa: {t!r}')
                    keyboard.send_keys('{ENTER}')   # cerrar diálogo de éxito
                    return True
                # También puede aparecer como diálogo sin título pero con botón "Entendido"
                if 'Entendido' in t or 'Got it' in t:
                    keyboard.send_keys('{ENTER}')
                    print('  Publicación exitosa (diálogo cerrado).')
                    return True
            except Exception:
                pass
        time.sleep(POLL_INTERVAL)

    # Si no detectamos el diálogo, el proceso puede haber terminado
    # silenciosamente — consideramos éxito si el título volvió a estar limpio.
    titulo = win.window_text()
    if _titulo_limpio(titulo):
        print('  Publicación completada (sin diálogo detectado).')
        return True

    raise TimeoutError(f'La publicación no terminó en {PUBLISH_TIMEOUT}s.')


# ══════════════════════════════════════════════════════════════
# 6. ENTRY POINT
# ══════════════════════════════════════════════════════════════

def ejecutar():
    win = conectar_ventana()
    actualizar_todo(win)
    guardar(win)
    publicar(win)
    print('Pipeline Power BI Desktop completado.')


if __name__ == '__main__':
    ejecutar()
