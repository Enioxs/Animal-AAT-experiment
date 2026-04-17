# -*- coding: utf-8 -*-
"""
AAT Animales + Priming — versión 220 trials (2×110) con inversión por bloque
------------------------------------------------------------------------------
- ID exigido: SXX_AAT_1 o SXX_AAT_2  (se guarda como "pre"/"post" en CSV; 1=pre, 2=post).
- Ensayo: Priming (frase) → Fix (+) → ISI → Imagen (377×377) con CUE → Respuesta → ITI.
- Regla por bloque:
    Bloque 1: CÍRCULO = ACERCAR (Approach)
    Bloque 2: CÍRCULO = ALEJAR  (Avoid)
- SOLO priming de SUFRIMIENTO (se elimina neutral).
- Cada imagen se muestra 2 veces: (Sufrimiento) × (Approach/Avoid).
- Balance por bloque: 55 Suf-AP, 55 Suf-AV (total 110).
- TRIGGER de PRIMING: 129=Sufrimiento (antes del onset del texto).
- TRIGGER de IMAGEN: valor de la columna `Numero` (1–110) enviado al onset.

CAMBIO PEDIDO (Trini):
- El código ahora SOLO usa imágenes cuyo `Numero` esté entre 1 y 110.
"""

# ============================== IMPORTS =====================================

import csv
import hashlib
import os
import random
import re
import sys
from functools import lru_cache
from os.path import join

import pandas as pd
import pygame
from pygame.locals import (
    FULLSCREEN as PYG_FULLSCREEN,
    USEREVENT,
    KEYUP,
    K_DOWN,
    K_ESCAPE,
    K_SPACE,
    K_UP,
    QUIT,
    Color,
)

# ===== LSL CONFIG (activar para usar Lab Streaming Layer) =====================
USE_LSL = True                 # <- activa LSL como backend de marcadores
LSL_STREAM_NAME = "AATMarkers" # nombre del stream que verás en EmotivPRO
LSL_STREAM_TYPE = "Markers"    # tipo estándar para streams de marcadores
LSL_CHANNEL_FORMAT = "int32"   # formato de canal (enteros)

try:
    from pylsl import StreamInfo, StreamOutlet, local_clock
except Exception as e:
    if USE_LSL:
        raise RuntimeError(
            "pylsl no está instalado. Instálalo con: pip install pylsl"
        ) from e
    else:
        StreamInfo = StreamOutlet = local_clock = None

# ============================== CONFIG ======================================

# Diseño experimental (AHORA: 2 bloques × 110 = 220 trials)
N_BLOCKS = 2
BLOCK_SIZE = 110
TOTAL_TRIALS = N_BLOCKS * BLOCK_SIZE          # 220
REPS_PER_IMAGE = 2                            # cada imagen dos veces (Suf × AP/AV)

# Pantalla / colores
FULLSCREEN = True
BACKGROUND = "lightgray"    # usado en fases de ensayo; las pantallas de UI usan gradiente
TEXT_COLOR = "black"

# Paleta UI (no altera lógica)
UI_BG_TOP = (244, 246, 250)
UI_BG_BOTTOM = (220, 226, 235)
UI_PANEL = (255, 255, 255)
UI_PANEL_BORDER = (210, 216, 225)
UI_PRIMARY = (55, 94, 148)        # azul sobrio
UI_ACCENT = (144, 97, 192)        # lila suave
UI_MUTED = (105, 112, 121)

# Tiempos (ms)
PRIME_MS = 1500           # duración del priming (texto)
FIX_MS = 500              # cruz de fijación (después del priming)
ISI_MS = 200              # intervalo entre fix e imagen
RESP_WINDOW_MS = 3000
ITI_MS = 500

# Imagen / Animación (AAT)
BASE_IMG_SIZE = 377       # pedido: 377x377 px
ZOOM_TIME_MS = 700        # duración de la animación de zoom
FPS = 60                  # objetivo de frames por segundo

# Entrada
JOYSTICK_AXIS = 1         # vertical en la mayoría de joysticks
JOY_THRESH = 0.75
KEY_APPROACH = K_DOWN
KEY_AVOID = K_UP


# Controles de "continuar" (compatibilidad: espacio, tecla 4 y botón 4 del joystick)
KEY_CONTINUE_1 = K_SPACE
KEY_CONTINUE_2 = pygame.K_4
KEY_CONTINUE_3 = pygame.K_KP4
JOY_CONTINUE_BUTTON = 3   # botón "4" en la mayoría de joysticks (índice 0-based)
# CUE-RESPUESTA (regla *dinámica* por bloque; geom = círculo/cuadrado; sentido cambia)
APPROACH_CUE = "circ"     # sólo para referencia semántica

# Archivos y rutas
EXCEL_ANIMALES = "codigos animales aat.xlsx"  # cols: Grupo, nombre, Codigo, Numero
EXCEL_PRIMING = "_priming x animales.xlsx"    # cols: Grupo, Tipo∈{Neutrales,Sufrimiento}, Frase
ANIMALS_DIR = join("media", "animales")
FIX_DIR = join("media", "fix")
CUE_CIRC_PATH = join(FIX_DIR, "circ.png")
CUE_RECT_PATH = join(FIX_DIR, "rect.png")
FONT_PATH = join("media", "Lexend-VariableFont_wght.ttf")

# Triggers (serial) — se mantienen por compatibilidad, pero no se usan si USE_LSL=True
SEND_TRIGGERS = True       # cambia a False si no quieres enviar por serial
SERIAL_PORT = "COM4"
BAUDRATE = 115200
TRIG_START = 254
TRIG_STOP = 255
TRIG_FIXATION = 244
TRIG_APPROACH = 247
TRIG_AVOID = 245
TRIG_CORRECT = 250
TRIG_INCORRECT = 251
TRIG_PRIME_NEU = 128       # (compatibilidad; no se usa en versión solo sufrimiento)
TRIG_PRIME_SUF = 129       # Sufrimiento

# =============================== UTIL =======================================

def safe_font(size):
    """Carga Lexend si existe; si no, recurre a Arial del sistema."""
    pygame.font.init()
    try:
        if os.path.exists(FONT_PATH):
            return pygame.font.Font(FONT_PATH, size)
    except Exception:
        pass
    return pygame.font.SysFont("arial", size)

# ----------- helpers para texto centrado, wrap y ajuste ---------------

def wrap_lines(font, text, max_width):
    """Devuelve lista de líneas envueltas a max_width (acepta saltos de línea)."""
    out = []
    for paragraph in text.split("\n"):
        if paragraph == "":
            out.append("")
            continue
        words = paragraph.split()
        cur = ""
        for w in words:
            t = (cur + " " + w).strip()
            if font.size(t)[0] <= max_width:
                cur = t
            else:
                if cur:
                    out.append(cur)
                cur = w
        if cur != "" or (len(out) == 0 or out[-1] != ""):
            out.append(cur)
    return out

def measure_block(font, lines, line_gap):
    """Alto total del bloque de líneas con un gap fijo."""
    if len(lines) == 0:
        return 0
    return len(lines) * font.get_height() + (len(lines) - 1) * line_gap

def autosize_and_wrap(base_size, min_size, rect_w, rect_h, text, line_gap=8, width_ratio=0.9):
    """
    Busca un tamaño de fuente (<=base_size y >=min_size) cuyo bloque envuelto
    quepa en (rect_w x rect_h). Devuelve (font, wrapped_lines).
    """
    size = base_size
    while size >= min_size:
        f = safe_font(size)
        maxw = int(rect_w * width_ratio)
        lines = wrap_lines(f, text, maxw)
        total_h = measure_block(f, lines, line_gap)
        if total_h <= rect_h:
            return f, lines
        size -= 2
    # fallback
    f = safe_font(min_size)
    maxw = int(rect_w * width_ratio)
    lines = wrap_lines(f, text, maxw)
    return f, lines

def blit_centered_block(surface, rect, font, lines, color, line_gap=8):
    """Pinta las líneas centradas horizontal y verticalmente dentro de rect."""
    total_h = measure_block(font, lines, line_gap)
    y = rect.y + (rect.h - total_h) // 2
    for ln in lines:
        s = font.render(ln, True, color)
        r = s.get_rect(centerx=rect.centerx, y=y)
        surface.blit(s, r)
        y += font.get_height() + line_gap

def rng_for(subject_id, salt):
    """Genera RNG determinista por sujeto + sal."""
    h = hashlib.sha256(f"{subject_id}-{salt}".encode()).hexdigest()
    return random.Random(int(h[:8], 16))

def find_image_by_code(code, animals_dir=ANIMALS_DIR):
    """Busca la primera imagen cuyo nombre empiece con el código."""
    code_l = str(code).lower()
    for root, _, files in os.walk(animals_dir):
        for f in files:
            name = f.lower()
            if name.startswith(code_l) and os.path.splitext(name)[1] in (".png", ".jpg", ".jpeg"):
                return os.path.join(root, f)
    return None

@lru_cache(maxsize=512)
def load_image_cached(path):
    """Carga y cachea una imagen con alpha."""
    surf = pygame.image.load(path)
    return surf.convert_alpha()

# --------------------------- UI Helpers (bonito) -----------------------------

def draw_vertical_gradient(surface, top_rgb, bottom_rgb):
    """Rellena con gradiente vertical suave (para pantallas de UI)."""
    w, h = surface.get_size()
    tr, tg, tb = top_rgb
    br, bg, bb = bottom_rgb
    for y in range(h):
        t = y / max(1, h - 1)
        r = int(tr + (br - tr) * t)
        g = int(tg + (bg - tg) * t)
        b = int(tb + (bb - tb) * t)
        pygame.draw.line(surface, (r, g, b), (0, y), (w, y))

def draw_soft_shadow(surface, rect, radius=18, alpha=80, spread=12):
    """Sombra difusa para paneles (no afecta imágenes/estímulos)."""
    shadow = pygame.Surface((rect.w + spread * 2, rect.h + spread * 2), pygame.SRCALPHA)
    pygame.draw.rect(shadow, (0, 0, 0, alpha), shadow.get_rect(), border_radius=radius + 6)
    shadow = pygame.transform.smoothscale(shadow, (shadow.get_width(), shadow.get_height()))
    surface.blit(shadow, (rect.x - spread, rect.y - spread))

def draw_panel(surface, rect, border_color=UI_PANEL_BORDER, fill=UI_PANEL, radius=18):
    """Panel con borde y esquinas redondeadas."""
    draw_soft_shadow(surface, rect, radius=radius)
    pygame.draw.rect(surface, fill, rect, border_radius=radius)
    pygame.draw.rect(surface, border_color, rect, 2, border_radius=radius)

def fade_in(screen, draw_fn, duration_ms=400):
    """Transición de fundido (0→1) ejecutando draw_fn en cada frame."""
    clock = pygame.time.Clock()
    start = pygame.time.get_ticks()
    overlay = pygame.Surface(screen.get_size(), pygame.SRCALPHA)
    while True:
        draw_fn()  # dibuja contenido base
        now = pygame.time.get_ticks()
        t = min(1.0, (now - start) / max(1, duration_ms))
        overlay.fill((0, 0, 0, int((1.0 - t) * 180)))
        screen.blit(overlay, (0, 0))
        pygame.display.flip()
        if t >= 1.0:
            break
        clock.tick(FPS)

# =========================== TRIGGERS: SERIAL + LSL ===========================

class SerialSender:
    """Wrapper seguro para enviar un byte (0–255) por serial."""
    def __init__(self, port, baud, enabled=True):
        self.enabled = enabled
        self.ser = None
        if not enabled:
            return
        try:
            import serial
            self.ser = serial.Serial(port=port, baudrate=baud)
            print(f"[TRIG] Serial abierto en {port}")
        except Exception as e:
            print(f"[TRIG] No se pudo abrir serial {port}: {e}")
            self.enabled = False

    def send(self, val):
        if not self.enabled or self.ser is None:
            return
        try:
            ival = int(val)
            if 0 <= ival <= 255:
                self.ser.write((ival).to_bytes(1, "little"))
        except Exception as e:
            print(f"[TRIG] Error trigger {val}: {e}")

    def close(self):
        try:
            if self.ser:
                self.ser.close()
        except Exception:
            pass

class MarkerSink:
    """
    Abstracción para mandar marcadores a uno o ambos backends (LSL / Serial).
    Si USE_LSL=True, usa LSL. Si no, cae a Serial si está habilitado.
    """
    def __init__(self, lsl_outlet=None, serial_sender=None):
        self.lsl_outlet = lsl_outlet
        self.serial_sender = serial_sender

    def send(self, val):
        ival = int(val)
        # LSL
        if self.lsl_outlet is not None:
            try:
                self.lsl_outlet.push_sample([ival], local_clock() if local_clock else None)
            except Exception as e:
                print(f"[TRIG][LSL] Error enviando {ival}: {e}")
        # Serial
        if self.serial_sender is not None:
            self.serial_sender.send(ival)

    def close(self):
        if self.serial_sender is not None:
            try:
                self.serial_sender.close()
            except Exception:
                pass

# ============================== DISPLAY ======================================

def init_display():
    """Inicializa ventana (fullscreen u 1280×720) y devuelve screen, res, center."""
    pygame.init()
    # Importante: inicializar joystick temprano para que pygame capture eventos de botones
    pygame.joystick.init()

    if FULLSCREEN:
        res = (pygame.display.Info().current_w, pygame.display.Info().current_h)
        screen = pygame.display.set_mode(res, PYG_FULLSCREEN)
    else:
        res = (1280, 720)
        screen = pygame.display.set_mode(res)
    pygame.display.set_caption("AAT Animales + Priming (regla dinámica por bloque)")
    center = (res[0] // 2, res[1] // 2)
    return screen, res, center

def draw_text_center(screen, res, font, text, color):
    """Compatibilidad: centrado simple."""
    screen.fill(Color(BACKGROUND))
    lines = wrap_lines(font, text, int(res[0] * 0.75))
    blit_centered_block(screen, pygame.Rect(0, 0, res[0], res[1]), font, lines, color)
    pygame.display.flip()

def show_fixation(screen, center, font, trig=None):
    """Muestra la cruz de fijación centrada y opcionalmente envía trigger."""
    screen.fill(Color(BACKGROUND))
    plus = font.render("+", True, Color(TEXT_COLOR))
    box = plus.get_rect(center=center)
    screen.blit(plus, box)
    pygame.display.flip()
    if trig:
        trig.send(TRIG_FIXATION)
    pygame.time.wait(FIX_MS)

def show_image(screen, res, image_path, cue_surface, base_size=BASE_IMG_SIZE):
    """Dibuja la imagen centrada (377×377) y el cue centrado encima (sin contorno)."""
    screen.fill(Color(BACKGROUND))
    pic = load_image_cached(image_path)
    pic = pygame.transform.smoothscale(pic, (base_size, base_size))
    pic_rect = pic.get_rect(center=(res[0] // 2, res[1] // 2))
    screen.blit(pic, pic_rect)
    cue_rect = cue_surface.get_rect(center=(res[0] // 2, res[1] // 2))
    screen.blit(cue_surface, cue_rect)
    pygame.display.flip()

def zoom_animation(screen, res, image_path, cue_surface, target_factor, duration_ms):
    """
    Anima desde BASE_IMG_SIZE hasta BASE_IMG_SIZE*target_factor en 'duration_ms'.
    target_factor > 1.0 = acercar; < 1.0 = alejar. (sin contorno).
    """
    if duration_ms <= 0 or target_factor == 1.0:
        return

    clock = pygame.time.Clock()
    pic = load_image_cached(image_path)
    start_size = float(BASE_IMG_SIZE)
    end_size = float(BASE_IMG_SIZE) * float(target_factor)
    start_ticks = pygame.time.get_ticks()

    def ease_out_cubic(t):
        t1 = 0.0 if t < 0.0 else (1.0 if t > 1.0 else t)
        return 1.0 - (1.0 - t1) ** 3

    while True:
        now = pygame.time.get_ticks()
        elapsed = now - start_ticks
        t = elapsed / float(duration_ms)

        for e in pygame.event.get():
            if e.type == QUIT or (e.type == KEYUP and e.key == K_ESCAPE):
                pygame.quit()
                sys.exit()

        if t >= 1.0:
            cur_size = int(end_size)
        else:
            k = ease_out_cubic(t)
            cur_size = int(start_size + (end_size - start_size) * k)

        screen.fill(Color(BACKGROUND))
        scaled = pygame.transform.smoothscale(pic, (cur_size, cur_size))
        rect = scaled.get_rect(center=(res[0] // 2, res[1] // 2))
        screen.blit(scaled, rect)
        screen.blit(cue_surface, cue_surface.get_rect(center=(res[0] // 2, res[1] // 2)))
        pygame.display.flip()

        if t >= 1.0:
            break
        clock.tick(FPS)

# =============================== ENTRADA =====================================


# ============================ CONTINUAR / UI ================================

def is_continue_event(e):
    """Retorna True si el evento corresponde a 'continuar' (espacio/4/joystick botón 4)."""
    if e.type in (pygame.KEYDOWN, pygame.KEYUP):
        if e.key in (KEY_CONTINUE_1, KEY_CONTINUE_2, KEY_CONTINUE_3):
            return True
    if e.type in (pygame.JOYBUTTONDOWN, pygame.JOYBUTTONUP):
        # Muchos joysticks reportan botón '4' como índice 3; aceptamos también cualquier botón
        if getattr(e, "button", None) == JOY_CONTINUE_BUTTON:
            return True
        # Fallback: cualquier botón del joystick también sirve para no dejar a la persona atrapada
        if getattr(e, "button", None) is not None:
            return True
    return False

def wait_for_continue(timeout_ms=None):
    """
    Espera de forma robusta a que el/la participante continúe.
    - Acepta ESPACIO, tecla 4 (arriba del teclado), 4 del numpad, y botón 4 del joystick.
    - Si timeout_ms no es None, retorna True si se detecta continuar, False si expira.
    """
    clock = pygame.time.Clock()
    start = pygame.time.get_ticks()

    # Limpiar eventos viejos (evita 'doble pulsación' accidental)
    pygame.event.pump()
    pygame.event.clear()

    while True:
        now = pygame.time.get_ticks()
        if timeout_ms is not None and (now - start) >= timeout_ms:
            return False

        for e in pygame.event.get():
            if e.type == QUIT:
                pygame.quit(); sys.exit()
            if e.type == KEYUP and e.key == K_ESCAPE:
                pygame.quit(); sys.exit()

            if is_continue_event(e):
                # Limpia para que el mismo botón no "salte" dos pantallas seguidas
                pygame.event.clear()
                return True

        clock.tick(60)

def break_countdown_screen(screen, res, center, seconds=10):
    """Muestra un descanso con contador visible y avanza automático al terminar."""
    clock = pygame.time.Clock()
    start = pygame.time.get_ticks()
    total_ms = int(seconds * 1000)

    # Tipografías
    title_font = safe_font(44)
    big_font = safe_font(96)
    small_font = safe_font(26)

    while True:
        now = pygame.time.get_ticks()
        elapsed = now - start
        remaining_ms = max(0, total_ms - elapsed)
        remaining_s = int((remaining_ms + 999) // 1000)

        # Permitir saltar (por si necesitan seguir rápido)
        for e in pygame.event.get():
            if e.type == QUIT:
                pygame.quit(); sys.exit()
            if e.type == KEYUP and e.key == K_ESCAPE:
                pygame.quit(); sys.exit()
            if is_continue_event(e):
                return

        # Dibujo UI
        draw_vertical_gradient(screen, UI_BG_TOP, UI_BG_BOTTOM)
        panel = pygame.Rect(0, 0, int(res[0] * 0.62), int(res[1] * 0.40))
        panel.center = center
        draw_panel(screen, panel)

        # Título
        title_rect = pygame.Rect(panel.x + 24, panel.y + 18, panel.w - 48, 70)
        blit_centered_block(screen, title_rect, title_font, ["Descanso"], UI_PRIMARY, line_gap=0)

        # Número grande
        num_rect = pygame.Rect(panel.x + 24, title_rect.bottom + 6, panel.w - 48, 140)
        blit_centered_block(screen, num_rect, big_font, [str(remaining_s)], Color("black"), line_gap=0)

        # Barra de progreso
        bar_w = int(panel.w * 0.78)
        bar_h = 16
        bar_x = panel.centerx - bar_w // 2
        bar_y = num_rect.bottom + 14
        pygame.draw.rect(screen, (230, 233, 239), (bar_x, bar_y, bar_w, bar_h), border_radius=10)
        progress = 1.0 - (remaining_ms / total_ms if total_ms > 0 else 1.0)
        fill_w = int(bar_w * max(0.0, min(1.0, progress)))
        pygame.draw.rect(screen, UI_ACCENT, (bar_x, bar_y, fill_w, bar_h), border_radius=10)

        # Texto pie
        foot_rect = pygame.Rect(panel.x + 24, bar_y + 26, panel.w - 48, panel.bottom - (bar_y + 40))
        blit_centered_block(
            screen,
            foot_rect,
            small_font,
            ["Descansa", "Continuará automáticamente…"],
            UI_MUTED,
            line_gap=6
        )

        pygame.display.flip()

        if elapsed >= total_ms:
            return

        clock.tick(60)

def get_joystick_or_none():
    """Devuelve el primer joystick inicializado o None."""
    pygame.joystick.init()
    if pygame.joystick.get_count() > 0:
        joy = pygame.joystick.Joystick(0)
        joy.init()
        return joy
    return None

def wait_response(joystick, axis_number=JOYSTICK_AXIS, window_ms=RESP_WINDOW_MS):
    """Espera respuesta (↑/↓ o joystick) hasta 'window_ms' y devuelve (resp, RT)."""
    start = pygame.time.get_ticks()
    TIMEOUT = USEREVENT + 1
    if window_ms and window_ms > 0:
        pygame.time.set_timer(TIMEOUT, window_ms, loops=1)

    while True:
        for e in pygame.event.get():
            if e.type == QUIT or (e.type == KEYUP and e.key == K_ESCAPE):
                pygame.quit()
                sys.exit()

            if e.type == KEYUP:
                if e.key == KEY_APPROACH:
                    pygame.time.set_timer(TIMEOUT, 0)
                    pygame.event.clear()
                    return "in", pygame.time.get_ticks() - start
                if e.key == KEY_AVOID:
                    pygame.time.set_timer(TIMEOUT, 0)
                    pygame.event.clear()
                    return "out", pygame.time.get_ticks() - start

            if joystick and e.type == pygame.JOYAXISMOTION and e.axis == axis_number:
                axis = round(joystick.get_axis(axis_number), 2)
                if axis > JOY_THRESH:
                    pygame.time.set_timer(TIMEOUT, 0)
                    pygame.event.clear()
                    return "in", pygame.time.get_ticks() - start
                if axis < -JOY_THRESH:
                    pygame.time.set_timer(TIMEOUT, 0)
                    pygame.event.clear()
                    return "out", pygame.time.get_ticks() - start

            if e.type == TIMEOUT:
                pygame.time.set_timer(TIMEOUT, 0)
                pygame.event.clear()
                return None, pygame.time.get_ticks() - start

# ================================ DATOS ======================================

def load_tables():
    """Carga excels de animales y priming, normalizando columnas básicas."""
    df_anim = pd.read_excel(EXCEL_ANIMALES)
    df_prime = pd.read_excel(EXCEL_PRIMING)

    for c in ["Grupo", "nombre", "Codigo"]:
        df_anim[c] = df_anim[c].astype(str).str.strip()

    if "Numero" not in df_anim.columns:
        df_anim["Numero"] = range(1, len(df_anim) + 1)

    # ================== CAMBIO: SOLO NUMERO 1..110 ==================
    df_anim = df_anim[df_anim["Numero"].between(1, 110)].reset_index(drop=True)
    # ================================================================

    for c in ["Grupo", "Tipo", "Frase"]:
        df_prime[c] = df_prime[c].astype(str).str.strip()

    return df_anim, df_prime

def pick_prime_phrase(df_prime, grupo, prime_tipo, rng):
    """Toma una frase de priming para el grupo/tipo; fallback al grupo si no hay tipo."""
    sub = df_prime[(df_prime["Grupo"] == grupo) & (df_prime["Tipo"] == prime_tipo)]
    if len(sub) == 0:
        sub = df_prime[df_prime["Grupo"] == grupo]
    if len(sub) == 0:
        return ""
    idx = rng.randrange(len(sub))
    return sub.iloc[idx]["Frase"]

# -------- clave: construcción 220 trials (2×110) con rotación y balance por bloque ----

def build_master_trials(df_anim, df_prime, subject_id):
    """
    Crea 2 bloques × 110 (220) ensayos con balance SOLO SUFRIMIENTO:
    - Selecciona 110 imágenes (sin reemplazo si hay suficientes; con reemplazo si faltan).
    - Para cada imagen, define las 2 condiciones: Suf-AP, Suf-AV.
    - Rotación por bloque: para imagen i y bloque b, condición = combo[(i+b)%2]
      => en cada bloque hay 55 Approach y 55 Avoid, y cada imagen aparece 1 vez por bloque.
    - El cue mostrado se decidirá en run_block según la regla del bloque (se invierte).
    """
    rng = rng_for(subject_id, "master")
    needed_images = TOTAL_TRIALS // REPS_PER_IMAGE  # 110

    if len(df_anim) >= needed_images:
        df_sel = df_anim.sample(
            needed_images, replace=False, random_state=rng.randrange(2**32)
        ).copy()
    else:
        faltan = needed_images - len(df_anim)
        extra = df_anim.sample(
            faltan, replace=True, random_state=rng.randrange(2**32)
        ).copy()
        df_sel = pd.concat([df_anim.copy(), extra], ignore_index=True)

    # combos SOLO sufrimiento
    combos = [
        ("Sufrimiento", "Approach"),
        ("Sufrimiento", "Avoid"),
    ]

    # Barajamos las imágenes una vez, para que la rotación por bloque reparta 50/50
    df_sel = df_sel.sample(frac=1.0, random_state=rng.randrange(2**32)).reset_index(drop=True)

    # Construimos 2 bloques: para cada imagen i, en bloque b usamos combos[(i+b)%2]
    blocks = []
    for b in range(N_BLOCKS):
        trials_b = []
        meta_b = []
        for i, row in df_sel.iterrows():
            prime_tipo, desired_resp = combos[(i + b) % 2]  # rotación
            grupo  = str(row["Grupo"]).strip()
            codigo = str(row["Codigo"]).strip()
            numero = int(row["Numero"])
            nombre = str(row.get("nombre", "")).strip()
            img = find_image_by_code(codigo, ANIMALS_DIR)
            if img is None:
                fallback = join(ANIMALS_DIR, "placeholder.png")
                img = fallback if os.path.exists(fallback) else CUE_RECT_PATH  # algo seguro

            # frase SOLO de sufrimiento (por grupo)
            frase = pick_prime_phrase(df_prime, grupo, prime_tipo, rng)

            trials_b.append((img, None))
            meta_b.append(
                {
                    "grupo": grupo,
                    "codigo": codigo,
                    "numero": numero,
                    "nombre": nombre,
                    "prime_tipo": prime_tipo,       # siempre "Sufrimiento"
                    "prime_frase": frase,
                    "desired_resp": desired_resp,   # "Approach" o "Avoid"
                }
            )

        # shuffle suave dentro del bloque
        pairs = list(zip(trials_b, meta_b))
        rng.shuffle(pairs)
        trials_b, meta_b = zip(*pairs)
        blocks.append((list(trials_b), list(meta_b)))

    return blocks

# ====================== MAPEO DE REGLA Y ACIERTO DINÁMICO ====================

def block_rule_for(block_num):
    """Devuelve 'approach_circ' o 'avoid_circ' según el bloque 1..2."""
    return "approach_circ" if block_num in (1,) else "avoid_circ"

def cue_for_desired_resp(desired_resp, block_rule):
    """
    Dado el objetivo (Approach/Avoid) y la regla del bloque, decide el cue geométrico mostrado.
    - approach_circ:   Approach→círculo, Avoid→cuadrado
    - avoid_circ:      Approach→cuadrado, Avoid→círculo
    """
    if block_rule == "approach_circ":
        return "circ" if desired_resp == "Approach" else "rect"
    else:
        return "rect" if desired_resp == "Approach" else "circ"

def is_correct(cue_geom, resp, block_rule):
    """
    Evalúa acierto según cue mostrado y regla del bloque.
    - approach_circ: círculo≡acercar, cuadrado≡alejar
    - avoid_circ:    círculo≡alejar,  cuadrado≡acercar
    """
    if resp is None:
        return False
    if block_rule == "approach_circ":
        return (cue_geom == "circ" and resp == "in") or (cue_geom == "rect" and resp == "out")
    else:  # avoid_circ
        return (cue_geom == "circ" and resp == "out") or (cue_geom == "rect" and resp == "in")

# ================================ BLOQUES ====================================

def run_block(
    screen,
    res,
    center,
    fonts,
    cues,
    trials,
    meta,
    trig,
    csvwriter,
    subj_name,
    condition,
    uid,
    block_num,
):
    """Ejecuta un bloque completo con instrucciones y logging."""
    char, bigchar = fonts
    cue_circ, cue_rect = cues
    joystick = get_joystick_or_none()

    # Regla de este bloque
    block_rule = block_rule_for(block_num)

    # Mano solicitada por instrucción (se registra en el CSV)
    mano_instruida = "Derecha" if block_num == 1 else "Izquierda"
    regla_txt = (
        "CÍRCULO = ACERCAR (hacia ti) | CUADRADO = ALEJAR (hacia la pantalla)"
        if block_rule == "approach_circ"
        else "CÍRCULO = ALEJAR (hacia la pantalla) | CUADRADO = ACERCAR (hacia ti)"
    )

    # Instrucciones del bloque (UI con centrado real y anti-solape)
    def draw_instructions():
        draw_vertical_gradient(screen, UI_BG_TOP, UI_BG_BOTTOM)
        panel = pygame.Rect(0, 0, int(res[0] * 0.75), int(res[1] * 0.62))
        panel.center = center
        draw_panel(screen, panel)

        # Título
        title_rect = pygame.Rect(panel.x + 24, panel.y + 20, panel.w - 48, int(panel.h * 0.22))
        title_font, title_lines = autosize_and_wrap(
            base_size=96, min_size=44,
            rect_w=title_rect.w, rect_h=title_rect.h,
            text=f"Bloque {block_num}",
            line_gap=6, width_ratio=0.95
        )
        blit_centered_block(screen, title_rect, title_font, title_lines, UI_PRIMARY, line_gap=6)

        # Cuerpo
        body_rect = pygame.Rect(panel.x + 32, title_rect.bottom + 8, panel.w - 64, int(panel.h * 0.54))
        body_text = (
            "Verás un PRIMING (frase) y luego una CRUZ DE FIJACIÓN (+).\n"
            "Después aparecerá la imagen de un animal con un CÍRCULO o CUADRADO en el centro.\n"
            f"Regla de este bloque: {regla_txt}\n"
            f"{'USA LA MANO DERECHA.' if block_num == 1 else 'USA LA MANO IZQUIERDA.'}\n"
            "Responde con joystick alejando y acercándolo. Responde con la mayor rapidez posible."
        )
        body_font, body_lines = autosize_and_wrap(
            base_size=36, min_size=26,
            rect_w=body_rect.w, rect_h=body_rect.h,
            text=body_text, line_gap=10, width_ratio=0.94
        )
        blit_centered_block(screen, body_rect, body_font, body_lines, Color("black"), line_gap=10)

        # Pie
        foot_rect = pygame.Rect(panel.x + 24, body_rect.bottom + 8, panel.w - 48, panel.bottom - (body_rect.bottom + 24))
        foot_font, foot_lines = autosize_and_wrap(
            base_size=30, min_size=22,
            rect_w=foot_rect.w, rect_h=foot_rect.h,
            text="Pulsa ESPACIO o 4 (teclado/joystick) para comenzar.",
            line_gap=6, width_ratio=0.95
        )
        blit_centered_block(screen, foot_rect, foot_font, foot_lines, UI_MUTED, line_gap=6)

    fade_in(screen, draw_instructions, duration_ms=350)

    # Espera inicio
    wait_for_continue()

# Ensayos del bloque
    for idx, (img_path, _) in enumerate(trials, start=1):
        # ===== PRIMING (Trigger + texto) =====
        # En versión SOLO SUFRIMIENTO:
        if trig:
            trig.send(TRIG_PRIME_SUF)

        frase = meta[idx-1]["prime_frase"]
        if frase:
            def draw_prime():
                draw_vertical_gradient(screen, UI_BG_TOP, UI_BG_BOTTOM)
                panel = pygame.Rect(0, 0, int(res[0] * 0.80), int(res[1] * 0.40))
                panel.center = center
                draw_panel(screen, panel)

                text_rect = pygame.Rect(panel.x + 32, panel.y + 24, panel.w - 64, panel.h - 48)
                wrap_font, wrap_lines = autosize_and_wrap(
                    base_size=42, min_size=26,
                    rect_w=text_rect.w, rect_h=text_rect.h,
                    text=frase, line_gap=6, width_ratio=0.92
                )
                blit_centered_block(screen, text_rect, wrap_font, wrap_lines, Color("black"), line_gap=6)

            fade_in(screen, draw_prime, duration_ms=180)
            pygame.time.wait(PRIME_MS)

        # ===== FIX DESPUÉS DEL PRIMING =====
        show_fixation(screen, center, char, trig=trig)

        # ===== ISI =====
        screen.fill(Color(BACKGROUND)); pygame.display.flip()
        pygame.time.wait(ISI_MS)

        # ===== TRIGGER por IMAGEN (Numero) justo antes del onset =====
        img_code = int(meta[idx-1]["numero"])
        trig_val = img_code if 1 <= img_code <= 253 else 0
        if trig:
            trig.send(trig_val)

        # ===== Determinar CUE según regla y objetivo =====
        desired_resp = meta[idx-1]["desired_resp"]  # "Approach" / "Avoid"
        cue_geom = cue_for_desired_resp(desired_resp, block_rule)
        cue_surface = cue_circ if cue_geom == "circ" else cue_rect
        # ===== IMAGEN + CUE (onset) =====
        show_image(screen, res, img_path, cue_surface)

        # ===== RESPUESTA =====
        resp, rt = wait_response(joystick, axis_number=JOYSTICK_AXIS, window_ms=RESP_WINDOW_MS)
        if resp == "in" and trig:  trig.send(TRIG_APPROACH)
        elif resp == "out" and trig: trig.send(TRIG_AVOID)

        # ===== ZOOM (feedback kinestésico) =====
        if resp == "in":
            zoom_animation(screen, res, img_path, cue_surface, target_factor=1.8, duration_ms=ZOOM_TIME_MS)
        elif resp == "out":
            zoom_animation(screen, res, img_path, cue_surface, target_factor=0.5, duration_ms=ZOOM_TIME_MS)

        # ===== ACIERTO =====
        ok = is_correct(cue_geom, resp, block_rule)
        if trig:
            trig.send(TRIG_CORRECT if ok else TRIG_INCORRECT)

        # ===== CSV =====
        tipo_cue = "Círculo" if cue_geom == "circ" else "Cuadrado"
        respuesta_real = "Approach" if resp == "in" else ("Avoid" if resp == "out" else "None")
        input_dev = "Joystick" if joystick else "Keyboard"
        trig_prime = TRIG_PRIME_SUF

        csvwriter.writerow(
            [
                subj_name,                         # ID
                condition,                         # Condicion: 'pre' o 'post'
                block_num,                         # Bloque (1..2)
                idx,                               # TrialEnBloque (1..110)
                meta[idx-1]["grupo"],              # Grupo
                meta[idx-1]["codigo"],             # Codigo
                meta[idx-1]["numero"],             # Numero (IdImagen)
                meta[idx-1]["nombre"],             # Nombre (si disponible)
                meta[idx-1]["prime_tipo"],         # PrimeTipo (siempre Sufrimiento)
                meta[idx-1]["prime_frase"],        # PrimeFrase
                block_rule,                        # BlockRule: approach_circ / avoid_circ
                tipo_cue,                          # CueMostrado: Círculo/Cuadrado
                desired_resp,                      # RespuestaObjetivo: Approach/Avoid
                respuesta_real,                    # RespuestaReal: Approach/Avoid/None
                int(ok),                           # Acierto (0/1)
                (rt if rt is not None else ""),    # RT_ms
                input_dev,                         # FormaRespuesta
                trig_prime,                        # TrigPrime (129)
                trig_val,                          # TrigImagen (Numero)
                mano_instruida,                   # ManoInstruida
            ]
        )

        # ===== ITI =====
        screen.fill(Color(BACKGROUND)); pygame.display.flip()
        pygame.time.wait(ITI_MS)

# ============================== PANTALLA ID ==================================

def id_input_screen(screen, res, center) -> str:
    """Pantalla de ingreso de ID (SXX_AAT_1 / SXX_AAT_2) con UI profesional centrada."""
    title_font = safe_font(44)
    hint_font = safe_font(28)
    input_font_size = 42
    input_font = safe_font(input_font_size)

    subj_name = ""
    pat = re.compile(r"^S\d{2}_AAT_[12]$", flags=re.IGNORECASE)
    error_msg = ""

    def id_already_used(sid: str, data_dir: str = "data") -> bool:
        """
        Retorna True si ya existe un archivo en /data que termine con _{sid}.csv (case-insensitive).
        Esto evita repetir el mismo ID en la misma carpeta de datos.
        """
        if not os.path.isdir(data_dir):
            return False
        sid_l = sid.lower()
        for fn in os.listdir(data_dir):
            if fn.lower().endswith(f"_{sid_l}.csv"):
                return True
        return False

    caret_show = True
    clock = pygame.time.Clock()
    last_toggle = 0

    def draw():
        nonlocal input_font
        draw_vertical_gradient(screen, UI_BG_TOP, UI_BG_BOTTOM)
        panel = pygame.Rect(0, 0, int(res[0] * 0.70), int(res[1] * 0.50))
        panel.center = center
        draw_panel(screen, panel)

        # Título centrado
        title_rect = pygame.Rect(panel.x + 24, panel.y + 18, panel.w - 48, int(panel.h * 0.22))
        t_lines = [ "Ingresa tu ID de participante" ]
        blit_centered_block(screen, title_rect, title_font, t_lines, UI_PRIMARY, line_gap=4)

        # Hint centrado
        hint_rect = pygame.Rect(panel.x + 24, title_rect.bottom, panel.w - 48, int(panel.h * 0.12))
        h_lines = [ "(ejemplo: S00_AAT_1 o S12_AAT_2)" ]
        blit_centered_block(screen, hint_rect, hint_font, h_lines, UI_MUTED, line_gap=4)

        # Caja input
        box_rect = pygame.Rect(0, 0, int(panel.w * 0.78), 64)
        box_rect.center = (panel.centerx, panel.centery + 10)
        pygame.draw.rect(screen, (255, 255, 255), box_rect, border_radius=12)
        pygame.draw.rect(screen, (210, 216, 225), box_rect, 2, border_radius=12)

        # Ajuste de fuente si se pasa
        max_text_w = box_rect.w - 28
        tmp_size = input_font.get_height()
        tmp_font = input_font
        while tmp_font.size(subj_name)[0] > max_text_w and tmp_size > 22:
            tmp_size -= 2
            tmp_font = safe_font(tmp_size)
        input_font = tmp_font

        display_txt = input_font.render(subj_name, True, Color("black"))
        txt_rect = display_txt.get_rect()
        txt_rect.midleft = (box_rect.x + 14, box_rect.centery)
        screen.blit(display_txt, txt_rect)

        # Caret parpadeante
        if caret_show:
            caret_x = txt_rect.right + 2
            pygame.draw.rect(screen, UI_MUTED, (caret_x, box_rect.y + 10, 2, box_rect.h - 20))

        # Error centrado
        if error_msg:
            err_font = safe_font(24)
            err_rect = pygame.Rect(panel.x + 24, box_rect.bottom + 10, panel.w - 48, 30)
            blit_centered_block(screen, err_rect, err_font, [error_msg], Color("red"), line_gap=0)

        # Pie
        foot_font = safe_font(24)
        foot_rect = pygame.Rect(panel.x + 24, panel.bottom - 48, panel.w - 48, 28)
        blit_centered_block(screen, foot_rect, foot_font,
                            ["Presiona ENTER para continuar   |   ESC para salir"], UI_MUTED, line_gap=0)

    fade_in(screen, draw, duration_ms=300)

    while True:
        now = pygame.time.get_ticks()
        if now - last_toggle > 450:
            caret_show = not caret_show
            last_toggle = now

        for e in pygame.event.get():
            if e.type == QUIT:
                pygame.quit(); sys.exit()
            if e.type == KEYUP and e.key == K_ESCAPE:
                pygame.quit(); sys.exit()
            if e.type == pygame.KEYDOWN:
                if e.key == pygame.K_RETURN:
                    if subj_name and pat.match(subj_name):
                        return subj_name
                    else:
                        error_msg = "Formato inválido. Usa SXX_AAT_1 o SXX_AAT_2 (X debe ser 1 o 2)."
                elif e.key == pygame.K_BACKSPACE:
                    subj_name = subj_name[:-1]
                elif len(subj_name) < 12 and e.unicode.isprintable():
                    if not (len(subj_name) == 0 and e.unicode.isspace()):
                        subj_name += e.unicode

        draw()
        pygame.display.flip()
        clock.tick(60)

# ================================= MAIN ======================================

def main():
    # Validaciones mínimas
    required = [EXCEL_ANIMALES, EXCEL_PRIMING, CUE_CIRC_PATH, CUE_RECT_PATH, ANIMALS_DIR]
    for path in required:
        if not os.path.exists(path):
            print(f"[ERROR] No se encuentra: {path}")
            sys.exit(1)

    # Pygame + fuentes + ventana
    screen, res, center = init_display()

    # Pantalla de ID
    subj_name = id_input_screen(screen, res, center)
    uid = subj_name
    condition_raw = subj_name.split("_")[-1]  # "1" o "2"
    # Normalización pedida: 1=pre, 2=post
    condition = "pre" if condition_raw == "1" else ("post" if condition_raw == "2" else condition_raw)

    # Cargar tablas
    df_anim, df_prime = load_tables()

    # CSV
    os.makedirs("data", exist_ok=True)
    from time import gmtime, strftime
    csv_name = join("data", f"{strftime('%Y-%m-%d_%H-%M-%S', gmtime())}_{subj_name}.csv")
    f = open(csv_name, "w", newline="", encoding="utf-8")
    writer = csv.writer(f)

    # Cabeceras CSV
    writer.writerow(
        [
            "ID", "Condicion", "Bloque", "TrialEnBloque",
            "Grupo", "Codigo", "Numero", "Nombre",
            "PrimeTipo", "PrimeFrase",
            "BlockRule", "CueMostrado",
            "RespuestaObjetivo", "RespuestaReal",
            "Acierto", "RT_ms", "FormaRespuesta",
            "TrigPrime", "TrigImagen", "ManoInstruida"
        ]
    )

    # Fuentes
    char = safe_font(32)
    bigchar = safe_font(96)

    # Cues
    cue_circ = pygame.image.load(CUE_CIRC_PATH).convert_alpha()
    cue_rect = pygame.image.load(CUE_RECT_PATH).convert_alpha()

    # ===== Triggers (LSL o Serial) =====
    lsl_outlet = None
    serial_obj = None

    if USE_LSL:
        info = StreamInfo(
            name=LSL_STREAM_NAME,
            type=LSL_STREAM_TYPE,
            channel_count=1,
            nominal_srate=0.0,               # event stream
            channel_format=LSL_CHANNEL_FORMAT,
            source_id="aat_triggers_source",
        )
        lsl_outlet = StreamOutlet(info)
    else:
        serial_obj = SerialSender(SERIAL_PORT, BAUDRATE, enabled=SEND_TRIGGERS)

    trig = MarkerSink(lsl_outlet=lsl_outlet, serial_sender=serial_obj)
    trig.send(TRIG_START)

    try:
        # Bienvenida
        def draw_welcome():
            draw_vertical_gradient(screen, UI_BG_TOP, UI_BG_BOTTOM)
            panel = pygame.Rect(0, 0, int(res[0] * 0.70), int(res[1] * 0.46))
            panel.center = center
            draw_panel(screen, panel)

            title_rect = pygame.Rect(panel.x + 24, panel.y + 20, panel.w - 48, int(panel.h * 0.35))
            t_font, t_lines = autosize_and_wrap(96, 44, title_rect.w, title_rect.h,
                                                "AAT Animales + Priming", line_gap=6, width_ratio=0.95)
            blit_centered_block(screen, title_rect, t_font, t_lines, UI_PRIMARY, line_gap=6)

            mid_rect = pygame.Rect(panel.x + 24, title_rect.bottom, panel.w - 48, int(panel.h * 0.35))
            m_font, m_lines = autosize_and_wrap(36, 24, mid_rect.w, mid_rect.h,
                                                "Sigue las instrucciones en pantalla.", line_gap=6, width_ratio=0.95)
            blit_centered_block(screen, mid_rect, m_font, m_lines, Color("black"), line_gap=6)

            foot_rect = pygame.Rect(panel.x + 24, mid_rect.bottom, panel.w - 48, panel.bottom - (mid_rect.bottom + 20))
            f_font, f_lines = autosize_and_wrap(28, 20, foot_rect.w, foot_rect.h,
                                                "Pulsa ESPACIO o 4 (teclado/joystick) para continuar",
                                                line_gap=4, width_ratio=0.95)
            blit_centered_block(screen, foot_rect, f_font, f_lines, UI_MUTED, line_gap=4)

        fade_in(screen, draw_welcome, duration_ms=350)

        wait_for_continue()

    # Construcción de ensayos (220)
        blocks = build_master_trials(df_anim, df_prime, subj_name)

        # Correr bloques (2 × 110)
        for b, (trials, meta) in enumerate(blocks, start=1):
            run_block(
                screen, res, center,
                (char, bigchar), (cue_circ, cue_rect),
                trials, meta, trig, writer, subj_name, condition, uid, b
            )
            if b < N_BLOCKS:
                # Descanso con contador visible (10s) y avance automático
                break_countdown_screen(screen, res, center, seconds=10)

        # Fin
        def draw_end():
            draw_vertical_gradient(screen, UI_BG_TOP, UI_BG_BOTTOM)
            panel = pygame.Rect(0, 0, int(res[0] * 0.60), int(res[1] * 0.34))
            panel.center = center
            draw_panel(screen, panel)

            text_rect = pygame.Rect(panel.x + 24, panel.y + 24, panel.w - 48, panel.h - 48)
            t_font, t_lines = autosize_and_wrap(42, 26, text_rect.w, text_rect.h,
                                                "El experimento ha terminado. ¡Gracias!", 6, 0.95)
            blit_centered_block(screen, text_rect, t_font, t_lines, UI_PRIMARY, 6)

        fade_in(screen, draw_end, duration_ms=280)

    finally:
        # Cierre robusto (incluye salidas por ESC/SystemExit)
        try:
            if trig:
                trig.send(TRIG_STOP)
        except Exception:
            pass
        try:
            if trig:
                trig.close()
        except Exception:
            pass
        try:
            f.close()
        except Exception:
            pass
        try:
            pygame.time.wait(800)
            pygame.quit()
        except Exception:
            pass

# ================================== BOOT =====================================

if __name__ == "__main__":
    main()
