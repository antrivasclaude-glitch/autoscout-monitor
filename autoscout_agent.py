"""
AutoScout24 Agent — Versión Railway
- Configuración: config.json (en el mismo directorio)
- Credenciales Google: variable de entorno GOOGLE_CREDENTIALS_JSON (base64)
- Contraseña email: variable de entorno EMAIL_PASSWORD
- Estado entre ejecuciones: guardado en Google Sheets (hoja oculta "Estado")
- Logs: stdout (Railway los captura automáticamente)
"""

import os
import sys
import json
import time
import base64
import smtplib
import logging
import tempfile
import requests
from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials


# ══════════════════════════════════════════════════════════════
#  RUTAS
# ══════════════════════════════════════════════════════════════

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def ruta(filename: str) -> str:
    return os.path.join(SCRIPT_DIR, filename)


# ══════════════════════════════════════════════════════════════
#  CARGA DE CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════

def cargar_config() -> dict:
    config_path = ruta("config.json")
    if not os.path.exists(config_path):
        print(f"[ERROR FATAL] No se encuentra config.json en: {config_path}")
        sys.exit(1)
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # La contraseña del email viene de variable de entorno (más seguro)
    cfg["email"]["password"] = os.environ.get("EMAIL_PASSWORD", "")
    if not cfg["email"]["password"]:
        print("[ERROR FATAL] Variable de entorno EMAIL_PASSWORD no definida.")
        sys.exit(1)

    return cfg


CFG = cargar_config()


# ══════════════════════════════════════════════════════════════
#  LOGGING — stdout (Railway lo captura y muestra en dashboard)
# ══════════════════════════════════════════════════════════════

NIVEL = getattr(logging, CFG["log"].get("nivel", "INFO").upper(), logging.INFO)

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)-8s] %(funcName)-28s — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger("autoscout")


# ══════════════════════════════════════════════════════════════
#  CREDENCIALES GOOGLE — desde variable de entorno
# ══════════════════════════════════════════════════════════════

def obtener_credenciales_google() -> str:
    """
    Decodifica GOOGLE_CREDENTIALS_JSON (base64) y lo guarda en un
    archivo temporal. Devuelve la ruta al archivo temporal.
    """
    b64 = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if not b64:
        log.critical("Variable de entorno GOOGLE_CREDENTIALS_JSON no definida.")
        sys.exit(1)

    try:
        json_bytes = base64.b64decode(b64)
        # Validar que es JSON correcto
        json.loads(json_bytes)
    except Exception as e:
        log.critical(f"GOOGLE_CREDENTIALS_JSON no es un base64 válido: {e}")
        sys.exit(1)

    # Escribir en archivo temporal (se borra al terminar el proceso)
    tmp = tempfile.NamedTemporaryFile(
        suffix=".json", delete=False, mode="wb"
    )
    tmp.write(json_bytes)
    tmp.close()
    log.debug(f"Credenciales Google escritas en temporal: {tmp.name}")
    return tmp.name


# ══════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════

def limpiar_precio(texto: str) -> int:
    import re
    nums = re.sub(r"[^\d]", "", texto)
    return int(nums) if nums else 0


def extraer_atributo(card, claves: list, selector_fallback: str) -> str:
    for clave in claves:
        val = card.get(f"data-{clave}", "")
        if val:
            return val
    elems = card.select(selector_fallback)
    return elems[0].get_text(strip=True) if elems else ""


# ══════════════════════════════════════════════════════════════
#  SCRAPING AUTOSCOUT24
# ══════════════════════════════════════════════════════════════

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}


def construir_url(pagina: int = 1) -> str:
    b = CFG["busqueda"]
    base   = f"https://www.autoscout24.{b['pais']}/lst/{b['marca']}/{b['modelo']}"
    params = (
        f"?sort=age&desc=0"
        f"&priceto={b['precio_max']}"
        f"&fregfrom={b['anio_min']}"
        f"&kmto={b['km_max']}"
        f"&page={pagina}"
    )
    return base + params


def extraer_datos_anuncio(card) -> dict | None:
    try:
        anuncio_id = card.get("id", "") or card.get("data-id", "")
        if not anuncio_id:
            link = card.select_one("a[href*='/anuncio/'], a[href*='/annonce/'], a[href*='/offerte/']")
            if link:
                href = link.get("href", "")
                anuncio_id = href.split("/")[-1].split("?")[0]

        link_elem   = card.select_one("a[href]")
        pais        = CFG["busqueda"]["pais"]
        href        = link_elem.get("href", "") if link_elem else ""
        url_anuncio = href if href.startswith("http") else f"https://www.autoscout24.{pais}{href}"

        titulo_elem = card.select_one("h2, .cldt-summary-headline, [data-testid='title']")
        titulo      = titulo_elem.get_text(strip=True) if titulo_elem else "Sin título"

        precio_elem = card.select_one(
            ".cldt-price, [data-testid='price'], .Price_price__wEzIV, "
            "span[class*='price'], div[class*='Price']"
        )
        precio = limpiar_precio(precio_elem.get_text(strip=True) if precio_elem else "0")

        anio        = extraer_atributo(card, ["year", "anio"], "span.cldt-summary-version-block")
        km          = extraer_atributo(card, ["mileage", "km"], "span.cldt-summary-version-block")

        comb_elem   = card.select_one("[data-testid='fuel-type'], .cldt-summary-tags span")
        combustible = comb_elem.get_text(strip=True) if comb_elem else ""

        ubic_elem   = card.select_one(
            "[data-testid='location'], .cldt-summary-seller-contact-location, "
            "span[class*='location'], div[class*='Location']"
        )
        ubicacion = ubic_elem.get_text(strip=True) if ubic_elem else ""

        if not anuncio_id or precio == 0:
            log.debug(f"  Descartado (sin ID o precio=0): '{titulo[:40]}'")
            return None

        return {
            "id":              str(anuncio_id),
            "titulo":          titulo,
            "precio":          precio,
            "anio":            anio,
            "km":              km,
            "combustible":     combustible,
            "ubicacion":       ubicacion,
            "url":             url_anuncio,
            "fecha_detectado": date.today().isoformat(),
        }
    except Exception as e:
        log.warning(f"Error parseando card: {e}")
        return None


def obtener_anuncios_pagina(url: str) -> list[dict]:
    log.debug(f"GET {url}")
    anuncios = []
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        log.debug(f"HTTP {resp.status_code} — {len(resp.content)} bytes")
        resp.raise_for_status()

        soup  = BeautifulSoup(resp.text, "html.parser")
        cards = soup.select("article.cldt-summary-full-item, [data-testid='result-item']")
        if not cards:
            cards = soup.select("article[id^='listing']")

        log.debug(f"Cards encontradas: {len(cards)}")
        for i, card in enumerate(cards):
            anuncio = extraer_datos_anuncio(card)
            if anuncio:
                anuncios.append(anuncio)
                log.debug(f"  [{i+1}] {anuncio['titulo'][:50]} — {anuncio['precio']:,}€")

        log.info(f"Página OK → {len(anuncios)} anuncios de {len(cards)} cards")

    except requests.exceptions.Timeout:
        log.error(f"Timeout al conectar con {url}")
    except requests.exceptions.HTTPError as e:
        log.error(f"HTTP Error {e.response.status_code} en {url}")
    except Exception as e:
        log.error(f"Error inesperado scrapeando {url}: {e}", exc_info=True)

    return anuncios


def scrape_todos_los_anuncios() -> list[dict]:
    b        = CFG["busqueda"]
    max_pags = b.get("max_paginas", 5)
    pausa    = b.get("pausa_entre_paginas_seg", 3)
    todos    = []

    log.info(
        f"Iniciando scraping — {b['marca']} {b['modelo']} | "
        f"hasta {b['precio_max']:,}€ | desde {b['anio_min']} | "
        f"hasta {b['km_max']:,}km | país={b['pais']}"
    )

    for pagina in range(1, max_pags + 1):
        url      = construir_url(pagina)
        log.info(f"── Página {pagina}/{max_pags}")
        anuncios = obtener_anuncios_pagina(url)

        if not anuncios:
            log.info(f"Página {pagina} sin resultados — fin del scraping")
            break

        todos.extend(anuncios)
        if pagina < max_pags:
            log.debug(f"Esperando {pausa}s...")
            time.sleep(pausa)

    vistos, unicos = set(), []
    for a in todos:
        if a["id"] not in vistos:
            vistos.add(a["id"])
            unicos.append(a)

    duplicados = len(todos) - len(unicos)
    if duplicados:
        log.info(f"Eliminados {duplicados} duplicados")

    log.info(f"Scraping completado — {len(unicos)} anuncios únicos")
    return unicos


# ══════════════════════════════════════════════════════════════
#  GOOGLE SHEETS — conexión compartida
# ══════════════════════════════════════════════════════════════

_spreadsheet = None   # caché de la hoja abierta
_creds_file  = None   # ruta al archivo temporal de credenciales


def conectar_sheets():
    global _spreadsheet, _creds_file
    if _spreadsheet:
        return _spreadsheet

    if not _creds_file:
        _creds_file = obtener_credenciales_google()

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds  = Credentials.from_service_account_file(_creds_file, scopes=scopes)
    client = gspread.authorize(creds)

    nombre = CFG["google_sheets"]["nombre_hoja"]
    try:
        _spreadsheet = client.open(nombre)
        log.debug(f"Hoja '{nombre}' encontrada")
    except gspread.SpreadsheetNotFound:
        _spreadsheet = client.create(nombre)
        _spreadsheet.share(CFG["email"]["destino"], perm_type="user", role="writer")
        log.info(f"Hoja '{nombre}' creada y compartida con {CFG['email']['destino']}")

    return _spreadsheet


# ══════════════════════════════════════════════════════════════
#  ESTADO — guardado en Google Sheets (Railway es efímero)
# ══════════════════════════════════════════════════════════════

HOJA_ESTADO = "_Estado_"


def cargar_estado() -> dict:
    """Lee el estado previo desde la hoja oculta _Estado_ de Google Sheets."""
    try:
        spreadsheet = conectar_sheets()
        try:
            ws = spreadsheet.worksheet(HOJA_ESTADO)
        except gspread.WorksheetNotFound:
            log.info("Hoja de estado no existe — primera ejecución")
            return {}

        contenido = ws.acell("A1").value
        if not contenido:
            log.info("Hoja de estado vacía — primera ejecución")
            return {}

        estado = json.loads(contenido)
        log.info(f"Estado cargado desde Sheets — {len(estado)} anuncios conocidos")
        return estado

    except Exception as e:
        log.error(f"Error cargando estado desde Sheets: {e}", exc_info=True)
        return {}


def guardar_estado(anuncios: list[dict]):
    """Guarda el estado actual en la hoja oculta _Estado_ de Google Sheets."""
    try:
        spreadsheet = conectar_sheets()
        try:
            ws = spreadsheet.worksheet(HOJA_ESTADO)
        except gspread.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(HOJA_ESTADO, rows=1, cols=1)
            log.debug("Hoja de estado creada en Sheets")

        estado = {a["id"]: a for a in anuncios}
        ws.update("A1", [[json.dumps(estado, ensure_ascii=False)]])
        log.info(f"Estado guardado en Sheets — {len(estado)} anuncios")

    except Exception as e:
        log.error(f"Error guardando estado en Sheets: {e}", exc_info=True)


def detectar_cambios(actuales: list[dict], anteriores: dict):
    nuevos, bajadas = [], []

    for anuncio in actuales:
        aid = anuncio["id"]
        if aid not in anteriores:
            log.debug(f"NUEVO: {anuncio['titulo'][:50]} — {anuncio['precio']:,}€")
            nuevos.append(anuncio)
        else:
            precio_ant = anteriores[aid].get("precio", 0)
            precio_act = anuncio.get("precio", 0)
            if precio_act > 0 and precio_ant > 0 and precio_act < precio_ant:
                diferencia  = precio_ant - precio_act
                porcentaje  = round(diferencia / precio_ant * 100, 1)
                anuncio["precio_anterior"]   = precio_ant
                anuncio["diferencia_precio"] = diferencia
                anuncio["porcentaje_bajada"] = porcentaje
                log.debug(
                    f"BAJADA: {anuncio['titulo'][:40]} — "
                    f"{precio_ant:,}€ → {precio_act:,}€ (-{porcentaje}%)"
                )
                bajadas.append(anuncio)

    log.info(f"Cambios detectados — {len(nuevos)} nuevos | {len(bajadas)} bajadas")
    return nuevos, bajadas


# ══════════════════════════════════════════════════════════════
#  GOOGLE SHEETS — datos de anuncios
# ══════════════════════════════════════════════════════════════

CABECERAS = [
    "ID", "Título", "Precio (€)", "Precio Anterior (€)", "Bajada (€)", "Bajada (%)",
    "Año", "Km", "Combustible", "Ubicación", "Estado", "Fecha Detectado", "URL"
]


def _fila_anuncio(a: dict, ids_nuevos: set, ids_bajadas: set) -> list:
    aid = a["id"]
    if aid in ids_bajadas:
        estado = f"⬇ BAJADA {a.get('porcentaje_bajada')}%"
    elif aid in ids_nuevos:
        estado = "🆕 NUEVO"
    else:
        estado = "Conocido"
    return [
        a.get("id",""),              a.get("titulo",""),
        a.get("precio",0),           a.get("precio_anterior",""),
        a.get("diferencia_precio",""), a.get("porcentaje_bajada",""),
        a.get("anio",""),            a.get("km",""),
        a.get("combustible",""),     a.get("ubicacion",""),
        estado,                      a.get("fecha_detectado",""),
        a.get("url",""),
    ]


def actualizar_sheets(anuncios: list[dict], nuevos: list[dict], bajadas: list[dict]) -> str:
    log.info("Actualizando Google Sheets...")
    try:
        spreadsheet = conectar_sheets()
        ids_nuevos  = {a["id"] for a in nuevos}
        ids_bajadas = {a["id"] for a in bajadas}

        # ── Hoja principal ──────────────────────────────────
        try:
            ws = spreadsheet.worksheet("Todos los anuncios")
        except gspread.WorksheetNotFound:
            ws = spreadsheet.add_worksheet("Todos los anuncios", rows=2000, cols=15)

        if not ws.row_values(1):
            ws.append_row(CABECERAS)

        ids_existentes = {
            fila[0]: i
            for i, fila in enumerate(ws.get_all_values()[1:], start=2)
            if fila
        }

        nuevas_filas, actualizadas = [], 0
        for a in anuncios:
            fila = _fila_anuncio(a, ids_nuevos, ids_bajadas)
            if a["id"] in ids_existentes:
                n = ids_existentes[a["id"]]
                ws.update(f"A{n}:M{n}", [fila])
                actualizadas += 1
            else:
                nuevas_filas.append(fila)

        if nuevas_filas:
            ws.append_rows(nuevas_filas)

        log.info(
            f"Hoja principal — {len(nuevas_filas)} filas nuevas | {actualizadas} actualizadas"
        )

        # ── Hoja del día ────────────────────────────────────
        nombre_hoy = f"Día {date.today().isoformat()}"
        try:
            ws_hoy = spreadsheet.worksheet(nombre_hoy)
            ws_hoy.clear()
        except gspread.WorksheetNotFound:
            ws_hoy = spreadsheet.add_worksheet(nombre_hoy, rows=200, cols=15)

        filas_hoy = [CABECERAS] + [
            _fila_anuncio(a, ids_nuevos, ids_bajadas)
            for a in (nuevos + bajadas)
        ]
        if len(filas_hoy) > 1:
            ws_hoy.update("A1", filas_hoy)
            log.info(f"Hoja '{nombre_hoy}' → {len(filas_hoy)-1} entradas")
        else:
            log.info("Sin cambios hoy — hoja del día queda vacía")

        log.info(f"Sheets OK: {spreadsheet.url}")
        return spreadsheet.url

    except Exception as e:
        log.error(f"Error actualizando Sheets: {e}", exc_info=True)
        return ""


# ══════════════════════════════════════════════════════════════
#  EMAIL
# ══════════════════════════════════════════════════════════════

def _card_html(a: dict, tipo: str) -> str:
    precio_str = f"{a.get('precio', 0):,} €".replace(",", ".")
    if tipo == "bajada":
        precio_ant = f"{a.get('precio_anterior', 0):,} €".replace(",", ".")
        diferencia = f"{a.get('diferencia_precio', 0):,} €".replace(",", ".")
        pct        = a.get("porcentaje_bajada", 0)
        extra = (
            f'<div style="color:#d32f2f;font-weight:bold;margin-top:4px;">'
            f'⬇ Antes: {precio_ant} → Ahora: {precio_str} (-{diferencia} / -{pct}%)</div>'
        )
    else:
        extra = f'<div style="color:#1976d2;font-weight:bold;margin-top:4px;">💶 {precio_str}</div>'

    return f"""
    <div style="border:1px solid #e0e0e0;border-radius:8px;padding:14px;
                margin-bottom:12px;background:#fafafa;">
        <div style="font-weight:600;font-size:15px;color:#212121;">{a.get('titulo','')}</div>
        {extra}
        <div style="color:#616161;font-size:13px;margin-top:6px;">
            📅 {a.get('anio','?')} &nbsp;|&nbsp;
            🛣️ {a.get('km','?')} &nbsp;|&nbsp;
            ⛽ {a.get('combustible','?')} &nbsp;|&nbsp;
            📍 {a.get('ubicacion','?')}
        </div>
        <div style="margin-top:8px;">
            <a href="{a.get('url','#')}"
               style="background:#1565c0;color:white;padding:6px 14px;
                      border-radius:4px;text-decoration:none;font-size:13px;">
                Ver anuncio →
            </a>
        </div>
    </div>"""


def enviar_email(nuevos: list[dict], bajadas: list[dict], url_sheets: str):
    if not nuevos and not bajadas:
        log.info("Sin cambios — no se envía email")
        return

    b     = CFG["busqueda"]
    marca = b["marca"].title()
    modelo= b["modelo"].title()
    hoy   = date.today().strftime("%d/%m/%Y")
    em    = CFG["email"]

    log.info(f"Preparando email — {len(nuevos)} nuevos | {len(bajadas)} bajadas")

    sec_nuevos = ""
    if nuevos:
        items = "".join(_card_html(a, "nuevo") for a in nuevos)
        sec_nuevos = (
            f'<h2 style="color:#1565c0;border-bottom:2px solid #1565c0;padding-bottom:6px;">'
            f'🆕 Nuevos anuncios ({len(nuevos)})</h2>{items}'
        )

    sec_bajadas = ""
    if bajadas:
        items = "".join(_card_html(a, "bajada") for a in bajadas)
        sec_bajadas = (
            f'<h2 style="color:#c62828;border-bottom:2px solid #c62828;padding-bottom:6px;">'
            f'⬇️ Bajadas de precio ({len(bajadas)})</h2>{items}'
        )

    link_sheets = (
        f'<p><a href="{url_sheets}" style="background:#388e3c;color:white;'
        f'padding:8px 18px;border-radius:4px;text-decoration:none;">'
        f'📊 Ver hoja de cálculo</a></p>'
    ) if url_sheets else ""

    html = f"""
    <html><body style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;padding:20px;">
        <div style="background:#1565c0;color:white;padding:16px 24px;border-radius:8px 8px 0 0;">
            <h1 style="margin:0;font-size:20px;">🚗 AutoScout24 — {marca} {modelo}</h1>
            <div style="opacity:.85;font-size:13px;margin-top:4px;">
                {hoy} &nbsp;|&nbsp; {len(nuevos)} nuevos &nbsp;·&nbsp; {len(bajadas)} bajadas
            </div>
        </div>
        <div style="background:white;border:1px solid #e0e0e0;border-top:none;
                    padding:20px;border-radius:0 0 8px 8px;">
            {sec_nuevos}{sec_bajadas}{link_sheets}
            <hr style="margin-top:24px;border:none;border-top:1px solid #eee;">
            <p style="color:#9e9e9e;font-size:11px;">
                {marca} {modelo} | Máx {b['precio_max']:,}€ |
                Desde {b['anio_min']} | Máx {b['km_max']:,}km
            </p>
        </div>
    </body></html>"""

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = (
            f"🚗 AutoScout24 {marca} {modelo} — "
            f"{len(nuevos)} nuevos, {len(bajadas)} bajadas — {hoy}"
        )
        msg["From"] = em["origen"]
        msg["To"]   = em["destino"]
        msg.attach(MIMEText(html, "html"))

        with smtplib.SMTP_SSL(em["servidor_smtp"], em["puerto_smtp"]) as server:
            server.login(em["origen"], em["password"])
            server.sendmail(em["origen"], em["destino"], msg.as_string())

        log.info(f"Email enviado a {em['destino']}")

    except smtplib.SMTPAuthenticationError:
        log.error(
            "Error autenticación SMTP — usa la contraseña de APLICACIÓN de Gmail "
            "(16 caracteres), no tu contraseña normal"
        )
    except Exception as e:
        log.error(f"Error enviando email: {e}", exc_info=True)


# ══════════════════════════════════════════════════════════════
#  LIMPIEZA DE ARCHIVOS TEMPORALES
# ══════════════════════════════════════════════════════════════

def limpiar_temporales():
    global _creds_file
    if _creds_file and os.path.exists(_creds_file):
        try:
            os.unlink(_creds_file)
            log.debug(f"Archivo temporal eliminado: {_creds_file}")
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════

def main():
    inicio    = datetime.now()
    separador = "═" * 60

    log.info(separador)
    log.info(f"AutoScout24 Agent — INICIO {inicio.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info(separador)

    try:
        log.info("PASO 1/5 — Cargando estado anterior (desde Google Sheets)")
        estado_anterior = cargar_estado()

        log.info("PASO 2/5 — Scraping AutoScout24")
        anuncios_actuales = scrape_todos_los_anuncios()

        if not anuncios_actuales:
            log.warning("No se obtuvieron anuncios. Posible bloqueo temporal.")
            return

        log.info("PASO 3/5 — Detectando cambios")
        nuevos, bajadas = detectar_cambios(anuncios_actuales, estado_anterior)

        log.info("PASO 4/5 — Actualizando Google Sheets")
        url_sheets = actualizar_sheets(anuncios_actuales, nuevos, bajadas)

        log.info("PASO 5/5 — Enviando email de resumen")
        enviar_email(nuevos, bajadas, url_sheets)

        guardar_estado(anuncios_actuales)

    except Exception as e:
        log.critical(f"Error crítico: {e}", exc_info=True)
        raise
    finally:
        limpiar_temporales()
        duracion = (datetime.now() - inicio).total_seconds()
        log.info(separador)
        log.info(f"AutoScout24 Agent — FIN | Duración: {duracion:.1f}s")
        log.info(separador)


if __name__ == "__main__":
    main()
