"""
AutoScout24 Agent — Versión GitHub Actions
- Config en config.json (múltiples búsquedas, filtros opcionales)
- Credenciales Google: variable GOOGLE_CREDENTIALS_JSON (base64)
- Contraseña email:    variable EMAIL_PASSWORD
- Estado entre ejecuciones: Google Sheets (pestaña _Estado_)
- Scraping: extracción desde JSON de Next.js + fallback CSS mejorado
"""

import os, sys, json, time, base64, re, smtplib, logging, tempfile
import requests
from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials


# ══════════════════════════════════════════════════════════════
#  CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def ruta(f): return os.path.join(SCRIPT_DIR, f)

def cargar_config() -> dict:
    path = ruta("config.json")
    if not os.path.exists(path):
        print(f"[FATAL] No se encuentra config.json en {path}"); sys.exit(1)
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    cfg["email"]["password"] = os.environ.get("EMAIL_PASSWORD", "")
    if not cfg["email"]["password"]:
        print("[FATAL] Variable de entorno EMAIL_PASSWORD no definida."); sys.exit(1)
    return cfg

CFG = cargar_config()


# ══════════════════════════════════════════════════════════════
#  LOGGING
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
#  CREDENCIALES GOOGLE
# ══════════════════════════════════════════════════════════════

_creds_file = None

def obtener_credenciales_google() -> str:
    global _creds_file
    if _creds_file and os.path.exists(_creds_file):
        return _creds_file
    b64 = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if not b64:
        log.critical("Variable GOOGLE_CREDENTIALS_JSON no definida."); sys.exit(1)
    try:
        json_bytes = base64.b64decode(b64)
        json.loads(json_bytes)
    except Exception as e:
        log.critical(f"GOOGLE_CREDENTIALS_JSON inválido: {e}"); sys.exit(1)
    tmp = tempfile.NamedTemporaryFile(suffix=".json", delete=False, mode="wb")
    tmp.write(json_bytes); tmp.close()
    _creds_file = tmp.name
    log.debug(f"Credenciales Google en temporal: {_creds_file}")
    return _creds_file

def limpiar_temporales():
    global _creds_file
    if _creds_file and os.path.exists(_creds_file):
        try: os.unlink(_creds_file)
        except: pass


# ══════════════════════════════════════════════════════════════
#  SCRAPING — HELPERS DE PARSEO DE CAMPOS
# ══════════════════════════════════════════════════════════════

def limpiar_numero(texto: str) -> int:
    """Convierte '21.350 €' o '45.900 km' en 21350 / 45900."""
    nums = re.sub(r"[^\d]", "", str(texto))
    return int(nums) if nums else 0

def es_precio_razonable(n: int) -> bool:
    return 500 < n < 500_000

def es_km_razonable(n: int) -> bool:
    return 0 <= n < 2_000_000


# ══════════════════════════════════════════════════════════════
#  SCRAPING — EXTRACCIÓN DESDE __NEXT_DATA__ (método principal)
# ══════════════════════════════════════════════════════════════

def _buscar_listings_en_json(obj, depth=0):
    """Busca recursivamente listas de anuncios en el JSON de Next.js."""
    if depth > 12:
        return None
    if isinstance(obj, list) and len(obj) >= 1:
        primer = obj[0]
        if isinstance(primer, dict) and any(
            k in primer for k in ("id", "price", "url", "vehicleDetails", "vehicle")
        ):
            return obj
    if isinstance(obj, dict):
        for key in ("listings", "entities", "results", "items", "hits", "data"):
            if key in obj:
                r = _buscar_listings_en_json(obj[key], depth + 1)
                if r:
                    return r
        for v in obj.values():
            if isinstance(v, (dict, list)):
                r = _buscar_listings_en_json(v, depth + 1)
                if r:
                    return r
    return None


def _parsear_listing_json(item: dict, pais: str) -> dict | None:
    """Convierte un anuncio del JSON de Next.js en nuestro diccionario."""
    try:
        aid = str(item.get("id") or item.get("guid") or "")
        if not aid:
            return None

        # URL
        url_raw = (item.get("url") or item.get("link") or
                   item.get("detailPageUrl") or "")
        if url_raw.startswith("http"):
            url = url_raw
        elif url_raw:
            url = f"https://www.autoscout24.{pais}{url_raw}"
        else:
            url = f"https://www.autoscout24.{pais}/anuncios/{aid}"

        # Título
        titulo = ""
        for key in ("title", "name", "headline"):
            if item.get(key):
                titulo = str(item[key]).strip(); break
        if not titulo:
            v = item.get("vehicle") or {}
            make  = (v.get("make") or {}).get("name", "")
            model = (v.get("model") or {}).get("name", "")
            desc  = v.get("description", "") or v.get("modelVersionInput", "")
            titulo = f"{make} {model} {desc}".strip() or "Sin título"

        # Precio
        precio = 0
        p = item.get("price") or item.get("pricing") or {}
        if isinstance(p, dict):
            precio = int(p.get("value") or p.get("amount") or p.get("gross") or 0)
        elif isinstance(p, (int, float)):
            precio = int(p)
        if not es_precio_razonable(precio):
            precio = 0

        # Año
        anio = ""
        vd = item.get("vehicleDetails") or item.get("vehicle") or {}
        for key in ("firstRegistration", "firstRegistrationDate", "yearOfProduction", "year"):
            val = vd.get(key) or item.get(key) or ""
            if val:
                m = re.search(r"(19|20)\d{2}", str(val))
                if m:
                    anio = m.group(); break

        # Km
        km_val = 0
        for key in ("mileage", "kilometer", "km", "odometer"):
            val = vd.get(key) or item.get(key) or 0
            if val:
                km_val = limpiar_numero(str(val)); break
        km = f"{km_val:,} km".replace(",", ".") if km_val else ""

        # Combustible
        combustible = ""
        for key in ("fuelType", "fuel", "fuelTypeDetails"):
            val = vd.get(key) or item.get(key) or ""
            if isinstance(val, dict):
                val = val.get("name") or val.get("value") or ""
            if val:
                combustible = str(val).strip(); break

        # Ubicación
        ubicacion = ""
        seller = item.get("seller") or item.get("location") or {}
        loc = seller.get("location") or seller if isinstance(seller, dict) else {}
        for key in ("city", "zip", "region", "countryName", "address"):
            val = loc.get(key) or ""
            if val:
                ubicacion = str(val).strip(); break
        if not ubicacion:
            ubicacion = (item.get("location") or {}).get("city", "") if isinstance(item.get("location"), dict) else ""

        if not aid or precio == 0:
            return None

        return {
            "id":              aid,
            "titulo":          titulo,
            "precio":          precio,
            "anio":            anio,
            "km":              km,
            "combustible":     combustible,
            "ubicacion":       ubicacion,
            "url":             url,
            "fecha_detectado": date.today().isoformat(),
        }
    except Exception as e:
        log.debug(f"Error parseando JSON item: {e}")
        return None


def extraer_desde_next_data(soup, pais: str) -> list[dict]:
    """Extrae anuncios del JSON embebido de Next.js (__NEXT_DATA__)."""
    script = soup.find("script", {"id": "__NEXT_DATA__"})
    if not script or not script.string:
        return []
    try:
        data = json.loads(script.string)
    except json.JSONDecodeError:
        return []

    listings_raw = _buscar_listings_en_json(data)
    if not listings_raw:
        log.debug("No se encontraron listings en __NEXT_DATA__")
        return []

    anuncios = []
    for item in listings_raw:
        a = _parsear_listing_json(item, pais)
        if a:
            anuncios.append(a)

    log.debug(f"__NEXT_DATA__: {len(anuncios)} anuncios parseados de {len(listings_raw)} items")
    return anuncios


# ══════════════════════════════════════════════════════════════
#  SCRAPING — FALLBACK CSS MEJORADO
# ══════════════════════════════════════════════════════════════

def _extraer_precio_css(card) -> int:
    """Busca el precio evitando confundirlo con km u otros números."""
    # 1. Selectores específicos por prioridad
    selectores = [
        "[data-testid='price']",
        "[data-testid='regular-price']",
        ".cldt-price",
        "[class*='Price_price']",
        "[class*='price__']",
        "strong[class*='Price']",
        "p[class*='Price']",
    ]
    for sel in selectores:
        el = card.select_one(sel)
        if el:
            txt = el.get_text(strip=True)
            if "€" in txt:
                n = limpiar_numero(txt)
                if es_precio_razonable(n):
                    return n

    # 2. Buscar texto con € filtrando km/h y valores raros
    for el in card.find_all(string=re.compile(r"[\d\.]+\s*€")):
        txt = str(el).strip()
        if "km" in txt.lower() or "km/h" in txt.lower():
            continue
        n = limpiar_numero(txt)
        if es_precio_razonable(n):
            return n

    return 0


def _extraer_km_css(card) -> str:
    """Extrae los km del anuncio."""
    selectores = [
        "[data-testid='mileage']",
        "[data-testid='km']",
        "[class*='mileage']",
        "[class*='Mileage']",
    ]
    for sel in selectores:
        el = card.select_one(sel)
        if el:
            txt = el.get_text(strip=True)
            n = limpiar_numero(txt)
            if es_km_razonable(n) and n > 0:
                return f"{n:,} km".replace(",", ".")

    for el in card.find_all(string=re.compile(r"[\d\.]+\s*km", re.IGNORECASE)):
        txt = str(el).strip()
        if "km/h" in txt.lower():
            continue
        n = limpiar_numero(txt)
        if es_km_razonable(n) and n > 0:
            return f"{n:,} km".replace(",", ".")

    return ""


def _extraer_anio_css(card) -> str:
    """Extrae el año de primera matriculación."""
    selectores = [
        "[data-testid='first-registration']",
        "[data-testid='year']",
        "[class*='firstReg']",
        "[class*='FirstReg']",
    ]
    for sel in selectores:
        el = card.select_one(sel)
        if el:
            m = re.search(r"(19|20)\d{2}", el.get_text())
            if m: return m.group()

    for el in card.find_all(string=re.compile(r"\b(19|20)\d{2}\b")):
        m = re.search(r"\b(20[012]\d|199\d)\b", str(el))
        if m: return m.group()

    return ""


def _extraer_combustible_css(card) -> str:
    selectores = [
        "[data-testid='fuel-type']",
        "[data-testid='fuel']",
        "[class*='FuelType']",
        "[class*='fuel']",
    ]
    for sel in selectores:
        el = card.select_one(sel)
        if el:
            txt = el.get_text(strip=True)
            if txt and len(txt) < 30:
                return txt
    return ""


def _extraer_ubicacion_css(card) -> str:
    selectores = [
        "[data-testid='location']",
        "[data-testid='seller-location']",
        ".cldt-summary-seller-contact-location",
        "[class*='Location']",
        "[class*='location']",
    ]
    for sel in selectores:
        el = card.select_one(sel)
        if el:
            txt = el.get_text(strip=True)
            if txt and len(txt) < 60:
                return txt
    return ""


def _extraer_url_css(card, pais: str) -> str:
    """Busca el enlace al anuncio individual (no a filtros ni paginación)."""
    # AutoScout24 listings URL pattern: /anuncios/... o /annonce/... etc
    patrones = ["/anuncios/", "/annonce/", "/annuncio/", "/offerte/", "/aanbod/"]
    for a in card.find_all("a", href=True):
        href = a["href"]
        if any(p in href for p in patrones):
            if href.startswith("http"):
                return href
            return f"https://www.autoscout24.{pais}{href}"

    # Fallback: cualquier enlace con el ID del anuncio
    aid = card.get("id", "")
    if aid:
        for a in card.find_all("a", href=True):
            href = a["href"]
            if aid in href:
                if href.startswith("http"):
                    return href
                return f"https://www.autoscout24.{pais}{href}"

    # Último recurso: primer enlace de la card
    a = card.select_one("a[href]")
    if a:
        href = a["href"]
        if href.startswith("http"):
            return href
        return f"https://www.autoscout24.{pais}{href}"

    return ""


def extraer_desde_css(soup, pais: str) -> list[dict]:
    """Extrae anuncios mediante selectores CSS como fallback."""
    cards = (
        soup.select("article[id^='listing']") or
        soup.select("article.cldt-summary-full-item") or
        soup.select("[data-testid='result-item']") or
        soup.select("article[id]")
    )
    log.debug(f"CSS fallback: {len(cards)} cards encontradas")

    anuncios = []
    for card in cards:
        try:
            aid = card.get("id", "") or card.get("data-id", "")
            if not aid:
                continue

            titulo_el = card.select_one("h2, [data-testid='title'], .cldt-summary-headline")
            titulo = titulo_el.get_text(strip=True) if titulo_el else "Sin título"

            precio     = _extraer_precio_css(card)
            km         = _extraer_km_css(card)
            anio       = _extraer_anio_css(card)
            combustible= _extraer_combustible_css(card)
            ubicacion  = _extraer_ubicacion_css(card)
            url        = _extraer_url_css(card, pais)

            if not aid or precio == 0:
                continue

            anuncios.append({
                "id":              str(aid),
                "titulo":          titulo,
                "precio":          precio,
                "anio":            anio,
                "km":              km,
                "combustible":     combustible,
                "ubicacion":       ubicacion,
                "url":             url,
                "fecha_detectado": date.today().isoformat(),
            })
        except Exception as e:
            log.debug(f"Error CSS card: {e}")

    return anuncios


# ══════════════════════════════════════════════════════════════
#  SCRAPING — PRINCIPAL
# ══════════════════════════════════════════════════════════════

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
}


def construir_url(b: dict, pagina: int = 1) -> str:
    """Construye la URL de búsqueda con solo los filtros que tengan valor."""
    marca  = b["marca"].lower().replace(" ", "-")
    modelo = b["modelo"].lower().replace(" ", "-")
    pais   = b.get("pais", "es")
    base   = f"https://www.autoscout24.{pais}/lst/{marca}/{modelo}"

    params = [f"sort=age", "desc=0", f"page={pagina}"]

    def add(key, val):
        if val is not None and val != "" and val != 0:
            params.append(f"{key}={val}")

    add("pricefrom", b.get("precio_min"))
    add("priceto",   b.get("precio_max"))
    add("fregfrom",  b.get("anio_min"))
    add("fregto",    b.get("anio_max"))
    add("kmfrom",    b.get("km_min"))
    add("kmto",      b.get("km_max"))

    return base + "?" + "&".join(params)


def obtener_anuncios_pagina(url: str, pais: str) -> list[dict]:
    log.debug(f"GET {url}")
    try:
        resp = requests.get(url, headers=HEADERS, timeout=25)
        log.debug(f"HTTP {resp.status_code} — {len(resp.content):,} bytes")
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")

        # Intentar __NEXT_DATA__ primero
        anuncios = extraer_desde_next_data(soup, pais)
        if anuncios:
            log.info(f"  JSON: {len(anuncios)} anuncios")
            return anuncios

        # Fallback CSS
        anuncios = extraer_desde_css(soup, pais)
        log.info(f"  CSS:  {len(anuncios)} anuncios")
        return anuncios

    except requests.exceptions.Timeout:
        log.error(f"Timeout: {url}")
    except requests.exceptions.HTTPError as e:
        log.error(f"HTTP {e.response.status_code}: {url}")
    except Exception as e:
        log.error(f"Error scraping {url}: {e}", exc_info=True)
    return []


def scrape_busqueda(b: dict) -> list[dict]:
    """Ejecuta el scraping completo de una búsqueda."""
    max_pags = b.get("max_paginas", 5)
    pausa    = b.get("pausa_seg", 3)
    pais     = b.get("pais", "es")
    todos    = []

    filtros = []
    for k, label in [("precio_min","desde"), ("precio_max","hasta"), ("anio_min","año desde"),
                     ("anio_max","año hasta"), ("km_min","km desde"), ("km_max","km hasta")]:
        if b.get(k) is not None:
            filtros.append(f"{label}:{b[k]}")

    log.info(f"Buscando: {b['marca']} {b['modelo']} | {b['pais'].upper()} | {' | '.join(filtros) or 'sin filtros extra'}")

    for pagina in range(1, max_pags + 1):
        url      = construir_url(b, pagina)
        log.info(f"── Página {pagina}/{max_pags}")
        anuncios = obtener_anuncios_pagina(url, pais)
        if not anuncios:
            log.info("  Página vacía — fin del scraping")
            break
        todos.extend(anuncios)
        if pagina < max_pags:
            time.sleep(pausa)

    # Deduplicar por ID
    vistos, unicos = set(), []
    for a in todos:
        if a["id"] not in vistos:
            vistos.add(a["id"])
            unicos.append(a)

    log.info(f"Scraping completado — {len(unicos)} anuncios únicos ({len(todos)-len(unicos)} duplicados eliminados)")
    return unicos


# ══════════════════════════════════════════════════════════════
#  GOOGLE SHEETS — CONEXIÓN
# ══════════════════════════════════════════════════════════════

_spreadsheet = None

def conectar_sheets():
    global _spreadsheet
    if _spreadsheet:
        return _spreadsheet

    creds = Credentials.from_service_account_file(
        obtener_credenciales_google(),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
    )
    client = gspread.authorize(creds)
    nombre = CFG["google_sheets"]["nombre_hoja"]

    try:
        _spreadsheet = client.open(nombre)
        log.debug(f"Hoja '{nombre}' encontrada")
    except gspread.SpreadsheetNotFound:
        _spreadsheet = client.create(nombre)
        _spreadsheet.share(CFG["email"]["destino"], perm_type="user", role="writer")
        log.info(f"Hoja '{nombre}' creada y compartida")

    return _spreadsheet


# ══════════════════════════════════════════════════════════════
#  ESTADO — GUARDADO EN GOOGLE SHEETS POR BÚSQUEDA
# ══════════════════════════════════════════════════════════════

def _nombre_hoja_estado(nombre_busqueda: str) -> str:
    safe = re.sub(r"[^\w\s-]", "", nombre_busqueda)[:20].strip()
    return f"_Estado_{safe}_"

def cargar_estado(nombre_busqueda: str) -> dict:
    try:
        sp  = conectar_sheets()
        nom = _nombre_hoja_estado(nombre_busqueda)
        try:
            ws = sp.worksheet(nom)
        except gspread.WorksheetNotFound:
            log.info(f"[{nombre_busqueda}] Sin estado previo — primera ejecución")
            return {}
        contenido = ws.acell("A1").value
        if not contenido:
            return {}
        estado = json.loads(contenido)
        log.info(f"[{nombre_busqueda}] Estado cargado — {len(estado)} anuncios conocidos")
        return estado
    except Exception as e:
        log.error(f"Error cargando estado [{nombre_busqueda}]: {e}", exc_info=True)
        return {}

def guardar_estado(nombre_busqueda: str, anuncios: list[dict]):
    try:
        sp  = conectar_sheets()
        nom = _nombre_hoja_estado(nombre_busqueda)
        try:
            ws = sp.worksheet(nom)
        except gspread.WorksheetNotFound:
            ws = sp.add_worksheet(nom, rows=1, cols=1)
        estado = {a["id"]: a for a in anuncios}
        ws.update("A1", [[json.dumps(estado, ensure_ascii=False)]])
        log.info(f"[{nombre_busqueda}] Estado guardado — {len(estado)} anuncios")
    except Exception as e:
        log.error(f"Error guardando estado [{nombre_busqueda}]: {e}", exc_info=True)


# ══════════════════════════════════════════════════════════════
#  DETECCIÓN DE CAMBIOS
# ══════════════════════════════════════════════════════════════

def detectar_cambios(actuales: list[dict], anteriores: dict):
    nuevos, bajadas = [], []
    for a in actuales:
        aid = a["id"]
        if aid not in anteriores:
            nuevos.append(a)
        else:
            p_ant = anteriores[aid].get("precio", 0)
            p_act = a.get("precio", 0)
            if p_act > 0 and p_ant > 0 and p_act < p_ant:
                diferencia = p_ant - p_act
                a["precio_anterior"]   = p_ant
                a["diferencia_precio"] = diferencia
                a["porcentaje_bajada"] = round(diferencia / p_ant * 100, 1)
                bajadas.append(a)
    log.info(f"  Cambios: {len(nuevos)} nuevos | {len(bajadas)} bajadas de precio")
    return nuevos, bajadas


# ══════════════════════════════════════════════════════════════
#  GOOGLE SHEETS — ACTUALIZACIÓN DE DATOS
# ══════════════════════════════════════════════════════════════

CABECERAS = [
    "ID", "Título", "Precio (€)", "Precio Anterior (€)", "Bajada (€)", "Bajada (%)",
    "Año", "Km", "Combustible", "Ubicación", "Estado", "Fecha Detectado", "URL"
]

def _fila(a: dict, ids_nuevos: set, ids_bajadas: set) -> list:
    aid = a["id"]
    if aid in ids_bajadas:
        estado = f"⬇ BAJADA {a.get('porcentaje_bajada')}%"
    elif aid in ids_nuevos:
        estado = "🆕 NUEVO"
    else:
        estado = "Conocido"
    return [
        a.get("id",""),                a.get("titulo",""),
        a.get("precio",0),             a.get("precio_anterior",""),
        a.get("diferencia_precio",""), a.get("porcentaje_bajada",""),
        a.get("anio",""),              a.get("km",""),
        a.get("combustible",""),       a.get("ubicacion",""),
        estado,                        a.get("fecha_detectado",""),
        a.get("url",""),
    ]

def actualizar_sheets_busqueda(nombre: str, anuncios: list, nuevos: list, bajadas: list) -> str:
    try:
        sp = conectar_sheets()
        ids_nuevos  = {a["id"] for a in nuevos}
        ids_bajadas = {a["id"] for a in bajadas}

        # ── Hoja de la búsqueda (todos los anuncios) ──────────
        nombre_ws = nombre[:50]
        try:
            ws = sp.worksheet(nombre_ws)
        except gspread.WorksheetNotFound:
            ws = sp.add_worksheet(nombre_ws, rows=2000, cols=15)
            log.info(f"  Hoja '{nombre_ws}' creada")

        if not ws.row_values(1):
            ws.append_row(CABECERAS)

        ids_existentes = {
            fila[0]: i
            for i, fila in enumerate(ws.get_all_values()[1:], start=2)
            if fila
        }

        nuevas_filas, actualizadas = [], 0
        for a in anuncios:
            f = _fila(a, ids_nuevos, ids_bajadas)
            if a["id"] in ids_existentes:
                n = ids_existentes[a["id"]]
                ws.update(f"A{n}:M{n}", [f])
                actualizadas += 1
            else:
                nuevas_filas.append(f)
        if nuevas_filas:
            ws.append_rows(nuevas_filas)
        log.info(f"  Sheets '{nombre_ws}': +{len(nuevas_filas)} nuevas, {actualizadas} actualizadas")

        # ── Hoja del día ──────────────────────────────────────
        nombre_hoy = f"{date.today().isoformat()} {nombre[:20]}"
        try:
            ws_hoy = sp.worksheet(nombre_hoy)
            ws_hoy.clear()
        except gspread.WorksheetNotFound:
            ws_hoy = sp.add_worksheet(nombre_hoy, rows=200, cols=15)

        filas_hoy = [CABECERAS] + [_fila(a, ids_nuevos, ids_bajadas) for a in (nuevos + bajadas)]
        if len(filas_hoy) > 1:
            ws_hoy.update("A1", filas_hoy)
            log.info(f"  Hoja del día: {len(filas_hoy)-1} entradas")

        return sp.url

    except Exception as e:
        log.error(f"Error Sheets [{nombre}]: {e}", exc_info=True)
        return ""


# ══════════════════════════════════════════════════════════════
#  EMAIL
# ══════════════════════════════════════════════════════════════

def _card_html(a: dict, tipo: str) -> str:
    precio_str = f"{a.get('precio', 0):,} €".replace(",", ".")
    if tipo == "bajada":
        p_ant = f"{a.get('precio_anterior', 0):,} €".replace(",", ".")
        dif   = f"{a.get('diferencia_precio', 0):,} €".replace(",", ".")
        pct   = a.get("porcentaje_bajada", 0)
        precio_html = (
            f'<div style="color:#d32f2f;font-weight:bold;margin-top:4px;">'
            f'⬇ Antes: {p_ant} → Ahora: {precio_str} (-{dif} / -{pct}%)</div>'
        )
    else:
        precio_html = f'<div style="color:#1565c0;font-weight:bold;margin-top:4px;">💶 {precio_str}</div>'

    anio        = a.get("anio", "") or "—"
    km          = a.get("km", "")  or "—"
    combustible = a.get("combustible", "") or "—"
    ubicacion   = a.get("ubicacion", "")  or "—"
    url         = a.get("url", "#")
    titulo      = a.get("titulo", "Sin título")

    return f"""
    <div style="border:1px solid #e0e0e0;border-radius:8px;padding:14px;
                margin-bottom:12px;background:#fafafa;">
        <div style="font-weight:600;font-size:15px;color:#212121;">{titulo}</div>
        {precio_html}
        <div style="color:#616161;font-size:13px;margin-top:6px;">
            📅 {anio} &nbsp;|&nbsp; 🛣️ {km} &nbsp;|&nbsp;
            ⛽ {combustible} &nbsp;|&nbsp; 📍 {ubicacion}
        </div>
        <div style="margin-top:8px;">
            <a href="{url}"
               style="background:#1565c0;color:white;padding:6px 14px;
                      border-radius:4px;text-decoration:none;font-size:13px;">
                Ver anuncio →
            </a>
        </div>
    </div>"""


def _seccion_busqueda_html(nombre: str, nuevos: list, bajadas: list) -> str:
    if not nuevos and not bajadas:
        return ""

    partes = [f"""
    <div style="margin:20px 0 8px;padding:10px 14px;background:#e8eaf6;
                border-left:4px solid #3949ab;border-radius:0 6px 6px 0;">
        <strong style="color:#1a237e;font-size:15px;">🔍 {nombre}</strong>
        <span style="color:#5c6bc0;font-size:12px;margin-left:10px;">
            {len(nuevos)} nuevos · {len(bajadas)} bajadas
        </span>
    </div>"""]

    if nuevos:
        partes.append(
            f'<h3 style="color:#1565c0;font-size:14px;margin:12px 0 8px;">'
            f'🆕 Nuevos anuncios ({len(nuevos)})</h3>'
        )
        partes.extend(_card_html(a, "nuevo") for a in nuevos)

    if bajadas:
        partes.append(
            f'<h3 style="color:#c62828;font-size:14px;margin:12px 0 8px;">'
            f'⬇️ Bajadas de precio ({len(bajadas)})</h3>'
        )
        partes.extend(_card_html(a, "bajada") for a in bajadas)

    return "\n".join(partes)


def enviar_email(resultados: list[dict], url_sheets: str):
    """
    resultados: lista de dicts con keys nombre, nuevos, bajadas
    """
    total_nuevos  = sum(len(r["nuevos"])  for r in resultados)
    total_bajadas = sum(len(r["bajadas"]) for r in resultados)

    if total_nuevos == 0 and total_bajadas == 0:
        log.info("Sin cambios hoy — no se envía email")
        return

    hoy = date.today().strftime("%d/%m/%Y")
    em  = CFG["email"]

    secciones = "\n".join(
        _seccion_busqueda_html(r["nombre"], r["nuevos"], r["bajadas"])
        for r in resultados
        if r["nuevos"] or r["bajadas"]
    )

    busquedas_activas = ", ".join(r["nombre"] for r in resultados)

    link_sheets = (
        f'<p style="margin-top:20px;"><a href="{url_sheets}" '
        f'style="background:#388e3c;color:white;padding:8px 18px;'
        f'border-radius:4px;text-decoration:none;">📊 Ver hoja de cálculo</a></p>'
    ) if url_sheets else ""

    html = f"""
    <html><body style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;padding:20px;">
        <div style="background:#1565c0;color:white;padding:16px 24px;border-radius:8px 8px 0 0;">
            <h1 style="margin:0;font-size:20px;">🚗 AutoScout24 Monitor</h1>
            <div style="opacity:.85;font-size:13px;margin-top:4px;">
                {hoy} &nbsp;|&nbsp; {total_nuevos} nuevos &nbsp;·&nbsp; {total_bajadas} bajadas
            </div>
        </div>
        <div style="background:white;border:1px solid #e0e0e0;border-top:none;
                    padding:20px;border-radius:0 0 8px 8px;">
            {secciones}
            {link_sheets}
            <hr style="margin-top:24px;border:none;border-top:1px solid #eee;">
            <p style="color:#9e9e9e;font-size:11px;">Búsquedas: {busquedas_activas}</p>
        </div>
    </body></html>"""

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = (
            f"🚗 AutoScout24 — {total_nuevos} nuevos, {total_bajadas} bajadas — {hoy}"
        )
        msg["From"] = em["origen"]
        msg["To"]   = em["destino"]
        msg.attach(MIMEText(html, "html"))

        with smtplib.SMTP_SSL(em["servidor_smtp"], em["puerto_smtp"]) as server:
            server.login(em["origen"], em["password"])
            server.sendmail(em["origen"], em["destino"], msg.as_string())

        log.info(f"Email enviado a {em['destino']} ({total_nuevos} nuevos, {total_bajadas} bajadas)")

    except smtplib.SMTPAuthenticationError:
        log.error("Error autenticación SMTP — usa contraseña de APLICACIÓN de Gmail (16 chars)")
    except Exception as e:
        log.error(f"Error enviando email: {e}", exc_info=True)


# ══════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════

def main():
    inicio    = datetime.now()
    sep       = "═" * 60
    log.info(sep)
    log.info(f"AutoScout24 Agent — INICIO {inicio.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info(sep)

    busquedas = [b for b in CFG.get("busquedas", []) if b.get("activa", True)]
    if not busquedas:
        log.warning("No hay búsquedas activas en config.json")
        return

    log.info(f"Búsquedas activas: {len(busquedas)}")

    resultados  = []
    url_sheets  = ""

    try:
        for i, b in enumerate(busquedas, 1):
            nombre = b.get("nombre", f"Búsqueda {i}")
            log.info(sep)
            log.info(f"BÚSQUEDA {i}/{len(busquedas)}: {nombre}")
            log.info(sep)

            # 1. Estado anterior
            estado_anterior = cargar_estado(nombre)

            # 2. Scraping
            anuncios = scrape_busqueda(b)
            if not anuncios:
                log.warning(f"[{nombre}] Sin anuncios — posible bloqueo temporal")
                resultados.append({"nombre": nombre, "nuevos": [], "bajadas": []})
                continue

            # 3. Detectar cambios
            nuevos, bajadas = detectar_cambios(anuncios, estado_anterior)

            # 4. Google Sheets
            url_sheets = actualizar_sheets_busqueda(nombre, anuncios, nuevos, bajadas) or url_sheets

            # 5. Guardar estado
            guardar_estado(nombre, anuncios)

            resultados.append({"nombre": nombre, "nuevos": nuevos, "bajadas": bajadas})

        # 6. Email con resultados de todas las búsquedas
        log.info(sep)
        log.info("Enviando email de resumen")
        enviar_email(resultados, url_sheets)

    except Exception as e:
        log.critical(f"Error crítico: {e}", exc_info=True)
        raise
    finally:
        limpiar_temporales()
        duracion = (datetime.now() - inicio).total_seconds()
        log.info(sep)
        log.info(f"AutoScout24 Agent — FIN | Duración: {duracion:.1f}s")
        log.info(sep)


if __name__ == "__main__":
    main()
