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

def _recopilar_listings_del_json(data: dict) -> list[dict]:
    """
    Recorre todo el JSON de Next.js buscando objetos que parezcan anuncios.
    Maneja el patrón Redux normalizado (IDs en array + datos en byId/dict).
    """
    encontrados = []

    def _buscar(obj, depth=0):
        if depth > 15:
            return
        if isinstance(obj, dict):
            tiene_precio   = any(k in obj for k in ("price", "pricing", "prices"))
            tiene_id       = any(k in obj for k in ("id", "guid", "listingId"))
            tiene_vehiculo = any(k in obj for k in
                                 ("vehicle", "vehicleDetails", "make", "model",
                                  "firstRegistration", "mileage"))
            if tiene_precio and tiene_id and tiene_vehiculo:
                encontrados.append(obj)
                return  # no buscar dentro para evitar duplicados
            for v in obj.values():
                if isinstance(v, (dict, list)):
                    _buscar(v, depth + 1)
        elif isinstance(obj, list):
            for item in obj:
                if isinstance(item, (dict, list)):
                    _buscar(item, depth + 1)

    _buscar(data)
    return encontrados


def _safe_dict(val) -> dict:
    """Devuelve val si es dict, o {} en caso contrario."""
    return val if isinstance(val, dict) else {}

def _safe_str(val) -> str:
    """Devuelve val como string si no es dict/list, o '' en caso contrario."""
    if val is None or isinstance(val, (dict, list)):
        return ""
    return str(val).strip()

def _parsear_listing_json(item, pais: str) -> dict | None:
    """Convierte un anuncio del JSON de Next.js en nuestro diccionario."""
    if not isinstance(item, dict):
        return None
    try:
        aid = _safe_str(item.get("id") or item.get("guid") or item.get("listingId") or "")
        if not aid:
            return None

        # URL
        url_raw = _safe_str(
            item.get("url") or item.get("link") or item.get("detailPageUrl") or
            item.get("absoluteUrl") or item.get("shareUrl") or item.get("href") or ""
        )
        if url_raw.startswith("http"):
            url = url_raw
        elif url_raw.startswith("/"):
            url = f"https://www.autoscout24.{pais}{url_raw}"
        else:
            url = ""

        # Título — evitar duplicar partes del modelo en la descripción
        titulo = ""
        for key in ("title", "name", "headline"):
            val = item.get(key)
            if val and isinstance(val, str):
                titulo = val.strip(); break
        if not titulo:
            v    = _safe_dict(item.get("vehicle"))
            make = _safe_dict(v.get("make")).get("name", "") or _safe_str(v.get("make"))
            model= _safe_dict(v.get("model")).get("name", "") or _safe_str(v.get("model"))
            desc = _safe_str(v.get("description") or v.get("modelVersionInput") or
                             v.get("version") or "")
            # Evitar duplicar si desc ya empieza con lo que hay en model
            if model and desc and desc.lower().startswith(model.lower().split()[-1].lower()):
                titulo = f"{make} {model} {desc[len(model.split()[-1]):].strip()}".strip()
            else:
                titulo = " ".join(filter(None, [make, model, desc])).strip()
            titulo = re.sub(r'\s+', ' ', titulo) or "Sin título"

        # Precio
        precio = 0
        p = item.get("price") or item.get("pricing") or {}
        if isinstance(p, dict):
            # 1) campos numéricos directos
            raw = p.get("value") or p.get("amount") or p.get("gross") or 0
            if raw:
                try: precio = int(raw)
                except: precio = limpiar_numero(str(raw))
            # 2) priceFormatted: "€ 28.899" o "28.899 €"
            if not es_precio_razonable(precio):
                pf = _safe_str(p.get("priceFormatted", ""))
                if pf:
                    precio = limpiar_numero(pf)
        elif isinstance(p, (int, float)):
            precio = int(p)
        # 3) fallback: tracking.price (siempre es numérico como string)
        if not es_precio_razonable(precio):
            tracking = _safe_dict(item.get("tracking"))
            pt = _safe_str(tracking.get("price", ""))
            if pt:
                try: precio = int(pt)
                except: precio = limpiar_numero(pt)
        if not es_precio_razonable(precio):
            precio = 0

        # vehicleDetails y vehicle como dicts seguros
        vd   = _safe_dict(item.get("vehicleDetails"))
        v_alt= _safe_dict(item.get("vehicle"))

        # Año
        anio = ""
        for key in ("firstRegistration", "firstRegistrationDate", "yearOfProduction", "year"):
            val = vd.get(key) or v_alt.get(key) or item.get(key) or ""
            val = _safe_str(val)
            if val:
                m = re.search(r"(19|20)\d{2}", val)
                if m: anio = m.group(); break

        # Km
        km_val = 0
        for key in ("mileage", "kilometer", "km", "odometer"):
            val = vd.get(key) or v_alt.get(key) or item.get(key) or 0
            if val and not isinstance(val, (dict, list)):
                km_val = limpiar_numero(str(val)); break
        km = f"{km_val:,} km".replace(",", ".") if km_val else ""

        # Combustible
        combustible = ""
        for key in ("fuelType", "fuel", "fuelTypeDetails", "fuelCategory"):
            val = vd.get(key) or v_alt.get(key) or item.get(key) or ""
            if isinstance(val, dict):
                val = val.get("name") or val.get("value") or val.get("key") or ""
            val = _safe_str(val)
            if val and len(val) < 40:
                combustible = val; break

        # Ubicación
        ubicacion = ""
        seller = _safe_dict(item.get("seller"))
        loc    = _safe_dict(seller.get("location") or item.get("location"))
        for key in ("city", "zip", "region", "countryName", "address"):
            val = _safe_str(loc.get(key) or "")
            if val:
                ubicacion = val; break

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

    listings_raw = _recopilar_listings_del_json(data)
    if not listings_raw:
        log.debug("No se encontraron listings en __NEXT_DATA__")
        return []

    anuncios = []
    for item in listings_raw:
        a = _parsear_listing_json(item, pais)
        if a:
            anuncios.append(a)

    log.debug(f"__NEXT_DATA__: {len(anuncios)} anuncios parseados de {len(listings_raw)} objetos")
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



# Tipos de combustible conocidos en ES / DE / EN
_COMBUSTIBLES = {
    "gasolina", "diésel", "diesel", "híbrido", "híbrida",
    "eléctrico", "eléctrica", "glp", "gnc", "gas natural",
    "hidrógeno", "gasolina/eléctrico", "diésel/eléctrico",
    "gasolina/gas natural", "gas licuado (glp)",
    # Alemán
    "benzin", "elektro", "hybrid", "erdgas", "autogas",
    "wasserstoff", "plug-in-hybrid", "mild-hybrid",
    # Inglés
    "petrol", "electric", "lpg", "cng", "hydrogen",
}

def _extraer_combustible_css(card) -> str:
    # 1. Selectores data-testid específicos
    for sel in ["[data-testid='fuel-type']", "[data-testid='fuel']",
                "[data-testid='fuelType']"]:
        el = card.select_one(sel)
        if el:
            txt = el.get_text(strip=True)
            if txt:
                return txt

    # 2. Buscar cualquier texto que coincida con tipos conocidos
    for el in card.find_all(string=True):
        txt = el.strip()
        if not txt or len(txt) > 40:
            continue
        if txt.lower() in _COMBUSTIBLES:
            return txt
        # Coincidencia parcial para tipos compuestos
        for tipo in _COMBUSTIBLES:
            if len(tipo) > 4 and tipo in txt.lower():
                return txt

    return ""


def _extraer_ubicacion_css(card) -> str:
    # 1. Selectores data-testid específicos
    for sel in ["[data-testid='location']", "[data-testid='seller-location']",
                "[data-testid='city']", ".cldt-summary-seller-contact-location"]:
        el = card.select_one(sel)
        if el:
            txt = el.get_text(strip=True)
            if txt and len(txt) < 60:
                return txt

    # 2. Clases con "location" o "Location" (sin importar el prefijo hash)
    for el in card.find_all(attrs={"class": True}):
        clases = " ".join(el.get("class", []))
        if "ocation" in clases or "seller" in clases.lower():
            txt = el.get_text(strip=True)
            # Descartar textos que sean precio, km, año u otros campos
            if (txt and 2 < len(txt) < 60
                    and "€" not in txt
                    and "km" not in txt.lower()
                    and not re.match(r"^\d{4}$", txt)
                    and not re.match(r"^[\d\.]+$", txt)):
                return txt

    # 3. Buscar elemento que contenga un SVG de pin de mapa (📍)
    #    AutoScout24 suele usar un SVG justo antes del texto de ciudad
    for svg in card.find_all("svg"):
        siguiente = svg.find_next_sibling(string=True)
        if siguiente:
            txt = siguiente.strip()
            if txt and len(txt) < 60 and "€" not in txt and "km" not in txt.lower():
                return txt
        parent = svg.parent
        if parent:
            txt = parent.get_text(strip=True)
            if (txt and len(txt) < 60 and "€" not in txt
                    and "km" not in txt.lower()
                    and not re.match(r"^\d+$", txt)):
                return txt

    return ""


def _extraer_url_css(card, pais: str) -> str:
    """Devuelve la URL del anuncio: el primer enlace del card que no sea de dealer/filtro."""
    base    = f"https://www.autoscout24.{pais}"
    EXCLUIR = ["haendler", "dealer", "concessionnaire", "rivenditore",
               "vendeur", "javascript:", "?sort=", "?atype=", "/lst/"]

    for a in card.find_all("a", href=True):
        href = a.get("href", "").strip()
        if not href or len(href) < 6 or href.startswith("#"):
            continue
        if any(e in href for e in EXCLUIR):
            continue
        return href if href.startswith("http") else base + href

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
            if titulo_el:
                titulo = re.sub(r'\s+', ' ', titulo_el.get_text(separator=" ", strip=True))
            else:
                titulo = "Sin título"

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
    """Construye la URL de busqueda con los filtros configurados.

    Parametros fijos (igual que una busqueda manual en AutoScout24):
      atype=C                  -> coches
      cy=E                     -> pais del TLD
      damaged_listing=exclude  -> excluir accidentados
      desc=0                   -> orden ascendente
      powertype=kw             -> potencia en kW
      sort=price               -> ordenar por precio ascendente
      ustate=N,U               -> nuevos y de ocasion
    """
    marca  = b["marca"].lower().replace(" ", "-")
    modelo = b["modelo"].lower().replace(" ", "-")
    pais   = b.get("pais", "es")
    base   = f"https://www.autoscout24.{pais}/lst/{marca}/{modelo}"

    params = [
        "atype=C",
        "cy=E",
        "damaged_listing=exclude",
        "desc=0",
        f"page={pagina}",
        "powertype=kw",
        "sort=price",
        "ustate=N%2CU",
    ]

    def add(key, val):
        if val is not None and val != "" and val != 0:
            params.append(f"{key}={val}")

    add("fuel",      b.get("fuel"))
    add("fregfrom",  b.get("anio_min"))
    add("fregto",    b.get("anio_max"))
    add("kmfrom",    b.get("km_min"))
    add("kmto",      b.get("km_max"))
    add("pricefrom", b.get("precio_min"))
    add("priceto",   b.get("precio_max"))

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

def _tabla_anuncios_html(anuncios: list, tipo: str) -> str:
    """Genera una tabla HTML compacta con una fila por anuncio."""
    if not anuncios:
        return ""

    if tipo == "nuevo":
        cabecera_color = "#1565c0"
        titulo_seccion = f"🆕 Nuevos anuncios ({len(anuncios)})"
    else:
        cabecera_color = "#c62828"
        titulo_seccion = f"⬇️ Bajadas de precio ({len(anuncios)})"

    filas = []
    for i, a in enumerate(anuncios):
        bg = "#ffffff" if i % 2 == 0 else "#f8f9fa"
        precio_str = f"{a.get('precio', 0):,} €".replace(",", ".")
        anio   = a.get("anio", "") or "—"
        km     = a.get("km", "") or "—"
        comb   = a.get("combustible", "") or "—"
        ubic   = a.get("ubicacion", "") or "—"
        url    = a.get("url", "") or ""
        titulo = a.get("titulo", "Sin título")

        if tipo == "bajada":
            p_ant = f"{a.get('precio_anterior', 0):,} €".replace(",", ".")
            pct   = a.get("porcentaje_bajada", 0)
            precio_cell = (
                f'<span style="text-decoration:line-through;color:#999;">{p_ant}</span> '
                f'<span style="color:#c62828;font-weight:bold;">{precio_str}</span> '
                f'<span style="color:#c62828;font-size:11px;">(-{pct}%)</span>'
            )
        else:
            precio_cell = f'<span style="color:#1565c0;font-weight:bold;">{precio_str}</span>'

        link_cell = (
            f'<a href="{url}" style="background:#1565c0;color:white;padding:3px 9px;'
            f'border-radius:4px;text-decoration:none;font-size:11px;white-space:nowrap;">Ver →</a>'
        ) if url else "—"

        filas.append(f"""
        <tr style="background:{bg};border-bottom:1px solid #e0e0e0;">
            <td style="padding:7px 10px;font-size:12px;max-width:220px;">{titulo}</td>
            <td style="padding:7px 10px;font-size:12px;white-space:nowrap;">{precio_cell}</td>
            <td style="padding:7px 10px;font-size:12px;text-align:center;">{anio}</td>
            <td style="padding:7px 10px;font-size:12px;white-space:nowrap;">{km}</td>
            <td style="padding:7px 10px;font-size:12px;">{comb}</td>
            <td style="padding:7px 10px;font-size:12px;max-width:160px;">{ubic}</td>
            <td style="padding:7px 10px;text-align:center;">{link_cell}</td>
        </tr>""")

    return f"""
    <h3 style="color:{cabecera_color};font-size:14px;margin:16px 0 8px;">{titulo_seccion}</h3>
    <div style="overflow-x:auto;">
    <table style="width:100%;border-collapse:collapse;font-size:12px;
                  border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;">
        <thead>
            <tr style="background:{cabecera_color};color:white;text-align:left;">
                <th style="padding:8px 10px;font-weight:600;">Título</th>
                <th style="padding:8px 10px;font-weight:600;">Precio</th>
                <th style="padding:8px 10px;font-weight:600;">Año</th>
                <th style="padding:8px 10px;font-weight:600;">Km</th>
                <th style="padding:8px 10px;font-weight:600;">Combustible</th>
                <th style="padding:8px 10px;font-weight:600;">Ubicación</th>
                <th style="padding:8px 10px;font-weight:600;">Link</th>
            </tr>
        </thead>
        <tbody>{"".join(filas)}</tbody>
    </table>
    </div>"""


def _seccion_busqueda_html(nombre: str, nuevos: list, bajadas: list) -> str:
    if not nuevos and not bajadas:
        return ""

    cabecera = f"""
    <div style="margin:24px 0 10px;padding:10px 14px;background:#e8eaf6;
                border-left:4px solid #3949ab;border-radius:0 6px 6px 0;">
        <strong style="color:#1a237e;font-size:15px;">🔍 {nombre}</strong>
        <span style="color:#5c6bc0;font-size:12px;margin-left:10px;">
            {len(nuevos)} nuevos · {len(bajadas)} bajadas
        </span>
    </div>"""

    return cabecera + _tabla_anuncios_html(nuevos, "nuevo") + _tabla_anuncios_html(bajadas, "bajada")


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
