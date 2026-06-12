# -*- coding: utf-8 -*-
"""
main.py - הגרסה היציבה והמלאה (ללא Threading, עם סידור דפים בטוח)
"""
from __future__ import annotations
import os
import shutil
import sys
import json
import base64, difflib, re
import time
import win32com.client
import pythoncom
import requests
from typing import Dict, Any, Optional
# הגדרת קידוד
sys.stdout.reconfigure(encoding='utf-8')
# ייבוא הפונקציות הגרפיות
from illustrator_ops import run_jsx, open_and_color_template, place_and_simulate_print, update_size_label, delete_side_assets, delete_print_layer_only, save_pdf, clean_layout, apply_extra_colors, delete_information_layer, set_order_number_in_simulation, remove_order_number_from_simulation, close_all_illustrator_documents, VARIABLE_PRINT_COLS, VARIABLE_PRINT_GAP_MIN
from vectorizer_ops import convert_to_svg
# --- הגדרות ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_DOWNLOAD_DIR = os.path.join(BASE_DIR, "temp_downloads")
TEMP_AI_DIR = os.path.join(BASE_DIR, "temp_ai_files")
if not os.path.exists(TEMP_DOWNLOAD_DIR): os.makedirs(TEMP_DOWNLOAD_DIR)
# --- קונפיגורציה ---
try:
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    default_docs = os.path.join(os.path.expanduser("~"), "Documents", "Auto_Print_Output")
    SAVE_FOLDER = config.get('save_folder_path', default_docs)
except:
    SAVE_FOLDER = os.path.join(os.path.expanduser("~"), "Documents", "Auto_Print_Output")
ORDERS_ROOT_DIR = SAVE_FOLDER
if not os.path.exists(ORDERS_ROOT_DIR): os.makedirs(ORDERS_ROOT_DIR)
# --- תבניות ---
TEMPLATES = {
    '90 Bag': os.path.join(BASE_DIR, 'Simulations', '90 Bag.ai'),
    '50 Bag': os.path.join(BASE_DIR, 'Simulations', '50 Bag.ai'),
    '30 Bag': os.path.join(BASE_DIR, 'Simulations', '30 Bag.ai'),
    'Bandana Cap': os.path.join(BASE_DIR, 'Simulations', 'Bandana Cap.ai'),
    'Baby Bodysuit': os.path.join(BASE_DIR, 'Simulations', 'Baby Bodysuit.ai'),
    'Apron': os.path.join(BASE_DIR, 'Simulations', 'Apron.ai'),
    'Buff': os.path.join(BASE_DIR, 'Simulations', 'Buff.ai'),
    'Boxers': os.path.join(BASE_DIR, 'Simulations', 'Boxers.ai'),
    'Beanie': os.path.join(BASE_DIR, 'Simulations', 'Beanie.ai'),
    'Chef Jacket': os.path.join(BASE_DIR, 'Simulations', 'Chef Jacket.ai'),
    'Cargo Pants': os.path.join(BASE_DIR, 'Simulations', 'Cargo Pants.ai'),
    'Canvas Bag': os.path.join(BASE_DIR, 'Simulations', 'Canvas Bag.ai'),
    'Flag 80-110': os.path.join(BASE_DIR, 'Simulations', 'Flag 80-110.ai'),
    'Flag Car': os.path.join(BASE_DIR, 'Simulations', 'Flag Car.ai'),
    'Fashion Vest': os.path.join(BASE_DIR, 'Simulations', 'Fashion Vest.ai'),
    'Drawstring Bag': os.path.join(BASE_DIR, 'Simulations', 'Drawstring Bag.ai'),
    'Fleece1': os.path.join(BASE_DIR, 'Simulations', 'Fleece1.ai'),
    'Fleece Blanket': os.path.join(BASE_DIR, 'Simulations', 'Fleece Blanket.ai'),
    'Flag 150-100': os.path.join(BASE_DIR, 'Simulations', 'Flag 150-100.ai'),
    'High Visibility Vest': os.path.join(BASE_DIR, 'Simulations', 'High Visibility Vest.ai'),
    'Hat': os.path.join(BASE_DIR, 'Simulations', 'Hat.ai'),
    'Fleece2': os.path.join(BASE_DIR, 'Simulations', 'Fleece2.ai'),
    'Kippah': os.path.join(BASE_DIR, 'Simulations', 'Kippah.ai'),
    'Hoodie': os.path.join(BASE_DIR, 'Simulations', 'Hoodie.ai'),
    'Hoodie T-shirt': os.path.join(BASE_DIR, 'Simulations', 'Hoodie T-shirt.ai'),
    'Long Baby Bodysuit': os.path.join(BASE_DIR, 'Simulations', 'Long Baby Bodysuit.ai'),
    'Legionnaire Hat': os.path.join(BASE_DIR, 'Simulations', 'Legionnaire Hat.ai'),
    'Lab Coat': os.path.join(BASE_DIR, 'Simulations', 'Lab Coat.ai'),
    'Long Short': os.path.join(BASE_DIR, 'Simulations', 'Long Short.ai'),
    'Long Polo': os.path.join(BASE_DIR, 'Simulations', 'Long Polo.ai'),
    'Long Chef Jacket': os.path.join(BASE_DIR, 'Simulations', 'Long Chef Jacket.ai'),
    'Overalls': os.path.join(BASE_DIR, 'Simulations', 'Overalls.ai'),
    'Neck Warmer': os.path.join(BASE_DIR, 'Simulations', 'Neck Warmer.ai'),
    'Mesh Laundry Basket': os.path.join(BASE_DIR, 'Simulations', 'Mesh Laundry Basket.ai'),
    'Laundry Basket': os.path.join(BASE_DIR, 'Simulations', 'Laundry Basket.ai'),
    'Scarf': os.path.join(BASE_DIR, 'Simulations', 'Scarf.ai'),
    'Raglan Shirt': os.path.join(BASE_DIR, 'Simulations', 'Raglan Shirt.ai'),
    'Polo': os.path.join(BASE_DIR, 'Simulations', 'Polo.ai'),
    'Sweater': os.path.join(BASE_DIR, 'Simulations', 'Sweater.ai'),
    'Softshell': os.path.join(BASE_DIR, 'Simulations', 'Softshell.ai'),
    'Short': os.path.join(BASE_DIR, 'Simulations', 'Short.ai'),
    'Triangular Bandana': os.path.join(BASE_DIR, 'Simulations', 'Triangular Bandana.ai'),
    'Towel': os.path.join(BASE_DIR, 'Simulations', 'Towel.ai'),
    'Tactical Vest': os.path.join(BASE_DIR, 'Simulations', 'Tactical Vest.ai'),
    'Sweatpants': os.path.join(BASE_DIR, 'Simulations', 'Sweatpants.ai'),
    'Zippered Hoodie': os.path.join(BASE_DIR, 'Simulations', 'Zippered Hoodie.ai'),
    'Wide Brimmed Hat': os.path.join(BASE_DIR, 'Simulations', 'Wide Brimmed Hat.ai'),
    'Undershirt': os.path.join(BASE_DIR, 'Simulations', 'Undershirt.ai'),
    'Mesh hat': os.path.join(BASE_DIR, 'Simulations', 'Mesh hat.ai'),
    'Combined hat': os.path.join(BASE_DIR, 'Simulations', 'Combined hat.ai'),
    'Shirt': os.path.join(BASE_DIR, 'Simulations', 'Short.ai'),
}
API_ID = "vkd2vcts24ywdpk"
API_SECRET = "r20rqffqdcv6vj0ahukmiu9i8ma6ur4g0e1a5o9c7vugsoracpk8"
EXTENDED_COLOR_MAP = {
    'צבעוני': 'ORIGINAL', 'מקורי': 'ORIGINAL', 'ללא שינוי': 'ORIGINAL', 'צבעוני (ללא שינוי)': 'ORIGINAL',
    'שחור': '#000000', 'לבן': '#FFFFFF', 'אדום': '#cc2127', 'צהוב': '#fff200', 'כתום': '#f7941d',
    'זהב': '#FFD700', 'גולד': '#FFD700', 'כסף': '#C0C0C0', 'סילבר': '#C0C0C0', 'ברונזה': '#CD7F32',
    'צהוב זוהר': '#fff200', 'כתום זוהר': '#f7941d', 'ירוק זוהר': '#8dc63f',
    'אפור': '#808080', 'אפור מלנץ': '#b3b3b3', 'אפור מלנץ\'': '#b3b3b3', 'מלנץ': '#b3b3b3',
    'אנטרציט': '#36454F', 'אפור עכבר': '#4d4d4d', 'גרפיט': '#383838',
    'כאמל': '#c2b59b', 'קאמל': '#c2b59b', 'חאקי': '#c2b59b', 'שמנת': '#FFFDD0',
    'בז': '#F5F5DC', 'בז\'': '#F5F5DC', 'אוף וויט': '#c2b59b', 'אווף ויט': '#c2b59b',
    'מוקה': '#967969', 'חום': '#8B4513',
    'אוף וויט כאמל': '#c2b59b', 'אווף ויט-כאמל': '#c2b59b', 'שמנת כאמל': '#c2b59b',
    'כחול': '#0000FF', 'נייבי': '#0e2d4e', 'כחול נייבי': '#0e2d4e', 'ניבי': '#0e2d4e',
    'רויאל': '#1d4483', 'כחול רויאל': '#1d4483', 'תכלת': '#58b7de', 'טורקיז': '#029faa',
    'ים': '#40E0D0', 'פטרול': '#005f6a',
    'ירוק': '#8dc63f', 'ירוק בקבוק': '#033A1E', 'בקבוק': '#033A1E', 'ירוק תפוח': '#8DB600',
    'תפוח': '#8DB600', 'זיית': '#4f4e20', 'זית': '#4f4e20', 'ירוק זית': '#4f4e20', 'מנטה': '#98FF98',
    'ורוד': '#FFC0CB', 'ורוד בייבי': '#f1b1d0', 'ורוד ביבי': '#f1b1d0', 'ביבי': '#f1b1d0', 'בייבי': '#f1b1d0', 'פוקסיה': '#ec008c',
    'ורוד פוקסיה': '#ec008c', 'סגול': '#311d72', 'סגול כהה': '#4B0082', 'חציל': '#4B0082',
    'ליילך': '#C8A2C8', 'בורדו': '#8c191f', 'יין': '#800000',
}
# -----------------------------------------------------------
# פונקציות עזר (ללא שימוש ב-concurrent)
# -----------------------------------------------------------
_HEX6_RE = re.compile(r'^[0-9A-Fa-f]{6}$')

def try_parse_hex_color(name: str) -> Optional[str]:
    """מחזיר #RRGGBB אם הקלט הוא בדיוק 6 תווים הקסדצימליים (אותיות גדולות/קטנות)."""
    if not name or not isinstance(name, str):
        return None
    s = name.strip()
    if s.startswith('#'):
        s = s[1:]
    if len(s) != 6 or not _HEX6_RE.match(s):
        return None
    return '#' + s.upper()

def get_contrasting_print_color(bg_hex):
    if not bg_hex: return '#FFFFFF'
    h = bg_hex.lstrip('#')
    try:
        r, g, b = tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
        luminance = (0.299 * r + 0.587 * g + 0.114 * b)
        return '#FFFFFF' if luminance < 128 else '#000000'
    except: return '#FFFFFF'
def get_hex_smart(name, return_none_on_fail=False):
    if not name or not isinstance(name, str): return None if return_none_on_fail else '#FFFFFF'
    name_clean = name.strip()
    hex_from_code = try_parse_hex_color(name_clean)
    if hex_from_code:
        return hex_from_code
    if name_clean in EXTENDED_COLOR_MAP: return EXTENDED_COLOR_MAP[name_clean]
    matches = difflib.get_close_matches(name_clean, EXTENDED_COLOR_MAP.keys(), n=1, cutoff=0.5)
    if matches: return EXTENDED_COLOR_MAP[matches[0]]
    return None if return_none_on_fail else '#FFFFFF'
def resolve_print_color(req, shirt):
    txt = str(req).strip() if req else ""
    found = get_hex_smart(txt, True)
    if found == 'ORIGINAL': return None
    if found: return found
    return get_contrasting_print_color(shirt)
def get_hex(name):
    val = get_hex_smart(name)
    return val if val != 'ORIGINAL' else None

def _is_gray_color_name(name: str) -> bool:
    n = name.strip()
    return n == 'אפור' or n.startswith('אפור') or n in ('מלנץ', 'גרפיט', 'אנטרציט')

def _is_burgundy_color_name(name: str) -> bool:
    return name.strip() in ('בורדו', 'יין', 'אדום')

def _is_olive_color_name(name: str) -> bool:
    n = name.strip()
    return n in ('זית', 'זיית', 'ירוק זית') or n.startswith('ירוק זית')

def _is_black_color_name(name: str) -> bool:
    return name.strip() == 'שחור'

def normalize_split_product_color(col_raw: str) -> str:
    """אפור+בורדו/אדום: תמיד אפור ראשון (Side1), בורדו/אדום שני (Side2).
    זית+שחור: צבע יחיד – שחור בלבד."""
    if not col_raw or '-' not in col_raw:
        return col_raw
    parts = [p.strip() for p in col_raw.split("-")]
    if len(parts) != 2:
        return col_raw
    first, second = parts[0], parts[1]
    if (_is_olive_color_name(first) and _is_black_color_name(second)) or (
        _is_black_color_name(first) and _is_olive_color_name(second)
    ):
        return 'שחור'
    if _is_gray_color_name(first) and _is_burgundy_color_name(second):
        return col_raw
    if _is_burgundy_color_name(first) and _is_gray_color_name(second):
        return f"{second}-{first}"
    return col_raw

def get_unique_filename(path):
    if not os.path.exists(path): return path
    base, ext = os.path.splitext(path)
    counter = 1
    while True:
        new_path = f"{base} ({counter}){ext}"
        if not os.path.exists(new_path): return new_path
        counter += 1
def download_image(url_or_base64, filename_prefix):
    try:
        if not url_or_base64:
            return None
        if isinstance(url_or_base64, str) and url_or_base64.startswith("blob:"):
            print(f"   > Download skip: blob URL not supported ({url_or_base64[:80]})")
            return None
        if url_or_base64.startswith('data:'):
            header, encoded = url_or_base64.split(',', 1)
            file_ext = '.png'
            if 'image/svg' in header: file_ext = '.svg'
            elif 'pdf' in header: file_ext = '.pdf'
            elif 'illustrator' in header or 'postscript' in header: file_ext = '.ai'
            path = os.path.join(TEMP_DOWNLOAD_DIR, f"{filename_prefix}{file_ext}")
            with open(path, 'wb') as f: f.write(base64.b64decode(encoded))
            return path
        elif os.path.exists(url_or_base64) or (len(url_or_base64)>1 and url_or_base64[1]==':'):
            if not os.path.exists(url_or_base64):
                print(f"   > Download skip: local path not found: {url_or_base64}")
                return None
            _, ext = os.path.splitext(url_or_base64)
            path = os.path.join(TEMP_DOWNLOAD_DIR, f"{filename_prefix}{ext or '.png'}")
            if os.path.abspath(url_or_base64) != os.path.abspath(path): shutil.copy(url_or_base64, path)
            return path
        elif url_or_base64.startswith('http'):
            ext = ".png"
            lower = url_or_base64.lower()
            if '.pdf' in lower: ext = ".pdf"
            elif '.svg' in lower: ext = ".svg"
            elif '.ai' in lower: ext = ".ai"
            path = os.path.join(TEMP_DOWNLOAD_DIR, f"{filename_prefix}{ext}")
            r = requests.get(url_or_base64, stream=True, timeout=120)
            if r.status_code == 200:
                content_type = (r.headers.get('content-type') or '').lower()
                if ext == ".png":
                    if 'pdf' in content_type: ext = ".pdf"
                    elif 'svg' in content_type: ext = ".svg"
                    elif 'postscript' in content_type or 'illustrator' in content_type: ext = ".ai"
                    if ext != ".png":
                        path = os.path.join(TEMP_DOWNLOAD_DIR, f"{filename_prefix}{ext}")
                with open(path, 'wb') as f: shutil.copyfileobj(r.raw, f)
                print(f"   > Downloaded: {os.path.basename(path)} ({r.status_code})")
                return path
            print(f"   > Download failed ({r.status_code}): {url_or_base64[:120]}")
    except Exception as e:
        print(f"   > Download error for {str(url_or_base64)[:120]}: {e}")
    return None
def vec_single(d: Dict, f: str, id: str, sec: str) -> Optional[str]:
    if not d.get('exists'): return None
    if not d.get('file') or not os.path.exists(d['file']): return None
    skip = d.get('no_vectorization', False)
    orig_dst = os.path.join(f, f"{d['prefix']}_{os.path.basename(d['file'])}")
    dst = get_unique_filename(orig_dst)
    if skip:
        shutil.copy(d['file'], dst)
        return dst
    else:
        shutil.copy(d['file'], dst)
        return convert_to_svg(dst, id, sec)

SIDE_KEYS = ['front', 'back', 'right_sleeve', 'left_sleeve']
SIDE_TO_PREFIX = {'front': 'F', 'back': 'B', 'right_sleeve': 'RS', 'left_sleeve': 'LS'}


def normalize_side_name(name: str) -> Optional[str]:
    if not name:
        return None
    n = str(name).strip().lower()
    aliases = {
        'front': 'front', 'f': 'front', 'קידמי': 'front', 'pf': 'front',
        'back': 'back', 'b': 'back', 'אחורי': 'back', 'pb': 'back',
        'right_sleeve': 'right_sleeve', 'rs': 'right_sleeve', 'right': 'right_sleeve',
        'left_sleeve': 'left_sleeve', 'ls': 'left_sleeve', 'left': 'left_sleeve',
    }
    if n in aliases:
        return aliases[n]
    return n if n in SIDE_KEYS else None


SIDE_DEFAULTS = {
    'front': {'prefix': 'F', 'category': 'A4', 'label': 'size_Front', 'heb': 'קדמי'},
    'back': {'prefix': 'B', 'category': 'A4', 'label': 'size_Back', 'heb': 'אחורי'},
    'right_sleeve': {'prefix': 'RS', 'category': 'Sleeve', 'label': 'size_Right_Sleeve', 'heb': 'שרוול ימין'},
    'left_sleeve': {'prefix': 'LS', 'category': 'Sleeve', 'label': 'size_Left_Sleeve', 'heb': 'שרוול שמאל'},
}


def side_is_active(loc_data: Dict[str, Any]) -> bool:
    if not loc_data:
        return False
    ref = loc_data.get('reuse_print_from') or loc_data.get('reusePrintFrom')
    if ref and isinstance(ref, dict):
        return True
    return bool(loc_data.get('exists'))


def resolve_reuse_product_number(ref_product, products_all) -> int:
    try:
        n = int(ref_product)
    except (TypeError, ValueError):
        raise ValueError(f"Invalid reuse_print_from product: {ref_product}")
    if 1 <= n <= len(products_all):
        return n
    for i, prod in enumerate(products_all):
        if str(prod.get('item_index', '')).strip() == str(ref_product):
            return i + 1
    raise ValueError(f"reuse_print_from product {ref_product} not found in order")


def enrich_side_data(side: str, data: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    merged = dict(SIDE_DEFAULTS.get(side, {}))
    if data:
        merged.update({k: v for k, v in data.items() if v is not None and v != ''})
    return merged


def parse_reuse_print_from(loc_data: Dict[str, Any]) -> Optional[tuple[int, str]]:
    if not loc_data:
        return None
    ref = loc_data.get('reuse_print_from') or loc_data.get('reusePrintFrom')
    if not ref or not isinstance(ref, dict):
        return None
    src_product = ref.get('product') if ref.get('product') is not None else ref.get('prod')
    src_location = ref.get('location') or ref.get('side') or ref.get('loc')
    if src_product is None or not src_location:
        return None
    try:
        src_product = int(src_product)
    except (TypeError, ValueError):
        return None
    src_side = normalize_side_name(src_location)
    if not src_side:
        raise ValueError(f"Invalid reuse_print_from location: {src_location}")
    return src_product, src_side


def side_data_for_cache(loc_data: Dict[str, Any]) -> Dict[str, Any]:
    skip_keys = {'file', 'file_url', 'reuse_print_from', 'reusePrintFrom'}
    return {k: v for k, v in loc_data.items() if k not in skip_keys and not str(k).startswith('_')}


def download_all_side_files(products, order_prefix: str) -> None:
    """מוריד קבצים לכל המוצרים מראש — גם כשמוצר מאוחר הוא מקור ל-reuse."""
    for i, prod in enumerate(products):
        for loc in SIDE_KEYS:
            loc_d = prod.get(loc, {}) or {}
            if not side_is_active(loc_d) or parse_reuse_print_from(loc_d):
                continue
            if loc_d.get('file_url') and not loc_d.get('file'):
                path = download_image(loc_d['file_url'], f"{order_prefix}_{i}_{loc}")
                if path:
                    loc_d['file'] = path


def resolve_side_assets(
    src_p: int,
    src_s: str,
    products_all,
    folder: str,
    side_cache: Dict,
    visiting: Optional[set] = None,
) -> Dict[str, Any]:
    """מחזיר svg + side_data למקור — גם אם המקור מאוחר יותר ברשימה."""
    cache_key = (src_p, src_s)
    if cache_key in side_cache:
        return side_cache[cache_key]
    if visiting is None:
        visiting = set()
    if cache_key in visiting:
        raise ValueError(f"Circular reuse detected at product {src_p} {src_s}")
    visiting = visiting | {cache_key}

    src_prod = products_all[src_p - 1]
    src_loc_raw = src_prod.get(src_s, {}) or {}
    reuse = parse_reuse_print_from(src_loc_raw)
    if reuse:
        chain_p = resolve_reuse_product_number(reuse[0], products_all)
        return resolve_side_assets(chain_p, reuse[1], products_all, folder, side_cache, visiting)

    src_loc = enrich_side_data(src_s, src_loc_raw)
    if not side_is_active(src_loc):
        raise ValueError(f"Source product {src_p} {src_s} has no print")
    if not src_loc.get('file') or not os.path.exists(src_loc['file']):
        raise ValueError(f"Source product {src_p} {src_s} missing file (download failed?)")
    svg = vec_single(src_loc, folder, API_ID, API_SECRET)
    if not svg:
        raise ValueError(f"Source product {src_p} {src_s} vectorization failed")
    entry = {'svg': svg, 'side_data': side_data_for_cache(src_loc)}
    side_cache[cache_key] = entry
    return entry

# -----------------------------------------------------------
# עיבוד בודד (ללא שינוי, זה עובד טוב)
# -----------------------------------------------------------
def process_single_product_to_temp(order, idx, folder, is_wholesale=False, order_id=None, products_all=None, side_cache=None):
    if side_cache is None:
        side_cache = {}
    if products_all is None:
        products_all = [order]
    pythoncom.CoInitialize()
    doc = None
    app = None
    try:
        prod = order.get('product_type', 'Shirt')
        print(f"\n>> Processing Product {idx+1}: {prod}")
        t_path = TEMPLATES.get(prod)
        if not t_path or not os.path.exists(t_path): return None
        sides = SIDE_KEYS
        svgs = {}
        original_sides = {s: dict(order.get(s, {}) or {}) for s in sides}
        for s in sides:
            loc = original_sides[s]
            if not side_is_active(loc):
                continue
            reuse_ref = parse_reuse_print_from(loc)
            if reuse_ref:
                ref_product, src_s = reuse_ref
                src_p = resolve_reuse_product_number(ref_product, products_all)
                if src_p < 1 or src_p > len(products_all):
                    raise ValueError(f"Product {idx+1} {s}: reuse_print_from product {src_p} out of range")
                if src_p - 1 == idx and src_s == s:
                    raise ValueError(f"Product {idx+1} {s}: cannot reuse from itself")
                cached = resolve_side_assets(src_p, src_s, products_all, folder, side_cache)
                svgs[s] = cached['svg']
                order[s] = {
                    **enrich_side_data(src_s, cached['side_data']),
                    'exists': True,
                    '_skip_print_layer': True,
                }
                print(f"   > [Product {idx+1}] {s}: reuse print from product {src_p} {src_s} (simulation only)")
                continue
            loc = enrich_side_data(s, loc)
            order[s] = loc
            own_key = (idx + 1, s)
            if own_key in side_cache:
                svgs[s] = side_cache[own_key]['svg']
            else:
                res = vec_single(loc, folder, API_ID, API_SECRET)
                if res:
                    svgs[s] = res
        col_raw = normalize_split_product_color(order.get('product_color_hebrew', ""))
        parts = [p.strip() for p in col_raw.split("-")] if "-" in col_raw else [col_raw]
        h1 = get_hex_smart(parts[0])
        h2 = get_hex_smart(parts[1]) if len(parts) >= 2 else h1
        is_split = len(parts) >= 2
        doc, app = open_and_color_template(t_path, h1, h2, is_split, prod)
        extra = order.get('extra_colors_hebrew', [])
        extra_data = []
        for name in extra:
            p = [x.strip() for x in name.split("-")]
            pair = [get_hex_smart(c) for c in p[:2] if get_hex_smart(c)!='ORIGINAL']
            if pair: extra_data.append(pair)
        apply_extra_colors(app, extra_data)
        for s in sides:
            d = enrich_side_data(s, order.get(s, {}) or {})
            if not side_is_active(d) or not svgs.get(s):
                continue
            is_r = d.get('no_vectorization', False)
            fc = resolve_print_color(d.get('req_color_hebrew'), h1)
            cp = fc if fc!='#FFFFFF' else '#000000'
            prefix = d.get('prefix') or SIDE_TO_PREFIX.get(s, 'F')
            w = place_and_simulate_print(doc, app, svgs[s], prefix, d.get('category', 'A4'), cp, fc, is_r)
            label = d.get('label') or SIDE_DEFAULTS.get(s, {}).get('label', f'size_{s}')
            heb = d.get('heb') or SIDE_DEFAULTS.get(s, {}).get('heb', '')
            if w>0: update_size_label(doc, app, label, w, heb)
            if d.get('_skip_print_layer'):
                delete_print_layer_only(app, prefix)
                print(f"   > [Product {idx+1}] {s}: removed Print layer (reuse)")
            orig = enrich_side_data(s, original_sides[s])
            reuse_from = parse_reuse_print_from(orig)
            if not reuse_from:
                side_cache[(idx + 1, s)] = {
                    'svg': svgs[s],
                    'side_data': side_data_for_cache(orig),
                }
            else:
                src_p = resolve_reuse_product_number(reuse_from[0], products_all)
                src_s = reuse_from[1]
                side_cache[(idx + 1, s)] = side_cache[(src_p, src_s)]
        rs_active = side_is_active(order.get('right_sleeve', {}))
        ls_active = side_is_active(order.get('left_sleeve', {}))
        if not rs_active and not ls_active:
            delete_side_assets(doc, app, "Print_Sleeves", "size_Right_Sleeve")
            run_jsx(app, "try{app.activeDocument.textFrames.getByName('size_Left_Sleeve').remove();}catch(e){}")
        elif not rs_active:
            run_jsx(app, "try{app.activeDocument.textFrames.getByName('size_Right_Sleeve').remove();}catch(e){}")
        elif not ls_active:
            run_jsx(app, "try{app.activeDocument.textFrames.getByName('size_Left_Sleeve').remove();}catch(e){}")
        if not side_is_active(order.get('front', {})): delete_side_assets(doc, app, "Print_Front", "size_Front")
        if not side_is_active(order.get('back', {})): delete_side_assets(doc, app, "Print_Back", "size_Back")
        clean_layout(app)
        # מוסיפים מספר הזמנה ל"NumberOrder" בכל המוצרים; במוצרים 2+ מוחקים את התיבה – נשאר רק בראשון
        if order_id:
            set_order_number_in_simulation(app, order_id)
            if idx >= 1:
                remove_order_number_from_simulation(app)
        # אם זה סיטונאי, מוחקים את שכבה/קבוצה "information" מתוך "Simulation"
        if is_wholesale:
            print(f"   > [Product {idx+1}] מחיקת שכבה 'information'...")
            delete_information_layer(app)
        out_name = f"temp_{idx}.ai"
        out_path = os.path.join(TEMP_AI_DIR, out_name)
        doc.SaveAs(out_path)
        doc.Close(2)
        print(f"   > Saved: {out_name}")
        return out_path
    except Exception as e:
        print(f"Error processing product {idx+1}: {e}")
        import traceback
        traceback.print_exc()
        if doc:
            try:
                doc.Close(2)
            except Exception:
                pass
        return None
    finally:
        pythoncom.CoUninitialize()
# -----------------------------------------------------------
# פונקציית האיחוד: ה-Super Script (חישוב גובה ורוחב דינמי חכם)
# -----------------------------------------------------------
def create_and_run_merge_script(files_list, output_pdf, order_data=None):
    pythoncom.CoInitialize()
    if not files_list: return
# חילוץ רשימת המוצרים שהוזמנו כדי למנוע כפילויות בסרגל הצד
    ordered_products = []
    if order_data and 'products' in order_data:
        # אוסף את כל ה-product_type הייחודיים מההזמנה
        ordered_products = list(set([str(p.get('product_type')) for p in order_data['products']]))
    # --- הדפסה ללוג לצורך בקרה (מה שביקשת) ---
    print("\n" + "="*40)
    print(f"📊 SIDEBAR DATA CONTROL:")
    print(f"🛒 Products in current order: {ordered_products}")
    # בדיקה אילו מהמוצרים האלו קיימים ברשימת האופציות של הסרגל
    upsell_list = ["Apron", "Drawstring Bag", "Wide Brimmed Hat", "Neck Warmer", "Canvas Bag", "Polo", "Fleece1", "Beanie", "Boxers", "Short", "Hoodie", "Hat"]
    filtered_out = [p for p in ordered_products if p in upsell_list]
    if filtered_out:
        print(f"🚫 Products to be FILTERED OUT from sidebar: {filtered_out}")
    else:
        print(f"✅ No products from the order match the sidebar options.")
    print("="*40 + "\n")
    # ------------------------------------------
    # קריאת נתוני simulation_ad מה-order_data
    sim_ad = order_data.get('simulation_ad', {}) or {} if order_data else {}
    print(f"\n🔍 DEBUG: sim_ad = {sim_ad}")
    show_sidebar = bool(sim_ad.get('enabled', False))
    upsell_mode = sim_ad.get('mode', 'random')
    manual_list = sim_ad.get('selected_products', [])
    print(f"🔍 DEBUG: show_sidebar = {show_sidebar}")
    print(f"🔍 DEBUG: upsell_mode = {upsell_mode}")
    print(f"🔍 DEBUG: manual_list = {manual_list}")
    # הכנת אובייקט הגדרות ל-JSX
    job_config = {
        "sidebar_path": config.get('sidebar_template_path', "").replace("\\", "/"),
        "ordered_products": ordered_products,
        "show_sidebar": show_sidebar,
        "upsell_mode": upsell_mode,
        "manual_products": manual_list,
    }
    # כתיבת הקונפיג לקובץ JSON זמני
    with open(os.path.join(BASE_DIR, "current_job.json"), "w", encoding="utf-8") as f:
        json.dump(job_config, f, ensure_ascii=False, indent=4)    # ------------------------------------------
    js_files = [f.replace("\\", "/") for f in files_list]
    sidebar_logic_path = os.path.join(BASE_DIR, "sidebar_logic.jsx").replace("\\", "/")
    sidebar_exists = os.path.exists(sidebar_logic_path)
    print(f"🔍 DEBUG: sidebar_logic.jsx exists: {sidebar_exists}")
    print(f"🔍 DEBUG: sidebar_logic.jsx path: {sidebar_logic_path}")
    jsx_content = f"""
    #target illustrator
    var files = {json.dumps(js_files)};
    function openFileSafe(path) {{
        var f = new File(path);
        if (!f.exists) throw new Error("File not found: " + path);
        var target = f.fsName.toLowerCase();
        for (var d = 0; d < app.documents.length; d++) {{
            try {{
                if (app.documents[d].fullName && app.documents[d].fullName.fsName.toLowerCase() === target) {{
                    return app.documents[d];
                }}
            }} catch(e) {{}}
        }}
        return app.open(f);
    }}
    function main() {{
        if (files.length === 0) return;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        var maxWidth = 0; var maxHeight = 0;
        for (var i = 0; i < files.length; i++) {{
            var tempDoc = openFileSafe(files[i]);
            var m = calculateLayoutMetrics(tempDoc);
            if (m.width > maxWidth) maxWidth = m.width;
            if (m.height > maxHeight) maxHeight = m.height;
            tempDoc.close(SaveOptions.DONOTSAVECHANGES);
        }}
        var GAP_X = {int(VARIABLE_PRINT_GAP_MIN)};
        var GAP_Y = {int(VARIABLE_PRINT_GAP_MIN + 100)};
        var STEP_X = maxWidth + GAP_X;
        var STEP_Y = maxHeight + GAP_Y;
        var COLS = {int(VARIABLE_PRINT_COLS)};
        $.writeln("Merge grid: " + files.length + " products, STEP_X=" + STEP_X + " STEP_Y=" + STEP_Y + " COLS=" + COLS);
        var masterDoc = openFileSafe(files[0]);
        organizeMasterContent(masterDoc);
        for (var j = 1; j < files.length; j++) {{
            var col = j % COLS;
            var row = Math.floor(j / COLS);
            processNextFileFast(masterDoc, files[j], (j+1).toString(), col * STEP_X, -(row * STEP_Y));
        }}
        var sideFile = new File("{sidebar_logic_path}");
        var showSidebar = {str(show_sidebar).lower()};
        $.writeln("=== MERGE SCRIPT: Sidebar check ===");
        $.writeln("showSidebar: " + showSidebar);
        $.writeln("sideFile exists: " + sideFile.exists);
        $.writeln("masterDoc name: " + masterDoc.name);
        if (showSidebar && sideFile.exists) {{
            $.writeln("--- Calling sidebar_logic.jsx ---");
            // העברת masterDoc כמשתנה גלובלי ל-sidebar_logic.jsx
            var targetDoc = masterDoc;
            $.writeln("Activating targetDoc: " + targetDoc.name);
            targetDoc.activate();
            app.activeDocument = targetDoc;
            $.writeln("Active document after activation: " + app.activeDocument.name);
            // הגדרת משתנה גלובלי ש-sidebar_logic.jsx יוכל להשתמש בו
            $.global.targetMasterDoc = targetDoc;
            $.writeln("targetMasterDoc set in global");
            $.writeln("Evaluating sidebar_logic.jsx file...");
            $.evalFile(sideFile);
            $.writeln("sidebar_logic.jsx execution completed");
            // וידוא ש-masterDoc נשאר פעיל אחרי הוספת התפריט הצד
            $.writeln("Reactivating targetDoc after sidebar logic...");
            targetDoc.activate();
            app.activeDocument = targetDoc;
            $.writeln("Final active document: " + app.activeDocument.name);
            $.writeln("Document layers count: " + targetDoc.layers.length);
            $.writeln("Document artboards count: " + targetDoc.artboards.length);
            $.writeln("--- Sidebar logic complete ---");
        }} else {{
            $.writeln("Sidebar skipped (showSidebar=" + showSidebar + ", fileExists=" + sideFile.exists + ")");
        }}
        reorderArtboardsSafe(masterDoc);
    }}
    // --- שאר הפונקציות (calculateLayoutMetrics, organizeMasterContent וכו') נשארות ללא שינוי ---
    function calculateLayoutMetrics(doc) {{
        var minX = Infinity; var maxX = -Infinity; var maxY = -Infinity; var minY = Infinity;
        for (var i = 0; i < doc.artboards.length; i++) {{
            var r = doc.artboards[i].artboardRect;
            if (r[0] < minX) minX = r[0]; if (r[2] > maxX) maxX = r[2];
            if (r[1] > maxY) maxY = r[1]; if (r[3] < minY) minY = r[3];
        }}
        return {{ width: Math.abs(maxX - minX), height: Math.abs(maxY - minY) }};
    }}
    function organizeMasterContent(doc) {{
        app.executeMenuCommand('unlockAll'); app.executeMenuCommand('showAll');
        var l1 = doc.layers.add(); l1.name = "1";
        for (var i = doc.layers.length - 1; i >= 0; i--) {{
            var lay = doc.layers[i];
            if (lay != l1) lay.move(l1, ElementPlacement.PLACEATEND);
        }}
    }}
    function fastCopyLayer(srcLayer, destLayer, offX, offY) {{
        if (srcLayer.pageItems.length > 0) {{
            var tempGrp = srcLayer.groupItems.add();
            for (var i = srcLayer.pageItems.length - 1; i >= 0; i--) {{
                if (srcLayer.pageItems[i] != tempGrp) {{
                    srcLayer.pageItems[i].move(tempGrp, ElementPlacement.PLACEATEND);
                }}
            }}
            try {{
                var dup = tempGrp.duplicate(destLayer, ElementPlacement.PLACEATBEGINNING);
                dup.translate(offX, offY);
                while (dup.pageItems.length > 0) {{
                    dup.pageItems[0].move(destLayer, ElementPlacement.PLACEATBEGINNING);
                }}
                dup.remove();
                while (tempGrp.pageItems.length > 0) {{
                    tempGrp.pageItems[0].move(srcLayer, ElementPlacement.PLACEATBEGINNING);
                }}
                tempGrp.remove();
            }} catch(e) {{}}
        }}
        for (var j = 0; j < srcLayer.layers.length; j++) {{
            var sSub = srcLayer.layers[j];
            var dSub = destLayer.layers.add();
            dSub.name = sSub.name;
            fastCopyLayer(sSub, dSub, offX, offY);
        }}
    }}
    function processNextFileFast(masterDoc, srcPath, layerName, offX, offY) {{
        var srcDoc = openFileSafe(srcPath);
        var abData = [];
        for(var i=0; i<srcDoc.artboards.length; i++) abData.push({{rect: srcDoc.artboards[i].artboardRect, name: srcDoc.artboards[i].name}});
        masterDoc.activate();
        var mainLayer = masterDoc.layers.add(); mainLayer.name = layerName;
        for (var k = 0; k < srcDoc.layers.length; k++) {{
            var sLay = srcDoc.layers[k];
            var dLay = mainLayer.layers.add();
            dLay.name = sLay.name;
            fastCopyLayer(sLay, dLay, offX, offY);
        }}
        srcDoc.close(SaveOptions.DONOTSAVECHANGES);
        masterDoc.activate();
        for(var n=0; n<abData.length; n++){{
            var d = abData[n];
            var newAb = masterDoc.artboards.add([d.rect[0]+offX, d.rect[1]+offY, d.rect[2]+offX, d.rect[3]+offY]);
            newAb.name = "P" + layerName + "_" + d.name;
        }}
    }}
    function reorderArtboardsSafe(doc) {{
        var oldAbs = [];
        for (var i = 0; i < doc.artboards.length; i++) oldAbs.push({{rect: doc.artboards[i].artboardRect, name: doc.artboards[i].name}});
        var newOrder = [];
        for (var i = 0; i < oldAbs.length; i++) if (oldAbs[i].name.indexOf("Simulation") > -1) newOrder.push(oldAbs[i]);
        for (var i = 0; i < oldAbs.length; i++) if (oldAbs[i].name.indexOf("Simulation") === -1) newOrder.push(oldAbs[i]);
        for (var j = 0; j < newOrder.length; j++) doc.artboards.add(newOrder[j].rect).name = newOrder[j].name;
        var len = oldAbs.length;
        for (var k = 0; k < len; k++) doc.artboards[0].remove();
    }}
    main();
    """
    script_path = os.path.join(BASE_DIR, "run_merge_batch.jsx")
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(jsx_content)
    try:
        print(f"\n🔍 DEBUG: About to run merge script...")
        print(f"🔍 DEBUG: Script path: {script_path}")
        print(f"🔍 DEBUG: Files to merge: {files_list}")
        for fp in files_list:
            if not os.path.exists(fp):
                raise FileNotFoundError(f"Merge input missing: {fp}")
            print(f"🔍 DEBUG: File OK: {fp} ({os.path.getsize(fp)} bytes)")
        app = win32com.client.Dispatch("Illustrator.Application")
        print(f"🔍 DEBUG: Illustrator app connected")
        close_all_illustrator_documents(app)
        time.sleep(0.5)
        print(f"🔍 DEBUG: Running JavaScript file...")
        app.DoJavaScriptFile(script_path)
        print(f"🔍 DEBUG: JavaScript file executed")
        print(f"🔍 DEBUG: Number of open documents: {app.Documents.Count}")
        if app.Documents.Count > 0:
            doc = app.ActiveDocument
            # וידוא שהמסמך הנכון פעיל (temp_0 - קובץ האיחוד)
            doc_name = doc.Name
            print(f"DEBUG: Active document before save: {doc_name}")
            # אם יש כמה מסמכים פתוחים, נחפש את temp_0
            if "temp_0" not in doc_name and app.Documents.Count > 1:
                print(f"🔍 DEBUG: temp_0 not in active doc, searching...")
                for i in range(1, app.Documents.Count + 1):
                    try:
                        temp_doc = app.Documents[i]
                        print(f"🔍 DEBUG: Checking document {i}: {temp_doc.Name}")
                        if "temp_0" in temp_doc.Name:
                            doc = temp_doc
                            doc.Activate()
                            print(f"✅ DEBUG: Found and activated temp_0: {doc.Name}")
                            break
                    except Exception as e:
                        print(f"⚠️ DEBUG: Error checking document {i}: {e}")
                        continue
            # בדיקה כמה שכבות יש במסמך (כדי לראות אם התפריט הצד נוסף)
            try:
                layers_count = doc.Layers.Count
                print(f"🔍 DEBUG: Document has {layers_count} layers")
                # בדיקה אם יש שכבה בשם Sidebar_Layer
                has_sidebar = False
                for i in range(1, layers_count + 1):
                    try:
                        layer = doc.Layers[i]
                        if "Sidebar" in layer.Name:
                            has_sidebar = True
                            print(f"✅ DEBUG: Found Sidebar layer: {layer.Name}")
                            break
                    except:
                        continue
                if not has_sidebar:
                    print(f"⚠️ WARNING: No Sidebar layer found in document!")
            except Exception as e:
                print(f"⚠️ DEBUG: Error checking layers: {e}")
            pdf_options = win32com.client.Dispatch("Illustrator.PDFSaveOptions")
            pdf_options.PDFPreset = "[High Quality Print]"
            pdf_options.PreserveEditability = True
            final_path = output_pdf.replace("/", "\\")
            print(f"💾 DEBUG: Saving PDF to: {final_path}")
            doc.SaveAs(final_path, pdf_options)
            print(f"✅ DEBUG: PDF saved successfully")
            doc.Close(2)
            print(f"✅ DEBUG: Document closed")
    except Exception as e:
        print(f"Error during execution/save: {e}")
        import traceback
        traceback.print_exc()
        # --- MAIN ENTRY ---
if __name__ == "__main__":
    if len(sys.argv) > 1:
        try:
            # טעינת הנתונים שהגיעו מהקריאה
            full_data = json.loads(sys.argv[1])
            mode = full_data.get("mode", "standard")
            # שלב 1: חילוץ מספר הזמנה מהנתונים האמיתיים (לא משתנה קבוע)
            raw_order_id = str(full_data.get('order_id', '0000')).strip()
            # לקיחת 4 ספרות אחרונות בלבד
            current_order_4 = raw_order_id[-4:]
            # הדפסה לבדיקה בטרמינל (תוכלי לראות מה הוא באמת מזהה)
            print(f"DEBUG: Processing Order ID: {raw_order_id} | Final Name will be: {current_order_4} | mode={mode}")
            products = full_data.get('products', [])
            # שלב 2: יצירת תיקיית פלט (לפי ה-4 ספרות)
            order_folder = os.path.join(ORDERS_ROOT_DIR, current_order_4)
            if not os.path.exists(order_folder):
                os.makedirs(order_folder)
            # ניקוי ויצירה מחדש של תיקיית הקבצים הזמניים בצורה בטוחה
            if os.path.exists(TEMP_AI_DIR):
                try:
                    shutil.rmtree(TEMP_AI_DIR)
                    import time
                    time.sleep(0.5) # השהיה קלה כדי לתת למערכת ההפעלה לשחרר את התיקייה
                except:
                    pass
            # השינוי הקריטי: הוספת exist_ok=True
            os.makedirs(TEMP_AI_DIR, exist_ok=True)
            generated_files = []
            # שלב 3: עיבוד המוצרים
            is_wholesale = full_data.get('is_wholesale', False)
            print(f"\n{'='*50}")
            print(f"📦 WHOLESALE MODE: {is_wholesale}")
            if is_wholesale:
                print(f"   → השכבה 'information' תימחק מכל המוצרים")
            print(f"{'='*50}\n")
            if mode == "variable":
                from variable_print import download_variable_product_files, process_variable_product_to_temp

                if len(products) != 1:
                    print("⚠️ Variable mode requires exactly 1 product")
                else:
                    prod = products[0]
                    download_variable_product_files(prod, current_order_4)
                    ai_file = process_variable_product_to_temp(
                        prod, order_folder, is_wholesale, order_id=raw_order_id
                    )
                    if ai_file:
                        generated_files.append(ai_file)
            else:
                side_cache = {}
                download_all_side_files(products, current_order_4)
                for i, prod in enumerate(products):
                    ai_file = process_single_product_to_temp(
                        prod, i, order_folder, is_wholesale,
                        order_id=raw_order_id, products_all=products, side_cache=side_cache
                    )
                    if ai_file:
                        generated_files.append(ai_file)
                    else:
                        print(f"⚠️ Product {i+1} skipped – not included in merge")
            print(f"\n📊 Merge input: {len(generated_files)}/{len(products)} products")
            # שלב 4: איחוד לקובץ PDF סופי
            if generated_files:
                # הגדרת נתיב הקובץ הסופי
                final_pdf = os.path.join(order_folder, f"{current_order_4}.pdf")
                # אם הקובץ קיים, נוסיף מספר בסוגריים
                counter = 1
                while os.path.exists(final_pdf):
                    final_pdf = os.path.join(order_folder, f"{current_order_4} ({counter}).pdf")
                    counter += 1
                print(f"DEBUG: Saving Final PDF to: {final_pdf}")
                create_and_run_merge_script(generated_files, final_pdf, full_data)
            else:
                print("No files created.")
                sys.exit(1)
        except Exception as e:
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
