# -*- coding: utf-8 -*-

"""
main.py - רץ על תיקים עם הזמנות
גרסה ניידת - שומרת בתיקיית המסמכים
"""

from __future__ import annotations
import os
import shutil
import datetime
import sys
import concurrent.futures
import pythoncom
import json
import difflib 
import requests 
import base64
from typing import Dict, Any, Optional

# הגדרת קידוד פלט לוודא עברית תקינה בלוגים
sys.stdout.reconfigure(encoding='utf-8')

# נניח שהייבוא הבא קיים ועדכני
from illustrator_ops import open_and_color_template, place_and_simulate_print, update_size_label, delete_side_assets, save_pdf, clean_layout, apply_extra_colors
from vectorizer_ops import convert_to_svg

# --- הגדרות נתיבים דינמיות ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_DOWNLOAD_DIR = os.path.join(BASE_DIR, "temp_downloads") # תיקייה זמנית להורדות

# === השינוי מתחיל כאן ===
try:
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    # שליפת הנתונים למשתנים
    # אם כתוב נתיב בקובץ - ניקח אותו. אם לא - נלך למסמכים
    default_docs = os.path.join(os.path.expanduser("~"), "Documents", "Auto_Print_Output")
    SAVE_FOLDER = config.get('save_folder_path', default_docs) 
    IS_TEST_MODE = config.get('is_test_mode', False)

    print(f"Loaded config: Saving to {SAVE_FOLDER}, Test Mode: {IS_TEST_MODE}")

except FileNotFoundError:
    print("Error: config.json not found! Using defaults.")
    SAVE_FOLDER = os.path.join(os.path.expanduser("~"), "Documents", "Auto_Print_Output")
    IS_TEST_MODE = False

# עדכון הנתיב הראשי לפי מה שיצא לנו למעלה
ORDERS_ROOT_DIR = SAVE_FOLDER

if not os.path.exists(ORDERS_ROOT_DIR):
    os.makedirs(ORDERS_ROOT_DIR)

if not os.path.exists(TEMP_DOWNLOAD_DIR):
    os.makedirs(TEMP_DOWNLOAD_DIR)

TEMPLATES = {
    'Shirt': os.path.join(BASE_DIR, 'Simulations', 'Short.ai'), 
    'Sweater': os.path.join(BASE_DIR, 'Simulations', 'Sweater.ai'),
    'Hoodie': os.path.join(BASE_DIR, 'Simulations', 'Hoodie.ai'), 
    'Zippered Hoodie': os.path.join(BASE_DIR, 'Simulations', 'Zippered Hoodie.ai'), 
}

# פרטי API (וקטוריזציה)
API_ID = "vkd2vcts24ywdpk"
API_SECRET = "r20rqffqdcv6vj0ahukmiu9i8ma6ur4g0e1a5o9c7vugsoracpk8"

# --- מילון צבעים מורחב (השארתי את הנתונים שלך) ---
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
    'רויאל': '#1d4483', 'כחול רויאל': '#1d4483', 'תכלת': '#00aeef', 'טורקיז': '#029faa',
    'ים': '#40E0D0', 'פטרול': '#005f6a',
    'ירוק': '#8dc63f', 'ירוק בקבוק': '#064422', 'בקבוק': '#064422', 'ירוק תפוח': '#8DB600',
    'תפוח': '#8DB600', 'זיית': '#4f4e20', 'זית': '#4f4e20', 'ירוק זית': '#4f4e20', 'מנטה': '#98FF98',
    'ורוד': '#FFC0CB', 'ורוד בייבי': '#f1b1d0', 'ורוד ביבי': '#f1b1d0', 'ביבי': '#f1b1d0', 'בייבי': '#f1b1d0', 'פוקסיה': '#ec008c',
    'ורוד פוקסיה': '#ec008c', 'סגול': '#311d72', 'סגול כהה': '#4B0082', 'חציל': '#4B0082',
    'ליילך': '#C8A2C8', 'בורדו': '#8c191f', 'יין': '#800000',
}

# --- פונקציות עזר (ללא שינוי, לניקיון) ---

def get_contrasting_print_color(bg_hex):
    if not bg_hex: return '#FFFFFF'
    h = bg_hex.lstrip('#')
    try:
        r, g, b = tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
        luminance = (0.299 * r + 0.587 * g + 0.114 * b)
        return '#FFFFFF' if luminance < 128 else '#000000'
    except:
        return '#FFFFFF'

def get_hex_smart(name, return_none_on_fail=False):
    if not name or not isinstance(name, str): return None if return_none_on_fail else '#000000'
    name_clean = name.strip()
    if name_clean in EXTENDED_COLOR_MAP:
        return EXTENDED_COLOR_MAP[name_clean]
    matches = difflib.get_close_matches(name_clean, EXTENDED_COLOR_MAP.keys(), n=1, cutoff=0.5)
    if matches:
        print(f"DEBUG: Typo fixed: '{name_clean}' -> '{matches[0]}'")
        return EXTENDED_COLOR_MAP[matches[0]]
    return None if return_none_on_fail else '#000000'

def resolve_print_color(req_color_name, shirt_hex):
    txt = str(req_color_name).strip() if req_color_name else ""
    
    # 1. בדיקה במילון (תופס את 'צבעוני')
    found_val = get_hex_smart(txt, return_none_on_fail=True)
    if found_val == 'ORIGINAL':
        return None # לא צובעים
    if found_val:
        return found_val 

    # 2. בדיקה אם ריק -> שחור/לבן
    if not txt:
        return get_contrasting_print_color(shirt_hex)

    # 3. בדיקה מפורשת לשחור/לבן או ג'יבריש
    return get_contrasting_print_color(shirt_hex)

def get_hex(name):
    val = get_hex_smart(name)
    return val if val != 'ORIGINAL' else None

def get_print_colors(name):
    # הפונקציה הזו הוחזרה כדי למנוע את השגיאה שראית בתמונה
    # היא לא בשימוש בלוגיקה החדשה, אבל חייבת להיות קיימת
    return '#000000', '#000000'

def get_unique_filename(path):
    if not os.path.exists(path):
        return path
        
    base, ext = os.path.splitext(path)
    counter = 1
    while True:
        new_path = f"{base} ({counter}){ext}"
        if not os.path.exists(new_path):
            return new_path
        counter += 1

# פונקציה להורדת תמונה (תומכת גם בקישור וגם בקובץ שהועלה ידנית)
def download_image(url_or_base64, filename_prefix):
    try:
        # מקרה א': המשתמש העלה קובץ (Base64)
        if url_or_base64.startswith('data:'):
            print(f"Processing uploaded file for: {filename_prefix}")
            
            # חילוץ המידע מהפורמט: data:image/png;base64,.....
            header, encoded = url_or_base64.split(',', 1)
            
            # זיהוי סיומת
            file_ext = '.png' # ברירת מחדל
            
            # --- הוספנו זיהוי ל-SVG כאן ---
            if 'image/svg+xml' in header:
                file_ext = '.svg'
            elif 'image/jpeg' in header or 'image/jpg' in header:
                file_ext = '.jpg'
            elif 'application/pdf' in header:
                file_ext = '.pdf'
            
            final_filename = f"{filename_prefix}{file_ext}"
            path = os.path.join(TEMP_DOWNLOAD_DIR, final_filename)
            
            # המרה חזרה לקובץ ושמירה
            with open(path, 'wb') as f:
                f.write(base64.b64decode(encoded))
                
            return path

        # מקרה ב': קובץ קיים במערכת (URL רגיל)
        else:
            print(f"Downloading URL: {url_or_base64}")
            
            # ניקוי פרמטרים מיותרים מה-URL אם יש
            clean_url = url_or_base64.split('#')[0] 
            
            ext = ".png" # ברירת מחדל
            lower_url = clean_url.lower()
            
            # --- הוספנו זיהוי ל-SVG גם כאן ---
            if '.svg' in lower_url:
                ext = ".svg"
            elif '.pdf' in lower_url: 
                ext = ".pdf"
            elif '.jpg' in lower_url or '.jpeg' in lower_url: 
                ext = ".jpg"
            
            final_filename = f"{filename_prefix}{ext}"
            path = os.path.join(TEMP_DOWNLOAD_DIR, final_filename)

            response = requests.get(clean_url, stream=True)
            if response.status_code == 200:
                with open(path, 'wb') as f:
                    shutil.copyfileobj(response.raw, f)
                return path
            else:
                print(f"Failed to download. Status: {response.status_code}")
                return None

    except Exception as e:
        print(f"Error saving image: {e}")
        return None

# 📢 פונקציה מעודכנת: מטפלת במתג 'ללא וקטור' 
def vec_single(d: Dict[str, Any], f: str, id: str, sec: str) -> Optional[str]:
    if not d.get('exists'): return None
    
    # 📢 קליטת המתג החדש
    skip_vector = d.get('no_vectorization', False) 

    if not d.get('file') or not os.path.exists(d['file']): 
        print(f"File missing for vectorization: {d.get('file')}")
        return None
    
    original_src_path = d['file']
    # יצירת שם יעד (מבוסס על קובץ המקור)
    original_dst = os.path.join(f, f"{d['prefix']}_{os.path.basename(original_src_path)}")
    dst = get_unique_filename(original_dst)

    # --- לוגיקת וקטוריזציה מותנית ---
    if skip_vector:
        print(f"V Skipping vectorization for {d['prefix']}. Using original file (Raster).")
        # מעתיקים את הקובץ המקורי (PNG/JPG) ישירות לתיקיית היעד
        shutil.copy(original_src_path, dst)
        return dst # מחזירים את הנתיב לקובץ הרסטר המקורי
    else:
        # תהליך וקטוריזציה רגיל
        print(f"V Starting vectorization for {d['prefix']}...")
        shutil.copy(original_src_path, dst) # העתקה לתיקיית היעד לפני המרה
        # ה-convert_to_svg אמור להחזיר את הנתיב לקובץ ה-SVG שנוצר
        return convert_to_svg(dst, id, sec) 
    # ---------------------------------

def clean_temp_folder():
    """מוחקת את כל הקבצים הזמניים בתיקיית ההורדות"""
    try:
        if os.path.exists(TEMP_DOWNLOAD_DIR):
            # מחיקת התיקייה וכל תוכנה
            shutil.rmtree(TEMP_DOWNLOAD_DIR)
            # יצירה מחדש של התיקייה ריקה לפעם הבאה
            os.makedirs(TEMP_DOWNLOAD_DIR)
            print("V Temp folder cleaned.")
    except Exception as e:
        print(f"Warning: Could not clean temp folder: {e}")

def process_order(order: Dict[str, Any]):
    pythoncom.CoInitialize() 
    doc = None
    app = None
    
    try:
        oid = str(order['order_id'])
        
        # === התיקון: שימוש ב-4 ספרות אחרונות ===
        short_name = oid[-4:]
        # =======================================

        prod = order.get('product_type', 'Shirt')
        print(f"\n=== Processing {oid} -> Folder: {short_name} ===")
        
        # יצירת התיקייה לפי השם הקצר
        folder = os.path.join(ORDERS_ROOT_DIR, short_name)
        if not os.path.exists(folder): os.makedirs(folder)
        
        t_path = TEMPLATES.get(prod)
        
        if not t_path or not os.path.exists(t_path): 
            print(f"X Template not found: {t_path}")
            return

        sides = ['front', 'back', 'right_sleeve', 'left_sleeve']
        svgs = {}
        
        # 📢 המרה ל-SVG או העתקת הרסטר (מעודכן עם order.get)
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
            fs = {ex.submit(vec_single, order.get(s, {}), folder, API_ID, API_SECRET): s for s in sides}
            for f in concurrent.futures.as_completed(fs):
                svgs[fs[f]] = f.result()

        if IS_TEST_MODE:
            print("--- TEST MODE ON ---")
            print("שומר קבצי ביניים של ווקטור לבדיקה...")
            # כאן את יכולה להוסיף את הפקודה שרצית שתקרה רק בבדיקות
        
        try:
            # קבלת צבע החולצה הראשי
            col = get_hex_smart(order.get('product_color_hebrew'))
            doc, app = open_and_color_template(t_path, col, prod)

            # === הוספה חדשה: טיפול בצבעים נוספים (ריבועים) ===
            extra_heb_names = order.get('extra_colors_hebrew', [])
            extra_hex_list = []

            if extra_heb_names:
                print(f"Processing extra colors: {extra_heb_names}")
                for name in extra_heb_names:
                    h = get_hex_smart(name)
                    if h and h != 'ORIGINAL':
                        extra_hex_list.append(h)
            
            apply_extra_colors(app, extra_hex_list)
            # =================================================

        except Exception as e:
            print(f"X AI Initialization Error: {e}")
            if app: 
                try: app.Quit()
                except: pass
            return
            
        # 📢 הרצת המיקום והסימולציה לכל צד (מעודכן עם is_raster)
        for s in sides:
            d = order.get(s, {}) # שימוש ב-get כדי לוודא שקיימים נתונים
            if d.get('exists') and svgs.get(s):
                
                # 📢 קליטת המתג החדש
                is_raster = d.get('no_vectorization', False)
                
                # --- התיקון החכם כאן ---
                final_color = resolve_print_color(d.get('req_color_hebrew'), col)
                
                # ברירת מחדל: הדפסה והדמיה זהים
                cs = final_color # צבע להדמיה
                cp = final_color # צבע להדפסה

                # תנאי מיוחד ללבן:
                if final_color == '#FFFFFF':
                    cp = '#000000'
                # -----------------------

                # 📢 העברת הדגל החדש לפונקציית הסימולציה ב-Illustrator
                # אם is_raster=True, illustrator_ops צריך לדעת לא לעשות הסרת רקע
                w = place_and_simulate_print(doc, app, svgs[s], d['prefix'], d['category'], cp, cs, is_raster=is_raster)

                if w > 0: 
                    update_size_label(doc, app, d['label'], w, d.get('heb', ''))

        # === לוגיקה לשרוולים (ללא שינוי) ===
        rs_data = order.get('right_sleeve', {})
        ls_data = order.get('left_sleeve', {})
        rs_exists = rs_data.get('exists', False)
        ls_exists = ls_data.get('exists', False)

        if not rs_exists and not ls_exists:
            delete_side_assets(doc, app, "Print_Sleeves", "size_Right_Sleeve")
            try: app.DoJavaScript("try{app.activeDocument.textFrames.getByName('size_Left_Sleeve').remove();}catch(e){}")
            except: pass

        elif not rs_exists and ls_exists:
            try: app.DoJavaScript("try{app.activeDocument.textFrames.getByName('size_Right_Sleeve').remove();}catch(e){}")
            except: pass

        elif rs_exists and not ls_exists:
            try: app.DoJavaScript("try{app.activeDocument.textFrames.getByName('size_Left_Sleeve').remove();}catch(e){}")
            except: pass

        if not order['front'].get('exists'):
            delete_side_assets(doc, app, "Print_Front", "size_Front")
        if not order['back'].get('exists'):
            delete_side_assets(doc, app, "Print_Back", "size_Back")

        # === ניקוי ריבועי העזר (ללא שינוי) ===
        print("Cleaning up sizing boxes...")
        clean_layout(app)
        # =====================================

        # שמירת PDF
        base_pdf_path = os.path.join(folder, f"{short_name}.pdf")
        final_pdf_path = get_unique_filename(base_pdf_path)
        
        save_pdf(doc, final_pdf_path)
        print(f"V Finished! Saved to: {final_pdf_path}")
        
    except Exception as general_e:

        print(f"!!! FATAL ERROR processing {order.get('order_id', 'Unknown')}: {general_e}")

        if app: 
            try: app.Quit()
            except: pass
            
    finally:
        clean_temp_folder()
        pythoncom.CoUninitialize()


# --- נקודת הכניסה ---
if __name__ == "__main__":
    if len(sys.argv) > 1:
        try:
            # 1. קריאת ה-JSON
            json_str = sys.argv[1]
            order_data = json.loads(json_str)
            
            # === התיקון: חילוץ 4 ספרות אחרונות ===
            full_id = str(order_data.get('order_id', '0000'))
            short_id = full_id[-4:] 
            print(f"Processing Order: {full_id} (Short: {short_id})")
            # =====================================

            # 2. הורדת קבצים
            for loc in ['front', 'back', 'right_sleeve', 'left_sleeve']:
                loc_data = order_data.get(loc, {})
                if loc_data.get('exists'):
                    url = loc_data.get('file_url')
                    if url:
                        # זיהוי סיומת
                        ext = ".png"
                        lower_url = url.lower()
                        if '.pdf' in lower_url: ext = ".pdf"
                        elif '.jpg' in lower_url or '.jpeg' in lower_url: ext = ".jpg"
                        elif '.svg' in lower_url: ext = ".svg"
                        
                        # === שינוי: שם הקובץ עם המזהה הקצר ===
                        filename_prefix = f"{short_id}_{loc}"
                        # =====================================
                        
                        print(f"Downloading {loc} as {ext}...")
                        local_path = download_image(url, filename_prefix)
                        
                        if local_path:
                            order_data[loc]['file'] = local_path
                        else:
                            print(f"Error: Could not download file for {loc}")
                            order_data[loc]['exists'] = False

            # 3. הפעלת העיבוד באילוסטרייטור
            process_order(order_data)

        except Exception as e:
            print(f"Error in main execution: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("No order data provided.")