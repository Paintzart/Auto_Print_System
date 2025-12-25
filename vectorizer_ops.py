# -*- coding: utf-8 -*-

"""
vectorizer_ops.py
פשוט ויציב: מעלה קובץ ל־vectorizer.ai ומוריד SVG.
כולל מנגנון Caching (זיכרון) למניעת המרות כפולות.
כולל תמיכה בהמרת PDF לתמונה לפני שליחה.
"""

import requests
import os
import time
import hashlib
import json
import shutil
import fitz  # ספרייה לטיפול ב-PDF (PyMuPDF)

# מיקום קובץ הזיכרון - יישמר באותה תיקייה של הסקריפט
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(BASE_DIR, "vector_cache.json")

# --- טעינת הגדרות ---
try:
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    IS_TEST_MODE = config.get('is_test_mode', False)
except:
    IS_TEST_MODE = False
# --------------------

def load_cache():
    """טוען את מסד הנתונים של ההמרות"""
    if not os.path.exists(CACHE_FILE):
        return {}
    try:
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

def save_cache(cache_data):
    """שומר את מסד הנתונים"""
    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache_data, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Warning: Could not save cache: {e}")

def get_file_hash(path):
    """מייצר חתימה ייחודית (Hash) לפי תוכן הקובץ"""
    sha256 = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            sha256.update(chunk)
    return sha256.hexdigest()

def convert_pdf_to_png(pdf_path):
    """ממיר דף ראשון של PDF לתמונה איכותית (PNG)"""
    try:
        print(f"   > Converting PDF to PNG for vectorization: {os.path.basename(pdf_path)}")
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)  # לוקח את העמוד הראשון
        
        # הגדרת רזולוציה גבוהה (Zoom x3 = 216 DPI בערך)
        # זה חשוב כדי שהווקטור יצא איכותי ומדויק
        matrix = fitz.Matrix(3.0, 3.0) 
        pix = page.get_pixmap(matrix=matrix, alpha=True)
        
        output_png = pdf_path.replace(".pdf", ".png")
        # אם השם לא השתנה (במקרה מוזר), נוסיף סיומת
        if output_png == pdf_path:
            output_png = pdf_path + ".png"
            
        pix.save(output_png)
        doc.close()
        return output_png
    except Exception as e:
        print(f"X Error converting PDF to PNG: {e}")
        return None

def convert_to_svg(image_path, api_id, api_secret, retries=1):
    if not os.path.exists(image_path): return None

    # --- שלב 0: בדיקה אם הקובץ הוא כבר SVG ---
    if image_path.lower().endswith(".svg"):
        print(f"   V Input is already SVG. Skipping conversion -> {os.path.basename(image_path)}")
        return image_path

    # --- שלב 0.5: בדיקה אם זה PDF והמרה לתמונה ---
    # ה-API לא מקבל PDF, אז אנחנו הופכים אותו לתמונה רגע לפני
    if image_path.lower().endswith(".pdf"):
        png_path = convert_pdf_to_png(image_path)
        if png_path and os.path.exists(png_path):
            image_path = png_path  # מעדכנים את הנתיב לעבוד מול ה-PNG החדש
        else:
            print("X Failed to convert PDF, trying to send original (might fail)...")

    # --- שלב 1: בדיקת זיכרון (Cache) ---
    try:
        file_hash = get_file_hash(image_path)
        cache = load_cache()

        if file_hash in cache:
            existing_svg_path = cache[file_hash]
            
            # בדיקה אם קובץ ה-SVG הישן עדיין קיים במחשב
            if os.path.exists(existing_svg_path):
                print(f"   V Image recognized from history! Copying existing SVG...")
                # כאן אנחנו שומרים את ה-SVG עם השם המקורי (בלי ה-PNG באמצע)
                base_original_name = os.path.splitext(image_path)[0].replace(".png", "") 
                output = base_original_name + ".svg"
                
                shutil.copy(existing_svg_path, output)
                print(f"   V Copied locally -> {os.path.basename(output)}")
                return output
            else:
                del cache[file_hash]
                save_cache(cache)
    
    except Exception as e:
        print(f"Warning: Cache check failed ({e}), proceeding to API.")

    # --- שלב 2: שליחה ל-API ---
    try:
        print(f"   -> Uploading {os.path.basename(image_path)}...")
        
        api_data = {}
        if IS_TEST_MODE:
            api_data['mode'] = 'test'

        with open(image_path, 'rb') as f:
            response = requests.post(
                'https://vectorizer.ai/api/v1/vectorize',
                files={'image': f},
                data=api_data,
                auth=(api_id, api_secret),
                timeout=60
            )
            
        if response.status_code != 200:
            print(f"X Vectorizer API returned {response.status_code}")
            return None
        
        # יצירת שם לקובץ SVG הסופי (מסירים את סיומת ה-png הזמנית אם יש)
        output = os.path.splitext(image_path)[0] + ".svg"
        
        with open(output, 'wb') as fo:
            fo.write(response.content)
        print(f"   V SVG Saved -> {os.path.basename(output)}")

        # --- שלב 3: שמירה בזיכרון לעתיד ---
        try:
            cache = load_cache()
            cache[file_hash] = output
            save_cache(cache)
        except Exception as e:
            print(f"Warning: Could not update cache: {e}")

        return output

    except Exception as e:
        print(f"X convert_to_svg Error: {e}")
        if retries > 0:
            time.sleep(1)
            return convert_to_svg(image_path, api_id, api_secret, retries-1)
        return None