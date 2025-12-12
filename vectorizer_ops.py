# -*- coding: utf-8 -*-
"""
vectorizer_ops.py
פשוט ויציב: מעלה קובץ ל־vectorizer.ai ומוריד SVG.
כולל מנגנון Caching (זיכרון) למניעת המרות כפולות.
"""
import requests
import os
import time
import hashlib
import json
import shutil

# מיקום קובץ הזיכרון - יישמר באותה תיקייה של הסקריפט
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(BASE_DIR, "vector_cache.json")

# --- שינוי 1: טעינת הגדרות (כדי לדעת אם אנחנו בבדיקה) ---
try:
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    IS_TEST_MODE = config.get('is_test_mode', False)
except:
    # אם אין קובץ הגדרות, נניח שאנחנו לא בבדיקה (ברירת מחדל לייצור)
    IS_TEST_MODE = False
# -------------------------------------------------------

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

def convert_to_svg(image_path, api_id, api_secret, retries=1):
    if not os.path.exists(image_path): return None

    # --- שלב 0: בדיקה אם הקובץ הוא כבר SVG ---
    if image_path.lower().endswith(".svg"):
        print(f"   V Input is already SVG. Skipping conversion -> {os.path.basename(image_path)}")
        return image_path

    # --- שלב 1: בדיקת זיכרון (Cache) ---
    try:
        file_hash = get_file_hash(image_path)
        cache = load_cache()

        if file_hash in cache:
            existing_svg_path = cache[file_hash]
            
            # בדיקה אם קובץ ה-SVG הישן עדיין קיים במחשב
            if os.path.exists(existing_svg_path):
                print(f"   V Image recognized from history! Copying existing SVG...")
                output = os.path.splitext(image_path)[0] + ".svg"
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
        
        # --- שינוי 2: בניית המידע שנשלח ל-API ---
        # אנחנו מגדירים את הפרמטרים במשתנה נפרד
        api_data = {}
        
        # רק אם אנחנו במצב בדיקה - נוסיף את השורה הזו
        if IS_TEST_MODE:
            api_data['mode'] = 'test'
        # ----------------------------------------

        with open(image_path, 'rb') as f:
            response = requests.post(
                'https://vectorizer.ai/api/v1/vectorize',
                files={'image': f},
                data=api_data,  # כאן אנחנו שולחים את המשתנה שיצרנו למעלה
                auth=(api_id, api_secret),
                timeout=60
            )
            
        if response.status_code != 200:
            print(f"X Vectorizer API returned {response.status_code}")
            return None
        
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