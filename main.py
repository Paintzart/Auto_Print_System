# -*- coding: utf-8 -*-
"""
main.py - הגרסה היציבה והמלאה (ללא Threading, עם סידור דפים בטוח)
"""

from __future__ import annotations
import os
import shutil
import sys
import json
import base64, difflib
import win32com.client
import pythoncom
import requests
from typing import Dict, Any, Optional

# הגדרת קידוד
sys.stdout.reconfigure(encoding='utf-8')

# ייבוא הפונקציות הגרפיות
from illustrator_ops import run_jsx, open_and_color_template, place_and_simulate_print, update_size_label, delete_side_assets, save_pdf, clean_layout, apply_extra_colors
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
    'רויאל': '#1d4483', 'כחול רויאל': '#1d4483', 'תכלת': '#00aeef', 'טורקיז': '#029faa',
    'ים': '#40E0D0', 'פטרול': '#005f6a',
    'ירוק': '#8dc63f', 'ירוק בקבוק': '#064422', 'בקבוק': '#064422', 'ירוק תפוח': '#8DB600',
    'תפוח': '#8DB600', 'זיית': '#4f4e20', 'זית': '#4f4e20', 'ירוק זית': '#4f4e20', 'מנטה': '#98FF98',
    'ורוד': '#FFC0CB', 'ורוד בייבי': '#f1b1d0', 'ורוד ביבי': '#f1b1d0', 'ביבי': '#f1b1d0', 'בייבי': '#f1b1d0', 'פוקסיה': '#ec008c',
    'ורוד פוקסיה': '#ec008c', 'סגול': '#311d72', 'סגול כהה': '#4B0082', 'חציל': '#4B0082',
    'ליילך': '#C8A2C8', 'בורדו': '#8c191f', 'יין': '#800000',
}

# -----------------------------------------------------------
# פונקציות עזר (ללא שימוש ב-concurrent)
# -----------------------------------------------------------
def get_contrasting_print_color(bg_hex):
    if not bg_hex: return '#FFFFFF'
    h = bg_hex.lstrip('#')
    try:
        r, g, b = tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
        luminance = (0.299 * r + 0.587 * g + 0.114 * b)
        return '#FFFFFF' if luminance < 128 else '#000000'
    except: return '#FFFFFF'

def get_hex_smart(name, return_none_on_fail=False):
    if not name or not isinstance(name, str): return None if return_none_on_fail else '#000000'
    name_clean = name.strip()
    if name_clean in EXTENDED_COLOR_MAP: return EXTENDED_COLOR_MAP[name_clean]
    matches = difflib.get_close_matches(name_clean, EXTENDED_COLOR_MAP.keys(), n=1, cutoff=0.5)
    if matches: return EXTENDED_COLOR_MAP[matches[0]]
    return None if return_none_on_fail else '#000000'

def resolve_print_color(req, shirt):
    txt = str(req).strip() if req else ""
    found = get_hex_smart(txt, True)
    if found == 'ORIGINAL': return None
    if found: return found
    return get_contrasting_print_color(shirt)

def get_hex(name):
    val = get_hex_smart(name)
    return val if val != 'ORIGINAL' else None

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
        if url_or_base64.startswith('data:'):
            header, encoded = url_or_base64.split(',', 1)
            file_ext = '.png'
            if 'image/svg' in header: file_ext = '.svg'
            elif 'pdf' in header: file_ext = '.pdf'
            path = os.path.join(TEMP_DOWNLOAD_DIR, f"{filename_prefix}{file_ext}")
            with open(path, 'wb') as f: f.write(base64.b64decode(encoded))
            return path
        elif os.path.exists(url_or_base64) or (len(url_or_base64)>1 and url_or_base64[1]==':'):
            if not os.path.exists(url_or_base64): return None
            _, ext = os.path.splitext(url_or_base64)
            path = os.path.join(TEMP_DOWNLOAD_DIR, f"{filename_prefix}{ext or '.png'}")
            if os.path.abspath(url_or_base64) != os.path.abspath(path): shutil.copy(url_or_base64, path)
            return path
        elif url_or_base64.startswith('http'):
            ext = ".png"
            if '.pdf' in url_or_base64.lower(): ext = ".pdf"
            elif '.svg' in url_or_base64.lower(): ext = ".svg"
            path = os.path.join(TEMP_DOWNLOAD_DIR, f"{filename_prefix}{ext}")
            r = requests.get(url_or_base64, stream=True)
            if r.status_code == 200:
                with open(path, 'wb') as f: shutil.copyfileobj(r.raw, f)
                return path
    except: pass
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

# -----------------------------------------------------------
# עיבוד בודד (ללא שינוי, זה עובד טוב)
# -----------------------------------------------------------
def process_single_product_to_temp(order, idx, folder):
    pythoncom.CoInitialize()
    doc = None
    app = None
    try:
        prod = order.get('product_type', 'Shirt')
        print(f"\n>> Processing Product {idx+1}: {prod}")
        t_path = TEMPLATES.get(prod)
        if not t_path or not os.path.exists(t_path): return None

        sides = ['front', 'back', 'right_sleeve', 'left_sleeve']
        svgs = {}
        # תיקון שגיאת concurrent: שימוש בלולאה רגילה
        for s in sides:
            res = vec_single(order.get(s, {}), folder, API_ID, API_SECRET)
            if res: svgs[s] = res

        col_raw = order.get('product_color_hebrew', "")
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
            d = order.get(s, {})
            if d.get('exists') and svgs.get(s):
                is_r = d.get('no_vectorization', False)
                fc = resolve_print_color(d.get('req_color_hebrew'), h1)
                cp = fc if fc!='#FFFFFF' else '#000000'
                w = place_and_simulate_print(doc, app, svgs[s], d['prefix'], d['category'], cp, fc, is_r)
                if w>0: update_size_label(doc, app, d['label'], w, d.get('heb',''))

        if not order.get('right_sleeve', {}).get('exists') and not order.get('left_sleeve', {}).get('exists'):
            delete_side_assets(doc, app, "Print_Sleeves", "size_Right_Sleeve")
            run_jsx(app, "try{app.activeDocument.textFrames.getByName('size_Left_Sleeve').remove();}catch(e){}")
        elif not order.get('right_sleeve', {}).get('exists'):
            run_jsx(app, "try{app.activeDocument.textFrames.getByName('size_Right_Sleeve').remove();}catch(e){}")
        elif not order.get('left_sleeve', {}).get('exists'):
            run_jsx(app, "try{app.activeDocument.textFrames.getByName('size_Left_Sleeve').remove();}catch(e){}")
        
        if not order.get('front', {}).get('exists'): delete_side_assets(doc, app, "Print_Front", "size_Front")
        if not order.get('back', {}).get('exists'): delete_side_assets(doc, app, "Print_Back", "size_Back")
        
        clean_layout(app)

        out_name = f"temp_{idx}.ai"
        out_path = os.path.join(TEMP_AI_DIR, out_name)
        doc.SaveAs(out_path)
        doc.Close(2)
        print(f"   > Saved: {out_name}")
        return out_path

    except Exception as e:
        print(f"Error processing {idx}: {e}")
        if app: 
            try: app.Quit() 
            except: pass
        return None
    finally:
        pythoncom.CoUninitialize()

# -----------------------------------------------------------
# פונקציית האיחוד: ה-Super Script (חישוב גובה ורוחב דינמי חכם)
# -----------------------------------------------------------
def create_and_run_merge_script(files_list, output_pdf):
    pythoncom.CoInitialize()
    
    if not files_list: return
    
    js_files = [f.replace("\\", "/") for f in files_list]
    js_output = output_pdf.replace("\\", "/")
    
    # === ה-JSX של הגריד החכם (4 בשורה + גובה דינמי) ===
    jsx_content = f"""
    #target illustrator
    
    var files = {json.dumps(js_files)};
    var finalPath = "{js_output}";
    
    function main() {{
        if (files.length === 0) return;
        
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        app.preferences.setBooleanPreference("preventPasteRemembersLayers", true); 
        
        // 1. פתיחת המאסטר (הקובץ הראשון)
        var masterFile = new File(files[0]);
        var masterDoc = app.open(masterFile);
        masterDoc.activate();
        
        organizeMasterContent(masterDoc);
        
        // === חישוב דינמי של גודל המוצר ===
        // מודדים את הגובה האמיתי של הקובץ כדי לדעת כמה לרדת שורה
        var metrics = calculateLayoutMetrics(masterDoc);
        var productWidth = metrics.width;
        var productHeight = metrics.height;
        
        // הגדרות גריד
        var GAP_X = 100; 
        var GAP_Y = 200; 
        var COLS = 4;    // 4 מוצרים בשורה
        
        var STEP_X = productWidth + GAP_X;
        var STEP_Y = productHeight + GAP_Y;
        
        // 2. לולאה על שאר הקבצים ומיקום לפי הגריד
        for (var i = 1; i < files.length; i++) {{
            var srcPath = files[i];
            
            var col = i % COLS; 
            var row = Math.floor(i / COLS);
            
            var offsetX = col * STEP_X;
            var offsetY = -(row * STEP_Y); // ירידה למטה לפי הגובה המחושב
            
            processNextFileFast(masterDoc, srcPath, (i+1).toString(), offsetX, offsetY);
        }}
        
        // 3. סידור דפים ומחיקת מיותרים
        reorderArtboardsSafe(masterDoc);
        
        // 4. שמירה
        saveFinalPDF(masterDoc, finalPath);
    }}
    
    function calculateLayoutMetrics(doc) {{
        var minX = Infinity;
        var maxX = -Infinity;
        var maxY = -Infinity; 
        var minY = Infinity;  
        
        for (var i = 0; i < doc.artboards.length; i++) {{
            var r = doc.artboards[i].artboardRect; 
            if (r[0] < minX) minX = r[0];
            if (r[2] > maxX) maxX = r[2];
            if (r[1] > maxY) maxY = r[1]; 
            if (r[3] < minY) minY = r[3]; 
        }}
        
        return {{
            width: Math.abs(maxX - minX),
            height: Math.abs(maxY - minY)
        }};
    }}
    
    function organizeMasterContent(doc) {{
        app.executeMenuCommand('unlockAll');
        app.executeMenuCommand('showAll');
        
        var l1 = doc.layers.add();
        l1.name = "1";
        
        for (var i = doc.layers.length - 1; i >= 0; i--) {{
            var lay = doc.layers[i];
            if (lay != l1) {{
                lay.move(l1, ElementPlacement.PLACEATEND);
            }}
        }}
    }}
    
    function fastCopyLayer(srcLayer, destLayer, offX, offY) {{
        if (srcLayer.pageItems.length > 0) {{
            var tempGroup = srcLayer.groupItems.add();
            var items = srcLayer.pageItems;
            for (var i = items.length - 1; i >= 0; i--) {{
                if (items[i] != tempGroup) {{
                    try {{ items[i].move(tempGroup, ElementPlacement.PLACEATEND); }} catch(e) {{}}
                }}
            }}
            try {{
                var dupGroup = tempGroup.duplicate(destLayer, ElementPlacement.PLACEATBEGINNING);
                dupGroup.translate(offX, offY);
                dupGroup.ungroup();
            }} catch(e) {{}}
        }}
        for (var j = srcLayer.layers.length - 1; j >= 0; j--) {{
            var srcSub = srcLayer.layers[j];
            var destSub = destLayer.layers.add();
            destSub.name = srcSub.name;
            fastCopyLayer(srcSub, destSub, offX, offY);
        }}
    }}
    
    function recursiveUnlock(layersCollection) {{
        for (var i = 0; i < layersCollection.length; i++) {{
            var lay = layersCollection[i];
            lay.locked = false;
            lay.visible = true;
            var items = lay.pageItems;
            for (var j = 0; j < items.length; j++) {{
                try {{ items[j].locked = false; items[j].hidden = false; }} catch(e) {{}}
            }}
            if (lay.layers.length > 0) {{
                recursiveUnlock(lay.layers);
            }}
        }}
    }}
    
    function processNextFileFast(masterDoc, srcPath, layerName, offX, offY) {{
        var srcDoc = app.open(new File(srcPath));
        srcDoc.activate();
        
        app.executeMenuCommand('unlockAll');
        app.executeMenuCommand('showAll');
        recursiveUnlock(srcDoc.layers); 
        
        var abData = [];
        for(var i=0; i<srcDoc.artboards.length; i++){{
            abData.push({{rect: srcDoc.artboards[i].artboardRect, name: srcDoc.artboards[i].name}});
        }}
        
        masterDoc.activate();
        var mainLayer = masterDoc.layers.add();
        mainLayer.name = layerName;
        
        var srcLayers = srcDoc.layers;
        for (var k = srcLayers.length - 1; k >= 0; k--) {{
            var sLay = srcLayers[k];
            var dLay = mainLayer.layers.add();
            dLay.name = sLay.name;
            fastCopyLayer(sLay, dLay, offX, offY);
        }}
        
        srcDoc.close(SaveOptions.DONOTSAVECHANGES);
        
        masterDoc.activate();
        for(var n=0; n<abData.length; n++){{
            var d = abData[n];
            var newLeft = d.rect[0] + offX;
            var newTop = d.rect[1] + offY;
            var newRight = d.rect[2] + offX;
            var newBottom = d.rect[3] + offY;
            
            var newAb = masterDoc.artboards.add([newLeft, newTop, newRight, newBottom]);
            newAb.name = "P" + layerName + "_" + d.name;
        }}
    }}
    
    function reorderArtboardsSafe(doc) {{
        var oldAbs = [];
        for (var i = 0; i < doc.artboards.length; i++) {{
            oldAbs.push({{
                rect: doc.artboards[i].artboardRect,
                name: doc.artboards[i].name
            }});
        }}
        
        var newOrder = [];
        for (var i = 0; i < oldAbs.length; i++) {{
            if (oldAbs[i].name.indexOf("Simulation") > -1) newOrder.push(oldAbs[i]);
        }}
        for (var i = 0; i < oldAbs.length; i++) {{
            if (oldAbs[i].name.indexOf("Simulation") === -1) newOrder.push(oldAbs[i]);
        }}
        
        for (var i = 0; i < newOrder.length; i++) {{
            var n = doc.artboards.add(newOrder[i].rect);
            n.name = newOrder[i].name;
        }}
        
        var deleteCount = oldAbs.length;
        for (var k = 0; k < deleteCount; k++) {{
            doc.artboards[0].remove();
        }}
    }}
    
    function saveFinalPDF(doc, path) {{
        var opts = new PDFSaveOptions();
        opts.preserveEditability = true;
        var f = new File(path);
        doc.saveAs(f, opts);
        doc.close(SaveOptions.DONOTSAVECHANGES);
    }}
    
    main();
    """
    
    script_path = os.path.join(BASE_DIR, "run_merge_batch.jsx")
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(jsx_content)
        
    print(f"Generated Merge Script: {script_path}")
    
    try:
        app = win32com.client.Dispatch("Illustrator.Application")
        app.DoJavaScriptFile(script_path)
        print("V Merge Script Completed.")
    except Exception as e:
        print(f"Error executing JSX: {e}")                  
# --- MAIN ENTRY ---
if __name__ == "__main__":
    if len(sys.argv) > 1:
        try:
            full_data = json.loads(sys.argv[1])
            order_id = str(full_data.get('order_id', '0000'))
            products = full_data.get('products', [])
            short_id = order_id[-4:]
            
            print(f"Starting Batch: {order_id} ({len(products)} items)")
            
            order_folder = os.path.join(ORDERS_ROOT_DIR, short_id)
            if not os.path.exists(order_folder): os.makedirs(order_folder)
            
            try: 
                if os.path.exists(TEMP_AI_DIR): shutil.rmtree(TEMP_AI_DIR)
            except: pass
            if not os.path.exists(TEMP_AI_DIR): os.makedirs(TEMP_AI_DIR)
            
            generated_files = []
            for i, prod in enumerate(products):
                for loc in ['front', 'back', 'right_sleeve', 'left_sleeve']:
                    loc_d = prod.get(loc, {})
                    if loc_d.get('exists') and loc_d.get('file_url'):
                        path = download_image(loc_d['file_url'], f"{short_id}_{i}_{loc}")
                        if path: loc_d['file'] = path
                
                ai_file = process_single_product_to_temp(prod, i, order_folder)
                if ai_file: generated_files.append(ai_file)
            
            if generated_files:
                base_name = short_id
                pdf_name = f"{base_name}.pdf"
                final_pdf = os.path.join(order_folder, pdf_name)

                counter = 1
                while os.path.exists(final_pdf):
                    pdf_name = f"{base_name} ({counter}).pdf"
                    final_pdf = os.path.join(order_folder, pdf_name)
                    counter += 1

                create_and_run_merge_script(generated_files, final_pdf)
            else:
                print("No files created.")
                
        except Exception as e:
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()