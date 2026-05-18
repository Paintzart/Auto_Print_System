import os
import json
import win32com.client
import pythoncom
import shutil
import time

# ========================================================
# סקריפטים של JSX (ג'אווה סקריפט לאילוסטרייטור)
# ========================================================

# 1. הסקריפט של "הבלש" - רץ בתוך אילוסטרייטור ומחזיר "true"/"false"
JSX_DETECT_LOGIC = """
#target illustrator
function checkColors() {
    try {
        var doc = app.activeDocument;
        var printLayerName = "%PRINT_LAYER%";
        var simSubName = "%SIM_SUB%";
        
        // 1. בדיקת שכבת ההדפסה - האם היא שחורה?
        var pLayer = null;
        try { pLayer = doc.layers.getByName(printLayerName); } catch(e) { return "false"; }
        
        if (!pLayer.visible) pLayer.visible = true;
        if (pLayer.pageItems.length === 0) return "false"; // שכבה ריקה
        
        var isPrintBlack = quickScanColor(pLayer.pageItems, false); // false = מחפש שחור
        if (!isPrintBlack) return "false"; // אם ההדפסה לא שחורה, אין מה להמשיך
        
        // 2. בדיקת שכבת ההדמיה - האם היא לבנה?
        var simLayer = null;
        try { simLayer = doc.layers.getByName("Simulation"); } catch(e) { return "false"; }
        
        var simSub = null;
        // חיפוש תת-השכבה או הקבוצה
        try { simSub = simLayer.layers.getByName(simSubName); } catch(e) {}
        if (!simSub) {
            try { simSub = simLayer.groupItems.getByName(simSubName); } catch(e) {}
        }
        
        if (!simSub) return "false"; // לא נמצאה ההדמיה המתאימה
        
        // בדיקה אם ההדמיה לבנה
        var items = (simSub.typename === "Layer") ? simSub.pageItems : (simSub.pageItems || simSub.pathItems);
        var isSimWhite = quickScanColor(items, true); // true = מחפש לבן
        
        return isSimWhite ? "true" : "false";
        
    } catch(e) { return "false"; }
}

function quickScanColor(items, detectWhite) {
    // בודק רק את ה-20 פריטים הראשונים כדי להיות מהיר
    var limit = 20;
    var count = 0;
    
    for (var i = 0; i < items.length; i++) {
        if (count >= limit) break;
        var item = items[i];
        
        if (item.typename === 'GroupItem') {
            if (quickScanColor(item.pageItems, detectWhite)) return true;
        } else if (item.typename === 'PathItem') {
            count++;
            var c = null;
            if (item.filled) c = item.fillColor;
            else if (item.stroked) c = item.strokeColor;
            
            if (c) {
                if (detectWhite && isWhite(c)) return true;
                if (!detectWhite && isBlack(c)) return true;
            }
        } else if (item.typename === 'CompoundPathItem') {
             if (item.pathItems.length > 0) {
                 var p = item.pathItems[0];
                 var c = p.filled ? p.fillColor : p.strokeColor;
                 if (c) {
                    if (detectWhite && isWhite(c)) return true;
                    if (!detectWhite && isBlack(c)) return true;
                 }
             }
        }
    }
    return false;
}

function isWhite(c) {
    if (c.typename === 'CMYKColor') return c.cyan===0 && c.magenta===0 && c.yellow===0 && c.black===0;
    if (c.typename === 'RGBColor') return c.red===255 && c.green===255 && c.blue===255;
    if (c.typename === 'GrayColor') return c.gray===0;
    return false;
}

function isBlack(c) {
    if (c.typename === 'CMYKColor') return c.black > 90 || (c.cyan>40 && c.black>80);
    if (c.typename === 'RGBColor') return c.red===0 && c.green===0 && c.blue===0;
    if (c.typename === 'GrayColor') return c.gray > 90;
    return false;
}

checkColors();
"""

# 2. הסקריפט של "הצבע" - הופך ללבן במהירות
JSX_RECOLOR_WHITE = """
#target illustrator
function runRecolor() {
    try {
        var doc = app.activeDocument;
        var layerName = "%TARGET_LAYER%";
        var l = doc.layers.getByName(layerName);
        
        var white = new CMYKColor();
        white.cyan=0; white.magenta=0; white.yellow=0; white.black=0;
        
        recolorItems(l.pageItems, white);
    } catch(e) {}
}

function recolorItems(items, color) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename === 'GroupItem') {
            recolorItems(item.pageItems, color);
        } else if (item.typename === 'PathItem' && !item.clipping) {
            item.filled = true;
            item.fillColor = color;
            item.stroked = false; // מבטל קו מתאר כדי שיהיה נקי
        } else if (item.typename === 'CompoundPathItem') {
            for (var j = 0; j < item.pathItems.length; j++) {
                var p = item.pathItems[j];
                if (!p.clipping) {
                    p.filled = true;
                    p.fillColor = color;
                    p.stroked = false;
                }
            }
        } else if (item.typename === 'TextFrame') {
            try { item.textRange.characterAttributes.fillColor = color; } catch(e){}
        }
    }
}
runRecolor();
"""

# ========================================================
# פונקציות עזר לפייתון
# ========================================================
def run_jsx_script(app, script_content):
    """מריץ את ה-JSX ומחזיר את התוצאה (אם יש)"""
    try:
        return app.DoJavaScript(script_content)
    except Exception as e:
        # print(f"JSX Error: {e}") # נמנעים מלהדפיס שגיאות כדי לא להפריע ל-Streamlit
        return None

# ========================================================
# הלוגיקה הראשית של אילוסטרייטור (Generates PDF files)
# ========================================================
def run_illustrator_split(source_file_path, order_number):
    try:
        # הגדרות נתיבים
        try:
            # מנסה לטעון נתיב שמירה מ-config.json
            with open('config.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
                root_save_folder = config.get('save_folder_path', os.path.join(os.path.expanduser("~"), "Documents", "Auto_Print_Output"))
        except:
            # אם אין קובץ config, שומרים בתיקיית ברירת מחדל
            root_save_folder = os.path.join(os.path.expanduser("~"), "Documents", "Auto_Print_Output")

        last_4 = str(order_number)[-4:]
        order_folder = os.path.join(root_save_folder, last_4)
        print_files_folder = os.path.join(order_folder, "קבצי הדפסה")
        if not os.path.exists(print_files_folder): os.makedirs(print_files_folder)

        # חיבור לאילוסטרייטור
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Illustrator.Application")
        try: app.UserInteractionLevel = -1 
        except: pass

        # 1. פתיחת המקור (לקריאה בלבד)
        source_doc = app.Open(source_file_path)

        jobs = [
            {"suffix": "PF", "layer": "Print_Front", "artboard": "Print_Front", "sim_sublayer": "S_Placement_Front"},
            {"suffix": "PB", "layer": "Print_Back", "artboard": "Print_Back", "sim_sublayer": "S_Placement_Back"},
            {"suffix": "PL", "layer": "Print_Left_Sleeve", "artboard": "Print_Sleeves", "sim_sublayer": "S_Placement_Left_Sleeve"},
            {"suffix": "PR", "layer": "Print_Right_Sleeve", "artboard": "Print_Sleeves", "sim_sublayer": "S_Placement_Right_Sleeve"},
        ]

        files_to_process = []
        total_jobs = len(jobs)
        
        for i, job in enumerate(jobs):
            
            # --- שלב א: בדיקה מקדימה (האם השכבה קיימת?) ---
            layer_exists = False
            try:
                source_doc.Layers(job["layer"])
                layer_exists = True
            except: pass
            
            if not layer_exists:
                yield (i + 1) / total_jobs, f"מדלג: {job['layer']} (לא קיים)"
                continue

            # --- שלב ב: הבלש (רץ בתוך JSX - מהיר!) ---
            yield (i + 1) / total_jobs, f"מנתח צבעים: {job['layer']}..."
            
            detect_script = JSX_DETECT_LOGIC.replace("%PRINT_LAYER%", job["layer"]).replace("%SIM_SUB%", job["sim_sublayer"])
            result_str = run_jsx_script(app, detect_script)
            should_turn_white = (str(result_str).lower() == "true")
            
            # בדיקה אם השכבה ריקה (אחרי בדיקת ה-JSX)
            try:
                if source_doc.Layers(job["layer"]).PageItems.Count == 0:
                    yield (i + 1) / total_jobs, f"מדלג: {job['layer']} (ריקה)"
                    continue
            except: pass

            # --- שלב ג: ביצוע (העתקה, צביעה, שמירה) ---
            yield (i + 1) / total_jobs, f"מעבד ושומר: {job['layer']}..."
            
            temp_work_file = os.path.join(print_files_folder, f"temp_{job['suffix']}.ai")
            
            # 1. העתקה מהירה בדיסק
            shutil.copyfile(source_file_path, temp_work_file)
            
            # 2. פתיחת העותק
            work_doc = app.Open(temp_work_file)
            
            # 3. צביעה ללבן (אם הבלש אמר) - רץ ב-JSX מהיר!
            if should_turn_white:
                recolor_script = JSX_RECOLOR_WHITE.replace("%TARGET_LAYER%", job["layer"])
                run_jsx_script(app, recolor_script)
            
            # 4. מחיקת שכבות מיותרות
            idx = work_doc.Layers.Count
            while idx >= 1:
                try:
                    lyr = work_doc.Layers(idx)
                    if lyr.Name.strip().lower() != job["layer"].strip().lower():
                        lyr.Locked = False; lyr.Visible = True
                        lyr.Delete()
                    else:
                        lyr.Locked = False; lyr.Visible = True # מוודאים ששכבת ההדפסה דלוקה
                except: pass
                idx -= 1
            
            # 5. מחיקת Artboards מיותרים
            target_ab_index = -1
            for ab_i in range(work_doc.Artboards.Count):
                if job["artboard"].lower() in work_doc.Artboards(ab_i+1).Name.lower():
                    target_ab_index = ab_i; break
            
            if target_ab_index != -1:
                work_doc.Artboards.SetActiveArtboardIndex(target_ab_index)
                idx = work_doc.Artboards.Count - 1
                while idx >= 0:
                    if idx != target_ab_index:
                        try: work_doc.Artboards(idx+1).Delete()
                        except: pass
                    idx -= 1
            
            # 6. שמירה כ-PDF
            filename = f"{last_4}_{job['suffix']}.pdf"
            full_save_path = os.path.join(print_files_folder, filename)
            
            pdf_options = win32com.client.Dispatch("Illustrator.PDFSaveOptions")
            pdf_options.PDFPreset = "[Press Quality]"
            pdf_options.ArtboardRange = "1"
            
            work_doc.SaveAs(full_save_path, pdf_options)
            work_doc.Close(2) # סגירת הזמני
            try: os.remove(temp_work_file) # מחיקת הזמני
            except: pass
            
            # שמירת הנתיב המלא של הקובץ שנוצר! (קריטי לפוטושופ)
            files_to_process.append(full_save_path)

        # סיום: סגירת המקור
        source_doc.Close(2)
        try: app.UserInteractionLevel = 1 
        except: pass
        
        # מחזירים את התיקייה ואת רשימת הנתיבים המלאה של הקבצים שנוצרו
        yield "DONE", (print_files_folder, files_to_process)

    except Exception as e:
        try: app.UserInteractionLevel = 1
        except: pass
        try: app.ActiveDocument.Close(2)
        except: pass
        raise e