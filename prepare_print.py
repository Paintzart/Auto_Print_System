import sys
import os
import win32com.client
import time
import json
import pythoncom
import shutil

# הגדרת קידוד לקונסול כדי לתמוך בעברית ונתיבים
sys.stdout.reconfigure(encoding='utf-8')

# ========================================================
# 1. סקריפטים של JSX (לוגיקה פנימית של Adobe)
# ========================================================

# בדיקה האם הגרפיקה שחורה והאם הרקע בהדמיה לבן
JSX_DETECT_LOGIC = """
#target illustrator
function checkDoubleCondition() {
    try {
        var doc = app.activeDocument;
        var printLayerName = "%PRINT_LAYER%";
        var simSubName = "%SIM_SUB%"; // שם תת-השכבה בהדמיה (למשל S_Placement_Front)
        
        // 1. בדיקת שכבת ההדפסה - האם היא שחורה?
        var pLayer = doc.layers.getByName(printLayerName);
        if (pLayer.pageItems.length === 0) return "false";
        var isPrintBlack = quickScanColor(pLayer.pageItems, false); // מחפש שחור
        
        if (!isPrintBlack) return "false";

        // 2. בדיקת שכבת ההדמיה - האם הלוגו שם לבן?
        var simLayer = doc.layers.getByName("Simulation");
        var simSub = null;
        try { simSub = simLayer.layers.getByName(simSubName); } catch(e) {
            try { simSub = simLayer.groupItems.getByName(simSubName); } catch(e) {}
        }
        
        if (!simSub) return "false";
        
        var items = (simSub.typename === "Layer") ? simSub.pageItems : (simSub.pageItems || simSub.pathItems);
        var isSimWhite = quickScanColor(items, true); // מחפש לבן
        
        // רק אם שניהם אמת - מחזירים true
        return (isPrintBlack && isSimWhite) ? "true" : "false";
        
    } catch(e) { return "false"; }
}

function quickScanColor(items, findWhite) {
    for (var i = 0; i < Math.min(items.length, 15); i++) {
        var item = items[i];
        var c = null;
        if (item.typename === 'PathItem' && item.filled) c = item.fillColor;
        else if (item.typename === 'TextFrame') c = item.textRange.characterAttributes.fillColor;
        
        if (c) {
            if (findWhite && isWhite(c)) return true;
            if (!findWhite && isBlack(c)) return true;
        }
    }
    return false;
}

function isWhite(c) {
    if (c.typename === 'CMYKColor') return c.black===0 && c.cyan===0 && c.magenta===0 && c.yellow===0;
    if (c.typename === 'RGBColor') return c.red===255 && c.green===255 && c.blue===255;
    return false;
}

function isBlack(c) {
    if (c.typename === 'CMYKColor') return c.black > 80;
    if (c.typename === 'RGBColor') return c.red < 20 && c.green < 20 && c.blue < 20;
    return false;
}

checkDoubleCondition();
"""

# צביעת הגרפיקה בלבן נקי (CMYK 0,0,0,0)
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
            item.filled = true; item.fillColor = color; item.stroked = false;
        } else if (item.typename === 'CompoundPathItem') {
            for (var j = 0; j < item.pathItems.length; j++) {
                var p = item.pathItems[j];
                if (!p.clipping) { p.filled = true; p.fillColor = color; p.stroked = false; }
            }
        }
    }
}
runRecolor();
"""

# יצירת ספוט לבן (Spot1) בפוטושופ ושמירה חסינה
JSX_PS_TEMPLATE = r'''
#target photoshop
function main() {
    app.displayDialogs = DialogModes.NO;
    var filePath = FILE_PATH;
    var savePath = SAVE_PATH; 
    var contractPx = CONTRACT_PX; 

    try {
        var idOpn = charIDToTypeID("Opn ");
        var desc1 = new ActionDescriptor();
        desc1.putPath(charIDToTypeID("null"), new File(filePath));
        var desc2 = new ActionDescriptor();
        desc2.putUnitDouble(charIDToTypeID("Rslt"), charIDToTypeID("#Rsl"), 300.0);
        desc1.putObject(charIDToTypeID("As  "), charIDToTypeID("PDFG"), desc2);
        executeAction(idOpn, desc1);

        var doc = app.activeDocument;
        
        // בחירת אזורים לא שקופים
        var idset = charIDToTypeID("setd");
        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
        desc.putReference(charIDToTypeID("null"), ref);
        var ref2 = new ActionReference();
        ref2.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("Trsp"));
        desc.putReference(charIDToTypeID("T   "), ref2);
        executeAction(idset, desc);

        if (contractPx > 0) {
            try {
                var idCntc = charIDToTypeID("Cntc");
                var descC = new ActionDescriptor();
                descC.putUnitDouble(charIDToTypeID("By  "), charIDToTypeID("#Pxl"), contractPx);
                executeAction(idCntc, descC);
            } catch(e) {}
        }

        var spotChannel = doc.channels.add();
        spotChannel.name = "Spot1";
        spotChannel.kind = ChannelType.SPOTCOLOR;
        var whiteColor = new SolidColor();
        whiteColor.rgb.red = 255; whiteColor.rgb.green = 255; whiteColor.rgb.blue = 255;
        spotChannel.color = whiteColor;
        doc.selection.store(spotChannel);
        doc.selection.deselect();

        var saveFile = new File(savePath);
        var saved = false;

        // מנגנון שמירה חסין (3 ניסיונות)
        try {
            var opts1 = new PDFSaveOptions();
            opts1.presetFile = "[Press Quality]";
            doc.saveAs(saveFile, opts1, true); saved = true;
        } catch(e) {
            try {
                var opts2 = new PDFSaveOptions();
                opts2.presetFile = "[High Quality Print]";
                doc.saveAs(saveFile, opts2, true); saved = true;
            } catch(e2) {
                var opts3 = new PDFSaveOptions();
                opts3.spotColors = true;
                doc.saveAs(saveFile, opts3, true); saved = true;
            }
        }
        doc.close(SaveOptions.DONOTSAVECHANGES);
        if (saved) { writeStatus("SUCCESS"); }
    } catch(e) { writeStatus("ERROR: " + e.message); }
}
function writeStatus(msg) {
    var f = new File(STATUS_PATH);
    f.open("w"); f.write(msg); f.close();
}
main();
'''

# ========================================================
# 2. פונקציות ניהול (Python)
# ========================================================

def run_illustrator_split(source_file_path, order_number, output_folder):
    """מבצע פיצול שכבות באילוסטרייטור וצביעה ללבן במידת הצורך"""
    files_created = []
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("Illustrator.Application")
    app.UserInteractionLevel = -1 

    # מפה שמקשרת בין שכבת הדפסה לשכבת המיקום שלה בתוך ה-Simulation
    sim_mapping = {
        "Print_Front": "S_Placement_Front",
        "Print_Back": "S_Placement_Back",
        "Print_Left_Sleeve": "S_Placement_Left_Sleeve",
        "Print_Right_Sleeve": "S_Placement_Right_Sleeve"
    }

    jobs = [
        {"suffix": "PF", "layer": "Print_Front", "artboard": "Print_Front"},
        {"suffix": "PB", "layer": "Print_Back", "artboard": "Print_Back"},
        {"suffix": "PL", "layer": "Print_Left_Sleeve", "artboard": "Print_Sleeves"},
        {"suffix": "PR", "layer": "Print_Right_Sleeve", "artboard": "Print_Sleeves"},
    ]

    source_doc = app.Open(source_file_path)
    last_4 = str(order_number)[-4:]

    for job in jobs:
        try:
            lyr = source_doc.Layers(job["layer"])
            if lyr.PageItems.Count == 0: continue
            
            temp_file = os.path.join(output_folder, f"temp_{job['suffix']}.ai")
            shutil.copyfile(source_file_path, temp_file)
            work_doc = app.Open(temp_file)

            # --- עדכון שלב הבדיקה והצביעה (תנאי כפול) ---
            sim_sub = sim_mapping.get(job["layer"], "")
            detect_script = JSX_DETECT_LOGIC.replace("%PRINT_LAYER%", job["layer"]).replace("%SIM_SUB%", sim_sub)
            
            detect_res = app.DoJavaScript(detect_script)
            
            if str(detect_res).lower() == "true":
                print(f"✨ Conditions met: White Print detected for {job['layer']}. Recoloring...")
                app.DoJavaScript(JSX_RECOLOR_WHITE.replace("%TARGET_LAYER%", job["layer"]))
            else:
                print(f"ℹ️ {job['layer']} does not require white recoloring.")
            # --------------------------------------------

            # --- שלב א: ניקוי שכבות (Layers) ---
            for i in range(work_doc.Layers.Count, 0, -1):
                l = work_doc.Layers(i)
                if l.Name.strip().lower() != job["layer"].strip().lower():
                    try:
                        l.Locked = False; l.Visible = True; l.Delete()
                    except: pass
                else:
                    l.Visible = True; l.Locked = False

            # --- שלב ב: ניקוי דפי עבודה (Artboards) ---
            target_ab_index = -1
            for ab_i in range(work_doc.Artboards.Count):
                if job["artboard"].lower() in work_doc.Artboards(ab_i+1).Name.lower():
                    target_ab_index = ab_i
                    break
            
            if target_ab_index != -1:
                work_doc.Artboards.SetActiveArtboardIndex(target_ab_index)
                for ab_i in range(work_doc.Artboards.Count, 0, -1):
                    if (ab_i-1) != target_ab_index:
                        try: work_doc.Artboards(ab_i).Delete()
                        except: pass

            final_pdf = os.path.join(output_folder, f"{last_4}_{job['suffix']}.pdf")
            pdf_opts = win32com.client.Dispatch("Illustrator.PDFSaveOptions")
            pdf_opts.PDFPreset = "[Press Quality]"
            
            work_doc.SaveAs(final_pdf, pdf_opts)
            work_doc.Close(2)
            if os.path.exists(temp_file): os.remove(temp_file)
            files_created.append(final_pdf)
            
        except Exception as e:
            print(f"Error processing {job['layer']}: {e}")
            continue

    source_doc.Close(2)
    return files_created

def run_photoshop_processing(files_list, contract_px):
    pythoncom.CoInitialize()
    try:
        ps = win32com.client.Dispatch("Photoshop.Application")
    except Exception as e:
        print(f"Error connecting to Photoshop: {e}")
        return

    for file_path in files_list:
        # וידוא שהנתיב תקין עבור ווינדוס
        abs_path = os.path.abspath(file_path).replace("\\", "/")
        status_path = abs_path + ".status.txt"
        
        # עדכון ה-JSX עם נתיבים נקיים לחלוטין
        jsx_code = JSX_PS_TEMPLATE.replace("FILE_PATH", json.dumps(abs_path)) \
                                  .replace("SAVE_PATH", json.dumps(abs_path)) \
                                  .replace("STATUS_PATH", json.dumps(status_path)) \
                                  .replace("CONTRACT_PX", str(contract_px))
        
        jsx_file = os.path.join(os.path.dirname(file_path), "temp_ps_runner.jsx")
        with open(jsx_file, "w", encoding="utf-8") as f:
            f.write(jsx_code)
        
        print(f"Sending to Photoshop: {os.path.basename(file_path)}")
        ps.DoJavaScript(f'$.evalFile("{jsx_file.replace(os.sep, "/")}")')
        
        # המתנה לסיום (הגדלתי את זמן ההמתנה ל-30 שניות)
        success = False
        for _ in range(300): 
            if os.path.exists(status_path):
                try:
                    with open(status_path, "r") as f:
                        if "SUCCESS" in f.read(): success = True
                    os.remove(status_path)
                except: pass
                break
            time.sleep(0.1)
            
        if success:
            print(f"✅ Photoshop Saved: {os.path.basename(file_path)}")
        else:
            print(f"❌ Photoshop Failed to save: {os.path.basename(file_path)}")

        if os.path.exists(jsx_file): os.remove(jsx_file)

# ========================================================
# 3. נקודת כניסה ראשית
# ========================================================

if __name__ == "__main__":
    if len(sys.argv) < 4: sys.exit()
    
    input_file, order_id, thickness = sys.argv[1], sys.argv[2], sys.argv[3]
    contract_px = int(thickness.replace("px", ""))

    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
        output_base = config.get("print_folder_path", "C:/Temp/Print")

    order_dir = os.path.join(output_base, str(order_id)[-4:])
    if not os.path.exists(order_dir): os.makedirs(order_dir)

    print("Step 1: Illustrator...")
    generated = run_illustrator_split(input_file, order_id, order_dir)
    
    if generated:
        print("Step 2: Photoshop...")
        run_photoshop_processing(generated, contract_px)
    
    print("Done.")
    