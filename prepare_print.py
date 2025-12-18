print("--- הסטארט-אפ של הסקריפט עובד ---") # שורה לבדיקה
import sys
import os
print("--- ספריות בסיסיות נטענו ---") # שורה לבדיקה
import win32com.client
print("--- win32com נטען בהצלחה ---") # שורה לבדיקה
import time
import json
import pythoncom
import shutil
print("--- כל הספריות נטענו ---") #


# הגדרת קידוד לקונסול

sys.stdout.reconfigure(encoding='utf-8')



# ========================================================

# 1. סקריפטים לאילוסטרייטור (זיהוי צבע וצביעה)

# ========================================================

JSX_DETECT_LOGIC = """

#target illustrator

function checkColors() {

    try {

        var doc = app.activeDocument;

        var printLayerName = "%PRINT_LAYER%";

        var pLayer = null;

        try { pLayer = doc.layers.getByName(printLayerName); } catch(e) { return "false"; }

        if (!pLayer.visible) pLayer.visible = true;

        if (pLayer.pageItems.length === 0) return "empty"; 

        var isBlack = quickScanBlack(pLayer.pageItems);

        return isBlack ? "true" : "false";

    } catch(e) { return "false"; }

}

function quickScanBlack(items) {

    var limit = 20; 

    var count = 0;

    for (var i = 0; i < items.length; i++) {

        if (count >= limit) break;

        var item = items[i];

        if (item.typename === 'GroupItem') {

            if (quickScanBlack(item.pageItems)) return true;

        } else if (item.typename === 'PathItem') {

            count++;

            var c = item.filled ? item.fillColor : (item.stroked ? item.strokeColor : null);

            if (c && isColorBlack(c)) return true;

        } else if (item.typename === 'CompoundPathItem' && item.pathItems.length > 0) {

             var p = item.pathItems[0];

             var c = p.filled ? p.fillColor : (p.stroked ? p.strokeColor : null);

             if (c && isColorBlack(c)) return true;

        }

    }

    return false;

}

function isColorBlack(c) {

    if (c.typename === 'CMYKColor') return c.black > 90;

    if (c.typename === 'RGBColor') return c.red < 10 && c.green < 10 && c.blue < 10;

    if (c.typename === 'GrayColor') return c.gray > 90;

    return false;

}

checkColors();

"""



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

        } else if (item.typename === 'TextFrame') {

            try { item.textRange.characterAttributes.fillColor = color; } catch(e){}

        }

    }

}

runRecolor();

"""



# ========================================================

# 2. סקריפט פוטושופ (מתוקן: פתרון לבעיית השמירה)

# ========================================================

JSX_PS_TEMPLATE = r'''

#target photoshop

function main() {

    app.displayDialogs = DialogModes.NO;

    app.preferences.rulerUnits = Units.PIXELS;

    

    var filePath = FILE_PATH;

    var savePath = SAVE_PATH; 

    var contractPx = CONTRACT_PX; 



    try {

        // 1. פתיחה

        var idOpn = charIDToTypeID("Opn ");

        var desc1 = new ActionDescriptor();

        desc1.putPath(charIDToTypeID("null"), new File(filePath));

        var desc2 = new ActionDescriptor();

        desc2.putEnumerated(charIDToTypeID("CrpT"), charIDToTypeID("CrpT"), charIDToTypeID("BndB"));

        desc2.putUnitDouble(charIDToTypeID("Rslt"), charIDToTypeID("#Rsl"), 300.000000);

        desc2.putEnumerated(charIDToTypeID("Md  "), charIDToTypeID("ClrS"), charIDToTypeID("RGBC"));

        desc1.putObject(charIDToTypeID("As  "), charIDToTypeID("PDFG"), desc2);

        executeAction(idOpn, desc1);



        var doc = app.activeDocument;



        // 2. בחירה (Load Selection)

        var idset = charIDToTypeID("setd");

        var desc = new ActionDescriptor();

        var ref = new ActionReference();

        ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));

        desc.putReference(charIDToTypeID("null"), ref);

        var ref2 = new ActionReference();

        ref2.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("Trsp"));

        desc.putReference(charIDToTypeID("T   "), ref2);

        executeAction(idset, desc);



        // 3. כיווץ (Contract)

        if (contractPx > 0) {

            try {

                var idCntc = charIDToTypeID("Cntc");

                var descC = new ActionDescriptor();

                descC.putUnitDouble(charIDToTypeID("By  "), charIDToTypeID("#Pxl"), contractPx);

                descC.putBoolean(charIDToTypeID("FrCn"), true);

                executeAction(idCntc, descC);

            } catch(e) {}

        }



        // 4. יצירת ספוט לבן

        try {

            var spotChannel = doc.channels.add();

            spotChannel.name = "Spot1";

            spotChannel.kind = ChannelType.SPOTCOLOR;

            spotChannel.opacity = 0; 

            var whiteColor = new SolidColor();

            whiteColor.rgb.red = 255; whiteColor.rgb.green = 255; whiteColor.rgb.blue = 255;

            spotChannel.color = whiteColor;

            doc.selection.store(spotChannel);

        } catch(e) {

             try { doc.selection.store(doc.channels.getByName("Spot1")); } catch(_){}

        }



        doc.selection.deselect();



        // 5. וידוא תצוגה (מדליק הכל)

        try {

            var visibleChannels = new Array();

            var componentChannels = doc.componentChannels;

            for (var i = 0; i < componentChannels.length; i++) {

                visibleChannels.push(componentChannels[i]);

            }

            try { visibleChannels.push(doc.channels.getByName("Spot1")); } catch(e){}

            doc.activeChannels = visibleChannels; 

            

            // וידוא נוסף שכל הערוצים גלויים

            for (var k=0; k<doc.channels.length; k++) {

                doc.channels[k].visible = true;

            }

        } catch(e) {}



        // 6. שמירה לקובץ חדש (עם מנגנון גיבוי משופר)

        var saveFile = new File(savePath);

        var saved = false;



        // נסיון 1: Press Quality

        try {

            var opts1 = new PDFSaveOptions();

            opts1.spotColors = true;

            opts1.alphaChannels = true;

            opts1.layers = true;

            opts1.presetFile = "[Press Quality]"; 

            doc.saveAs(saveFile, opts1, true, Extension.LOWERCASE);

            saved = true;

        } catch(e) {

            // נסיון 2: High Quality Print

            try {

                var opts2 = new PDFSaveOptions();

                opts2.spotColors = true;

                opts2.alphaChannels = true;

                opts2.layers = true;

                opts2.presetFile = "[High Quality Print]"; 

                doc.saveAs(saveFile, opts2, true, Extension.LOWERCASE);

                saved = true;

            } catch(e2) {

                // נסיון 3: ללא שום Preset (הגדרות ידניות בסיסיות) - זה התיקון לשגיאת Required value missing

                try {

                    var opts3 = new PDFSaveOptions();

                    opts3.spotColors = true;

                    opts3.alphaChannels = true;

                    opts3.layers = true;

                    opts3.encoding = PDFEncoding.JPEG;

                    opts3.downSample = PDFResample.NONE; 

                    // לא מגדירים presetFile בכלל!

                    

                    doc.saveAs(saveFile, opts3, true, Extension.LOWERCASE);

                    saved = true;

                } catch(e3) {

                    writeStatus("ERROR_SAVE: " + e3.message);

                }

            }

        }



        doc.close(SaveOptions.DONOTSAVECHANGES);

        

        if (saved) {

            writeStatus("SUCCESS");

        }



    } catch(e) {

        var f = new File(filePath + ".status.txt");

        f.open("w"); f.write("ERROR_MAIN: " + e.message); f.close();

    } finally {

        app.displayDialogs = DialogModes.ALL;

    }

}



function writeStatus(msg) {

    var f = new File(STATUS_PATH);

    f.open("w"); f.write(msg); f.close();

}



main();

'''



# ========================================================

# פונקציות עזר

# ========================================================

def run_jsx_script(app, script_content):

    try:

        return app.DoJavaScript(script_content)

    except:

        return None



# ========================================================

# שלב 1: אילוסטרייטור

# ========================================================

def run_illustrator_split(source_file_path, order_number, output_folder):

    files_created = []

    

    try:

        pythoncom.CoInitialize()

        app = win32com.client.Dispatch("Illustrator.Application")

        app.UserInteractionLevel = -1 

    except:

        print("Illustrator Error: Not open or installed.")

        return []



    jobs = [

        {"suffix": "PF", "layer": "Print_Front", "artboard": "Print_Front"},

        {"suffix": "PB", "layer": "Print_Back", "artboard": "Print_Back"},

        {"suffix": "PL", "layer": "Print_Left_Sleeve", "artboard": "Print_Sleeves"},

        {"suffix": "PR", "layer": "Print_Right_Sleeve", "artboard": "Print_Sleeves"},

    ]



    last_4 = str(order_number)[-4:]



    try:

        source_doc = app.Open(source_file_path)

    except:

        print(f"Failed to open source: {source_file_path}")

        return []



    for job in jobs:

        # בדיקה אם השכבה קיימת

        try:

            source_doc.Layers(job["layer"])

        except:

            print(f"Skipping {job['layer']} (Not found)")

            continue



        # בדיקה אם השכבה ריקה

        item_count = 0

        try:

            item_count = source_doc.Layers(job["layer"]).PageItems.Count

        except: pass

        

        if item_count == 0:

            print(f"Skipping {job['layer']} (Layer is empty)")

            continue



        print(f"Processing layer: {job['layer']}...")



        temp_work_file = os.path.join(output_folder, f"temp_work_{job['suffix']}.ai")

        shutil.copyfile(source_file_path, temp_work_file)

        

        try:

            work_doc = app.Open(temp_work_file)

        except:

            continue



        # צביעה ללבן אם צריך

        detect_script = JSX_DETECT_LOGIC.replace("%PRINT_LAYER%", job["layer"])

        res = run_jsx_script(app, detect_script)

        if str(res).lower() == "true":

            print(f"  -> Turning items to WHITE for {job['layer']}")

            recolor_script = JSX_RECOLOR_WHITE.replace("%TARGET_LAYER%", job["layer"])

            run_jsx_script(app, recolor_script)



        idx = work_doc.Layers.Count

        while idx >= 1:

            try:

                lyr = work_doc.Layers(idx)

                if lyr.Name.strip().lower() != job["layer"].strip().lower():

                    lyr.Locked = False; lyr.Visible = True

                    lyr.Delete()

                else:

                    lyr.Visible = True; lyr.Locked = False

            except: pass

            idx -= 1



        target_ab_index = -1

        for i in range(work_doc.Artboards.Count):

            if job["artboard"].lower() in work_doc.Artboards(i+1).Name.lower():

                target_ab_index = i; break

        

        if target_ab_index != -1:

            work_doc.Artboards.SetActiveArtboardIndex(target_ab_index)

            initial_count = work_doc.Artboards.Count

            for i in range(initial_count, 0, -1):

                if (i-1) != target_ab_index:

                    try: work_doc.Artboards(i).Delete()

                    except: pass

        

        filename = f"{last_4}_{job['suffix']}.pdf"

        full_save_path = os.path.join(output_folder, filename)

        

        if os.path.exists(full_save_path):

            try: os.remove(full_save_path)

            except: pass



        pdf_options = win32com.client.Dispatch("Illustrator.PDFSaveOptions")

        pdf_options.PDFPreset = "[Press Quality]"

        pdf_options.ArtboardRange = "1"

        

        work_doc.SaveAs(full_save_path, pdf_options)

        print(f"Illustrator Generated: {filename}")

        files_created.append(full_save_path)

        

        work_doc.Close(2) 

        try: os.remove(temp_work_file) 

        except: pass



    source_doc.Close(2) 

    try: app.UserInteractionLevel = 1 

    except: pass

    

    return files_created



# ========================================================

# שלב 2: פוטושופ

# ========================================================

def run_photoshop_processing(files_list, contract_px):

    if not files_list: return



    print("Starting Photoshop processing...")

    pythoncom.CoInitialize()

    try:

        ps = win32com.client.Dispatch("Photoshop.Application")

    except:

        return



    for file_path in files_list:

        if not os.path.exists(file_path): continue

        

        print(f"Processing in PS: {os.path.basename(file_path)}")

        

        temp_final_path = file_path.replace(".pdf", "_ready.pdf")

        if os.path.exists(temp_final_path): os.remove(temp_final_path)



        status_path = file_path + ".status.txt"

        # שימוש בנתיב עם סלאשים רגילים ל-JSX

        file_path_js = file_path.replace("\\", "/")

        temp_final_path_js = temp_final_path.replace("\\", "/")

        status_path_js = status_path.replace("\\", "/")

        

        jsx_path = os.path.join(os.path.dirname(file_path), "_temp_ps_script.jsx")

        

        if os.path.exists(status_path): os.remove(status_path)

        # וודאי שאין רווחים לפני ה-נקודה בכל שורה כזו
# החלפת הטקסט בצורה בטוחה ללא שבירת שורות
        jsx_code = JSX_PS_TEMPLATE.replace("FILE_PATH", json.dumps(file_path_js))
        jsx_code = jsx_code.replace("SAVE_PATH", json.dumps(temp_final_path_js))
        jsx_code = jsx_code.replace("STATUS_PATH", json.dumps(status_path_js))
        jsx_code = jsx_code.replace("CONTRACT_PX", str(contract_px))

        with open(jsx_path, "w", encoding="utf-8") as f: f.write(jsx_code)

        

        try: ps.DoJavaScript(f'$.evalFile("{jsx_path.replace(os.sep, "/")}")')

        except: pass



        success = False

        for _ in range(300): # 30 שניות המתנה

            if os.path.exists(status_path):

                time.sleep(0.5) 

                try: 

                    with open(status_path, "r") as f: status = f.read()

                    os.remove(status_path)

                    if "SUCCESS" in status: success = True

                    else: print(f"PS Script Reported: {status}")

                except: pass

                break

            time.sleep(0.1)

        

        try: os.remove(jsx_path)

        except: pass

        

        if success and os.path.exists(temp_final_path):

            try:

                os.remove(file_path)

                os.rename(temp_final_path, file_path)

                print(f"✅ Photoshop Finished & Saved: {os.path.basename(file_path)}")

            except Exception as e:

                print(f"⚠️ Error swapping files: {e}")

        else:

            print(f"❌ Photoshop Timeout/Error: {os.path.basename(file_path)}")



# ========================================================

# ראשי

# ========================================================

def main():

    if len(sys.argv) < 4:

        print("Usage: python prepare_print.py <file_path> <order_id> <thickness>")

        return



    input_file = sys.argv[1]

    order_id = sys.argv[2]

    thickness_str = sys.argv[3]



    try:

        contract_px = int(thickness_str.replace("px", "").strip())

    except:

        contract_px = 2



    # --- קריאת הגדרות מ-config.json ---

    config_path = os.path.join(os.path.dirname(__file__), 'config.json')

    if os.path.exists(config_path):

        with open(config_path, 'r', encoding='utf-8') as f:

            config = json.load(f)

            # שימוש בנתיב הייעודי לקבצי הדפסה (הכפתור הוורוד)

            base_output_path = config.get("print_folder_path", "C:/Temp/Print")

    else:

        base_output_path = r"C:\Temp\Print"



    last_4 = str(order_id)[-4:]

    # יצירת נתיב ישיר לתיקיית ההזמנה

    order_folder = os.path.join(base_output_path, last_4)

    

    if not os.path.exists(order_folder):

        os.makedirs(order_folder, exist_ok=True)



    # שלב 1: אילוסטרייטור

    print(f"--- Step 1: Illustrator Split ({order_id}) ---")

    generated_files = run_illustrator_split(input_file, order_id, order_folder)



    time.sleep(1)



    # שלב 2: פוטושופ

    if generated_files:

        print(f"--- Step 2: Photoshop Spot ({contract_px}px) ---")

        run_photoshop_processing(generated_files, contract_px)

    else:

        print("No files generated from Illustrator.")



    print("Done.")

if __name__ == "__main__":
    main()