import os
import json
import time
import pythoncom
import win32com.client

# =========================================================
# סקריפט JSX (הגרסה הסופית, עובדת)
# =========================================================
JSX_TEMPLATE = r'''
#target photoshop

function writeStatus(msg) {
    var f = new File(STATUS_PATH);
    f.open("w");
    f.write(msg);
    f.close();
}

function main() {
    
    app.preferences.rulerUnits = Units.PIXELS;
    app.preferences.typeUnits = TypeUnits.PIXELS;

    var filePath = FILE_PATH;
    var savePath = SAVE_PATH; 
    var contractPx = CONTRACT_PX; 

    try {
        // ---------------------------------------------------
        // 1. פתיחה (Action Manager)
        // ---------------------------------------------------
        var idOpn = charIDToTypeID("Opn ");
        var desc1 = new ActionDescriptor();
        desc1.putPath(charIDToTypeID("null"), new File(filePath));
        var desc2 = new ActionDescriptor();
        desc2.putEnumerated(charIDToTypeID("CrpT"), charIDToTypeID("CrpT"), charIDToTypeID("BndB"));
        desc2.putUnitDouble(charIDToTypeID("Rslt"), charIDToTypeID("#Rsl"), 300.000000);
        desc2.putEnumerated(charIDToTypeID("Md  "), charIDToTypeID("ClrS"), charIDToTypeID("RGBC"));
        desc2.putBoolean(charIDToTypeID("Trns"), true);
        desc2.putBoolean(charIDToTypeID("AntA"), true);
        desc1.putObject(charIDToTypeID("As  "), charIDToTypeID("PDFG"), desc2);
        executeAction(idOpn, desc1);

        var doc = app.activeDocument;
        if (doc.width.value < 1) { writeStatus("ERROR: File empty"); return; }

        // ---------------------------------------------------
        // 2. בחירה (Selection)
        // ---------------------------------------------------
        var idset = charIDToTypeID("setd");
        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
        desc.putReference(charIDToTypeID("null"), ref);
        var ref2 = new ActionReference();
        ref2.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("Trsp"));
        desc.putReference(charIDToTypeID("T   "), ref2);
        executeAction(idset, desc);

        // ---------------------------------------------------
        // 3. כיווץ (Contract)
        // ---------------------------------------------------
        try {
            var idCntc = charIDToTypeID("Cntc");
            var descC = new ActionDescriptor();
            descC.putUnitDouble(charIDToTypeID("By  "), charIDToTypeID("#Pxl"), contractPx);
            descC.putBoolean(charIDToTypeID("FrCn"), true);
            executeAction(idCntc, descC);
        } catch(e) {}

        // ---------------------------------------------------
        // 4. יצירת ספוט לבן (DOM)
        // ---------------------------------------------------
        var spotChannel;
        try {
            spotChannel = doc.channels.add();
            spotChannel.name = "Spot1";
            
            spotChannel.kind = ChannelType.SPOTCOLOR;
            spotChannel.opacity = 0; 

            var whiteColor = new SolidColor();
            whiteColor.rgb.red = 255;
            whiteColor.rgb.green = 255;
            whiteColor.rgb.blue = 255;
            spotChannel.color = whiteColor;

            doc.selection.store(spotChannel);

        } catch(e) {
             try { doc.selection.store(doc.channels.getByName("Spot1")); } catch(_){}
        }

        doc.selection.deselect();

        // ---------------------------------------------------
        // 5. וידוא תצוגה (הדלקת עיניים)
        // ---------------------------------------------------
        try {
            var visibleChannels = new Array();
            var componentChannels = doc.componentChannels;
            for (var i = 0; i < componentChannels.length; i++) {
                visibleChannels.push(componentChannels[i]);
            }
            if (spotChannel) visibleChannels.push(spotChannel);
            doc.activeChannels = visibleChannels;
        } catch(e) {}

        // ---------------------------------------------------
        // 6. שמירה ודריסה (Press Quality)
        // ---------------------------------------------------
        var saveFile = new File(savePath);
        
        try {
            // נסיון ראשון: שמירה עם הגדרת "Press Quality"
            var optsPress = new PDFSaveOptions();
            optsPress.alphaChannels = true;
            optsPress.spotColors = true;
            optsPress.layers = true;
            optsPress.presetFile = "[Press Quality]"; // השם הסטנדרטי לאיכות דפוס
            
            doc.saveAs(saveFile, optsPress, true, Extension.LOWERCASE);
            writeStatus("SUCCESS_SAVED_PRESS");

        } catch(e) {
            // נסיון שני (גיבוי): אם השם של הפריסט לא עבד, שומרים רגיל
            try {
                var optsDefault = new PDFSaveOptions();
                optsDefault.alphaChannels = true;
                optsDefault.spotColors = true;
                optsDefault.layers = true;
                doc.saveAs(saveFile, optsDefault, true, Extension.LOWERCASE);
                writeStatus("SUCCESS_SAVED_DEFAULT");
            } catch(e2) {
                writeStatus("ERROR_SAVE: " + e2.message);
                return;
            }
        }

        doc.close(SaveOptions.DONOTSAVECHANGES);

    } catch(e) {
        try { if(doc) doc.close(SaveOptions.DONOTSAVECHANGES); } catch(_) {}
        writeStatus("CRITICAL: " + e.message);
    }
}

main();
'''

# =========================================================
# Python Runner - הפונקציה המקושרת (run_photoshop_action)
# =========================================================
def run_photoshop_action(files_list, contract_px=2):
    """
    מריץ את תהליך הספוט הלבן על רשימת קבצים ספציפית.
    מחזיר גנרטור להתקדמות.
    """
    if not files_list:
        yield 1.0, "לא נמצאו קבצים לעיבוד."
        return

    pythoncom.CoInitialize()
    try:
        ps = win32com.client.Dispatch("Photoshop.Application")
    except Exception as e:
        yield 0.0, f"שגיאת חיבור לפוטושופ: {e}"
        return

    total_files = len(files_list)
    
    for i, full_path in enumerate(files_list):
        filename = os.path.basename(full_path)
        
        # בדיקה האם הקובץ קיים לפני שמתחילים
        if not os.path.exists(full_path):
            # מדלגים, אבל מעדכנים התקדמות
            yield (i + 1) / total_files, f"מדלג על {filename}: קובץ לא נמצא."
            continue

        # הגדרת נתיבים זמניים וסטטוס
        status_path = full_path + ".status.txt"
        jsx_path = os.path.join(os.path.dirname(full_path), "_temp_run.jsx")

        # הכנת ה-JSX
        jsx = JSX_TEMPLATE \
            .replace("FILE_PATH", json.dumps(full_path)) \
            .replace("SAVE_PATH", json.dumps(full_path)) \
            .replace("STATUS_PATH", json.dumps(status_path)) \
            .replace("CONTRACT_PX", str(contract_px))

        try:
            with open(jsx_path, "w", encoding="utf-8") as f:
                f.write(jsx)
        except Exception as e:
            yield (i + 1) / total_files, f"שגיאה בכתיבת JSX עבור {filename}: {e}"
            continue

        yield (i + 1) / total_files, f"מעבד: {filename}..."

        try:
            ps.DoJavaScript(f'$.evalFile("{jsx_path.replace("\\", "/")}")')
        except Exception as e:
            yield (i + 1) / total_files, f"שגיאת הרצה PS עבור {filename}: {e}"
            pass 

        # המתנה לקובץ סטטוס
        result = "TIMEOUT"
        for _ in range(100):
            if os.path.exists(status_path):
                try:
                    with open(status_path, "r", encoding="utf-8") as s:
                        result = s.read().strip()
                    time.sleep(0.1)
                    os.remove(status_path)
                    break
                except:
                    time.sleep(0.1)
            time.sleep(0.2)

        # עדכון סופי של הקובץ הנוכחי
        yield (i + 1) / total_files, f"סיים {filename}: [{result}]"
        
        try: os.remove(jsx_path)
        except: pass

    yield "DONE", "תהליך פוטושופ הושלם."