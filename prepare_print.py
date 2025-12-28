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
# יצירת קובץ סיכום הדמיות (הגרסה הסופית: Manual Grouping לשמירה על שכבות)
JSX_CREATE_SIM_SUMMARY = r'''
#target illustrator
function createSimSummary() {
    var sourceDoc = app.activeDocument;
    var savePath = SAVE_PATH;
    
    // איפוס: פתיחת הכל
    app.executeMenuCommand('unlockAll');
    app.executeMenuCommand('showAll');
    sourceDoc.selection = null;

    // --- שלב 1: איסוף תיקיות Simulation ---
    var simLayersData = []; 
    
    for (var i = 0; i < sourceDoc.layers.length; i++) {
        var mainLayer = sourceDoc.layers[i];
        var prodIdx = parseInt(mainLayer.name);
        
        if (!isNaN(prodIdx)) {
            try {
                // וידוא שהשכבה פתוחה
                mainLayer.visible = true;
                mainLayer.locked = false;
                
                var simLayer = mainLayer.layers.getByName("Simulation");
                simLayer.visible = true;
                simLayer.locked = false;
                
                if (simLayer.pageItems.length > 0 || simLayer.layers.length > 0) {
                    simLayersData.push({ id: prodIdx, layer: simLayer });
                }
            } catch(e) {}
        }
    }

    simLayersData.sort(function(a, b) { return a.id - b.id; });

    if (simLayersData.length === 0) return;

    // --- מקרה 1: הדמיה אחת ---
    if (simLayersData.length === 1) {
        var tempG = manualGroup(simLayersData[0].layer); 
        var targetAbIdx = 0;
        
        for (var k=0; k<sourceDoc.artboards.length; k++) {
            if (intersect(sourceDoc.artboards[k].artboardRect, tempG.visibleBounds)) {
                targetAbIdx = k;
                break;
            }
        }
        
        var opts = new PDFSaveOptions();
        opts.presetFile = "[Press Quality]";
        opts.artboardRange = (targetAbIdx + 1).toString();
        sourceDoc.saveAs(new File(savePath), opts);
        
        sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
        return;
    }

    // --- מקרה 2: ריבוי הדמיות ---
    var A4_W = 595.28;
    var A4_H = 841.89;
    var isLandscape = (simLayersData.length > 2); 
    
    var targetDoc = app.documents.add(DocumentColorSpace.CMYK, 
                                      isLandscape ? A4_H : A4_W, 
                                      isLandscape ? A4_W : A4_H);
    
    var cellW, cellH;
    if (!isLandscape) { cellW = A4_W; cellH = A4_H / 2; } 
    else { cellW = A4_H / 2; cellH = A4_W / 2; }

    var itemsPerSheet = isLandscape ? 4 : 2;
    var sheetCount = 1;

    for (var i = 0; i < simLayersData.length; i++) {
        var currentData = simLayersData[i];
        
        if (i > 0 && i % itemsPerSheet === 0) {
            targetDoc.artboards.add(targetDoc.artboards[0].artboardRect);
            sheetCount++;
        }

        var tempGroup = manualGroup(currentData.layer);
        var targetGroup = tempGroup.duplicate(targetDoc.activeLayer, ElementPlacement.PLACEATEND);
        
        if (targetGroup) {
            app.activeDocument = targetDoc; 
            
            var indexOnPage = i % itemsPerSheet;
            var col = 0, row = 0;
            
            if (!isLandscape) { col = 0; row = indexOnPage; } 
            else { col = indexOnPage % 2; row = Math.floor(indexOnPage / 2); }

            var margin = 20;
            var targetW = cellW - (margin * 2);
            var targetH = cellH - (margin * 2);
            
            var scaleX = (targetW / targetGroup.width) * 100;
            var scaleY = (targetH / targetGroup.height) * 100;
            var scale = Math.min(scaleX, scaleY);
            if (scale > 100) scale = 100; 
            
            targetGroup.resize(scale, scale, true, true, true, true, scale);

            var abRect = targetDoc.artboards[sheetCount-1].artboardRect;
            var abLeft = abRect[0];
            var abTop = abRect[1];
            var cellCenterX = abLeft + (col * cellW) + (cellW / 2);
            var cellCenterY = abTop - (row * cellH) - (cellH / 2);

            targetGroup.position = [
                cellCenterX - (targetGroup.width / 2),
                cellCenterY + (targetGroup.height / 2)
            ];
        }
    }

    var saveOpts = new PDFSaveOptions();
    saveOpts.presetFile = "[Press Quality]";
    targetDoc.saveAs(new File(savePath), saveOpts);
    targetDoc.close(SaveOptions.DONOTSAVECHANGES);
    
    sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
}

function manualGroup(layerObj) {
    var newGroup = layerObj.groupItems.add();
    newGroup.name = "TempContainer";
    
    var items = layerObj.pageItems;
    for (var i = items.length - 1; i >= 0; i--) {
        var item = items[i];
        if (item !== newGroup) {
            try { item.move(newGroup, ElementPlacement.PLACEATBEGINNING); } catch(e){}
        }
    }
    flattenSubLayers(layerObj, newGroup);
    return newGroup;
}

function flattenSubLayers(parentLayer, targetGroup) {
    for (var i = parentLayer.layers.length - 1; i >= 0; i--) {
        var subLayer = parentLayer.layers[i];
        for (var j = subLayer.pageItems.length - 1; j >= 0; j--) {
            try { subLayer.pageItems[j].move(targetGroup, ElementPlacement.PLACEATBEGINNING); } catch(e){}
        }
        flattenSubLayers(subLayer, targetGroup);
    }
}

function intersect(r1, r2) {
    return !(r2[0] > r1[2] || r2[2] < r1[0] || r2[1] < r1[3] || r2[3] > r1[1]);
}

createSimSummary();
'''

# 1. סקריפטים של JSX (לוגיקה פנימית של Adobe)
# ========================================================

def get_layer_suffix(layer_name):
    """מזהה את הסיומת רק ל-4 הסוגים המותרים"""
    name = layer_name.lower()
    
    # זיהוי מדויק לפי סדר חשיבות
    if "front" in name: return "PF"
    if "back" in name: return "PB"
    
    # בדיקה ספציפית לשרוולים (ימין/שמאל)
    if "sleeve" in name:
        if "right" in name: return "PR"
        if "left" in name: return "PL"
        
    return None # אם זה לא אחד מאלה, נתעלם

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
    """מבצע פיצול חכם לפי מוצרים (1, 2...) ולפי שכבות הדפסה"""
    files_created = []
    
    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch("Illustrator.Application")
    except:
        print("Error: Could not connect to Illustrator.")
        return []
        
    app.UserInteractionLevel = -1 

    print(f"Opening Master File: {os.path.basename(source_file_path)}")
    try:
        master_doc = app.Open(source_file_path)
    except:
        print("Error opening file.")
        return []

    # --- שלב 1: מיפוי וסינון המשימות (Jobs) ---
    job_list = [] 
    
    for main_layer in master_doc.Layers:
        prod_idx = main_layer.Name.strip()
        
        if not prod_idx.isdigit():
            continue
            
        for sub_layer in main_layer.Layers:
            if not sub_layer.Name.lower().startswith("print"):
                continue

            if sub_layer.PageItems.Count == 0:
                print(f"Skipping empty layer: Product {prod_idx} - {sub_layer.Name}")
                continue

            suffix = get_layer_suffix(sub_layer.Name)
            if not suffix: 
                continue 

            # לוגיקת שרוולים
            base_ab_name = sub_layer.Name 
            if "sleeve" in base_ab_name.lower():
                base_ab_name = "Print_Sleeves"

            target_artboard = base_ab_name
            if prod_idx != "1":
                target_artboard = f"P{prod_idx}_{base_ab_name}"
            
            artboard_exists = False
            for ab in master_doc.Artboards:
                if target_artboard.lower() in ab.Name.lower():
                    artboard_exists = True
                    break
            
            if not artboard_exists:
                print(f"Skipping - No matching artboard found for: {target_artboard}")
                continue

            job_list.append({
                "prod_idx": prod_idx,
                "layer_name": sub_layer.Name,
                "artboard_target": target_artboard,
                "suffix": suffix
            })

    master_doc.Close(2) 

    if not job_list:
        print("No valid print jobs found (Skipped empty layers/missing artboards).")
        return []

    # --- שלב 2: ביצוע המשימות ---
    last_4 = str(order_number)[-4:]
    
    for job in job_list:
        print(f">> Processing Product {job['prod_idx']}: {job['layer_name']} ({job['suffix']})")
        
        base_name = f"{last_4}_{job['suffix']}"
        if job['prod_idx'] != "1":
            base_name = f"{last_4}_{job['prod_idx']}_{job['suffix']}"
            
        final_pdf_name = f"{base_name}.pdf"
        final_pdf_path = os.path.join(output_folder, final_pdf_name)
        
        # === שינוי לדריסה: מחיקת קובץ קיים במקום יצירת עותק ===
        if os.path.exists(final_pdf_path):
            try: os.remove(final_pdf_path)
            except: pass
        # ======================================================

        temp_ai = os.path.join(output_folder, "temp_work.ai")
        shutil.copyfile(source_file_path, temp_ai)
        
        try:
            work_doc = app.Open(temp_ai)
            
            # 1. ניקוי מוצרים
            for i in range(work_doc.Layers.Count, 0, -1):
                l = work_doc.Layers(i)
                if l.Name != job['prod_idx']:
                    try: l.Locked = False; l.Visible = True; l.Delete()
                    except: pass
                else:
                    l.Visible = True; l.Locked = False
                    for j in range(l.Layers.Count, 0, -1):
                        sub = l.Layers(j)
                        if sub.Name != job['layer_name']:
                            try: sub.Locked = False; sub.Visible = True; sub.Delete()
                            except: pass
                        else:
                            sub.Visible = True; sub.Locked = False
            
            # 2. ניקוי ארטבורדים
            target_idx = -1
            for k in range(work_doc.Artboards.Count):
                ab_name = work_doc.Artboards(k+1).Name
                if job['artboard_target'].lower() in ab_name.lower():
                    target_idx = k
                    break
            
            if target_idx != -1:
                work_doc.Artboards.SetActiveArtboardIndex(target_idx)
                for k in range(work_doc.Artboards.Count, 0, -1):
                    if (k-1) != target_idx:
                        try: work_doc.Artboards(k).Delete()
                        except: pass
            
            try:
                detect_res = app.DoJavaScript(JSX_DETECT_LOGIC)
                if str(detect_res).lower() == "true":
                    print("   > Black Print detected. Converting to White...")
                    app.DoJavaScript(JSX_RECOLOR_WHITE)
            except: pass

            pdf_opts = win32com.client.Dispatch("Illustrator.PDFSaveOptions")
            pdf_opts.PDFPreset = "[Press Quality]"
            
            work_doc.SaveAs(final_pdf_path, pdf_opts)
            work_doc.Close(2)
            
            files_created.append(final_pdf_path)
            print(f"   > Created: {final_pdf_name}")

        except Exception as e:
            print(f"Error processing job: {e}")
            try: work_doc.Close(2)
            except: pass
        
        if os.path.exists(temp_ai):
            try: os.remove(temp_ai)
            except: pass

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











        jsx_code = (


            JSX_PS_TEMPLATE


            .replace("FILE_PATH", json.dumps(abs_path))


            .replace("SAVE_PATH", json.dumps(abs_path))


            .replace("STATUS_PATH", json.dumps(status_path))


            .replace("CONTRACT_PX", str(contract_px))


        )











        











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
def create_simulation_summary_file(source_file_path, order_number, output_folder):
    """יוצר קובץ PDF מסכם של ההדמיות"""
    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch("Illustrator.Application")
    except:
        return

    last_4 = str(order_number)[-4:]
    final_pdf_name = f"{last_4}_print.pdf"
    final_pdf_path = os.path.join(output_folder, final_pdf_name)
    
    # === דריסה אם קיים ===
    if os.path.exists(final_pdf_path):
        try: os.remove(final_pdf_path)
        except: pass
    
    js_save_path = os.path.abspath(final_pdf_path).replace("\\", "/")
    
    print(f">> Generating Simulation Summary: {final_pdf_name}")
    
    try:
        app.Open(source_file_path)
        script_to_run = JSX_CREATE_SIM_SUMMARY.replace("SAVE_PATH", json.dumps(js_save_path))
        app.DoJavaScript(script_to_run)
        print(f"   > Created Summary: {final_pdf_name}")
    except Exception as e:
        print(f"   Error creating summary: {e}")

# 3. נקודת כניסה ראשית
# ========================================================
if __name__ == "__main__":
    if len(sys.argv) < 4: 
        print("Usage: script.py <input_file> <order_id> <thickness>")
        sys.exit()
    
    input_file, order_id, thickness = sys.argv[1], sys.argv[2], sys.argv[3]
    contract_px = int(thickness.replace("px", ""))

    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
        output_base = config.get("print_folder_path", "C:/Temp/Print")

    order_dir = os.path.join(output_base, str(order_id)[-4:])
    if not os.path.exists(order_dir): os.makedirs(order_dir)

    print("Step 1: Illustrator Splitting...")
    generated = run_illustrator_split(input_file, order_id, order_dir)
    
    # --- התוספת החדשה: יצירת קובץ הדמיות ---
    create_simulation_summary_file(input_file, order_id, order_dir)
    # ---------------------------------------
    
    if generated:
        print("Step 2: Photoshop Processing...")
        run_photoshop_processing(generated, contract_px)
    
    print("Done.")