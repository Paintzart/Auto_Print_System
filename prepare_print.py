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

# מיפוי מיקום (מהשרת) לסיומת – כדי להמיר ל־layer_name מתוך job_list
LOCATION_TO_SUFFIX = {
    "front": "PF",         # הדפס קידמי
    "back": "PB",          # הדפס אחורי
    "right_sleeve": "PR",  # שרוול ימין
    "left_sleeve": "PL",   # שרוול שמאל
}
def get_simulation_sublayer(print_layer_name):
    """מחזיר את שם תת-השכבה בהדמיה לפי שם שכבות ההדפסה"""
    name = print_layer_name.lower()
    if "front" in name: return "S_Placement_Front"
    if "back" in name: return "S_Placement_Back"
    if "sleeve" in name:
        if "right" in name: return "S_Placement_Right_Sleeve"
        if "left" in name: return "S_Placement_Left_Sleeve"
    return None
# בדיקה האם הגרפיקה שחורה והאם הרקע בהדמיה לבן
JSX_DETECT_LOGIC = """
#target illustrator
function checkDoubleCondition() {
    try {
        var doc = app.activeDocument;
        var printLayerName = "%PRINT_LAYER%";
        var simSubName = "%SIM_SUB%"; // שם תת-השכבה בהדמיה (למשל S_Placement_Front)
        var productIndex = "%PROD_IDX%"; // מספר המוצר (1, 2, וכו')
        // 1. בדיקת שכבת ההדפסה - האם היא שחורה?
        var pLayer = doc.layers.getByName(printLayerName);
        if (!pLayer.visible) pLayer.visible = true;
        if (pLayer.pageItems.length === 0) return "false";
        var isPrintBlack = quickScanColor(pLayer.pageItems, false); // מחפש שחור
        // 2. בדיקת שכבת ההדמיה - האם הלוגו שם לבן?
        // שכבות ההדמיה נמצאות בתוך שכבות המוצר (1, 2, וכו')
        var simLayer = null;
        try {
            // מחפש את שכבות ההדמיה בתוך שכבות המוצר
            var productLayer = doc.layers.getByName(productIndex);
            simLayer = productLayer.layers.getByName("Simulation");
        } catch(e) {
            // נסיון ברמה הראשית (למקרה של מבנה שונה)
            try { simLayer = doc.layers.getByName("Simulation"); } catch(e2) { return "false"; }
        }
        var simSub = null;
        try { simSub = simLayer.layers.getByName(simSubName); } catch(e) {
            try { simSub = simLayer.groupItems.getByName(simSubName); } catch(e) {}
        }
        if (!simSub) {
            // אין תת-שכבת הדמיה – בודקים רק לפי שחור בהדפסה
            return isPrintBlack ? "true" : "false";
        }
        var items = (simSub.typename === "Layer") ? simSub.pageItems : (simSub.pageItems || simSub.pathItems);
        var isSimWhite = quickScanColor(items, true); // מחפש לבן
        // רק אם אחד מהם (או שניהם) – זיהוי בהדמיה (לבן) או בקבצי הדפסה (שחור)
        return (isPrintBlack || isSimWhite) ? "true" : "false";
    } catch(e) { return "false"; }
}
function quickScanColor(items, findWhite) {
    // בודק את כל הפריטים
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename === 'GroupItem') {
            if (quickScanColor(item.pageItems, findWhite)) return true;
        } else if (item.typename === 'PathItem') {
            var c = null;
            if (item.filled) c = item.fillColor;
            else if (item.stroked) c = item.strokeColor;
            if (c) {
                if (findWhite && isWhite(c)) return true;
                if (!findWhite && isBlack(c)) return true;
            }
        } else if (item.typename === 'CompoundPathItem') {
             if (item.pathItems.length > 0) {
                 var p = item.pathItems[0];
                 var c = p.filled ? p.fillColor : p.strokeColor;
                 if (c) {
                    if (findWhite && isWhite(c)) return true;
                    if (!findWhite && isBlack(c)) return true;
                 }
             }
        }
        if (item.typename === 'PathItem') {
            count++;
            if (item.filled) c = item.fillColor;
            else if (item.stroked) c = item.strokeColor;
        }
        else if (item.typename === 'CompoundPathItem') {
             if (item.pathItems.length > 0) {
                 var p = item.pathItems[0];
                 var c = p.filled ? p.fillColor : p.strokeColor;
                 if (c) {
                    if (findWhite && isWhite(c)) return true;
                    if (!findWhite && isBlack(c)) return true;
                 }
             }
        }
        if (c) {
            if (findWhite && isWhite(c)) return true;
            if (!findWhite && isBlack(c)) return true;
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
checkDoubleCondition();
"""
# צביעת הגרפיקה בלבן נקי (CMYK 0,0,0,0)
JSX_CREATE_PRINT_FROM_SCRATCH = r"""
#target illustrator
// יצירת קובץ Print מאפס עבור מוצר 1 בלבד
function createPrintFromScratch() {
    var sourceDoc = app.activeDocument;
    var savePath = SAVE_PATH;
    try {
        // בדיקה אם יש רק מוצר 1
        var product1Layer = null;
        try {
            product1Layer = sourceDoc.layers.getByName("1");
        } catch(e) {
            return "ERROR: Product 1 layer not found";
        }
        if (!product1Layer) {
            return "ERROR: Product 1 layer not found";
        }
        // מציאת שכבות ההדמיה של מוצר 1
        var simLayer = null;
        try {
            simLayer = product1Layer.layers.getByName("Simulation");
        } catch(e) {
            return "ERROR: Simulation layer not found in Product 1";
        }
        if (!simLayer || (simLayer.pageItems.length === 0 && simLayer.layers.length === 0)) {
            return "ERROR: Simulation layer is empty";
        }
        // יצירת קבוצה זמנית מהשכבה
        var tempGroup = manualGroup(simLayer);
        // גודל A4 לרוחב (landscape)
        var A4_W = 841.89;  // רוחב A4 לרוחב בפוינטים
        var A4_H = 595.28;  // גובה A4 לרוחב בפוינטים
        var margin = 20;    // שוליים
        // יצירת מסמך חדש בגודל A4 לרוחב
        var newDoc = app.documents.add(DocumentColorSpace.CMYK, A4_W, A4_H);
        // הגדרת הארטבורד ל-A4
        var abRect = newDoc.artboards[0].artboardRect;
        var abLeft = abRect[0];
        var abTop = abRect[1];
        var abRight = abLeft + A4_W;
        var abBottom = abTop - A4_H;
        newDoc.artboards[0].artboardRect = [abLeft, abTop, abRight, abBottom];
        // העתקת הקבוצה למסמך החדש
        var newGroup = tempGroup.duplicate(newDoc.activeLayer, ElementPlacement.PLACEATEND);
        // חישוב גודל התוכן הזמין (A4 פחות שוליים)
        var availableWidth = A4_W - (margin * 2);
        var availableHeight = A4_H - (margin * 2);
        // חישוב יחס הגדלה
        var scaleX = (availableWidth / newGroup.width) * 100;
        var scaleY = (availableHeight / newGroup.height) * 100;
        var scale = Math.min(scaleX, scaleY);
        // הגדלת התוכן לגודל הדף (תמיד, גם אם הוא קטן יותר)
        newGroup.resize(scale, scale, true, true, true, true, scale);
        // מיקום במרכז הארטבורד A4
        var abCenterX = abLeft + (A4_W / 2);
        var abCenterY = abTop - (A4_H / 2);
        newGroup.position = [
            abCenterX - (newGroup.width / 2),
            abCenterY + (newGroup.height / 2)
        ];
        // שמירה
        var pdfOpts = new PDFSaveOptions();
        pdfOpts.presetFile = "[Press Quality]";
        newDoc.saveAs(new File(savePath), pdfOpts);
        newDoc.close(SaveOptions.DONOTSAVECHANGES);
        sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
        return "SUCCESS";
    } catch(e) {
        return "ERROR: " + e.message;
    }
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
            try { subLayer.pageItems[j].move(targetGroup, ElementPlacement.PLACEATBEGINNING); } catch(e) {}
        }
        flattenSubLayers(subLayer, targetGroup);
    }
}
function intersect(r1, r2) {
    return !(r2[0] > r1[2] || r2[2] < r1[0] || r2[1] < r1[3] || r2[3] > r1[1]);
}
createPrintFromScratch();
"""
JSX_CHECK_AND_RECOLOR_ALL = r"""
#target illustrator
// סקריפט לבדיקת צבעים וצביעה - בודק את כל שכבות ההדפסה ומצבע אוטומטית
function checkAndRecolorAllLayers() {
    var doc = app.activeDocument;
    var results = [];
    // רשימת כל שכבות ההדפסה האפשריות
    var printLayers = [];
    var simSubNames = [];
    // סריקת כל השכבות הראשיות
    for (var i = 0; i < doc.layers.length; i++) {
        var mainLayer = doc.layers[i];
        var layerName = mainLayer.name;
        // בדיקה אם זו שכבה של מוצר (1, 2, וכו')
        if (/^\d+$/.test(layerName)) {
            var productIndex = layerName;
            // סריקת תת-שכבות של ההדפסה
            for (var j = 0; j < mainLayer.layers.length; j++) {
                var subLayer = mainLayer.layers[j];
                var subName = subLayer.name;
                // זיהוי שכבות הדפסה
                if (subName.indexOf("Print_") === 0) {
                    printLayers.push({
                        product: productIndex,
                        layer: subName
                    });
                    // קביעת שם תת-השכבה בהדמיה
                    var simSub = getSimulationSubName(subName);
                    simSubNames.push(simSub);
                }
            }
        }
    }
    // בדיקה וצביעה לכל שכבה
    var recoloredCount = 0;
    for (var k = 0; k < printLayers.length; k++) {
        var printInfo = printLayers[k];
        var simSubName = simSubNames[k];
        try {
            var needsRecolor = checkDoubleCondition(doc, printInfo.product, printInfo.layer, simSubName);
            var wasRecolored = false;
            // אם צריך צביעה - מבצעים את הצביעה
            if (needsRecolor) {
                try {
                    wasRecolored = recolorLayerToWhite(doc, printInfo.product, printInfo.layer);
                    if (wasRecolored) recoloredCount++;
                } catch(recolorErr) {
                    // שגיאה בצביעה - נמשיך
                }
            }
            results.push({
                product: printInfo.product,
                layer: printInfo.layer,
                needsRecolor: needsRecolor,
                wasRecolored: wasRecolored,
                simSub: simSubName
            });
        } catch(e) {
            results.push({
                product: printInfo.product,
                layer: printInfo.layer,
                needsRecolor: false,
                wasRecolored: false,
                error: e.message
            });
        }
    }
    // בניית הודעת התוצאות
    var message = "=== תוצאות בדיקת צבעים ===\\n\\n";
    if (results.length === 0) {
        message += "לא נמצאו שכבות הדפסה בקובץ.\\n";
    } else {
        for (var r = 0; r < results.length; r++) {
            var res = results[r];
            var status = "";
            if (res.needsRecolor) {
                if (res.wasRecolored) {
                    status = "✅ צריך צביעה - בוצעה צביעה ללבן";
                } else {
                    status = "✅ צריך צביעה - שגיאה בביצוע הצביעה";
                }
            } else {
                status = "❌ אין צורך בצביעה";
            }
            message += "מוצר " + res.product + " - " + res.layer + ":\\n";
            message += "  " + status + "\\n";
            if (res.simSub) {
                message += "  (שכבת הדמיה: " + res.simSub + ")\\n";
            }
            if (res.error) {
                message += "  ⚠️ שגיאה: " + res.error + "\\n";
            }
            message += "\\n";
        }
    }
    message += "=== סיום בדיקה ===\\n";
    if (recoloredCount > 0) {
        message += "\\nסה\\\"כ " + recoloredCount + " שכבות נצבעו ללבן.";
    }
    return message;
}
function getSimulationSubName(printLayerName) {
    var name = printLayerName.toLowerCase();
    if (name.indexOf("front") !== -1) return "S_Placement_Front";
    if (name.indexOf("back") !== -1) return "S_Placement_Back";
    if (name.indexOf("sleeve") !== -1) {
        if (name.indexOf("right") !== -1) return "S_Placement_Right_Sleeve";
        if (name.indexOf("left") !== -1) return "S_Placement_Left_Sleeve";
    }
    return null;
}
function findSimSubByName(simLayer, simSubName) {
    var key = simSubName.toLowerCase();
    try {
        for (var i = 0; i < simLayer.layers.length; i++) {
            if (simLayer.layers[i].name.toLowerCase().indexOf("front") !== -1 && key.indexOf("front") !== -1) return simLayer.layers[i];
            if (simLayer.layers[i].name.toLowerCase().indexOf("back") !== -1 && key.indexOf("back") !== -1) return simLayer.layers[i];
            if (simLayer.layers[i].name.toLowerCase().indexOf("right") !== -1 && key.indexOf("right") !== -1) return simLayer.layers[i];
            if (simLayer.layers[i].name.toLowerCase().indexOf("left") !== -1 && key.indexOf("left") !== -1) return simLayer.layers[i];
        }
    } catch(e) {}
    return null;
}
function checkDoubleCondition(doc, productIndex, printLayerName, simSubName) {
    try {
        // 1. בדיקת שכבת ההדפסה - האם היא שחורה?
        var pLayer = null;
        try {
            pLayer = doc.layers.getByName(printLayerName);
        } catch(e) {
            try {
                var productLayer = doc.layers.getByName(productIndex);
                pLayer = productLayer.layers.getByName(printLayerName);
            } catch(e2) {
                return false;
            }
        }
        if (!pLayer || pLayer.pageItems.length === 0) return false;
        if (!pLayer.visible) pLayer.visible = true;
        var isPrintBlack = quickScanColor(pLayer.pageItems, false);
        // 2. בדיקת שכבת ההדמיה - האם הלוגו שם לבן?
        var simLayer = null;
        try {
            var productLayer = doc.layers.getByName(productIndex);
            simLayer = productLayer.layers.getByName("Simulation");
        } catch(e) {
            try { simLayer = doc.layers.getByName("Simulation"); } catch(e2) { return false; }
        }
        if (!simLayer) return false;
        var simSub = null;
        try {
            simSub = simLayer.layers.getByName(simSubName);
        } catch(e) {
            try { simSub = simLayer.groupItems.getByName(simSubName); } catch(e) {}
        }
        if (!simSub && simSubName) {
            simSub = findSimSubByName(simLayer, simSubName);
        }
        var items;
        if (!simSub) {
            // fallback: אין תת-שכבות (למשל PDF) – בודקים את שכבת Simulation כולה
            items = simLayer.pageItems || [];
        } else {
            items = (simSub.typename === "Layer") ? simSub.pageItems : (simSub.pageItems || simSub.pathItems || []);
        }
        var isSimWhite = (items.length > 0) ? quickScanColor(items, true) : false;
        // זיהוי בשניהם: צובעים אם שחור בהדפסה או לבן בהדמיה (גם וגם – לא או־או)
        return (isPrintBlack || isSimWhite);
    } catch(e) {
        return false;
    }
}
function quickScanColor(items, findWhite) {
    if (!items || items.length === 0) return false;
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename === 'GroupItem') {
            if (quickScanColor(item.pageItems, findWhite)) return true;
        } else if (item.typename === 'PathItem') {
            var c = null;
            if (item.filled) c = item.fillColor;
            else if (item.stroked) c = item.strokeColor;
            if (c) {
                if (findWhite && isWhite(c)) return true;
                if (!findWhite && isBlack(c)) return true;
            }
        } else if (item.typename === 'CompoundPathItem') {
             if (item.pathItems.length > 0) {
                 var p = item.pathItems[0];
                 var c = p.filled ? p.fillColor : p.strokeColor;
                 if (c) {
                    if (findWhite && isWhite(c)) return true;
                    if (!findWhite && isBlack(c)) return true;
                 }
             }
        } else if (item.typename === 'TextFrame') {
            if (item.textRange.characterAttributes.fillColor) {
                var c = item.textRange.characterAttributes.fillColor;
                if (findWhite && isWhite(c)) return true;
                if (!findWhite && isBlack(c)) return true;
            }
        }
    }
    return false;
}
function isWhite(c) {
    if (!c) return false;
    if (c.typename === 'CMYKColor') return c.cyan <= 10 && c.magenta <= 10 && c.yellow <= 10 && c.black <= 10;
    if (c.typename === 'RGBColor') return c.red >= 245 && c.green >= 245 && c.blue >= 245;
    if (c.typename === 'GrayColor') return c.gray <= 10;
    if (c.typename === 'SpotColor') return (typeof c.tint !== 'undefined' && c.tint <= 10);
    return false;
}
function isBlack(c) {
    if (!c) return false;
    if (c.typename === 'CMYKColor') return c.black >= 70 || (c.cyan >= 20 && c.black >= 60);
    if (c.typename === 'RGBColor') return c.red <= 30 && c.green <= 30 && c.blue <= 30;
    if (c.typename === 'GrayColor') return c.gray >= 70;
    return false;
}
function recolorLayerToWhite(doc, productIndex, printLayerName) {
    try {
        var pLayer = null;
        try {
            pLayer = doc.layers.getByName(printLayerName);
        } catch(e) {
            try {
                var productLayer = doc.layers.getByName(productIndex);
                pLayer = productLayer.layers.getByName(printLayerName);
            } catch(e2) {
                return false;
            }
        }
        if (!pLayer || pLayer.pageItems.length === 0) return false;
        if (!pLayer.visible) pLayer.visible = true;
        var white = new CMYKColor();
        white.cyan = 0;
        white.magenta = 0;
        white.yellow = 0;
        white.black = 0;
        recolorItems(pLayer.pageItems, white);
        return true;
    } catch(e) {
        return false;
    }
}
function recolorItems(items, color) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename === 'GroupItem') {
            recolorItems(item.pageItems, color);
        } else if (item.typename === 'PathItem' && !item.clipping) {
            item.filled = true;
            item.fillColor = color;
            item.stroked = false;
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
            if (item.textRange.characterAttributes.fillColor) {
                item.textRange.characterAttributes.fillColor = color;
            }
        }
    }
}
checkAndRecolorAllLayers();
"""
# הסרת שכבה/קבוצה "information" מתוך Simulation בכל מוצר (גם אם נעולה) – להכנת קובץ הדפסה
JSX_REMOVE_INFORMATION_FROM_SIMULATION = r"""
#target illustrator
function removeInformationFromSimulation() {
    var doc = app.activeDocument;
    for (var i = 0; i < doc.layers.length; i++) {
        var mainLayer = doc.layers[i];
        if (!/^\d+$/.test(mainLayer.name)) continue;
        try {
            var simLayer = mainLayer.layers.getByName("Simulation");
            simLayer.locked = false;
            simLayer.visible = true;
            // נסיון למחוק שכבת "information"
            try {
                var infoLayer = simLayer.layers.getByName("information");
                infoLayer.locked = false;
                infoLayer.visible = true;
                infoLayer.remove();
            } catch (e1) {
                // אולי זו קבוצה (group) בשם "information"
                try {
                    var infoGroup = simLayer.groupItems.getByName("information");
                    infoGroup.remove();
                } catch (e2) {}
            }
        } catch (e) {}
    }
}
removeInformationFromSimulation();
"""
# שינוי הארטבורד ל-A4 לרוחב ומרכוז התוכן (למצב 2Pocket)
JSX_RESIZE_ARTBOARD_A4_LANDSCAPE = r"""
#target illustrator
function resizeArtboardA4Landscape() {
    var doc = app.activeDocument;
    if (doc.artboards.length === 0) return "NO_ARTBOARD";
    var ab = doc.artboards[0];
    var A4_W = 841.89;
    var A4_H = 595.28;
    var margin = 20;
    var abRect = ab.artboardRect;
    var oldLeft = abRect[0], oldTop = abRect[1];
    ab.artboardRect = [oldLeft, oldTop, oldLeft + A4_W, oldTop - A4_H];
    var availW = A4_W - (margin * 2);
    var availH = A4_H - (margin * 2);
    var items = [];
    for (var i = 0; i < doc.pageItems.length; i++) items.push(doc.pageItems[i]);
    if (items.length === 0) return "OK";
    var group = doc.groupItems.add();
    group.name = "A4FitGroup";
    for (var m = doc.pageItems.length - 1; m >= 0; m--) {
        var item = doc.pageItems[m];
        if (item !== group) try { item.move(group, ElementPlacement.PLACEATBEGINNING); } catch(e) {}
    }
    var gW = group.width;
    var gH = group.height;
    if (gW <= 0) gW = 1;
    if (gH <= 0) gH = 1;
    var scaleGX = (availW / gW) * 100;
    var scaleGY = (availH / gH) * 100;
    var scaleG = Math.min(scaleGX, scaleGY);
    if (scaleG > 100) scaleG = 100;
    group.resize(scaleG, scaleG, true, true, true, true, scaleG);
    var newCX = oldLeft + A4_W / 2;
    var newCY = oldTop - A4_H / 2;
    group.position = [newCX - group.width / 2, newCY + group.height / 2];
    return "OK";
}
resizeArtboardA4Landscape();
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
        }
    }
}
runRecolor();
"""
# צביעה ללבן של שכבה בודדת לפי מוצר + שם שכבה (הוראה מהשרת, בלי זיהוי אוטומטי)
JSX_RECOLOR_ONE_LAYER_WHITE = r"""
#target illustrator
function runRecolorOneLayer() {
    var productIndex = "%PROD_IDX%";
    var printLayerName = "%LAYER_NAME%";
    try {
        var doc = app.activeDocument;
        var pLayer = null;
        try {
            var productLayer = doc.layers.getByName(productIndex);
            pLayer = productLayer.layers.getByName(printLayerName);
        } catch(e) {
            return "skip";
        }
        if (!pLayer || pLayer.pageItems.length === 0) return "skip";
        if (!pLayer.visible) pLayer.visible = true;
        var white = new CMYKColor();
        white.cyan = 0; white.magenta = 0; white.yellow = 0; white.black = 0;
        recolorItems(pLayer.pageItems, white);
        return "ok";
    } catch(e) { return "error"; }
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
            if (item.textRange.characterAttributes.fillColor) {
                item.textRange.characterAttributes.fillColor = color;
            }
        }
    }
}
runRecolorOneLayer();
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
def run_illustrator_split(source_file_path, order_number, output_folder, order_payload_path=None):
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
    # --- טעינת דגל הדפס קידמי 2Pocket (מהפיילוד) ---
    front_print_2pocket = False
    if order_payload_path and os.path.exists(order_payload_path):
        try:
            with open(order_payload_path, "r", encoding="utf-8") as f:
                payload_pre = json.load(f)
            front_print_2pocket = payload_pre.get("front_print_2pocket") or payload_pre.get("frontPrint2Pocket") or False
        except Exception:
            pass

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
            # מצב 2Pocket: רק הדפס קידמי, חיפוש ארטבורד שמכיל "2Pocket" + A4 לרוחב
            if front_print_2pocket and suffix == "PF":
                target_2pocket = None
                try:
                    for k in range(1, master_doc.Artboards.Count + 1):
                        ab = master_doc.Artboards(k)
                        ab_name_lower = ab.Name.lower()
                        if "2pocket" not in ab_name_lower:
                            continue
                        if prod_idx == "1":
                            if "p2" in ab_name_lower or "p3" in ab_name_lower:
                                continue
                            target_2pocket = ab.Name
                            break
                        if f"p{prod_idx}" in ab_name_lower or prod_idx in ab_name_lower:
                            target_2pocket = ab.Name
                            break
                    if not target_2pocket and prod_idx == "1":
                        for k in range(1, master_doc.Artboards.Count + 1):
                            ab = master_doc.Artboards(k)
                            if "2pocket" in ab.Name.lower():
                                target_2pocket = ab.Name
                                break
                except Exception as e:
                    print(f"   [אזהרה] שגיאה בסריקת ארטבורדים (2Pocket): {e}")
                if target_2pocket:
                    job_list.append({
                        "prod_idx": prod_idx,
                        "layer_name": sub_layer.Name,
                        "artboard_target": target_2pocket,
                        "suffix": suffix,
                        "a4_landscape": True
                    })
                continue
            # לוגיקת שרוולים (גישה לפי אינדקס – מונעת שגיאת COM "there is no document")
            base_ab_name = sub_layer.Name
            if "sleeve" in base_ab_name.lower():
                base_ab_name = "Print_Sleeves"
            target_artboard = base_ab_name
            if prod_idx != "1":
                target_artboard = f"P{prod_idx}_{base_ab_name}"
            artboard_exists = False
            try:
                for k in range(1, master_doc.Artboards.Count + 1):
                    ab = master_doc.Artboards(k)
                    if target_artboard.lower() in ab.Name.lower():
                        artboard_exists = True
                        break
            except Exception as e:
                print(f"   [אזהרה] שגיאה בסריקת ארטבורדים: {e}")
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
    # --- רשימת "הדפס לבן" ו־item_quantities מהשרת – נקרא מקובץ שהשרת שולח (order_payload_path) ---
    white_print_locations = []
    quantity_map = {}  # (product, suffix) -> quantity
    if order_payload_path and os.path.exists(order_payload_path):
        try:
            with open(order_payload_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            white_print_locations = payload.get("white_print_locations") or []
            if not isinstance(white_print_locations, list):
                white_print_locations = []
            item_quantities = payload.get("item_quantities") or []
            if not isinstance(item_quantities, list):
                item_quantities = []
            # בניית מפת כמות לפי (מוצר, צד)
            for iq in item_quantities:
                prod = str(iq.get("product") or iq.get("prod_idx") or "")
                loc = str(iq.get("location") or "").strip().lower()
                qty = iq.get("quantity")
                if prod and loc and qty is not None:
                    suffix = LOCATION_TO_SUFFIX.get(loc)
                    if suffix:
                        quantity_map[(prod, suffix)] = int(qty)
            print(f"   [דיבוג] נטען מ־payload: {len(white_print_locations)} מיקומי הדפס לבן, {len(quantity_map)} כמויות")
        except Exception as e:
            print(f"   [דיבוג] שגיאה בקריאת payload: {e}")
    else:
        if order_payload_path:
            print(f"   [דיבוג] קובץ payload לא נמצא: {order_payload_path}")
        else:
            print(f"   [דיבוג] לא הועבר נתיב payload – זיהוי אוטומטי")
    # --- שלב 2: ביצוע המשימות ---
    last_4 = str(order_number)[-4:]
    for job in job_list:
        print(f">> Processing Product {job['prod_idx']}: {job['layer_name']} ({job['suffix']})")
        base_name = f"{last_4}_{job['suffix']}"
        if job['prod_idx'] != "1":
            base_name = f"{last_4}_{job['prod_idx']}_{job['suffix']}"
        # הוספת כמות לשם הקובץ אם נשלחה (לפי מוצר וצד) – בסוגריים
        qty = quantity_map.get((str(job['prod_idx']), job['suffix']))
        if qty is not None:
            base_name = f"{base_name}({qty})"
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
            # הסרת שכבה/קבוצה "information" מתוך Simulation בכל מוצר (גם אם נעולה)
            try:
                app.DoJavaScript(JSX_REMOVE_INFORMATION_FROM_SIMULATION)
            except Exception:
                pass
            # צביעה ללבן: רק אם הרשימה ריקה – זיהוי אוטומטי (שחור בהדפסה + לבן בהדמיה). אם יש רשימה – צובעים רק את מה שברשימה, בלי שום בדיקה.
            try:
                if white_print_locations:
                    # רשימה מהשרת – רק מה שברשימה יצבע ללבן, בלי בדיקת צבע. שום דבר אחר לא נבדק ולא נצבע.
                    print(f"   > Recoloring to white by server list only ({len(white_print_locations)} location(s)), no color check.")
                    for loc in white_print_locations:
                        prod = loc.get("product") or loc.get("prod_idx")
                        location = loc.get("location")
                        layer_name = loc.get("layer") or loc.get("layer_name")
                        if not prod:
                            continue
                        if location is not None and location != "":
                            suffix = LOCATION_TO_SUFFIX.get(str(location).strip().lower())
                            if not suffix:
                                continue
                            layer_name = None
                            for j in job_list:
                                if j["prod_idx"] == str(prod) and j["suffix"] == suffix:
                                    layer_name = j["layer_name"]
                                    break
                        if not layer_name:
                            continue
                        jsx_one = JSX_RECOLOR_ONE_LAYER_WHITE.replace("%PROD_IDX%", str(prod)).replace("%LAYER_NAME%", str(layer_name))
                        try:
                            app.DoJavaScript(jsx_one)
                        except Exception:
                            pass
                    print(f"   > Server white-print recolor done.")
                else:
                    # רשימה ריקה – זיהוי אוטומטי: בודקים שחור בהדפסה + לבן בהדמיה ורק אז צובעים.
                    print(f"   > Checking and recoloring all print layers (auto)...")
                    result_message = app.DoJavaScript(JSX_CHECK_AND_RECOLOR_ALL)
                    print(f"   > Color check completed")
            except Exception as e:
                print(f"   > Error: Color detection/recoloring failed: {e}")
                import traceback
                traceback.print_exc()
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
            # מצב 2Pocket: שינוי גודל הדף ל-A4 לרוחב ומרכוז התוכן
            if job.get("a4_landscape"):
                try:
                    app.DoJavaScript(JSX_RESIZE_ARTBOARD_A4_LANDSCAPE)
                    print(f"   > A4 לרוחב (2Pocket) הוחל")
                except Exception as e:
                    print(f"   > אזהרה: שינוי A4 לרוחב נכשל: {e}")
            pdf_opts = win32com.client.Dispatch("Illustrator.PDFSaveOptions")
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
    ps = None
    max_attempts = 5
    delay_seconds = 3
    for attempt in range(1, max_attempts + 1):
        try:
            ps = win32com.client.Dispatch("Photoshop.Application")
            break
        except Exception as e:
            if attempt < max_attempts:
                print(f"   ניסיון {attempt}/{max_attempts}: חיבור ל-Photoshop נכשל, ממתין {delay_seconds} שניות...")
                time.sleep(delay_seconds)
            else:
                print(f"Error connecting to Photoshop: {e}")
                print("")
                print("  טיפים:")
                print("  • פתח את Photoshop ידנית לפני לחיצה על הכפתור (ורד/ורוד).")
                print("  • המתן עד שהחלון של Photoshop נטען לגמרי ואז הרץ שוב.")
                print("  • אם השגיאה חוזרת: הפעל את Select Graphic Automation Server כמנהל (Run as administrator).")
                print("")
                return
    if ps is None:
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
    # בדיקה אם יש רק מוצר 1
    is_only_product_1 = False
    try:
        temp_doc = app.Open(source_file_path)
        product_count = 0
        has_product_1 = False
        for layer in temp_doc.Layers:
            if layer.Name.strip().isdigit():
                product_count += 1
                if layer.Name.strip() == "1":
                    has_product_1 = True
        temp_doc.Close(2)
        is_only_product_1 = (product_count == 1 and has_product_1)
    except:
        pass
    try:
        app.Open(source_file_path)
        # הסרת שכבה/קבוצה "information" מתוך Simulation בכל מוצר (גם אם נעולה)
        try:
            app.DoJavaScript(JSX_REMOVE_INFORMATION_FROM_SIMULATION)
        except Exception:
            pass
        if is_only_product_1:
            # יצירת הקובץ מאפס עבור מוצר 1 בלבד
            print(f"   > Creating Print file from scratch for Product 1 only...")
            script_to_run = JSX_CREATE_PRINT_FROM_SCRATCH.replace("SAVE_PATH", json.dumps(js_save_path))
            result = app.DoJavaScript(script_to_run)
            if "SUCCESS" in str(result):
                print(f"   > Created Summary: {final_pdf_name}")
            else:
                print(f"   > Error creating from scratch: {result}")
                # נסיון עם הלוגיקה הרגילה
                script_to_run = JSX_CREATE_SIM_SUMMARY.replace("SAVE_PATH", json.dumps(js_save_path))
                app.DoJavaScript(script_to_run)
                print(f"   > Created Summary (fallback): {final_pdf_name}")
        else:
            # לוגיקה רגילה - עבור מספר מוצרים
            script_to_run = JSX_CREATE_SIM_SUMMARY.replace("SAVE_PATH", json.dumps(js_save_path))
            app.DoJavaScript(script_to_run)
            print(f"   > Created Summary: {final_pdf_name}")
    except Exception as e:
        print(f"   Error creating summary: {e}")
# 3. נקודת כניסה ראשית
# ========================================================
if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: script.py <input_file> <order_id> <thickness> [order_payload_path]")
        sys.exit()
    input_file, order_id, thickness = sys.argv[1], sys.argv[2], sys.argv[3]
    order_payload_path = sys.argv[4] if len(sys.argv) > 4 else None  # קובץ JSON מהשרת עם white_print_locations
    print(f"   [דיבוג] ארגומנטים: {len(sys.argv)}, order_payload_path={'כן: ' + order_payload_path if order_payload_path else 'לא'}")
    if order_payload_path:
        print(f"   [דיבוג] קובץ קיים? {os.path.exists(order_payload_path)}")
    contract_px = int(thickness.replace("px", ""))
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
        output_base = config.get("print_folder_path", "C:/Temp/Print")
    # יצירת תיקיית הבסיס אם היא לא קיימת (ללא תיקיית משנה לפי מספר הזמנה)
    if not os.path.exists(output_base): os.makedirs(output_base)
    print("Step 1: Illustrator Splitting...")
    generated = run_illustrator_split(input_file, order_id, output_base, order_payload_path)
    # --- התוספת החדשה: יצירת קובץ הדמיות ---
    create_simulation_summary_file(input_file, order_id, output_base)
    # ---------------------------------------
    if generated:
        print("Step 2: Photoshop Processing...")
        run_photoshop_processing(generated, contract_px)
    print("Done.")
