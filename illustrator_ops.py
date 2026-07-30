# -*- coding: utf-8 -*-
from __future__ import annotations
import win32com.client
import os
import uuid
import time
import json
from typing import Tuple, Optional
# --- הגדרות גלובליות ---
am = {
    "F": "Print_Front",
    "B": "Print_Back",
    "RS": "Print_Sleeves",
    "LS": "Print_Sleeves"
}
def hex_to_rgb(h: Optional[str]) -> Tuple[int, int, int]:
    if not h: return (0,0,0)
    h = h.lstrip('#')
    if len(h) != 6: return (0,0,0)
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
def run_jsx(app, s: str):
    """מריץ את ה-JSX עם הגנה מפני קריסות"""
    try:
        app.DoJavaScript(s)
    except Exception as e:
        print(f"   > JSX warning: {e}")


def save_native_ai(doc, path: str) -> None:
    """שמירה כ-.ai native (לא PDF עטוף ב-.ai) — נדרש לאיחוד merge."""
    path = os.path.abspath(path)
    folder = os.path.dirname(path)
    if folder:
        os.makedirs(folder, exist_ok=True)
    # ניסיון COM קודם — יציב יותר למסמכים גדולים (51 artboards)
    try:
        opts = win32com.client.Dispatch("Illustrator.IllustratorSaveOptions")
        opts.PDFCompatible = False
        opts.Compressed = False
        doc.SaveAs(path, opts)
        if not is_pdf_disguised_as_ai(path):
            return
        print(f"   ⚠ save_native_ai COM: {os.path.basename(path)} still PDF — retry JSX")
    except Exception as e:
        print(f"   ⚠ save_native_ai COM failed: {e}")
    safe_path = path.replace("\\", "/").replace('"', '\\"')
    safe_name = str(doc.Name).replace("\\", "\\\\").replace('"', '\\"').replace("'", "\\'")
    jsx = f"""
    #target illustrator
    (function() {{
        var wanted = "{safe_name}";
        var doc = null;
        for (var i = 0; i < app.documents.length; i++) {{
            if (app.documents[i].name === wanted) {{ doc = app.documents[i]; break; }}
        }}
        if (!doc) doc = app.activeDocument;
        if (!doc) return "no_doc";
        app.activeDocument = doc;
        var f = new File("{safe_path}");
        var opts = new IllustratorSaveOptions();
        opts.pdfCompatible = false;
        try {{ opts.compressed = true; }} catch(e) {{}}
        try {{
            opts.compatibility = Compatibility.ILLUSTRATOR24;
        }} catch(e2) {{
            try {{ opts.compatibility = Compatibility.ILLUSTRATOR17; }} catch(e3) {{}}
        }}
        doc.saveAs(f, opts);
        return "ok";
    }})();
    """
    app = doc.Application
    app.DoJavaScript(jsx)
    if is_pdf_disguised_as_ai(path):
        print(f"   ⚠ save_native_ai: {os.path.basename(path)} still PDF — retry uncompressed")
        jsx_retry = f"""
        #target illustrator
        (function() {{
            var wanted = "{safe_name}";
            var doc = null;
            for (var i = 0; i < app.documents.length; i++) {{
                if (app.documents[i].name === wanted) {{ doc = app.documents[i]; break; }}
            }}
            if (!doc) doc = app.activeDocument;
            if (!doc) return "no_doc";
            app.activeDocument = doc;
            var f = new File("{safe_path}");
            var opts = new IllustratorSaveOptions();
            opts.pdfCompatible = false;
            opts.compressed = false;
            doc.saveAs(f, opts);
            return "ok";
        }})();
        """
        app.DoJavaScript(jsx_retry)


def is_pdf_disguised_as_ai(path: str) -> bool:
    try:
        with open(path, "rb") as f:
            return f.read(5) == b"%PDF-"
    except OSError:
        return False
# --- סקריפטים JSX ---
JSX_CLEAN_MAGIC = """
#target illustrator
// פונקציה להשוואת צבעים
function isSameColor(c1, c2) {
    if (!c1 || !c2) return false;
    if (c1.typename !== c2.typename) return false;
    var t = 1;
    if (c1.typename === 'RGBColor') {
        return Math.abs(c1.red - c2.red) <= t &&
               Math.abs(c1.green - c2.green) <= t &&
               Math.abs(c1.blue - c2.blue) <= t;
    }
    if (c1.typename === 'CMYKColor') {
        return c1.cyan === c2.cyan && c1.magenta === c2.magenta &&
               c1.yellow === c2.yellow && c1.black === c2.black;
    }
    if (c1.typename === 'GrayColor') {
        return Math.abs(c1.gray - c2.gray) <= t;
    }
    return false;
}
// פונקציה לניקוי שאריות קטנות בצבע של הרקע
function removeInternalParts(container, bgCol) {
    for (var i = container.pageItems.length - 1; i >= 0; i--) {
        var item = container.pageItems[i];
        if (item.typename === 'GroupItem') {
            removeInternalParts(item, bgCol);
        }
        else if ((item.typename === 'PathItem' || item.typename === 'CompoundPathItem') && !item.clipping) {
            var colorMatch = false;
            if (item.typename === 'PathItem' && item.filled && isSameColor(item.fillColor, bgCol)) colorMatch = true;
            if (item.typename === 'CompoundPathItem' && item.pathItems.length > 0 &&
                item.pathItems[0].filled && isSameColor(item.pathItems[0].fillColor, bgCol)) colorMatch = true;
            if (colorMatch) {
                item.remove();
            }
        }
    }
}
function run(ln, grpN, r, g, b, doC, isRaster) {
    // בדיקה ראשונית: אם זה רסטר, מדלגים
    if (isRaster === true) {
        if (doC === true) {
            var c = new RGBColor(); c.red=r; c.green=g; c.blue=b;
            try {
                var groupRefresh = app.activeDocument.pageItems.getByName(grpN);
                if(groupRefresh) colRec(groupRefresh, c);
            } catch(e) {}
        }
        return;
    }
    // ----------------------------------------------------
    try {
        var doc = app.activeDocument;
        var group = doc.pageItems.getByName(grpN);
        // 1. ניקוי "זבל" ראשוני מתחתית הקבוצה (קוים שקופים וכו')
        // נעשה את זה פעמיים כדי לוודא שניקינו לכלוך
        for(var k=0; k<2; k++){
            try {
                var c = group.pageItems.length;
                if (c > 0) {
                    var last = group.pageItems[c - 1];
                    // אם זה path ללא מילוי וללא קו - למחוק
                    if (last.typename === "PathItem" && !last.filled && !last.stroked) last.remove();
                }
            } catch(e){}
        }
        if (group.typename === 'GroupItem' && group.pageItems.length > 0) {
            var gb = group.visibleBounds; // [Left, Top, Right, Bottom]
            var totalW = group.width;
            var totalH = group.height;
            var totalArea = totalW * totalH;
            var detectedBgColor = null;
            var keepPeeling = true;
            var safetyCounter = 0; // למנוע לולאה אינסופית
            // 2. לולאת "קילוף" - בודקים רק מלמטה!
            while (keepPeeling && group.pageItems.length > 0 && safetyCounter < 10) {
                safetyCounter++;
                // באילוסטרייטור: האינדקס הגבוה (length-1) הוא בדרך כלל הפריט הכי תחתון בקבוצה (Back)
                // אבל זה תלוי איך הקובץ נבנה. בדרך כלל הסריקה היא הפוכה.
                // בקוד הקודם עשינו i-- שזה אומר שהתחלנו מ- length-1.
                // לכן נבדוק את הפריט באינדקס [length-1] (הכי תחתון)
                var idx = group.pageItems.length - 1;
                var item = group.pageItems[idx];
                var iArea = item.width * item.height;
                var ib = item.visibleBounds;
                // בדיקת מגע בקצוות
                var tolerance = 2.0;
                var edgesTouching = 0;
                if (Math.abs(ib[0] - gb[0]) < tolerance) edgesTouching++; // L
                if (Math.abs(ib[1] - gb[1]) < tolerance) edgesTouching++; // T
                if (Math.abs(ib[2] - gb[2]) < tolerance) edgesTouching++; // R
                if (Math.abs(ib[3] - gb[3]) < tolerance) edgesTouching++; // B
                var isBackground = false;
                // תנאי א: נוגע ב-4 קצוות (רקע מלא)
                if (edgesTouching === 4) isBackground = true;
                // תנאי ב: נוגע ב-3 קצוות (חצי רקע) - חייב להיות לפחות 20% מהשטח כדי לא למחוק פסים דקים
                else if (edgesTouching === 3 && iArea > (totalArea * 0.20)) isBackground = true;
                // תנאי ג: נוגע ב-2 קצוות - חייב להיות גדול (40%) - רקע פינתי
                else if (edgesTouching >= 2 && iArea > (totalArea * 0.40)) isBackground = true;
                // תנאי ד: ענק ללא קשר לקצוות (95%)
                else if (iArea > (totalArea * 0.95)) isBackground = true;
                if (isBackground) {
                    // זיהינו רקע!
                    // נשמור את הצבע (רק של הרקע הראשון שנמצא)
                    if (!detectedBgColor) {
                        if (item.typename === 'PathItem' && item.filled) detectedBgColor = item.fillColor;
                        else if (item.typename === 'CompoundPathItem' && item.pathItems.length > 0 && item.pathItems[0].filled)
                            detectedBgColor = item.pathItems[0].fillColor;
                    }
                    // מחיקה
                    item.remove();
                    // ממשיכים בלולאה (keepPeeling נשאר true) כדי לבדוק את השכבה שמתחתיה שנחשפה עכשיו
                } else {
                    // הגענו לפריט שהוא לא רקע (למשל הלוגו)
                    // עוצרים מיד!!
                    keepPeeling = false;
                }
            }
            // 3. ניקוי עדין (חורים באותיות) - רק אם זוהה צבע רקע
            if (detectedBgColor) {
                removeInternalParts(group, detectedBgColor);
            }
        }
        // צביעה (אם נדרש)
        if (doC === true) {
            var c = new RGBColor(); c.red=r; c.green=g; c.blue=b;
            try {
                var groupRefresh = doc.pageItems.getByName(grpN);
                if(groupRefresh) colRec(groupRefresh, c);
            } catch(e) {}
        }
    } catch(e) { }
}
function colRec(it, c) {
    try {
        if (it.typename === 'GroupItem') {
            for (var i=0; i<it.pageItems.length; i++) colRec(it.pageItems[i], c);
        } else if (it.typename === 'PathItem' && !it.clipping) {
            if (it.stroked && !it.filled) { it.strokeColor = c; }
            it.filled=true; it.fillColor=c; it.stroked=false;
        } else if (it.typename === 'CompoundPathItem') {
            for (var j=0; j<it.pathItems.length; j++) {
                if (!it.pathItems[j].clipping) {
                    it.pathItems[j].filled=true; it.pathItems[j].fillColor=c; it.pathItems[j].stroked=false;
                }
            }
        }
    } catch(e) { }
}
try{
    var isR = ("%ISRASTER%" === "true");
    var doColor = ("%DOCOL%" === "true");
    run("%LNAME%", "%GNAME%", %R%, %G%, %B%, doColor, isR);
}catch(e){}
"""

JSX_RECOLOR_GROUP = """
#target illustrator
function colRecPrint(it, c) {
    try {
        if (it.typename === 'GroupItem') {
            for (var i=0; i<it.pageItems.length; i++) colRecPrint(it.pageItems[i], c);
        } else if (it.typename === 'PathItem' && !it.clipping) {
            if (it.stroked && !it.filled) { it.strokeColor = c; }
            it.filled=true; it.fillColor=c; it.stroked=false;
        } else if (it.typename === 'CompoundPathItem') {
            for (var j=0; j<it.pathItems.length; j++) {
                if (!it.pathItems[j].clipping) {
                    it.pathItems[j].filled=true; it.pathItems[j].fillColor=c; it.pathItems[j].stroked=false;
                }
            }
        } else if (it.typename === 'TextFrame') {
            try {
                var chars = it.textRange.characters;
                for (var ti = 0; ti < chars.length; ti++) {
                    chars[ti].characterAttributes.fillColor = c;
                    chars[ti].characterAttributes.filled = true;
                    chars[ti].characterAttributes.stroked = false;
                }
            } catch(e) {}
        }
    } catch(e) { }
}
function makePrintColor(r, g, b) {
    try {
        if (app.activeDocument.documentColorSpace == DocumentColorSpace.CMYK) {
            var cmykArr = app.convertSampleColor(
                ImageColorSpace.RGB, ColorModel.PROCESS, [r, g, b],
                ImageColorSpace.CMYK, ColorModel.PROCESS, []
            );
            var cm = new CMYKColor();
            cm.cyan = cmykArr[0];
            cm.magenta = cmykArr[1];
            cm.yellow = cmykArr[2];
            cm.black = cmykArr[3];
            return cm;
        }
    } catch(e) {}
    var rgb = new RGBColor();
    rgb.red = r; rgb.green = g; rgb.blue = b;
    return rgb;
}
try {
    var grpN = "%GNAME%";
    var group = app.activeDocument.pageItems.getByName(grpN);
    if (group) colRecPrint(group, makePrintColor(%R%, %G%, %B%));
} catch(e) {}
"""

JSX_DUPLICATE_AND_POS = """
#target illustrator
function runSim(originalName, simName, r, g, b, prefix, category, doRecolor) {
    try {
        var doc = app.activeDocument;
        var original = null;
        try { original = doc.pageItems.getByName(originalName); } catch(e) { return; }
        var simLayer = doc.layers.getByName("Simulation");
        var targetLayer = null;
        var sideName = "";
        if(prefix=="F") sideName = "Front";
        if(prefix=="B") sideName = "Back";
        if(prefix=="RS") sideName = "Right_Sleeve";
        if(prefix=="LS") sideName = "Left_Sleeve";
        try { targetLayer = simLayer.layers.getByName("S_Placement_" + sideName); }
        catch(e) { return; }
        targetLayer.visible = true;
        simLayer.visible = true;
        var simItem = original.duplicate(targetLayer, ElementPlacement.PLACEATEND);
        simItem.name = simName;
        simItem.hidden = false;
        if (doRecolor === true) {
            var c = new RGBColor(); c.red=r; c.green=g; c.blue=b;
            recolorItem(simItem, c);
        }
        doSmartPos(simItem, prefix, category);
        simItem.name = "";
    } catch(e) { }
}
function recolorItem(it, c) {
    if (it.typename === 'GroupItem') for(var i=0; i<it.pageItems.length; i++) recolorItem(it.pageItems[i], c);
    else if (it.typename === 'PathItem' && !it.clipping) { it.filled=true; it.fillColor=c; it.stroked=false; }
    else if (it.typename === 'CompoundPathItem') for(var j=0; j<it.pathItems.length; j++) if(!it.pathItems[j].clipping) { it.pathItems[j].filled=true; it.pathItems[j].fillColor=c; it.pathItems[j].stroked=false; }
}
function getDist(p1, p2) { return Math.sqrt(Math.pow(p2[0] - p1[0], 2) + Math.pow(p2[1] - p1[1], 2)); }
function doSmartPos(item, prefix, category) {
    var itemW = item.width; var itemH = item.height; if(itemH == 0) itemH = 1;
    var suffix = "A4_Square";
    var catLower = category.toLowerCase();
    if (category === "Sleeve2") suffix = "Sleeve2";
    else if (catLower.indexOf("sleeve") !== -1 || catLower.indexOf("9") !== -1 || catLower.indexOf("שרוול") !== -1) suffix = "Sleeve";
    else if (category === "Pocket") suffix = "Pocket";
    else if (category === "2Pocket") suffix = "2Pocket";
    else if (category === "2") { suffix = "2"; } // כאן היה חסר הסוגר שגרם לשגיאה!
    else if (category === "A3") suffix = "A3";
    else if (category === "A5") suffix = "A5";
    else if (category === "A4") {
        var ratio = itemW / itemH;
        if (ratio > 1.21) suffix = "A4_Landscape";
        else if (ratio < 0.75) suffix = "A4_Portrait";
        else suffix = "A4_Square";
    }
    var boxName = "S" + prefix + "_Box_" + suffix;
    var box = null;
    try { box = app.activeDocument.pageItems.getByName(boxName); } catch(e) { return; }
    var trueBoxW = 0; var trueBoxH = 0; var angleDeg = 0;
    if (box.typename === "PathItem" && box.pathPoints.length > 1) {
        var p0 = box.pathPoints[0].anchor; var p1 = box.pathPoints[1].anchor; var p2 = box.pathPoints[2].anchor;
        var d01 = getDist(p0, p1); var d12 = getDist(p1, p2);
        if (suffix === "Sleeve") { trueBoxW = Math.max(d01, d12); trueBoxH = Math.min(d01, d12); } else { trueBoxW = d01; trueBoxH = d12; }
        var angleRad = Math.atan2(p1[1] - p0[1], p1[0] - p0[0]); angleDeg = angleRad * 180 / Math.PI;
    } else { trueBoxW = box.width; trueBoxH = box.height; }
    var scaleW = (trueBoxW / itemW) * 100.0; var scaleH = (trueBoxH / itemH) * 100.0;
    var scale = Math.min(scaleW, scaleH);
    item.resize(scale, scale);
    var b = box.visibleBounds;
    var cx = b[0] + (b[2] - b[0])/2.0; var cy = b[1] - (b[1] - b[3])/2.0;
    item.position = [cx - item.width/2.0, cy + item.height/2.0];
    if (Math.abs(angleDeg) > 0.5) {
        if (Math.abs(angleDeg) > 90) angleDeg += 180;
        item.rotate(angleDeg);
    }
}
try { var doRec = ("%DORECOLOR%" === "true"); runSim("%ORIG%", "%SIM%", %R%, %G%, %B%, "%PRE%", "%CAT%", doRec); } catch(e) { }
"""
JSX_SMART_POS = """
#target illustrator
function findNamedItem(container, name) {
    try {
        if (container.pageItems) {
            for (var i = 0; i < container.pageItems.length; i++) {
                var it = container.pageItems[i];
                if (it.name === name) return it;
                if (it.typename === "GroupItem") {
                    var r = findNamedItem(it, name);
                    if (r) return r;
                }
            }
        }
        if (container.layers) {
            for (var j = 0; j < container.layers.length; j++) {
                var r2 = findNamedItem(container.layers[j], name);
                if (r2) return r2;
            }
        }
    } catch(e) {}
    return null;
}
function smartPos(itemName, prefix, category, resizeArtboard, isPrint, artboardName, layerName) {
    try {
        var doc = app.activeDocument;
        var item = doc.pageItems.getByName(itemName);
        item.hidden = false;
        var itemW = item.width; var itemH = item.height; if(itemH == 0) itemH = 1;
        var suffix = "A4_Square";
        var catLower = category.toLowerCase();
        // 1. זיהוי תיבת המיקום
        if (category === "Sleeve2") suffix = "Sleeve2";
        else if (catLower.indexOf("sleeve") !== -1 || catLower.indexOf("9") !== -1 || catLower.indexOf("שרוול") !== -1) suffix = "Sleeve";
        else if (category === "Pocket") suffix = "Pocket";
        else if (category === "2Pocket") suffix = "2Pocket";
        else if (category === "2") suffix = "2";
        else if (category === "A3") suffix = "A3";
        else if (category === "A5") suffix = "A5";
        else if (category === "A4") {
            var ratio = itemW / itemH;
            if (ratio > 1.21) suffix = "A4_Landscape";
            else if (ratio < 0.75) suffix = "A4_Portrait";
            else suffix = "A4_Square";
        }
        var boxPrefix = isPrint ? prefix : "S" + prefix;
        var boxName = boxPrefix + "_Box_" + suffix;
        var box = null;
        var usedTargetLayerBox = false;
        if (layerName) {
            try {
                var targetLayer = doc.layers.getByName(layerName);
                box = findNamedItem(targetLayer, boxName);
                if (box) usedTargetLayerBox = true;
            } catch(e) {}
        }
        if (!box) {
            try { box = doc.pageItems.getByName(boxName); } catch(e) {
                box = findNamedItem(doc, boxName);
            }
        }
        if (!box) return;
        var b = box.visibleBounds;
        var boxW = b[2] - b[0]; var boxH = b[1] - b[3];
        var cx = b[0] + boxW/2.0; var cy = b[1] - boxH/2.0;
        var scale = 100.0;
        // שינוי גודל האלמנט
        if (suffix === "Sleeve") {
            var maxW = 255.0; var maxH = 170.0;
            var scaleW = (maxW / itemW) * 100.0;
            var scaleH = (maxH / itemH) * 100.0;
            scale = Math.min(scaleW, scaleH);
        } else {
            var sW = (boxW / itemW) * 100.0; var sH = (boxH / itemH) * 100.0;
            scale = Math.min(sW, sH);
        }
        item.resize(scale, scale);
        item.position = [cx - item.width/2.0, cy + item.height/2.0];
        // בדיקה נוספת: אם A4 ריבוע, אחד מהמימדים קטן מ-19.9 ס"מ והשני הוא 24 ס"מ, להגדיל ל-19.9 ס"מ (רק בקבצי הדפסה!)
        // 19.9 ס"מ = 19.9 * 28.34645 = 564.194 נקודות
        // 24 ס"מ = 24 * 28.34645 = 680.3148 נקודות
        if (isPrint && category === "A4" && suffix === "A4_Square") {
            var currentW = item.width;
            var currentH = item.height;
            var minSizePoints = 19.9 * 28.34645; // המרה מס"מ לנקודות
            var targetSizePoints = 24.0 * 28.34645; // 24 ס"מ בנקודות
            var tolerance = 2.0; // טולרנס לבדיקת "בדיוק" 24 ס"מ (±2 נקודות)
            // בודקים אם אחד קטן מ-19.9 והשני הוא 24 ס"מ (בטולרנס)
            var wIsSmall = currentW < minSizePoints;
            var hIsSmall = currentH < minSizePoints;
            var wIs24 = Math.abs(currentW - targetSizePoints) <= tolerance;
            var hIs24 = Math.abs(currentH - targetSizePoints) <= tolerance;
            if ((wIsSmall && hIs24) || (hIsSmall && wIs24)) {
                // מחשבים את ה-scale הנוסף הנדרש מהגודל הנוכחי
                // כדי שהמימד הקטן יהיה לפחות 19.9 ס"מ (564.194 נקודות)
                var scaleW_additional = (minSizePoints / currentW) * 100.0;
                var scaleH_additional = (minSizePoints / currentH) * 100.0;
                // לוקחים את ה-scale הגדול יותר כדי שהמימד הקטן יהיה לפחות 19.9 ס"מ
                var additionalScale = Math.max(scaleW_additional, scaleH_additional);
                // מגדילים את האלמנט ב-scale נוסף (יחסי לגודל הנוכחי)
                item.resize(additionalScale, additionalScale);
                // מעדכנים את המיקום לאחר השינוי
                item.position = [cx - item.width/2.0, cy + item.height/2.0];
            }
        }
        // 2. שינוי גודל הדף (רק אם זה לא A4 מרובע!)
        // התיקון כאן: הוספת התנאי suffix !== "A4_Square"
        if (isPrint && resizeArtboard === true && artboardName && suffix !== "A4_Square") {
            try {
                var ab = doc.artboards.getByName(artboardName);
                var oldRect = ab.artboardRect;
                var newW = 595.28; // A4 width
                var newH = 841.89; // A4 height
                if (suffix === "2Pocket") {
                    newW = 841.89;
                    newH = 595.28;
                } else if (boxW > boxH) {
                    newW = 841.89;
                    newH = 595.28;
                }
                var minX = oldRect[0] + (newW / 2);
                var maxX = oldRect[2] - (newW / 2);
                var minY = oldRect[3] + (newH / 2);
                var maxY = oldRect[1] - (newH / 2);
                var targetX = cx;
                var targetY = cy;
                var finalX = Math.max(minX, Math.min(targetX, maxX));
                var finalY = Math.min(maxY, Math.max(minY, targetY));
                ab.artboardRect = [
                    finalX - newW/2,
                    finalY + newH/2,
                    finalX + newW/2,
                    finalY - newH/2
                ];
            } catch(e) { }
        }
        // הדפס משתנה: הזזה לארטבורד הממוספר (Print_Back_2 וכו')
        if (isPrint && artboardName && /_\d+$/.test(artboardName) && !usedTargetLayerBox) {
            var idxMatch = artboardName.match(/_(\d+)$/);
            var slotIdx = idxMatch ? parseInt(idxMatch[1], 10) : 1;
            if (slotIdx > 1) {
                try {
                    var abPrefix = artboardName.replace(/_\d+$/, "");
                    var refAb = doc.artboards.getByName(abPrefix + "_1");
                    var targetAb = doc.artboards.getByName(artboardName);
                    if (refAb && targetAb) {
                        var rr = refAb.artboardRect;
                        var tr = targetAb.artboardRect;
                        var dx = ((tr[0] + tr[2]) / 2) - ((rr[0] + rr[2]) / 2);
                        var dy = ((tr[1] + tr[3]) / 2) - ((rr[1] + rr[3]) / 2);
                        item.translate(dx, dy);
                    }
                } catch(e) { }
            }
        }
    } catch(e) { }
}
try { var isRes = ("%RES%" === "true"); var isP = ("%ISP%" === "true"); smartPos("%ITEM%", "%PRE%", "%CAT%", isRes, isP, "%ABNAME%", "%LNAME%"); } catch(e) { }
"""
JSX_COLOR_PROD = """
#target illustrator
function col(it, r, g, b, sr, sg, sb, noStroke) {
    var f = new RGBColor(); f.red=r; f.green=g; f.blue=b;
    var s = new RGBColor(); s.red=sr; s.green=sg; s.blue=sb;
    if(it.name && (it.name.indexOf("String")!==-1 || it.name.indexOf("מיתר")!==-1)) return;
    if(it.typename==='PathItem' && !it.clipping){
        it.filled=true; it.fillColor=f;
        it.stroked=true; it.strokeColor=s; it.strokeWidth=1; if(noStroke){ it.stroked=false; } // דילוג על מיתר עבור Hoodie
    } else if(it.typename==='CompoundPathItem'){
        for(var i=0; i<it.pathItems.length; i++){
            var p=it.pathItems[i];
            if(!p.clipping){ p.filled=true; p.fillColor=f; p.stroked=true; p.strokeColor=s; if(noStroke){ p.stroked=false; } }
        }
    } else if(it.typename==='GroupItem'){
        for(var j=0; j<it.pageItems.length; j++) col(it.pageItems[j], r, g, b, sr, sg, sb, noStroke);
    }
}
try {
    var d = app.activeDocument;
    var l = d.layers.getByName("Simulation");
    var mainGrp = null;
    try { mainGrp = l.groupItems.getByName("Simulation"); } catch(e) {}
    if(!mainGrp) try { mainGrp = l.groupItems.getByName("%PROD%"); } catch(e) {}
    if(!mainGrp) try { mainGrp = l.groupItems.getByName("Shirt"); } catch(e) {}
    if(mainGrp) {
        var isSplit = ("%IS_SPLIT%" === "true");
        var prodName = "%PROD%";
        var isHoodie = (prodName.indexOf("Hoodie") !== -1 || prodName.indexOf("Zippered") !== -1);
        var hasSide1 = false; var hasSide2 = false;
        var s1, s2;
        try { s1 = mainGrp.groupItems.getByName("Side1"); hasSide1 = true; } catch(e) {}
        try { s2 = mainGrp.groupItems.getByName("Side2"); hasSide2 = true; } catch(e) {}
        if(isSplit && hasSide1 && hasSide2) {
            col(s1, %R1%, %G1%, %B1%, %SR1%, %SG1%, %SB1%, isHoodie);
            col(s2, %R2%, %G2%, %B2%, %SR2%, %SG2%, %SB2%, isHoodie);
        } else {
            col(mainGrp, %R1%, %G1%, %B1%, %SR1%, %SG1%, %SB1%, isHoodie);
        }
    }
} catch(e) { }
"""
JSX_DEL = """
#target illustrator
try{app.activeDocument.artboards.getByName("%AB%").remove();}catch(e){}
try{app.activeDocument.textFrames.getByName("%TF%").remove();}catch(e){}
"""
JSX_CLEAN_BOXES = """
#target illustrator
(function() {
    function removeBoxesRecursive(container) {
        if (!container) return;
        try {
            if (container.pageItems) {
                for (var i = container.pageItems.length - 1; i >= 0; i--) {
                    try {
                        var it = container.pageItems[i];
                        var nm = it.name || "";
                        if (nm.indexOf("_Box_") !== -1) {
                            it.remove();
                        } else if (it.typename === "GroupItem") {
                            removeBoxesRecursive(it);
                        }
                    } catch(e) {}
                }
            }
        } catch(e) {}
        try {
            if (container.layers) {
                for (var j = 0; j < container.layers.length; j++) {
                    removeBoxesRecursive(container.layers[j]);
                }
            }
        } catch(e) {}
    }
    try { removeBoxesRecursive(app.activeDocument); } catch(e) {}
})();
"""
JSX_EXTRA_COLORS = """
#target illustrator
try {
    var doc = app.activeDocument;
    var container = null;
    try { container = doc.layers.getByName("Simulation").groupItems.getByName("Box_Color"); } catch(e) {}
    if (!container) try { container = doc.layers.getByName("Simulation").layers.getByName("Box_Color"); } catch(e) {}
    if (container) {
        var allData = %COLOR_ARRAY%;
        function applyStyle(item, rgb) {
            var c = new RGBColor(); c.red = rgb[0]; c.green = rgb[1]; c.blue = rgb[2];
            var black = new RGBColor(); black.red = 0; black.green = 0; black.blue = 0;
            item.filled = true; item.fillColor = c;
            item.stroked = true; item.strokeColor = black; item.strokeWidth = 0.5;
        }
        for (var i = 1; i <= 24; i++) {
            try {
                var box = container.pageItems.getByName("Color_" + i);
                if (i <= allData.length) {
                    var colors = allData[i-1];
                    if (colors.length === 1) {
                        applyStyle(box, colors[0]);
                    } else if (colors.length >= 2) {
                        // שימוש ב-geometricBounds לדיוק מתמטי (ללא ה-Stroke בחישוב)
                        var b = box.geometricBounds;
                        var left = b[0]; var top = b[1]; var right = b[2]; var bottom = b[3];
                        var w = right - left; var h = top - bottom;
                        var leftRect = box.parent.pathItems.rectangle(top, left, w/2, h);
                        applyStyle(leftRect, colors[0]);
                        var rightRect = box.parent.pathItems.rectangle(top, left + w/2, w/2, h);
                        applyStyle(rightRect, colors[1]);
                        box.remove();
                    }
                } else { box.remove(); }
            } catch(e) {}
        }
    }
} catch(e) {}
"""
JSX_MEASURE_FINAL = """
#target illustrator
try {
    var doc = app.activeDocument;
    // מחפש את הפריט לפי השם הייחודי שנתנו לו
    var item = doc.pageItems.getByName("%NAME%");
    // מחזיר את הרוחב הנוכחי והאמיתי אחרי כל השינויים
    item.width;
} catch(e) {
    0;
}
"""
# -------------------------
# פונקציות עזר
# -------------------------
def close_all_illustrator_documents(app, save: bool = False) -> None:
    """סוגר את כל המסמכים הפתוחים ב-Illustrator (נדרש לפני merge)."""
    close_opt = 1 if save else 2
    for _ in range(20):
        try:
            if app.Documents.Count == 0:
                break
            app.Documents(1).Close(close_opt)
        except Exception as e:
            print(f"⚠️ Could not close Illustrator document: {e}")
            try:
                run_jsx(app, "try{app.documents[0].close(SaveOptions.DONOTSAVECHANGES);}catch(e){}")
            except Exception:
                break
        time.sleep(0.2)
    try:
        remaining = app.Documents.Count
    except Exception:
        remaining = 0
    if remaining:
        print(f"⚠️ WARNING: {remaining} Illustrator document(s) still open")
    else:
        print("🔍 DEBUG: All Illustrator documents closed")

def get_doc_safe(app):
    for i in range(5):
        try:
            if app.Documents.Count > 0:
                return app.ActiveDocument
        except:
            time.sleep(0.5)
    return None
def get_layer(doc, name):
    def _walk(layer):
        try:
            if layer.Name == name:
                layer.Visible = True
                layer.Locked = False
                return layer
        except Exception:
            pass
        try:
            for i in range(1, layer.Layers.Count + 1):
                found = _walk(layer.Layers(i))
                if found:
                    return found
        except Exception:
            pass
        return None

    try:
        l = doc.Layers(name)
        l.Visible = True
        l.Locked = False
        return l
    except Exception:
        pass
    try:
        for i in range(1, doc.Layers.Count + 1):
            found = _walk(doc.Layers(i))
            if found:
                return found
    except Exception:
        pass
    return None
def clean_arts(grp):
    try:
        if grp.PageItems.Count > 0:
            last = grp.PageItems(grp.PageItems.Count)
            if getattr(last, "TypeName", "") in ["GroupItem", "PathItem"]:
                last.Delete()
    except: pass
# -------------------------
# פונקציות ראשיות
# -------------------------
def update_size_label(doc, app, name, w, txt):
    if w <= 1: return
    width_cm = int(round(w / 28.34645))
    final_text = f"{width_cm} ס\"מ רוחב הדפס {txt}"
    jsx = f"""
    try {{
        var doc = app.activeDocument;
        var simLayer = null;
        try {{ simLayer = doc.layers.getByName("Simulation"); }} catch(e) {{}}
        if (simLayer) {{
            function find(container, n) {{
                try {{ return container.textFrames.getByName(n); }} catch(e) {{}}
                if (container.groupItems) {{
                    for (var i=0; i<container.groupItems.length; i++) {{
                        var r = find(container.groupItems[i], n);
                        if (r) return r;
                    }}
                }}
                return null;
            }}
            var t = find(simLayer, "{name}");
            if (t) {{
                var p = t;
                while(p) {{ if(p.locked) p.locked=false; try{{p=p.parent;}}catch(e){{break;}} if(p.typename=="Layer") break; }}
                t.contents = '{final_text}';
            }}
        }}
    }} catch(e) {{ }}
    """
    run_jsx(app, jsx)
VARIABLE_SIDE_BASE = {
    "front": {"layer": "Print_Front", "artboard": "Print_Front", "prefix": "F"},
    "back": {"layer": "Print_Back", "artboard": "Print_Back", "prefix": "B"},
    "right_sleeve": {"layer": "Print_Right_Sleeve", "artboard": "Print_Sleeves", "prefix": "RS"},
    "left_sleeve": {"layer": "Print_Left_Sleeve", "artboard": "Print_Sleeves", "prefix": "LS"},
}

# פריסת ארטבורדים להדפס משתנה – רשת עם מרווחים, ללא חפיפות
VARIABLE_PRINT_COLS = 5
VARIABLE_PRINT_GAP_MIN = 150
VARIABLE_PRINT_MAX_COLS = 11
VARIABLE_PRINT_MAX_ROWS = 10


def variable_grid_columns(count: int) -> int:
    """עמודות לרשת — מקס ~11 בשורה (מגבלת קנבס עם artboards קיימים), עד 10 שורות."""
    if count <= 1:
        return 1
    if count <= 5:
        return count
    rows_at_max = (count + VARIABLE_PRINT_MAX_COLS - 1) // VARIABLE_PRINT_MAX_COLS
    if rows_at_max <= VARIABLE_PRINT_MAX_ROWS:
        return min(VARIABLE_PRINT_MAX_COLS, count)
    for cols in range(5, VARIABLE_PRINT_MAX_COLS + 1):
        rows = (count + cols - 1) // cols
        if rows <= VARIABLE_PRINT_MAX_ROWS:
            return cols
    return VARIABLE_PRINT_MAX_COLS


def activate_document(app, doc_name: str) -> bool:
    if not doc_name:
        return True
    safe = json.dumps(str(doc_name))
    jsx = f"""
    #target illustrator
    (function() {{
        var want = {safe};
        for (var i = 0; i < app.documents.length; i++) {{
            if (app.documents[i].name === want) {{
                app.activeDocument = app.documents[i];
                return "1";
            }}
        }}
        return "0";
    }})();
    """
    try:
        return str(app.DoJavaScript(jsx)).strip() == "1"
    except Exception as e:
        print(f"   > JSX warning: {e}")
        return False


def get_template_artboard_size(app, template_path: str) -> tuple[float, float]:
    """גודל artboard פעיל בתבנית AI (רוחב/גובה בנקודות)."""
    if not template_path or not os.path.exists(template_path):
        return 0.0, 0.0
    safe = template_path.replace("\\", "/").replace('"', '\\"')
    jsx = f"""
    #target illustrator
    (function() {{
        var f = new File("{safe}");
        if (!f.exists) return "0|0";
        var d = null;
        try {{ d = app.open(f); }} catch(e) {{ return "0|0"; }}
        var idx = 0;
        try {{ idx = d.artboards.getActiveArtboardIndex(); }} catch(e) {{}}
        var r = d.artboards[idx].artboardRect;
        var w = r[2] - r[0];
        var h = r[1] - r[3];
        d.close(SaveOptions.DONOTSAVECHANGES);
        return w + "|" + h;
    }})();
    """
    try:
        raw = str(app.DoJavaScript(jsx)).strip()
        parts = raw.split("|")
        return float(parts[0]), float(parts[1])
    except Exception:
        return 0.0, 0.0


def get_template_content_size(app, template_path: str) -> tuple[float, float]:
    """גודל תוכן ההדפסה על artboard התבנית (לא גודל הדף) — לרשת variants."""
    if not template_path or not os.path.exists(template_path):
        return 0.0, 0.0
    safe = template_path.replace("\\", "/").replace('"', '\\"')
    jsx = f"""
    #target illustrator
    (function() {{
        var f = new File("{safe}");
        if (!f.exists) return "0|0";
        var d = null;
        try {{ d = app.open(f); }} catch(e) {{ return "0|0"; }}
        var idx = 0;
        try {{ idx = d.artboards.getActiveArtboardIndex(); }} catch(e) {{}}
        var abRect = d.artboards[idx].artboardRect;
        var left = 1e15, top = -1e15, right = -1e15, bottom = 1e15;
        function grow(it) {{
            try {{
                var b = it.visibleBounds;
                if (b[0] < left) left = b[0];
                if (b[1] > top) top = b[1];
                if (b[2] > right) right = b[2];
                if (b[3] < bottom) bottom = b[3];
            }} catch(e) {{}}
        }}
        function scanLayer(layer) {{
            if (!layer || !layer.visible) return;
            for (var pi = 0; pi < layer.pageItems.length; pi++) {{
                var it = layer.pageItems[pi];
                var nm = it.name || "";
                if (nm.indexOf("_Box_") !== -1) continue;
                grow(it);
            }}
            for (var li = 0; li < layer.layers.length; li++) scanLayer(layer.layers[li]);
        }}
        for (var i = 0; i < d.layers.length; i++) scanLayer(d.layers[i]);
        d.close(SaveOptions.DONOTSAVECHANGES);
        if (right <= left + 1 || top <= bottom + 1) {{
            var w = abRect[2] - abRect[0];
            var h = abRect[1] - abRect[3];
            return w + "|" + h;
        }}
        return (right - left) + "|" + (top - bottom);
    }})();
    """
    try:
        raw = str(app.DoJavaScript(jsx)).strip()
        parts = raw.split("|")
        return float(parts[0]), float(parts[1])
    except Exception:
        return 0.0, 0.0


# Illustrator יציב עד ~30 artboards — מעל זה: artboard יחיד + grid
VARIABLE_SIM_MAX_ARTBOARDS = 30


def variable_sim_use_layers_only(variant_count: int) -> bool:
    return variant_count > VARIABLE_SIM_MAX_ARTBOARDS


def create_variable_prints_document(
    app,
    main_doc_name: str,
    side: str,
    page_w: float,
    page_h: float,
) -> str:
    """מסמך AI חדש — artboard ראשון בגודל התבנית + שכבת Print עם box מהמסמך הראשי."""
    cfg = VARIABLE_SIDE_BASE.get(side)
    if not cfg or page_w < 1 or page_h < 1:
        return ""
    if main_doc_name and not activate_document(app, main_doc_name):
        return ""
    base_layer = cfg["layer"]
    base_ab = cfg["artboard"]
    w = float(page_w)
    h = float(page_h)
    jsx = f"""
    #target illustrator
    (function() {{
        var mainDoc = app.activeDocument;
        var pageW = {w};
        var pageH = {h};
        var baseLayerName = {json.dumps(base_layer)};
        var baseAbName = {json.dumps(base_ab)};
        function findLayerByName(container, name) {{
            if (!container || !container.layers) return null;
            for (var li = 0; li < container.layers.length; li++) {{
                var lyr = container.layers[li];
                if (lyr.name === name) return lyr;
                var sub = findLayerByName(lyr, name);
                if (sub) return sub;
            }}
            return null;
        }}
        function hasPlacementBox(item) {{
            var n = item.name || "";
            if (n.indexOf("_Box_") !== -1) return true;
            if (item.typename === "GroupItem" && item.pageItems) {{
                for (var k = 0; k < item.pageItems.length; k++) {{
                    if (hasPlacementBox(item.pageItems[k])) return true;
                }}
            }}
            return false;
        }}
        var srcLayer = findLayerByName(mainDoc, baseLayerName);
        var newDoc = app.documents.add(DocumentColorSpace.CMYK, pageW, pageH);
        try {{ newDoc.artboards[0].name = baseAbName; }} catch(e) {{}}
        var dstLayer = newDoc.layers[0];
        dstLayer.name = baseLayerName;
        dstLayer.locked = false;
        if (srcLayer && srcLayer.pageItems) {{
            for (var pi = 0; pi < srcLayer.pageItems.length; pi++) {{
                var it = srcLayer.pageItems[pi];
                if (!hasPlacementBox(it)) continue;
                try {{
                    it.duplicate(dstLayer, ElementPlacement.PLACEATEND);
                }} catch(e2) {{}}
            }}
        }}
        return newDoc.name;
    }})();
    """
    try:
        name = str(app.DoJavaScript(jsx)).strip()
        if name and name != "undefined":
            print(f"   > Variable prints document: {name} ({w:.0f}x{h:.0f} pt)")
            return name
    except Exception as e:
        print(f"   > JSX warning (create prints doc): {e}")
    return ""


def ensure_variable_print_slot(
    app,
    side: str,
    slot_idx: int,
    total: int,
    doc_name: str = "",
    layers_only: bool = False,
    cell_w: float = 0,
    cell_h: float = 0,
    ab_w: float = 0,
    ab_h: float = 0,
    reposition_grid: bool = True,
) -> bool:
    """יוצר/מאמת slot — artboard _N בגודל תבנית (cell_w/h) או layers-only."""
    if slot_idx < 1 or total < 1:
        return False
    cfg = VARIABLE_SIDE_BASE.get(side)
    if not cfg:
        return False
    if doc_name and not activate_document(app, doc_name):
        return False
    base_layer = cfg["layer"]
    base_ab = cfg["artboard"]
    ab_prefix = cfg["layer"] if side in ("right_sleeve", "left_sleeve") else base_layer
    cols = variable_grid_columns(total) if total > 5 else min(VARIABLE_PRINT_COLS, max(1, total))
    gap_min = VARIABLE_PRINT_GAP_MIN
    layers_only_js = "true" if layers_only else "false"
    reposition_js = "true" if reposition_grid else "false"
    jsx = f"""
    #target illustrator
    (function() {{
        var doc = app.activeDocument;
        var slotIdx = {int(slot_idx)};
        var total = {int(total)};
        var cols = {int(cols)};
        var gapMin = {int(gap_min)};
        var layersOnly = {layers_only_js};
        var REPOSITION = {reposition_js};
        var CELL_W = {float(cell_w)};
        var CELL_H = {float(cell_h)};
        var AB_W = {float(ab_w)};
        var AB_H = {float(ab_h)};
        var baseLayerName = {json.dumps(base_layer)};
        var baseAbName = {json.dumps(base_ab)};
        var abPrefix = {json.dumps(ab_prefix)};
        var slotLayerName = baseLayerName + "_" + slotIdx;
        var slotAbName = abPrefix + "_" + slotIdx;

        function findLayerByName(container, name) {{
            if (!container || !container.layers) return null;
            for (var li = 0; li < container.layers.length; li++) {{
                var lyr = container.layers[li];
                if (lyr.name === name) return lyr;
                var sub = findLayerByName(lyr, name);
                if (sub) return sub;
            }}
            return null;
        }}
        function getPrintLayerContainer(baseLayer) {{
            try {{
                var p = baseLayer.parent;
                if (p && p.typename === "Layer") return p;
            }} catch(e) {{}}
            return doc;
        }}
        function addPrintLayer(container, layerName) {{
            var nl;
            if (container.typename === "Layer") nl = container.layers.add();
            else nl = doc.layers.add();
            nl.name = layerName;
            nl.zOrder(ZOrderMethod.BRINGTOFRONT);
            return nl;
        }}
        function addArtboardSafe(rect, abW, abH, gapX, gapY) {{
            var r = [rect[0], rect[1], rect[2], rect[3]];
            if ((r[2] - r[0]) < 1 || (r[1] - r[3]) < 1) return null;
            var tries = [
                r,
                [r[0], r[1] - (abH + gapY), r[2], r[3] - (abH + gapY)],
                [r[0] + (r[2] - r[0]) + gapMin, r[1], r[2] + (r[2] - r[0]) + gapMin, r[3]]
            ];
            for (var ti = 0; ti < tries.length; ti++) {{
                var t = tries[ti];
                try {{ return doc.artboards.add(t); }} catch(e1) {{
                    try {{
                        var ab = doc.artboards.add();
                        ab.artboardRect = t;
                        return ab;
                    }} catch(e2) {{}}
                }}
            }}
            return null;
        }}
        function hasPlacementBox(item) {{
            var n = item.name || "";
            if (n.indexOf("_Box_") !== -1) return true;
            if (item.typename === "GroupItem" && item.pageItems) {{
                for (var k = 0; k < item.pageItems.length; k++) {{
                    if (hasPlacementBox(item.pageItems[k])) return true;
                }}
            }}
            return false;
        }}
        function clearPrintContent(layer) {{
            if (!layer || !layer.pageItems) return;
            for (var i = layer.pageItems.length - 1; i >= 0; i--) {{
                try {{
                    var it = layer.pageItems[i];
                    if (!it) continue;
                    var n = it.name || "";
                    if (n.indexOf("_Box_") !== -1) continue;
                    it.remove();
                }} catch(e) {{}}
            }}
        }}
        function translateLayerContents(layer, dx, dy) {{
            if (!layer) return;
            for (var i = 0; i < layer.pageItems.length; i++) layer.pageItems[i].translate(dx, dy);
            for (var j = 0; j < layer.layers.length; j++) translateLayerContents(layer.layers[j], dx, dy);
        }}
        function centerOf(rect) {{
            return [ (rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2 ];
        }}
        function duplicateBoxes(fromLayer, toLayer, dx, dy) {{
            for (var pi = 0; pi < fromLayer.pageItems.length; pi++) {{
                var boxItem = fromLayer.pageItems[pi];
                if (!hasPlacementBox(boxItem)) continue;
                var dup = boxItem.duplicate(toLayer, ElementPlacement.PLACEATEND);
                dup.translate(dx, dy);
            }}
        }}
        function boundsOfAllArtboards() {{
            var left = 1e15, top = -1e15, right = -1e15, bottom = 1e15;
            for (var bi = 0; bi < doc.artboards.length; bi++) {{
                var r = doc.artboards[bi].artboardRect;
                if (r[0] < left) left = r[0];
                if (r[1] > top) top = r[1];
                if (r[2] > right) right = r[2];
                if (r[3] < bottom) bottom = r[3];
            }}
            return {{ left: left, top: top, right: right, bottom: bottom }};
        }}
        function isPrintGridArtboardName(nm) {{
            if (!nm) return false;
            if (nm === baseLayerName || nm === baseAbName) return true;
            var prefix = abPrefix + "_";
            if (nm.indexOf(prefix) === 0) {{
                var rest = nm.substring(prefix.length);
                return rest.length > 0 && !isNaN(parseInt(rest, 10));
            }}
            return false;
        }}
        function boundsOfProductArtboards() {{
            var left = 1e15, top = -1e15, right = -1e15, bottom = 1e15;
            var found = false;
            for (var bi = 0; bi < doc.artboards.length; bi++) {{
                var nm = doc.artboards[bi].name || "";
                if (isPrintGridArtboardName(nm)) continue;
                var r = doc.artboards[bi].artboardRect;
                found = true;
                if (r[0] < left) left = r[0];
                if (r[1] > top) top = r[1];
                if (r[2] > right) right = r[2];
                if (r[3] < bottom) bottom = r[3];
            }}
            if (!found) return boundsOfAllArtboards();
            return {{ left: left, top: top, right: right, bottom: bottom }};
        }}
        function slotRectFromSlot1(slot1Rect, w, h, gapX, gapY, columns, index1Based) {{
            var idx = index1Based - 1;
            var col = idx % columns;
            var row = Math.floor(idx / columns);
            var left = slot1Rect[0] + col * (w + gapX);
            var top = slot1Rect[1] - row * (h + gapY);
            return [left, top, left + w, top - h];
        }}
        function findPlacementBoxItem(layer) {{
            if (!layer || !layer.pageItems) return null;
            for (var pi = 0; pi < layer.pageItems.length; pi++) {{
                var it = layer.pageItems[pi];
                if (hasPlacementBox(it)) return it;
            }}
            return null;
        }}
        function cellMetrics(srcLayer, s1, slotIdx, columns, gapMin) {{
            var w = s1[2] - s1[0];
            var h = s1[1] - s1[3];
            var abW = (AB_W > 0) ? AB_W : w;
            var abH = (AB_H > 0) ? AB_H : h;
            var ox = s1[0], oy = s1[1];
            if (abW < 1) abW = w;
            if (abH < 1) abH = h;
            var gapX = Math.max(gapMin, Math.round(abW * 0.08));
            var gapY = Math.max(gapMin, Math.round(abH * 0.08));
            var maxRowsBand = {int(VARIABLE_PRINT_MAX_ROWS)};
            var idx = slotIdx - 1;
            var col = idx % columns;
            var row = Math.floor(idx / columns);
            var rowBand = Math.floor(row / maxRowsBand);
            var rowInBand = row % maxRowsBand;
            var bandW = columns * (abW + gapX);
            var oxBand = ox + rowBand * (bandW + gapMin * 2);
            var cellLeft = oxBand + col * (abW + gapX);
            var cellTop = oy - rowInBand * (abH + gapY);
            var c1x = ox + abW / 2;
            var c1y = oy - abH / 2;
            var cx = cellLeft + abW / 2;
            var cy = cellTop - abH / 2;
            return {{
                cellW: abW,
                cellH: abH,
                dx: cx - c1x,
                dy: cy - c1y,
                rect: [cellLeft, cellTop, cellLeft + abW, cellTop - abH]
            }};
        }}

        app.executeMenuCommand("unlockAll");
        if (findLayerByName(doc, slotLayerName)) return 1;

        if (slotIdx === 1) {{
            var baseLayer = findLayerByName(doc, baseLayerName);
            if (!baseLayer) return 0;
            var srcAb = null;
            try {{ srcAb = doc.artboards.getByName(baseAbName); }} catch(e) {{ return 0; }}
            if (!srcAb) return 0;
            baseLayer.locked = false;
            baseLayer.visible = true;
            clearPrintContent(baseLayer);
            var origRect = srcAb.artboardRect;
            var abW = origRect[2] - origRect[0];
            var abH = origRect[1] - origRect[3];
            if (AB_W > 0 && AB_H > 0) {{
                abW = AB_W;
                abH = AB_H;
            }} else if (CELL_W > 0 && CELL_H > 0) {{
                abW = CELL_W;
                abH = CELL_H;
            }}
            if (abW < 1) abW = 595.28;
            if (abH < 1) abH = 841.89;
            var gapX = Math.max(gapMin, Math.round(abW * 0.08));
            var gapY = Math.max(gapMin, Math.round(abH * 0.08));
            var targetRect = origRect;
            if (total > 5 && !layersOnly && REPOSITION) {{
                var bb = boundsOfProductArtboards();
                var anchorLeft = bb.right + gapX;
                var anchorTop = bb.top;
                targetRect = slotRectFromSlot1(
                    [anchorLeft, anchorTop, anchorLeft + abW, anchorTop - abH],
                    abW, abH, gapX, gapY, cols, 1
                );
                translateLayerContents(baseLayer, targetRect[0] - origRect[0], targetRect[1] - origRect[1]);
            }} else if (total > 5 && layersOnly && REPOSITION) {{
                var bb2 = boundsOfProductArtboards();
                var anchorLeft2 = bb2.right + gapX;
                var anchorTop2 = bb2.top;
                targetRect = slotRectFromSlot1(
                    [anchorLeft2, anchorTop2, anchorLeft2 + abW, anchorTop2 - abH],
                    abW, abH, gapX, gapY, cols, 1
                );
                translateLayerContents(baseLayer, targetRect[0] - origRect[0], targetRect[1] - origRect[1]);
            }} else if (AB_W > 0 && AB_H > 0 && !layersOnly) {{
                targetRect = [origRect[0], origRect[1], origRect[0] + abW, origRect[1] - abH];
            }}
            if (!layersOnly) {{
                srcAb.artboardRect = targetRect;
            }}
            srcAb.name = slotAbName;
            baseLayer.name = slotLayerName;
            return 1;
        }}

        var srcLayer = findLayerByName(doc, baseLayerName + "_1");
        var srcAb1 = null;
        try {{ srcAb1 = doc.artboards.getByName(abPrefix + "_1"); }} catch(e) {{ return 0; }}
        if (!srcLayer || !srcAb1) return 0;
        var s1 = srcAb1.artboardRect;
        var m = cellMetrics(srcLayer, s1, slotIdx, cols, gapMin);
        if (!layersOnly) {{
            var gapX2 = Math.max(gapMin, Math.round(m.cellW * 0.08));
            var gapY2 = Math.max(gapMin, Math.round(m.cellH * 0.08));
            var newAb = addArtboardSafe(m.rect, m.cellW, m.cellH, gapX2, gapY2);
            if (!newAb) return "ab_fail:" + slotIdx;
            newAb.name = slotAbName;
        }}
        var printHost = getPrintLayerContainer(srcLayer);
        addPrintLayer(printHost, slotLayerName);
        return "1";
    }})();
    """
    try:
        raw = str(app.DoJavaScript(jsx)).strip()
        if raw == "1":
            return True
        if raw.startswith("ab_fail:"):
            print(f"   > Slot {slot_idx}: artboard create failed at grid position")
        return False
    except Exception as e:
        print(f"   > JSX warning (slot {slot_idx}): {e}")
        return False


def setup_variable_print_slots(
    app,
    side: str,
    count: int,
    doc_name: str = "",
    layers_only: bool = False,
    cell_w: float = 0,
    cell_h: float = 0,
    ab_w: float = 0,
    ab_h: float = 0,
    reposition_grid: bool = True,
    anchor_only: bool = False,
) -> int:
    """יוצר שכבות + artboards Print_Back_1..N — גודל תא = תבנית (cell_w/h)."""
    if count < 1:
        return 0
    cfg = VARIABLE_SIDE_BASE.get(side)
    if not cfg:
        return 0
    cols = variable_grid_columns(count) if count > 5 else min(VARIABLE_PRINT_COLS, max(1, count))
    rows = (count + cols - 1) // cols
    gap_min = VARIABLE_PRINT_GAP_MIN
    mode = "layers-only, 1 artboard" if layers_only else f"{cols} cols x {rows} rows, gap>={gap_min}pt"
    if cell_w > 0 and cell_h > 0:
        mode += f", artboard={ab_w:.0f}x{ab_h:.0f}pt" if ab_w > 0 else f", cell={cell_w:.0f}x{cell_h:.0f}pt"
    if anchor_only:
        mode += ", anchor slot 1"
    print(f"   > Variable grid {side}: {count} slots ({mode})")
    if doc_name and not activate_document(app, doc_name):
        print(f"   > ERROR: document not active: {doc_name}")
        return 0
    if layers_only:
        cleanup_variable_layers_only_extras(app, side, doc_name)
    last_slot = 1 if anchor_only else count
    ready = 0
    for slot in range(1, last_slot + 1):
        if ensure_variable_print_slot(
            app,
            side,
            slot,
            count,
            doc_name,
            layers_only,
            cell_w,
            cell_h,
            ab_w,
            ab_h,
            reposition_grid,
        ):
            ready += 1
        elif slot == 1:
            print(f"   > ERROR: failed to create slot 1 ({cfg['layer']}_1)")
            return 0
    if anchor_only:
        print(f"   > Variable grid {side}: anchor ready — up to {count} pages on demand")
        return ready
    if ready < count:
        print(f"   > Variable grid {side}: only {ready}/{count} slots created")
    else:
        print(f"   > Variable grid {side}: {ready} slots ready")
    return ready


def remove_variable_placement_boxes(
    app, side: str, count: int, doc_name: str = ""
) -> None:
    """מוחק תיבות _Box_ מיותרות מכל שכבות ההדפסה המשתנה."""
    if count < 1:
        return
    cfg = VARIABLE_SIDE_BASE.get(side)
    if not cfg:
        return
    base_layer = cfg["layer"]
    if doc_name and not activate_document(app, doc_name):
        return
    jsx = f"""
    #target illustrator
    (function() {{
        var doc = app.activeDocument;
        var base = {json.dumps(base_layer)};
        var total = {int(count)};
        function findLayerByName(container, name) {{
            if (!container || !container.layers) return null;
            for (var li = 0; li < container.layers.length; li++) {{
                var lyr = container.layers[li];
                if (lyr.name === name) return lyr;
                var sub = findLayerByName(lyr, name);
                if (sub) return sub;
            }}
            return null;
        }}
        function isBoxItem(item) {{
            try {{
                var n = item.name || "";
                if (n.indexOf("_Box_") !== -1) return true;
                if (item.typename === "GroupItem" && item.pageItems) {{
                    for (var k = 0; k < item.pageItems.length; k++) {{
                        if (isBoxItem(item.pageItems[k])) return true;
                    }}
                }}
            }} catch(e) {{}}
            return false;
        }}
        function removeBoxesFromLayer(layer) {{
            if (!layer || !layer.pageItems) return;
            for (var pi = layer.pageItems.length - 1; pi >= 0; pi--) {{
                try {{
                    if (isBoxItem(layer.pageItems[pi])) layer.pageItems[pi].remove();
                }} catch(e) {{}}
            }}
            for (var si = 0; si < layer.layers.length; si++) removeBoxesFromLayer(layer.layers[si]);
        }}
        app.executeMenuCommand("unlockAll");
        removeBoxesFromLayer(findLayerByName(doc, base));
        for (var i = 1; i <= total; i++) {{
            removeBoxesFromLayer(findLayerByName(doc, base + "_" + i));
        }}
    }})();
    """
    run_jsx(app, jsx)


def cleanup_stray_print_on_simulation(
    app, prefix: str, doc_name: str = ""
) -> None:
    """מסיר הדפסות זולגות משכבת ההדמיה — קבוצות P_X_s* וכפילויות."""
    if doc_name and not activate_document(app, doc_name):
        return
    safe_prefix = prefix.replace("\\", "\\\\").replace('"', '\\"')
    side_map = {"F": "Front", "B": "Back", "RS": "Right_Sleeve", "LS": "Left_Sleeve"}
    sim_side = side_map.get(prefix, "Back")
    jsx = f"""
    #target illustrator
    (function() {{
        var doc = app.activeDocument;
        var printPrefix = "P_{safe_prefix}_s";
        var simSide = "{sim_side}";
        function removePrintGroups(container) {{
            if (!container || !container.pageItems) return;
            for (var pi = container.pageItems.length - 1; pi >= 0; pi--) {{
                try {{
                    var it = container.pageItems[pi];
                    var nm = it.name || "";
                    if (nm.indexOf(printPrefix) === 0 || nm.indexOf("P_{safe_prefix}_") === 0) {{
                        it.remove();
                    }}
                }} catch(e) {{}}
            }}
            if (container.layers) {{
                for (var li = 0; li < container.layers.length; li++) {{
                    removePrintGroups(container.layers[li]);
                }}
            }}
        }}
        function dedupeSimPlacement() {{
            try {{
                var simLayer = doc.layers.getByName("Simulation");
                var target = simLayer.layers.getByName("S_Placement_" + simSide);
                var keepers = [];
                for (var i = target.pageItems.length - 1; i >= 0; i--) {{
                    try {{
                        var it = target.pageItems[i];
                        var nm = it.name || "";
                        if (nm.indexOf("_Box_") !== -1) continue;
                        if (nm.indexOf("P_") === 0) {{ it.remove(); continue; }}
                        keepers.push(it);
                    }} catch(e) {{}}
                }}
                for (var k = 1; k < keepers.length; k++) {{
                    try {{ keepers[k].remove(); }} catch(e2) {{}}
                }}
            }} catch(e) {{}}
        }}
        try {{
            var simLayer = doc.layers.getByName("Simulation");
            removePrintGroups(simLayer);
        }} catch(e) {{}}
        dedupeSimPlacement();
    }})();
    """
    run_jsx(app, jsx)


def cleanup_orphan_variable_print_artboards(
    app, side: str, doc_name: str = ""
) -> None:
    """מסיר artboard/שכבת Print_Back יתומים אחרי ש-Print_Back_1 קיים (מונע כפילות דף 1)."""
    cfg = VARIABLE_SIDE_BASE.get(side)
    if not cfg:
        return
    base_layer = cfg["layer"]
    base_ab = cfg["artboard"]
    if doc_name and not activate_document(app, doc_name):
        return
    jsx = f"""
    #target illustrator
    (function() {{
        var doc = app.activeDocument;
        var baseLayer = {json.dumps(base_layer)};
        var baseAb = {json.dumps(base_ab)};
        function findLayerByName(container, name) {{
            if (!container || !container.layers) return null;
            for (var li = 0; li < container.layers.length; li++) {{
                var lyr = container.layers[li];
                if (lyr.name === name) return lyr;
                var sub = findLayerByName(lyr, name);
                if (sub) return sub;
            }}
            return null;
        }}
        var has1 = false;
        try {{
            doc.artboards.getByName(baseLayer + "_1");
            has1 = true;
        }} catch(e) {{}}
        if (!has1) return;
        for (var ai = doc.artboards.length - 1; ai >= 0; ai--) {{
            if (doc.artboards[ai].name === baseAb) {{
                try {{ doc.artboards.remove(ai); }} catch(e2) {{}}
            }}
        }}
        var orphan = findLayerByName(doc, baseLayer);
        if (orphan && orphan.name === baseLayer) {{
            try {{ orphan.remove(); }} catch(e3) {{}}
        }}
    }})();
    """
    run_jsx(app, jsx)


def cleanup_variable_layers_only_extras(
    app, side: str, doc_name: str = ""
) -> None:
    """מוחק artboards יתומים Print_*_2..N מריצות קודמות (layers-only = artboard _1 בלבד)."""
    cfg = VARIABLE_SIDE_BASE.get(side)
    if not cfg:
        return
    ab_prefix = cfg["layer"]
    if doc_name and not activate_document(app, doc_name):
        return
    safe_prefix = json.dumps(ab_prefix)
    jsx = f"""
    #target illustrator
    (function() {{
        var doc = app.activeDocument;
        var abPrefix = {safe_prefix};
        for (var i = doc.artboards.length - 1; i >= 0; i--) {{
            var nm = doc.artboards[i].name;
            if (nm.indexOf(abPrefix + "_") !== 0) continue;
            var m = nm.match(/_(\\d+)$/);
            if (m && parseInt(m[1], 10) > 1) {{
                try {{ doc.artboards.remove(i); }} catch(e) {{}}
            }}
        }}
    }})();
    """
    run_jsx(app, jsx)


def expand_variable_layers_only_artboard(
    app,
    side: str,
    count: int,
    doc_name: str = "",
    cell_w: float = 0,
    cell_h: float = 0,
) -> None:
    """מרחיב artboard _1 לכיסוי רשת התוכן (גודל תיאורטי + bounds בפועל)."""
    if count < 1:
        return
    cfg = VARIABLE_SIDE_BASE.get(side)
    if not cfg:
        return
    ab_prefix = cfg["layer"]
    cols = variable_grid_columns(count) if count > 5 else min(VARIABLE_PRINT_COLS, max(1, count))
    gap_min = VARIABLE_PRINT_GAP_MIN
    if doc_name and not activate_document(app, doc_name):
        return
    safe_prefix = json.dumps(ab_prefix)
    jsx = f"""
    #target illustrator
    (function() {{
        var doc = app.activeDocument;
        var abPrefix = {safe_prefix};
        var total = {int(count)};
        var COLS = {int(cols)};
        var CELL_W = {float(cell_w)};
        var CELL_H = {float(cell_h)};
        var gapMin = {int(gap_min)};
        var pad = 30;
        function findPlacementBoxItem(layer) {{
            if (!layer || !layer.pageItems) return null;
            for (var pi = 0; pi < layer.pageItems.length; pi++) {{
                var it = layer.pageItems[pi];
                if (hasPlacementBox(it)) return it;
            }}
            return null;
        }}
        function findLayerByName(container, name) {{
            if (!container || !container.layers) return null;
            for (var li = 0; li < container.layers.length; li++) {{
                var lyr = container.layers[li];
                if (lyr.name === name) return lyr;
                var sub = findLayerByName(lyr, name);
                if (sub) return sub;
            }}
            return null;
        }}
        function hasPlacementBox(item) {{
            var n = item.name || "";
            if (n.indexOf("_Box_") !== -1) return true;
            if (item.typename === "GroupItem" && item.pageItems) {{
                for (var k = 0; k < item.pageItems.length; k++) {{
                    if (hasPlacementBox(item.pageItems[k])) return true;
                }}
            }}
            return false;
        }}
        function growBounds(b, item) {{
            try {{
                var vb = item.visibleBounds;
                if (vb[0] < b.left) b.left = vb[0];
                if (vb[1] > b.top) b.top = vb[1];
                if (vb[2] > b.right) b.right = vb[2];
                if (vb[3] < b.bottom) b.bottom = vb[3];
            }} catch(e) {{}}
        }}
        function scanLayer(layer, b) {{
            if (!layer) return;
            for (var pi = 0; pi < layer.pageItems.length; pi++) {{
                var it = layer.pageItems[pi];
                if (hasPlacementBox(it)) continue;
                growBounds(b, it);
            }}
            for (var si = 0; si < layer.layers.length; si++) scanLayer(layer.layers[si], b);
        }}
        var ab1 = null;
        try {{ ab1 = doc.artboards.getByName(abPrefix + "_1"); }} catch(e) {{ return; }}
        if (!ab1) return;
        var s1 = ab1.artboardRect;
        var b = {{ left: s1[0], top: s1[1], right: s1[2], bottom: s1[3] }};
        for (var i = 1; i <= total; i++) {{
            scanLayer(findLayerByName(doc, abPrefix + "_" + i), b);
        }}
        var originX = s1[0], originY = s1[1];
        var layer1 = findLayerByName(doc, abPrefix + "_1");
        var box1 = findPlacementBoxItem(layer1);
        if (box1) {{
            var bb1 = box1.visibleBounds;
            originX = bb1[0];
            originY = bb1[1];
        }}
        var thRight = b.right, thBottom = b.bottom;
        if (CELL_W > 0 && CELL_H > 0) {{
            var rows = Math.ceil(total / COLS);
            var gapX = Math.max(gapMin, Math.round(CELL_W * 0.08));
            var gapY = Math.max(gapMin, Math.round(CELL_H * 0.08));
            thRight = originX + COLS * CELL_W + (COLS - 1) * gapX;
            thBottom = originY - (rows * CELL_H + (rows - 1) * gapY);
        }}
        if (b.right <= b.left + 1 && thRight <= originX + 1) return;
        var newLeft = Math.min(s1[0], originX, b.left) - pad;
        var newTop = Math.max(s1[1], originY, b.top) + pad;
        var newRight = Math.max(s1[2], thRight, b.right) + pad;
        var newBottom = Math.min(s1[3], thBottom, b.bottom) - pad;
        if (newRight <= newLeft + 1 || newTop <= newBottom + 1) return;
        try {{
            ab1.artboardRect = [newLeft, newTop, newRight, newBottom];
            var w = newRight - newLeft;
            var h = newTop - newBottom;
            return "ok|" + Math.round(w) + "|" + Math.round(h);
        }} catch(e1) {{
            return "fail|" + e1;
        }}
    }})();
    """
    try:
        raw = str(app.DoJavaScript(jsx)).strip()
        if raw.startswith("ok|"):
            parts = raw.split("|")
            if len(parts) >= 3:
                print(f"   > Expanded {ab_prefix}_1 artboard to {parts[1]}x{parts[2]} pt")
        elif raw.startswith("fail|"):
            print(f"   > JSX warning (expand artboard): {raw[5:]}")
    except Exception as e:
        print(f"   > JSX warning (expand artboard): {e}")



def apply_text_overrides_in_layer(app, layer_name: str, overrides: dict) -> None:
    if not overrides or not layer_name:
        return
    pairs_js = json.dumps(
        {str(k): str(v) for k, v in overrides.items()},
        ensure_ascii=False,
    )
    safe_layer = layer_name.replace("\\", "\\\\").replace('"', '\\"')
    jsx = f"""
    #target illustrator
    (function() {{
        var doc = app.activeDocument;
        var overrides = {pairs_js};
        var layer = null;
        try {{ layer = doc.layers.getByName("{safe_layer}"); }} catch(e) {{ return; }}
        layer.locked = false;
        layer.visible = true;
        function findTf(container, name) {{
            try {{
                if (container.textFrames) {{
                    var tf = container.textFrames.getByName(name);
                    if (tf) return tf;
                }}
            }} catch(e) {{}}
            if (container.pageItems) {{
                for (var i = 0; i < container.pageItems.length; i++) {{
                    var it = container.pageItems[i];
                    if (it.typename === "TextFrame" && it.name === name) return it;
                    if (it.typename === "GroupItem") {{
                        var r = findTf(it, name);
                        if (r) return r;
                    }}
                }}
            }}
            if (container.layers) {{
                for (var j = 0; j < container.layers.length; j++) {{
                    var r2 = findTf(container.layers[j], name);
                    if (r2) return r2;
                }}
            }}
            return null;
        }}
        for (var key in overrides) {{
            if (!overrides.hasOwnProperty(key)) continue;
            var tf = findTf(layer, key);
            if (tf) {{
                tf.locked = false;
                tf.contents = overrides[key];
            }}
        }}
    }})();
    """
    run_jsx(app, jsx)


def illustrator_purge(app) -> None:
    """מנקה cache של Illustrator — מונע קריסות אחרי עשרות פעולות."""
    jsx = """
    #target illustrator
    (function() {
        try { app.redraw(); } catch(e) {}
        try { app.executeMenuCommand("purge"); } catch(e2) {}
    })();
    """
    try:
        app.DoJavaScript(jsx)
    except Exception:
        pass


def close_stray_template_documents(
    app, template_path: str, product_doc_name: str = ""
) -> None:
    """סוגר עותקי תבנית פתוחים שנשארו מהרצות קודמות."""
    base = os.path.splitext(os.path.basename(template_path))[0]
    safe_base = json.dumps(base)
    safe_prod = json.dumps(product_doc_name or "")
    jsx = f"""
    #target illustrator
    (function() {{
        var base = {safe_base};
        var prod = {safe_prod};
        for (var i = app.documents.length - 1; i >= 0; i--) {{
            try {{
                var dn = app.documents[i].name;
                if (prod && dn === prod) continue;
                if (dn.indexOf(base) === 0) {{
                    app.documents[i].close(SaveOptions.DONOTSAVECHANGES);
                }}
            }} catch(e) {{}}
        }}
    }})();
    """
    try:
        app.DoJavaScript(jsx)
    except Exception:
        pass


def open_variable_template_session(
    app, template_path: str, product_doc_name: str = ""
) -> str:
    """פותח תבנית AI פעם אחת לכל variants — בלי open/close בכל variant."""
    if not template_path or not os.path.exists(template_path):
        return ""
    close_stray_template_documents(app, template_path, product_doc_name)
    safe_path = json.dumps(template_path.replace("\\", "/"))
    jsx = f"""
    #target illustrator
    (function() {{
        var f = new File({safe_path});
        if (!f.exists) return "";
        var doc = app.open(f);
        return doc ? doc.name : "";
    }})();
    """
    try:
        name = str(app.DoJavaScript(jsx)).strip()
        if name and name != "undefined":
            return name
    except Exception as e:
        print(f"   > Template session open failed: {e}")
    return ""


def close_variable_template_session(app, template_doc_name: str) -> None:
    if not template_doc_name:
        return
    safe = json.dumps(template_doc_name)
    jsx = f"""
    #target illustrator
    (function() {{
        var want = {safe};
        for (var i = 0; i < app.documents.length; i++) {{
            if (app.documents[i].name === want) {{
                try {{ app.documents[i].close(SaveOptions.DONOTSAVECHANGES); }} catch(e) {{}}
                return;
            }}
        }}
    }})();
    """
    try:
        app.DoJavaScript(jsx)
    except Exception:
        pass


def capture_template_text_snapshot(app, template_doc_name: str) -> dict:
    """שומר תוכן TEXT frames מקורי לשחזור בין variants."""
    if not template_doc_name or not activate_document(app, template_doc_name):
        return {}
    safe = json.dumps(template_doc_name)
    jsx = f"""
    #target illustrator
    (function() {{
        var doc = app.activeDocument;
        var snap = {{}};
        function isVarName(nm) {{
            if (!nm) return false;
            if (/^(TEXT|NUM|NUMBER)[_]?\\d+$/i.test(nm)) return true;
            if (/^TEXT_(NAME|NUMBER|NUM\\d+)$/i.test(nm)) return true;
            return false;
        }}
        function walk(root) {{
            if (!root) return;
            try {{
                if (root.textFrames) {{
                    for (var t = 0; t < root.textFrames.length; t++) {{
                        try {{
                            var tf = root.textFrames[t];
                            var nm = tf.name || "";
                            if (isVarName(nm)) snap[nm] = String(tf.contents);
                        }} catch(e) {{}}
                    }}
                }}
            }} catch(e) {{}}
            try {{
                if (root.pageItems) {{
                    for (var i = 0; i < root.pageItems.length; i++) {{
                        try {{
                            if (root.pageItems[i].typename === "GroupItem") walk(root.pageItems[i]);
                        }} catch(e2) {{}}
                    }}
                }}
            }} catch(e) {{}}
            try {{
                if (root.layers) {{
                    for (var j = 0; j < root.layers.length; j++) walk(root.layers[j]);
                }}
            }} catch(e) {{}}
        }}
        walk(doc);
        var parts = [];
        for (var k in snap) {{
            if (snap.hasOwnProperty(k)) parts.push(k + "\\t" + snap[k]);
        }}
        return parts.join("\\n");
    }})();
    """
    try:
        raw = str(app.DoJavaScript(jsx)).strip()
        result: dict = {}
        for line in raw.split("\n"):
            if "\t" not in line:
                continue
            key, val = line.split("\t", 1)
            result[key] = val
        return result
    except Exception as e:
        print(f"   > Template snapshot failed: {e}")
        return {}


def place_variable_template_variant(
    app,
    template_path: str,
    product_doc_name: str,
    layer_name: str,
    artboard_name: str,
    prefix: str,
    text_overrides: Optional[dict] = None,
    image_files: Optional[dict] = None,
    image_raster_flags: Optional[dict] = None,
    outline_text: bool = True,
    skip_simulation: bool = False,
    sim_hex: Optional[str] = None,
    print_hex: Optional[str] = None,
    product_doc=None,
    category: str = "A4",
    shared_template: bool = True,
    slot_id: int = 1,
    layers_only: bool = False,
    grid_cols: int = VARIABLE_PRINT_COLS,
    total_slots: int = 1,
    sim_only: bool = False,
    center_on_artboard: bool = False,
    template_doc_name: str = "",
    keep_template_open: bool = False,
    text_snapshot: Optional[dict] = None,
) -> float:
    """מלביש תוכן מתבנית AI ב-1:1 — שומר גופן, גודל ומיקום יחסי בתוך הארטבורד (בלי scale ובלי מירכוז)."""
    if not template_path or not os.path.exists(template_path):
        print(f"   > Template missing: {template_path}")
        return 0.0
    text_overrides = text_overrides or {}
    image_files = image_files or {}
    image_raster_flags = image_raster_flags or {}
    safe_tpl = template_path.replace("\\", "/").replace('"', '\\"')
    safe_doc = product_doc_name.replace("\\", "\\\\").replace('"', '\\"').replace("'", "\\'")
    safe_layer = layer_name.replace("\\", "\\\\").replace('"', '\\"')
    safe_ab = artboard_name.replace("\\", "\\\\").replace('"', '\\"')
    text_js = json.dumps({str(k): str(v) for k, v in text_overrides.items()}, ensure_ascii=False)
    img_js = json.dumps(
        {
            str(k): os.path.abspath(p).replace("\\", "/")
            for k, p in image_files.items()
            if p and os.path.exists(p)
        },
        ensure_ascii=False,
    )
    img_raster_js = json.dumps(
        {str(k): bool(v) for k, v in image_raster_flags.items()},
        ensure_ascii=False,
    )
    sr, sg, sb = hex_to_rgb(sim_hex) if sim_hex else (0, 0, 0)
    pr, pg, pb = hex_to_rgb(print_hex) if print_hex else (0, 0, 0)
    do_print_color = "true" if print_hex else "false"
    do_sim_color = "true" if sim_hex else "false"
    do_sim = "false" if skip_simulation else "true"
    do_outline = "true" if outline_text else "false"
    safe_cat = json.dumps(str(category or "A4"))
    shared_tpl = "true" if shared_template else "false"
    group_name = f"P_{prefix}_s{int(slot_id)}"
    safe_group = group_name.replace("\\", "\\\\").replace('"', '\\"')
    cols = min(int(grid_cols), max(1, int(total_slots)))
    layers_only_js = "true" if layers_only else "false"
    sim_only_js = "true" if sim_only else "false"
    center_js = "true" if center_on_artboard else "false"
    safe_tpl_doc = json.dumps(template_doc_name or "")
    keep_tpl_js = "true" if keep_template_open else "false"
    tpl_base = os.path.splitext(os.path.basename(template_path))[0]
    safe_tpl_base = json.dumps(tpl_base)
    snap_js = json.dumps(text_snapshot or {}, ensure_ascii=False)
    jsx = f"""
    #target illustrator
    (function() {{
        var TEMPLATE_PATH = "{safe_tpl}";
        var TEMPLATE_DOC_NAME = {safe_tpl_doc};
        var TEMPLATE_BASE = {safe_tpl_base};
        var KEEP_TEMPLATE_OPEN = {keep_tpl_js};
        var TEXT_SNAPSHOT = {snap_js};
        var PRODUCT_DOC_NAME = "{safe_doc}";
        var LAYER_NAME = "{safe_layer}";
        var ARTBOARD_NAME = "{safe_ab}";
        var PREFIX = "{prefix}";
        var SLOT_IDX = {int(slot_id)};
        var GRID_COLS = {cols};
        var GRID_GAP = {int(VARIABLE_PRINT_GAP_MIN)};
        var LAYERS_ONLY = {layers_only_js};
        var SIM_ONLY = {sim_only_js};
        var CENTER_ON_ARTBOARD = {center_js};
        var TEXT_OVERRIDES = {text_js};
        var IMAGE_FILES = {img_js};
        var IMAGE_RASTER = {img_raster_js};
        var OUTLINE_TEXT = {do_outline};
        var DO_SIMULATION = {do_sim};
        var SIM_R = {sr}; var SIM_G = {sg}; var SIM_B = {sb};
        var PRINT_R = {pr}; var PRINT_G = {pg}; var PRINT_B = {pb};
        var DO_PRINT_COLOR = {do_print_color};
        var DO_SIM_COLOR = {do_sim_color};
        var CATEGORY = {safe_cat};
        var SHARED_TEMPLATE = {shared_tpl};
        var GROUP_NAME = "{safe_group}";

        function findLayerByName(container, name) {{
            if (!container || !container.layers) return null;
            for (var li = 0; li < container.layers.length; li++) {{
                var lyr = container.layers[li];
                if (lyr.name === name) return lyr;
                var sub = findLayerByName(lyr, name);
                if (sub) return sub;
            }}
            return null;
        }}

        function findNamedItem(container, name) {{
            try {{
                if (container.pageItems) {{
                    for (var i = 0; i < container.pageItems.length; i++) {{
                        var it = container.pageItems[i];
                        if (it.name === name) return it;
                        if (it.typename === "GroupItem") {{
                            var r = findNamedItem(it, name);
                            if (r) return r;
                        }}
                    }}
                }}
            }} catch(e) {{}}
            if (container.layers) {{
                for (var j = 0; j < container.layers.length; j++) {{
                    var r2 = findNamedItem(container.layers[j], name);
                    if (r2) return r2;
                }}
            }}
            try {{
                if (container.textFrames) {{
                    var tf = container.textFrames.getByName(name);
                    if (tf) return tf;
                }}
            }} catch(e) {{}}
            return null;
        }}
        function findGroup(container, name) {{
            var direct = findNamedItem(container, name);
            if (direct && direct.typename === "GroupItem") return direct;
            if (container.layers) {{
                for (var i = 0; i < container.layers.length; i++) {{
                    var r = findGroup(container.layers[i], name);
                    if (r) return r;
                }}
            }}
            if (container.pageItems) {{
                for (var j = 0; j < container.pageItems.length; j++) {{
                    var it = container.pageItems[j];
                    if (it.typename === "GroupItem" && it.name === name) return it;
                    if (it.typename === "GroupItem") {{
                        var r2 = findGroup(it, name);
                        if (r2) return r2;
                    }}
                }}
            }}
            return null;
        }}
        function fitItemToPlaceholderBounds(placed, bounds) {{
            var bw = bounds[2] - bounds[0];
            var bh = bounds[1] - bounds[3];
            var nb = placed.visibleBounds;
            var iw = nb[2] - nb[0];
            var ih = nb[1] - nb[3];
            if (iw > 0 && ih > 0) {{
                var sc = Math.min((bw / iw) * 100, (bh / ih) * 100);
                placed.resize(sc, sc);
                nb = placed.visibleBounds;
            }}
            var cx = bounds[0] + bw / 2;
            var cy = bounds[1] - bh / 2;
            var pcx = nb[0] + (nb[2] - nb[0]) / 2;
            var pcy = nb[1] - (nb[1] - nb[3]) / 2;
            placed.translate(cx - pcx, cy - pcy);
        }}
        function getHostLayer(item) {{
            var host = item;
            while (host && host.typename !== "Layer") {{
                host = host.parent;
            }}
            return host;
        }}
        function replaceNamedImage(container, name, filePath, isRaster) {{
            var item = findNamedItem(container, name);
            if (!item || !filePath) return false;
            var bounds = item.visibleBounds;
            var parent = item.parent;
            var hostLayer = parent.typename === "Layer" ? parent : getHostLayer(parent);
            if (!hostLayer) return false;
            var file = new File(filePath);
            if (!file.exists) return false;
            try {{ item.remove(); }} catch(e) {{ return false; }}
            var placed = null;
            if (isRaster) {{
                try {{
                    placed = hostLayer.placedItems.add();
                    placed.file = file;
                    placed.name = name;
                    try {{ placed.embed(); }} catch(e) {{}}
                }} catch(e) {{
                    placed = null;
                }}
            }} else {{
                try {{
                    placed = hostLayer.groupItems.createFromFile(file);
                    placed.name = name;
                    try {{
                        if (placed.pageItems && placed.pageItems.length > 0) {{
                            var last = placed.pageItems[placed.pageItems.length - 1];
                            if (last.typename === "GroupItem" || last.typename === "PathItem") {{
                                last.remove();
                            }}
                        }}
                    }} catch(e) {{}}
                }} catch(e) {{
                    placed = null;
                }}
            }}
            if (!placed) return false;
            var finalItem = findNamedItem(hostLayer, name);
            if (!finalItem && parent.typename !== "Layer") {{
                finalItem = findNamedItem(parent, name);
            }}
            if (finalItem) placed = finalItem;
            if (parent.typename !== "Layer") {{
                try {{ placed.move(parent, ElementPlacement.PLACEATEND); }} catch(e) {{}}
            }}
            try {{ fitItemToPlaceholderBounds(placed, bounds); }} catch(e) {{}}
            return true;
        }}
        function captureCharStyle(charItem) {{
            var ca = charItem.characterAttributes;
            var saved = {{
                font: ca.textFont,
                size: ca.size,
                leading: ca.leading,
                tracking: ca.tracking,
                fill: ca.fillColor,
                stroke: ca.strokeColor,
                stroked: ca.stroked,
                filled: ca.filled,
                hScale: ca.horizontalScale,
                vScale: ca.verticalScale
            }};
            try {{ saved.baselineShift = ca.baselineShift; }} catch(e) {{}}
            try {{ saved.dir = ca.direction; }} catch(e) {{}}
            try {{ saved.fontName = ca.textFont.name; }} catch(e) {{}}
            return saved;
        }}
        function captureTextStyle(tf) {{
            var saved = {{}};
            try {{ saved.kind = tf.kind; }} catch(e) {{}}
            try {{ saved.width = tf.width; }} catch(e) {{}}
            try {{ saved.height = tf.height; }} catch(e) {{}}
            try {{ saved.position = tf.position; }} catch(e) {{}}
            try {{ saved.orientation = tf.orientation; }} catch(e) {{}}
            try {{
                var tr = tf.textRange;
                var ca = tr.characterAttributes;
                var pa = tr.paragraphAttributes;
                saved.font = ca.textFont;
                try {{ saved.fontName = ca.textFont.name; }} catch(e2) {{}}
                saved.size = ca.size;
                saved.leading = ca.leading;
                saved.tracking = ca.tracking;
                saved.fill = ca.fillColor;
                saved.stroke = ca.strokeColor;
                saved.stroked = ca.stroked;
                saved.filled = ca.filled;
                saved.hScale = ca.horizontalScale;
                saved.vScale = ca.verticalScale;
                try {{ saved.baselineShift = ca.baselineShift; }} catch(e2) {{}}
                try {{ saved.dir = ca.direction; }} catch(e2) {{}}
                try {{ saved.justification = pa.justification; }} catch(e2) {{}}
                try {{ saved.autoLeading = pa.autoLeadingAmount; }} catch(e2) {{}}
                try {{ saved.leftIndent = pa.leftIndent; }} catch(e2) {{}}
                try {{ saved.rightIndent = pa.rightIndent; }} catch(e2) {{}}
                try {{ saved.firstLineIndent = pa.firstLineIndent; }} catch(e2) {{}}
                try {{ saved.spaceBefore = pa.spaceBefore; }} catch(e2) {{}}
                try {{ saved.spaceAfter = pa.spaceAfter; }} catch(e2) {{}}
                saved.paragraphs = [];
                var paras = tr.paragraphs;
                for (var pi = 0; pi < paras.length; pi++) {{
                    var pAttr = paras[pi].paragraphAttributes;
                    saved.paragraphs.push({{
                        justification: pAttr.justification,
                        autoLeading: pAttr.autoLeadingAmount,
                        leftIndent: pAttr.leftIndent,
                        rightIndent: pAttr.rightIndent,
                        firstLineIndent: pAttr.firstLineIndent,
                        spaceBefore: pAttr.spaceBefore,
                        spaceAfter: pAttr.spaceAfter
                    }});
                }}
            }} catch(e) {{
                try {{
                    var chars = tf.textRange.characters;
                    if (chars && chars.length > 0) {{
                        var cs = captureCharStyle(chars[0]);
                        for (var sk in cs) {{ if (cs.hasOwnProperty(sk)) saved[sk] = cs[sk]; }}
                    }}
                }} catch(e3) {{}}
            }}
            return saved;
        }}
        function applyParagraphStyleOnly(tf, saved) {{
            if (!tf || !saved) return;
            function applyParaAttrs(pa, src) {{
                if (!pa || !src) return;
                try {{
                    if (src.justification !== undefined) pa.justification = src.justification;
                    if (src.autoLeading !== undefined) pa.autoLeadingAmount = src.autoLeading;
                    if (src.leftIndent !== undefined) pa.leftIndent = src.leftIndent;
                    if (src.rightIndent !== undefined) pa.rightIndent = src.rightIndent;
                    if (src.firstLineIndent !== undefined) pa.firstLineIndent = src.firstLineIndent;
                    if (src.spaceBefore !== undefined) pa.spaceBefore = src.spaceBefore;
                    if (src.spaceAfter !== undefined) pa.spaceAfter = src.spaceAfter;
                }} catch(e) {{}}
            }}
            try {{ applyParaAttrs(tf.textRange.paragraphAttributes, saved); }} catch(e) {{}}
            try {{
                var paras = tf.textRange.paragraphs;
                var paraStyles = saved.paragraphs || [];
                for (var pi = 0; pi < paras.length; pi++) {{
                    var src = paraStyles.length > 0
                        ? paraStyles[Math.min(pi, paraStyles.length - 1)]
                        : saved;
                    applyParaAttrs(paras[pi].paragraphAttributes, src);
                }}
            }} catch(e) {{}}
        }}
        function restoreTextFrameGeometry(tf, saved) {{
            if (!tf || !saved) return;
            try {{
                if (saved.orientation !== undefined && saved.orientation !== null) tf.orientation = saved.orientation;
            }} catch(e) {{}}
            try {{
                if (saved.kind === TextType.AREATEXT) {{
                    if (saved.width > 0) tf.width = saved.width;
                    if (saved.height > 0) tf.height = saved.height;
                }}
                if (saved.position) tf.position = saved.position;
            }} catch(e) {{}}
        }}
        function applyCharStyle(charItem, saved) {{
            if (!charItem || !saved) return;
            var ca = charItem.characterAttributes;
            try {{
                if (saved.font) ca.textFont = saved.font;
                else if (saved.fontName) {{
                    try {{ ca.textFont = app.textFonts.getByName(saved.fontName); }} catch(e2) {{}}
                }}
                if (saved.size) ca.size = saved.size;
                if (saved.leading) ca.leading = saved.leading;
                ca.tracking = saved.tracking;
                if (saved.fill) ca.fillColor = saved.fill;
                if (saved.stroke) ca.strokeColor = saved.stroke;
                ca.stroked = saved.stroked;
                ca.filled = saved.filled;
                ca.horizontalScale = saved.hScale;
                ca.verticalScale = saved.vScale;
                if (saved.baselineShift !== undefined) ca.baselineShift = saved.baselineShift;
                if (saved.dir !== undefined) ca.direction = saved.dir;
            }} catch(e) {{}}
        }}
        function applyTextStyle(tf, saved) {{
            if (!tf || !saved) return;
            applyParagraphStyleOnly(tf, saved);
            try {{
                var chars = tf.textRange.characters;
                for (var ci = 0; ci < chars.length; ci++) applyCharStyle(chars[ci], saved);
            }} catch(e) {{}}
            restoreTextFrameGeometry(tf, saved);
        }}
        function applyStyleToAllChars(tf, saved) {{
            applyTextStyle(tf, saved);
        }}
        function getVisualCenter(item) {{
            var b = item.visibleBounds;
            return {{
                x: (b[0] + b[2]) / 2,
                y: b[1] - (b[1] - b[3]) / 2
            }};
        }}
        function lockVisualCenter(item, cx, cy) {{
            var b = item.visibleBounds;
            var nx = (b[0] + b[2]) / 2;
            var ny = b[1] - (b[1] - b[3]) / 2;
            item.translate(cx - nx, cy - ny);
        }}
        function setTextPreserveStyle(tf, newText) {{
            if (!tf || tf.typename !== "TextFrame") return;
            newText = String(newText);
            if (String(tf.contents) === newText) return;
            tf.locked = false;
            tf.hidden = false;
            var saved = captureTextStyle(tf);
            var centerBefore = getVisualCenter(tf);
            var chars = null;
            try {{ chars = tf.textRange.characters; }} catch(e) {{ chars = null; }}
            var oldLen = chars ? chars.length : 0;
            var newChars = [];
            for (var ni = 0; ni < newText.length; ni++) newChars.push(newText.charAt(ni));
            var newLen = newChars.length;
            if (oldLen > 0) {{
                var styleSource = saved;
                try {{
                    var cs0 = captureCharStyle(chars[0]);
                    for (var sk in cs0) {{ if (cs0.hasOwnProperty(sk)) styleSource[sk] = cs0[sk]; }}
                }} catch(e) {{}}
                if (newLen <= oldLen) {{
                    for (var i = 0; i < newLen; i++) {{
                        chars[i].contents = newChars[i];
                        applyCharStyle(chars[i], styleSource);
                    }}
                    for (var d = oldLen - 1; d >= newLen; d--) {{
                        chars[d].remove();
                    }}
                }} else {{
                    for (var j = 0; j < oldLen; j++) {{
                        chars[j].contents = newChars[j];
                        applyCharStyle(chars[j], styleSource);
                    }}
                    var anchor = tf.textRange.characters[oldLen - 1];
                    for (var e = oldLen; e < newLen; e++) {{
                        var nc = anchor.duplicate();
                        nc.contents = newChars[e];
                        applyCharStyle(nc, styleSource);
                        anchor = nc;
                    }}
                }}
            }} else {{
                try {{ tf.textRange.contents = newText; }} catch(err) {{ tf.contents = newText; }}
                try {{
                    var allChars = tf.textRange.characters;
                    for (var ac = 0; ac < allChars.length; ac++) applyCharStyle(allChars[ac], saved);
                }} catch(e) {{}}
            }}
            applyParagraphStyleOnly(tf, saved);
            try {{ tf.textRange.paragraphAttributes.justification = Justification.CENTER; }} catch(e) {{}}
            try {{
                if (saved.kind === TextType.AREATEXT) {{
                    if (saved.width > 0) tf.width = saved.width;
                    if (saved.height > 0) tf.height = saved.height;
                }}
                if (saved.orientation !== undefined && saved.orientation !== null) tf.orientation = saved.orientation;
            }} catch(e) {{}}
            lockVisualCenter(tf, centerBefore.x, centerBefore.y);
        }}
        function forEachTextFrame(root, fn) {{
            if (!root) return;
            try {{
                if (root.textFrames) {{
                    for (var t = 0; t < root.textFrames.length; t++) fn(root.textFrames[t]);
                }}
            }} catch(e) {{}}
            try {{
                if (root.pageItems) {{
                    for (var i = 0; i < root.pageItems.length; i++) {{
                        try {{
                            if (root.pageItems[i].typename === "GroupItem") forEachTextFrame(root.pageItems[i], fn);
                        }} catch(e2) {{}}
                    }}
                }}
            }} catch(e) {{}}
            try {{
                if (root.layers) {{
                    for (var j = 0; j < root.layers.length; j++) forEachTextFrame(root.layers[j], fn);
                }}
            }} catch(e) {{}}
        }}
        function isActiveOverrideKey(name, overrides) {{
            if (overrides[name]) return true;
            var upper = String(name).toUpperCase();
            for (var k in overrides) {{
                if (!overrides.hasOwnProperty(k)) continue;
                if (String(k).toUpperCase() === upper) return true;
                var cands = buildTextKeyCandidates(k);
                for (var ci = 0; ci < cands.length; ci++) {{
                    if (String(cands[ci]).toUpperCase() === upper) return true;
                }}
            }}
            return false;
        }}
        function isVariableTextFrameName(nm) {{
            if (!nm) return false;
            if (/^(TEXT|NUM|NUMBER)[_]?\\d+$/i.test(nm)) return true;
            if (/^TEXT_(NAME|NUMBER|NUM\\d+)$/i.test(nm)) return true;
            return false;
        }}
        function buildTextKeyCandidates(key) {{
            var want = String(key).toUpperCase();
            var candidates = [want, String(key)];
            var textMatch = want.match(/^TEXT[_]?(\\d+)$/);
            if (textMatch) {{
                var n = textMatch[1];
                candidates.push("TEXT" + n);
                candidates.push("TEXT_" + n);
            }}
            var numMatch = want.match(/^(NUM|NUMBER)[_]?(\d+)$/);
            if (numMatch) {{
                var num = numMatch[2];
                candidates.push("NUM" + num);
                candidates.push("NUM_" + num);
                candidates.push("NUMBER" + num);
                candidates.push("NUMBER_" + num);
            }}
            if (want === "TEXT_NAME" || want === "TEXTNAME") candidates.push("TEXT1");
            if (want === "TEXT_NUMBER" || want === "TEXTNUMBER") candidates.push("TEXT2");
            var unique = [];
            for (var ui = 0; ui < candidates.length; ui++) {{
                var c = candidates[ui];
                var seen = false;
                for (var uj = 0; uj < unique.length; uj++) {{
                    if (unique[uj] === c) {{ seen = true; break; }}
                }}
                if (!seen) unique.push(c);
            }}
            return unique;
        }}
        function findTextFrameForOverride(doc, key) {{
            var direct = findNamedItem(doc, key);
            if (direct && direct.typename === "TextFrame") return direct;
            var candidates = buildTextKeyCandidates(key);
            var found = null;
            forEachTextFrame(doc, function(tf) {{
                if (found) return;
                var nm = tf.name || "";
                for (var ci = 0; ci < candidates.length; ci++) {{
                    if (nm === candidates[ci]) {{ found = tf; return; }}
                    if (nm.toUpperCase() === String(candidates[ci]).toUpperCase()) {{ found = tf; return; }}
                }}
            }});
            return found;
        }}
        function removeUnusedNamedTexts(doc, overrides) {{
            var toRemove = [];
            forEachTextFrame(doc, function(tf) {{
                var nm = tf.name;
                if (!isVariableTextFrameName(nm)) return;
                if (!isActiveOverrideKey(nm, overrides)) toRemove.push(tf);
            }});
            for (var ri = 0; ri < toRemove.length; ri++) {{
                try {{ toRemove[ri].remove(); }} catch(e) {{}}
            }}
        }}
        function clearLayerPrintContent(layer) {{
            if (!layer) return;
            layer.locked = false;
            for (var i = layer.pageItems.length - 1; i >= 0; i--) {{
                var it = layer.pageItems[i];
                var n = it.name || "";
                if (n.indexOf("_Box_") !== -1) continue;
                try {{ it.remove(); }} catch(e) {{}}
            }}
        }}
        function ensureSinglePrintOnLayer(layer, keepName) {{
            if (!layer || !layer.pageItems) return;
            var nonBox = [];
            for (var i = 0; i < layer.pageItems.length; i++) {{
                try {{
                    var it = layer.pageItems[i];
                    if (!it) continue;
                    if ((it.name || "").indexOf("_Box_") !== -1) continue;
                    nonBox.push(it);
                }} catch(e) {{}}
            }}
            if (nonBox.length <= 1) {{
                if (nonBox.length === 1) nonBox[0].name = keepName;
                return;
            }}
            var keeper = null;
            for (var k = 0; k < nonBox.length; k++) {{
                if (nonBox[k].name === keepName) {{ keeper = nonBox[k]; break; }}
            }}
            if (!keeper) keeper = nonBox[nonBox.length - 1];
            for (var j = 0; j < nonBox.length; j++) {{
                if (nonBox[j] !== keeper) {{
                    try {{ nonBox[j].remove(); }} catch(e) {{}}
                }}
            }}
            keeper.name = keepName;
        }}
        function makePrintColor(r, g, b) {{
            try {{
                if (app.activeDocument.documentColorSpace == DocumentColorSpace.CMYK) {{
                    var cmykArr = app.convertSampleColor(
                        ImageColorSpace.RGB, ColorModel.PROCESS, [r, g, b],
                        ImageColorSpace.CMYK, ColorModel.PROCESS, []
                    );
                    var cm = new CMYKColor();
                    cm.cyan = cmykArr[0];
                    cm.magenta = cmykArr[1];
                    cm.yellow = cmykArr[2];
                    cm.black = cmykArr[3];
                    return cm;
                }}
            }} catch(e) {{}}
            var rgb = new RGBColor();
            rgb.red = r; rgb.green = g; rgb.blue = b;
            return rgb;
        }}
        function colRecPrint(it, c) {{
            try {{
                if (it.typename === "GroupItem") {{
                    for (var ci = 0; ci < it.pageItems.length; ci++) colRecPrint(it.pageItems[ci], c);
                }} else if (it.typename === "PathItem" && !it.clipping) {{
                    if (it.stroked && !it.filled) {{ it.strokeColor = c; }}
                    it.filled = true; it.fillColor = c; it.stroked = false;
                }} else if (it.typename === "CompoundPathItem") {{
                    for (var cj = 0; cj < it.pathItems.length; cj++) {{
                        if (!it.pathItems[cj].clipping) {{
                            it.pathItems[cj].filled = true;
                            it.pathItems[cj].fillColor = c;
                            it.pathItems[cj].stroked = false;
                        }}
                    }}
                }} else if (it.typename === "TextFrame") {{
                    try {{
                        var chars = it.textRange.characters;
                        for (var ti = 0; ti < chars.length; ti++) {{
                            chars[ti].characterAttributes.fillColor = c;
                            chars[ti].characterAttributes.filled = true;
                            chars[ti].characterAttributes.stroked = false;
                        }}
                    }} catch(e) {{}}
                }}
            }} catch(e) {{}}
        }}
        function recolorDeep(item, c) {{
            colRecPrint(item, c);
        }}
        function getSimBoxSuffix(category, itemW, itemH) {{
            var suffix = "A4_Square";
            var catLower = (category || "A4").toLowerCase();
            if (category === "Sleeve2") suffix = "Sleeve2";
            else if (catLower.indexOf("sleeve") !== -1 || catLower.indexOf("9") !== -1 || catLower.indexOf("\\u05e9\\u05e8\\u05d5\\u05d5\\u05dc") !== -1) suffix = "Sleeve";
            else if (category === "Pocket") suffix = "Pocket";
            else if (category === "2Pocket") suffix = "2Pocket";
            else if (category === "2") suffix = "2";
            else if (category === "A3") suffix = "A3";
            else if (category === "A5") suffix = "A5";
            else if (category === "A4") {{
                var ratio = itemW / itemH;
                if (ratio > 1.21) suffix = "A4_Landscape";
                else if (ratio < 0.75) suffix = "A4_Portrait";
                else suffix = "A4_Square";
            }}
            return suffix;
        }}
        function placeSimExactFromTemplate(simItem, layoutOff, tplW, tplH, prefix, category) {{
            var suffix = getSimBoxSuffix(category, tplW, tplH);
            var boxName = "S" + prefix + "_Box_" + suffix;
            var box = findNamedItem(prodDoc, boxName);
            if (!box) {{
                try {{ box = prodDoc.pageItems.getByName(boxName); }} catch(e) {{ return; }}
            }}
            if (!box) return;
            var b = box.visibleBounds;
            var boxW = b[2] - b[0];
            var boxH = b[1] - b[3];
            var scale = 1.0;
            if (tplW > 0 && tplH > 0) {{
                scale = Math.min(boxW / tplW, boxH / tplH);
            }}
            if (Math.abs(scale - 1.0) > 0.001) {{
                try {{
                    simItem.resize(scale * 100, scale * 100, true, true, true, true, scale * 100, Transformation.TOPLEFT);
                }} catch(e) {{
                    simItem.resize(scale * 100, scale * 100);
                }}
            }}
            var targetLeft = b[0] + layoutOff.relL * scale;
            var targetTop = b[1] - layoutOff.relT * scale;
            var sb = simItem.visibleBounds;
            simItem.translate(targetLeft - sb[0], targetTop - sb[1]);
        }}
        function outlineAllNamedTexts(doc) {{
            var frames = [];
            forEachTextFrame(doc, function(tf) {{
                var nm = tf.name;
                if (!isVariableTextFrameName(nm)) return;
                frames.push(tf);
            }});
            for (var oi = 0; oi < frames.length; oi++) {{
                try {{ frames[oi].createOutline(); }} catch(e) {{}}
            }}
        }}
        function outlineTextKeys(container, keys) {{
            for (var ki = 0; ki < keys.length; ki++) {{
                var tf = findNamedItem(container, keys[ki]);
                if (!tf || tf.typename !== "TextFrame") {{
                    tf = findTextFrameForOverride(container, keys[ki]);
                }}
                if (tf && tf.typename === "TextFrame") {{
                    try {{ tf.createOutline(); }} catch(e) {{}}
                }}
            }}
        }}
        function recolorItem(it, c) {{
            recolorDeep(it, c);
        }}
        function alignCenterToArtboard(item, abRect) {{
            var b = item.visibleBounds;
            var icx = b[0] + (b[2] - b[0]) / 2;
            var icy = b[1] - (b[1] - b[3]) / 2;
            var abCx = (abRect[0] + abRect[2]) / 2;
            var abCy = abRect[1] - (abRect[1] - abRect[3]) / 2;
            item.translate(abCx - icx, abCy - icy);
        }}
        function fitAndCenterOnArtboard(item, abRect, marginPt) {{
            var margin = marginPt || 24;
            var abW = (abRect[2] - abRect[0]) - margin * 2;
            var abH = (abRect[1] - abRect[3]) - margin * 2;
            if (abW < 1 || abH < 1) {{
                alignCenterToArtboard(item, abRect);
                return;
            }}
            var b = item.visibleBounds;
            var iw = b[2] - b[0];
            var ih = b[1] - b[3];
            if (iw < 1 || ih < 1) return;
            var scale = Math.min(abW / iw, abH / ih);
            if (scale > 1) scale = 1;
            if (Math.abs(scale - 1) > 0.001) {{
                try {{
                    item.resize(scale * 100, scale * 100, true, true, true, true, scale * 100, Transformation.CENTER);
                }} catch(eR) {{
                    item.resize(scale * 100, scale * 100);
                }}
            }}
            alignCenterToArtboard(item, abRect);
        }}
        function captureArtboardOffset(item, abRect) {{
            var b = item.visibleBounds;
            return {{
                relL: b[0] - abRect[0],
                relT: abRect[1] - b[1]
            }};
        }}
        function placeAtArtboardOffset(item, abRect, relL, relT) {{
            var b = item.visibleBounds;
            var targetLeft = abRect[0] + relL;
            var targetTop = abRect[1] - relT;
            item.translate(targetLeft - b[0], targetTop - b[1]);
        }}
        function resizeArtboardTopLeft(ab, w, h) {{
            var r = ab.artboardRect;
            ab.artboardRect = [r[0], r[1], r[0] + w, r[1] - h];
            return ab.artboardRect;
        }}
        function itemIntersectsArtboard(item, abRect) {{
            try {{
                var b = item.visibleBounds;
                if (b[2] <= abRect[0] || b[0] >= abRect[2]) return false;
                if (b[1] <= abRect[3] || b[3] >= abRect[1]) return false;
                return true;
            }} catch(e) {{ return false; }}
        }}
        function collectTopItemsOnArtboard(container, abRect, out) {{
            if (container.pageItems) {{
                for (var i = 0; i < container.pageItems.length; i++) {{
                    var it = container.pageItems[i];
                    if (it.hidden) continue;
                    if (itemIntersectsArtboard(it, abRect)) out.push(it);
                }}
            }}
            if (container.layers) {{
                for (var j = 0; j < container.layers.length; j++) {{
                    var sub = container.layers[j];
                    if (!sub.visible) continue;
                    sub.locked = false;
                    collectTopItemsOnArtboard(sub, abRect, out);
                }}
            }}
        }}
        function resolveCopyRoot(doc, abRect) {{
            var named = findGroup(doc, GROUP_NAME);
            if (named) return named;
            var varPrint = findGroup(doc, "VAR_PRINT");
            if (varPrint) return varPrint;
            var items = [];
            for (var li = 0; li < doc.layers.length; li++) {{
                var layer = doc.layers[li];
                if (!layer.visible) continue;
                layer.locked = false;
                collectTopItemsOnArtboard(layer, abRect, items);
            }}
            if (items.length === 0) return null;
            if (items.length === 1) return items[0];
            var hostLayer = items[0].layer;
            hostLayer.locked = false;
            var grp = hostLayer.groupItems.add();
            grp.name = GROUP_NAME;
            for (var ii = 0; ii < items.length; ii++) {{
                items[ii].moveToEnd(grp);
            }}
            return grp;
        }}
        function findProductDoc(name) {{
            for (var i = 0; i < app.documents.length; i++) {{
                if (app.documents[i].name === name) return app.documents[i];
            }}
            return null;
        }}
        function closeStrayTemplateDocs() {{
            for (var i = app.documents.length - 1; i >= 0; i--) {{
                try {{
                    var dn = app.documents[i].name;
                    if (dn === PRODUCT_DOC_NAME) continue;
                    if (TEMPLATE_DOC_NAME && dn === TEMPLATE_DOC_NAME) continue;
                    if (dn.indexOf(TEMPLATE_BASE) === 0) {{
                        app.documents[i].close(SaveOptions.DONOTSAVECHANGES);
                    }}
                }} catch(e) {{}}
            }}
        }}
        function restoreTextSnapshot(doc) {{
            for (var k in TEXT_SNAPSHOT) {{
                if (!TEXT_SNAPSHOT.hasOwnProperty(k)) continue;
                var tf = findTextFrameForOverride(doc, k);
                if (tf) {{
                    try {{ setTextPreserveStyle(tf, TEXT_SNAPSHOT[k]); }} catch(e) {{
                        try {{ tf.contents = TEXT_SNAPSHOT[k]; }} catch(e2) {{}}
                    }}
                }}
            }}
        }}

        var prodDoc = findProductDoc(PRODUCT_DOC_NAME);
        if (!prodDoc) {{
            for (var di = 0; di < app.documents.length; di++) {{
                var dn = app.documents[di].name;
                if (dn === PRODUCT_DOC_NAME || dn.indexOf(PRODUCT_DOC_NAME.replace(".ai", "")) !== -1) {{
                    prodDoc = app.documents[di];
                    break;
                }}
            }}
        }}
        if (!prodDoc) return -1;

        app.activeDocument = prodDoc;
        var tplDoc = null;
        if (TEMPLATE_DOC_NAME) {{
            tplDoc = findProductDoc(TEMPLATE_DOC_NAME);
            if (!tplDoc) return -7;
        }} else {{
            closeStrayTemplateDocs();
            tplDoc = app.open(new File(TEMPLATE_PATH));
        }}
        if (!tplDoc) return -7;
        app.activeDocument = tplDoc;
        app.executeMenuCommand("unlockAll");

        var tplAbIdx = tplDoc.artboards.getActiveArtboardIndex();
        var tplAb = tplDoc.artboards[tplAbIdx].artboardRect;
        var tplW = tplAb[2] - tplAb[0];
        var tplH = tplAb[1] - tplAb[3];

        restoreTextSnapshot(tplDoc);

        var textKeys = [];
        var hasOverrides = false;
        for (var tk in TEXT_OVERRIDES) {{
            if (!TEXT_OVERRIDES.hasOwnProperty(tk)) continue;
            hasOverrides = true;
            var tf = findTextFrameForOverride(tplDoc, tk);
            if (tf && tf.typename === "TextFrame") {{
                setTextPreserveStyle(tf, TEXT_OVERRIDES[tk]);
                textKeys.push(tf.name || tk);
            }}
        }}
        if (SHARED_TEMPLATE && hasOverrides) {{
            removeUnusedNamedTexts(tplDoc, TEXT_OVERRIDES);
        }}
        for (var ik in IMAGE_FILES) {{
            if (IMAGE_FILES.hasOwnProperty(ik)) {{
                var asRaster = IMAGE_RASTER[ik] === true;
                replaceNamedImage(tplDoc, ik, IMAGE_FILES[ik], asRaster);
            }}
        }}
        if (OUTLINE_TEXT) {{
            if (textKeys.length > 0) {{
                outlineTextKeys(tplDoc, textKeys);
            }} else if (!hasOverrides) {{
                outlineAllNamedTexts(tplDoc);
            }}
        }}

        var copyRoot = resolveCopyRoot(tplDoc, tplAb);
        if (!copyRoot) {{
            tplDoc.close(SaveOptions.DONOTSAVECHANGES);
            return -2;
        }}
        copyRoot.name = GROUP_NAME;
        var layoutOff = captureArtboardOffset(copyRoot, tplAb);

        if (SIM_ONLY !== true && DO_PRINT_COLOR) {{
            var pc = makePrintColor(PRINT_R, PRINT_G, PRINT_B);
            colRecPrint(copyRoot, pc);
        }}

        tplDoc.selection = null;
        copyRoot.selected = true;
        app.executeMenuCommand("copy");
        if (!KEEP_TEMPLATE_OPEN) {{
            tplDoc.close(SaveOptions.DONOTSAVECHANGES);
        }}

        prodDoc = findProductDoc(PRODUCT_DOC_NAME);
        if (!prodDoc) {{
            for (var di = 0; di < app.documents.length; di++) {{
                var dn = app.documents[di].name;
                if (dn === PRODUCT_DOC_NAME || dn.indexOf(PRODUCT_DOC_NAME.replace(".ai", "")) !== -1) {{
                    prodDoc = app.documents[di];
                    break;
                }}
            }}
        }}
        if (!prodDoc && app.documents.length > 0) prodDoc = app.documents[0];
        if (!prodDoc) return -1;

        app.activeDocument = prodDoc;
        app.executeMenuCommand("unlockAll");

        if (SIM_ONLY) {{
            try {{
                var simLayerOnly = prodDoc.layers.getByName("Simulation");
                var sideNameOnly = "Back";
                if (PREFIX === "F") sideNameOnly = "Front";
                else if (PREFIX === "RS") sideNameOnly = "Right_Sleeve";
                else if (PREFIX === "LS") sideNameOnly = "Left_Sleeve";
                var simTargetOnly = simLayerOnly.layers.getByName("S_Placement_" + sideNameOnly);
                simLayerOnly.visible = true;
                simTargetOnly.visible = true;
                simTargetOnly.locked = false;
                prodDoc.activeLayer = simTargetOnly;
                prodDoc.selection = null;
                app.executeMenuCommand("pasteInPlace");
                var simOnlyItem = null;
                if (prodDoc.selection && prodDoc.selection.length > 0) {{
                    simOnlyItem = prodDoc.selection[0];
                }}
                if (!simOnlyItem) return -4;
                simOnlyItem.hidden = false;
                placeSimExactFromTemplate(simOnlyItem, layoutOff, tplW, tplH, PREFIX, CATEGORY);
                if (DO_SIM_COLOR) {{
                    var scOnly = makePrintColor(SIM_R, SIM_G, SIM_B);
                    colRecPrint(simOnlyItem, scOnly);
                }}
                simOnlyItem.name = "";
                var simB = simOnlyItem.visibleBounds;
                return simB[2] - simB[0];
            }} catch(eSim) {{
                return -4;
            }}
        }}

        var targetLayer = findLayerByName(app.activeDocument, LAYER_NAME);
        if (!targetLayer) return -3;
        targetLayer.locked = false;
        targetLayer.visible = true;
        prodDoc.activeLayer = targetLayer;
        prodDoc.selection = null;
        clearLayerPrintContent(targetLayer);
        prodDoc.activeLayer = targetLayer;
        app.executeMenuCommand("pasteInPlace");

        var placed = null;
        if (prodDoc.selection && prodDoc.selection.length > 0) {{
            placed = prodDoc.selection[0];
        }}
        if (!placed) {{
            for (var pi = targetLayer.pageItems.length - 1; pi >= 0; pi--) {{
                var cand = targetLayer.pageItems[pi];
                var cn = cand.name || "";
                if (cn.indexOf("_Box_") !== -1) continue;
                placed = cand;
                break;
            }}
        }}
        if (!placed) return -4;

        try {{
            placed.name = GROUP_NAME;
            placed.locked = false;
            placed.hidden = false;
        }} catch(eName) {{
            placed = null;
            for (var pi2 = targetLayer.pageItems.length - 1; pi2 >= 0; pi2--) {{
                try {{
                    var cand2 = targetLayer.pageItems[pi2];
                    var cn2 = cand2.name || "";
                    if (cn2.indexOf("_Box_") !== -1) continue;
                    placed = cand2;
                    break;
                }} catch(e) {{}}
            }}
            if (!placed) return -4;
            try {{ placed.name = GROUP_NAME; }} catch(e2) {{ return -4; }}
        }}
        try {{
            if (placed.parent !== targetLayer) placed.moveToEnd(targetLayer);
        }} catch(e) {{}}
        ensureSinglePrintOnLayer(targetLayer, GROUP_NAME);

        function slotRectFromSlot1(slot1Rect, w, h, gapX, gapY, columns, index1Based) {{
            var idx = index1Based - 1;
            var col = idx % columns;
            var row = Math.floor(idx / columns);
            var left = slot1Rect[0] + col * (w + gapX);
            var top = slot1Rect[1] - row * (h + gapY);
            return [left, top, left + w, top - h];
        }}

        var targetAb = null;
        var abIdx = -1;
        for (var ai = 0; ai < prodDoc.artboards.length; ai++) {{
            if (prodDoc.artboards[ai].name === ARTBOARD_NAME) {{
                targetAb = prodDoc.artboards[ai];
                abIdx = ai;
                break;
            }}
        }}
        if (!targetAb) {{
            var abFallback = ARTBOARD_NAME.replace(/_\\d+$/, "") + "_1";
            try {{ targetAb = prodDoc.artboards.getByName(abFallback); }} catch(e) {{}}
        }}
        if (!targetAb) return -5;

        if (LAYERS_ONLY) {{
            var boxSuffix = getSimBoxSuffix(CATEGORY, tplW, tplH);
            var boxName = PREFIX + "_Box_" + boxSuffix;
            var box = findNamedItem(targetLayer, boxName);
            if (!box) {{
                try {{ box = targetLayer.pageItems.getByName(boxName); }} catch(e) {{ box = null; }}
            }}
            if (!box) return -6;
            var bb = box.visibleBounds;
            var boxRect = [bb[0], bb[1], bb[2], bb[3]];
            placeAtArtboardOffset(placed, boxRect, layoutOff.relL, layoutOff.relT);
        }} else {{
            var placeRect;
            if (abIdx >= 0) prodDoc.artboards.setActiveArtboardIndex(abIdx);
            placeRect = targetAb.artboardRect;
            if (CENTER_ON_ARTBOARD) {{
                fitAndCenterOnArtboard(placed, placeRect, 24);
            }} else {{
                placeRect = resizeArtboardTopLeft(targetAb, tplW, tplH);
                placeAtArtboardOffset(placed, placeRect, layoutOff.relL, layoutOff.relT);
            }}
        }}

        if (DO_PRINT_COLOR) {{
            var pc2 = makePrintColor(PRINT_R, PRINT_G, PRINT_B);
            colRecPrint(placed, pc2);
        }}

        if (DO_SIMULATION) {{
            try {{
                var simLayer = prodDoc.layers.getByName("Simulation");
                var sideName = "Back";
                if (PREFIX === "F") sideName = "Front";
                else if (PREFIX === "RS") sideName = "Right_Sleeve";
                else if (PREFIX === "LS") sideName = "Left_Sleeve";
                var simTarget = simLayer.layers.getByName("S_Placement_" + sideName);
                simLayer.visible = true;
                simTarget.visible = true;
                var simCopy = placed.duplicate(simTarget, ElementPlacement.PLACEATEND);
                simCopy.hidden = false;
                placeSimExactFromTemplate(simCopy, layoutOff, tplW, tplH, PREFIX, CATEGORY);
                if (DO_SIM_COLOR) {{
                    var c = makePrintColor(SIM_R, SIM_G, SIM_B);
                    colRecPrint(simCopy, c);
                }}
                simCopy.name = "";
            }} catch(e) {{}}
        }}

        var finalB = placed.visibleBounds;
        return finalB[2] - finalB[0];
    }})();
    """
    try:
        if product_doc is not None:
            try:
                product_doc.Activate()
            except Exception:
                pass
        activate_document(app, product_doc_name)
        val = None
        for attempt in range(3):
            res = app.DoJavaScript(jsx)
            if res is None or res == "":
                print("   > Template placement: empty JSX result")
                return 0.0
            val = float(res)
            if val >= 0 or val != -4.0:
                break
            time.sleep(0.4)
        if val is None:
            return 0.0
        if val < 0:
            err_map = {
                -1: "product document not found",
                -2: "no printable content on template artboard",
                -3: f"print layer not found: {layer_name}",
                -4: "paste from template failed",
                -5: f"artboard not found: {artboard_name}",
                -6: f"placement box not found on {layer_name}",
                -7: "template session lost — reopen Illustrator if this repeats",
            }
            print(f"   > Template placement failed: {err_map.get(int(val), val)}")
            return 0.0
        if print_hex:
            r, g, b = hex_to_rgb(print_hex)
            sc = JSX_RECOLOR_GROUP.replace('%GNAME%', group_name)
            sc = sc.replace('%R%', str(r)).replace('%G%', str(g)).replace('%B%', str(b))
            run_jsx(app, sc)
            time.sleep(0.2)
        return val
    except Exception as e:
        print(f"   > Template placement error: {e}")
        return 0.0


def refresh_template_side_colors(
    app,
    side: str,
    variants: list,
    loc: dict,
    h1: str,
    resolve_colors_fn,
) -> None:
    """מחיל צבע מחדש אחרי outline — לשכבות Print ו-Simulation."""
    prefix = loc.get("prefix") or {"front": "F", "back": "B", "right_sleeve": "RS", "left_sleeve": "LS"}.get(side, "F")
    side_map = {"F": "Front", "B": "Back", "RS": "Right_Sleeve", "LS": "Left_Sleeve"}
    sim_side = side_map.get(prefix, "Front")

    for vi, variant in enumerate(variants):
        slot_idx = vi + 1
        layer_name, _ = variable_layer_and_artboard(side, slot_idx)
        sim_hex, print_hex = resolve_colors_fn(variant, loc, h1)
        if sim_hex is None and print_hex is None:
            continue
        group_name = f"P_{prefix}_s{slot_idx}"
        if print_hex:
            r, g, b = hex_to_rgb(print_hex)
            safe_layer = layer_name.replace("\\", "\\\\").replace('"', '\\"')
            safe_group = group_name.replace("\\", "\\\\").replace('"', '\\"')
            jsx_print = f"""
            #target illustrator
            (function() {{
                function makePrintColor(r, g, b) {{
                    try {{
                        if (app.activeDocument.documentColorSpace == DocumentColorSpace.CMYK) {{
                            var cmykArr = app.convertSampleColor(
                                ImageColorSpace.RGB, ColorModel.PROCESS, [r, g, b],
                                ImageColorSpace.CMYK, ColorModel.PROCESS, []
                            );
                            var cm = new CMYKColor();
                            cm.cyan = cmykArr[0]; cm.magenta = cmykArr[1];
                            cm.yellow = cmykArr[2]; cm.black = cmykArr[3];
                            return cm;
                        }}
                    }} catch(e) {{}}
                    var rgb = new RGBColor();
                    rgb.red = r; rgb.green = g; rgb.blue = b;
                    return rgb;
                }}
                function colRecPrint(it, c) {{
                    try {{
                        if (it.typename === "GroupItem") {{
                            for (var ci = 0; ci < it.pageItems.length; ci++) colRecPrint(it.pageItems[ci], c);
                        }} else if (it.typename === "PathItem" && !it.clipping) {{
                            if (it.stroked && !it.filled) {{ it.strokeColor = c; }}
                            it.filled = true; it.fillColor = c; it.stroked = false;
                        }} else if (it.typename === "CompoundPathItem") {{
                            for (var cj = 0; cj < it.pathItems.length; cj++) {{
                                if (!it.pathItems[cj].clipping) {{
                                    it.pathItems[cj].filled = true;
                                    it.pathItems[cj].fillColor = c;
                                    it.pathItems[cj].stroked = false;
                                }}
                            }}
                        }} else if (it.typename === "TextFrame") {{
                            try {{
                                var chars = it.textRange.characters;
                                for (var ti = 0; ti < chars.length; ti++) {{
                                    chars[ti].characterAttributes.fillColor = c;
                                    chars[ti].characterAttributes.filled = true;
                                    chars[ti].characterAttributes.stroked = false;
                                }}
                            }} catch(e) {{}}
                        }}
                    }} catch(e) {{}}
                }}
                var doc = app.activeDocument;
                var layer = null;
                try {{ layer = doc.layers.getByName("{safe_layer}"); }} catch(e) {{ return; }}
                if (!layer) return;
                layer.locked = false;
                var grp = null;
                try {{ grp = layer.pageItems.getByName("{safe_group}"); }} catch(e) {{ return; }}
                if (grp) colRecPrint(grp, makePrintColor({r}, {g}, {b}));
            }})();
            """
            run_jsx(app, jsx_print)

        if vi == 0 and sim_hex:
            sr, sg, sb = hex_to_rgb(sim_hex)
            jsx_sim = f"""
            #target illustrator
            (function() {{
                function makePrintColor(r, g, b) {{
                    var rgb = new RGBColor();
                    rgb.red = r; rgb.green = g; rgb.blue = b;
                    return rgb;
                }}
                function colRecPrint(it, c) {{
                    try {{
                        if (it.typename === "GroupItem") {{
                            for (var ci = 0; ci < it.pageItems.length; ci++) colRecPrint(it.pageItems[ci], c);
                        }} else if (it.typename === "PathItem" && !it.clipping) {{
                            it.filled = true; it.fillColor = c; it.stroked = false;
                        }} else if (it.typename === "CompoundPathItem") {{
                            for (var cj = 0; cj < it.pathItems.length; cj++) {{
                                if (!it.pathItems[cj].clipping) {{
                                    it.pathItems[cj].filled = true;
                                    it.pathItems[cj].fillColor = c;
                                    it.pathItems[cj].stroked = false;
                                }}
                            }}
                        }}
                    }} catch(e) {{}}
                }}
                var doc = app.activeDocument;
                var simLayer = doc.layers.getByName("Simulation");
                simLayer.visible = true;
                var target = simLayer.layers.getByName("S_Placement_{sim_side}");
                target.visible = true;
                target.locked = false;
                var c = makePrintColor({sr}, {sg}, {sb});
                for (var i = 0; i < target.pageItems.length; i++) {{
                    var it = target.pageItems[i];
                    var nm = it.name || "";
                    if (nm.indexOf("_Box_") !== -1) continue;
                    colRecPrint(it, c);
                }}
            }})();
            """
            run_jsx(app, jsx_sim)


def place_and_simulate_print(
    doc,
    app,
    path,
    pre,
    cat,
    p_hex,
    s_hex,
    is_raster=False,
    layer_name=None,
    artboard_name=None,
    skip_simulation=False,
    should_update_size_label=True,
):
    print(f"--- Processing {pre} ---")
    l_map = {"F":"Print_Front","B":"Print_Back","RS":"Print_Right_Sleeve","LS":"Print_Left_Sleeve"}
    # וידוא מסמך
    doc = get_doc_safe(app)
    if not doc: return 0
    target_layer = layer_name or l_map[pre]
    target_ab = artboard_name if artboard_name is not None else am.get(pre, "")
    p_lay = get_layer(doc, target_layer)
    if not p_lay: return 0
    unique_name_print = f"P_{pre}_{uuid.uuid4().hex[:6]}"
    # משתנה זמני לבדיקה שההטמעה הצליחה
    initial_check_w = 0
    try:
        # בדיקה אם הקובץ הוא SVG אך מסומן כ"ללא וקטור" (Raster)
        is_svg_no_vector = is_raster and path.lower().endswith('.svg')
        if is_raster or is_svg_no_vector:
            # --- הטמעת רסטר (תמונה או SVG "ללא וקטור") ---
            safe_path = path.replace('\\', '\\\\')
            jsx_place_raster = f"""
            #target illustrator
            function placeRaster(filePath, layerName, itemName) {{
                try {{
                    var doc = app.activeDocument;
                    var layer = doc.layers.getByName(layerName);
                    var file = new File("{safe_path}");
                    var placedItem = layer.placedItems.add();
                    placedItem.file = file;
                    placedItem.name = itemName;
                    try {{ placedItem.embed(); }} catch(e) {{}}
                    return placedItem.width;
                }} catch(e) {{ return 0; }}
            }}
            placeRaster('{safe_path}', '{target_layer}', '{unique_name_print}');
            """
            raw_width = app.DoJavaScript(jsx_place_raster)
            initial_check_w = float(raw_width)
        else:
            # --- הטמעת וקטור רגילה (פתיחה ופירוק) ---
            imported_group = p_lay.GroupItems.CreateFromFile(path)
            clean_arts(imported_group)
            imported_group.Name = unique_name_print
            initial_check_w = imported_group.Width
    except Exception as e:
        print(f"Fatal Import Error: {e}")
        return 0
    if initial_check_w == 0: return 0
    # 1. ניקוי וצביעה
    r, g, b = (0,0,0)
    do_col = 'false'
    if p_hex:
        r, g, b = hex_to_rgb(p_hex)
        do_col = 'true'
    # מעבירים לסקריפט הניקוי האם זה רסטר. אם כן - הוא מדלג על הניקוי.
    is_raster_str = "true" if (is_raster or is_svg_no_vector) else "false"
    # מריצים את הניקוי (או מדלגים אם זה רסטר)
    sc = JSX_CLEAN_MAGIC.replace('%LNAME%', target_layer).replace('%GNAME%', unique_name_print)
    sc = sc.replace('%R%', str(r)).replace('%G%', str(g)).replace('%B%', str(b))
    sc = sc.replace('%DOCOL%', do_col)
    sc = sc.replace('%ISRASTER%', is_raster_str)
    run_jsx(app, sc)
    time.sleep(0.2)
    # 2. מיקום חכם ושינוי גודל
    resize = "true" if cat in ["Pocket", "2Pocket", "A4", "A5", "2"] else "false"
    is_p = "true"
    sc_pos = JSX_SMART_POS.replace('%ITEM%', unique_name_print)
    sc_pos = sc_pos.replace('%PRE%', pre).replace('%CAT%', cat)
    sc_pos = sc_pos.replace('%RES%', resize).replace('%ISP%', is_p)
    sc_pos = sc_pos.replace('%ABNAME%', target_ab).replace('%LNAME%', target_layer)
    run_jsx(app, sc_pos)
    # 3. הדמיה (שכפול) – רק כשלא skip_simulation
    if not skip_simulation:
        unique_name_sim = f"S_{pre}_{uuid.uuid4().hex[:6]}"
        should_recolor_sim = 'false'
        rs, gs, bs = (0,0,0)
        if s_hex:
            rs, gs, bs = hex_to_rgb(s_hex)
            should_recolor_sim = 'true'
        elif p_hex:
            rs, gs, bs = hex_to_rgb(p_hex)
            should_recolor_sim = 'true'
        sc_dup = JSX_DUPLICATE_AND_POS.replace('%ORIG%', unique_name_print)
        sc_dup = sc_dup.replace('%SIM%', unique_name_sim)
        sc_dup = sc_dup.replace('%R%', str(rs)).replace('%G%', str(gs)).replace('%B%', str(bs))
        sc_dup = sc_dup.replace('%PRE%', pre).replace('%CAT%', cat)
        sc_dup = sc_dup.replace('%DORECOLOR%', should_recolor_sim)
        run_jsx(app, sc_dup)
    p_lay.Visible = True
    # 4. === מדידה סופית ומדויקת ===
    final_true_width = 0
    try:
        measure_jsx = JSX_MEASURE_FINAL.replace("%NAME%", unique_name_print)
        res = app.DoJavaScript(measure_jsx)
        final_true_width = float(res)
    except:
        final_true_width = initial_check_w
    # 5. עדכון טקסט עם הרוחב הנכון
    if final_true_width > 0:
        target_tf = ""
        txt_suffix = ""
        if pre == "F":
            target_tf = "size_Front"
            txt_suffix = "קדמי"
        elif pre == "B":
            target_tf = "size_Back"
            txt_suffix = "אחורי"
        elif pre == "RS":
            target_tf = "size_Right_Sleeve"
            txt_suffix = "שרוול ימין"
        elif pre == "LS":
            target_tf = "size_Left_Sleeve"
            txt_suffix = "שרוול שמאל"
        if target_tf and should_update_size_label:
            update_size_label(doc, app, target_tf, final_true_width, txt_suffix)
    return final_true_width


def variable_layer_and_artboard(
    side: str, variant_index: int, layers_only: bool = False
) -> tuple[str, str]:
    cfg = VARIABLE_SIDE_BASE[side]
    layer = f"{cfg['layer']}_{variant_index}"
    if layers_only:
        artboard = f"{cfg['layer']}_1"
    else:
        artboard = f"{cfg['layer']}_{variant_index}"
    return layer, artboard
def open_and_color_template(path: str, h1: str, h2: str, is_split: bool, prod: str="Shirt"):
    print(f"--- Opening AI: {os.path.basename(path)} ---")
    app = win32com.client.Dispatch("Illustrator.Application")
    app.UserInteractionLevel = -1
    doc = app.Open(path)
    r1, g1, b1 = hex_to_rgb(h1)
    r2, g2, b2 = hex_to_rgb(h2)
    # חישוב צבע ה-Stroke (המיתר) - לבן לכהה, שחור לבהיר
    sr1, sg1, sb1 = (255, 255, 255) if (0.299*r1 + 0.587*g1 + 0.114*b1) < 128 else (0, 0, 0)
    sr2, sg2, sb2 = (255, 255, 255) if (0.299*r2 + 0.587*g2 + 0.114*b2) < 128 else (0, 0, 0)
    s = JSX_COLOR_PROD.replace('%PROD%', prod)
    s = s.replace('%IS_SPLIT%', "true" if is_split else "false")
    # החלפת ערכים לצד 1
    s = s.replace('%R1%', str(r1)).replace('%G1%', str(g1)).replace('%B1%', str(b1))
    s = s.replace('%SR1%', str(sr1)).replace('%SG1%', str(sg1)).replace('%SB1%', str(sb1))
    # החלפת ערכים לצד 2
    s = s.replace('%R2%', str(r2)).replace('%G2%', str(g2)).replace('%B2%', str(b2))
    s = s.replace('%SR2%', str(sr2)).replace('%SG2%', str(sg2)).replace('%SB2%', str(sb2))
    run_jsx(app, s)
    return doc, app
def set_order_number_in_simulation(app, order_id: str):
    """מעדכן או יוצר תיבת טקסט 'NumberOrder' בשכבת Simulation עם מספר ההזמנה (עמוד ראשון / מוצר 1)."""
    if not order_id:
        return
    # Escape single quotes in order_id for JS string
    safe_id = order_id.replace("\\", "\\\\").replace("'", "\\'").replace("\r", "").replace("\n", " ")
    jsx_code = """
    #target illustrator
    (function() {
        var orderId = '%ORDER_ID%';
        var doc = app.activeDocument;
        var simLayer = null;
        try { simLayer = doc.layers.getByName("Simulation"); } catch(e) { return 0; }
        if (!simLayer) return 0;
        simLayer.visible = true;
        simLayer.locked = false;
        function findTextFrame(container, name) {
            try {
                if (container.textFrames && container.textFrames.getByName) {
                    return container.textFrames.getByName(name);
                }
            } catch(e) {}
            if (container.pageItems) {
                for (var i = 0; i < container.pageItems.length; i++) {
                    var it = container.pageItems[i];
                    if (it.name === name && it.typename === "TextFrame") return it;
                    if (it.typename === "GroupItem" && it.pageItems.length > 0) {
                        var r = findTextFrame(it, name);
                        if (r) return r;
                    }
                }
            }
            if (container.layers) {
                for (var j = 0; j < container.layers.length; j++) {
                    var r = findTextFrame(container.layers[j], name);
                    if (r) return r;
                }
            }
            return null;
        }
        var tf = findTextFrame(doc, "NumberOrder") || findTextFrame(simLayer, "NumberOrder");
        if (tf) {
            tf.contents = orderId;
            if (tf.locked) tf.locked = false;
            return 1;
        }
        var rect = simLayer.visibleBounds;
        if (!rect || rect.length < 4) rect = [50, -50, 250, -100];
        var top = rect[1]; var left = rect[0];
        tf = simLayer.textFrames.add();
        tf.name = "NumberOrder";
        tf.contents = orderId;
        tf.position = [left, top - 20];
        try { tf.textRange.characterAttributes.size = 14; } catch(e) {}
        return 1;
    })();
    """.replace("%ORDER_ID%", safe_id)
    run_jsx(app, jsx_code)


def set_fabric_type_in_simulation(app, fabric_type: str):
    """מעדכן תיבת טקסט 'type' בשכבת Simulation עם ערך fabric_type (הדמיית חולצה)."""
    if not fabric_type:
        return
    safe = (
        str(fabric_type)
        .replace("\\", "\\\\")
        .replace("'", "\\'")
        .replace("\r", "")
        .replace("\n", " ")
    )
    jsx_code = """
    #target illustrator
    (function() {
        var fabricType = '%FABRIC_TYPE%';
        var doc = app.activeDocument;
        var simLayer = null;
        try { simLayer = doc.layers.getByName("Simulation"); } catch(e) { return 0; }
        if (!simLayer) return 0;
        simLayer.visible = true;
        simLayer.locked = false;
        function findTextFrame(container, name) {
            try {
                if (container.textFrames && container.textFrames.getByName) {
                    return container.textFrames.getByName(name);
                }
            } catch(e) {}
            if (container.pageItems) {
                for (var i = 0; i < container.pageItems.length; i++) {
                    var it = container.pageItems[i];
                    if (it.name === name && it.typename === "TextFrame") return it;
                    if (it.typename === "GroupItem" && it.pageItems.length > 0) {
                        var r = findTextFrame(it, name);
                        if (r) return r;
                    }
                }
            }
            if (container.layers) {
                for (var j = 0; j < container.layers.length; j++) {
                    var r = findTextFrame(container.layers[j], name);
                    if (r) return r;
                }
            }
            return null;
        }
        var tf = findTextFrame(doc, "type") || findTextFrame(simLayer, "type")
              || findTextFrame(doc, "Type") || findTextFrame(simLayer, "Type");
        if (!tf) return 0;
        var p = tf;
        while (p) {
            if (p.locked) p.locked = false;
            try { p = p.parent; } catch(e) { break; }
            if (p.typename == "Layer") break;
        }
        tf.contents = fabricType;
        return 1;
    })();
    """.replace("%FABRIC_TYPE%", safe)
    run_jsx(app, jsx_code)


def remove_order_number_from_simulation(app):
    """מוחק את תיבת הטקסט 'NumberOrder' משכבת Simulation (למוצרים 2 ומעלה – להשאיר מספר הזמנה רק בראשון)."""
    jsx_code = """
    #target illustrator
    (function() {
        var doc = app.activeDocument;
        var simLayer = null;
        try { simLayer = doc.layers.getByName("Simulation"); } catch(e) { return; }
        if (!simLayer) return;
        function findAndRemove(container, name) {
            try {
                if (container.textFrames && container.textFrames.getByName) {
                    var tf = container.textFrames.getByName(name);
                    if (tf) { tf.remove(); return true; }
                }
            } catch(e) {}
            if (container.pageItems) {
                for (var i = container.pageItems.length - 1; i >= 0; i--) {
                    var it = container.pageItems[i];
                    if (it.name === name && it.typename === "TextFrame") { it.remove(); return true; }
                    if (it.typename === "GroupItem" && it.pageItems.length > 0) {
                        if (findAndRemove(it, name)) return true;
                    }
                }
            }
            if (container.layers) {
                for (var j = container.layers.length - 1; j >= 0; j--) {
                    if (findAndRemove(container.layers[j], name)) return true;
                }
            }
            return false;
        }
        try { doc.textFrames.getByName("NumberOrder").remove(); } catch(e) {}
        findAndRemove(simLayer, "NumberOrder");
    })();
    """
    run_jsx(app, jsx_code)


def delete_side_assets(doc, app, ab: str, tf: str):
    run_jsx(app, JSX_DEL.replace('%AB%', ab).replace('%TF%', tf))

_PRINT_LAYER_BY_PREFIX = {
    "F": "Print_Front",
    "B": "Print_Back",
    "RS": "Print_Right_Sleeve",
    "LS": "Print_Left_Sleeve",
}


def delete_print_layer_only(app, prefix: str):
    """מוחק תוכן שכבת Print + artboard — משאיר Simulation ותווית size."""
    layer_name = _PRINT_LAYER_BY_PREFIX.get(prefix)
    ab_name = am.get(prefix)
    if not layer_name or not ab_name:
        return
    jsx = f"""
    #target illustrator
    (function() {{
        try {{
            var doc = app.activeDocument;
            try {{
                var layer = doc.layers.getByName("{layer_name}");
                layer.locked = false;
                layer.visible = true;
                while (layer.pageItems.length > 0) layer.pageItems[0].remove();
            }} catch(e) {{}}
            try {{ doc.artboards.getByName("{ab_name}").remove(); }} catch(e) {{}}
        }} catch(e) {{}}
    }})();
    """
    run_jsx(app, jsx)
def save_pdf(doc, path: str):
    try:
        o = win32com.client.Dispatch("Illustrator.PDFSaveOptions")
        o.PreserveEditability = True
        doc.SaveAs(path, o)
    except: pass
    finally:
        try: doc.Close(2)
        except: pass
def outline_text_in_layers(app, layer_names: list) -> None:
    """הופך TextFrames בשכבות הנתונות ל-outline."""
    if not layer_names:
        return
    layers_js = json.dumps([str(n) for n in layer_names], ensure_ascii=False)
    jsx = f"""
    #target illustrator
    (function() {{
        var layerNames = {layers_js};
        app.executeMenuCommand("unlockAll");
        function collectTextFrames(container, out) {{
            try {{
                if (container.textFrames) {{
                    for (var t = 0; t < container.textFrames.length; t++) out.push(container.textFrames[t]);
                }}
            }} catch(e) {{}}
            if (container.pageItems) {{
                for (var i = 0; i < container.pageItems.length; i++) {{
                    var it = container.pageItems[i];
                    if (it.typename === "TextFrame") out.push(it);
                    else if (it.typename === "GroupItem") collectTextFrames(it, out);
                }}
            }}
            if (container.layers) {{
                for (var j = 0; j < container.layers.length; j++) collectTextFrames(container.layers[j], out);
            }}
        }}
        var outlined = 0;
        for (var li = 0; li < layerNames.length; li++) {{
            var layer = null;
            try {{ layer = app.activeDocument.layers.getByName(layerNames[li]); }} catch(e) {{ continue; }}
            if (!layer) continue;
            layer.locked = false;
            layer.visible = true;
            var frames = [];
            collectTextFrames(layer, frames);
            for (var fi = 0; fi < frames.length; fi++) {{
                try {{
                    frames[fi].locked = false;
                    frames[fi].createOutline();
                    outlined++;
                }} catch(e) {{}}
            }}
        }}
        return outlined;
    }})();
    """
    try:
        res = app.DoJavaScript(jsx)
        count = int(res) if res not in (None, "") else 0
        print(f"   > Outlined {count} text frame(s) in print layers")
    except Exception as e:
        print(f"   > Outline pass warning: {e}")


def outline_document_text(app) -> None:
    """הופך TextFrames ל-outline — לפי שכבה בבatch כדי לשמור מיקום יחסי בין שדות."""
    jsx = """
    #target illustrator
    (function() {
        app.executeMenuCommand("unlockAll");
        var doc = app.activeDocument;
        function collectTextFrames(container, out) {
            try {
                if (container.textFrames) {
                    for (var t = 0; t < container.textFrames.length; t++) out.push(container.textFrames[t]);
                }
            } catch(e) {}
            if (container.pageItems) {
                for (var i = 0; i < container.pageItems.length; i++) {
                    var it = container.pageItems[i];
                    if (it.typename === "TextFrame") out.push(it);
                    else if (it.typename === "GroupItem") collectTextFrames(it, out);
                }
            }
        }
        function collectAllLayers(container, out) {
            if (!container.layers) return;
            for (var i = 0; i < container.layers.length; i++) {
                out.push(container.layers[i]);
                collectAllLayers(container.layers[i], out);
            }
        }
        function outlineFramesTogether(frames) {
            if (!frames || frames.length === 0) return 0;
            if (frames.length === 1) {
                try {
                    frames[0].locked = false;
                    frames[0].createOutline();
                    return 1;
                } catch(e) { return 0; }
            }
            doc.selection = null;
            for (var i = 0; i < frames.length; i++) {
                try {
                    frames[i].locked = false;
                    frames[i].selected = true;
                } catch(e) {}
            }
            try {
                app.executeMenuCommand("outline");
                doc.selection = null;
                return frames.length;
            } catch(e) {
                var n = 0;
                for (var j = 0; j < frames.length; j++) {
                    try { frames[j].createOutline(); n++; } catch(e2) {}
                }
                return n;
            }
        }
        var layers = [];
        collectAllLayers(doc, layers);
        var outlined = 0;
        for (var li = 0; li < layers.length; li++) {
            var layer = layers[li];
            try { layer.locked = false; } catch(e) {}
            var frames = [];
            collectTextFrames(layer, frames);
            outlined += outlineFramesTogether(frames);
        }
        try { doc.selection = null; } catch(e) {}
        return outlined;
    })();
    """
    try:
        res = app.DoJavaScript(jsx)
        count = int(res) if res not in (None, "") else 0
        print(f"   > Outlined {count} text frame(s) in final document (print + simulation + labels)")
    except Exception as e:
        print(f"   > Final outline warning: {e}")


def clean_layout(app):
    """מוחק את כל ריבועי העזר (Box) מהקובץ"""
    run_jsx(app, JSX_CLEAN_BOXES)
def delete_information_layer(app):
    """מוחק את השכבה/קבוצה 'information' מתוך 'Simulation' בכל מוצר"""
    jsx_code = """
    #target illustrator
    (function() {
        var doc = app.activeDocument;
        var deletedCount = 0;
        // פונקציה רקורסיבית למחיקת information מתוך Simulation
        function deleteInfoFromSimulation(container) {
            var found = false;
            // חיפוש שכבה Simulation
            try {
                var simLayer = null;
                // ניסיון למצוא Simulation כשכבה
                try {
                    simLayer = container.layers.getByName("Simulation");
                } catch(e) {
                    // אם לא נמצאה כשכבה, נחפש רקורסיבית
                    for (var i = 0; i < container.layers.length; i++) {
                        var result = deleteInfoFromSimulation(container.layers[i]);
                        if (result) found = true;
                    }
                }
                if (simLayer) {
                    // פתיחת נעילה של שכבה Simulation אם היא נעולה
                    if (simLayer.locked) {
                        simLayer.locked = false;
                    }
                    // ניסיון למחוק שכבה בשם "information"
                    try {
                        var infoLayer = simLayer.layers.getByName("information");
                        if (infoLayer) {
                            // פתיחת נעילה אם השכבה נעולה
                            if (infoLayer.locked) {
                                infoLayer.locked = false;
                            }
                            // פתיחת נעילה של כל תתי-השכבות
                            for (var j = 0; j < infoLayer.layers.length; j++) {
                                try {
                                    if (infoLayer.layers[j].locked) {
                                        infoLayer.layers[j].locked = false;
                                    }
                                } catch(e) {}
                            }
                            infoLayer.remove();
                            deletedCount++;
                            found = true;
                        }
                    } catch(e) {}
                    // ניסיון למחוק קבוצה בשם "information"
                    try {
                        var infoGroup = simLayer.groupItems.getByName("information");
                        if (infoGroup) {
                            // פתיחת נעילה אם הקבוצה נעולה
                            if (infoGroup.locked) {
                                infoGroup.locked = false;
                            }
                            // פתיחת נעילה של כל הפריטים בקבוצה
                            for (var k = 0; k < infoGroup.pageItems.length; k++) {
                                try {
                                    if (infoGroup.pageItems[k].locked) {
                                        infoGroup.pageItems[k].locked = false;
                                    }
                                } catch(e) {}
                            }
                            infoGroup.remove();
                            deletedCount++;
                            found = true;
                        }
                    } catch(e) {}
                    // חיפוש רקורסיבי בתוך תתי-שכבות (אם יש Simulation בתוך Simulation)
                    for (var l = 0; l < simLayer.layers.length; l++) {
                        var subResult = deleteInfoFromSimulation(simLayer.layers[l]);
                        if (subResult) found = true;
                    }
                }
            } catch(e) {}
            return found;
        }
        try {
            // גישה 1: חיפוש ישיר של Simulation במסמך (למוצר בודד)
            try {
                var directSim = doc.layers.getByName("Simulation");
                if (directSim) {
                    if (directSim.locked) directSim.locked = false;
                    // ניסיון למחוק information ישירות
                    try {
                        var infoLayer = directSim.layers.getByName("information");
                        if (infoLayer) {
                            if (infoLayer.locked) infoLayer.locked = false;
                            for (var j = 0; j < infoLayer.layers.length; j++) {
                                try { if (infoLayer.layers[j].locked) infoLayer.layers[j].locked = false; } catch(e) {}
                            }
                            infoLayer.remove();
                            deletedCount++;
                        }
                    } catch(e) {}
                    try {
                        var infoGroup = directSim.groupItems.getByName("information");
                        if (infoGroup) {
                            if (infoGroup.locked) infoGroup.locked = false;
                            for (var k = 0; k < infoGroup.pageItems.length; k++) {
                                try { if (infoGroup.pageItems[k].locked) infoGroup.pageItems[k].locked = false; } catch(e) {}
                            }
                            infoGroup.remove();
                            deletedCount++;
                        }
                    } catch(e) {}
                }
            } catch(e) {}
            // גישה 2: חיפוש דרך שכבות מוצרים (1, 2, 3, 4...) - למקרה של איחוד
            for (var i = 0; i < doc.layers.length; i++) {
                var mainLayer = doc.layers[i];
                var layerName = mainLayer.name;
                // בדיקה אם זו שכבה של מוצר (מספר)
                if (/^\\d+$/.test(layerName)) {
                    deleteInfoFromSimulation(mainLayer);
                }
            }
        } catch(e) {}
        // החזרת מספר הפריטים שנמחקו (לצורך log)
        return deletedCount;
    })();
    """
    try:
        result = app.DoJavaScript(jsx_code)
        deleted_count = int(result) if result else 0
        if deleted_count > 0:
            print(f"   ✓ נמחקו {deleted_count} שכבות/קבוצות 'information'")
        else:
            print(f"   ⚠ לא נמצאו שכבות/קבוצות 'information' למחיקה")
    except Exception as e:
        print(f"   ❌ שגיאה במחיקת שכבה 'information': {e}")
def apply_extra_colors(app, extra_data_list: list):
    import json
    # וודאי שאין שימוש בשם hex_list אם הוא לא הוגדר
    if not extra_data_list:
        extra_data_list = []
    formatted_rgb = []
    for pair in extra_data_list:
        rgb_pair = [list(hex_to_rgb(h)) for h in pair]
        formatted_rgb.append(rgb_pair)
    rgb_json = json.dumps(formatted_rgb)
    final_jsx = JSX_EXTRA_COLORS.replace("%COLOR_ARRAY%", rgb_json)
    run_jsx(app, final_jsx)
