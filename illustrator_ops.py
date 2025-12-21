# -*- coding: utf-8 -*-



from __future__ import annotations



import win32com.client

import os

import uuid

import time

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

        print(f"!!! JSX Error (Might be harmless): {e}")



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


    if (category === "Sleeve2") suffix = "Sleeve2"; // תוספת עבור שרוול גדול
    else if (catLower.indexOf("sleeve") !== -1 || catLower.indexOf("9") !== -1 || catLower.indexOf("שרוול") !== -1) suffix = "Sleeve";

    else if (category === "Pocket") suffix = "Pocket";

    else if (category === "A3") suffix = "A3";
    else if (category === "A5") suffix = "A5";

    else if (category === "A4") {

        var ratio = itemW / itemH;

        

        // === גבול מדויק: 1.21 ===

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



function smartPos(itemName, prefix, category, resizeArtboard, isPrint, artboardName) {

    try {

        var doc = app.activeDocument;

        var item = doc.pageItems.getByName(itemName);

        item.hidden = false;

        var itemW = item.width; var itemH = item.height; if(itemH == 0) itemH = 1;

        var suffix = "A4_Square";

        var catLower = category.toLowerCase();


        if (category === "Sleeve2") suffix = "Sleeve2";
        else if (catLower.indexOf("sleeve") !== -1 || catLower.indexOf("9") !== -1 || catLower.indexOf("שרוול") !== -1) suffix = "Sleeve";

        else if (category === "Pocket") suffix = "Pocket";

        else if (category === "A3") suffix = "A3";
        else if (category === "A5") suffix = "A5";
        else if (category === "A4") {

            var ratio = itemW / itemH;

            

            // === גבול מדויק: 1.21 ===

            if (ratio > 1.21) suffix = "A4_Landscape"; 

            else if (ratio < 0.75) suffix = "A4_Portrait"; 

            else suffix = "A4_Square";

        }



        var boxPrefix = isPrint ? prefix : "S" + prefix;

        var boxName = boxPrefix + "_Box_" + suffix;

        var box = null;

        try { box = doc.pageItems.getByName(boxName); } catch(e) { return; } 

        var b = box.visibleBounds; 

        var boxW = b[2] - b[0]; var boxH = b[1] - b[3];

        var cx = b[0] + boxW/2.0; var cy = b[1] - boxH/2.0;



        var scale = 100.0;

        

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



        if (isPrint && resizeArtboard === true && artboardName) {

            try {

                var ab = doc.artboards.getByName(artboardName);

                var doResize = false; var newW = 0; var newH = 0;

                if (suffix === "Pocket" || suffix === "A4_Portrait" || suffix === "A5") { newW = 595.28; newH = 841.89; doResize = true; }

                else if (suffix === "A4_Landscape") { newW = 841.89; newH = 595.28; doResize = true; }

                if (doResize) ab.artboardRect = [cx - newW/2.0, cy + newH/2.0, cx + newW/2.0, cy - newH/2.0];

            } catch(e) { }

        }

    } catch(e) { }

}



try { var isRes = ("%RES%" === "true"); var isP = ("%ISP%" === "true"); smartPos("%ITEM%", "%PRE%", "%CAT%", isRes, isP, "%ABNAME%"); } catch(e) { }



"""



JSX_COLOR_PROD = """

#target illustrator



function col(it,r,g,b,sr,sg,sb) {

    var f=new RGBColor();f.red=r;f.green=g;f.blue=b;

    var s=new RGBColor();s.red=sr;s.green=sg;s.blue=sb;

    

    // === החזרנו את המצב לקדמותו: ===

    // אם בשם יש "String" או "מיתר" - תעצור מיד ואל תצבע כלום.

    if(it.name && (it.name.indexOf("String")!==-1 || it.name.indexOf("מיתר")!==-1)) return;

    

    if(it.typename==='PathItem' && !it.clipping){

        it.filled=true; it.fillColor=f; 

        it.stroked=true; it.strokeColor=s; it.strokeWidth=1;

    }

    else if(it.typename==='CompoundPathItem'){

        for(var i=0;i<it.pathItems.length;i++){

            var p=it.pathItems[i];

            if(!p.clipping){

                p.filled=true; p.fillColor=f; 

                p.stroked=true; p.strokeColor=s;

            }

        }

    }

    else if(it.typename==='GroupItem'){

        for(var j=0;j<it.pageItems.length;j++) col(it.pageItems[j],r,g,b,sr,sg,sb);

    }

}



try{

    var d=app.activeDocument;

    var l=d.layers.getByName("Simulation");

    var g=null;

    

    // סדר החיפוש נשאר כמו שביקשת (קודם כל Simulation בתוך Simulation)

    try { g=l.groupItems.getByName("Simulation"); } catch(e) {}

    

    // אם לא מצא, מחפש את השם המשתנה (כמו Hoodie/Pants)

    if(!g) try { g=l.groupItems.getByName("%PROD%"); } catch(e) {}

    

    // ברירת מחדל אחרונה

    if(!g) try { g=l.groupItems.getByName("Shirt"); } catch(e) {}



    if(g) col(g,%R%,%G%,%B%,%SR%,%SG%,%SB%);

}catch(e){ }

"""



JSX_DEL = """

#target illustrator

try{app.activeDocument.artboards.getByName("%AB%").remove();}catch(e){}

try{app.activeDocument.textFrames.getByName("%TF%").remove();}catch(e){}

"""



JSX_CLEAN_BOXES = """

#target illustrator

try {

    var doc = app.activeDocument;

    // רצים בלולאה הפוכה (חשוב מאוד במחיקה)

    for (var i = doc.pageItems.length - 1; i >= 0; i--) {

        var item = doc.pageItems[i];

        // בדיקה: אם השם מכיל "_Box_", זה ריבוע עזר -> למחוק!

        if (item.name.indexOf("_Box_") !== -1) {

            item.remove();

        }

    }

} catch(e) {}

"""



JSX_EXTRA_COLORS = """

#target illustrator

try {

    var doc = app.activeDocument;

    var simLayer = doc.layers.getByName("Simulation");

    

    // === התיקון: חיפוש חכם (גם קבוצה וגם תת-שכבה) ===

    var colorContainer = null;

    

    // נסיון 1: האם זו קבוצה? (Group)

    try { colorContainer = simLayer.groupItems.getByName("Box_Color"); } catch(e) {}

    

    // נסיון 2: האם זו תת-שכבה? (Sub-Layer)

    if (!colorContainer) {

        try { colorContainer = simLayer.layers.getByName("Box_Color"); } catch(e) {}

    }

    // =================================================



    if (colorContainer) {

        // קבלת רשימת הצבעים מהפייתון

        var colors = %COLOR_ARRAY%; 



        if (colors.length === 0) {

            // אם אין צבעים נוספים - מוחקים את הקונטיינר

            colorContainer.remove();

        } else {

            // לולאה על כל ה-24 ריבועים

            for (var i = 1; i <= 24; i++) {

                var boxName = "Color_" + i;

                try {

                    // חיפוש הריבוע בתוך הקונטיינר (עובד גם לקבוצה וגם לשכבה)

                    var box = colorContainer.pageItems.getByName(boxName);

                    

                    if (i <= colors.length) {

                        // צביעה

                        var cVal = colors[i-1];

                        var newCol = new RGBColor();

                        newCol.red = cVal[0]; newCol.green = cVal[1]; newCol.blue = cVal[2];

                        

                        if (box.typename === 'PathItem') {

                            box.filled = true; box.fillColor = newCol;

                        } else if (box.typename === 'CompoundPathItem') {

                             if(box.pathItems.length > 0) { box.pathItems[0].filled = true; box.pathItems[0].fillColor = newCol; }

                        }

                    } else {

                        // מחיקת ריבועים מיותרים

                        box.remove();

                    }

                } catch(e) {

                    // ריבוע לא נמצא - מדלגים

                }

            }

        }

    }

} catch(e) {

    // שגיאה כללית (למשל שכבת Simulation חסרה)

}

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



def get_doc_safe(app):

    for i in range(5):

        try:

            if app.Documents.Count > 0:

                return app.ActiveDocument

        except:

            time.sleep(0.5)

    return None



def get_layer(doc, name):

    try:

        l = doc.Layers(name)

        l.Visible = True; l.Locked = False

        return l

    except: return None



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

    final_text = f"{width_cm} ס\"מ הדפס {txt}"

    

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



def place_and_simulate_print(doc, app, path, pre, cat, p_hex, s_hex, is_raster=False):

    print(f"--- Processing {pre} ---")

    

    l_map = {"F":"Print_Front","B":"Print_Back","RS":"Print_Right_Sleeve","LS":"Print_Left_Sleeve"}

    

    # וידוא מסמך

    doc = get_doc_safe(app)

    if not doc: return 0



    p_lay = get_layer(doc, l_map[pre])

    if not p_lay: return 0



    unique_name_print = f"P_{pre}_{uuid.uuid4().hex[:6]}"

    

    # משתנה זמני לבדיקה שההטמעה הצליחה

    initial_check_w = 0 



    try:

        if is_raster:

            # --- הטמעת רסטר (תמונה) ---

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

            placeRaster('{safe_path}', '{l_map[pre]}', '{unique_name_print}');

            """

            raw_width = app.DoJavaScript(jsx_place_raster)

            initial_check_w = float(raw_width)



        else:

            # --- הטמעת וקטור ---

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

    

    is_raster_str = "true" if is_raster else "false"

    

    # מריצים את הניקוי

    sc = JSX_CLEAN_MAGIC.replace('%LNAME%', l_map[pre]).replace('%GNAME%', unique_name_print)

    sc = sc.replace('%R%', str(r)).replace('%G%', str(g)).replace('%B%', str(b))

    sc = sc.replace('%DOCOL%', do_col)

    sc = sc.replace('%ISRASTER%', is_raster_str) 

    run_jsx(app, sc)

    

    time.sleep(0.2)



    # 2. מיקום חכם ושינוי גודל (כאן הגודל משתנה!)

    resize = "true" if cat in ["Pocket", "A4", "A5"] else "false"

    is_p = "true"

    ab_name = am.get(pre, "")

    

    sc_pos = JSX_SMART_POS.replace('%ITEM%', unique_name_print)

    sc_pos = sc_pos.replace('%PRE%', pre).replace('%CAT%', cat)

    sc_pos = sc_pos.replace('%RES%', resize).replace('%ISP%', is_p)

    sc_pos = sc_pos.replace('%ABNAME%', ab_name)

    

    run_jsx(app, sc_pos)

    

    # 3. הדמיה (שכפול)

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

    # אנחנו שואלים את אילוסטרייטור מה הרוחב *עכשיו*, אחרי הניקוי והשינוי גודל

    final_true_width = 0

    try:

        measure_jsx = JSX_MEASURE_FINAL.replace("%NAME%", unique_name_print)

        res = app.DoJavaScript(measure_jsx)

        final_true_width = float(res)

    except:

        final_true_width = initial_check_w # גיבוי למקרה של כישלון



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

        

        if target_tf:

            update_size_label(doc, app, target_tf, final_true_width, txt_suffix)

            

    return final_true_width

def open_and_color_template(path: str, hex_c: Optional[str], prod: str="Shirt"):

    print(f"--- Opening AI: {os.path.basename(path)} ---")

    app = win32com.client.Dispatch("Illustrator.Application")

    app.UserInteractionLevel = -1 

    doc = app.Open(path) 

    

    r, g, b = hex_to_rgb(hex_c) if hex_c else (255,255,255)

    sr, sg, sb = (255, 255, 255) if (0.299*r + 0.587*g + 0.114*b) < 128 else (0, 0, 0)



    s = JSX_COLOR_PROD.replace('%PROD%', prod)

    s = s.replace('%R%', str(r)).replace('%G%', str(g)).replace('%B%', str(b))

    s = s.replace('%SR%', str(sr)).replace('%SG%', str(sg)).replace('%SB%', str(sb))

    run_jsx(app, s)

    return doc, app



def delete_side_assets(doc, app, ab: str, tf: str):

    run_jsx(app, JSX_DEL.replace('%AB%', ab).replace('%TF%', tf))



def save_pdf(doc, path: str):

    try:

        o = win32com.client.Dispatch("Illustrator.PDFSaveOptions")

        o.PreserveEditability = True

        doc.SaveAs(path, o)

    except: pass

    finally:

        try: doc.Close(2)

        except: pass



def clean_layout(app):

    """מוחק את כל ריבועי העזר (Box) מהקובץ"""

    run_jsx(app, JSX_CLEAN_BOXES)



def apply_extra_colors(app, hex_list: list):

    """ מקבלת רשימה של קודי HEX לצבעים נוספים.

    מטפלת בתיקייה Box_Color: מוחקת אותה אם הרשימה ריקה, או צובעת את הריבועים אם יש תוכן."""

    if hex_list is None:

        hex_list = []

        

    # המרת רשימת ה-HEX לרשימה של [r, g, b] עבור הסקריפט

    rgb_array_str = "["

    for h in hex_list:

        if h:

            r, g, b = hex_to_rgb(h)

            rgb_array_str += f"[{r},{g},{b}],"

        

    # סגירת המערך (מחיקת הפסיק האחרון אם יש)

    if rgb_array_str.endswith(","):

        rgb_array_str = rgb_array_str[:-1]

    rgb_array_str += "]"



    # הרצת הסקריפט

    final_jsx = JSX_EXTRA_COLORS.replace("%COLOR_ARRAY%", rgb_array_str)

    run_jsx(app, final_jsx)