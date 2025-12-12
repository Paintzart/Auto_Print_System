# -*- coding: utf-8 -*-
from __future__ import annotations
import win32com.client
import os
import uuid
import time

# --- הגדרות גלובליות ---
am = {
    "F": "Print_Front", 
    "B": "Print_Back", 
    "RS": "Print_Sleeves", 
    "LS": "Print_Sleeves"
}

def hex_to_rgb(h):
    if not h: return (0,0,0)
    h = h.lstrip('#')
    if len(h) != 6: return (0,0,0)
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def run_jsx(app, s):
    """מריץ את ה-JSX עם הגנה מפני קריסות"""
    try:
        app.DoJavaScript(s)
    except Exception as e:
        print(f"!!! JSX Error (Might be harmless): {e}")

# --- סקריפטים JSX ---

JSX_CLEAN_MAGIC = """
#target illustrator

// פונקציה להשוואת צבעים (הקטנו את הטולרנס ל-1 כדי להיות מדויקים יותר)
function isSameColor(c1, c2) {
    if (!c1 || !c2) return false;
    if (c1.typename !== c2.typename) return false;
    
    var t = 1; // שינוי: רגישות גבוהה יותר (היה 3)
    
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

function run(ln, grpN, r, g, b, doC) {
    try {
        var doc = app.activeDocument;
        var group = doc.pageItems.getByName(grpN);
        
        // 1. ניקוי "זבל" בסוף הקבוצה (השכבה הכי תחתונה)
        try {
            var count = group.pageItems.length;
            if (count > 0) {
                var lastItem = group.pageItems[count - 1];
                if (lastItem.typename === "GroupItem" || lastItem.typename === "PathItem") {
                    lastItem.remove();
                }
            }
        } catch(e){}

        if (group.typename === 'GroupItem' && group.pageItems.length > 0) {
            
            var totalW = group.width;
            var totalH = group.height;
            var totalArea = totalW * totalH;
            var detectedBgColor = null;

            // 2. זיהוי הרקע הגדול (מעל 90%)
            // התיקון: ברגע שמוצאים אחד - מוחקים ועוצרים!
            for (var i = group.pageItems.length - 1; i >= 0; i--) {
                var item = group.pageItems[i];
                var iArea = item.width * item.height;
                
                if (iArea > (totalArea * 0.90)) {
                    // דגימת הצבע
                    if (!detectedBgColor) {
                        if (item.typename === 'PathItem' && item.filled) detectedBgColor = item.fillColor;
                        else if (item.typename === 'CompoundPathItem' && item.pathItems.length > 0 && item.pathItems[0].filled) 
                            detectedBgColor = item.pathItems[0].fillColor;
                    }
                    
                    // מחיקת הרקע הגדול
                    item.remove();
                    
                    // === התיקון החשוב: עצור כאן! ===
                    // אם מצאנו ומחקנו רקע ענק, אין סיבה להמשיך למחוק דברים ענקיים אחרים (כמו הלוגו)
                    break; 
                }
            }

            // 3. הפעלת הניקוי העדין (רק לחלקים קטנים באותו צבע)
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
    var doColor = ("%DOCOL%" === "true");
    run("%LNAME%", "%GNAME%", %R%, %G%, %B%, doColor); 
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
    if (catLower.indexOf("sleeve") !== -1 || catLower.indexOf("9") !== -1 || catLower.indexOf("שרוול") !== -1) suffix = "Sleeve";
    else if (category === "Pocket") suffix = "Pocket";
    else if (category === "A3") suffix = "A3";
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
        if (catLower.indexOf("sleeve") !== -1 || catLower.indexOf("9") !== -1 || catLower.indexOf("שרוול") !== -1) suffix = "Sleeve";
        else if (category === "Pocket") suffix = "Pocket";
        else if (category === "A3") suffix = "A3";
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
                if (suffix === "Pocket" || suffix === "A4_Portrait") { newW = 595.28; newH = 841.89; doResize = true; }
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

def place_and_simulate_print(doc, app, path, pre, cat, p_hex, s_hex):
    print(f"--- Processing {pre} ---")
    
    l_map = {"F":"Print_Front","B":"Print_Back","RS":"Print_Right_Sleeve","LS":"Print_Left_Sleeve"}
    
    doc = get_doc_safe(app)
    if not doc: 
        print("XXX Error: No Document Found at start.")
        return 0

    p_lay = get_layer(doc, l_map[pre])
    if not p_lay: return 0

    try:
        imported_group = p_lay.GroupItems.CreateFromFile(path)
        clean_arts(imported_group)
    except Exception as e:
        print(f"Import Error: {e}")
        return 0
    
    if not imported_group: return 0
    
    unique_name_print = f"P_{pre}_{uuid.uuid4().hex[:6]}"
    imported_group.Name = unique_name_print
    
    do_col = 'false'
    r, g, b = (0,0,0)
    
    if p_hex:
        r, g, b = hex_to_rgb(p_hex)
        do_col = 'true'
    
    # 1. ניקוי וצביעה
    sc = JSX_CLEAN_MAGIC.replace('%LNAME%', l_map[pre]).replace('%GNAME%', unique_name_print)
    sc = sc.replace('%R%', str(r)).replace('%G%', str(g)).replace('%B%', str(b))
    sc = sc.replace('%DOCOL%', do_col)
    run_jsx(app, sc)
    
    time.sleep(0.5)

    # 2. מיקום חכם ושינוי גודל דף
    resize = "true" if cat in ["Pocket", "A4"] else "false"
    is_p = "true"
    ab_name = am.get(pre, "")
    
    sc_pos = JSX_SMART_POS.replace('%ITEM%', unique_name_print)
    sc_pos = sc_pos.replace('%PRE%', pre).replace('%CAT%', cat)
    sc_pos = sc_pos.replace('%RES%', resize).replace('%ISP%', is_p)
    sc_pos = sc_pos.replace('%ABNAME%', ab_name)
    
    run_jsx(app, sc_pos)
    
    # 3. חישוב רוחב
    final_w = 0
    try:
        doc = get_doc_safe(app)
        final_w = doc.PageItems(unique_name_print).Width
        print(f"DEBUG: Width calculated: {final_w}")
    except:
        print("DEBUG: Could not read width, but ignoring to continue flow.")

    # 4. הדמיה (שכפול)
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
    
    # 5. עדכון טקסט - החלק המעודכן
    if final_w > 0:
        target_tf = ""
        txt_suffix = ""
        
        if pre == "F": 
            target_tf = "size_Front"
            txt_suffix = "קדמי"
        elif pre == "B": 
            target_tf = "size_Back"
            txt_suffix = "אחורי"
        # === הוספה עבור שרוולים ===
        elif pre == "RS":
            target_tf = "size_Right_Sleeve"
            txt_suffix = "שרוול ימין"
        elif pre == "LS":
            target_tf = "size_Left_Sleeve"
            txt_suffix = "שרוול שמאל"
        # ==========================
        
        if target_tf:
            update_size_label(doc, app, target_tf, final_w, txt_suffix)
            
    return final_w


def open_and_color_template(path, hex_c, prod="Shirt"):
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

def delete_side_assets(doc, app, ab, tf):
    run_jsx(app, JSX_DEL.replace('%AB%', ab).replace('%TF%', tf))

def save_pdf(doc, path):
    try:
        o = win32com.client.Dispatch("Illustrator.PDFSaveOptions")
        o.PreserveEditability = True
        doc.SaveAs(path, o)
    except: pass
    finally:
        try: doc.Close(2)
        except: pass