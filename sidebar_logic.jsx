#target illustrator
// --- פונקציה לקריאת הנתונים שפייתון שולח (נתיבים ושמות קבצים) ---
function getJobData(jobPath) {
    var dataFile = new File(jobPath || new File($.fileName).path + "/current_job.json");
    if (!dataFile.exists) return null;
    dataFile.encoding = "UTF-8";
    if (!dataFile.open("r")) {
        $.writeln("Could not open sidebar job settings: " + dataFile.fsName + " (" + dataFile.error + ")");
        return null;
    }
    try {
        var content = dataFile.read();
        if (!content || !/\S/.test(content)) {
            $.writeln("Sidebar job settings are empty: " + dataFile.fsName);
            return null;
        }
        return eval("(" + content + ")");
    } catch (err) {
        $.writeln("Could not read sidebar job settings: " + dataFile.fsName + " (" + err + ")");
        return null;
    } finally {
        dataFile.close();
    }
}
var job = typeof sidebarJobData !== "undefined" ? sidebarJobData :
    getJobData(typeof sidebarJobPath === "undefined" ? null : sidebarJobPath);
var showSidebar = job ? job.show_sidebar : true;
var orderedProducts = job ? job.ordered_products : [];
var upsellMode = job ? job.upsell_mode : "random";
var manualProducts = job ? (job.manual_products || []) : [];
// הגדרת משתנים דינמיים (מגיע מהקונפיג של פייתון)
var sidebarMasterPath = job ? job.sidebar_path : "C:/Users/yarde/OneDrive/Desktop/Auto_Print_System/Simulations/Sidebar_Template.ai";
// --- שאר הגדרות המערכת שלך ---
var allUpsellOptions = ["Apron", "Drawstring Bag", "Wide Brimmed Hat", "Neck Warmer", "Canvas Bag", "Polo", "Fleece1", "Beanie", "Boxers", "Short", "Hoodie", "Hat"];
var allPalettes = [
    // הקיימים שלך (עם תיקוני פסיקים)
    {'name': 'Vintage Rose', 'dark': '#523245', 'medium': '#926266', 'light': '#AB7579'},
    {'name': 'Dusty Purple', 'dark': '#4A3B52', 'medium': '#7D6682', 'light': '#A68CA8'},
    {'name': 'Steel Blue',   'dark': '#2C3E50', 'medium': '#5D788F', 'light': '#87A0B5'},
    {'name': 'Forest Green', 'dark': '#1B2E1F', 'medium': '#3E5C44', 'light': '#6A8A70'},
    {'name': 'Deep Ocean',   'dark': '#1A2A3A', 'medium': '#34495E', 'light': '#5D788F'},
    {'name': 'Desert Sand',  'dark': '#7E6B5A', 'medium': '#A69076', 'light': '#CDBBA7'},
    {'name': 'Midnight Gold','dark': '#2C2C2C', 'medium': '#5E543D', 'light': '#9C8E6B'},
    {'name': 'Terracotta',   'dark': '#6E2C1B', 'medium': '#A0522D', 'light': '#CD853F'},
    {'name': 'Antique Gold', 'dark': '#594D28', 'medium': '#9C8948', 'light': '#C7B370'},
    {'name': 'Sage Green',   'dark': '#3B453B', 'medium': '#6B7F6D', 'light': '#94A896'},
    // פסטלים וחדשים (בהירים ורכים יותר)
    {'name': 'Lavender Mist', 'dark': '#6D5D7E', 'medium': '#A194B2', 'light': '#D1C8DA'},
    {'name': 'Mint Sorbet',   'dark': '#4A6B5E', 'medium': '#87A696', 'light': '#C2D6CC'},
    {'name': 'Sky Pastel',    'dark': '#4D6E7C', 'medium': '#89A9B8', 'light': '#C2D5DE'},
    {'name': 'Peach Fuzz',    'dark': '#8E6252', 'medium': '#C9907A', 'light': '#E8C5B9'},
    {'name': 'Soft Ochre',    'dark': '#7A6B46', 'medium': '#B3A175', 'light': '#DED3B6'},
    {'name': 'Powder Blue',   'dark': '#556677', 'medium': '#8B99A8', 'light': '#C5CDD6'},
    {'name': 'Rose Quartz',   'dark': '#8C5E66', 'medium': '#BF8E96', 'light': '#E8C8CD'},
    {'name': 'Olive Wash',    'dark': '#5B634D', 'medium': '#939B7E', 'light': '#C6CBB6'},
    {'name': 'Warm Stone',    'dark': '#5E5A54', 'medium': '#969188', 'light': '#CDC9C2'},
    {'name': 'Lilac Dream',   'dark': '#655470', 'medium': '#9684A1', 'light': '#C8BCCC'}
];
function recolorSidebarUI(container, palette) {
    if (!container) return;
    var allItems = container.pageItems;
    for (var i = 0; i < allItems.length; i++) {
        var it = allItems[i];
        try {
            // 1. טיפול בצורות (PathItems)
            if (it.typename === "PathItem") {
                // צביעת מילוי - רק אם הוא אחד מצבעי המקור
                if (it.filled && it.fillColor.typename === "RGBColor") {
                    var hF = rgbToHex(it.fillColor);
                    if (hF === "2C3E50") it.fillColor = hexToRgb(palette.dark);
                    else if (hF === "5D788F") it.fillColor = hexToRgb(palette.medium);
                    else if (hF === "87A0B5") it.fillColor = hexToRgb(palette.light);
                }
                // צביעת מיתר (Stroke) - רק אם הוא אחד מצבעי המקור
                if (it.stroked && it.strokeColor.typename === "RGBColor") {
                    var hS = rgbToHex(it.strokeColor);
                    if (hS === "2C3E50") it.strokeColor = hexToRgb(palette.dark);
                    else if (hS === "5D788F") it.strokeColor = hexToRgb(palette.medium);
                    else if (hS === "87A0B5") it.strokeColor = hexToRgb(palette.light);
                }
            }
            // 2. טיפול בטקסט (TextFrame)
            else if (it.typename === "TextFrame") {
                var attr = it.textRange.characterAttributes;
                // צביעת מילוי טקסט - רק אם הוא כהה במקור
                if (attr.fillColor.typename === "RGBColor") {
                    var hTF = rgbToHex(attr.fillColor);
                    if (hTF === "2C3E50") attr.fillColor = hexToRgb(palette.dark);
                    else if (hTF === "5D788F") attr.fillColor = hexToRgb(palette.medium);
                    else if (hTF === "87A0B5") attr.fillColor = hexToRgb(palette.light);
                }
                // צביעת מיתר טקסט - רק אם הוא קיים ואחד מצבעי המקור
                if (attr.strokeWeight > 0 && attr.strokeColor.typename === "RGBColor") {
                    var hTS = rgbToHex(attr.strokeColor);
                    if (hTS === "2C3E50") attr.strokeColor = hexToRgb(palette.dark);
                    else if (hTS === "5D788F") attr.strokeColor = hexToRgb(palette.medium);
                    else if (hTS === "87A0B5") attr.strokeColor = hexToRgb(palette.light);
                }
            }
            // כניסה לקבוצות
            else if (it.typename === "GroupItem") {
                recolorSidebarUI(it, palette);
            }
        } catch (e) { continue; }
    }
}
// פונקציה לבחירת מוצרים חכמה - מסננת את מה שכבר הוזמן
function getSmartUpsellItems(allOptions, ordered) {
    var availableOptions = [];
    $.writeln("--- Smart Upsell Check ---");
    $.writeln("Ordered by customer: " + ordered.join(", "));
    // מעבר על כל האופציות האפשריות לסרגל
    for (var i = 0; i < allOptions.length; i++) {
        var currentOption = allOptions[i];
        var alreadyOrdered = false;
        // בדיקה האם האופציה הזו נמצאת ברשימת המוצרים שהוזמנו
        for (var j = 0; j < ordered.length; j++) {
            if (currentOption === ordered[j]) {
                alreadyOrdered = true;
                break;
            }
        }
        // רק אם המוצר לא הוזמן, נוסיף אותו לרשימת האפשרויות הזמינות
        if (!alreadyOrdered) {
            availableOptions.push(currentOption);
        } else {
            $.writeln("!!! Filtered out (already in order): " + currentOption);
        }
    }
    // אם הלקוח הזמין הכל ואין מספיק אופציות (פחות מ-3), נחזור לרשימה המלאה
    if (availableOptions.length < 3) {
        $.writeln("Warning: Not enough options left after filtering. Using full list.");
        availableOptions = allOptions;
    }
    var selected = getRandomItems(availableOptions, 3);
    $.writeln("Final selected for sidebar: " + selected.join(", "));
    $.writeln("--------------------------");
    return selected;
}
function getRandomItems(arr, count) {
    var workingArr = arr.slice(0);
    var res = [];
    var i = workingArr.length;
    var t, idx;
    // ערבוב (Fisher-Yates Shuffle)
    while (i--) {
        idx = Math.floor((i + 1) * Math.random());
        t = workingArr[idx];
        workingArr[idx] = workingArr[i];
        workingArr[i] = t;
    }
    return workingArr.slice(0, count);
}
var sidebarTemplateDoc = null;
var sidebarTemplateFile = null;

function prepareSimulationSidebar() {
    var selected = upsellMode === "manual" && manualProducts.length > 0 ?
        manualProducts.slice(0, 3) : getSmartUpsellItems(allUpsellOptions, orderedProducts);
    $.writeln("Selected sidebar products: " + selected.join(", "));

    sidebarTemplateFile = new File(sidebarMasterPath);
    if (!sidebarTemplateFile.exists) {
        throw new Error("Sidebar template is missing: " + sidebarTemplateFile.fsName);
    }
    sidebarTemplateDoc = app.open(sidebarTemplateFile);
    sidebarTemplateDoc.activate();
    if (sidebarTemplateDoc.documentColorSpace !== DocumentColorSpace.RGB) {
        app.executeMenuCommand("doc-color-rgb");
    }

    var sidebarLayer = sidebarTemplateDoc.layers.getByName("Sidebar");
    var productsLayer = sidebarTemplateDoc.layers.getByName("Products");
    for (var p = 0; p < 3; p++) {
        var placeholder = sidebarLayer.pageItems.getByName("Product_" + (p + 1));
        if (p < selected.length) {
            var product = productsLayer.pageItems.getByName(selected[p]);
            var copy = product.duplicate(sidebarLayer, ElementPlacement.PLACEATBEGINNING);
            copy.left = placeholder.left;
            copy.top = placeholder.top;
        }
        placeholder.remove();
    }
    productsLayer.remove();
    var palette = allPalettes[Math.floor(Math.random() * allPalettes.length)];
    recolorSidebarUI(sidebarTemplateDoc, palette);

    var artwork = sidebarLayer.groupItems.add();
    for (var i = sidebarLayer.pageItems.length - 1; i >= 1; i--) {
        sidebarLayer.pageItems[i].move(artwork, ElementPlacement.PLACEATBEGINNING);
    }
    return artwork;
}

function closeSimulationSidebar() {
    if (!sidebarTemplateDoc) return;
    try {
        sidebarTemplateDoc.activate();
        var active = app.activeDocument;
        if (active.fullName.fsName.toLowerCase() !== sidebarTemplateFile.fsName.toLowerCase()) {
            throw new Error("Could not activate the sidebar template: " + sidebarTemplateFile.fsName);
        }
        active.close(SaveOptions.DONOTSAVECHANGES);
    } catch (error) {
        $.writeln("Error closing sidebar template: " + error);
    } finally {
        sidebarTemplateDoc = null;
        sidebarTemplateFile = null;
    }
}

function hasSidebarLayer(container) {
    for (var i = 0; i < container.layers.length; i++) {
        var layer = container.layers[i];
        if (layer.name === "Sidebar_Layer" || hasSidebarLayer(layer)) return true;
    }
    return false;
}

function addSimulationSidebar(doc, artwork) {
    if (!showSidebar || (job && job.sidebar_already_applied) || hasSidebarLayer(doc)) return false;

    var layer1 = doc.layers.getByName("1");
    if (!findLogoByPriority(layer1)) {
        $.writeln("Sidebar skipped: no simulation logo found");
        return false;
    }

    if (!artwork) {
        artwork = prepareSimulationSidebar();
        doc.activate();
    }
    var abRect = doc.artboards[0].artboardRect;
    var isDark = checkIsProductDark(layer1);
    resizeLayer1Aggressive(layer1, abRect);
    var logoData = findLogoByPriority(layer1);
    if (!logoData) throw new Error("Missing simulation logo after resizing layer 1");

    var sidebarLayer = doc.layers.add();
    sidebarLayer.name = "Sidebar_Layer";
    doc.activeLayer = sidebarLayer;
    var pastedBar = artwork.duplicate(sidebarLayer, ElementPlacement.PLACEATBEGINNING);
    pastedBar.left = abRect[0];
    pastedBar.top = abRect[1] - (Math.abs(abRect[1] - abRect[3]) - pastedBar.height) / 2;

    var masterLogo = sidebarLayer.groupItems.add();
    for (var i = 0; i < logoData.item.pageItems.length; i++) {
        logoData.item.pageItems[i].duplicate(masterLogo, ElementPlacement.PLACEATEND);
    }
    masterLogo.visible = false;
    for (var k = 0; k < pastedBar.groupItems.length; k++) {
        var prod = pastedBar.groupItems[k];
        if (isDark) {
            var prodSim = findRecursive(prod, "Simulation");
            if (prodSim) forceColorBlackRecursive(prodSim);
        }

        var box = selectAndCleanupBoxes(prod, logoData.type);
        if (box) {
            var logo = masterLogo.duplicate(prod, ElementPlacement.PLACEATBEGINNING);
            logo.visible = true;
            var scale = Math.min(box.width / logo.width, box.height / logo.height) * 100;
            logo.resize(scale, scale, true, true, true, true, scale);
            logo.left = box.left + (box.width - logo.width) / 2;
            logo.top = box.top - (box.height - logo.height) / 2;
        }
        removeBoxesRecursive(prod);
    }
    masterLogo.remove();
    ungroupSidebarItems(pastedBar);
    $.writeln("Sidebar added to " + doc.name);
    return true;
}

function ungroupSidebarItems(group) {
    while (group.pageItems.length > 0) {
        group.pageItems[0].move(group, ElementPlacement.PLACEBEFORE);
    }
    group.remove();
}

// --- פונקציות התיקון הקריטיות ---
function resizeWholeLayer1Aggressive(doc, layer, abRect) {
    var items = layer.pageItems;
    if (items.length === 0) return;
    // איסוף הכל לקבוצה אחת
    var tempGrp = doc.groupItems.add();
    for (var i = items.length - 1; i >= 0; i--) {
        if (items[i] != tempGrp) items[i].move(tempGrp, ElementPlacement.PLACEATEND);
    }
    // חישוב יחס הקטנה ל-23 ס"מ (651.9 נקודות)
    var targetWidth = 23 * 28.346;
    var ratio = (targetWidth / tempGrp.width) * 100;
    // פקודת הקטנה ישירה
    tempGrp.resize(ratio, ratio, true, true, true, true, ratio);
    // הצמדה לימין המשטח (abRect[2])
    tempGrp.left = abRect[2] - tempGrp.width;
    tempGrp.top = abRect[1] - (Math.abs(abRect[1] - abRect[3]) - tempGrp.height) / 2;
    // פירוק הקבוצה חזרה לשכבה
    tempGrp.selected = true;
    app.executeMenuCommand('ungroup');
}
function forceColorBlackRecursive(container) {
    var black = new RGBColor(); black.red=0; black.green=0; black.blue=0;
    var white = new RGBColor(); white.red=255; white.green=255; white.blue=255;
    // מוודא שהאובייקט קיים ולא ריק
    if (!container || !container.pageItems) return;
    var allItems = container.pageItems;
    for (var i = 0; i < allItems.length; i++) {
        var itm = allItems[i];
        try {
            // אם זו קבוצה, נצבע את התוכן שלה
            if (itm.typename === "GroupItem") {
                forceColorBlackRecursive(itm);
            }
            // צביעה בטוחה של נתיבים
            else if (itm.typename === "PathItem") {
                itm.fillColor = black;
                itm.strokeColor = white;
                itm.stroked = true;
                itm.strokeWidth = 0.5;
            }
            // טיפול ב-Compound Paths (לפעמים הם גורמים לקריסה)
            else if (itm.typename === "CompoundPathItem") {
                for (var j = 0; j < itm.pathItems.length; j++) {
                    itm.pathItems[j].fillColor = black;
                    itm.pathItems[j].strokeColor = white;
                }
            }
        } catch (e) {
            // אם אובייקט ספציפי בעייתי, פשוט נדלג עליו במקום להקריס
            continue;
        }
    }
}
function checkIsProductDark(layer1) {
    var totalLum = 0;
    var count = 0;
    try {
        // מציאת הסימולציה בנתיב המדויק
        var sim = findRecursive(layer1, "Simulation");
        var model = findRecursive(sim, "Simulation") || sim;
        function scan(obj) {
            if (obj.typename === "PathItem" && obj.filled) {
                var lum = getLumFromColor(obj.fillColor);
                if (lum !== null) {
                    totalLum += lum;
                    count++;
                }
            } else if (obj.typename === "GroupItem") {
                for (var i = 0; i < obj.pageItems.length; i++) scan(obj.pageItems[i]);
            } else if (obj.typename === "CompoundPathItem") {
                for (var j = 0; j < obj.pathItems.length; j++) scan(obj.pathItems[j]);
            }
        }
        scan(model);
        if (count === 0) {
            return false;
        }
        var avg = totalLum / count;
        var isDark = (avg < 128);
        return isDark;
    } catch(e) {
        return false;
    }
}
// --- פונקציות תשתית (ללא שינוי) ---
function findRecursive(container, name) {
    if (container.typename === "Layer") {
        try { return container.layers.getByName(name); } catch(e) {}
        for (var l=0; l<container.layers.length; l++) {
            var resL = findRecursive(container.layers[l], name);
            if (resL) return resL;
        }
    }
    for (var i=0; i<container.pageItems.length; i++) {
        var itm = container.pageItems[i];
        if (itm.name.indexOf(name) !== -1) return itm;
        if (itm.typename === "GroupItem") {
            var resG = findRecursive(itm, name);
            if (resG) return resG;
        }
    }
    return null;
}
function selectAndCleanupBoxes(prod, logoType) {
    var boxes = [];
    function scan(obj) {
        for (var i=0; i<obj.pageItems.length; i++) {
            if (obj.pageItems[i].name.indexOf("Box") === 0) boxes.push(obj.pageItems[i]);
            if (obj.pageItems[i].typename === "GroupItem") scan(obj.pageItems[i]);
        }
    }
    scan(prod);
    if (boxes.length === 0) return null;
    var selected = boxes[0];
    if (logoType.indexOf("Front") !== -1 || logoType.indexOf("Sleeve") !== -1) {
        for (var j=0; j<boxes.length; j++) if (boxes[j].name.indexOf("Pocket") !== -1) selected = boxes[j];
    }
    return selected;
}
function removeBoxesRecursive(container) {
    for (var i = container.pageItems.length - 1; i >= 0; i--) {
        var itm = container.pageItems[i];
        if (itm.name.indexOf("Box") === 0) itm.remove();
        else if (itm.typename === "GroupItem") removeBoxesRecursive(itm);
    }
}
function findLogoByPriority(layer1) {
    var names = ["S_Placement_Front", "S_Placement_Back", "S_Placement_Left_Sleeve", "S_Placement_Right_Sleeve"];
    for (var i = 0; i < names.length; i++) {
        var found = findRecursive(layer1, names[i]);
        if (found && found.pageItems && found.pageItems.length > 0) return {item: found, type: names[i]};
    }
    return null;
}
function rgbToHex(c) {
    var r = Math.round(c.red).toString(16).toUpperCase(); if (r.length == 1) r = "0" + r;
    var g = Math.round(c.green).toString(16).toUpperCase(); if (g.length == 1) g = "0" + g;
    var b = Math.round(c.blue).toString(16).toUpperCase(); if (b.length == 1) b = "0" + b;
    return r + g + b;
}
function hexToRgb(h) {
    var s = h.replace("#",""); var c = new RGBColor();
    c.red = parseInt(s.substring(0,2), 16); c.green = parseInt(s.substring(2,4), 16); c.blue = parseInt(s.substring(4,6), 16);
    return c;
}
function resizeLayer1Aggressive(mainLayer, abRect) {
    try {
        var simContainer = mainLayer.layers.getByName("Simulation");
        // שלב א: הפיכת כל תתי-השכבות בתוך Simulation לקבוצות
        for (var i = simContainer.layers.length - 1; i >= 0; i--) {
            var subLayer = simContainer.layers[i];
            var newGrp = simContainer.groupItems.add();
            newGrp.name = subLayer.name;
            // העברת פריטים מהשכבה לקבוצה החדשה
            for (var j = subLayer.pageItems.length - 1; j >= 0; j--) {
                subLayer.pageItems[j].move(newGrp, ElementPlacement.PLACEATBEGINNING);
            }
            subLayer.remove(); // מחיקת השכבה הריקה
        }
        // שלב ב: איגוד כל התוכן של Simulation לקבוצה אחת לצורך הקטנה
        var items = simContainer.pageItems;
        var masterGrp = simContainer.groupItems.add();
        for (var k = items.length - 1; k >= 0; k--) {
            if (items[k] != masterGrp) items[k].move(masterGrp, ElementPlacement.PLACEATBEGINNING);
        }
        // שלב ג: הקטנה ל-23 ס"מ והצמדה לימין
        var targetWidth = 23 * 28.346;
        var ratio = (targetWidth / masterGrp.width) * 100;
        masterGrp.resize(ratio, ratio, true, true, true, true, ratio);
        masterGrp.left = abRect[2] - masterGrp.width;
        masterGrp.top = abRect[1] - (Math.abs(abRect[1] - abRect[3]) - masterGrp.height) / 2;
        // פירוק קבוצת ה-Master (משאיר את תתי-הקבוצות בשם המקורי)
        ungroupSidebarItems(masterGrp);
    } catch (e) {
        // אם Simulation היא קבוצה ולא שכבה
        try {
            var simGrp = mainLayer.groupItems.getByName("Simulation");
            var ratio = (23 * 28.346 / simGrp.width) * 100;
            simGrp.resize(ratio, ratio, true, true, true, true, ratio);
            simGrp.left = abRect[2] - simGrp.width;
            simGrp.top = abRect[1] - (Math.abs(abRect[1] - abRect[3]) - simGrp.height) / 2;
        } catch(err) { $.writeln("Error resizing: " + err); }
    }
}
function findFlex(parent, name) {
    if (parent.layers) {
        try { return parent.layers.getByName(name); } catch(e) {}
    }
    if (parent.pageItems) {
        for (var i = 0; i < parent.pageItems.length; i++) {
            if (parent.pageItems[i].name.indexOf(name) !== -1) return parent.pageItems[i];
        }
    }
    return null;
}
function getLumFromColor(color) {
    try {
        if (color.typename === "RGBColor") return (color.red * 0.299) + (color.green * 0.587) + (color.blue * 0.114);
        if (color.typename === "GrayColor") return 255 - (color.gray * 2.55);
        if (color.typename === "CMYKColor") {
            var r = 255 * (1 - color.cyan/100) * (1 - color.black/100);
            var g = 255 * (1 - color.magenta/100) * (1 - color.black/100);
            var b = 255 * (1 - color.yellow/100) * (1 - color.black/100);
            return (r * 0.299) + (g * 0.587) + (b * 0.114);
        }
    } catch(e) {}
    return null;
}
if (typeof sidebarJobData === "undefined" && app.documents.length > 0) {
    var standaloneDoc = typeof sidebarTargetDoc === "undefined" ? app.activeDocument : sidebarTargetDoc;
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
    try {
        addSimulationSidebar(standaloneDoc, null);
    } catch (error) {
        $.writeln("Error adding sidebar: " + error);
    } finally {
        closeSimulationSidebar();
        standaloneDoc.activate();
    }
}
