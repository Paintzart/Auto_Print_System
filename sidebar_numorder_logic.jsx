#target illustrator
/**
 * מקטין את הדף הראשון ומציב Sidebar_NumOrder.ai עם מספר הזמנה.
 *
 * מצב "side" (הדמיה אחת / לרוחב):
 *   Artboard 1, שכבת Sidebar, NUM1..NUM4 (ספרה לכל תיבה),
 *   רוחב 23 ס"מ, הצמדה לימין.
 *
 * מצב "bottom" (2+ הדמיות בעמוד):
 *   Artboard 2, שכבת sidebar2, NUM (כל המספר),
 *   גובה 24 ס"מ, הצמדה לתחתית.
 *
 * קלט:
 *   $.global.targetMasterDoc, numOrderLast4, numOrderSidebarPath,
 *   numOrderMode ("side"|"bottom"), showNumOrderSidebar
 * או current_job.json
 */
function getNumOrderJobData() {
    var scriptFile = new File($.fileName);
    var dataFile = new File(scriptFile.path + "/current_job.json");
    if (!dataFile.exists) return null;
    dataFile.open("r");
    var content = dataFile.read();
    dataFile.close();
    return eval("(" + content + ")");
}

function findTextFrameByName(container, name) {
    if (!container) return null;
    try {
        if (container.textFrames) {
            for (var i = 0; i < container.textFrames.length; i++) {
                if (container.textFrames[i].name === name) return container.textFrames[i];
            }
        }
    } catch (e) {}
    try {
        if (container.pageItems) {
            for (var j = 0; j < container.pageItems.length; j++) {
                var it = container.pageItems[j];
                if (it.typename === "TextFrame" && it.name === name) return it;
                if (it.typename === "GroupItem") {
                    var found = findTextFrameByName(it, name);
                    if (found) return found;
                }
            }
        }
    } catch (e2) {}
    try {
        if (container.layers) {
            for (var k = 0; k < container.layers.length; k++) {
                var foundL = findTextFrameByName(container.layers[k], name);
                if (foundL) return foundL;
            }
        }
    } catch (e3) {}
    return null;
}

function normalizeLast4(digits) {
    var d = String(digits || "");
    d = d.replace(/[^0-9]/g, "");
    while (d.length < 4) d = "0" + d;
    return d.slice(-4);
}

function setNumDigitsSide(container, digits) {
    var d = normalizeLast4(digits);
    for (var i = 1; i <= 4; i++) {
        var tf = findTextFrameByName(container, "NUM" + i);
        if (tf) {
            try { tf.contents = d.charAt(i - 1); } catch (e) {}
        } else {
            $.writeln("NUM" + i + " not found in sidebar");
        }
    }
    return d;
}

function setNumWhole(container, digits) {
    var d = normalizeLast4(digits);
    var tf = findTextFrameByName(container, "NUM");
    if (tf) {
        try { tf.contents = d; } catch (e) {}
    } else {
        $.writeln("NUM not found in sidebar2");
    }
    return d;
}

function docHasNumOrderSidebar(doc) {
    function scan(container) {
        if (!container || !container.layers) return false;
        for (var i = 0; i < container.layers.length; i++) {
            if (container.layers[i].name === "Sidebar_NumOrder_Layer") return true;
            if (scan(container.layers[i])) return true;
        }
        return false;
    }
    return scan(doc);
}

function getLayerByNames(doc, names) {
    for (var i = 0; i < names.length; i++) {
        try { return doc.layers.getByName(names[i]); } catch (e) {}
    }
    return null;
}

function resizeLayer1Aggressive(doc, mainLayer, abRect) {
    try {
        var simContainer = mainLayer.layers.getByName("Simulation");
        for (var i = simContainer.layers.length - 1; i >= 0; i--) {
            var subLayer = simContainer.layers[i];
            var newGrp = simContainer.groupItems.add();
            newGrp.name = subLayer.name;
            for (var j = subLayer.pageItems.length - 1; j >= 0; j--) {
                subLayer.pageItems[j].move(newGrp, ElementPlacement.PLACEATBEGINNING);
            }
            subLayer.remove();
        }
        var items = simContainer.pageItems;
        var masterGrp = simContainer.groupItems.add();
        for (var k = items.length - 1; k >= 0; k--) {
            if (items[k] != masterGrp) items[k].move(masterGrp, ElementPlacement.PLACEATBEGINNING);
        }
        app.redraw();
        var targetWidth = 23 * 28.346;
        var ratio = (targetWidth / masterGrp.width) * 100;
        masterGrp.resize(ratio, ratio, true, true, true, true, ratio);
        masterGrp.left = abRect[2] - masterGrp.width;
        masterGrp.top = abRect[1] - (Math.abs(abRect[1] - abRect[3]) - masterGrp.height) / 2;
        masterGrp.selected = true;
        app.executeMenuCommand('ungroup');
    } catch (e) {
        try {
            var simGrp = mainLayer.groupItems.getByName("Simulation");
            var ratio2 = (23 * 28.346 / simGrp.width) * 100;
            simGrp.resize(ratio2, ratio2, true, true, true, true, ratio2);
            simGrp.left = abRect[2] - simGrp.width;
            simGrp.top = abRect[1] - (Math.abs(abRect[1] - abRect[3]) - simGrp.height) / 2;
        } catch (err) {
            $.writeln("Error resizing Simulation: " + err);
        }
    }
}

function itemIntersectsArtboard(item, abRect) {
    try {
        var b = item.visibleBounds;
        return !(b[0] > abRect[2] || b[2] < abRect[0] || b[1] < abRect[3] || b[3] > abRect[1]);
    } catch (e) {
        return false;
    }
}

function collectFirstArtboardItems(doc, abRect) {
    var toMove = [];
    var layer = doc.activeLayer;
    for (var i = 0; i < layer.pageItems.length; i++) {
        var it = layer.pageItems[i];
        if (itemIntersectsArtboard(it, abRect)) toMove.push(it);
    }
    if (toMove.length === 0) {
        for (var li = 0; li < doc.layers.length; li++) {
            var lay = doc.layers[li];
            if (lay.name === "Sidebar_NumOrder_Layer" || lay.name === "Sidebar_Layer") continue;
            for (var pi = 0; pi < lay.pageItems.length; pi++) {
                var pit = lay.pageItems[pi];
                if (itemIntersectsArtboard(pit, abRect)) toMove.push(pit);
            }
        }
    }
    return toMove;
}

function resizeFirstArtboardSide(doc, abRect) {
    // הדמיה אחת: שכבה 1 / Simulation → רוחב 23 ס"מ לימין
    try {
        var layer1 = doc.layers.getByName("1");
        resizeLayer1Aggressive(doc, layer1, abRect);
        return;
    } catch (e) {}

    try {
        var toMove = collectFirstArtboardItems(doc, abRect);
        if (toMove.length === 0) {
            $.writeln("No content found to resize on artboard 0");
            return;
        }
        var masterGrp = doc.groupItems.add();
        for (var m = toMove.length - 1; m >= 0; m--) {
            try { toMove[m].move(masterGrp, ElementPlacement.PLACEATBEGINNING); } catch (e2) {}
        }
        app.redraw();
        if (masterGrp.width > 0) {
            var ratio = ((23 * 28.346) / masterGrp.width) * 100;
            masterGrp.resize(ratio, ratio, true, true, true, true, ratio);
            masterGrp.left = abRect[2] - masterGrp.width;
            masterGrp.top = abRect[1] - (Math.abs(abRect[1] - abRect[3]) - masterGrp.height) / 2;
        }
        masterGrp.selected = true;
        app.executeMenuCommand('ungroup');
    } catch (err) {
        $.writeln("Error resizing first artboard (side): " + err);
    }
}

function resizeFirstArtboardBottom(doc, abRect) {
    // 2+ הדמיות: גובה 24 ס"מ, צמוד לתחתית
    try {
        var toMove = collectFirstArtboardItems(doc, abRect);
        if (toMove.length === 0) {
            $.writeln("No content found to resize on artboard 0 (bottom mode)");
            return;
        }
        var masterGrp = doc.groupItems.add();
        for (var m = toMove.length - 1; m >= 0; m--) {
            try { toMove[m].move(masterGrp, ElementPlacement.PLACEATBEGINNING); } catch (e2) {}
        }
        app.redraw();
        if (masterGrp.height > 0) {
            var targetHeight = 24 * 28.346;
            var ratio = (targetHeight / masterGrp.height) * 100;
            masterGrp.resize(ratio, ratio, true, true, true, true, ratio);
            // מרכוז אופקי + הצמדה לתחתית הארטבורד
            var pageW = Math.abs(abRect[2] - abRect[0]);
            masterGrp.left = abRect[0] + (pageW - masterGrp.width) / 2;
            masterGrp.top = abRect[3] + masterGrp.height;
        }
        masterGrp.selected = true;
        app.executeMenuCommand('ungroup');
    } catch (err) {
        $.writeln("Error resizing first artboard (bottom): " + err);
    }
}

function groupLayerItemsOnArtboard(layer, srcAbRect) {
    var finalBar = layer.groupItems.add();
    for (var m = layer.pageItems.length - 1; m >= 1; m--) {
        var it = layer.pageItems[m];
        if (it === finalBar) continue;
        // רק פריטים שנמצאים על הארטבורד הרלוונטי
        if (srcAbRect && !itemIntersectsArtboard(it, srcAbRect)) continue;
        try { it.move(finalBar, ElementPlacement.PLACEATBEGINNING); } catch (e) {}
    }
    return finalBar;
}

function mainNumOrder() {
    $.writeln("=== SIDEBAR NUMORDER START ===");
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    var job = getNumOrderJobData();
    var show = true;
    if (typeof $.global.showNumOrderSidebar !== "undefined") {
        show = !!$.global.showNumOrderSidebar;
    } else if (job && typeof job.show_numorder_sidebar !== "undefined") {
        show = !!job.show_numorder_sidebar;
    }
    if (!show) {
        $.writeln("NumOrder sidebar disabled.");
        return;
    }

    var last4 = "";
    if (typeof $.global.numOrderLast4 !== "undefined" && $.global.numOrderLast4) {
        last4 = String($.global.numOrderLast4);
    } else if (job && job.order_last4) {
        last4 = String(job.order_last4);
    }
    if (!last4) {
        $.writeln("ERROR: No order_last4 provided");
        return;
    }

    var sidebarPath = "";
    if (typeof $.global.numOrderSidebarPath !== "undefined" && $.global.numOrderSidebarPath) {
        sidebarPath = String($.global.numOrderSidebarPath);
    } else if (job && job.sidebar_numorder_path) {
        sidebarPath = String(job.sidebar_numorder_path);
    }
    if (!sidebarPath) {
        $.writeln("ERROR: No Sidebar_NumOrder path");
        return;
    }

    var mode = "side";
    if (typeof $.global.numOrderMode !== "undefined" && $.global.numOrderMode) {
        mode = String($.global.numOrderMode);
    } else if (job && job.numorder_mode) {
        mode = String(job.numorder_mode);
    }
    if (mode !== "bottom") mode = "side";
    $.writeln("NumOrder mode: " + mode);

    if (app.documents.length === 0) {
        $.writeln("ERROR: No documents open");
        return;
    }

    var pdfDoc = null;
    if (typeof $.global.targetMasterDoc !== "undefined" && $.global.targetMasterDoc !== null) {
        pdfDoc = $.global.targetMasterDoc;
        pdfDoc.activate();
        app.activeDocument = pdfDoc;
    } else {
        pdfDoc = app.activeDocument;
    }

    if (docHasNumOrderSidebar(pdfDoc)) {
        $.writeln("Sidebar_NumOrder_Layer already exists — skip NumOrder.");
        return;
    }

    var abRect = pdfDoc.artboards[0].artboardRect;
    try {
        $.writeln("--- Resize first page ---");
        if (mode === "bottom") {
            resizeFirstArtboardBottom(pdfDoc, abRect);
        } else {
            resizeFirstArtboardSide(pdfDoc, abRect);
        }

        $.writeln("--- Open Sidebar_NumOrder ---");
        $.writeln("Path: " + sidebarPath);
        var sbFile = new File(sidebarPath);
        if (!sbFile.exists) {
            $.writeln("ERROR: Sidebar_NumOrder.ai not found");
            return;
        }
        var sbDoc = app.open(sbFile);
        try { app.executeMenuCommand('doc-color-rgb'); } catch (eRgb) {}

        // מצב bottom → Artboard 2 + sidebar2 ; מצב side → Artboard 1 + Sidebar
        var sbLayer = null;
        var srcAbIndex = 0;
        if (mode === "bottom") {
            srcAbIndex = (sbDoc.artboards.length > 1) ? 1 : 0;
            try { sbDoc.artboards.setActiveArtboardIndex(srcAbIndex); } catch (eAb) {}
            sbLayer = getLayerByNames(sbDoc, ["sidebar2", "Sidebar2", "Sidebar_2"]);
            if (!sbLayer) {
                $.writeln("sidebar2 layer not found — fallback to activeLayer");
                sbLayer = sbDoc.activeLayer;
            }
            setNumWhole(sbDoc, last4);
            setNumWhole(sbLayer, last4);
        } else {
            srcAbIndex = 0;
            try { sbDoc.artboards.setActiveArtboardIndex(0); } catch (eAb2) {}
            sbLayer = getLayerByNames(sbDoc, ["Sidebar", "sidebar"]);
            if (!sbLayer) {
                sbLayer = sbDoc.activeLayer;
                $.writeln("Using activeLayer as sidebar source: " + sbLayer.name);
            }
            setNumDigitsSide(sbDoc, last4);
            setNumDigitsSide(sbLayer, last4);
        }

        var srcAbRect = sbDoc.artboards[srcAbIndex].artboardRect;
        $.writeln("Source artboard " + (srcAbIndex + 1) + ": " + srcAbRect);

        var finalBar = groupLayerItemsOnArtboard(sbLayer, srcAbRect);
        if (!finalBar || finalBar.pageItems.length === 0) {
            $.writeln("ERROR: No sidebar items found on source artboard");
            sbDoc.close(SaveOptions.DONOTSAVECHANGES);
            return;
        }

        // שמירת המיקום היחסי לארטבורד המקור לפני ההעתקה
        var srcRelLeft = finalBar.left - srcAbRect[0];
        var srcRelTop = finalBar.top - srcAbRect[1];
        $.writeln("Source relative pos: left=" + srcRelLeft + " top=" + srcRelTop);

        pdfDoc.activate();
        app.activeDocument = pdfDoc;

        var sidebarFinalLayer;
        try {
            sidebarFinalLayer = pdfDoc.layers.getByName("Sidebar_NumOrder_Layer");
        } catch (eL) {
            sidebarFinalLayer = pdfDoc.layers.add();
            sidebarFinalLayer.name = "Sidebar_NumOrder_Layer";
        }
        pdfDoc.activeLayer = sidebarFinalLayer;

        var pastedBar = finalBar.duplicate(sidebarFinalLayer, ElementPlacement.PLACEATBEGINNING);
        sbDoc.close(SaveOptions.DONOTSAVECHANGES);

        pdfDoc.activate();
        app.activeDocument = pdfDoc;

        // מיקום זהה ליחס לארטבורד כמו בקובץ Sidebar_NumOrder.ai
        // (לא דריסה לשמאל/למעלה של העמוד — שומר את ההצמדה מהעיצוב)
        pastedBar.left = abRect[0] + srcRelLeft;
        pastedBar.top = abRect[1] + srcRelTop;
        $.writeln("Placed on target: left=" + pastedBar.left + " top=" + pastedBar.top);

        pastedBar.selected = true;
        try { app.executeMenuCommand('ungroup'); } catch (eU) {}

        $.writeln("=== SIDEBAR NUMORDER COMPLETE ===");
    } catch (e) {
        $.writeln("=== ERROR IN NUMORDER SIDEBAR ===");
        $.writeln(e.toString());
        try {
            if (typeof $.global.targetMasterDoc !== "undefined" && $.global.targetMasterDoc !== null) {
                $.global.targetMasterDoc.activate();
                app.activeDocument = $.global.targetMasterDoc;
            }
        } catch (err) {}
    }
}

mainNumOrder();
