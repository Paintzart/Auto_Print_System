
    #target illustrator
    var files = ["C:/Users/yarde/OneDrive/Desktop/Auto_Print_System/temp_ai_files/temp_0.ai", "C:/Users/yarde/OneDrive/Desktop/Auto_Print_System/temp_ai_files/temp_1.ai"];
    
    function main() {
        if (files.length === 0) return;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        
        var maxWidth = 0; var maxHeight = 0;
        for (var i = 0; i < files.length; i++) {
            var tempDoc = app.open(new File(files[i]));
            var m = calculateLayoutMetrics(tempDoc);
            if (m.width > maxWidth) maxWidth = m.width;
            if (m.height > maxHeight) maxHeight = m.height;
            tempDoc.close(SaveOptions.DONOTSAVECHANGES);
        }
        
        var STEP_X = maxWidth + 150; 
        var STEP_Y = maxHeight + 250;
        var COLS = 4;

        var masterDoc = app.open(new File(files[0]));
        organizeMasterContent(masterDoc);
        
        for (var j = 1; j < files.length; j++) {
            var col = j % COLS;
            var row = Math.floor(j / COLS);
            processNextFileFast(masterDoc, files[j], (j+1).toString(), col * STEP_X, -(row * STEP_Y));
        }
        
        var sideFile = new File("C:/Users/yarde/OneDrive/Desktop/Auto_Print_System/sidebar_logic.jsx");
        var showSidebar = false;
        if (showSidebar && sideFile.exists) { 
            $.evalFile(sideFile); 
        }
        reorderArtboardsSafe(masterDoc);
    }
    
    // --- שאר הפונקציות (calculateLayoutMetrics, organizeMasterContent וכו') נשארות ללא שינוי ---
    function calculateLayoutMetrics(doc) {
        var minX = Infinity; var maxX = -Infinity; var maxY = -Infinity; var minY = Infinity;  
        for (var i = 0; i < doc.artboards.length; i++) {
            var r = doc.artboards[i].artboardRect; 
            if (r[0] < minX) minX = r[0]; if (r[2] > maxX) maxX = r[2];
            if (r[1] > maxY) maxY = r[1]; if (r[3] < minY) minY = r[3]; 
        }
        return { width: Math.abs(maxX - minX), height: Math.abs(maxY - minY) };
    }
    
    function organizeMasterContent(doc) {
        app.executeMenuCommand('unlockAll'); app.executeMenuCommand('showAll');
        var l1 = doc.layers.add(); l1.name = "1";
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var lay = doc.layers[i];
            if (lay != l1) lay.move(l1, ElementPlacement.PLACEATEND);
        }
    }
    
    function fastCopyLayer(srcLayer, destLayer, offX, offY) {
        if (srcLayer.pageItems.length > 0) {
            var tempGrp = srcLayer.groupItems.add();
            for (var i = srcLayer.pageItems.length - 1; i >= 0; i--) {
                if (srcLayer.pageItems[i] != tempGrp) {
                    srcLayer.pageItems[i].move(tempGrp, ElementPlacement.PLACEATEND);
                }
            }
            try {
                var dup = tempGrp.duplicate(destLayer, ElementPlacement.PLACEATBEGINNING);
                dup.translate(offX, offY);
                while (dup.pageItems.length > 0) {
                    dup.pageItems[0].move(destLayer, ElementPlacement.PLACEATBEGINNING);
                }
                dup.remove();
                while (tempGrp.pageItems.length > 0) {
                    tempGrp.pageItems[0].move(srcLayer, ElementPlacement.PLACEATBEGINNING);
                }
                tempGrp.remove();
            } catch(e) {}
        }
        for (var j = 0; j < srcLayer.layers.length; j++) {
            var sSub = srcLayer.layers[j];
            var dSub = destLayer.layers.add();
            dSub.name = sSub.name;
            fastCopyLayer(sSub, dSub, offX, offY);
        }
    }
    
    function processNextFileFast(masterDoc, srcPath, layerName, offX, offY) {
        var srcDoc = app.open(new File(srcPath));
        var abData = [];
        for(var i=0; i<srcDoc.artboards.length; i++) abData.push({rect: srcDoc.artboards[i].artboardRect, name: srcDoc.artboards[i].name});
        masterDoc.activate();
        var mainLayer = masterDoc.layers.add(); mainLayer.name = layerName;
        for (var k = 0; k < srcDoc.layers.length; k++) {
            var sLay = srcDoc.layers[k];
            var dLay = mainLayer.layers.add();
            dLay.name = sLay.name;
            fastCopyLayer(sLay, dLay, offX, offY);
        }
        srcDoc.close(SaveOptions.DONOTSAVECHANGES);
        masterDoc.activate();
        for(var n=0; n<abData.length; n++){
            var d = abData[n];
            var newAb = masterDoc.artboards.add([d.rect[0]+offX, d.rect[1]+offY, d.rect[2]+offX, d.rect[3]+offY]);
            newAb.name = "P" + layerName + "_" + d.name;
        }
    }
    
    function reorderArtboardsSafe(doc) {
        var oldAbs = [];
        for (var i = 0; i < doc.artboards.length; i++) oldAbs.push({rect: doc.artboards[i].artboardRect, name: doc.artboards[i].name});
        var newOrder = [];
        for (var i = 0; i < oldAbs.length; i++) if (oldAbs[i].name.indexOf("Simulation") > -1) newOrder.push(oldAbs[i]);
        for (var i = 0; i < oldAbs.length; i++) if (oldAbs[i].name.indexOf("Simulation") === -1) newOrder.push(oldAbs[i]);
        for (var j = 0; j < newOrder.length; j++) doc.artboards.add(newOrder[j].rect).name = newOrder[j].name;
        var len = oldAbs.length;
        for (var k = 0; k < len; k++) doc.artboards[0].remove();
    }
    
    main();
    