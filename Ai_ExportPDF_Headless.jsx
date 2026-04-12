// ============================================================
//                    CACHE SYSTEM
// ============================================================
var _objectCache = { byName: {}, textFrames: {}, built: false };

function buildCache(doc) {
    if (_objectCache.built) return;
    _objectCache.byName = {};
    _objectCache.textFrames = {};
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name.toLowerCase().indexOf("template") >= 0)
            walkLayerForCache(doc.layers[i]);
    }
    _objectCache.built = true;
}

function walkLayerForCache(layer) {
    for (var i = 0; i < layer.pageItems.length; i++)
        cacheItem(layer.pageItems[i]);
    for (var j = 0; j < layer.layers.length; j++)
        walkLayerForCache(layer.layers[j]);
}

function cacheItem(item) {
    if (item.name && item.name !== "") {
        if (!_objectCache.byName[item.name]) _objectCache.byName[item.name] = [];
        _objectCache.byName[item.name].push(item);
    }
    if (item.typename === "GroupItem") {
        for (var i = 0; i < item.pageItems.length; i++)
            cacheItem(item.pageItems[i]);
    }
    // TH1: TextFrame có name → cache theo name (bất kể contents)
    // TH2: TextFrame có pattern *x* trong contents → cache theo x
    if (item.typename === "TextFrame") {
        if (item.name && item.name !== "") {
            if (!_objectCache.textFrames[item.name]) _objectCache.textFrames[item.name] = [];
            _objectCache.textFrames[item.name].push(item);
        }
        // Tìm tất cả pattern *x* trong contents
        var contents = item.contents || "";
        var re = /\*([^*]+)\*/g;
        var match;
        while ((match = re.exec(contents)) !== null) {
            var key = match[1];
            if (!_objectCache.textFrames[key]) _objectCache.textFrames[key] = [];
            // Tránh duplicate
            var dup = false;
            for (var d = 0; d < _objectCache.textFrames[key].length; d++)
                if (_objectCache.textFrames[key][d] === item) { dup = true; break; }
            if (!dup) _objectCache.textFrames[key].push(item);
        }
    }
}

function findObjectsByName(name, doc) {
    if (!_objectCache.built) buildCache(doc);
    return _objectCache.byName[name] || [];
}

function findTextFramesByName(name, doc) {
    if (!_objectCache.built) buildCache(doc);
    return _objectCache.textFrames[name] || [];
}

// ============================================================
//                    LAYER HELPERS
// ============================================================
function findTemplateLayerName(obj) {
    var current = obj.parent;
    while (current) {
        if (current.typename === "Layer" && current.name.toLowerCase().indexOf("template") >= 0)
            return current.name;
        current = current.parent;
    }
    return "";
}

function createTempLayerAboveTemplate(doc, templateLayerName, tempLayerName) {
    try {
        var templateLayer = doc.layers.getByName(templateLayerName);
        
        // ✅ Lưu trạng thái gốc
        var wasLocked  = templateLayer.locked;
        var wasVisible = templateLayer.visible;
        
        // ✅ Bật lên và mở khóa trước khi move
        if (!wasVisible) templateLayer.visible = true;
        if (wasLocked)   templateLayer.locked  = false;
        
        var newLayer = doc.layers.add();
        newLayer.name = tempLayerName;
        newLayer.move(templateLayer, ElementPlacement.PLACEBEFORE);
        
        // ✅ Khôi phục trạng thái gốc
        if (wasLocked)    templateLayer.locked  = true;
        if (!wasVisible)  templateLayer.visible = false;
        
        return newLayer;
    } catch(e) {
        var newLayer = doc.layers.add();
        newLayer.name = tempLayerName;
        return newLayer;
    }
}

function removeAllTempLayers(tempLayersMap) {
    for (var key in tempLayersMap)
        if (tempLayersMap.hasOwnProperty(key))
            try { tempLayersMap[key].remove(); } catch(e) {}
}

function removeOldPreviewLayers(doc) {
    var toRemove = [];
    for (var i = 0; i < doc.layers.length; i++)
        if (doc.layers[i].name.indexOf("_PREVIEW") >= 0)
            toRemove.push(doc.layers[i]);
    for (var j = 0; j < toRemove.length; j++)
        try { toRemove[j].remove(); } catch(e) {}
}

function getOrCreateTempLayer(doc, templateLayerName, tempLayersMap, timestamp, isPreview) {
    if (tempLayersMap[templateLayerName]) return tempLayersMap[templateLayerName];
    var tempName = isPreview ? templateLayerName + "_PREVIEW" : templateLayerName + "_Temp_" + timestamp;
    var newLayer = createTempLayerAboveTemplate(doc, templateLayerName, tempName);
    tempLayersMap[templateLayerName] = newLayer;
    return newLayer;
}

function getTopGroupInLayer(item) {
    var current = item;
    var topGroup = null;
    while (current.parent) {
        if (current.parent.typename === "Layer") {
            if (current.typename === "GroupItem") topGroup = current;
            break;
        }
        current = current.parent;
    }
    return topGroup;
}

// ============================================================
//                    TEXT REPLACEMENT HELPERS
// ============================================================
function groupReplacementsByAncestor(textReplacements, doc) {
    var allItems = [];
    for (var t = 0; t < textReplacements.length; t++) {
        var rep = textReplacements[t];
        var textObjects = findTextFramesByName(rep.objectName, doc);
        for (var i = 0; i < textObjects.length; i++)
            allItems.push({ textObject: textObjects[i], replacement: rep });
    }
    if (allItems.length === 0) return [];

    var groups = [];
    for (var i = 0; i < allItems.length; i++) {
        var item = allItems[i];
        var topGroup = getTopGroupInLayer(item.textObject);
        var container = topGroup || item.textObject;
        var found = false;
        for (var g = 0; g < groups.length; g++) {
            if (groups[g].container === container) {
                var repExists = false;
                for (var r = 0; r < groups[g].replacements.length; r++)
                    if (groups[g].replacements[r].objectName === item.replacement.objectName)
                        { repExists = true; break; }
                if (!repExists) groups[g].replacements.push(item.replacement);
                found = true;
                break;
            }
        }
        if (!found)
            groups.push({ container: container, isGroup: topGroup !== null, replacements: [item.replacement] });
    }
    return groups;
}

function applyTextReplacementsToContainer(container, replacements) {
    for (var i = 0; i < replacements.length; i++) {
        var rep = replacements[i];
        var pattern = '*' + rep.objectName + '*';
        // TH1: Tìm theo tên textframe
        var namedFrames = findTextFramesInContainer(container, rep.objectName);
        if (namedFrames.length > 0) {
            for (var j = 0; j < namedFrames.length; j++) {
                namedFrames[j].contents = rep.newContent;
                var sizeObj = findSizeObjectInContainer(container);
                if (sizeObj) adjustTextToFitWidth(namedFrames[j], sizeObj.width);
            }
        }
        // TH2: Tìm pattern *x* trong contents
        applyPatternReplacement(container, pattern, rep.newContent, container);
    }
}

function applyPatternReplacement(container, pattern, newContent, rootContainer) {
    if (container.typename !== "GroupItem") return;
    for (var i = 0; i < container.pageItems.length; i++) {
        var item = container.pageItems[i];
        if (item.typename === "TextFrame" && item.contents.indexOf(pattern) > -1) {
            item.contents = item.contents.replace(new RegExp(escapeRegExp(pattern), 'g'), newContent);
            var sizeObj = findSizeObjectInContainer(rootContainer);
            if (sizeObj) adjustTextToFitWidth(item, sizeObj.width);
        } else if (item.typename === "GroupItem") {
            applyPatternReplacement(item, pattern, newContent, rootContainer);
        }
    }
}

// ============================================================
//                    PARSE COLUMNS
// ============================================================
function parseColumns(columns) {
    var result = { fileToPlace: null, objectName: null, layersToShow: [], textReplacements: [] };
    for (var j = 0; j < columns.length; j++) {
        var col = columns[j].replace(/^\s+|\s+$/g, '');
        if (col.indexOf('@') == 0) {
            var parts = col.substring(1).split('@');
            if (parts.length == 2) { result.objectName = parts[0]; result.fileToPlace = parts[1]; }
            else { result.fileToPlace = parts[0]; result.objectName = null; }
        } else if (col.indexOf('*') == 0) {
            var textParts = col.substring(1).split('*');
            if (textParts.length == 2) result.textReplacements.push({ objectName: textParts[0], newContent: textParts[1] });
        } else if (col != '') {
            result.layersToShow.push(col);
        }
    }
    return result;
}

// ============================================================
//                    MAIN FUNCTION
// ============================================================
function processSingleRow(configStr) {
    try {
        if (!app.documents.length) return "ERROR: Không có document AI nào đang mở!";
        var activeDoc = app.activeDocument;
        var config = eval("(" + configStr + ")");
        var isPreview = config.isPreview === true;

        if (isPreview) removeOldPreviewLayers(activeDoc);

        var parsed = parseColumns(config.columns);
        var timestamp = (new Date()).getTime();
        var tempLayersMap = {};

        // BƯỚC 1: Phân tích & tạo layer tạm TRƯỚC khi ẩn
        buildCache(activeDoc);

        if (parsed.objectName) {
            var targetObjects = findObjectsByName(parsed.objectName, activeDoc);
            for (var i = 0; i < targetObjects.length; i++) {
                var tplName = findTemplateLayerName(targetObjects[i]);
                if (tplName && !tempLayersMap[tplName])
                    getOrCreateTempLayer(activeDoc, tplName, tempLayersMap, timestamp, isPreview);
            }
        }
        if (parsed.textReplacements.length > 0) {
            for (var t = 0; t < parsed.textReplacements.length; t++) {
                var txObjs = findTextFramesByName(parsed.textReplacements[t].objectName, activeDoc);
                for (var i = 0; i < txObjs.length; i++) {
                    var tplName = findTemplateLayerName(txObjs[i]);
                    if (tplName && !tempLayersMap[tplName])
                        getOrCreateTempLayer(activeDoc, tplName, tempLayersMap, timestamp, isPreview);
                }
            }
        }

        // BƯỚC 2: Ẩn layers (trừ Template)
        hideAllLayers(activeDoc);

        // BƯỚC 3: Xử lý hình ảnh
        if (parsed.fileToPlace) {
            var sourceFile = new File(config.sourceFolder + '/' + parsed.fileToPlace);
            if (!sourceFile.exists) return "ERROR: File hình ảnh không tồn tại '" + parsed.fileToPlace + "'";

            if (parsed.objectName) {
                var targetObjects = findObjectsByName(parsed.objectName, activeDoc);
                if (targetObjects.length == 0) return "ERROR: Không tìm thấy Object '" + parsed.objectName + "'";

                for (var idx = 0; idx < targetObjects.length; idx++) {
                    var tObj = targetObjects[idx];
                    var tplName = findTemplateLayerName(tObj);
                    var tempLayer = tempLayersMap[tplName] || tempLayersMap["_fallback"];
                    if (!tempLayer) {
                        tempLayer = activeDoc.layers.add();
                        tempLayer.name = "Temp_" + timestamp;
                        tempLayer.visible = true;
                        tempLayersMap["_fallback"] = tempLayer;
                    }
                    var pc = getParentContainer(tObj);
                    if (pc) {
                        var cc = pc.duplicate(tempLayer, ElementPlacement.PLACEATBEGINNING);
                        replaceObjectInContainer(cc, parsed.objectName, sourceFile);
                    } else {
                        var pi = tempLayer.placedItems.add();
                        pi.file = sourceFile;
                        pi.width = tObj.width; pi.height = tObj.height; pi.position = tObj.position;
                    }
                }
            } else {
                var tempLayer = activeDoc.layers.add();
                tempLayer.name = isPreview ? "PREVIEW_IMG" : "Temp_" + timestamp;
                tempLayer.visible = true;
                tempLayersMap["_noObject"] = tempLayer;
                var pi = tempLayer.placedItems.add();
                pi.file = sourceFile;
            }
        }

        // BƯỚC 4: Xử lý text
        if (parsed.textReplacements.length > 0) {
            var groups = groupReplacementsByAncestor(parsed.textReplacements, activeDoc);
            for (var g = 0; g < groups.length; g++) {
                var grp = groups[g];
                var tplName = findTemplateLayerName(grp.container);
                if (tplName === "") tplName = "_fallbackText";
                var tempLayer = tempLayersMap[tplName];
                if (!tempLayer) {
                    tempLayer = activeDoc.layers.add();
                    tempLayer.name = "Temp_" + timestamp;
                    tempLayer.visible = true;
                    tempLayersMap[tplName] = tempLayer;
                }
                if (grp.isGroup) {
                    var copiedContainer = grp.container.duplicate(tempLayer, ElementPlacement.PLACEATBEGINNING);
                    applyTextReplacementsToContainer(copiedContainer, grp.replacements);
                } else {
                    var copiedText = grp.container.duplicate(tempLayer, ElementPlacement.PLACEATBEGINNING);
                    for (var r = 0; r < grp.replacements.length; r++) {
                        var rep = grp.replacements[r];
                        var pattern = '*' + rep.objectName + '*';
                        if (copiedText.name === rep.objectName)
                            copiedText.contents = rep.newContent;
                        else if (copiedText.contents.indexOf(pattern) > -1)
                            copiedText.contents = copiedText.contents.replace(new RegExp(escapeRegExp(pattern), 'g'), rep.newContent);
                        var sizeObj = findSizeObjectInContainer(tempLayer);
                        if (sizeObj) adjustTextToFitWidth(copiedText, sizeObj.width);
                    }
                }
            }
        }

        // BƯỚC 5: Hiện layers
        for (var k = 0; k < parsed.layersToShow.length; k++)
            showLayer(activeDoc, parsed.layersToShow[k]);

        // BƯỚC 6: Preview hoặc xuất PDF
        if (isPreview) { app.redraw(); return "OK"; }

        var pdfFile = new File(config.outputFolder + '/' + config.pdfName);
        var pdfOptions = new PDFSaveOptions();
        pdfOptions.pDFPreset = config.presetName;
        activeDoc.saveAs(pdfFile, pdfOptions);
        removeAllTempLayers(tempLayersMap);
        return "OK";
    } catch (e) {
        return "ERROR: " + e.message;
    }
}

// ============================================================
//                    CHECK BEFORE EXPORT
// ============================================================
function checkAllDataBeforeExport(configList, sourceFolderPath) {
    try {
        if (!app.documents.length) return "LỖI: Không có document AI nào đang mở!";
        var doc = app.activeDocument;
        var sourceFolder = new Folder(sourceFolderPath);
        var errors = [];
        buildCache(doc);

        for (var i = 0; i < configList.length; i++) {
            var config = configList[i];
            var columns = config.columns;
            for (var j = 0; j < columns.length; j++) {
                var col = columns[j].replace(/^\s+|\s+$/g, '');
                if (col == '') continue;
                if (col.indexOf('@') == 0) {
                    var parts = col.substring(1).split('@');
                    if (parts.length == 2) {
                        if (findObjectsByName(parts[0], doc).length == 0)
                            errors.push("File " + config.pdfName + ": Object '@" + parts[0] + "' không tìm thấy");
                        if (!new File(sourceFolder.fsName + '/' + parts[1]).exists)
                            errors.push("File " + config.pdfName + ": File ảnh '" + parts[1] + "' không tồn tại");
                    } else {
                        if (!new File(sourceFolder.fsName + '/' + parts[0]).exists)
                            errors.push("File " + config.pdfName + ": File ảnh '" + parts[0] + "' không tồn tại");
                    }
                } else if (col.indexOf('*') == 0) {
                    var textParts = col.substring(1).split('*');
                    if (textParts.length == 2 && findTextFramesByName(textParts[0], doc).length == 0)
                        errors.push("File " + config.pdfName + ": Text '*" + textParts[0] + "*' không tìm thấy");
                } else {
                    var layerFound = false;
                    for (var k = 0; k < doc.layers.length; k++)
                        if (doc.layers[k].name === col) { layerFound = true; break; }
                    if (!layerFound)
                        errors.push("File " + config.pdfName + ": Layer '" + col + "' không tồn tại");
                }
            }
        }

        if (errors.length > 0) {
            var msg = "PHÁT HIỆN " + errors.length + " LỖI DỮ LIỆU:\n================================\n\n";
            for (var i = 0; i < errors.length; i++) msg += errors[i] + "\n";
            msg += "\n================================\nVui lòng sửa lỗi trước khi xuất!";
            return msg;
        }
        return "OK";
    } catch (e) {
        return "Lỗi kịch bản kiểm tra: " + e.message;
    }
}

// ============================================================
//                    UTILITY FUNCTIONS
// ============================================================
function hideAllLayers(doc) {
    for (var i = 0; i < doc.layers.length; i++) {
        var layer = doc.layers[i];
        if (layer.name.toLowerCase().indexOf("template") >= 0) {
            // ✅ Bật layer template lên (dù đang tắt)
            layer.visible = true;
        } else {
            layer.visible = false;
        }
    }
}

function showLayer(doc, layerName) {
    try { doc.layers.getByName(layerName).visible = true; } catch(e) {}
}

function getParentContainer(item) {
    var parent = item.parent;
    if (item.typename === "TextFrame" && parent.typename === "Story")
        parent = item.parent.textFrames[0].parent;
    if (parent.typename === "GroupItem") return parent;
    if (parent.typename === "CompoundPathItem" && parent.pathItems.length > 0)
        try { if (parent.pathItems[0].clipping === true) return parent; } catch(e) {}
    if (parent.typename === "GroupItem" && parent.clipped === true) return parent;
    if (parent.typename !== "Layer" && parent.typename !== "Document") return getParentContainer(parent);
    return null;
}

function replaceObjectInContainer(container, objectName, sourceFile) {
    var targetItems = findObjectsInContainer(container, objectName);
    for (var i = 0; i < targetItems.length; i++) {
        var t = targetItems[i];
        var ow = t.width, oh = t.height, op = t.position, tp = t.parent;
        t.remove();
        var pi = (tp.typename === "GroupItem") ? tp.placedItems.add() : container.layer.placedItems.add();
        pi.file = sourceFile; pi.width = ow; pi.height = oh; pi.position = op; pi.name = objectName;
    }
}

function findObjectsInContainer(container, name) {
    var found = [];
    if (container.typename === "GroupItem") {
        for (var i = 0; i < container.pageItems.length; i++) {
            var item = container.pageItems[i];
            if (item.name === name) found.push(item);
            if (item.typename === "GroupItem") found = found.concat(findObjectsInContainer(item, name));
        }
    } else if (container.typename === "CompoundPathItem") {
        for (var i = 0; i < container.pathItems.length; i++)
            if (container.pathItems[i].name === name) found.push(container.pathItems[i]);
    }
    return found;
}

function findTextFramesInContainer(container, name) {
    var found = [];
    if (container.typename === "GroupItem") {
        for (var i = 0; i < container.pageItems.length; i++) {
            var item = container.pageItems[i];
            if (item.typename === "TextFrame" && item.name === name) found.push(item);
            if (item.typename === "GroupItem") found = found.concat(findTextFramesInContainer(item, name));
        }
    }
    return found;
}

function findSizeObjectInContainer(container) {
    if (container.typename === "GroupItem") {
        for (var i = 0; i < container.pageItems.length; i++) {
            var item = container.pageItems[i];
            if (item.typename === "PathItem" && item.name === "<size>") return item;
            if (item.typename === "GroupItem") { var f = findSizeObjectInContainer(item); if (f) return f; }
        }
    }
    return null;
}

function adjustTextToFitWidth(textFrame, maxWidth) {
    try {
        var targetWidth = maxWidth * 0.95;
        if (textFrame.width > targetWidth) {
            var just = textFrame.textRange.paragraphAttributes.justification;
            var cx = textFrame.position[0] + (textFrame.width / 2);
            var oy = textFrame.position[1];
            textFrame.width = targetWidth;
            if (just == Justification.CENTER) textFrame.position = [cx - (targetWidth / 2), oy];
            else if (just == Justification.RIGHT) textFrame.position = [textFrame.position[0] + textFrame.width - targetWidth, oy];
        }
    } catch(e) {}
}

function escapeRegExp(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}