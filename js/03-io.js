function toggleImportMenu() {
    var menu = document.getElementById('importMenu');
    if (!menu)
        return;
    var isShow = menu.classList.contains('show');
    menu.classList.toggle('show', !isShow);
    if (!isShow)
        loadImportFormatList();
}
function loadImportFormatList() {
    postToAHK({
        action: 'getFormatList'
    })
}
function onImportFormatList(formats) {
    _importFormats = (formats || []).map(function (f) {
        return {
            name: f.name,
            cols: parseFormatCSV(f.content)
        }
    });
    var list = document.getElementById('importFormatList'); // Đây vẫn là div bên trong menu
    if (!list)
        return;
    if (!_importFormats.length) {
        list.innerHTML = '<div class="dropdown-item-placeholder">Không có format</div>';
        return
    }
    var html = '';
    for (var i = 0; i < _importFormats.length; i++) {
        var f = _importFormats[i],
        n = f.cols ? f.cols.length : 0;
        html += '<div class="dropdown-item" onclick="importWithFormat(' + i + ')" title="' + n + ' cột"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15.5 2H8.6c-.4 0-.8.2-1.1.5-.3.3-.5.7-.5 1.1v12.8c0 .4.2.8.5 1.1.3.3.7.5 1.1.5h9.8c.4 0 .8-.2 1.1-.5.3-.3.5-.7.5-1.1V6.5L15.5 2z"></path><path d="M3 7.6v12.8c0 .4.2.8.5 1.1.3.3.7.5 1.1.5h9.8"></path><path d="M15 2v5h5"></path></svg>' + escHtml(f.name) + '</div>';
    }
    list.innerHTML = html;
}
async function importExcelFull() {
    var m = document.getElementById('importMenu');
    if (m)
        m.classList.remove('show');
    await jsImportExcel(null)
}
async function importWithFormat(idx) {
    var m = document.getElementById('importMenu');
    if (m)
        m.classList.remove('show');
    var fmt = _importFormats[idx];
    if (!fmt)
        return;
    if (!fmt.cols || !fmt.cols.length) {
        toast('Format "' + fmt.name + '" không hợp lệ');
        return
    }
    await jsImportExcel(fmt)
}
async function jsImportExcel(format) {
    try {
        var opts = {
            types: [{
                    description: 'Excel files',
                    accept: {
                        'application/octet-stream': ['.xlsx', '.xlsm', '.xls']
                    }
                }
            ],
            multiple: !1
        };
        if (window._lastDirHandle)
            opts.startIn = window._lastDirHandle;
        const [fileHandle] = await window.showOpenFilePicker(opts);
        window._lastDirHandle = fileHandle;
        const file = await fileHandle.getFile();
        const buffer = await file.arrayBuffer();
        var workbook = XLSX.read(buffer, {
            type: 'array',
            cellDates: !0,
            raw: !0
        });
        var sheetName = workbook.SheetNames.length > 1 ? await pickSheet(workbook.SheetNames) : workbook.SheetNames[0];
        if (!sheetName)
            return;
        var sheet = workbook.Sheets[sheetName];
        var jsonData = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: ''
        });
        if (!jsonData.length) {
            toast('Sheet trống!');
            return
        }
        var isAppend = false;
        var hasData = !1;
        for (var i = 0; i < _currentState.data.length; i++) {
            for (var j = 0; j < _currentState.data[i].length; j++) {
                if (_currentState.data[i][j] !== '') {
                    hasData = !0;
                    break
                }
            }
            if (hasData)
                break
        }
        if (hasData) {
            const result = await Swal.fire({
                title: 'Bảng đang có dữ liệu',
                text: 'Bạn muốn nối tiếp dữ liệu mới hay thay thế hoàn toàn?',
                icon: 'question',
                showDenyButton: true,
                confirmButtonText: 'Nối tiếp',
                denyButtonText: 'Thay thế',
                confirmButtonColor: '#1a73e8',
                denyButtonColor: '#d93025',
            });
            if (result.isConfirmed) {
                isAppend = true;
            } else if (!result.isDenied) {
                return;
            }
        }
        if (!format) {
            var maxCols = 0;
            for (var r = 0; r < jsonData.length; r++)
                if (jsonData[r].length > maxCols)
                    maxCols = jsonData[r].length;
            var newHeaders = [];
            for (var i = 0; i < maxCols; i++)
                newHeaders.push('Column ' + (i + 1));
            var newRows = jsonData.map(function (row) {
                var r = [];
                for (var c = 0; c < maxCols; c++)
                    r.push(_xlCellToStr(row[c]));
                return r
            });
            if (isAppend)
                appendData(newRows, file.name + ' [' + sheetName + ']');
            else
                loadDataDirect(newHeaders, newRows, file.name + ' [' + sheetName + ']')
        } else {
            _applyFormatImport(jsonData, format, file.name, sheetName, isAppend)
        }
    } catch (e) {
        if (e.name !== 'AbortError')
            toast('Lỗi: ' + e.message);
    }
}
function _xlCellToStr(v) {
    if (v === null || v === undefined)
        return '';
    if (v instanceof Date) {
        var y = v.getFullYear(),
        mo = v.getMonth() + 1,
        d = v.getDate();
        return y + '/' + (mo < 10 ? '0' + mo : mo) + '/' + (d < 10 ? '0' + d : d)
    }
    return String(v)
}
function _applyFormatImport(jsonData, format, fileName, sheetName, isAppend) {
    var cols = format.cols;
    if (!cols || !cols.length) {
        toast('Format trống');
        return
    }
    var baseColDef = null,
    globalStartRow = Infinity;
    for (var i = 0; i < cols.length; i++) {
        if (cols[i].isBase)
            baseColDef = cols[i];
        if (!cols[i].isEmpty && cols[i].startRow < globalStartRow)
            globalStartRow = cols[i].startRow
    }
    if (!isFinite(globalStartRow) || globalStartRow < 1)
        globalStartRow = 1;
    var startIdx = globalStartRow - 1,
    endIdx = jsonData.length - 1;
    if (baseColDef) {
        var baseXlCol = _xlColToIdx(baseColDef.colLetter),
        baseStartIdx = baseColDef.startRow - 1;
        endIdx = -1;
        for (var r = jsonData.length - 1; r >= baseStartIdx; r--) {
            if (!jsonData[r])
                continue;
            var rawVal = jsonData[r][baseXlCol];
            if (rawVal === undefined || rawVal === null || rawVal === 0)
                continue;
            var cellVal = (rawVal instanceof Date) ? _xlCellToStr(rawVal) : String(rawVal);
            cellVal = cellVal.replace(/[\s\u3000\u00A0\u200B\u200C\u200D\uFEFF\t\r\n]/g, '');
            if (cellVal === '' || cellVal === '0' || cellVal === '0.0')
                continue;
            endIdx = r;
            break
        }
        if (endIdx < startIdx) {
            toast('Không có dữ liệu');
            return
        }
    }
    var newHeaders = [],
    newColWidths = [];
    for (var ci = 0; ci < cols.length; ci++) {
        newHeaders.push(cols[ci].name || ('Col ' + (ci + 1)));
        newColWidths.push(cols[ci].width)
    }
    var newRows = [];
    for (var rowOffset = 0; rowOffset <= endIdx - startIdx; rowOffset++) {
        var row = [];
        for (var ci = 0; ci < cols.length; ci++) {
            var colDef = cols[ci],
            excelRow = startIdx + rowOffset;
            if (excelRow < colDef.startRow - 1 || colDef.isEmpty) {
                row.push('');
                continue
            }
            var xlColIdx = _xlColToIdx(colDef.colLetter);
            var raw = jsonData[excelRow] ? jsonData[excelRow][xlColIdx] : '';
            row.push(_xlCellToStr(raw !== undefined ? raw : ''))
        }
        newRows.push(row)
    }
    if (isAppend) {
        appendData(newRows, fileName + ' [' + sheetName + '] — ' + format.name)
    } else {
        loadDataDirect(newHeaders, newRows, fileName + ' [' + sheetName + '] — ' + format.name);
        setTimeout(function () {
            _currentState.hiddenCols = new Set();
            _currentState.fixedWidthCols = new Set();
            for (var ci = 0; ci < cols.length; ci++) {
                if (cols[ci].hidden) {
                    _currentState.colWidths[ci] = 0;
                    _currentState.hiddenCols.add(ci)
                } else if (cols[ci].width > 0) {
                    _currentState.colWidths[ci] = cols[ci].width;
                    _currentState.fixedWidthCols.add(ci)
                } else {
                    var maxW = measureTextWidth(newHeaders[ci] || '') + 14;
                    for (var r = 0; r < _currentState.data.length; r++) {
                        var w = measureTextWidth(_currentState.data[r][ci] || '') + 14;
                        if (w > maxW)
                            maxW = w
                    }
                    _currentState.colWidths[ci] = Math.max(MIN_COL_WIDTH, Math.min(maxW, 400))
                }
            }
            rebuildPrefixSums();
            recalcAllRowHeights();
            fullRender();
            toast('✅ ' + format.name + ' tải thành công')
        }, 60)
    }
}
function parseFormatCSV(content) {
    if (!content)
        return [];
    var cols = [],
    lines = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n'),
    firstLine = '';
    for (var i = 0; i < lines.length; i++) {
        var t = lines[i].trim();
        if (!t || t.charAt(0) === ';')
            continue;
        firstLine = t;
        break
    }
    if (!firstLine)
        return [];
    var cells = _parseCSVRow(firstLine);
    for (var ci = 0; ci < cells.length; ci++) {
        var cell = cells[ci].trim();
        if (!cell)
            continue;
        cell = cell.replace(/^[""\u201C\u201D]+|[""\u201C\u201D]+$/g, '');
        var parts = cell.split(','),
        name = (parts[0] || '').trim().replace(/^["]+|["]+$/g, '');
        var colRef = (parts[1] || '').trim().replace(/^["]+|["]+$/g, ''),
        widthStr = (parts[2] || '').trim().replace(/^["]+|["]+$/g, '');
        var opt = (parts[3] || '').trim().toLowerCase().replace(/^["]+|["]+$/g, ''),
        colLetter = '',
        startRow = 1;
        if (colRef) {
            var refMatch = colRef.match(/^([A-Za-z]+)(\d+)$/);
            if (refMatch) {
                colLetter = refMatch[1].toUpperCase();
                startRow = parseInt(refMatch[2]) || 1
            }
        }
        var widthVal = parseInt(widthStr);
        cols.push({
            name: name,
            colLetter: colLetter,
            startRow: startRow,
            width: isNaN(widthVal) ? -1 : widthVal,
            hidden: widthVal === 0,
            isBase: opt === 'base',
            isEmpty: colLetter === ''
        })
    }
    return cols
}
function _parseCSVRow(line) {
    var fields = [],
    field = '',
    inQ = !1,
    i = 0;
    while (i < line.length) {
        var c = line[i];
        if (inQ) {
            if (c === '"') {
                if (i + 1 < line.length && line[i + 1] === '"') {
                    field += '"';
                    i += 2;
                    continue
                }
                inQ = !1
            } else {
                field += c
            }
        } else {
            if (c === '"') {
                inQ = !0
            } else if (c === ',') {
                fields.push(field);
                field = ''
            } else {
                field += c
            }
        }
        i++
    }
    fields.push(field);
    return fields
}
function _xlColToIdx(letter) {
    if (!letter)
        return 0;
    var s = String(letter).replace(/\d+/g, '').toUpperCase(),
    n = 0;
    for (var i = 0; i < s.length; i++)
        n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1
}
async function pickSheet(names) {
    const inputOptions = {};
    names.forEach(name => {
        inputOptions[name] = name;
    });
    const { value: sheetName } = await Swal.fire({
        title: 'Chọn Sheet Excel',
        input: 'select',
        inputOptions: inputOptions,
        inputPlaceholder: 'Chọn một sheet để import',
        showCancelButton: true,
        confirmButtonText: 'Chọn',
        cancelButtonText: 'Hủy'
    });
    return sheetName || null;
}
function loadDataDirect(newHeaders, newData, fileName) {
    _currentState.headers = newHeaders;
    _currentState.data = newData;
    _currentState.colWidths = [];
    for (var i = 0; i < newHeaders.length; i++)
        _currentState.colWidths.push(120);
    _currentState.manualRowHeights = {};
    _currentState.undoStack = [];
    _currentState.redoStack = [];
    _currentState.displayCache = [];

    clearSel();
    selectedRows = [];
    selectedCols = [];

    if (activeTabNum === 1) {
        _currentState.rowLayerData = {};
        _currentState.currentRowSel = 0;
        document.getElementById('activeRowNum').textContent = '—';
        refreshActiveLayersList([]);
    }

    _currentState.fileName = fileName;
    _currentState._cleanFileName = fileName.split(' | ')[0].replace('📄 ', '');

    calcAllRowHeights();
    rebuildPrefixSums();
    var gw = document.getElementById('gridWrapper');
    if (gw) {
        gw.scrollTop = 0;
        gw.scrollLeft = 0
    }
    initHyperFormula();
    rebuildDisplayCache();
    fullRender();
    _drawFileStatus();

    toast('Đã tải mới ' + newData.length + ' dòng');
}
function appendData(newRows, fileName) {
    saveState();
    var startRow = _currentState.data.length;
    var addedCount = 0;
    for (var i = 0; i < newRows.length; i++) {
        var row = newRows[i];
        var normalizedRow = [];
        for (var c = 0; c < _currentState.headers.length; c++) {
            normalizedRow.push(row[c] !== undefined ? row[c] : '')
        }
        _currentState.data.push(normalizedRow);
        _currentState.rowHeights.push(MIN_ROW_HEIGHT);
        addedCount++;
    }
    for (var r = startRow; r < _currentState.data.length; r++) {
        recalcRowHeight(r)
    }
    for (var c = 0; c < _currentState.headers.length; c++) {
        if (_currentState.hiddenCols.has(c) || _currentState.fixedWidthCols.has(c))
            continue;
        var maxW = _currentState.colWidths[c] || MIN_COL_WIDTH;
        for (var r = startRow; r < _currentState.data.length; r++) {
            if (!_currentState.data[r][c])
                continue;
            var w = measureTextWidth(_currentState.data[r][c]) + 16;
            if (w > maxW)
                maxW = w
        }
        _currentState.colWidths[c] = Math.min(maxW, 800)
    }
    rebuildPrefixSums();
    syncHyperFormula();
    rebuildDisplayCache();
    fullRender();
    toast('✅ Đã nối thêm ' + addedCount + ' dòng');
    markAsModified();
    _currentState.fileName = 'Đã nối: ' + fileName + ' | Tổng: ' + _currentState.data.length + ' dòng';
    _drawFileStatus();
    setTimeout(function () {
        scrollToCell(startRow + 1, 1)
    }, 100);
}
function clearTableData() {
    const defaultState = createDefaultWorkspaceState();
    _currentState.headers = defaultState.headers;
    _currentState.data = defaultState.data;
    _currentState.hiddenCols = new Set();
    _currentState.fixedWidthCols = new Set();
    _currentState.fileName = '';
    _currentState._cleanFileName = '';
    _currentState.isModified = false;

    if (activeTabNum === 1) {
        allLayersFromAI = [];
        layerColors = {};
        _currentState.rowMatchedRules = {};
        renderAllLayersList();
    }
    if (activeTabNum === 3) {
        _currentState.fileHandle = null;
        _currentState._realFilePath = null;
    }

    loadDataDirect(defaultState.headers, defaultState.data, '');
    toast('🗑️ Đã xóa dữ liệu');
    _drawFileStatus();
}
function exportXlsx() {
    try {
        var exportHeaders = [];
        for (var c = 0; c < _currentState.headers.length; c++) {
            exportHeaders.push(_currentState.headers[c] + (_currentState.hiddenCols.has(c) ? "__HIDDEN__" : ""))
        }
        var allLayersWithColor = [];
        for (var i = 0; i < allLayersFromAI.length; i++) {
            var lname = allLayersFromAI[i];
            var lcolor = layerColors[lname] || '200,200,200';
            allLayersWithColor.push(lname + '|' + lcolor)
        }
        exportHeaders.push(allLayersWithColor.join(';;'));
        var exportData = [];
        for (var r = 0; r < _currentState.data.length; r++) {
            var rowData = _currentState.data[r].slice();
            var activeLayers = _currentState.rowLayerData[r + 1] || [];
            rowData.push(activeLayers.join(';;'));
            exportData.push(rowData)
        }
        var wb = XLSX.utils.book_new();
        var ws = XLSX.utils.aoa_to_sheet([exportHeaders].concat(exportData));
        XLSX.utils.book_append_sheet(wb, ws, 'Data');
        XLSX.writeFile(wb, 'Export_Data_Layers.xlsx');
        toast('✅ Đã export thành công (Bao gồm Layers)')
    } catch (e) {
        toast('❌ Lỗi export: ' + e.message)
    }
}
async function importProcessedExcel() {
    var m = document.getElementById('importMenu');
    if (m)
        m.classList.remove('show');
    try {
        var opts = {
            types: [{
                    description: 'Excel',
                    accept: {
                        'application/octet-stream': ['.xlsx', '.xlsm', '.xls']
                    }
                }
            ],
            multiple: !1
        };
        const [fileHandle] = await window.showOpenFilePicker(opts);
        const file = await fileHandle.getFile();
        const buffer = await file.arrayBuffer();
        var workbook = XLSX.read(buffer, {
            type: 'array',
            cellDates: !0,
            raw: !0
        });
        var jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {
            header: 1,
            defval: ''
        });
        if (!jsonData.length || !jsonData[0].length) {
            toast('Sheet trống!');
            return
        }
        var rawHeaders = jsonData[0];
        var lastColIdx = rawHeaders.length - 1;
        var allLayersStr = rawHeaders[lastColIdx];
        allLayersFromAI = [];
        layerColors = {};
        if (allLayersStr) {
            var layersArr = allLayersStr.split(';;');
            for (var i = 0; i < layersArr.length; i++) {
                var parts = layersArr[i].trim().split('|');
                if (parts[0]) {
                    allLayersFromAI.push(parts[0]);
                    layerColors[parts[0]] = parts[1] || '200,200,200'
                }
            }
        }
        var newHeaders = [];
        var newHiddenCols = new Set();
        for (var c = 0; c < lastColIdx; c++) {
            var h = String(rawHeaders[c]);
            if (h.endsWith("__HIDDEN__")) {
                newHiddenCols.add(c);
                newHeaders.push(h.replace("__HIDDEN__", ""))
            } else {
                newHeaders.push(h)
            }
        }
        var newRows = [];
        var tempRowLayerData = {};
        for (var r = 1; r < jsonData.length; r++) {
            var row = jsonData[r];
            var activeStr = row[lastColIdx];
            var normalizedRow = [];
            for (var c = 0; c < newHeaders.length; c++)
                normalizedRow.push(_xlCellToStr(row[c]));
            newRows.push(normalizedRow);
            if (activeStr) {
                tempRowLayerData[r] = activeStr.split(';;').map(function (s) {
                    return s.trim()
                }).filter(function (s) {
                    return s
                })
            } else {
                tempRowLayerData[r] = []
            }
        }
        loadDataDirect(newHeaders, newRows, file.name + ' (Đã xử lý)');
        _currentState.rowLayerData = tempRowLayerData;
        _currentState.hiddenCols = newHiddenCols;
        for (var c = 0; c < newHeaders.length; c++)
            _currentState.colWidths[c] = _currentState.hiddenCols.has(c) ? 0 : 120;
        rebuildPrefixSums();
        renderAllLayersList();
        if (_currentState.currentRowSel > 0)
            refreshActiveLayersList(_currentState.rowLayerData[_currentState.currentRowSel] || []);
        toast('🔄 Đã nạp Data & Layers. Đang tô màu chữ...');
        setTimeout(function () {
            if (currentRules && currentRules.length > 0) {
                applyRulesFromGrid(!0)
            } else {
                toast('⚠️ Bạn chưa nạp file Rules! Không thể tô màu chữ.');
                fullRender()
            }
        }, 300)
    } catch (e) {
        if (e.name !== 'AbortError')
            toast('❌ Lỗi: ' + e.message);
    }
}
async function parseCSVTextAndLoadTab3(text, fileName) {
    var lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    if (!lines.length || !lines[0]) {
        toast("File CSV rỗng!");
        return
    }
    var newData = [];
    var maxCols = 0;
    for (var i = 0; i < lines.length; i++) {
        if (lines[i].trim() === "")
            continue;
        var row = _parseCSVRow(lines[i]);
        if (row.length > maxCols)
            maxCols = row.length;
        newData.push(row)
    }
    var newHeaders = [];
    for (var c = 1; c <= maxCols; c++) {
        newHeaders.push('Column ' + c)
    }
    for (var r = 0; r < newData.length; r++) {
        while (newData[r].length < maxCols) {
            newData[r].push("")
        }
    }
    loadDataDirect(newHeaders, newData, fileName);
    for (let i = 0; i < newHeaders.length; i++)
        _currentState.colWidths[i] = 120;
    rebuildPrefixSums();
    fullRender();
    _drawFileStatus();
    toast("✅ Đã nạp CSV thành công!")
}
function exportPDFTab3() {
    if (!_currentState.data || _currentState.data.length === 0 || _currentState.data[0][0] === "") {
        toast("⚠️ Bảng dữ liệu trống!");
        return
    }
    let rowsToExport = [];
    if (selectedRows && selectedRows.length > 0) {
        let sortedRows = selectedRows.slice().sort((a, b) => a - b);
        for (let i = 0; i < sortedRows.length; i++) {
            if (_currentState.data[sortedRows[i] - 1])
                rowsToExport.push(_currentState.data[sortedRows[i] - 1]);
        }
    } else if (selRange) {
        let rStart = Math.min(selRange.r1, selRange.r2),
        rEnd = Math.max(selRange.r1, selRange.r2);
        for (let r = rStart; r <= rEnd; r++) {
            if (_currentState.data[r - 1])
                rowsToExport.push(_currentState.data[r - 1]);
        }
    } else if (selCells && selCells.size > 0) {
        let rSet = new Set();
        selCells.forEach(k => rSet.add(parseInt(k.split(',')[0])));
        let rArr = Array.from(rSet).sort((a, b) => a - b);
        for (let i = 0; i < rArr.length; i++) {
            if (_currentState.data[rArr[i] - 1])
                rowsToExport.push(_currentState.data[rArr[i] - 1]);
        }
    } else if (currentCell) {
        if (_currentState.data[currentCell.row - 1])
            rowsToExport.push(_currentState.data[currentCell.row - 1]);
    }
    if (rowsToExport.length === 0)
        rowsToExport = _currentState.data;
    toast(`🚀 Chuẩn bị xuất ${rowsToExport.length} dòng...`);
    let hasLink = !1;
    for (let r = 0; r < rowsToExport.length; r++) {
        for (let c = 0; c < rowsToExport[r].length; c++) {
            if (String(rowsToExport[r][c]).trim().startsWith('@')) {
                hasLink = !0;
                break
            }
        }
        if (hasLink)
            break
    }
    window._tempExportData = rowsToExport;
    window._tempHasLink = hasLink;
    postToAHK({
        action: 'requestPdfExportUI'
    })
}
function showPdfExportDialog(presets) {
    let old = document.getElementById('_aiExportDialog');
    if (old)
        old.remove();
    let overlay = document.createElement('div');
    overlay.id = '_aiExportDialog';
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999;display:flex;align-items:center;justify-content:center;';
    let box = document.createElement('div');
    box.style.cssText = 'background:#fff;border-radius:10px;padding:20px 24px;min-width:380px;max-width:450px;box-shadow:0 8px 32px rgba(0,0,0,.25);font-family:Google Sans,sans-serif; display: flex; flex-direction: column; gap: 14px;';
    let h = document.createElement('div');
    h.style.cssText = 'font-size:15px;font-weight:600;color:#1a73e8;text-transform:uppercase;letter-spacing:0.5px;border-bottom:1px solid #eee;padding-bottom:10px;';
    h.textContent = '🖨️ Xuất PDF tự động (Illustrator)';
    box.appendChild(h);
    let linkDisabled = window._tempHasLink ? '' : 'disabled';
    let linkStyle = window._tempHasLink ? '' : 'opacity:0.5; pointer-events:none;';
    let linkPlaceholder = window._tempHasLink ? 'Chưa chọn...' : 'Không có hình ảnh (Bỏ qua)';
    let srcGrp = document.createElement('div');
    srcGrp.innerHTML = `<div style="font-size:12px; font-weight:600; color:#5f6368; margin-bottom:4px; ${linkStyle}">1. Thư mục chứa File Link (Hình, PDF...):</div>
                        <div style="display:flex; gap:8px; ${linkStyle}">
                            <input type="text" id="aiSrcFolder" readonly placeholder="${linkPlaceholder}" ${linkDisabled} style="flex:1; padding:7px; border:1px solid #dadce0; border-radius:4px; font-size:12px; background:#f8f9fa;">
                            <button onclick="postToAHK({action:'browseFolder', target:'aiSrcFolder'})" ${linkDisabled} style="padding:0 12px; border:1px solid #dadce0; background:#fff; cursor:pointer; border-radius:4px;">Chọn</button>
                        </div>`;
    box.appendChild(srcGrp);
    let outGrp = document.createElement('div');
    outGrp.innerHTML = `<div style="font-size:12px; font-weight:600; color:#5f6368; margin-bottom:4px;">2. Thư mục XUẤT PDF:</div>
                        <div style="display:flex; gap:8px;">
                            <input type="text" id="aiOutFolder" readonly placeholder="Chưa chọn..." style="flex:1; padding:7px; border:1px solid #dadce0; border-radius:4px; font-size:12px; background:#f8f9fa;">
                            <button onclick="postToAHK({action:'browseFolder', target:'aiOutFolder'})" style="padding:0 12px; border:1px solid #dadce0; background:#fff; cursor:pointer; border-radius:4px;">Chọn</button>
                        </div>`;
    box.appendChild(outGrp);
    let preGrp = document.createElement('div');
    let selHtml = `<div style="font-size:12px; font-weight:600; color:#5f6368; margin-bottom:4px;">3. Chọn PDF Preset:</div>
                   <select id="aiPresetSel" style="width:100%; border:1px solid #dadce0; border-radius:4px; font-size:13px; outline:none;" size="8">`;
    for (let i = 0; i < presets.length; i++) {
        let sel = (presets[i].includes("High Quality") || presets[i].includes("Illustrator")) ? "selected" : "";
        selHtml += `<option value="${presets[i]}" ${sel}>${presets[i]}</option>`
    }
    selHtml += `</select>`;
    preGrp.innerHTML = selHtml;
    box.appendChild(preGrp);
    let btnRow = document.createElement('div');
    btnRow.style.cssText = 'display:flex;justify-content:flex-end;gap:10px;margin-top:10px;padding-top:14px;border-top:1px solid #eee;';
    let btnCancel = document.createElement('button');
    btnCancel.textContent = 'Hủy';
    btnCancel.style.cssText = 'padding:7px 18px;border:1px solid #dadce0;border-radius:5px;background:#fff;color:#5f6368;font-size:12.5px;cursor:pointer;';
    btnCancel.onclick = () => {
        overlay.remove();
        window._tempExportData = null;
        window._tempHasLink = null
    };
    let btnOk = document.createElement('button');
    btnOk.textContent = 'Bắt đầu Xuất';
    btnOk.style.cssText = 'padding:7px 22px;border:1px solid #1557b0;border-radius:5px;background:#1a73e8;color:#fff;font-size:12.5px;font-weight:600;cursor:pointer;';
    btnOk.onclick = () => {
        let src = document.getElementById('aiSrcFolder').value;
        let out = document.getElementById('aiOutFolder').value;
        let pre = document.getElementById('aiPresetSel').value;
        if (window._tempHasLink && !src) {
            toast("Bảng chứa File Link (@). Vui lòng chọn Thư mục Link!");
            return
        }
        if (!out) {
            toast("Vui lòng chọn Thư mục Xuất PDF!");
            return
        }
        overlay.remove();
        let exportData = window._tempExportData || _currentState.data;
        postToAHK({
            action: 'startAiExportThread',
            csvData: exportData,
            sourceFolder: src,
            outputFolder: out,
            presetName: pre
        });
        window._tempExportData = null;
        window._tempHasLink = null
    };
    btnRow.appendChild(btnCancel);
    btnRow.appendChild(btnOk);
    box.appendChild(btnRow);
    overlay.appendChild(box);
    document.body.appendChild(overlay)
}
function triggerAiPreview(rIdx) {
    if (activeTabNum !== 3)
        return;
    let cb = document.getElementById('aiPreviewCb');
    if (!cb || !cb.checked)
        return;
    if (rIdx < 1 || rIdx > _currentState.data.length)
        return;
    let rowData = _currentState.data[rIdx - 1];
    let cols = [];
    let hasLink = !1;
    for (let i = 1; i < rowData.length; i++) {
        let val = String(rowData[i]).trim();
        cols.push(val);
        if (val.startsWith('@'))
            hasLink = !0
    }
    postToAHK({
        action: 'previewAiRow',
        columns: cols,
        hasLink: hasLink
    })
}
function setFolderPath(targetId, path) {
    if (path && document.getElementById(targetId)) {
        document.getElementById(targetId).value = path
    }
}
function updateAiProgress(current, total, filename) {
    var overlay = document.getElementById('aiProgressOverlay');
    if (overlay.style.display === 'none')
        overlay.style.display = 'flex';
    var percent = Math.round((current / total) * 100);
    document.getElementById('aiProgressBar').style.width = percent + '%';
    document.getElementById('aiProgressText').textContent = 'Đang xuất: ' + filename + ' (' + current + '/' + total + ')'
}
function finishAiExport(message) {
    document.getElementById('aiProgressOverlay').style.display = 'none';
    document.getElementById('aiProgressBar').style.width = '0%';
    let old = document.getElementById('_aiReportDialog');
    if (old)
        old.remove();
    let overlay = document.createElement('div');
    overlay.id = '_aiReportDialog';
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999;display:flex;align-items:center;justify-content:center;';
    let box = document.createElement('div');
    box.style.cssText = 'background:#fff;border-radius:10px;padding:20px 24px;box-shadow:0 8px 32px rgba(0,0,0,.25);font-family:Google Sans,sans-serif; display:flex; flex-direction:column; gap:14px;';
    let h = document.createElement('div');
    h.style.cssText = 'font-size:15px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px;border-bottom:1px solid #eee;padding-bottom:10px;';
    h.style.color = message.indexOf("LỖI") >= 0 ? '#d93025' : '#188038';
    h.textContent = message.indexOf("LỖI") >= 0 ? '⚠️ Báo cáo Lỗi Dữ Liệu' : '✅ Báo cáo Xuất PDF';
    box.appendChild(h);
    let content = document.createElement('textarea');
    content.value = message;
    content.readOnly = !0;
    content.style.cssText = 'width:500px; height:300px; min-width:300px; min-height:150px; padding:10px; border:1px solid #dadce0; border-radius:4px; font-size:13px; font-family:"Roboto Mono", monospace; resize:both; background:#f8f9fa; outline:none; white-space:pre;';
    box.appendChild(content);
    let btnRow = document.createElement('div');
    btnRow.style.cssText = 'display:flex;justify-content:flex-end;margin-top:4px;';
    let btnOk = document.createElement('button');
    btnOk.textContent = 'Đóng';
    btnOk.style.cssText = 'padding:7px 22px;border:1px solid #dadce0;border-radius:5px;background:#fff;color:#202124;font-size:13px;font-weight:600;cursor:pointer;';
    btnOk.onclick = () => overlay.remove();
    btnRow.appendChild(btnOk);
    box.appendChild(btnRow);
    overlay.appendChild(box);
    document.body.appendChild(overlay)
}
async function cancelAiExport() {
    const result = await Swal.fire({
        title: 'Dừng xuất PDF?',
        text: "Bạn có chắc muốn dừng tiến trình hiện tại không?",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#d33',
        cancelButtonColor: '#3085d6',
        confirmButtonText: 'Có, dừng lại!',
        cancelButtonText: 'Không'
    });
    if (result.isConfirmed) {
        postToAHK({
            action: 'cancelAiExportThread'
        });
    }
}
function onPreviewCheckboxChange(cb) {
    if (!cb.checked) {
        postToAHK({
            action: 'clearAiPreviewLayers'
        })
    } else {
        let rIdx = currentCell ? currentCell.row : (selectedRows.length > 0 ? selectedRows[0] : 0);
        if (rIdx > 0)
            triggerAiPreview(rIdx);
    }
}
async function saveCsvTab3() {
    if (!_currentState.data || _currentState.data.length === 0)
        return toast("Bảng trống!");
    var csvContent = "";
    _currentState.data.forEach(function (rowArray) {
        var row = rowArray.map(function (cell) {
            var str = String(cell || '');
            if (str.includes(',') || str.includes('"') || str.includes('\n'))
                return '"' + str.replace(/"/g, '""') + '"';
            return str
        }).join(',');
        csvContent += row + "\r\n"
    });
    if (ws3_state._realFilePath) {
        postToAHK({
            action: 'saveCsvViaAHK',
            filePath: ws3_state._realFilePath,
            content: csvContent
        });
        return
    }
    try {
        var BOM = "\uFEFF";
        var handle = ws3_state.fileHandle;
        if (!handle) {
            handle = await window.showSaveFilePicker({
                suggestedName: ws3_state._cleanFileName || 'export.csv',
                types: [{
                        description: 'CSV File',
                        accept: {
                            'text/csv': ['.csv']
                        }
                    }
                ]
            });
            ws3_state.fileHandle = handle;
            ws3_state._cleanFileName = handle.name;
            ws3_state.fileName = handle.name;
        }
        var perm = await handle.queryPermission({
            mode: 'readwrite'
        });
        if (perm !== 'granted')
            perm = await handle.requestPermission({
                mode: 'readwrite'
            });
        if (perm !== 'granted')
            return toast("❌ Không có quyền ghi file!");
        var writable = await handle.createWritable();
        await writable.write(new Blob([BOM + csvContent], {
                type: 'text/csv;charset=utf-8'
            }));
        await writable.close();
        ws3_state.isModified = !1;
        _drawFileStatus();
        toast("✅ Đã lưu file!")
    } catch (err) {
        if (err.name !== 'AbortError')
            toast("❌ Lỗi lưu file: " + err.message);
    }
}
async function saveCsvTab3As() {
    if (!_currentState.data || _currentState.data.length === 0)
        return toast("Bảng trống!");
    try {
        var newHandle = await window.showSaveFilePicker({
            suggestedName: (ws3_state._cleanFileName) ? ('copy_of_' + ws3_state._cleanFileName) : 'export.csv',
            types: [{
                    description: 'CSV File',
                    accept: {
                        'text/csv': ['.csv']
                    }
                }
            ]
        });
        ws3_state.fileHandle = newHandle;
        ws3_state._cleanFileName = newHandle.name;
        ws3_state.fileName = newHandle.name;
        await saveCsvTab3()
    } catch (err) {
        if (err.name !== 'AbortError')
            toast("❌ Không thể lưu file: " + err.message);
    }
}
async function onImportCsvFromAHK(text, fileName, filePath) {
    _isImporting = !0;
    ws3_state.fileHandle = null;
    ws3_state._cleanFileName = fileName;
    ws3_state.fileName = fileName;
    ws3_state._realFilePath = filePath;
    ws3_state.isModified = !1;
    try {
        await parseCSVTextAndLoadTab3(text, fileName)
    } finally {
        _isImporting = !1;
        ws3_state.isModified = !1;
        _drawFileStatus()
    }
}
function markAsModified() {
    if (_isImporting)
        return;
    if (activeTabNum !== 3)
        return;
    ws3_state.isModified = !0;
    _drawFileStatus();
}
function _drawFileStatus() {
    var statusEl,
    state;
    if (activeTabNum === 1) {
        statusEl = document.getElementById('fileStatus');
        state = ws1_state;
    } else if (activeTabNum === 3) {
        statusEl = document.getElementById('fileStatus3');
        state = ws3_state;
    } else {
        return;
    }
    if (!statusEl || !state)
        return;

    var name = state.fileName || '';
    if (!name) {
        statusEl.textContent = '📄 Chưa import file';
        statusEl.style.color = 'var(--md-on-surface-3)';
        return;
    }

    var modifiedPrefix = state.isModified ? '● ' : '📄 ';
    statusEl.textContent = modifiedPrefix + name;
    statusEl.style.color = state.isModified ? '#d93025' : 'var(--md-on-surface-3)';
}
