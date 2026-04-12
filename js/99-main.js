// ============================================
// NỘI DUNG MỚI CỦA FILE: 99-main.js
// ============================================

function toggleLayerDetach() {
    _layerFloating ? dockLayerPane() : detachLayerPane()
}
function detachLayerPane() {
    var pane = document.getElementById('layerPane');
    var btn = document.getElementById('layerDetachBtn');
    var divider = document.getElementById('divider');
    var rect = pane.getBoundingClientRect();
    var ph = document.createElement('div');
    ph.id = 'layerPlaceholder';
    pane.parentNode.insertBefore(ph, pane);
    document.body.appendChild(pane);
    pane.classList.add('floating');
    pane.style.left = rect.left + 'px';
    pane.style.top = rect.top + 'px';
    pane.style.width = rect.width + 'px';
    pane.style.height = rect.height + 'px';
    if (divider)
        divider.style.display = 'none';
    btn.textContent = '⊟';
    btn.title = 'Gắn lại vào cửa sổ';
    _layerFloating = !0
}
function dockLayerPane() {
    var pane = document.getElementById('layerPane');
    var btn = document.getElementById('layerDetachBtn');
    var ph = document.getElementById('layerPlaceholder');
    var divider = document.getElementById('divider');
    pane.classList.remove('floating');
    pane.style.left = '';
    pane.style.top = '';
    pane.style.width = '';
    pane.style.height = '';
    if (ph) {
        ph.parentNode.insertBefore(pane, ph);
        ph.remove()
    }
    if (divider)
        divider.style.display = '';
    btn.textContent = '⊞';
    btn.title = 'Tách panel';
    _layerFloating = !1
}
function setupLayerPaneDrag() {
    var header = document.getElementById('layerPaneHeader');
    if (!header)
        return;
    header.addEventListener('mousedown', function (e) {
        if (!_layerFloating)
            return;
        if (e.target.closest('.btn'))
            return;
        _layerDragging = !0;
        var pane = document.getElementById('layerPane');
        var rect = pane.getBoundingClientRect();
        _layerDragOffX = e.clientX - rect.left;
        _layerDragOffY = e.clientY - rect.top;
        e.preventDefault()
    });
    header.addEventListener('dblclick', function (e) {
        if (e.target.closest('.btn'))
            return;
        if (_layerFloating)
            dockLayerPane();
    });
    document.addEventListener('mousemove', function (e) {
        if (!_layerDragging)
            return;
        var pane = document.getElementById('layerPane');
        var x = e.clientX - _layerDragOffX;
        var y = e.clientY - _layerDragOffY;
        x = Math.max(0, Math.min(window.innerWidth - pane.offsetWidth, x));
        y = Math.max(0, Math.min(window.innerHeight - pane.offsetHeight, y));
        pane.style.left = x + 'px';
        pane.style.top = y + 'px'
    });
    document.addEventListener('mouseup', function () {
        _layerDragging = !1
    })
}
function setupDivider() {
    var divider = document.getElementById('divider'),
    gridPane = document.getElementById('gridPane'),
    layerPane = document.getElementById('layerPane');
    if (!divider)
        return;
    var dragging = !1,
    startX = 0,
    startLayerW = 0;
    divider.addEventListener('mousedown', function (e) {
        dragging = !0;
        startX = e.clientX;
        startLayerW = layerPane.offsetWidth;
        divider.classList.add('dragging');
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
        e.preventDefault()
    });
    document.addEventListener('mousemove', function (e) {
        if (!dragging)
            return;
        layerPane.style.width = layerPane.style.minWidth = layerPane.style.maxWidth = Math.max(180, Math.min(500, startLayerW + (startX - e.clientX))) + 'px'
    });
    document.addEventListener('mouseup', function () {
        if (!dragging)
            return;
        dragging = !1;
        divider.classList.remove('dragging');
        document.body.style.cursor = '';
        document.body.style.userSelect = ''
    })
}
function setupLayerDivider() {
    var divider = document.getElementById('layerDivider');
    if (!divider)
        return;
    var allList = document.getElementById('allLayersList'),
    activeList = document.getElementById('activeLayersList'),
    dragging = !1,
    startY = 0,
    startAllH = 0,
    startActiveH = 0;
    divider.addEventListener('mousedown', function (e) {
        dragging = !0;
        startY = e.clientY;
        startAllH = allList.offsetHeight;
        startActiveH = activeList.offsetHeight;
        divider.classList.add('dragging');
        document.body.style.cursor = 'row-resize';
        document.body.style.userSelect = 'none';
        e.preventDefault()
    });
    document.addEventListener('mousemove', function (e) {
        if (!dragging)
            return;
        var delta = e.clientY - startY,
        newAllH = Math.max(60, startAllH + delta),
        newActiveH = Math.max(60, startActiveH - delta),
        totalH = startAllH + startActiveH;
        if (newAllH + newActiveH > totalH) {
            if (delta > 0) {
                newAllH = totalH - 60;
                newActiveH = 60
            } else {
                newAllH = 60;
                newActiveH = totalH - 60
            }
        }
        allList.style.flex = 'none';
        activeList.style.flex = 'none';
        allList.style.height = newAllH + 'px';
        activeList.style.height = newActiveH + 'px'
    });
    document.addEventListener('mouseup', function () {
        if (!dragging)
            return;
        dragging = !1;
        divider.classList.remove('dragging');
        document.body.style.cursor = '';
        document.body.style.userSelect = ''
    })
}
function showCtx(x, y) {
    var m = document.getElementById('ctxMenu');
    m.style.left = x + 'px';
    m.style.top = y + 'px';
    var btnUnhideAdjacent = document.getElementById('ctxUnhideCol');
    if (btnUnhideAdjacent) {
        var hasHiddenAdjacent = !1;
        if (_currentState.hiddenCols.size > 0) {
            var checkCols = [];
            if (selectedCols.length > 0)
                checkCols = selectedCols;
            else if (selRange) {
                for (var c = selRange.c1; c <= selRange.c2; c++)
                    checkCols.push(c);
            } else if (currentCell)
                checkCols.push(currentCell.col);
            for (var i = 0; i < checkCols.length; i++) {
                var cIdx = checkCols[i] - 1;
                if (cIdx > 0 && _currentState.hiddenCols.has(cIdx - 1)) {
                    hasHiddenAdjacent = !0;
                    break
                }
                if (cIdx < _currentState.headers.length - 1 && _currentState.hiddenCols.has(cIdx + 1)) {
                    hasHiddenAdjacent = !0;
                    break
                }
            }
        }
        btnUnhideAdjacent.style.display = hasHiddenAdjacent ? 'flex' : 'none'
    }
    m.classList.add('show');
    if (x + m.offsetWidth > window.innerWidth)
        m.style.left = Math.max(0, x - m.offsetWidth) + 'px';
    if (y + m.offsetHeight > window.innerHeight)
        m.style.top = Math.max(0, y - m.offsetHeight) + 'px'
}
function hideCtx() {
    document.getElementById('ctxMenu').classList.remove('show')
}
function setupEvents() {
    var gw = document.getElementById('gridWrapper'),
    ticking = !1;
    gw.addEventListener('scroll', function () {
        if (!ticking) {
            ticking = !0;
            requestAnimationFrame(function () {
                renderVisibleRows();
                updateInfo();
                ticking = !1
            })
        }
    }, {
        passive: !0
    });
    gw.addEventListener('wheel', function (e) {
        if (e.ctrlKey) {
            e.preventDefault();
            e.deltaY < 0 ? zoomIn() : zoomOut()
        }
    }, {
        passive: !1
    });
    document.addEventListener('mousedown', function (e) {
        var ae = document.activeElement;
        if (ae && ae.tagName === 'INPUT' && ae.type === 'number' && !e.target.closest('input'))
            ae.blur();
        if (!e.target.closest('.search-panel') && ae && (ae.id === 'searchInput' || ae.id === 'replaceInput'))
            ae.blur();
        if (e.target.classList.contains('col-resize')) {
            if (e.detail >= 2) {
                e.preventDefault();
                return
            }
            isResizingCol = !0;
            resizeColIndex = parseInt(e.target.getAttribute('data-col'));
            resizeStartX = e.pageX;
            resizeStartWidth = _currentState.colWidths[resizeColIndex - 1] || 120;
            resizingMultiCols = [];
            if (selectedCols.length > 1 && selectedCols.indexOf(resizeColIndex) >= 0) {
                resizingMultiCols = selectedCols.slice();
                resizeMultiStartWidths = {};
                for (var i = 0; i < selectedCols.length; i++)
                    resizeMultiStartWidths[selectedCols[i]] = _currentState.colWidths[selectedCols[i] - 1] || 120
            }
            e.target.classList.add('active');
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';
            e.preventDefault();
            return
        }
        if (e.target.classList.contains('row-resize-handle')) {
            if (e.detail >= 2) {
                e.preventDefault();
                return
            }
            isResizingRow = !0;
            resizeRowIndex = parseInt(e.target.getAttribute('data-row'));
            resizeStartY = e.pageY;
            resizeStartHeight = _currentState.rowHeights[resizeRowIndex - 1] || MIN_ROW_HEIGHT;
            resizingMultiRows = [];
            if (selectedRows.length > 1 && selectedRows.indexOf(resizeRowIndex) >= 0) {
                resizingMultiRows = selectedRows.slice();
                resizeMultiStartHeights = {};
                for (var i = 0; i < selectedRows.length; i++)
                    resizeMultiStartHeights[selectedRows[i]] = _currentState.rowHeights[selectedRows[i] - 1] || MIN_ROW_HEIGHT
            }
            e.target.classList.add('active');
            document.body.style.cursor = 'row-resize';
            document.body.style.userSelect = 'none';
            e.preventDefault();
            return
        }
        var ch = e.target.closest('.col-header');
        if (ch && !e.target.classList.contains('col-resize')) {
            var cn = parseInt(ch.getAttribute('data-col'));
            if (e.button === 2) {
                if (selectedCols.indexOf(cn) < 0)
                    handleColSelect(cn, !1, !1);
                return
            }
            handleColSelect(cn, e.ctrlKey || e.metaKey, e.shiftKey);
            e.preventDefault();
            return
        }
        var rh = e.target.closest('.row-header');
        if (rh) {
            var rn = parseInt(rh.getAttribute('data-row'));
            if (e.button === 2) {
                if (selectedRows.indexOf(rn) < 0)
                    handleRowSelect(rn, !1, !1);
                return
            }
            handleRowSelect(rn, e.ctrlKey || e.metaKey, e.shiftKey);
            triggerAiPreview(rn);
            e.preventDefault();
            return
        }
        var dc = e.target.closest('.data-cell');
        if (dc) {
            var info = getCellInfo(dc);
            if (editingCell) {
                if (editingCell !== dc)
                    exitEditMode();
                else
                    return
            }
            if (e.button === 2) {
                if (!isCellSelected(info.row, info.col)) {
                    clearSel();
                    selectedRows = [];
                    selectedCols = [];
                    selCells = null;
                    selRange = null;
                    currentCell = info;
                    anchorCell = info;
                    dc.classList.add('current')
                }
                return
            }
            if (e.shiftKey && anchorCell) {
                setSelRange(anchorCell, info);
                currentCell = info;
                return
            }
            if (e.ctrlKey || e.metaKey) {
                selectionType = 'ctrl';
                var overlay = document.getElementById('selectionOverlay');
                if (overlay)
                    overlay.style.display = 'none';
                if (!selCells) {
                    selCells = new Set();
                    if (selRange) {
                        for (var r = selRange.r1; r <= selRange.r2; r++) {
                            if (isRowHidden(r - 1) || isRowHiddenBySearch(r - 1))
                                continue;
                            for (var c = selRange.c1; c <= selRange.c2; c++) {
                                if (!_currentState.hiddenCols.has(c - 1))
                                    selCells.add(r + ',' + c);
                            }
                        }
                        selRange = null
                    } else if (currentCell) {
                        selCells.add(currentCell.row + ',' + currentCell.col)
                    }
                    selectedRows = [];
                    selectedCols = []
                }
                var key = info.row + ',' + info.col;
                if (selCells.has(key)) {
                    selCells.delete(key);
                    if (currentCell && currentCell.row === info.row && currentCell.col === info.col) {
                        currentCell = null
                    }
                } else {
                    selCells.add(key);
                    currentCell = info
                }
                anchorCell = info;
                renderVisibleRows();
                e.preventDefault();
                return
            }
            isDragging = !0;
            dragStartCell = info;
            clearSel();
            selectedRows = [];
            selectedCols = [];
            selCells = null;
            selRange = null;
            currentCell = info;
            anchorCell = info;
            dc.classList.add('current');
            if (activeTabNum === 1) {
                _currentState.currentRowSel = info.row;
                document.getElementById('activeRowNum').textContent = info.row;
                refreshActiveLayersList(_currentState.rowLayerData[info.row] || []);
                renderAllLayersList_forRow(info.row);
            }
            updateHeaderHighlight();
            triggerAiPreview(info.row);
            return
        }
        if (!e.target.closest('#gridWrapper') && !e.target.closest('.ctx-menu') && !e.target.closest('.search-panel') && !e.target.closest('.modal-overlay') && !e.target.closest('.import-split-btn') && !e.target.closest('.layer-pane') && !e.target.closest('.toolbar') && !e.target.closest('#_aiExportDialog') && !e.target.closest('#_aiReportDialog') && !e.target.closest('#_dorDialog') && !e.target.closest('#_sideDialog')) {
            if (editingCell)
                exitEditMode();
            clearSelFull();
            selectedRows = [];
            selectedCols = [];
            renderVisibleRows()
        }
    });
    document.addEventListener('mousemove', function (e) {
        if (isResizingCol) {
            var delta = e.pageX - resizeStartX;
            if (resizingMultiCols.length > 0) {
                for (var i = 0; i < resizingMultiCols.length; i++)
                    _currentState.colWidths[resizingMultiCols[i] - 1] = Math.max(MIN_COL_WIDTH, (resizeMultiStartWidths[resizingMultiCols[i]] || 120) + delta);
            } else {
                _currentState.colWidths[resizeColIndex - 1] = Math.max(MIN_COL_WIDTH, resizeStartWidth + delta)
            }
            rebuildPrefixSums();
            renderVisibleRows();
            renderHeader();
            return
        }
        if (isResizingRow) {
            var delta = e.pageY - resizeStartY;
            if (resizingMultiRows.length > 0) {
                for (var i = 0; i < resizingMultiRows.length; i++) {
                    var nh = Math.max(MIN_ROW_HEIGHT, (resizeMultiStartHeights[resizingMultiRows[i]] || MIN_ROW_HEIGHT) + delta);
                    _currentState.rowHeights[resizingMultiRows[i] - 1] = nh;
                    _currentState.manualRowHeights[resizingMultiRows[i] - 1] = nh
                }
            } else {
                var nh = Math.max(MIN_ROW_HEIGHT, resizeStartHeight + delta);
                _currentState.rowHeights[resizeRowIndex - 1] = nh;
                _currentState.manualRowHeights[resizeRowIndex - 1] = nh
            }
            rebuildPrefixSums();
            renderVisibleRows();
            return
        }
        if (isDragging && dragStartCell && !editingCell) {
            var dc = e.target.closest('.data-cell');
            if (dc) {
                var info = getCellInfo(dc);
                if (!dragCurrentCell || dragCurrentCell.row !== info.row || dragCurrentCell.col !== info.col) {
                    dragCurrentCell = info;
                    setSelRange(dragStartCell, dragCurrentCell);
                    updateHeaderHighlight()
                }
            }
        }
    });
    document.addEventListener('mouseup', function (e) {
        if (isResizingCol) {
            isResizingCol = !1;
            document.querySelectorAll('.col-resize.active').forEach(function (el) {
                el.classList.remove('active')
            });
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            justFinishedResize = !0;
            setTimeout(function () {
                justFinishedResize = !1
            }, 50);
            recalcAllRowHeights();
            rebuildPrefixSums();
            fullRender();
            return
        }
        if (isResizingRow) {
            isResizingRow = !1;
            document.querySelectorAll('.row-resize-handle.active').forEach(function (el) {
                el.classList.remove('active')
            });
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            justFinishedResize = !0;
            setTimeout(function () {
                justFinishedResize = !1
            }, 50);
            return
        }
        if (isDragging) {
            isDragging = !1;
            dragCurrentCell = null;
            autoUpdateSearchScope();
            return
        }
        isDragging = !1;
        dragCurrentCell = null
    });
    document.addEventListener('dblclick', function (e) {
        if (justFinishedResize)
            return;
        var dc = e.target.closest('.data-cell');
        if (dc) {
            enterEditMode(dc);
            e.preventDefault();
            return
        }
        if (e.target.classList.contains('col-resize')) {
            var cIdx = parseInt(e.target.getAttribute('data-col')) - 1,
            maxW = MIN_COL_WIDTH,
            nw = measureTextWidth(String(cIdx + 1));
            if (nw > maxW)
                maxW = nw;
            for (var r = 0; r < _currentState.data.length; r++) {
                var v = _currentState.data[r][cIdx];
                if (!v)
                    continue;
                var w = measureTextWidth(v);
                if (w > maxW)
                    maxW = w
            }
            _currentState.colWidths[cIdx] = Math.min(maxW, 800);
            rebuildPrefixSums();
            recalcAllRowHeights();
            fullRender();
            return
        }
        if (e.target.classList.contains('row-resize-handle')) {
            var ri = parseInt(e.target.getAttribute('data-row')) - 1;
            delete _currentState.manualRowHeights[ri];
            _currentState.rowHeights[ri] = calcRowHeight(ri);
            rebuildPrefixSums();
            renderVisibleRows();
            return
        }
    });
    document.addEventListener('contextmenu', function (e) {
        var tg = e.target.closest('.data-cell, .row-header, .col-header');
        if (tg) {
            e.preventDefault();
            showCtx(e.clientX, e.clientY)
        }
    });
    document.addEventListener('click', function(e) {
        if (!e.target.closest('.ctx-menu')) hideCtx();
        
        if (!e.target.closest('.dropdown-wrapper')) {
            document.querySelectorAll('.dropdown-menu.show').forEach(menu => {
                menu.classList.remove('show');
            });
        }
    });
    document.addEventListener('keydown', function (e) {
        if (e.ctrlKey && e.key.toLowerCase() === 's') {
            if (activeTabNum === 2) {
                e.preventDefault();
                saveRules();
                return
            }
            if (activeTabNum === 3) {
                e.preventDefault();
                saveCsvTab3();
                return
            }
        }
        if (activeTabNum === 2) {
            if (e.ctrlKey) {
                var k2 = e.key.toLowerCase();
                if (k2 === 'z' || k2 === 'y')
                    e.preventDefault();
            }
            return
        }
        var ae = document.activeElement;
        if (ae && ae.closest && ae.closest('.modal-overlay'))
            return;
        if (ae && ae.tagName === 'INPUT' && ae.type === 'number') {
            if (e.ctrlKey && e.key.toLowerCase() === 'a') {
                e.stopPropagation();
                return
            }
            if (e.key === 'Escape') {
                ae.blur();
                e.preventDefault()
            }
            return
        }
        if (ae && (ae.id === 'searchInput' || ae.id === 'replaceInput' || ae.id === 'excludeInput')) {
            if (e.key === 'Escape') {
                toggleSearchPanel();
                e.preventDefault()
            }
            return
        }
        if (editingCell) {
            if (e.key === 'Escape') {
                exitEditMode();
                e.preventDefault()
            } else if (e.key === 'Enter' && !e.shiftKey) {
                exitEditMode();
                moveCurrentCell(0, 1);
                e.preventDefault()
            } else if (e.key === 'Tab') {
                exitEditMode();
                moveCurrentCell(e.shiftKey ? -1 : 1, 0);
                e.preventDefault()
            }
            return
        }
        if (e.key === 'ArrowUp') {
            e.shiftKey ? expandSelection(0, -1) : moveCurrentCell(0, -1);
            e.preventDefault()
        } else if (e.key === 'ArrowDown') {
            e.shiftKey ? expandSelection(0, 1) : moveCurrentCell(0, 1);
            e.preventDefault()
        } else if (e.key === 'ArrowLeft') {
            e.shiftKey ? expandSelection(-1, 0) : moveCurrentCell(-1, 0);
            e.preventDefault()
        } else if (e.key === 'ArrowRight') {
            e.shiftKey ? expandSelection(1, 0) : moveCurrentCell(1, 0);
            e.preventDefault()
        } else if (e.key === 'Tab') {
            moveCurrentCell(e.shiftKey ? -1 : 1, 0);
            e.preventDefault()
        } else if (e.key === 'Enter') {
            if (currentCell)
                enterEditModeByInfo(currentCell);
            e.preventDefault()
        } else if (e.key === 'Delete' || e.key === 'Backspace') {
            deleteContent();
            e.preventDefault()
        } else if (e.key === 'F2') {
            if (currentCell)
                enterEditModeByInfo(currentCell);
            e.preventDefault()
        } else if (e.ctrlKey) {
            var k = e.key.toLowerCase();
            if (k === 'c') {
                var selText = window.getSelection().toString();
                if (selText.length > 0 && !e.target.closest('#gridWrapper')) {}
                else {
                    copySelection();
                    e.preventDefault()
                }
            } else if (k === 'x') {
                cutSelection();
                e.preventDefault()
            } else if (k === 'z') {
                undo();
                e.preventDefault()
            } else if (k === 'y') {
                redo();
                e.preventDefault()
            } else if (k === 'a') {
                if (e.target.closest('.layer-pane')) {}
                else {
                    selectAll();
                    e.preventDefault()
                }
            } else if (k === 'f') {
                toggleSearchPanel();
                setTimeout(function () {
                    var si = document.getElementById('searchInput');
                    if (si) {
                        si.focus();
                        si.select()
                    }
                }, 100);
                e.preventDefault()
            } else if (k === '0') {
                zoomReset();
                e.preventDefault()
            } else if (k === '=' || k === '+') {
                zoomIn();
                e.preventDefault()
            } else if (k === '-') {
                zoomOut();
                e.preventDefault()
            }
        } else if (e.key === 'Escape') {
            if (searchTerm)
                clearSearch();
            else {
                clearSelFull();
                selectedRows = [];
                selectedCols = [];
                fullRender()
            }
            e.preventDefault()
        } else if (e.key.length === 1 && !e.ctrlKey && !e.altKey && currentCell) {
            enterEditModeByInfo(currentCell);
            var cc = currentCell;
            setTimeout(function () {
                var c = document.querySelector('.data-cell[data-row="' + cc.row + '"][data-col="' + cc.col + '"]');
                if (c) {
                    c.innerText = e.key;
                    moveCursorToEnd(c)
                }
            }, 0)
        }
    });
    document.addEventListener('paste', function (e) {
        if (editingCell || (document.activeElement && ['searchInput', 'replaceInput', 'excludeInput'].indexOf(document.activeElement.id) >= 0) || !currentCell)
            return;
        var cd = e.clipboardData || window.clipboardData;
        e.preventDefault();
        saveState();
        processPasteData(cd.getData('text/plain') || '', cd.getData('text/html') || '')
    })
}
document.addEventListener('DOMContentLoaded', function () {
    // Khởi tạo colWidths cho cả 2 workspace
    for (var i = 0; i < ws1_state.headers.length; i++)
        ws1_state.colWidths.push(120);
    for (var i = 0; i < ws3_state.headers.length; i++)
        ws3_state.colWidths.push(120);

    createMeasureDivs();
    calcAllRowHeights(); // Sẽ tính cho _currentState (mặc định là ws1_state)
    rebuildPrefixSums();
    fullRender();
    setupEvents();
    setupDivider();
    setupLayerDivider();
    setupRulesTableResize();
    setupLayerPaneDrag()
});
window.addEventListener('beforeunload', function (e) {
    if (ws3_state && ws3_state.isModified) {
        e.preventDefault();
        e.returnValue = '';
        return ''
    }
});
function setupRulesTableResize() {
    var rulesTable = document.getElementById('rulesTable');
    if (!rulesTable)
        return;
    var isResizing = !1;
    var resizeColIndex = -1;
    var resizeStartX = 0;
    var resizeStartWidth = 0;
    var thElements = rulesTable.querySelectorAll('th');
    rulesTable.addEventListener('mousedown', function (e) {
        if (e.target.classList.contains('rules-col-resize')) {
            isResizing = !0;
            resizeColIndex = parseInt(e.target.getAttribute('data-col'));
            resizeStartX = e.pageX;
            var currentStyleWidth = thElements[resizeColIndex].style.width;
            resizeStartWidth = parseInt(currentStyleWidth, 10) || thElements[resizeColIndex].offsetWidth;
            e.target.classList.add('active');
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';
            e.preventDefault()
        }
    });
    document.addEventListener('mousemove', function (e) {
        if (!isResizing)
            return;
        var delta = e.pageX - resizeStartX;
        var newWidth = Math.max(10, resizeStartWidth + delta);
        rulesColWidths[resizeColIndex] = newWidth;
        thElements[resizeColIndex].style.width = newWidth + 'px'
    });
    document.addEventListener('mouseup', function () {
        if (!isResizing)
            return;
        isResizing = !1;
        var activeHandle = document.querySelector('.rules-col-resize.active');
        if (activeHandle)
            activeHandle.classList.remove('active');
        document.body.style.cursor = '';
        document.body.style.userSelect = ''
    })
}
function toggleDropdownMenu(btn, menuId) {
    // Đóng tất cả các menu khác
    document.querySelectorAll('.dropdown-menu.show').forEach(openMenu => {
        if (openMenu.id !== menuId) {
            openMenu.classList.remove('show');
        }
    });
    // Toggle menu hiện tại
    const menu = document.getElementById(menuId);
    if (menu) {
        menu.classList.toggle('show');
    }
}
