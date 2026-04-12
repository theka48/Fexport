// ============================================
// NỘI DUNG MỚI CỦA FILE: 10-formula.js
// ============================================

function onFormulaList(formulas) {
    _formulaList = formulas || [];
    var menu = document.getElementById('formulaMenu');
    if (!menu)
        return;

    const selectedIdx = _formulaList.findIndex(f => f.name === document.getElementById('formulaDdlLabel').textContent);

    let html = '';
    for (var i = 0; i < _formulaList.length; i++) {
        html += `<div class="dropdown-item ${selectedIdx === i ? 'selected' : ''}" onclick="selectFormula(${i})">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="4 17 10 11 4 5"></polyline><line x1="12" y1="19" x2="20" y2="19"></line></svg>
                    ${escHtml(_formulaList[i].name)}
                 </div>`;
    }
    menu.innerHTML = html;
}

function selectFormula(idx) {
    const label = document.getElementById('formulaDdlLabel');
    if (label && _formulaList[idx]) {
        label.textContent = _formulaList[idx].name;
    }
}

function runFormula() {
    const label = document.getElementById('formulaDdlLabel').textContent;
    if (!label || label === 'Formula...') {
        toast('Chọn Formula!');
        return
    }
    const formula = _formulaList.find(f => f.name === label);
    if (!formula)
        return;
    _execFormula(formula);
}

function reloadFormulaList() {
    postToAHK({
        action: 'getFormulaList'
    });
    toast('Đang tải danh sách Formula...')
}

function selectDialog(options, title) {
    return new Promise(function (resolve) {
        var old = document.getElementById('_formulaDialog');
        if (old)
            old.remove();
        var overlay = document.createElement('div');
        overlay.id = '_formulaDialog';
        overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999;display:flex;align-items:center;justify-content:center;';
        var box = document.createElement('div');
        box.style.cssText = 'background:#fff;border-radius:10px;padding:20px 24px;min-width:280px;max-width:400px;box-shadow:0 8px 32px rgba(0,0,0,.25);font-family:Google Sans,sans-serif;';
        var h = document.createElement('div');
        h.style.cssText = 'font-size:13px;font-weight:600;color:#202124;margin-bottom:14px;';
        h.textContent = title || 'Chọn giá trị';
        box.appendChild(h);
        var sel = document.createElement('select');
        sel.style.cssText = 'width:100%;padding:7px 10px;border:1px solid #dadce0;border-radius:5px;font-size:13px;outline:none;background:#f8f9fa;color:#202124;cursor:pointer;margin-bottom:16px;';
        sel.onfocus = function () {
            this.style.borderColor = '#1a73e8';
            this.style.boxShadow = '0 0 0 2px #e8f0fe'
        };
        sel.onblur = function () {
            this.style.borderColor = '#dadce0';
            this.style.boxShadow = 'none'
        };
        for (var i = 0; i < options.length; i++) {
            var o = options[i],
            opt = document.createElement('option');
            opt.value = typeof o === 'object' ? o.value : o;
            opt.textContent = typeof o === 'object' ? o.label : o;
            sel.appendChild(opt)
        }
        box.appendChild(sel);
        var btnRow = document.createElement('div');
        btnRow.style.cssText = 'display:flex;justify-content:flex-end;gap:8px;';
        var btnCancel = document.createElement('button');
        btnCancel.textContent = 'Hủy';
        btnCancel.style.cssText = 'padding:7px 18px;border:1px solid #dadce0;border-radius:5px;background:#fff;color:#5f6368;font-size:12.5px;cursor:pointer;';
        btnCancel.onclick = function () {
            overlay.remove();
            resolve(null)
        };
        var btnOk = document.createElement('button');
        btnOk.textContent = 'OK';
        btnOk.style.cssText = 'padding:7px 22px;border:1px solid #1557b0;border-radius:5px;background:#1a73e8;color:#fff;font-size:12.5px;font-weight:600;cursor:pointer;';
        btnOk.onclick = function () {
            overlay.remove();
            resolve(sel.value)
        };
        overlay.onkeydown = function (e) {
            if (e.key === 'Enter') {
                btnOk.click();
                e.preventDefault()
            }
            if (e.key === 'Escape') {
                btnCancel.click();
                e.preventDefault()
            }
        };
        btnRow.appendChild(btnCancel);
        btnRow.appendChild(btnOk);
        box.appendChild(btnRow);
        overlay.appendChild(box);
        document.body.appendChild(overlay);
        sel.focus()
    })
}
function _execFormula(formula) {
    saveState();
    var ctx = {
        data: _currentState.data,
        headers: _currentState.headers,
        selRange: selRange,
        selCells: selCells,
        selectedRows: selectedRows,
        selectedCols: selectedCols,
        currentCell: currentCell,
        selectDialog: selectDialog,
        toast: toast,
        confirm: async function (message) {
            const result = await Swal.fire({
                title: 'Xác nhận',
                text: message,
                icon: 'question',
                showCancelButton: true,
                confirmButtonColor: '#3085d6',
                cancelButtonColor: '#d33',
                confirmButtonText: 'Đồng ý',
                cancelButtonText: 'Hủy'
            });
            return result.isConfirmed;
        },
        setData: function (d, h) {
            if (h)
                _currentState.headers = h;
            _currentState.data = d;
            _currentState.manualRowHeights = {};
            calcAllRowHeights();
            rebuildPrefixSums();
            rebuildDisplayCache();
            fullRender();
        },
        refresh: function () {
            _currentState.manualRowHeights = {};
            calcAllRowHeights();
            rebuildPrefixSums();
            rebuildDisplayCache();
            fullRender();
        },
        mapSelected: function (cb) {
            var changedCells = [];
            if (selRange) {
                for (var r = selRange.r1; r <= selRange.r2; r++) {
                    for (var c = selRange.c1; c <= selRange.c2; c++) {
                        var ri = r - 1,
                        ci = c - 1;
                        if (_currentState.data[ri]) {
                            var origVal = _currentState.data[ri][ci];
                            var res = cb(ri, ci, origVal);
                            if (res !== undefined && String(res) !== origVal) {
                                _currentState.data[ri][ci] = String(res);
                                changedCells.push({
                                    r: ri,
                                    c: ci
                                });
                            }
                        }
                    }
                }
            } else if (selCells) {
                selCells.forEach(function (k) {
                    var p = k.split(','),
                    ri = parseInt(p[0]) - 1,
                    ci = parseInt(p[1]) - 1;
                    if (_currentState.data[ri]) {
                        var origVal = _currentState.data[ri][ci];
                        var res = cb(ri, ci, origVal);
                        if (res !== undefined && String(res) !== origVal) {
                            _currentState.data[ri][ci] = String(res);
                            changedCells.push({
                                r: ri,
                                c: ci
                            });
                        }
                    }
                })
            }
            for (var i = 0; i < changedCells.length; i++) {
                updateCellDisplayCache(changedCells[i].r, changedCells[i].c);
            }
        },
        mapRows: function (cb) {
            var rows = selectedRows.length > 0 ? selectedRows : (selRange ? Array.from({
                        length: selRange.r2 - selRange.r1 + 1
                    }, function (_, i) {
                        return selRange.r1 + i
                    }) : []);
            for (var i = 0; i < rows.length; i++) {
                var ri = rows[i] - 1;
                if (_currentState.data[ri]) {
                    var res = cb(ri, _currentState.data[ri].slice());
                    if (res !== undefined)
                        _currentState.data[ri] = res
                }
            }
        },
        colByName: function (name) {
            var n = (name || '').toLowerCase();
            for (var i = 0; i < _currentState.headers.length; i++)
                if ((_currentState.headers[i] || '').toLowerCase() === n)
                    return i;
            return -1
        },
        addColumn: function (headerName, def) {
            _currentState.headers.push(headerName || 'Column ' + (_currentState.headers.length + 1));
            _currentState.colWidths.push(120);
            for (var r = 0; r < _currentState.data.length; r++)
                _currentState.data[r].push(def !== undefined ? String(def) : '');
        }
    };
    try {
        var keys = Object.keys(ctx),
        vals = keys.map(function (k) {
            return ctx[k]
        });
        var fn = new(Function.prototype.bind.apply(Function, [null].concat(keys).concat([formula.content])))();
        var result = fn.apply(ctx, vals);
        if (result && typeof result.then === 'function') {
            result.then(function () {
                rebuildDisplayCache();
                fullRender();
                toast('✅ Hoàn tất');
            }).catch(function (e) {
                undo();
                toast('❌ Lỗi: ' + e.message);
                console.error(e);
            });
        } else {
            rebuildDisplayCache();
            fullRender();
            toast('✅ Hoàn tất');
        }
    } catch (e) {
        undo();
        toast('❌ Lỗi: ' + e.message);
        console.error(e);
    }
}
