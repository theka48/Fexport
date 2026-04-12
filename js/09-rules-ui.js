// ============================================
// NỘI DUNG MỚI CỦA FILE: 09-rules-ui.js
// ============================================

var svgEyeOpen = '<svg viewBox="0 0 24 24"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>';
var svgEyeClosed = '<svg viewBox="0 0 24 24"><path d="M12 7c2.76 0 5 2.24 5 5 0 .65-.13 1.26-.36 1.83l2.92 2.92c1.51-1.26 2.7-2.89 3.43-4.75-1.73-4.39-6-7.5-11-7.5-1.4 0-2.74.25-3.98.7l2.16 2.16C10.74 7.13 11.35 7 12 7zM2 4.27l2.28 2.28.46.46C3.08 8.3 1.78 10.02 1 12c1.73 4.39 6 7.5 11 7.5 1.55 0 3.03-.3 4.38-.84l.42.42L19.73 22 21 20.73 3.27 3 2 4.27zM7.53 9.8l1.55 1.55c-.05.21-.08.43-.08.65 0 1.66 1.34 3 3 3 .22 0 .44-.03.65-.08l1.55 1.55c-.67.33-1.41.53-2.2.53-2.76 0-5-2.24-5-5 0-.79.2-1.53.53-2.2zm4.31-.78l3.15 3.15.02-.16c0-1.66-1.34-3-3-3l-.17.01z"/></svg>';
function toggleEyeClick(btn, layerName) {
    if (!_currentState.currentRowSel) { toast('Chọn một dòng trước'); return; }
    var itemDiv = btn.closest('.layer-item');
    var isChecked = itemDiv.classList.contains('active-item');
    var newChecked = !isChecked;
    var layers = _currentState.rowLayerData[_currentState.currentRowSel] ? _currentState.rowLayerData[_currentState.currentRowSel].slice() : [];
    if (newChecked) { if (layers.indexOf(layerName) < 0) layers.push(layerName); } 
    else { var idx = layers.indexOf(layerName); if (idx >= 0) layers.splice(idx, 1); }
    _currentState.rowLayerData[_currentState.currentRowSel] = layers;
    postToAHK({ action: 'toggleLayer', row: _currentState.currentRowSel, layer: layerName, checked: newChecked.toString() });
    refreshActiveLayersList(layers);
    renderAllLayersList_forRow(_currentState.currentRowSel);
    renderVisibleRows();
}
function copyLayerName(text) {
    if (navigator.clipboard && window.isSecureContext) { navigator.clipboard.writeText(text); } 
    else { var ta = document.createElement('textarea'); ta.value = text; ta.style.cssText = 'position:fixed;left:-9999px'; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); } catch (e) {} document.body.removeChild(ta); }
    toast('📋 Đã copy: ' + text);
}
function onLayersImported(data) {
    allLayersFromAI = [];
    layerColors = {};
    if (data && data.length) {
        for (var i = 0; i < data.length; i++) {
            if (typeof data[i] === 'object') {
                allLayersFromAI.push(data[i].name);
                layerColors[data[i].name] = data[i].color;
            } else {
                allLayersFromAI.push(data[i]);
            }
        }
    }
    renderAllLayersList();
    toast('Đã import ' + allLayersFromAI.length + ' layers từ AI');
    if (_currentState.currentRowSel > 0)
        refreshActiveLayersList(_currentState.rowLayerData[_currentState.currentRowSel] || []);
}
function renderAllLayersList() { renderAllLayersList_forRow(_currentState.currentRowSel); }
function renderAllLayersList_forRow(rowNum) {
    var container = document.getElementById('allLayersList');
    if (!allLayersFromAI.length) { container.innerHTML = '<div class="layer-empty">Chưa import layers từ AI</div>'; return; }
    var activeLayers = rowNum > 0 ? (_currentState.rowLayerData[rowNum] || []) : [];
    var html = '';
    for (var i = 0; i < allLayersFromAI.length; i++) {
        var name = allLayersFromAI[i];
        var checked = activeLayers.indexOf(name) >= 0;
        var c = layerColors[name] || '200,200,200';
        var colorDot = '<span style="width:12px;height:12px;min-width:12px;background:rgb(' + c + ');border-radius:3px;display:inline-block;border:1px solid rgba(0,0,0,0.15);"></span>';
        html += '<div class="layer-item' + (checked ? ' active-item' : '') + '">';
        html += '<span class="layer-eye-btn" title="Bật/Tắt Layer" onclick="toggleEyeClick(this, \'' + escHtml(name).replace(/'/g, "\\'") + '\')">';
        html += (checked ? svgEyeOpen : svgEyeClosed) + '</span>';
        html += colorDot + '&nbsp;';
        html += '<span style="user-select:text; cursor:text;">' + escHtml(name) + '</span>';
        html += '</div>';
    }
    container.innerHTML = html;
}
function refreshActiveLayersList(activeLayers) {
    var container = document.getElementById('activeLayersList');
    if (!activeLayers || !activeLayers.length) { container.innerHTML = '<div class="layer-empty">Không có layer nào</div>'; return; }
    var ordered = [];
    for (var i = 0; i < allLayersFromAI.length; i++) { if (activeLayers.indexOf(allLayersFromAI[i]) >= 0) { ordered.push(allLayersFromAI[i]); } }
    for (var j = 0; j < activeLayers.length; j++) { if (allLayersFromAI.indexOf(activeLayers[j]) < 0) { ordered.push(activeLayers[j]); } }
    var html = '';
    for (var k = 0; k < ordered.length; k++) {
        var name = ordered[k];
        var c = layerColors[name] || '200,200,200';
        var colorDot = '<span style="width:12px;height:12px;min-width:12px;background:rgb(' + c + ');border-radius:3px;display:inline-block;border:1px solid rgba(0,0,0,0.15);"></span>';
        html += '<div class="layer-item active-item">';
        html += '<span class="layer-eye-btn" title="Tắt Layer" onclick="toggleEyeClick(this, \'' + escHtml(name).replace(/'/g, "\\'") + '\')">' + svgEyeOpen + '</span>';
        html += colorDot + '&nbsp;';
        html += '<span style="user-select:text; cursor:text;">' + escHtml(name) + '</span>';
        html += '</div>';
    }
    container.innerHTML = html;
}
function onRowLayersLoaded(layers) {
    if (_currentState.currentRowSel > 0) _currentState.rowLayerData[_currentState.currentRowSel] = layers || [];
    refreshActiveLayersList(layers || []);
    renderAllLayersList();
}
function clearActiveRowLayers() {
    var rowsToClear = new Set();
    if (selectedRows.length > 0) { selectedRows.forEach(function (r) { rowsToClear.add(r); }); } 
    else if (selRange) { for (var r = selRange.r1; r <= selRange.r2; r++) { if (!isRowHidden(r - 1) && !isRowHiddenBySearch(r - 1)) rowsToClear.add(r); } }
    else if (selCells && selCells.size > 0) { selCells.forEach(function (k) { var r = parseInt(k.split(',')[0]); if (!isRowHidden(r - 1) && !isRowHiddenBySearch(r - 1)) rowsToClear.add(r); }); }
    else if (currentCell) { rowsToClear.add(currentCell.row); }
    if (rowsToClear.size === 0) { toast('Chọn ít nhất 1 dòng để xóa'); return; }
    var arrRows = Array.from(rowsToClear);
    for (var i = 0; i < arrRows.length; i++) {
        var r = arrRows[i];
        _currentState.rowLayerData[r] = [];
        postToAHK({ action: 'clearRowLayers', row: r });
    }
    refreshActiveLayersList([]);
    renderAllLayersList();
    renderVisibleRows();
    toast('🗑️ Đã xóa layer của ' + arrRows.length + ' dòng');
}
function hasActiveLayers(rowNum) {
    if (activeTabNum !== 1) return false;
    var l = _currentState.rowLayerData[rowNum];
    return l && l.length > 0;
}
function loadRulesFromDdlTab1() { var ddl = document.getElementById('rulesDdlTab1'); if (ddl && ddl.value) { if (document.getElementById('rulesDdl')) document.getElementById('rulesDdl').value = ddl.value; postToAHK({ action: 'loadRulesFile', file: ddl.value }); } }
function loadRulesFromDdl() { var e = document.getElementById("rulesDdl"); if (e && e.value) { postToAHK({ action: "loadRulesFile", file: e.value }); } else { currentRules = []; rulesSelectedRow = -1; renderRulesTable(); } }
function onRulesFileList(e) {
    if (typeof e === 'string') try { e = JSON.parse(e) } catch (t) { e = [] }
    
    // ---> HÀM MỚI ĐỂ RENDER DROPDOWN RULES <---
    function renderRulesDropdown(menuId, selectedFile) {
        const menu = document.getElementById(menuId);
        if (!menu) return;

        let html = `<div class="dropdown-item ${!selectedFile ? 'selected' : ''}" onclick="selectRulesFile('')">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="4.93" y1="4.93" x2="19.07" y2="19.07"></line></svg>
                        <em>(Không dùng)</em>
                    </div>
                    <div class="dropdown-sep"></div>`;

        (e || []).forEach(file => {
            html += `<div class="dropdown-item ${selectedFile === file ? 'selected' : ''}" onclick="selectRulesFile('${file}')">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="8" y1="6" x2="21" y2="6"></line><line x1="8" y1="12" x2="21" y2="12"></line><line x1="8" y1="18" x2="21" y2="18"></line><line x1="3" y1="6" x2="3.01" y2="6"></line><line x1="3" y1="12" x2="3.01" y2="12"></line><line x1="3" y1="18" x2="3.01" y2="18"></line></svg>
                        ${escHtml(file)}
                    </div>`;
        });
        html += `<div class="dropdown-sep"></div>
                 <div class="dropdown-item" onclick="postToAHK({action:&#34;browseRulesFile&#34;})">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path></svg>
                    Mở file khác...
                 </div>`;
        menu.innerHTML = html;
    }
    
    // Cập nhật cả 2 menu
    const selectedFile = document.getElementById('rulesDdlLabel')?.textContent || document.getElementById('rulesDdlLabelTab1')?.textContent;
    renderRulesDropdown('rulesMenuTab1', selectedFile);
    renderRulesDropdown('rulesMenu', selectedFile); // Tab 2 cũng có menu
}
function selectRulesFile(fileName) {
    const label1 = document.getElementById('rulesDdlLabelTab1');
    const label2 = document.getElementById('rulesDdlLabel'); // Tab 2
    if (fileName) {
        if(label1) label1.textContent = fileName;
        if(label2) label2.textContent = fileName;
        postToAHK({ action: "loadRulesFile", file: fileName });
    } else {
        if(label1) label1.textContent = "Chọn Rules...";
        if(label2) label2.textContent = "Chọn Rules...";
        currentRules = [];
        rulesSelectedRow = -1;
        renderRulesTable();
    }
}
function onRulesLoaded(rules) { if (typeof rules === 'string') { try { rules = JSON.parse(rules); } catch (e) { rules = []; } } currentRules = rules || []; renderRulesTable(); }
function renderRulesTable() { var ths = document.querySelectorAll("#rulesTable th"); if (ths.length === rulesColWidths.length) { ths.forEach(function (th, i) { th.style.width = rulesColWidths[i] + 'px'; }); } var tbody = document.getElementById('rulesBody'); if (!tbody) return; var html = ''; for (var i = 0; i < currentRules.length; i++) { var r = currentRules[i], en = r.Enabled === '1'; var color = r.Color || ''; var rowStyle = color ? 'background:' + color + ';' : ''; var dotStyle = color ? 'background:' + color + ';border:2px solid rgba(0,0,0,0.2);' : 'background:#e8eaed;border:2px solid #bdc1c6;'; html += '<tr class="' + (en ? '' : 'rule-disabled') + ' ' + (rulesSelectedRow === i ? 'selected' : '') + '" data-idx="' + i + '" style="' + rowStyle + '" onclick="selectRulesRow(' + i + ')" ondblclick="editRule()">'; html += '<td class="rule-check"><input type="checkbox" class="rules-chk" ' + (en ? 'checked' : '') + ' onchange="toggleRuleEnabled(' + i + ',this.checked)" onclick="event.stopPropagation()"></td>'; html += '<td style="text-align:center;padding:2px;">'; html += '<span onclick="openColorPicker(' + i + ',this,event)" style="display:inline-block;width:16px;height:16px;border-radius:50%;cursor:pointer;' + dotStyle + '"></span>'; html += '</td>'; html += '<td>' + escHtml(r.Priority) + '</td>'; html += '<td>' + escHtml(r.RuleID) + '</td>'; html += '<td><span class="rule-type-badge">' + escHtml(r.Type) + '</span></td>'; html += '<td>' + escHtml(r.Column) + '</td>'; html += '<td>' + escHtml(r.Match) + '</td>'; html += '<td>' + escHtml(r.Output) + '</td>'; html += '<td>' + escHtml(r.Options) + '</td>'; html += '<td>' + escHtml(r.Description) + '</td>'; html += '</tr>'; } tbody.innerHTML = html; }
function selectRulesRow(idx) { rulesSelectedRow = idx; renderRulesTable(); }
function toggleRuleEnabled(idx, checked) { currentRules[idx].Enabled = checked ? '1' : '0'; renderRulesTable(); }
function saveRules() { var ddl = document.getElementById('rulesDdl'), fileName = ddl && ddl.value ? ddl.value : prompt('Nhập tên file:', 'new_rules.csv'); if (!fileName) return; if (!fileName.endsWith('.csv')) fileName += '.csv'; postToAHK({ action: 'saveRules', fileName: fileName, rules: currentRules }); }
function addRule() { rulesEditRow = -1; openRuleDialog(null); }
function editRule() { if (rulesSelectedRow < 0 || rulesSelectedRow >= currentRules.length) return toast('Chọn rule để sửa'); rulesEditRow = rulesSelectedRow; openRuleDialog(currentRules[rulesSelectedRow]); }
function copyRule() { if (rulesSelectedRow < 0 || rulesSelectedRow >= currentRules.length) return toast('Chọn rule để copy'); var orig = currentRules[rulesSelectedRow]; rulesEditRow = -1; openRuleDialog(Object.assign({}, orig, { RuleID: orig.RuleID + '_copy' })); }
async function deleteRule() { if (rulesSelectedRow < 0 || rulesSelectedRow >= currentRules.length) { return toast('Chọn rule để xóa'); } const result = await Swal.fire({ title: 'Xác nhận xóa?', text: `Bạn có chắc muốn xóa vĩnh viễn Rule "${currentRules[rulesSelectedRow].RuleID}"?`, icon: 'warning', showCancelButton: true, confirmButtonColor: '#d33', cancelButtonColor: '#3085d6', confirmButtonText: 'Xóa!', cancelButtonText: 'Hủy' }); if (result.isConfirmed) { currentRules.splice(rulesSelectedRow, 1); rulesSelectedRow = Math.min(rulesSelectedRow, currentRules.length - 1); renderRulesTable(); toast('Đã xóa rule.'); } }
function moveRuleUp() { if (rulesSelectedRow <= 0 || rulesSelectedRow >= currentRules.length) return; var tmp = currentRules[rulesSelectedRow]; currentRules[rulesSelectedRow] = currentRules[rulesSelectedRow - 1]; currentRules[rulesSelectedRow - 1] = tmp; rulesSelectedRow--; renderRulesTable(); }
function moveRuleDown() { if (rulesSelectedRow < 0 || rulesSelectedRow >= currentRules.length - 1) return; var tmp = currentRules[rulesSelectedRow]; currentRules[rulesSelectedRow] = currentRules[rulesSelectedRow + 1]; currentRules[rulesSelectedRow + 1] = tmp; rulesSelectedRow++; renderRulesTable(); }
function openRuleDialog(rule) { var isEdit = rule !== null, now = new Date(), ts = now.getFullYear().toString() + padZ(now.getMonth() + 1) + padZ(now.getDate()) + padZ(now.getHours()) + padZ(now.getMinutes()) + padZ(now.getSeconds()); setV('mPriority', isEdit ? rule.Priority : '50'); setV('mRuleID', isEdit ? rule.RuleID : 'R' + ts); setDDL('mType', isEdit ? rule.Type : 'exact'); setV('mColumn', isEdit ? rule.Column : 'A'); setV('mMatch', isEdit ? rule.Match : ''); setV('mOutput', isEdit ? rule.Output : ''); setV('mDescription', isEdit ? rule.Description : ''); document.getElementById('mEnabled').checked = !isEdit || rule.Enabled === '1'; var opts = isEdit ? rule.Options : ''; setDDL('mNormalize', getOption(opts, 'normalize') || 'none'); setDDL('mInputMode', getOption(opts, 'inputmode') || 'contains'); setDDL('mMatchType', getOption(opts, 'matchtype') || 'boundary'); setV('mOptExclude', getOption(opts, 'exclude')); setV('mOptChild', getOption(opts, 'child')); setV('mOptGroup', getOption(opts, 'group')); setV('mOptStop', getOption(opts, 'stop')); document.getElementById('mChkExclude').checked = getOption(opts, 'exclude') !== ''; document.getElementById('mChkChild').checked = getOption(opts, 'child') !== ''; document.getElementById('mChkLimit').checked = getOption(opts, 'limit') === 'level'; document.getElementById('ruleModalTitle').textContent = isEdit ? 'Sửa Rule' : 'Thêm Rule'; document.getElementById('ruleModal').style.display = 'flex'; updateModalState(); }
function closeRuleDialog() { document.getElementById('ruleModal').style.display = 'none'; }
function updateModalState() { var type = document.getElementById('mType').value; setEnabled('mColumn', ['exact', 'contains', 'range', 'combine', 'search', 'empty'].indexOf(type) >= 0); setEnabled('mMatch', ['exact', 'contains', 'range', 'combine'].indexOf(type) >= 0); setEnabled('mInputMode', ['contains', 'combine'].indexOf(type) >= 0); setEnabled('mOutput', ['exact', 'contains', 'range', 'combine', 'empty', 'always'].indexOf(type) >= 0); setEnabled('mMatchType', ['exact', 'contains', 'combine', 'search', 'empty'].indexOf(type) >= 0); setEnabled('mNormalize', type !== 'always' && type !== 'empty'); setEnabled('mChkChild', type === 'range'); setEnabled('mOptChild', type === 'range'); }
function saveRuleDialog() { var optParts = [], norm = document.getElementById('mNormalize').value, im = document.getElementById('mInputMode').value, mt = document.getElementById('mMatchType').value; if (norm !== 'none') optParts.push('normalize:' + norm); if (im !== 'contains') optParts.push('inputmode:' + im); if (mt !== 'boundary') optParts.push('matchtype:' + mt); if (document.getElementById('mChkExclude').checked && getV('mOptExclude')) optParts.push('exclude:' + getV('mOptExclude')); if (document.getElementById('mChkChild').checked && getV('mOptChild')) optParts.push('child:' + getV('mOptChild')); if (document.getElementById('mChkLimit').checked) optParts.push('limit:level'); if (getV('mOptGroup')) optParts.push('group:' + getV('mOptGroup')); if (getV('mOptStop')) optParts.push('stop:' + getV('mOptStop')); var rule = { Priority: getV('mPriority'), RuleID: getV('mRuleID'), Type: document.getElementById('mType').value, Column: getV('mColumn'), Match: getV('mMatch'), Output: getV('mOutput'), Options: optParts.join('|'), Enabled: document.getElementById('mEnabled').checked ? '1' : '0', Description: getV('mDescription'), Color: (rulesEditRow >= 0 && rulesEditRow < currentRules.length) ? currentRules[rulesEditRow].Color || '' : '' }; if (rulesEditRow >= 0 && rulesEditRow < currentRules.length) currentRules[rulesEditRow] = rule; else currentRules.push(rule); closeRuleDialog(); renderRulesTable(); }
function openColorPicker(idx, dotEl, e) { e.stopPropagation(); if (_colorPickerEl) { _colorPickerEl.remove(); _colorPickerEl = null; } var colors = [ '', '#fce8e6', '#fef3c0', '#e6f4ea', '#e8f0fe', '#fce4ec', '#f3e8fd', '#e0f2f1', '#fff3e0', '#f1f8e9', '#e3f2fd', '#ffccbc', '#d7ccc8', '#cfd8dc', '#b2dfdb', '#c8e6c9' ]; var picker = document.createElement('div'); picker.style.cssText = 'position:fixed;background:#fff;border:1px solid #dadce0;border-radius:8px;padding:8px;box-shadow:0 4px 12px rgba(0,0,0,.2);z-index:9999;display:flex;flex-wrap:wrap;gap:6px;width:168px;'; var rect = dotEl.getBoundingClientRect(); picker.style.left = rect.left + 'px'; picker.style.top = (rect.bottom + 4) + 'px'; var currentColor = currentRules[idx].Color || ''; colors.forEach(function (c) { var swatch = document.createElement('div'); var isNone = c === ''; swatch.style.cssText = 'width:22px;height:22px;border-radius:4px;cursor:pointer;border:2px solid ' + (c === currentColor ? '#1a73e8' : 'rgba(0,0,0,0.15)') + ';background:' + (isNone ? '#fff' : c) + ';position:relative;'; if (isNone) { swatch.title = 'Không màu'; swatch.innerHTML = '<svg viewBox="0 0 22 22" style="position:absolute;inset:0;width:100%;height:100%;"><line x1="2" y1="2" x2="20" y2="20" stroke="#d93025" stroke-width="2"/></svg>'; } swatch.onclick = function () { setRuleColor(idx, c); picker.remove(); _colorPickerEl = null; }; swatch.onmouseover = function () { this.style.transform = 'scale(1.2)'; }; swatch.onmouseout = function () { this.style.transform = 'scale(1)'; }; picker.appendChild(swatch); }); var customRow = document.createElement('div'); customRow.style.cssText = 'width:100%;display:flex;align-items:center;gap:6px;margin-top:4px;border-top:1px solid #eee;padding-top:6px;'; var colorInput = document.createElement('input'); colorInput.type = 'color'; colorInput.value = currentColor || '#ffffff'; colorInput.style.cssText = 'width:32px;height:24px;border:1px solid #dadce0;border-radius:4px;cursor:pointer;padding:0;'; colorInput.oninput = function () { setRuleColor(idx, this.value); }; var label = document.createElement('span'); label.textContent = 'Màu tùy chỉnh'; label.style.cssText = 'font-size:11px;color:#5f6368;'; customRow.appendChild(colorInput); customRow.appendChild(label); picker.appendChild(customRow); document.body.appendChild(picker); _colorPickerEl = picker; setTimeout(function () { document.addEventListener('click', function closePicker() { if (_colorPickerEl) { _colorPickerEl.remove(); _colorPickerEl = null; } document.removeEventListener('click', closePicker); }); }, 0); }
function setRuleColor(idx, color) { if (!currentRules[idx]) return; currentRules[idx].Color = color; renderRulesTable(); }
function refreshAndReload() { postToAHK({ action: 'refreshRules' }); setTimeout(function () { var ddl = document.getElementById('rulesDdl'); if (ddl && ddl.value) { postToAHK({ action: 'loadRulesFile', file: ddl.value }); } }, 300); }