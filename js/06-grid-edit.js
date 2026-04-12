// ============================================
// NỘI DUNG MỚI CỦA FILE: 06-grid-edit.js
// ============================================

function enterEditMode(cell){if(editingCell&&editingCell!==cell)exitEditMode();editingCell=cell;var info=getCellInfo(cell);var raw=_currentState.data[info.row-1]?_currentState.data[info.row-1][info.col-1]||'':'';cell.innerText=raw;cell.contentEditable='true';cell.classList.add('editing');cell.focus();moveCursorToEnd(cell)}
function enterEditModeByInfo(info){var cell=document.querySelector('.data-cell[data-row="'+info.row+'"][data-col="'+info.col+'"]');if(cell)enterEditMode(cell);}
function exitEditMode(){if(!editingCell)return;var info=getCellInfo(editingCell);var r=info.row-1;var c=info.col-1;var val=editingCell.innerText;if(_currentState.data[r]&&_currentState.data[r][c]!==val){saveState();_currentState.data[r][c]=val;
// ---> THAY ĐỔI: Cập nhật cache cho ô vừa sửa <---
updateCellDisplayCache(r, c);
// ---> KẾT THÚC THAY ĐỔI <---
delete _currentState.manualRowHeights[r];recalcRowHeight(r);rebuildPrefixSums();if(hf){hf.setCellContents({sheet:0,row:r,col:c},[[val]])}
markAsModified()}
editingCell.contentEditable='false';editingCell.classList.remove('editing');editingCell=null;renderVisibleRows()}
function editCell(){if(currentCell)enterEditModeByInfo(currentCell);}
function moveCursorToEnd(el){var range=document.createRange(),sel=window.getSelection();range.selectNodeContents(el);range.collapse(!1);sel.removeAllRanges();sel.addRange(range)}
function deleteContent(){var count=getSelectionCount();if(count===0)return;saveState();var affected={};forEachSelected(function(r,c){var ri = r - 1; var ci = c - 1; if(_currentState.data[ri]){_currentState.data[ri][ci]='';
// ---> THAY ĐỔI: Cập nhật cache cho ô vừa xóa <---
updateCellDisplayCache(ri, ci);
affected[ri]=!0}});for(var r in affected){delete _currentState.manualRowHeights[parseInt(r)];recalcRowHeight(parseInt(r))}
rebuildPrefixSums();renderVisibleRows();markAsModified();if(searchTerm)refreshSearch();}
function insertRowAbove(){var tr=currentCell?currentCell.row:1;if(selectedRows.length>0)tr=Math.min.apply(null,selectedRows);saveState();_currentState.data.splice(tr-1,0,Array(_currentState.headers.length).fill(''));_currentState.rowHeights.splice(tr-1,0,MIN_ROW_HEIGHT);
// ---> THAY ĐỔI: Chèn một hàng cache trống <---
_currentState.displayCache.splice(tr-1, 0, []);
var nm={};for(var k in _currentState.manualRowHeights){var ki=parseInt(k);nm[ki>=tr-1?ki+1:ki]=_currentState.manualRowHeights[k]}_currentState.manualRowHeights=nm;rebuildPrefixSums();fullRender();toast('Đã thêm dòng trên');markAsModified()}
function insertRowBelow(){var tr=currentCell?currentCell.row:_currentState.data.length;if(selectedRows.length>0)tr=Math.max.apply(null,selectedRows);saveState();_currentState.data.splice(tr,0,Array(_currentState.headers.length).fill(''));_currentState.rowHeights.splice(tr,0,MIN_ROW_HEIGHT);
// ---> THAY ĐỔI: Chèn một hàng cache trống <---
_currentState.displayCache.splice(tr, 0, []);
var nm={};for(var k in _currentState.manualRowHeights){var ki=parseInt(k);nm[ki>=tr?ki+1:ki]=_currentState.manualRowHeights[k]}_currentState.manualRowHeights=nm;rebuildPrefixSums();fullRender();toast('Đã thêm dòng dưới');markAsModified()}
function insertColLeft(){var tc=currentCell?currentCell.col:1;saveState();_currentState.headers.splice(tc-1,0,'Column '+(_currentState.headers.length+1));_currentState.colWidths.splice(tc-1,0,120);for(var i=0;i<_currentState.data.length;i++)_currentState.data[i].splice(tc-1,0,'');_currentState.manualRowHeights={};recalcAllRowHeights();
// ---> THAY ĐỔI: Rebuild toàn bộ cache vì cấu trúc cột thay đổi <---
rebuildDisplayCache();
fullRender();markAsModified()}
function insertColRight(){var tc=currentCell?currentCell.col:_currentState.headers.length;saveState();_currentState.headers.splice(tc,0,'Column '+(_currentState.headers.length+1));_currentState.colWidths.splice(tc,0,120);for(var i=0;i<_currentState.data.length;i++)_currentState.data[i].splice(tc,0,'');_currentState.manualRowHeights={};recalcAllRowHeights();
// ---> THAY ĐỔI: Rebuild toàn bộ cache vì cấu trúc cột thay đổi <---
rebuildDisplayCache();
fullRender();markAsModified()}
function deleteSelectedRows(){var rd=selectedRows.length>0?selectedRows.slice():(currentCell?[currentCell.row]:[]);if(rd.length===0){toast('Chọn dòng trước');return}
saveState();rd.sort(function(a,b){return b-a});if(rd.length>=_currentState.data.length){_currentState.data=[];_currentState.rowHeights=[];_currentState.manualRowHeights={};
// ---> THAY ĐỔI: Xóa cache <---
_currentState.displayCache = [];
_currentState.data.push(Array(_currentState.headers.length).fill(''));_currentState.rowHeights.push(MIN_ROW_HEIGHT);clearSel();selectedRows=[];selectedCols=[];currentCell={row:1,col:1};rebuildPrefixSums();fullRender();toast('Đã xóa sạch bảng. Tạo 1 dòng trống mới.');return}
for(var j=0;j<rd.length;j++){var rowIndex = rd[j]-1; _currentState.data.splice(rowIndex,1);_currentState.rowHeights.splice(rowIndex,1);
// ---> THAY ĐỔI: Xóa cache của dòng tương ứng <---
_currentState.displayCache.splice(rowIndex, 1);
}
_currentState.manualRowHeights={};clearSel();selectedRows=[];selectedCols=[];if(currentCell&&currentCell.row>_currentState.data.length){currentCell.row=_currentState.data.length}
rebuildPrefixSums();fullRender();toast('Đã xóa '+rd.length+' dòng');markAsModified()}
function deleteSelectedColumns(){if(!selRange&&!selCells){toast('Chọn cột trước');return}saveState();var cols={};forEachSelected(function(r,c){cols[c]=!0});var cn=Object.keys(cols).map(Number).sort(function(a,b){return b-a});if(cn.length>=_currentState.headers.length){toast('Giữ ít nhất 1 cột!');return}
for(var j=0;j<cn.length;j++){var c=cn[j]-1;_currentState.headers.splice(c,1);_currentState.colWidths.splice(c,1);for(var r=0;r<_currentState.data.length;r++)_currentState.data[r].splice(c,1);}
clearSel();selectedRows=[];selectedCols=[];_currentState.manualRowHeights={};recalcAllRowHeights();
// ---> THAY ĐỔI: Rebuild toàn bộ cache vì cấu trúc cột thay đổi <---
rebuildDisplayCache();
fullRender();toast('Đã xóa '+cn.length+' cột');markAsModified()}
function copySelection(){if(!selRange&&!selCells&&!currentCell)return;var minR=Infinity,maxR=-1,minC=Infinity,maxC=-1,cm={};forEachSelected(function(r,c){if(isRowHidden(r-1))return;minR=Math.min(minR,r);maxR=Math.max(maxR,r);minC=Math.min(minC,c);maxC=Math.max(maxC,c);if(!cm[r])cm[r]={};cm[r][c]=_currentState.data[r-1]?_currentState.data[r-1][c-1]||'':''});if(minR===Infinity)return;var pl=[],hl='<table>';for(var r=minR;r<=maxR;r++){var cols=[];hl+='<tr>';for(var c=minC;c<=maxC;c++){var v=(cm[r]&&cm[r][c]!==undefined)?cm[r][c]:'';cols.push(v);hl+='<td>'+escHtml(v)+'</td>'}
pl.push(cols.join('\t'));hl+='</tr>'}
hl+='</table>';copyToClipboard(pl.join('\r\n'),hl);var count=getSelectionCount();toast('📋 Đã copy '+count+' ô')}
function cutSelection(){copySelection();deleteContent()}
function pasteData(){if(!currentCell){toast('Chọn ô đích');return}saveState();pendingPasteCallback=function(text,html){processPasteData(text,html)};triggerPasteFromMenu()}
function processPasteData(text,html){if(!currentCell)return;var rows=[];if(html&&html.indexOf('<table')>=0)rows=parseHtmlTable(html);if(!rows.length&&text){text=text.replace(/\r\n/g,'\n').replace(/\r/g,'\n');if(text.endsWith('\n'))text=text.slice(0,-1);rows=text.split('\n').map(function(l){return l.split('\t')})}
if(!rows.length){toast('Không có dữ liệu');return}
var affected={},isSingle=rows.length===1&&rows[0].length===1,hasMulti=getSelectionCount()>1;if(isSingle&&hasMulti){var fv=rows[0][0];forEachSelected(function(r,c){var ri=r-1; var ci=c-1; if(_currentState.data[ri]){_currentState.data[ri][ci]=fv;
// ---> THAY ĐỔI: Cập nhật cache cho ô vừa paste <---
updateCellDisplayCache(ri, ci);
affected[ri]=!0}})}else{var sr=currentCell.row,sc=currentCell.col;var maxPasteCols=0;for(var i=0;i<rows.length;i++){if(rows[i].length>maxPasteCols)maxPasteCols=rows[i].length}
var needC=sc+maxPasteCols-1;while(_currentState.headers.length<needC){_currentState.headers.push("Column "+(_currentState.headers.length+1));_currentState.colWidths.push(120);for(var r=0;r<_currentState.data.length;r++){_currentState.data[r].push('')}}
var needR=sr+rows.length-1;while(_currentState.data.length<needR){_currentState.data.push(Array(_currentState.headers.length).fill(''));_currentState.rowHeights.push(MIN_ROW_HEIGHT);
// ---> THAY ĐỔI: Thêm cache cho dòng mới <---
_currentState.displayCache.push([]);}
for(var i=0;i<rows.length;i++){for(var j=0;j<rows[i].length;j++){var tr=sr+i-1,tc=sc+j-1;if(tr<_currentState.data.length&&tc<_currentState.headers.length){_currentState.data[tr][tc]=rows[i][j];
// ---> THAY ĐỔI: Cập nhật cache cho ô vừa paste <---
updateCellDisplayCache(tr, tc);
affected[tr]=!0}}}}
for(var r in affected){delete _currentState.manualRowHeights[parseInt(r)];recalcRowHeight(parseInt(r))}
syncHyperFormula();rebuildPrefixSums();fullRender();if(searchTerm)refreshSearch();markAsModified()}
function copyToClipboard(pt,ht){if(navigator.clipboard&&navigator.clipboard.write){try{navigator.clipboard.write([new ClipboardItem({'text/html':new Blob([ht],{type:'text/html'}),'text/plain':new Blob([pt],{type:'text/plain'})})]);return}catch(e){}}
var ta=document.createElement('textarea');ta.value=pt;ta.style.cssText='position:fixed;left:-9999px';document.body.appendChild(ta);ta.select();try{document.execCommand('copy')}catch(e){}document.body.removeChild(ta)}
function triggerPasteFromMenu(){var ta=document.createElement('textarea');ta.style.cssText='position:fixed;left:-9999px;opacity:0';document.body.appendChild(ta);ta.focus();ta.addEventListener('paste',function(e){var cd=e.clipboardData||window.clipboardData,t=cd.getData('text/plain')||'',h=cd.getData('text/html')||'';e.preventDefault();document.body.removeChild(ta);if(pendingPasteCallback){var cb=pendingPasteCallback;pendingPasteCallback=null;cb(t,h)}});try{document.execCommand('paste')}catch(e){document.body.removeChild(ta);toast('Dùng Ctrl+V');pendingPasteCallback=null}}
function parseHtmlTable(html){var rows=[],doc=new DOMParser().parseFromString(html,'text/html'),trs=doc.querySelectorAll('tr');for(var i=0;i<trs.length;i++){var cells=trs[i].querySelectorAll('td,th'),row=[];for(var j=0;j<cells.length;j++)row.push(cells[j].innerText||cells[j].textContent||'');if(row.length>0)rows.push(row);}
return rows}
function saveState(){
    _currentState.undoStack.push(JSON.stringify({
        data: _currentState.data, headers: _currentState.headers, colWidths: _currentState.colWidths, 
        rowHeights: _currentState.rowHeights, manualRowHeights: _currentState.manualRowHeights,
        hiddenCols: Array.from(_currentState.hiddenCols),
        // ---> THAY ĐỔI: Lưu cả displayCache vào Undo <---
        displayCache: _currentState.displayCache
    }));
    if(_currentState.undoStack.length>20)_currentState.undoStack.shift();
    _currentState.redoStack=[];
}
function undo(){
    if(!_currentState.undoStack.length){toast('Không có undo');return}
    _currentState.redoStack.push(JSON.stringify({
        data: _currentState.data, headers: _currentState.headers, colWidths: _currentState.colWidths, 
        rowHeights: _currentState.rowHeights, manualRowHeights: _currentState.manualRowHeights,
        hiddenCols: Array.from(_currentState.hiddenCols),
        // ---> THAY ĐỔI: Lưu cả displayCache vào Redo <---
        displayCache: _currentState.displayCache
    }));
    var s=JSON.parse(_currentState.undoStack.pop());
    _currentState.data=s.data;
    _currentState.headers=s.headers;
    _currentState.colWidths=s.colWidths||[];
    _currentState.rowHeights=s.rowHeights||[];
    _currentState.manualRowHeights=s.manualRowHeights||{};
    _currentState.hiddenCols = new Set(s.hiddenCols || []);
    // ---> THAY ĐỔI: Khôi phục displayCache <---
    _currentState.displayCache = s.displayCache || [];
    clearSel();selectedRows=[];selectedCols=[];fullRender();toast('Đã undo');
}
function redo(){
    if(!_currentState.redoStack.length){toast('Không có redo');return}
    _currentState.undoStack.push(JSON.stringify({
        data: _currentState.data, headers: _currentState.headers, colWidths: _currentState.colWidths, 
        rowHeights: _currentState.rowHeights, manualRowHeights: _currentState.manualRowHeights,
        hiddenCols: Array.from(_currentState.hiddenCols),
        // ---> THAY ĐỔI: Lưu cả displayCache vào Undo <---
        displayCache: _currentState.displayCache
    }));
    var s=JSON.parse(_currentState.redoStack.pop());
    _currentState.data=s.data;
    _currentState.headers=s.headers;
    _currentState.colWidths=s.colWidths||[];
    _currentState.rowHeights=s.rowHeights||[];
    _currentState.manualRowHeights=s.manualRowHeights||{};
    _currentState.hiddenCols = new Set(s.hiddenCols || []);
    // ---> THAY ĐỔI: Khôi phục displayCache <---
    _currentState.displayCache = s.displayCache || [];
    clearSel();selectedRows=[];selectedCols=[];fullRender();toast('Đã redo');
}
function hideSelectedCols(){var colsToHide=new Set();if(selectedCols.length>0){for(var i=0;i<selectedCols.length;i++)colsToHide.add(selectedCols[i]-1);}else if(selRange){for(var c=selRange.c1;c<=selRange.c2;c++)colsToHide.add(c-1);}else if(selCells&&selCells.size>0){selCells.forEach(function(k){colsToHide.add(parseInt(k.split(',')[1])-1)})}else if(currentCell){colsToHide.add(currentCell.col-1)}
if(colsToHide.size===0){toast('Chọn ít nhất 1 cột để ẩn');return}
if(_currentState.hiddenCols.size+colsToHide.size>=_currentState.headers.length){toast('Không thể ẩn tất cả các cột!');return}
saveState();colsToHide.forEach(function(cIdx){_currentState.hiddenCols.add(cIdx);_currentState.colWidths[cIdx]=0});clearSel();rebuildPrefixSums();fullRender();toast('👁️‍🗨️ Đã ẩn '+colsToHide.size+' cột')}
function unhideAdjacentCols(){var colsToUnhide=new Set();var checkCols=[];if(selectedCols.length>0)checkCols=selectedCols;else if(selRange){for(var c=selRange.c1;c<=selRange.c2;c++)checkCols.push(c);}else if(currentCell)checkCols.push(currentCell.col);for(var i=0;i<checkCols.length;i++){var cIdx=checkCols[i]-1;var leftIdx=cIdx-1;while(leftIdx>=0&&_currentState.hiddenCols.has(leftIdx)){colsToUnhide.add(leftIdx);leftIdx--}
var rightIdx=cIdx+1;while(rightIdx<_currentState.headers.length&&_currentState.hiddenCols.has(rightIdx)){colsToUnhide.add(rightIdx);rightIdx++}}
if(colsToUnhide.size===0){toast('Không có cột nào bị ẩn kề bên');return}
saveState();colsToUnhide.forEach(function(cIdx){_currentState.hiddenCols.delete(cIdx);if(_currentState.colWidths[cIdx]===0)_currentState.colWidths[cIdx]=120});rebuildPrefixSums();fullRender();toast('👁️ Đã hiện lại '+colsToUnhide.size+' cột')}
function unhideAllCols(){if(_currentState.hiddenCols.size===0){toast('Không có cột nào bị ẩn');return}
saveState();var count=_currentState.hiddenCols.size;_currentState.hiddenCols.forEach(function(cIdx){if(_currentState.colWidths[cIdx]===0)_currentState.colWidths[cIdx]=120});_currentState.hiddenCols.clear();rebuildPrefixSums();fullRender();toast('👁️ Đã hiện lại tất cả '+count+' cột')}