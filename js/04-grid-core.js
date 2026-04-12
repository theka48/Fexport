// ============================================
// NỘI DUNG MỚI CỦA FILE: 04-grid-core.js
// ============================================

function initHyperFormula(){
    if(!window.HyperFormula || !_currentState.data)return;
    try{
        hf=HyperFormula.buildFromArray(_currentState.data,{licenseKey:'gpl-v3'})
    }catch(e){
        if(e.message&&e.message.includes('Language not registered')){
            try{var enUS=HyperFormula.languages.getLanguage('en-US');HyperFormula.languages.registerLanguage('enUS',enUS);hf=HyperFormula.buildFromArray(_currentState.data,{licenseKey:'gpl-v3',language:'enUS'})
            }catch(e2){toast('❌ Lỗi HyperFormula: '+e2.message)}
        }else{toast('❌ Lỗi HyperFormula: '+e.message)}
    }
}
function syncHyperFormula(){if(!hf)return;hf.setSheetContent(0,_currentState.data)}
function evalFormula(row,col){if(!hf)return null;try{var val=hf.getCellValue({sheet:0,row:row,col:col});if(val&&val.type==='ERROR')return'#ERROR!';return val}catch(e){return'#ERROR!'}}
function isFormula(val){return val&&String(val).charAt(0)==='='}
function getDisplayValue(rowIdx,colIdx){var raw=_currentState.data[rowIdx]&&_currentState.data[rowIdx][colIdx]!==undefined?_currentState.data[rowIdx][colIdx]:'';if(isFormula(raw)){if(!hf)initHyperFormula();return evalFormula(rowIdx,colIdx)}
return raw}
function rebuildPrefixSums(){
    rowTopCache=new Float64Array(_currentState.data.length+1);
    for(var i=0;i<_currentState.data.length;i++)rowTopCache[i+1]=rowTopCache[i]+(_currentState.rowHeights[i]||MIN_ROW_HEIGHT);
    colLeftCache=new Float64Array(_currentState.headers.length+1);
    for(var i=0;i<_currentState.headers.length;i++){
        var w=_currentState.hiddenCols.has(i)?0:(_currentState.colWidths[i]||120);
        colLeftCache[i+1]=colLeftCache[i]+w
    }
}
function getRowTop(r){return rowTopCache?rowTopCache[r]:0}
function getColLeft(c){return colLeftCache?colLeftCache[c]:0}
function getTotalTableWidth(){return colLeftCache?colLeftCache[_currentState.headers.length]:0}
function calculateTotalHeight(){return rowTopCache?rowTopCache[_currentState.data.length]:0}
function findRowAtY(y){if(!rowTopCache||_currentState.data.length===0)return 0;var lo=0,hi=_currentState.data.length-1;while(lo<hi){var mid=(lo+hi+1)>>1;if(rowTopCache[mid]<=y)lo=mid;else hi=mid-1}return lo}
function createMeasureDivs(){measureDiv=document.createElement('div');measureDiv.className='measure-cell';document.body.appendChild(measureDiv);measureWidthDiv=document.createElement('div');measureWidthDiv.className='measure-width';document.body.appendChild(measureWidthDiv)}
function measureTextHeight(text,width){if(!measureDiv)return MIN_ROW_HEIGHT;measureDiv.style.width=width+'px';if(!text||text==='')return MIN_ROW_HEIGHT;measureDiv.innerText=text;return Math.max(MIN_ROW_HEIGHT,measureDiv.offsetHeight)}
function measureTextWidth(text){if(!measureWidthDiv||!text||text==='')return 50;var lines=String(text).split('\n'),maxW=30;for(var i=0;i<lines.length;i++){measureWidthDiv.innerText=lines[i];if(measureWidthDiv.offsetWidth>maxW)maxW=measureWidthDiv.offsetWidth}
return maxW+2}
function calcRowHeight(r){if(r<0||r>=_currentState.data.length)return MIN_ROW_HEIGHT;var maxH=MIN_ROW_HEIGHT;for(var c=0;c<_currentState.data[r].length;c++){var v=_currentState.data[r][c];if(!v||v==='')continue;var h=measureTextHeight(v,_currentState.colWidths[c]||120);if(h>maxH)maxH=h}return maxH}
function calcAllRowHeights(){_currentState.rowHeights=new Array(_currentState.data.length);for(var r=0;r<_currentState.data.length;r++){var ch=calcRowHeight(r);_currentState.rowHeights[r]=_currentState.manualRowHeights[r]!==undefined?Math.max(_currentState.manualRowHeights[r],ch):ch}}
function recalcRowHeight(r){if(r<0||r>=_currentState.data.length)return;var ch=calcRowHeight(r);_currentState.rowHeights[r]=_currentState.manualRowHeights[r]!==undefined?Math.max(_currentState.manualRowHeights[r],ch):ch}
function recalcAllRowHeights(){for(var r=0;r<_currentState.data.length;r++)recalcRowHeight(r);}
function autoFitAll(){toast('Đang auto fit...');for(var ci=0;ci<_currentState.headers.length;ci++){if(_currentState.hiddenCols.has(ci)||_currentState.fixedWidthCols.has(ci))continue;var maxW=MIN_COL_WIDTH,nw=measureTextWidth((ci+1)+' - '+(_currentState.headers[ci]||''))+20;if(nw>maxW)maxW=nw;for(var r=0;r<_currentState.data.length;r++){var v=_currentState.data[r][ci];if(!v)continue;var w=measureTextWidth(v)+16;if(w>maxW)maxW=w}
_currentState.colWidths[ci]=Math.min(maxW,800)}
_currentState.manualRowHeights={};recalcAllRowHeights();rebuildPrefixSums();fullRender();toast('Đã auto fit')}
function renderHeader(){var thead=document.getElementById('tableHead'),html='<tr><th class="corner-cell" onclick="selectAll()">✓</th>';for(var c=0;c<_currentState.headers.length;c++){if(_currentState.hiddenCols.has(c))continue;var w=_currentState.colWidths[c]||120,isSel=selectedCols.indexOf(c+1)>=0,isFixed=c<_currentState.fixedCols,isBorder=c===_currentState.fixedCols-1;var cls='col-header'+(isSel?' sel':'')+(isFixed?' fixed-col':'')+(isBorder?' fixed-col-border':'');var leftStyle=isFixed?'left:'+(50+getColLeft(c))+'px;':'';html+='<th class="'+cls+'" data-col="'+(c+1)+'" style="width:'+w+'px;min-width:'+w+'px;'+leftStyle+'">';if(activeTabNum===3){html+='<span>'+colToLetter(c+1)+'</span>'}else{html+='<span>'+colToLetter(c+1)+escHtml(_currentState.headers[c])+'</span>'}
html+='<div class="col-resize" data-col="'+(c+1)+'"></div></th>'}
thead.innerHTML=html+'</tr>'}
function renderVisibleRows(){var gw=document.getElementById('gridWrapper'),scale=zoomLevel/100;var scrollTop=(gw.scrollTop/scale)-28,viewH=gw.clientHeight/scale;var buf=50;var visibleRowIndices=[];for(var ri=0;ri<_currentState.data.length;ri++){if(!isRowHidden(ri)&&!isRowHiddenBySearch(ri))visibleRowIndices.push(ri);}
var visRowTops=new Float64Array(visibleRowIndices.length+1);for(var vi=0;vi<visibleRowIndices.length;vi++){visRowTops[vi+1]=visRowTops[vi]+(_currentState.rowHeights[visibleRowIndices[vi]]||MIN_ROW_HEIGHT)}
var totalVisH=visRowTops[visibleRowIndices.length];var sizer=document.getElementById('scrollSizer');if(!sizer){sizer=document.createElement('div');sizer.id='scrollSizer';sizer.style.cssText='position:absolute; top:0; left:0; width:1px; z-index:-1; pointer-events:none; visibility:hidden;';gw.appendChild(sizer)}
sizer.style.height=totalVisH+'px';var startVI=0,targetTop=Math.max(0,scrollTop-buf),lo=0,hi=visibleRowIndices.length-1;while(lo<hi){var mid=(lo+hi+1)>>1;if(visRowTops[mid]<=targetTop)lo=mid;else hi=mid-1}startVI=lo;var endVI=startVI,cumH=0;while(endVI<visibleRowIndices.length&&cumH<viewH+buf*2){cumH+=_currentState.rowHeights[visibleRowIndices[endVI]]||MIN_ROW_HEIGHT;endVI++}
endVI=Math.min(visibleRowIndices.length-1,endVI);var tbody=document.getElementById('tableBody');var html='<tr style="height:0;line-height:0;font-size:0;visibility:hidden;"><td class="corner-cell" style="padding:0;border:none;height:0;"></td>';for(var c=0;c<_currentState.headers.length;c++){if(!_currentState.hiddenCols.has(c))html+='<td style="width:'+(_currentState.colWidths[c]||120)+'px;padding:0;border:none;height:0;"></td>'}
html+='</tr>';var selCount=getSelectionCount();var isMultiRange=selCount>1;function buildRowHTML(vi){var ri=visibleRowIndices[vi],rn=ri+1,rh=_currentState.rowHeights[ri]||MIN_ROW_HEIGHT;var isFixedRow=ri<_currentState.fixedRows,isBorderRow=ri===_currentState.fixedRows-1;var fixedTopPx=isFixedRow?(28+getRowTop(ri)):0;var rHtml='<tr data-row="'+rn+'" style="height:'+rh+'px">';var isFullRowRange=(selRange&&rn>=selRange.r1&&rn<=selRange.r2&&selRange.c1===1&&selRange.c2===_currentState.headers.length);var isRowHeaderSel=(selectedRows.indexOf(rn)>=0)||isFullRowRange;var hasLayer=hasActiveLayers(rn);var rhCls='row-header'+(isRowHeaderSel?' sel':'')+(isFixedRow?' fixed-row':'')+(isBorderRow?' fixed-row-border':'');var rowHeaderZ=isFixedRow?190:100;rHtml+='<td class="'+rhCls+'" data-row="'+rn+'" style="height:'+rh+'px; top:'+(isFixedRow?fixedTopPx+'px':'auto')+'; z-index:'+rowHeaderZ+';">'+rn+'<div class="row-resize-handle" data-row="'+rn+'"></div></td>';for(var c=0;c<_currentState.headers.length;c++){if(_currentState.hiddenCols.has(c))continue;var isFixedCol=c<_currentState.fixedCols,isBorderCol=c===_currentState.fixedCols-1;var cellLeft=isFixedCol?(50+getColLeft(c)):0;var isSel=isCellSelected(rn,c+1);var isCur=currentCell&&currentCell.row===rn&&currentCell.col===c+1;var inRange=isSel;var isCtrlSel=selCells&&selectionType==='ctrl'&&selCells.has(rn+','+(c+1));var cls='data-cell'+(inRange&&isMultiRange?' sel-range':'')+(isCtrlSel?' sel-ctrl':'')+(isCur?' current':'')+(isSearchMatch(rn,c+1)?' search-match':'')+(isSearchCurrent(rn,c+1)?' search-current':'')+(isFixedCol?' fixed-col':'')+(isBorderCol?' fixed-col-border':'')+(isFixedRow?' fixed-row':'')+(isBorderRow?' fixed-row-border':'');var style='width:'+(_currentState.colWidths[c]||120)+'px;height:'+rh+'px;';if(isFixedCol&&isFixedRow){style+='left:'+cellLeft+'px; top:'+fixedTopPx+'px; z-index: 180;'}else if(isFixedCol){style+='left:'+cellLeft+'px; z-index: 60;'}else if(isFixedRow){style+='top:'+fixedTopPx+'px; z-index: 170;'}
var cellValue=getDisplayValue(ri,c);var displayHTML=(_currentState.displayCache[ri]&&_currentState.displayCache[ri][c]!==undefined)?_currentState.displayCache[ri][c]:escHtml(cellValue);rHtml+='<td class="'+cls+'" data-row="'+rn+'" data-col="'+(c+1)+'" contenteditable="false" style="'+style+'" tabindex="-1">'+displayHTML+'</td>'}
rHtml+='</tr>';return rHtml}
var numFixedToDraw=0;for(var vi=0;vi<visibleRowIndices.length;vi++){if(visibleRowIndices[vi]<_currentState.fixedRows){html+=buildRowHTML(vi);numFixedToDraw++}else{break}}
if(startVI<numFixedToDraw){startVI=numFixedToDraw}
var paddingHeight=visRowTops[startVI]-visRowTops[numFixedToDraw];if(paddingHeight>0){html+='<tr style="height:'+paddingHeight+'px; pointer-events:none;"><td colspan="'+(_currentState.headers.length+1)+'" style="padding:0;border:none;"></td></tr>'}
for(var vi=startVI;vi<=endVI;vi++){html+=buildRowHTML(vi)}
tbody.innerHTML=html;updateInfo();requestAnimationFrame(updateSelectionOverlay)}
function fullRender(){if(!_currentState) return; renderHeader();renderVisibleRows()}
function updateInfo(){var vis=0;for(var r=0;r<_currentState.data.length;r++){if(!isRowHidden(r)&&!isRowHiddenBySearch(r))vis++}
safeSetText('info','Dòng: '+vis+(vis<_currentState.data.length?' / '+_currentState.data.length:''))}
function setFixedRows(val){_currentState.fixedRows=Math.max(0,Math.min(parseInt(val)||0,_currentState.data.length));document.getElementById('fixedRowsInput').value=_currentState.fixedRows;fullRender()}
function setFixedCols(val){_currentState.fixedCols=Math.max(0,Math.min(parseInt(val)||0,_currentState.headers.length));document.getElementById('fixedColsInput').value=_currentState.fixedCols;fullRender()}
function zoomIn(){setZoom(Math.min(200,zoomLevel+10))}
function zoomOut(){setZoom(Math.max(50,zoomLevel-10))}
function zoomReset(){setZoom(100)}
function setZoom(level){zoomLevel=level;document.getElementById('mainTable').style.zoom=(level/100);document.getElementById('zoomLevel').textContent=level+'%';document.documentElement.style.setProperty('--sel-border-width',(1/(level/100))+'px');rebuildPrefixSums();requestAnimationFrame(renderVisibleRows)}
function formatCellWithRules(value,rn,cn){var escaped=escHtml(value);if(!_currentState.rowMatchedRules||!_currentState.rowMatchedRules[rn]||!value)return escaped;var rules=_currentState.rowMatchedRules[rn];var rawVal=String(value);var rangesToHighlight=[];var findMatches=function(pattern,isRegex,isExact){if(!pattern)return;if(isRegex){try{var rx=new RegExp(pattern,'gi');var match;while((match=rx.exec(rawVal))!==null){if(match[0].length>0)rangesToHighlight.push({start:match.index,end:match.index+match[0].length});if(!rx.global)break}}catch(e){}}else if(isExact){if(rawVal.toLowerCase()===pattern.toLowerCase())rangesToHighlight.push({start:0,end:rawVal.length})}else{var lowerRaw=rawVal.toLowerCase(),lowerPat=pattern.toLowerCase(),idx=0;while((idx=lowerRaw.indexOf(lowerPat,idx))>=0){rangesToHighlight.push({start:idx,end:idx+pattern.length});idx+=pattern.length}}};for(var i=0;i<rules.length;i++){var rule=rules[i];var cols=(rule.Column||"").split('+');var inputMode=getOption(rule.Options,'inputmode')||'contains';for(var ci=0;ci<cols.length;ci++){var cRef=cols[ci];var isMatchCol=(cRef===String(cn)||cRef===colToLetter(cn));if(!isMatchCol)continue;if(rule.Type==='exact'||rule.Type==='contains'){var mvs=(rule.Match||"").split('|');for(var j=0;j<mvs.length;j++)findMatches(mvs[j].trim(),inputMode==='regex',inputMode==='exact'||rule.Type==='exact');}else if(rule.Type==='combine'){var matchParts=(rule.Match||"").split('+');if(matchParts[ci]&&matchParts[ci]!=='*'){var mvs=matchParts[ci].split('|');for(var j=0;j<mvs.length;j++)findMatches(mvs[j].trim(),inputMode==='regex',inputMode==='exact');}}else if(rule.Type==='search'){rangesToHighlight.push({start:0,end:rawVal.length})}}}
if(rangesToHighlight.length===0)return escaped;rangesToHighlight.sort(function(a,b){return a.start-b.start});var merged=[rangesToHighlight[0]];for(var i=1;i<rangesToHighlight.length;i++){var curr=rangesToHighlight[i],last=merged[merged.length-1];if(curr.start<=last.end)last.end=Math.max(last.end,curr.end);else merged.push(curr)}
var result="",lastIdx=0;for(var i=0;i<merged.length;i++){var m=merged[i];result+=escHtml(rawVal.substring(lastIdx,m.start));result+='<span class="rule-match-text">'+escHtml(rawVal.substring(m.start,m.end))+'</span>';lastIdx=m.end}
result+=escHtml(rawVal.substring(lastIdx));return result}
function updateHeaderHighlight(){var activeRows=new Set();var activeCols=new Set();if(selectedRows.length>0){for(var i=0;i<selectedRows.length;i++)
activeRows.add(selectedRows[i]);for(var c=1;c<=_currentState.headers.length;c++)
activeCols.add(c);}else if(selectedCols.length>0){for(var i=0;i<selectedCols.length;i++)
activeCols.add(selectedCols[i]);for(var r=1;r<=_currentState.data.length;r++)
activeRows.add(r);}else if(selRange){for(var r=selRange.r1;r<=selRange.r2;r++)activeRows.add(r);for(var c=selRange.c1;c<=selRange.c2;c++)activeCols.add(c);}else if(selCells&&selCells.size>0){selCells.forEach(function(k){var p=k.split(',');activeRows.add(parseInt(p[0]));activeCols.add(parseInt(p[1]))})}else if(currentCell){activeRows.add(currentCell.row);activeCols.add(currentCell.col)}
document.querySelectorAll('.col-header').forEach(function(th){var c=parseInt(th.getAttribute('data-col'));if(selectedCols.indexOf(c)>=0){th.classList.add('sel')}else if(activeCols.has(c)){th.classList.add('sel')}else{th.classList.remove('sel')}});document.querySelectorAll('.row-header').forEach(function(td){var r=parseInt(td.getAttribute('data-row'));if(selectedRows.indexOf(r)>=0){td.classList.add('sel')}else if(activeRows.has(r)){td.classList.add('sel')}else{td.classList.remove('sel')}})}
function rebuildDisplayCache(){_currentState.displayCache=[];for(var r=0;r<_currentState.data.length;r++){var rowCache=[];for(var c=0;c<_currentState.headers.length;c++){rowCache.push(formatCellWithRules(getDisplayValue(r,c),r+1,c+1))}
_currentState.displayCache.push(rowCache)}}
function updateCellDisplayCache(r,c){if(!_currentState.displayCache[r])_currentState.displayCache[r]=[];_currentState.displayCache[r][c]=formatCellWithRules(getDisplayValue(r,c),r+1,c+1)}