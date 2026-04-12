// ============================================
// NỘI DUNG MỚI CỦA FILE: 07-search.js
// ============================================

function toggleSearchPanel(){var p=document.getElementById('searchPanel');if(p.classList.contains('show')){p.classList.remove('show');clearSearch()}else{captureSearchScope();p.classList.add('show');updateScopeLabel();var inp=document.getElementById('searchInput');inp.focus();inp.select();if(inp.value)doSearch(inp.value);}makeDraggable()}
function makeDraggable(){var p=document.getElementById('searchPanel'),h=p.querySelector('.search-panel-header'),dr=!1,sx,sy,ox,oy;h.onmousedown=function(e){if(e.target.classList.contains('search-close'))return;dr=!0;sx=e.clientX;sy=e.clientY;var r=p.getBoundingClientRect();ox=r.left;oy=r.top;e.preventDefault()};document.addEventListener('mousemove',function(e){if(!dr)return;p.style.left=(ox+e.clientX-sx)+'px';p.style.top=(oy+e.clientY-sy)+'px';p.style.right='auto'});document.addEventListener('mouseup',function(){dr=!1})}
function getSearchOptions(){return{matchCase:document.getElementById('optMatchCase').checked,wholeCell:document.getElementById('optWholeCell').checked,regex:document.getElementById('optRegex').checked,exclude:(document.getElementById('excludeInput')||{}).value||''}}
function normalizeForSearch(s,o){if(!s)return'';s=String(s);if(!o.matchCase)s=s.toLowerCase();return s}
function cellMatchesSearch(v,t,o){var nv=normalizeForSearch(v,o),nt=normalizeForSearch(t,o);if(!nt)return!1;if(o.regex){try{var re=new RegExp(nt,o.matchCase?'g':'gi');if(o.wholeCell)return re.test(nv)&&nv.replace(re,'').length===0;return re.test(nv)}catch(e){return!1}}
if(o.wholeCell)return nv===nt;return nv.indexOf(nt)>=0}
function refreshSearch(){var term=document.getElementById('searchInput').value,excl=(document.getElementById('excludeInput')||{}).value||'';if(term||excl)doSearch(term);}
function doSearch(term){searchTerm=term;searchResults=[];searchCurrentIdx=-1;searchMatchSet=null;searchMatchedRows=new Set();var opts=getSearchOptions(),hasSearch=term&&term!=='',hasExclude=opts.exclude&&opts.exclude!=='';if(!hasSearch&&!hasExclude){updateSearchInfo();if(searchFilterRows){document.getElementById('gridWrapper').scrollTop=0;fullRender()}else{renderVisibleRows()}return}
var excludedRows=new Set();if(hasExclude){for(var r=0;r<_currentState.data.length;r++){if(searchScope&&searchScope.rows&&searchScope.rows.indexOf(r+1)<0)continue;for(var c=0;c<_currentState.data[r].length;c++){if(searchScope&&searchScope.cols&&searchScope.cols.indexOf(c+1)<0)continue;if(_currentState.data[r][c]&&cellMatchesSearch(_currentState.data[r][c],opts.exclude,opts)){excludedRows.add(r+1);break}}}}
for(var r=0;r<_currentState.data.length;r++){if(searchScope&&searchScope.rows&&searchScope.rows.indexOf(r+1)<0)continue;if(excludedRows.has(r+1))continue;for(var c=0;c<_currentState.data[r].length;c++){if(searchScope&&searchScope.cols&&searchScope.cols.indexOf(c+1)<0)continue;if(!hasSearch||cellMatchesSearch(_currentState.data[r][c],term,opts)){searchResults.push({row:r+1,col:c+1});searchMatchedRows.add(r+1)}}}
searchMatchSet=new Set();for(var i=0;i<searchResults.length;i++)searchMatchSet.add(searchResults[i].row+','+searchResults[i].col);if(searchResults.length>0){searchCurrentIdx=0;if(currentCell)for(var j=0;j<searchResults.length;j++)if(searchResults[j].row>currentCell.row||(searchResults[j].row===currentCell.row&&searchResults[j].col>=currentCell.col)){searchCurrentIdx=j;break}}
excludeResults=[];excludeCurrentIdx=-1;if(hasExclude){for(var r=0;r<_currentState.data.length;r++){if(searchScope&&searchScope.rows&&searchScope.rows.indexOf(r+1)<0)continue;for(var c=0;c<_currentState.data[r].length;c++){if(searchScope&&searchScope.cols&&searchScope.cols.indexOf(c+1)<0)continue;if(_currentState.data[r][c]&&cellMatchesSearch(_currentState.data[r][c],opts.exclude,opts))excludeResults.push({row:r+1,col:c+1})}}}
updateSearchInfo();updateExcludeInfo();if(searchFilterRows){document.getElementById('gridWrapper').scrollTop=0;fullRender()}else{renderVisibleRows()}
if(searchResults.length>0)setTimeout(navigateToSearchResult,50);}
function triggerSearch(){doSearch(document.getElementById('searchInput').value)}
function findNext(){var t=document.getElementById('searchInput').value;if(t!==searchTerm)return triggerSearch();if(searchResults.length){searchCurrentIdx=(searchCurrentIdx+1)%searchResults.length;navigateToSearchResult();updateSearchInfo()}}
function findPrev(){var t=document.getElementById('searchInput').value;if(t!==searchTerm)return triggerSearch();if(searchResults.length){searchCurrentIdx=(searchCurrentIdx-1+searchResults.length)%searchResults.length;navigateToSearchResult();updateSearchInfo()}}
function findNextExclude(){var t=document.getElementById('searchInput').value;if(t!==searchTerm)return triggerSearch();if(excludeResults.length){excludeCurrentIdx=(excludeCurrentIdx+1)%excludeResults.length;navigateToExcludeResult();updateExcludeInfo()}}
function findPrevExclude(){var t=document.getElementById('searchInput').value;if(t!==searchTerm)return triggerSearch();if(excludeResults.length){excludeCurrentIdx=(excludeCurrentIdx-1+excludeResults.length)%excludeResults.length;navigateToExcludeResult();updateExcludeInfo()}}
function handleSearchKey(e){if(e.key==='Enter'){e.preventDefault();e.shiftKey?findPrev():findNext()}else if(e.key==='Escape'){toggleSearchPanel();e.preventDefault()}}
function handleExcludeKey(e){if(e.key==='Enter'){e.preventDefault();e.shiftKey?findPrevExclude():findNextExclude()}else if(e.key==='Escape'){toggleSearchPanel();e.preventDefault()}}
function handleReplaceKey(e){if(e.key==='Enter'){e.preventDefault();e.shiftKey?replaceAll():replaceCurrent()}else if(e.key==='Escape'){toggleSearchPanel();e.preventDefault()}}
function navigateToSearchResult(){if(searchCurrentIdx<0||searchCurrentIdx>=searchResults.length)return;var r=searchResults[searchCurrentIdx];currentCell={row:r.row,col:r.col};renderVisibleRows();setTimeout(function(){scrollToCell(r.row,r.col);renderVisibleRows()},20)}
function navigateToExcludeResult(){if(excludeCurrentIdx<0||excludeCurrentIdx>=excludeResults.length)return;var r=excludeResults[excludeCurrentIdx];currentCell={row:r.row,col:r.col};renderVisibleRows();setTimeout(function(){scrollToCell(r.row,r.col);renderVisibleRows()},20)}
function clearSearch(){searchTerm='';searchResults=[];searchCurrentIdx=-1;searchMatchSet=null;searchMatchedRows=new Set();searchScope=null;excludeResults=[];excludeCurrentIdx=-1;if(document.getElementById('searchInput'))document.getElementById('searchInput').value='';if(document.getElementById('excludeInput'))document.getElementById('excludeInput').value='';if(document.getElementById('optFilterRows'))document.getElementById('optFilterRows').checked=!1;searchFilterRows=!1;updateSearchInfo();updateExcludeInfo();updateScopeLabel();fullRender()}
function updateSearchInfo(){var el=document.getElementById('searchInfo');if(el){el.textContent=searchResults.length?(searchCurrentIdx+1)+' / '+searchResults.length:(searchTerm?'Không tìm thấy':'');el.style.color=searchResults.length?'#666':'#c00'}}
function updateExcludeInfo(){var el=document.getElementById('excludeInfo'),excl=(document.getElementById('excludeInput')||{}).value||'';if(el){el.textContent=(!excl||!excludeResults.length)?(excl?'Không có':''):((excludeCurrentIdx>=0?(excludeCurrentIdx+1):'—')+' / '+excludeResults.length);el.style.color='#c00'}}
function isSearchMatch(r,c){return searchMatchSet?searchMatchSet.has(r+','+c):!1}
function isSearchCurrent(r,c){if(searchCurrentIdx<0||searchCurrentIdx>=searchResults.length)return!1;var s=searchResults[searchCurrentIdx];return s.row===r&&s.col===c}
function replaceCurrent(){var term=document.getElementById('searchInput').value;if(!term||term.trim()===''){toast('Nhập từ khóa!');return}
if(!searchResults.length||searchCurrentIdx<0){toast('Không có kết quả');return}
var rv=(document.getElementById('replaceInput')||{}).value||'',opts=getSearchOptions();saveState();var res=searchResults[searchCurrentIdx],r=res.row-1,c=res.col-1,cv=_currentState.data[r][c]||'';
var originalValue = cv;
if(opts.wholeCell)_currentState.data[r][c]=rv;else if(opts.regex){try{_currentState.data[r][c]=cv.replace(new RegExp(searchTerm,opts.matchCase?'g':'gi'),rv)}catch(e){toast('Regex lỗi');return}}else{var idx=normalizeForSearch(cv,opts).indexOf(normalizeForSearch(searchTerm,opts));if(idx>=0)_currentState.data[r][c]=cv.substring(0,idx)+rv+cv.substring(idx+searchTerm.length);}
// ---> THAY ĐỔI: Chỉ cập nhật cache nếu giá trị thực sự thay đổi <---
if (originalValue !== _currentState.data[r][c]) {
    updateCellDisplayCache(r, c);
}
delete _currentState.manualRowHeights[r];recalcRowHeight(r);rebuildPrefixSums();var si=searchCurrentIdx;doSearch(searchTerm);if(searchResults.length>0){searchCurrentIdx=Math.min(si,searchResults.length-1);navigateToSearchResult();updateSearchInfo();renderVisibleRows()}toast('Đã thay thế 1 ô')}
function replaceAll(){var term=document.getElementById('searchInput').value;if(!term||term.trim()===''){toast('Nhập từ khóa!');return}
if(!searchResults.length){toast('Không có kết quả');return}
var rv=(document.getElementById('replaceInput')||{}).value||'',opts=getSearchOptions(),cnt=0;saveState();var affected={};
var changedCells = [];
for(var i=0;i<searchResults.length;i++){var res=searchResults[i],r=res.row-1,c=res.col-1,cv=_currentState.data[r][c]||'';
var originalValue = cv;
var newValue = cv;
if(opts.wholeCell)newValue=rv;else if(opts.regex){try{newValue=cv.replace(new RegExp(searchTerm,opts.matchCase?'g':'gi'),rv)}catch(e){continue}}else{var result='',sf=0,orig=cv;while(!0){var idx2=normalizeForSearch(orig.substring(sf),opts).indexOf(normalizeForSearch(searchTerm,opts));if(idx2<0){result+=orig.substring(sf);break}result+=orig.substring(sf,sf+idx2)+rv;sf+=idx2+searchTerm.length}newValue=result}
if (originalValue !== newValue) {
    _currentState.data[r][c] = newValue;
    affected[r]=!0;
    changedCells.push({r: r, c: c});
    cnt++;
}}
// ---> THAY ĐỔI: Cập nhật cache cho tất cả các ô đã thay đổi <---
for (var i=0; i<changedCells.length; i++) {
    updateCellDisplayCache(changedCells[i].r, changedCells[i].c);
}
for(var rk in affected){delete _currentState.manualRowHeights[parseInt(rk)];recalcRowHeight(parseInt(rk))}rebuildPrefixSums();doSearch(searchTerm);toast('Đã thay thế '+cnt+' ô')}
function toggleSearchFilterRows(){searchFilterRows=document.getElementById('optFilterRows').checked;searchMatchedRows=(searchFilterRows&&searchResults.length>0)?new Set(searchResults.map(function(sr){return sr.row})):new Set();var gw=document.getElementById('gridWrapper');if(gw)gw.scrollTop=0;fullRender();if(searchFilterRows&&searchResults.length>0){searchCurrentIdx=0;setTimeout(navigateToSearchResult,50)}}
function isRowHidden(ri){return!1}
function isRowHiddenBySearch(ri){return searchFilterRows&&searchMatchedRows.size>0&&!searchMatchedRows.has(ri+1)}
function captureSearchScope(){if(selectedCols.length>0){searchScope={rows:null,cols:selectedCols.filter(function(c){return!_currentState.hiddenCols.has(c-1)})};return}
if(selectedRows.length>0){searchScope={rows:selectedRows.filter(function(r){return!isRowHidden(r-1)&&!isRowHiddenBySearch(r-1)}),cols:null};return}
var rSet=new Set(),cSet=new Set(),cnt=0;if(selRange){for(var r=selRange.r1;r<=selRange.r2;r++){if(isRowHidden(r-1)||isRowHiddenBySearch(r-1))continue;for(var c=selRange.c1;c<=selRange.c2;c++){if(!_currentState.hiddenCols.has(c-1)){rSet.add(r);cSet.add(c);cnt++}}}}else if(selCells){selCells.forEach(function(k){var p=k.split(','),r=parseInt(p[0]),c=parseInt(p[1]);if(!isRowHidden(r-1)&&!isRowHiddenBySearch(r-1)&&!_currentState.hiddenCols.has(c-1)){rSet.add(r);cSet.add(c);cnt++}})}
if(cnt>1){searchScope={rows:Array.from(rSet).sort(function(a,b){return a-b}),cols:Array.from(cSet).sort(function(a,b){return a-b})};return}searchScope=null}
function updateScopeLabel(){var el=document.getElementById('searchScopeLabel');if(!el)return;if(!searchScope){el.textContent='Toàn bộ';el.style.color='var(--md-on-surface-3)'}else if(searchScope.rows&&!searchScope.cols){el.textContent='Dòng: '+searchScope.rows.length;el.style.color='var(--md-primary)'}else if(!searchScope.rows&&searchScope.cols){el.textContent='Cột: '+searchScope.cols.length;el.style.color='var(--md-primary)'}else{el.textContent='Vùng: '+searchScope.rows.length+'D x '+searchScope.cols.length+'C';el.style.color='var(--md-primary)'}}
function clearSearchScope(){searchScope=null;updateScopeLabel();if(searchTerm)doSearch(searchTerm);}
function autoUpdateSearchScope(){var sp=document.getElementById('searchPanel');if(sp&&sp.classList.contains('show')){captureSearchScope();updateScopeLabel()}}
function switchSearchTab(tab){document.getElementById('sTabFind').classList.toggle('active',tab==='find');document.getElementById('sTabReplace').classList.toggle('active',tab==='replace');document.getElementById('rowExclude').style.display=(tab==='find')?'flex':'none';document.getElementById('rowReplace').style.display=(tab==='replace')?'flex':'none';document.getElementById('actionFind').style.display=(tab==='find')?'flex':'none';document.getElementById('actionReplace').style.display=(tab==='replace')?'flex':'none';document.getElementById('optFilterWrap').style.display=(tab==='find')?'flex':'none';document.getElementById('searchInput').focus()}