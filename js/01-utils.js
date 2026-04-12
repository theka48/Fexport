// ============================================
// NỘI DUNG MỚI CỦA FILE: 01-utils.js
// ============================================

function toast(msg){var t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');setTimeout(function(){t.classList.remove('show')},1800)}
function safeSetText(id,v){var el=document.getElementById(id);if(el)el.textContent=v}
function setV(id,v){var el=document.getElementById(id);if(el)el.value=v}
function getV(id){var el=document.getElementById(id);return el?el.value:''}
function setDDL(id,v){var el=document.getElementById(id);if(!el)return;for(var i=0;i<el.options.length;i++)
if(el.options[i].value===v){el.selectedIndex=i;return}}
function setEnabled(id,v){var el=document.getElementById(id);if(el)el.disabled=!v}
function padZ(n){return n<10?'0'+n:String(n)}
function escHtml(s){if(!s)return'';return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function normalizeText(text,mode){if(!text)return'';if(mode==='asc')return toHalfWidth(String(text));if(mode==='jis')return toFullWidth(String(text));return String(text)}
function toHalfWidth(s){return s.replace(/[\uFF01-\uFF5E]/g,function(c){return String.fromCharCode(c.charCodeAt(0)-0xFEE0)}).replace(/\u3000/g,' ')}
function toFullWidth(s){return s.replace(/[\u0021-\u007E]/g,function(c){return String.fromCharCode(c.charCodeAt(0)+0xFEE0)}).replace(/ /g,'\u3000')}
function getOption(options,key){if(!options)return'';var parts=options.split('|');for(var i=0;i<parts.length;i++){var p=parts[i].trim();if(p.indexOf(key+':')===0)return p.slice(key.length+1);}
return''}

// HÀM saveWorkspaceState và loadWorkspaceState KHÔNG CÒN CẦN THIẾT NỮA VÀ SẼ BỊ XÓA
// function saveWorkspaceState() { ... }
// function loadWorkspaceState(wsNum) { ... }

async function switchTab(n){
    if(activeTabNum===n)return;

    // Sử dụng state của Tab 3 để kiểm tra
    if(activeTabNum === 3 && ws3_state.isModified){
        const result = await Swal.fire({
            title: 'Lưu thay đổi?',
            text: "Dữ liệu ở Tab CSV & PDF đã được chỉnh sửa.",
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#1a73e8',
            cancelButtonColor: '#80868b',
            confirmButtonText: 'Tự động Lưu & Chuyển',
            cancelButtonText: 'Ở lại'
        });
        if (result.isConfirmed) {
            try { await saveCsvTab3(); } catch(e) { console.warn('Auto-save failed:', e); }
        } else {
            return;
        }
    }
    
    // --- PHẦN THAY ĐỔI QUAN TRỌNG ---
    activeTabNum = n;
    _currentState = (n === 1) ? ws1_state : ws3_state;
    
    // Ẩn/hiện các panel và tab
    var panels=document.querySelectorAll('.tab-panel');
    for(var i=0;i<panels.length;i++)panels[i].style.display='none';
    var tabs=document.querySelectorAll('.sidebar-tab');
    for(var i=0;i<tabs.length;i++)tabs[i].classList.remove('active');
    var panel=document.getElementById('tabPanel'+n);
    if(panel)panel.style.display='flex';
    var tab=document.getElementById('sideTab'+n);
    if(tab)tab.classList.add('active');
    
    // Reset search
    var sp=document.getElementById('searchPanel');
    if(sp)sp.classList.remove('show');
    clearSearch();

    // Di chuyển các thành phần UI chung
    var gridPane=document.getElementById('gridPane');
    var toolbarEdit=document.getElementById('sharedEditToolbar');
    if(n===1){
        document.getElementById('splitArea').insertBefore(gridPane,document.getElementById('divider'));
        document.getElementById('tabPanel1').insertBefore(toolbarEdit,document.getElementById('splitArea'));
    } else if(n===3){
        document.getElementById('tab3GridContainer').appendChild(gridPane);
        document.getElementById('tabPanel3').insertBefore(toolbarEdit,document.getElementById('tab3GridContainer'));
    }

    // Cập nhật UI từ state mới
    document.getElementById('fixedRowsInput').value = _currentState.fixedRows;
    document.getElementById('fixedColsInput').value = _currentState.fixedCols;

    // Render lại hoàn toàn với state mới
    // Dùng setTimeout để đảm bảo DOM đã được di chuyển xong
    setTimeout(() => {
        clearSelFull(); // Xóa lựa chọn cũ
        initHyperFormula(); // Khởi tạo lại công thức cho state mới
        fullRender(); // Render lại grid
        _drawFileStatus(); // Vẽ lại status file cho tab hiện tại
    }, 50);
}
function colToLetter(colNum){var temp,letter='';while(colNum>0){temp=(colNum-1)%26;letter=String.fromCharCode(temp+65)+letter;colNum=(colNum-temp-1)/26}
return letter}
function letterToCol(letter){var col=0,len=letter.length;for(var i=0;i<len;i++){col+=(letter.charCodeAt(i)-64)*Math.pow(26,len-i-1)}
return col}