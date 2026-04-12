// ============================================
// NỘI DUNG MỚI CỦA FILE: 00-state.js
// ============================================

// --- CÁC BIẾN TOÀN CỤC CŨ (SẼ BỊ XÓA) ---
// var headers=['','','','',''];
// var data=[['','','','',''],['','','','',''],['','','','','']];
// var undoStack=[],redoStack=[];
// var colWidths=[],rowHeights=[],manualRowHeights={};
// var hiddenCols=new Set(),fixedWidthCols=new Set();

// --- STATE CỐ ĐỊNH (KHÔNG THAY ĐỔI THEO TAB) ---
var MIN_ROW_HEIGHT=28,MIN_COL_WIDTH=30;
var selRange=null,selCells=null,selectedRows=[],lastSelectedRow=-1;
var selectedCols=[],lastSelectedCol=-1,currentCell=null,editingCell=null;
var rowTopCache=null,colLeftCache=null;
var isResizingCol=!1,isResizingRow=!1;
var resizeColIndex=-1,resizeRowIndex=-1;
var resizeStartX=0,resizeStartY=0,resizeStartWidth=0,resizeStartHeight=0;
var resizingMultiCols=[],resizeMultiStartWidths={};
var resizingMultiRows=[],resizeMultiStartHeights={};
var justFinishedResize=!1;
var measureDiv=null,measureWidthDiv=null;
var anchorCell=null,isDragging=!1,dragStartCell=null,dragCurrentCell=null;
var pendingPasteCallback=null;
var searchResults=[],searchCurrentIdx=-1,searchTerm='';
var searchMatchSet=null,excludeResults=[],excludeCurrentIdx=-1;
var searchFilterRows=!1,searchMatchedRows=new Set(),searchScope=null;
var selectionType='range';
var zoomLevel=100;
var activeTabNum=1;
var _isImporting=!1;
var rulesColWidths=[30,25,45,90,75,55,160,160,200,300];
var _colorPickerEl=null;var _layerFloating=!1;var _layerDragging=!1;
var _layerDragOffX=0;var _layerDragOffY=0;var _layerPaneCollapsed=!1;
var hf=null;
var _importFormats=[],_formulaList=[];

// --- STATE DÀNH RIÊNG CHO TỪNG WORKSPACE (TAB) ---

// Hàm trợ giúp để tạo một đối tượng state mặc định
function createDefaultWorkspaceState() {
    return {
        // Grid data
        headers: ['', '', '', '', ''],
        data: [['', '', '', '', ''], ['', '', '', '', ''], ['', '', '', '', ''], ['', '', '', '', ''], ['', '', '', '', '']],
        // Grid layout
        colWidths: [],
        rowHeights: [],
        manualRowHeights: {},
        hiddenCols: new Set(),
        fixedWidthCols: new Set(),
        fixedRows: 0,
        fixedCols: 0,
        // History
        undoStack: [],
        redoStack: [],
        // File info
        isModified: false,
        fileName: '', // Tên file hiển thị (VD: "data.xlsx")
        _cleanFileName: '', // Tên file không có tiền tố
        _realFilePath: null, // Đường dẫn thật trên đĩa (nếu có, cho Tab 3)
        fileHandle: null, // File handle cho File System Access API
        // Tab-specific data
        displayCache: [], // Cache HTML của cell để render nhanh hơn
        // Tab 1 specific
        rowLayerData: {},
        currentRowSel: 0,
        rowMatchedRules: {},
    };
}

// Khởi tạo các workspace
var ws1_state = createDefaultWorkspaceState();
var ws3_state = createDefaultWorkspaceState();

// CON TRỎ TRẠNG THÁI HIỆN TẠI - CỰC KỲ QUAN TRỌNG
var _currentState = ws1_state;

// --- STATE CHUNG CỦA TOÀN APP (ít thay đổi) ---
// Tab 1
var allLayersFromAI=[], layerColors = {};
// Tab 2
var currentRules=[],rulesEditRow=0,rulesSelectedRow=-1,_lastHoverRow=0;