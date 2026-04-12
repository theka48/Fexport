#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent()
SetWorkingDir(A_ScriptDir)
#Include WebView2.ahk
; #Include XL.ahk

; ============================================================
;                        GLOBALS
; ============================================================

global aiThread := 0
global aiThreadRunning := false
global totalExportRows := 0

; ĐỔI SỐ NÀY ĐỂ TRÁNH TRÙNG LẶP (0x0400 + 0x1500)
global WM_AI_PROGRESS := 0x1901 
global WM_AI_DONE     := 0x1902

OnMessage(WM_AI_PROGRESS, OnAiProgress, 20)
OnMessage(WM_AI_DONE,     OnAiDone,     20)

global wvc          := unset
global wv           := unset
global wvReady      := false
global pageReady    := false
global TempHtmlPath := A_Temp "\xl_import_wv2.html"

global rulesFolder  := A_ScriptDir "\Rules"
global currentRules := []   ; [{Priority,RuleID,Type,Column,Match,Output,Options,Enabled,Description}]
global allLayers    := []   ; layer names từ Illustrator
global rowLayerData := Map() ; rowIndex → [layerName, ...]
global aiApp        := ""
global aiPreviewFolder := ""
global cancelFlagFile := A_Temp "\ai_export_cancel.flag"

OnExit(OnUnload)

; ============================================================
;                        FONTS
; ============================================================
OnLoad() {
    local fonts := ["NotoSansJP-Medium", "NotoSansJP-Regular", "NotoSansJP-SemiBold"]
    for f in fonts
        DllCall("Gdi32.dll\AddFontResourceEx", "Str", "Lib\Fonts\" f ".ttf", "UInt", 0x10, "UInt", 0)
    if !DirExist(rulesFolder)
        DirCreate(rulesFolder)
}

OnUnload(ExitReason, ExitCode) {
    local fonts := ["NotoSansJP-Medium", "NotoSansJP-Regular", "NotoSansJP-SemiBold"]
    for f in fonts
        DllCall("Gdi32.dll\RemoveFontResourceEx", "Str", "Lib\Fonts\" f ".ttf", "UInt", 0x10, "UInt", 0)
}

OnLoad()

; ============================================================
;                        GUI
; ============================================================
MyGui := Gui("+Resize", "XL Import")
MyGui.MarginX := MyGui.MarginY := 0
MyGui.OnEvent("Size",  GuiResize)
MyGui.OnEvent("Close", GuiClose)

wvContainer  := MyGui.Add("Text", "x0 y0 w1400 h800 vWVC")
LoadingLabel := MyGui.Add("Text", "x0 y0 w1400 h800 Center vLoadingText", "Đang khởi động...")
LoadingLabel.SetFont("s14")
MyGui.Show("w1400 h800")

SetTimer(InitWebView2, -1)

; ============================================================
;                        WEBVIEW2 INIT
; ============================================================
; InitWebView2() {
;     global wvc, wv, wvReady
;     try {
;         wvc := WebView2.CreateControllerAsync(wvContainer.Hwnd).await2()
;         wv  := wvc.CoreWebView2
;         wv.Settings.AreDefaultContextMenusEnabled := false
;         wv.Settings.IsZoomControlEnabled          := false
;         ; wv.Settings.AreDevToolsEnabled            := true
;         wv.add_WebMessageReceived(OnWebMessage)
;         wv.add_NavigationCompleted(OnNavigationCompleted)
;         wvReady := true
;         LoadPageHTML()
;     } catch as e {
;         MsgBox("Lỗi WebView2:`n" e.Message, "Lỗi", "Icon!")
;     }
; }

InitWebView2() {
    global wvc, wv, wvReady
    try {
        ; === ĐÂY LÀ PHẦN SỬA LỖI ===

        ; 1. Tạo một đối tượng Options
        local creationOptions := Object()
        creationOptions.DefaultBackgroundColor := 0 ; 0 = Transparent

        ; 2. Truyền đối tượng options này làm tham số thứ hai
        wvc := WebView2.CreateControllerAsync(wvContainer.Hwnd, creationOptions).await2()
        
        ; === KẾT THÚC PHẦN SỬA LỖI ===

        wv  := wvc.CoreWebView2
        wv.Settings.AreDefaultContextMenusEnabled := false
        wv.Settings.IsZoomControlEnabled          := false
        ; wv.Settings.AreDevToolsEnabled            := true
        wv.add_WebMessageReceived(OnWebMessage)
        wv.add_NavigationCompleted(OnNavigationCompleted)
        wvReady := true
        LoadPageHTML()
    } catch as e {
        MsgBox("Lỗi WebView2:`n" e.Message, "Lỗi", "Icon!")
    }
}

; Build HTML một lần duy nhất → ghi file → navigate
LoadPageHTML() {
    global wv, TempHtmlPath, pageReady
    pageReady := false
    html := BuildHTMLContent()
    try FileDelete(TempHtmlPath)
    f := FileOpen(TempHtmlPath, "w", "UTF-8-RAW")
    f.Write(Chr(0xFEFF) html)
    f.Close()
    wv.Navigate("file:///" StrReplace(TempHtmlPath, "\", "/"))
}

BuildHTMLContent() {
    html := FileRead(A_ScriptDir "\xl_import.html", "UTF-8")
    
    ; 1. CSS
    cssContent := ""
    Loop Files, A_ScriptDir "\css\*.css"
    {
        cssContent .= "/* --- " A_LoopFileName " --- */`n"
        cssContent .= FileRead(A_LoopFileFullPath, "UTF-8") "`n`n"
    }

    ; 2. JS - Load HyperFormula trước từ Lib\
    jsContent := ""
    hfPath    := A_ScriptDir "\Lib\hyperformula.full.min.js"
    if FileExist(hfPath)
        jsContent .= FileRead(hfPath, "UTF-8") "`n`n"

    Loop Files, A_ScriptDir "\js\*.js"
    {
        jsContent .= "/* --- " A_LoopFileName " --- */`n"
        jsContent .= FileRead(A_LoopFileFullPath, "UTF-8") "`n`n"
    }

    html := StrReplace(html, "<!-- INJECT_CSS -->", "<style>`n" cssContent "</style>")
    html := StrReplace(html, "<!-- INJECT_JS -->", "<script>`n" jsContent "</script>")
    
    return html
}

OnNavigationCompleted(handler, args) {
    global pageReady, MyGui
    pageReady := true
    try MyGui["LoadingText"].Visible := false
    ; try wv.OpenDevToolsWindow()
    SetTimer(() => RefreshRulesListToJS(),  -200)
    SetTimer(() => SendFormulaListToJS(),   -300)
}

; ============================================================
;                        RESIZE / CLOSE
; ============================================================
GuiResize(thisGui, MinMax, Width, Height) {
    global wvc
    if MinMax = -1
        return
    thisGui["WVC"].Move(0, 0, Width, Height)
    try wvc.Fill()
}

GuiClose(thisGui) {
    global TempHtmlPath, wvc, wv
    try {
        wv.remove_NavigationCompleted(OnNavigationCompleted)
        wv.remove_WebMessageReceived(OnWebMessage)
        wv := 0
    }
    try FileDelete(TempHtmlPath)
    try wvc.Close()
    wvc := 0
    ExitApp()
}

; ============================================================
;                        MESSAGES FROM JS
; ============================================================
OnWebMessage(handler, args) {
    raw    := args.TryGetWebMessageAsString()
    action := JSONGet(raw, "action")
    switch action {
        case "importLayers":    SetTimer(() => ImportLayersFromAI(),     -1)
        case "applyRules":      SetTimer(() => ApplyAllRulesAHK(),       -1)
        case "saveRules":       SetTimer(() => SaveRulesFromJS(raw),     -1)
        case "loadRulesFile":   SetTimer(() => LoadRulesFileFromJS(raw), -1)
        case "browseRulesFile": SetTimer(() => BrowseRulesFile(),        -1)
        case "refreshRules":    SetTimer(() => RefreshRulesListToJS(),   -1)
        case "rowSelected":     SetTimer(() => OnRowSelected(raw),       -1)
        case "toggleLayer":     SetTimer(() => OnToggleLayer(raw),       -1)
        case "clearRowLayers":  SetTimer(() => ClearRowLayers(raw),      -1)
        case "testAI":          SetTimer(() => TestAIConnection(),       -1)
        case "getFormatList":   SetTimer(() => SendFormatListToJS(),     -1)
        case "getFormulaList":  SetTimer(() => SendFormulaListToJS(),    -1)
        case "requestPdfExportUI":   SetTimer(() => GetAiPdfPresets(), -1)
        case "browseFolder":         SetTimer(() => BrowseFolderForJS(raw), -1)
        case "startAiExportThread":  SetTimer(() => StartAiExport(raw), -1)
        case "cancelAiExportThread": SetTimer(() => CancelAiExport(), -1)
        case "previewAiRow":         SetTimer(() => PreviewAiRow(raw), -1)
        case "clearAiPreviewLayers": SetTimer(() => ClearAiPreviewLayers(), -1)
        case "importCsvTab3": SetTimer(() => ImportCsvTab3AHK(), -1)
        case "saveCsvViaAHK": SetTimer(() => SaveCsvViaAHK(raw), -1)
        
    }
}

SendFormatListToJS() {
    global wv
    formatFolder := A_ScriptDir "\Format"
    formats := []

    ; Hàm nội bộ — đọc tất cả *.csv trong một thư mục
    CollectFromDir(dir, &fmts) {
        Loop Files, dir "\*.csv" {
            baseName := StrReplace(A_LoopFileName, ".csv", "")
            try {
                ; Đọc toàn bộ nội dung file — JS sẽ parse từng dòng
                fileContent := FileRead(A_LoopFilePath, "UTF-8")
            } catch {
                continue
            }
            ; Bỏ qua file hoàn toàn rỗng hoặc chỉ có comment
            hasDef := false
            for ln in StrSplit(fileContent, "`n", "`r") {
                t := Trim(ln)
                if t != "" && SubStr(t, 1, 1) != ";" {
                    hasDef := true
                    break
                }
            }
            if hasDef
                fmts.Push({ name: baseName, content: fileContent })
        }
    }

    if DirExist(formatFolder) {
        ; Quét thẳng trong Format\
        CollectFromDir(formatFolder, &formats)
        ; Quét các subfolder 1 cấp
        Loop Files, formatFolder "\*", "D"
            CollectFromDir(A_LoopFileFullPath, &formats)
    }

    if formats.Length = 0 {
        try wv.ExecuteScript("onImportFormatList([]);")
        return
    }

    ; Build JSON thủ công — gửi {name, content} cho mỗi file
    json := "["
    for i, f in formats {
        if i > 1
            json .= ","
        json .= '{"name":"' EscJS(f.name) '","content":"' EscJS(f.content) '"}'
    }
    json .= "]"
    try wv.ExecuteScript("onImportFormatList(" json ");")
}

; ============================================================
;                        FORMULA LIST
; ============================================================
SendFormulaListToJS() {
    global wv
    formulaFolder := A_ScriptDir "\Formula"
    formulas := []

    if DirExist(formulaFolder) {
        Loop Files, formulaFolder "\*.js" {
            try {
                fileContent := FileRead(A_LoopFilePath, "UTF-8")
            } catch {
                continue
            }
            baseName := StrReplace(A_LoopFileName, ".js", "")
            formulas.Push({ name: baseName, content: fileContent })
        }
    }

    if formulas.Length = 0 {
        try wv.ExecuteScript("onFormulaList([]);")
        return
    }

    json := "["
    for i, f in formulas {
        if i > 1
            json .= ","
        json .= '{"name":"' EscJS(f.name) '","content":"' EscJS(f.content) '"}'
    }
    json .= "]"
    try wv.ExecuteScript("onFormulaList(" json ");")
}

ImportLayersFromAI() {
    global wv, allLayers, aiApp
    aiApp := ""
    ; for ver in [30, 29, 28, 27, 26, 25, 24] {
    for ver in [25] {
        try {
            aiApp := ComObjActive("Illustrator.Application." ver)
            break
        }
    }
    if !IsObject(aiApp) {
        JSToast("Không kết nối được Illustrator")
        return
    }
    if aiApp.Documents.Count = 0 {
        JSToast("Chưa mở document nào trong Illustrator")
        return
    }
    
    doc    := aiApp.ActiveDocument
    layers := doc.Layers
    
    jsonStr := "["
    Loop layers.Count {
        lyr := layers.Item(A_Index)
        name := lyr.Name
        
        ; Mặc định màu xám nếu không đọc được màu
        r := 200, g := 200, b := 200 
        try {
            c := lyr.Color
            r := Floor(c.Red)
            g := Floor(c.Green)
            b := Floor(c.Blue)
        }
        
        if (A_Index > 1)
            jsonStr .= ","
            
        ; SỬ DỤNG NHÁY ĐƠN VÀ DẤU CHẤM ĐỂ NỐI CHUỖI CHUẨN AHK V2
        jsonStr .= '{"name":"' . EscJS(name) . '","color":"' . r . ',' . g . ',' . b . '"}'
    }
    jsonStr .= "]"

    try wv.ExecuteScript("onLayersImported(" jsonStr ");")
}

TestAIConnection() {
    global wv
    ; for ver in [30, 29, 28, 27, 26, 25, 24] {
    for ver in [25] {
        try {
            ai      := ComObjActive("Illustrator.Application." ver)
            docName := ai.Documents.Count > 0 ? ai.ActiveDocument.Name : "(no doc)"
            JSToast("✅ Kết nối AI OK: " docName)
            return
        }
    }
    JSToast("❌ Không kết nối được Illustrator")
}

; ============================================================
;                        ROW SELECTION / LAYER TOGGLE
; ============================================================
OnRowSelected(raw) {
    global wv, rowLayerData
    rowStr := JSONGet(raw, "row")
    if rowStr = ""
        return
    rowIdx := Integer(rowStr)
    if rowIdx < 1
        return
    layers := rowLayerData.Has(rowIdx) ? rowLayerData[rowIdx] : []
    try wv.ExecuteScript("onRowLayersLoaded(" JSON.stringify(layers) ");")
}

OnToggleLayer(raw) {
    global rowLayerData
    rowStr := JSONGet(raw, "row")
    if rowStr = ""
        return
    rowIdx    := Integer(rowStr)
    layerName := JSONGet(raw, "layer")
    checked   := JSONGet(raw, "checked") = "true"
    if rowIdx < 1
        return
    layers    := rowLayerData.Has(rowIdx) ? rowLayerData[rowIdx].Clone() : []
    if checked {
        found := false
        for l in layers
            if l = layerName { 
                found := true
             break 
            }
        if !found
            layers.Push(layerName)
    } else {
        newLayers := []
        for l in layers
            if l != layerName
                newLayers.Push(l)
        layers := newLayers
    }
    rowLayerData[rowIdx] := layers
}

ClearRowLayers(raw) {
    global rowLayerData, wv
    rowStr := JSONGet(raw, "row")
    if rowStr = ""
        return
    rowIdx := Integer(rowStr)
    if rowIdx > 0
        rowLayerData[rowIdx] := []
    try wv.ExecuteScript("onRowLayersLoaded([]);")
}

; ============================================================
;                        APPLY ALL RULES
; ============================================================
ApplyAllRulesAHK() {
    global wv, currentRules, allLayers
    if !currentRules.Length {
        JSToast("Chưa load rules!")
        return
    }
    if !allLayers.Length {
        JSToast("Chưa import layers từ Illustrator!")
        return
    }
    try wv.ExecuteScript("sendTableDataForRules();")
}

; ============================================================
;                        RULES FILE MANAGEMENT
; ============================================================
RefreshRulesListToJS() {
    global wv, rulesFolder
    files := []
    Loop Files, rulesFolder "\*.csv"
        files.Push(A_LoopFileName)
    try wv.ExecuteScript("onRulesFileList(" JSON.stringify(files) ");")
}

LoadRulesFileFromJS(raw) {
    global wv, rulesFolder
    fileName := JSONGet(raw, "file")
    filePath := rulesFolder "\" fileName
    if !FileExist(filePath) {
        JSToast("Không tìm thấy file: " fileName)
        return
    }
    LoadRulesCSV(filePath)
}

BrowseRulesFile() {
    global rulesFolder
    filePath := FileSelect("1", rulesFolder, "Chọn file Rules", "CSV (*.csv)")
    if !filePath
        return
    LoadRulesCSV(filePath)
}

LoadRulesCSV(filePath) {
    global wv, currentRules
    currentRules := []
    try {
        content  := FileRead(filePath, "UTF-8")
        lines    := StrSplit(content, "`n", "`r")
        isHeader := true
        for line in lines {
            if !line || Trim(line) = ""
                continue
            if isHeader {
                isHeader := false
                continue
            }
            cols := ParseCSVLine(line)
            if cols.Length < 8
                continue
            currentRules.Push({
                Priority:    cols[1],
                RuleID:      cols[2],
                Type:        cols[3],
                Column:      cols[4],
                Match:       cols[5],
                Output:      cols[6],
                Options:     cols[7],
                Enabled:     cols[8],
                Description: cols.Length >= 9 ? cols[9] : "",
                Color:       cols.Length >= 10 ? cols[10] : ""
            })
        }
        try wv.ExecuteScript("onRulesLoaded(" JSON.stringify(currentRules) ");")
        SplitPath(filePath, &fname)
        JSToast("Đã load: " fname " (" currentRules.Length " rules)")
    } catch as e {
        JSToast("Lỗi đọc file: " SubStr(e.Message, 1, 80))
    }
}

SaveRulesFromJS(raw) {
    global wv, rulesFolder, currentRules
    fileName := JSONGet(raw, "fileName")
    if fileName = ""
        fileName := "new_rules.csv"
    filePath := rulesFolder "\" fileName

    ; Parse rules từ JSON do JS gửi
    rulesJson    := JSONGetArray(raw, "rules")
    
    ; BẢN SỬA LỖI: Kiểm tra xem biến JSON có tồn tại hàm parse hay không.
    try {
        parsedRules  := JSON.parse(rulesJson)
        if IsObject(parsedRules) && parsedRules.Length > 0
            currentRules := parsedRules
    } catch {
        JSToast("Lỗi phân tích JSON khi lưu Rule!")
        return
    }

    ; Build CSV
    content := "Priority,RuleID,Type,Column,Match,Output,Options,Enabled,Description`n"
    for rule in currentRules {
        ; TÙY VÀO THƯ VIỆN JSON AHK CỦA BẠN, NÓ CÓ THỂ LÀ MAP [] HOẶC OBJECT .
        ; Tôi dùng toán tử try/catch hoặc Has để an toàn 100% lấy được giá trị
        
        pri := (rule.Has("Priority") ? rule["Priority"] : (rule.HasProp("Priority") ? rule.Priority : ""))
        rid := (rule.Has("RuleID") ? rule["RuleID"] : (rule.HasProp("RuleID") ? rule.RuleID : ""))
        typ := (rule.Has("Type") ? rule["Type"] : (rule.HasProp("Type") ? rule.Type : ""))
        col := (rule.Has("Column") ? rule["Column"] : (rule.HasProp("Column") ? rule.Column : ""))
        mat := (rule.Has("Match") ? rule["Match"] : (rule.HasProp("Match") ? rule.Match : ""))
        out := (rule.Has("Output") ? rule["Output"] : (rule.HasProp("Output") ? rule.Output : ""))
        opt := (rule.Has("Options") ? rule["Options"] : (rule.HasProp("Options") ? rule.Options : ""))
        enb := (rule.Has("Enabled") ? rule["Enabled"] : (rule.HasProp("Enabled") ? rule.Enabled : ""))
        dsc := (rule.Has("Description") ? rule["Description"] : (rule.HasProp("Description") ? rule.Description : ""))
        clr := (rule.Has("Color") ? rule["Color"] : (rule.HasProp("Color") ? rule.Color : ""))

        content .= pri ","
        content .= rid ","
        content .= typ ","
        content .= col ","
        content .= '"' StrReplace(mat, '"', '""') '",'
        content .= '"' StrReplace(out, '"', '""') '",'
        content .= '"' StrReplace(opt, '"', '""') '",'
        content .= enb ","
        content .= '"' StrReplace(dsc, '"', '""') '",'
        content .= '"' StrReplace(clr, '"', '""') '"' "`n"
    }
    try {
        if FileExist(filePath)
            FileDelete(filePath)
        FileAppend(content, filePath, "UTF-8")
        JSToast("✅ Đã lưu: " fileName)
        RefreshRulesListToJS()
    } catch as e {
        JSToast("Lỗi lưu: " SubStr(e.Message, 1, 80))
    }
}

; ============================================================
;                        HELPERS
; ============================================================

; Gọi hàm toast() trong JS
JSToast(msg) {
    global wv
    try wv.ExecuteScript("toast('" EscJS(msg) "');")
}

; Escape chuỗi để nhúng vào JS string literal (dấu nháy đơn)
EscJS(s) {
    s := StrReplace(s, "\",  "\\")
    s := StrReplace(s, '"',  '\"')
    s := StrReplace(s, "'",  "\'")
    s := StrReplace(s, "`n", "\n")
    s := StrReplace(s, "`r", "\r")
    return s
}

; Lấy giá trị string của key trong JSON (dùng cho các message ngắn từ JS)
JSONGet(json, key) {
    needle := '"' key '":"'
    pos    := InStr(json, needle)
    if !pos
        return ""
    pos += StrLen(needle)
    result := ""
    esc    := false
    Loop Parse, SubStr(json, pos), "" {
        c := A_LoopField
        if esc {
            result .= c
            esc := false
            continue
        }
        if c = "\" {
            esc := true
            continue
        }
        if c = '"'
            break
        result .= c
    }
    return result
}

; Trích array JSON dạng [...] theo key (dùng để lấy "rules":[...])
JSONGetArray(json, key) {
    needle := '"' key '":['
    pos    := InStr(json, needle)
    if !pos
        return "[]"
    pos   += StrLen(needle) - 1
    depth := 0
    result := ""
    inQ   := false
    esc   := false
    Loop Parse, SubStr(json, pos), "" {
        c := A_LoopField
        result .= c
        if esc {
            esc := false
            continue
        }
        if c = "\" && inQ {
            esc := true
            continue
        }
        if c = '"' {
            inQ := !inQ
            continue
        }
        if !inQ {
            if c = "["
                depth++
            else if c = "]" {
                depth--
                if depth = 0
                    break
            }
        }
    }
    return result
}

; Parse một dòng CSV (hỗ trợ quoted fields và "" escaped quotes bên trong)
ParseCSVLine(line) {
    cols    := []
    current := ""
    inQuote := false
    i       := 1
    len     := StrLen(line)
    while i <= len {
        c := SubStr(line, i, 1)
        if inQuote {
            if c = '"' {
                ; Kiểm tra "" (escaped quote)
                if SubStr(line, i + 1, 1) = '"' {
                    current .= '"'
                    i += 2
                    continue
                }
                inQuote := false
            } else {
                current .= c
            }
        } else {
            if c = '"' {
                inQuote := true
            } else if c = "," {
                cols.Push(current)
                current := ""
            } else {
                current .= c
            }
        }
        i++
    }
    cols.Push(current)
    return cols
}

; ============================================================
;               MULTI-THREADING AI EXPORT (HTML UI + thqby AHK_H)
; ============================================================

GetAiPdfPresets() {
    global wv
    aiApp := ""
    
    ; Quét tìm COM của phiên bản AI đang mở (Giống hệt Tab 1)
    ; for ver in [30, 29, 28, 27, 26, 25, 24] {
    for ver in [25] {
        try {
            aiApp := ComObjActive("Illustrator.Application." ver)
            break
        }
    }
    
    if (aiApp = "") {
        JSToast("Lỗi: Không tìm thấy Illustrator đang mở!")
        return
    }
    
    try {
        presets := aiApp.PDFPresetsList
        json := "["
        Loop presets.MaxIndex() + 1 {
            json .= '"' presets[A_Index - 1] '"' (A_Index <= presets.MaxIndex() ? "," : "")
        }
        json .= "]"
        
        wv.ExecuteScript("showPdfExportDialog(" json ");")
    } catch as e {
        JSToast("Lỗi đọc PDF Preset: " e.Message)
    }
}

BrowseFolderForJS(raw) {
    global wv
    targetId := JSONGet(raw, "target")
    folder := DirSelect("", 3, "Chọn thư mục")
    if (folder != "") {
        script := "setFolderPath('" targetId "', '" StrReplace(folder, "\", "\\") "');"
        wv.ExecuteScript(script)
    }
}

StartAiExport(raw) {
    global wv, aiThread, aiThreadRunning, totalExportRows, MyGui, cancelFlagFile

    if (aiThreadRunning) {
        JSToast("Tiến trình AI đang chạy rồi!")
        return
    }

    ; ✅ Chỉ cần xóa file flag cũ (nếu có)
    try FileDelete(cancelFlagFile)

    try {
        parsedObj    := JSON.parse(raw)
        csvData      := parsedObj.Has("csvData") ? parsedObj["csvData"] : parsedObj.csvData
        sourceFolder := parsedObj.Has("sourceFolder") ? parsedObj["sourceFolder"] : parsedObj.sourceFolder
        outputFolder := parsedObj.Has("outputFolder") ? parsedObj["outputFolder"] : parsedObj.outputFolder
        presetName   := parsedObj.Has("presetName") ? parsedObj["presetName"] : parsedObj.presetName
    } catch {
        JSToast("Lỗi phân tích dữ liệu JSON từ WebView2!")
        return
    }
    
    tempDataFile := A_Temp "\ai_export_data.txt"
    try FileDelete(tempDataFile)
    jsxPath := A_ScriptDir "\Ai_ExportPDF_Headless.jsx"
    
    FileAppend(sourceFolder "`n" outputFolder "`n" presetName "`n", tempDataFile, "UTF-8")
    
    validRows := 0
    for row in csvData {
        if (row.Length = 0 || row[1] = "")
            continue
            
        pdfName := StrReplace(row[1], '\', '\\')
        pdfName := StrReplace(pdfName, '"', '\"')
        
        columnsJson := "["
        loop row.Length {
            if (A_Index > 1) {
                colVal := StrReplace(row[A_Index], '\', '\\')
                colVal := StrReplace(colVal, '"', '\"')
                colVal := StrReplace(colVal, "`n", "\n")
                columnsJson .= '"' colVal '"' (A_Index < row.Length ? "," : "")
            }
        }
        columnsJson .= "]"
        
        rowJson := '{"pdfName":"' pdfName '", "columns":' columnsJson '}'
        FileAppend(rowJson "`n", tempDataFile, "UTF-8")
        validRows++
    }
    
    if (validRows = 0) {
        JSToast("Không có dòng dữ liệu hợp lệ để xuất!")
        return
    }
    totalExportRows := validRows
    
    wv.ExecuteScript("document.getElementById('aiProgressOverlay').style.display = 'flex';")
    wv.ExecuteScript("updateAiProgress(0, " totalExportRows ", 'Đang khởi động Illustrator...');")
    
    aiThreadRunning := true
    hwndMain := MyGui.Hwnd
    
    threadFile := A_ScriptDir "\Ai_Export_Thread.ahk"
    threadCode := FileRead(threadFile, "UTF-8")
    
    ; ✅ Không còn tham số nào liên quan đến cancel hay start flag
    aiThread := Worker(threadCode, '"' hwndMain '" "' tempDataFile '" "' jsxPath '"')
}

CancelAiExport() {
    global aiThreadRunning, wv, cancelFlagFile
    if (aiThreadRunning) {
        ; ✅ Chỉ tạo file flag để yêu cầu worker dừng
        FileAppend("cancel", cancelFlagFile)
        
        ; ✅ Cập nhật UI để báo cho người dùng biết là đang chờ hủy
        wv.ExecuteScript("updateAiProgress(totalExportRows, totalExportRows, 'Đang chờ hủy...');")
    }
}

OnAiProgress(wParam, lParam, msg, hwnd) {
    global wv, totalExportRows
    
    ; ✅ Bỏ hoàn toàn logic `isCancelled`
    if (lParam = 0)
        return 0

    current := wParam
    try { 
        pdfName := StrGet(lParam) 
    } catch { 
        pdfName := "(Đang xuất...)" 
    }
    script := "updateAiProgress(" current ", " totalExportRows ", '" EscJS(pdfName) "');"
    try wv.ExecuteScript(script)
    return 0
}

OnAiDone(wParam, lParam, msg, hwnd) {
    global wv, aiThreadRunning, aiThread
    
    ; ✅ Reset trạng thái running
    aiThreadRunning := false
    aiThread := 0
    
    ; ✅ Bỏ logic `isCancelled`
    if (lParam = 0 || !lParam)
        return 0
    
    try { 
        report := StrGet(lParam) 
    } catch { 
        report := "Lỗi: Không thể đọc báo cáo" 
    }
    script := "finishAiExport('" EscJS(report) "');"
    try wv.ExecuteScript(script)
    return 0
}

PreviewAiRow(raw) {
    global aiPreviewFolder
    try {
        parsedObj := JSON.parse(raw)
        cols := parsedObj.Has("columns") ? parsedObj["columns"] : parsedObj.columns
        hasLink := parsedObj.Has("hasLink") ? parsedObj["hasLink"] : parsedObj.hasLink
        
        ; Nếu có hình ảnh, yêu cầu Folder 1 lần duy nhất cho phiên làm việc
        if (hasLink && aiPreviewFolder = "") {
            aiPreviewFolder := DirSelect("", 3, "Vui lòng chọn Thư mục chứa File Link (@) để XEM TRƯỚC (Preview)")
            if (aiPreviewFolder = "") {
                JSToast("⚠️ Đã hủy Xem trước vì thiếu Thư mục Link!")
                return
            }
        }
        
        columnsJson := "["
        loop cols.Length {
            colVal := StrReplace(cols[A_Index], '\', '\\')
            colVal := StrReplace(colVal, '"', '\"')
            columnsJson .= '"' colVal '"' (A_Index < cols.Length ? "," : "")
        }
        columnsJson .= "]"
        
        rowJson := '{"isPreview":true, "columns":' columnsJson ', "sourceFolder":"' StrReplace(aiPreviewFolder, '\', '\\') '"}'
        
        aiApp := ""
        ; for ver in [30, 29, 28, 27, 26, 25, 24] {
        for ver in [25] {
            try { 
                aiApp := ComObjActive("Illustrator.Application." ver)
                break 
            }
        }
        if (aiApp = "") 
            return
        
        jsxPath := A_ScriptDir "\Ai_ExportPDF_Headless.jsx"
        jsxContent := FileRead(jsxPath, "UTF-8")
        
        safeConfig := StrReplace(rowJson, '\', '\\')
        safeConfig := StrReplace(safeConfig, '"', '\"')
        safeConfig := StrReplace(safeConfig, "'", "\'")
        
        scriptToRun := "var configStr = '" safeConfig "';`n" jsxContent "`nprocessSingleRow(configStr);"
        
        try {
            result := aiApp.DoJavaScript(scriptToRun)
            
            ; Nếu kết quả bắt đầu bằng chữ LỖI hoặc ERROR, bắn thẳng bảng báo cáo lên Web
            if (InStr(result, "LỖI") || InStr(result, "ERROR")) {
                errorMsg := StrReplace(result, "\", "\\")
                errorMsg := StrReplace(errorMsg, "`n", "\n")
                errorMsg := StrReplace(errorMsg, "`r", "")
                errorMsg := StrReplace(errorMsg, "'", "\'")
                
                wv.ExecuteScript("finishAiExport('⚠️ PREVIEW THẤT BẠI:\\n" errorMsg "');")
            }
        } catch as e {
            wv.ExecuteScript("finishAiExport('⚠️ PREVIEW LỖI HỆ THỐNG:\\n" StrReplace(e.Message, "\", "\\") "');")
        }
    }
}

; Hàm dọn dẹp Layer Preview khi người dùng tắt Checkbox
ClearAiPreviewLayers() {
    aiApp := ""
    ; for ver in [30, 29, 28, 27, 26, 25, 24] {
    for ver in [25] {
        try { 
            aiApp := ComObjActive("Illustrator.Application." ver)
        break 
        }
    }
    if (aiApp = "" || aiApp.Documents.Count = 0) 
        return
    
    ; ✅ Sửa lại cú pháp
    cleanScript := "
    (
        var doc = app.activeDocument;
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            if (doc.layers[i].name.indexOf('_PREVIEW') > -1) {
                try { doc.layers[i].remove(); } catch(e) {}
            }
        }
        app.redraw();
    )"
    
    try aiApp.DoJavaScript(cleanScript)
}

ImportCsvTab3AHK() {
    global wv
    filePath := FileSelect("1",, "Chọn file CSV", "CSV (*.csv)")
    if !filePath
        return
    try {
        content := FileRead(filePath, "UTF-8")
    } catch {
        try content := FileRead(filePath, "CP932")
        catch {
            JSToast("❌ Không đọc được file!")
            return
        }
    }
    SplitPath(filePath, &fileName)
    safeContent  := EscJS(content)
    safeName     := EscJS(fileName)
    safeFilePath := EscJS(filePath)  ; ✨ Gửi luôn đường dẫn thật
    try wv.ExecuteScript("onImportCsvFromAHK('" safeContent "','" safeName "','" safeFilePath "');")
}

SaveCsvViaAHK(raw) {
    global wv
    try {
        obj      := JSON.parse(raw)
        filePath := obj["filePath"]
        content  := obj["content"]
    } catch as e {
        JSToast("❌ Lỗi parse JSON: " e.Message)
        return
    }

    if !filePath {
        JSToast("❌ Không có đường dẫn file!")
        return
    }

    try {
        if FileExist(filePath)
            FileDelete(filePath)
        f := FileOpen(filePath, "w", "UTF-8")
        f.Write(content)
        f.Close()
        try wv.ExecuteScript("ws3_state.isModified=false;_drawFileStatus();toast('✅ Đã lưu file!');")
    } catch as e {
        JSToast("❌ Lỗi lưu file: " SubStr(e.Message, 1, 80))
    }
}
