#Requires AutoHotkey v2.0
#NoTrayIcon

; Nhận tham số từ Luồng chính
hwndMain       := A_Args[1]
tempDataFile   := A_Args[2]
jsxPath        := A_Args[3]

cancelFlagFile := A_Temp "\ai_export_cancel.flag"
WM_AI_PROGRESS := 0x1901
WM_AI_DONE     := 0x1902

; Xóa flag cũ (nếu có) khi bắt đầu để đảm bảo chạy đúng
try FileDelete(cancelFlagFile)

aiApp := ""
for ver in [30, 29, 28, 27, 26, 25, 24] {
    try {
        aiApp := ComObjActive("Illustrator.Application." ver)
        break
    }
}

if (aiApp = "") {
    SendMessage(WM_AI_DONE, 0, StrPtr("Lỗi: Illustrator chưa mở!"),, "ahk_id " hwndMain)
    ExitApp()
}

jsxContent := FileRead(jsxPath, "UTF-8")
lines := StrSplit(FileRead(tempDataFile, "UTF-8"), "`n", "`r")
sourceFolder := lines[1]
outputFolder := lines[2]
presetName   := lines[3]

totalRows := lines.Length - 4
successCount := 0
failList := ""
wasCancelled := false
progressBuf := ""

; ... (Đoạn code checkScript giữ nguyên, không cần thay đổi) ...
; 1. TẠO SCRIPT KIỂM TRA LỖI (Gửi toàn bộ data qua AI để check trước)
checkScript := "var configList = ["
Loop totalRows {
    idx := A_Index + 3
    if (lines[idx] != "")
        checkScript .= lines[idx] ","
}
checkScript .= "];`n"
checkScript .= "var sourceFolder = '" StrReplace(sourceFolder, '\', '\\') "';`n"
checkScript .= jsxContent "`n"
checkScript .= "checkAllDataBeforeExport(configList, sourceFolder);"

; Thực thi kiểm tra lỗi
try {
    checkResult := aiApp.DoJavaScript(checkScript)
    if (checkResult != "OK") {
        SendMessage(WM_AI_DONE, 0, StrPtr(checkResult),, "ahk_id " hwndMain)
        ExitApp()
    }
} catch as e {
    SendMessage(WM_AI_DONE, 0, StrPtr("Lỗi quá trình kiểm tra Data: " e.Message),, "ahk_id " hwndMain)
    ExitApp()
}

; 2. NẾU KHÔNG CÓ LỖI -> BẮT ĐẦU XUẤT TỪNG FILE
Loop totalRows {
    ; ✅ KIỂM TRA HỦY
    if FileExist(cancelFlagFile) {
        wasCancelled := true
        break ; Thoát khỏi vòng lặp
    }
    
    idx := A_Index + 3
    rowJson := lines[idx]
    if (rowJson = "")
        continue
        
    RegExMatch(rowJson, '"pdfName":"(.*?)"', &match)
    pdfName := match ? match[1] : "Unknown"
    
    progressBuf := pdfName
    PostMessage(WM_AI_PROGRESS, A_Index, StrPtr(progressBuf),, "ahk_id " hwndMain)
    
    jsxConfig := SubStr(rowJson, 1, StrLen(rowJson)-1)
    jsxConfig .= ',"sourceFolder":"' StrReplace(sourceFolder, '\', '\\') '"'
    jsxConfig .= ',"outputFolder":"' StrReplace(outputFolder, '\', '\\') '"'
    jsxConfig .= ',"presetName":"' presetName '"}'
    
    safeConfig := StrReplace(jsxConfig, '\', '\\')
    safeConfig := StrReplace(safeConfig, '"', '\"')
    safeConfig := StrReplace(safeConfig, "'", "\'")
    
    scriptToRun := "var configStr = '" safeConfig "';`n" jsxContent "`nprocessSingleRow(configStr);"
    
    try {
        result := aiApp.DoJavaScript(scriptToRun)
        if (result != "OK") {
            failList .= pdfName " - " result "`n"
        } else {
            successCount++
        }
    } catch as e {
        failList .= pdfName " - Lỗi: " e.Message "`n"
    }
}

; ✅ TẠO BÁO CÁO CUỐI CÙNG
report := ""
if (wasCancelled) {
    report := "Tiến trình đã được người dùng hủy."
    try FileDelete(cancelFlagFile) ; Dọn dẹp file flag
} else {
    report := "Xuất xong: " successCount " file thành công!`n"
    if (failList != "")
        report .= "`nCHI TIẾT LỖI:`n" failList
}
    
; Gửi báo cáo về luồng chính và tự thoát
SendMessage(WM_AI_DONE, 0, StrPtr(report),, "ahk_id " hwndMain)
ExitApp()