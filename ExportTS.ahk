#NoEnv
#SingleInstance Force
SetTitleMatchMode, 2
SetKeyDelay, 50
SetMouseDelay, 50

; ============================================
; DZH AdvisorXs Time & Sales Export Script
; Press Ctrl+Shift+E to export all T&S windows
; Optimized for speed
; ============================================

; Configuration
global ExportFolder := "C:\Users\siycm1.CGSCIMB\Desktop\Data\TS"
global FirstExportDone := false

; Hotkey to start export: Ctrl+Shift+E
^+e::
    ; First, find all standalone T&S windows
    windowIDs := []
    windowTitles := []
   
    WinGet, ids, List, ahk_class TaqFloatingForm ahk_exe AdvisorXs.exe
   
    Loop, %ids% {
        thisID := ids%A_Index%
        WinGetTitle, thisTitle, ahk_id %thisID%
       
        ; Skip if title contains "QTrack" (docked windows)
        if InStr(thisTitle, "QTrack")
            continue
       
        ; Skip if title is empty
        if (thisTitle = "")
            continue
       
        windowIDs.Push(thisID)
        windowTitles.Push(thisTitle)
    }
   
    numWindows := windowIDs.Length()
   
    if (numWindows = 0) {
        MsgBox, No standalone Time & Sales windows found!`n`nMake sure you have T&S windows open (not docked).
        return
    }
   
    ; Ask if this is first export of the day
    if (!FirstExportDone) {
        MsgBox, 4, First Export of Day?, Is this the FIRST export of the day?`n`nYes = Select CSV format and type filenames`nNo = Just press Enter (filenames already set)
        IfMsgBox Yes
            isFirstExport := true
        else
            isFirstExport := false
    } else {
        isFirstExport := false
    }
   
    MsgBox, 4, Export Time & Sales, Found %numWindows% standalone T&S windows.`n`nExport folder:`n%ExportFolder%`n`nFirst export: %isFirstExport%`n`nClick Yes to start.
    IfMsgBox No
        return
   
    exportCount := 0
    startTime := A_TickCount
   
    Loop, %numWindows% {
        idx := A_Index
        thisID := windowIDs[idx]
        thisTitle := windowTitles[idx]
       
        ; Extract stock name from title
        stockName := ExtractStockName(thisTitle)
       
        if (stockName = "UNKNOWN") {
            MsgBox, Could not extract stock name from:`n%thisTitle%`n`nSkipping...
            continue
        }
       
        ; Activate this window
        WinActivate, ahk_id %thisID%
        WinWaitActive, ahk_id %thisID%, , 2
        Sleep, 150
       
        ; Right-click on T&S control
        ControlClick, TDrawGrid1, ahk_id %thisID%, , Right
        Sleep, 250
       
        ; Navigate menu and save
        if (SaveExport(stockName, isFirstExport)) {
            exportCount++
        }
       
        Sleep, 150
    }
   
    ; Mark first export as done
    FirstExportDone := true
   
    ; Calculate elapsed time
    elapsedTime := (A_TickCount - startTime) / 1000
   
    MsgBox, Export complete!`n`nExported: %exportCount% / %numWindows% stocks`nTime: %elapsedTime% seconds`nFolder: %ExportFolder%
    return

; Extract stock name from window title
; "SINGTEL.SG [4.46 +0.07]" → "SINGTEL"
ExtractStockName(title) {
    pos := InStr(title, ".SG")
    if (pos > 0) {
        name := SubStr(title, 1, pos - 1)
        name := Trim(name)
        return name
    }
    return "UNKNOWN"
}

; Navigate export menu and save
SaveExport(stockName, isFirstExport) {
    ; Navigate to Export menu item (7th item down)
    Send, {Down}{Down}{Down}{Down}{Down}{Down}{Down}
    Sleep, 100
   
    ; Move right to open submenu
    Send, {Right}
    Sleep, 150
   
    ; Click Save As
    Send, {Enter}
    Sleep, 300
   
    ; Wait for Save dialog
    WinWait, Save As, , 3
    if ErrorLevel {
        MsgBox, Save dialog did not appear for %stockName%
        Send, {Escape}{Escape}{Escape}
        return false
    }
    Sleep, 150
   
    if (isFirstExport) {
        ; First export: select CSV and type filename
       
        ; Tab to "Save as type" dropdown
        Send, {Tab}
        Sleep, 100
       
        ; Open dropdown and select CSV (2nd option)
        Send, {Down}{Down}
        Sleep, 100
        Send, {Enter}
        Sleep, 150
       
        ; Tab back to filename field
        Send, {Shift down}{Tab}{Shift up}
        Sleep, 100
       
        ; Type filename
        fileName := stockName . ".csv"
        Send, ^a
        Sleep, 50
        SendRaw, %fileName%
        Sleep, 150
    }
   
    ; Press Enter to save
    Send, {Enter}
    Sleep, 300
   
    ; Handle overwrite confirmation
    IfWinExist, Confirm
    {
        Send, y
        Sleep, 150
    }
   
    return true
}

; Emergency stop
Escape::
    MsgBox, Script stopped.
    Reload
    return