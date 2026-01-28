#NoEnv
#SingleInstance Force
SetBatchLines, -1          ; Maximum speed
SetTitleMatchMode, 2
SetKeyDelay, 20            ; Faster key presses
SetMouseDelay, 20          ; Faster mouse clicks
SetWinDelay, 50            ; Faster window operations

; ============================================
; DZH AdvisorXs Time & Sales Export Script
; Press ` (backtick) to export all T&S windows
; SPEED OPTIMIZED
; ============================================

; Configuration
global ExportFolder := "C:\Users\siycm1.CGSCIMB\Desktop\Data\TS"
global FirstExportDone := false
global ExcelFile := "OrderFlowX_V5.xlsm"
global MacroName := "QuickRankingUpdate"

; Hotkey to start export: Backtick (`)
`::
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
        Sleep, 50

        ; Right-click on T&S control
        ControlClick, TDrawGrid1, ahk_id %thisID%, , Right
        Sleep, 100

        ; Navigate menu and save
        if (SaveExport(stockName, isFirstExport)) {
            exportCount++
        }

        Sleep, 50
    }
   
    ; Mark first export as done
    FirstExportDone := true
   
    ; Calculate elapsed time
    elapsedTime := (A_TickCount - startTime) / 1000

    ; Auto-run Excel macro after export (no prompt)
    macroSuccess := RunExcelMacro()

    ; Only show dialog on error
    if (exportCount < numWindows || !macroSuccess) {
        errorMsg := "Export issues detected:`n`n"
        if (exportCount < numWindows)
            errorMsg .= "Exported: " . exportCount . " / " . numWindows . " stocks`n"
        if (!macroSuccess)
            errorMsg .= "Excel macro failed to run`n"
        errorMsg .= "`nTime: " . elapsedTime . " seconds"
        MsgBox, 16, Export Error, %errorMsg%
    }
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
    Send, {Down 7}
    Sleep, 50

    ; Move right to open submenu
    Send, {Right}
    Sleep, 80

    ; Click Save As
    Send, {Enter}
    Sleep, 150

    ; Wait for Save dialog
    WinWait, Save As, , 3
    if ErrorLevel {
        MsgBox, Save dialog did not appear for %stockName%
        Send, {Escape}{Escape}{Escape}
        return false
    }
    Sleep, 80

    if (isFirstExport) {
        ; First export: select CSV and type filename

        ; Tab to "Save as type" dropdown
        Send, {Tab}
        Sleep, 50

        ; Open dropdown and select CSV (2nd option)
        Send, {Down}{Down}
        Sleep, 50
        Send, {Enter}
        Sleep, 80

        ; Tab back to filename field
        Send, +{Tab}
        Sleep, 50

        ; Type filename
        fileName := stockName . ".csv"
        Send, ^a
        Sleep, 30
        SendRaw, %fileName%
        Sleep, 80
    }

    ; Press Enter to save
    Send, {Enter}
    Sleep, 150

    ; Handle overwrite confirmation
    IfWinExist, Confirm
    {
        Send, y
        Sleep, 80
    }

    return true
}

; Run Excel macro after export completes
RunExcelMacro() {
    global ExcelFile, MacroName, ExportFolder

    ; Find and activate Excel window
    WinActivate, %ExcelFile%
    Sleep, 300

    ; Check if window became active
    WinGetActiveTitle, activeTitle
    if !InStr(activeTitle, ExcelFile) {
        ; Try partial match - activate any Excel window
        WinActivate, ahk_exe EXCEL.EXE
        Sleep, 300
    }

    ; Open Macro dialog (Alt+F8)
    Send, !{F8}
    Sleep, 500

    ; Wait for Macro dialog to appear
    WinWait, Macro, , 5
    if ErrorLevel {
        return false
    }
    Sleep, 200

    ; Type macro name and run
    SendRaw, %MacroName%
    Sleep, 200
    Send, {Enter}
    Sleep, 1000  ; Wait for macro to start and show folder picker

    ; Wait for folder picker dialog - try multiple possible titles
    ; VBA FileDialog title is "Select Folder Containing CSV Files"
    found := false
    Loop, 10 {
        ; Check for various possible dialog titles
        IfWinExist, Select Folder
        {
            WinActivate, Select Folder
            found := true
            break
        }
        IfWinExist, Browse For Folder
        {
            WinActivate, Browse For Folder
            found := true
            break
        }
        IfWinExist, Browse
        {
            WinActivate, Browse
            found := true
            break
        }
        Sleep, 200
    }

    if (!found) {
        return false
    }
    Sleep, 400

    ; For Windows folder picker dialog, type path in the folder name field
    ; Try multiple methods to ensure path is entered correctly

    ; Method 1: Try Alt+D for address bar
    Send, !d
    Sleep, 150

    ; Clear and type the path
    Send, ^a
    Sleep, 80
    SendRaw, %ExportFolder%
    Sleep, 300

    ; Press Enter to navigate to the folder
    Send, {Enter}
    Sleep, 800  ; Wait for folder to load

    ; Now we need to press Enter or Tab+Enter to select the folder
    ; The "Select Folder" button should be focused or we can Tab to it
    Send, {Tab}
    Sleep, 100
    Send, {Enter}
    Sleep, 300

    ; If dialog still open, try direct Enter
    IfWinExist, Select Folder
    {
        Send, {Enter}
        Sleep, 200
    }
    IfWinExist, Browse
    {
        Send, {Enter}
        Sleep, 200
    }

    return true
}

; Emergency stop
Escape::
    MsgBox, Script stopped.
    Reload
    return