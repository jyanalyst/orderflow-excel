Attribute VB_Name = "Module3"
Sub GenerateRankingTable()
    ' Scan all sheets, extract last-row signals, compile ranking table
    ' Side-by-side layout: BULL (A-I) | BEAR (K-S)
    ' Appends to existing data with timestamp separator

    Dim ws As Worksheet
    Dim rankWs As Worksheet
    Dim sheetName As String
    Dim batchTimestamp As String

    ' Data collection arrays
    Dim bullData() As Variant
    Dim bearData() As Variant
    Dim bullCount As Long
    Dim bearCount As Long

    ' Position trackers
    Dim bullStartRow As Long
    Dim bearStartRow As Long
    Dim i As Long

    ' Signal variables
    Dim signalType As String
    Dim signalStatus As String
    Dim accelCount As Variant
    Dim entryPrice As Double
    Dim bullishFlag As String
    Dim bearishFlag As String
    Dim tickerCode As String
    Dim lastRow As Long

    Application.ScreenUpdating = False

    ' ---------------------------------------------------------
    ' CREATE BATCH TIMESTAMP
    ' ---------------------------------------------------------
    batchTimestamp = "Batch: " & Format(Now, "DD-MMM-YYYY HH:MM")

    ' ---------------------------------------------------------
    ' INITIALIZE DATA ARRAYS (max 100 stocks each)
    ' ---------------------------------------------------------
    ReDim bullData(1 To 100, 1 To 9)
    ReDim bearData(1 To 100, 1 To 9)
    bullCount = 0
    bearCount = 0
   
    ' ---------------------------------------------------------
    ' SCAN ALL SHEETS - COLLECT SIGNALS
    ' ---------------------------------------------------------
    For Each ws In ThisWorkbook.Worksheets
        sheetName = ws.Name
       
        ' Skip system sheets
        If sheetName <> "Data" And sheetName <> "OrderFlow" And sheetName <> "Ranking" _
           And LCase(sheetName) <> "data" And LCase(sheetName) <> "orderflow" Then
           
            ' Find last row with data
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
           
            If lastRow >= 2 Then
                ' Extract last row data
                signalType = ws.Cells(lastRow, 12).Value   ' Column L
                signalStatus = ws.Cells(lastRow, 17).Value ' Column Q
                accelCount = ws.Cells(lastRow, 18).Value   ' Column R
                entryPrice = ws.Cells(lastRow, 13).Value   ' Column M

                ' Look up ticker code from Watchlist sheet
                tickerCode = LookupTickerCode(sheetName)

                ' Check if ticker is in Bullish or Bearish sheets
                If CheckTickerInSheet(tickerCode, "Bullish") Then
                    bullishFlag = "Bullish"
                Else
                    bullishFlag = ""
                End If

                If CheckTickerInSheet(tickerCode, "Bearish") Then
                    bearishFlag = "Bearish"
                Else
                    bearishFlag = ""
                End If

                ' Collect BULL signals
                If signalType = "BULL" And signalStatus <> "" Then
                    bullCount = bullCount + 1
                    bullData(bullCount, 1) = ""  ' Rank (fill after sorting)
                    bullData(bullCount, 2) = sheetName
                    bullData(bullCount, 3) = tickerCode
                    bullData(bullCount, 4) = signalType
                    bullData(bullCount, 5) = signalStatus
                    bullData(bullCount, 6) = accelCount
                    bullData(bullCount, 7) = entryPrice
                    bullData(bullCount, 8) = bullishFlag
                    bullData(bullCount, 9) = bearishFlag
                End If

                ' Collect BEAR signals
                If signalType = "BEAR" And signalStatus <> "" Then
                    bearCount = bearCount + 1
                    bearData(bearCount, 1) = ""  ' Rank (fill after sorting)
                    bearData(bearCount, 2) = sheetName
                    bearData(bearCount, 3) = tickerCode
                    bearData(bearCount, 4) = signalType
                    bearData(bearCount, 5) = signalStatus
                    bearData(bearCount, 6) = accelCount
                    bearData(bearCount, 7) = entryPrice
                    bearData(bearCount, 8) = bullishFlag
                    bearData(bearCount, 9) = bearishFlag
                End If
            End If
        End If
    Next ws
   
    ' ---------------------------------------------------------
    ' GET OR CREATE RANKING SHEET
    ' ---------------------------------------------------------
    On Error Resume Next
    Set rankWs = ThisWorkbook.Sheets("Ranking")
    On Error GoTo 0
   
    If rankWs Is Nothing Then
        ' Create new Ranking sheet at position 3
        Set rankWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(2))
        rankWs.Name = "Ranking"
        bullStartRow = 1
        bearStartRow = 1
    Else
        ' Find next empty row for BULL (column A)
        bullStartRow = rankWs.Cells(rankWs.Rows.count, 1).End(xlUp).Row + 1
        If bullStartRow = 2 And rankWs.Cells(1, 1).Value = "" Then
            bullStartRow = 1  ' Sheet is empty
        End If
       
        ' Find next empty row for BEAR (column K)
        bearStartRow = rankWs.Cells(rankWs.Rows.count, 11).End(xlUp).Row + 1
        If bearStartRow = 2 And rankWs.Cells(1, 11).Value = "" Then
            bearStartRow = 1  ' Column K is empty
        End If
    End If
   
    ' Add blank row separator if appending (not first batch)
    If bullStartRow > 1 Then bullStartRow = bullStartRow + 1
    If bearStartRow > 1 Then bearStartRow = bearStartRow + 1
   
    ' ---------------------------------------------------------
    ' WRITE BULL SECTION (Columns A-I)
    ' ---------------------------------------------------------
    ' Sort bull data FIRST (needed for correct watchlist order)
    If bullCount > 0 Then
        Call SortSignalArray(bullData, bullCount)
    End If

    With rankWs
        ' Timestamp row
        .Cells(bullStartRow, 1).Value = batchTimestamp
        .Cells(bullStartRow, 1).Font.Bold = True
        .Cells(bullStartRow, 1).Font.Italic = True
        .Cells(bullStartRow, 1).Interior.Color = RGB(198, 224, 180)  ' Light green
        .Range(.Cells(bullStartRow, 1), .Cells(bullStartRow, 9)).Merge
        bullStartRow = bullStartRow + 1

        ' TradingView watchlist string row
        .Cells(bullStartRow, 2).Value = BuildTradingViewString(bullData, bullCount)
        .Cells(bullStartRow, 2).Font.Italic = True
        .Cells(bullStartRow, 2).Font.Color = RGB(80, 80, 80)
        bullStartRow = bullStartRow + 1

        ' Header row
        .Cells(bullStartRow, 1).Value = "Rank"
        .Cells(bullStartRow, 2).Value = "Stock"
        .Cells(bullStartRow, 3).Value = "Ticker"
        .Cells(bullStartRow, 4).Value = "Entry_Price"
        .Cells(bullStartRow, 5).Value = "Accel_Count"
        .Cells(bullStartRow, 6).Value = "Bullish"
        .Cells(bullStartRow, 7).Value = "Bearish"
        .Cells(bullStartRow, 8).Value = "Signal_Type"
        .Cells(bullStartRow, 9).Value = "Signal_Status"

        .Range(.Cells(bullStartRow, 1), .Cells(bullStartRow, 9)).Font.Bold = True
        .Range(.Cells(bullStartRow, 1), .Cells(bullStartRow, 9)).Interior.Color = RGB(169, 208, 142)
        .Range(.Cells(bullStartRow, 1), .Cells(bullStartRow, 9)).HorizontalAlignment = xlCenter
        bullStartRow = bullStartRow + 1
    End With

    ' Write bull data (already sorted above)
    If bullCount > 0 Then
        ' Write data and add rank numbers
        For i = 1 To bullCount
            bullData(i, 1) = i  ' Rank
            rankWs.Cells(bullStartRow + i - 1, 1).Value = bullData(i, 1)
            rankWs.Cells(bullStartRow + i - 1, 2).Value = bullData(i, 2)
            rankWs.Cells(bullStartRow + i - 1, 3).Value = bullData(i, 3)
            rankWs.Cells(bullStartRow + i - 1, 4).Value = bullData(i, 7)
            rankWs.Cells(bullStartRow + i - 1, 5).Value = bullData(i, 6)
            rankWs.Cells(bullStartRow + i - 1, 6).Value = bullData(i, 8)
            rankWs.Cells(bullStartRow + i - 1, 7).Value = bullData(i, 9)
            rankWs.Cells(bullStartRow + i - 1, 8).Value = bullData(i, 4)
            rankWs.Cells(bullStartRow + i - 1, 9).Value = bullData(i, 5)

            ' Highlight based on status
            Call HighlightSignalRow(rankWs, bullStartRow + i - 1, 1, 9, _
                                    CStr(bullData(i, 5)), i, "BULL")
        Next i
    End If
   
    ' ---------------------------------------------------------
    ' WRITE BEAR SECTION (Columns K-S)
    ' ---------------------------------------------------------
    ' Sort bear data FIRST (needed for correct watchlist order)
    If bearCount > 0 Then
        Call SortSignalArray(bearData, bearCount)
    End If

    With rankWs
        ' Timestamp row
        .Cells(bearStartRow, 11).Value = batchTimestamp
        .Cells(bearStartRow, 11).Font.Bold = True
        .Cells(bearStartRow, 11).Font.Italic = True
        .Cells(bearStartRow, 11).Interior.Color = RGB(244, 204, 204)  ' Light red
        .Range(.Cells(bearStartRow, 11), .Cells(bearStartRow, 19)).Merge
        bearStartRow = bearStartRow + 1

        ' TradingView watchlist string row
        .Cells(bearStartRow, 12).Value = BuildTradingViewString(bearData, bearCount)
        .Cells(bearStartRow, 12).Font.Italic = True
        .Cells(bearStartRow, 12).Font.Color = RGB(80, 80, 80)
        bearStartRow = bearStartRow + 1

        ' Header row
        .Cells(bearStartRow, 11).Value = "Rank"
        .Cells(bearStartRow, 12).Value = "Stock"
        .Cells(bearStartRow, 13).Value = "Ticker"
        .Cells(bearStartRow, 14).Value = "Entry_Price"
        .Cells(bearStartRow, 15).Value = "Accel_Count"
        .Cells(bearStartRow, 16).Value = "Bearish"
        .Cells(bearStartRow, 17).Value = "Bullish"
        .Cells(bearStartRow, 18).Value = "Signal_Type"
        .Cells(bearStartRow, 19).Value = "Signal_Status"

        .Range(.Cells(bearStartRow, 11), .Cells(bearStartRow, 19)).Font.Bold = True
        .Range(.Cells(bearStartRow, 11), .Cells(bearStartRow, 19)).Interior.Color = RGB(230, 145, 145)
        .Range(.Cells(bearStartRow, 11), .Cells(bearStartRow, 19)).HorizontalAlignment = xlCenter
        bearStartRow = bearStartRow + 1
    End With

    ' Write bear data (already sorted above)
    If bearCount > 0 Then
        ' Write data and add rank numbers
        For i = 1 To bearCount
            bearData(i, 1) = i  ' Rank
            rankWs.Cells(bearStartRow + i - 1, 11).Value = bearData(i, 1)
            rankWs.Cells(bearStartRow + i - 1, 12).Value = bearData(i, 2)
            rankWs.Cells(bearStartRow + i - 1, 13).Value = bearData(i, 3)
            rankWs.Cells(bearStartRow + i - 1, 14).Value = bearData(i, 7)
            rankWs.Cells(bearStartRow + i - 1, 15).Value = bearData(i, 6)
            rankWs.Cells(bearStartRow + i - 1, 16).Value = bearData(i, 9)
            rankWs.Cells(bearStartRow + i - 1, 17).Value = bearData(i, 8)
            rankWs.Cells(bearStartRow + i - 1, 18).Value = bearData(i, 4)
            rankWs.Cells(bearStartRow + i - 1, 19).Value = bearData(i, 5)

            ' Highlight based on status
            Call HighlightSignalRow(rankWs, bearStartRow + i - 1, 11, 19, _
                                    CStr(bearData(i, 5)), i, "BEAR")
        Next i
    End If
   
    ' ---------------------------------------------------------
    ' AUTO-FIT AND POSITION
    ' ---------------------------------------------------------
    ' AutoFit columns except B and L (watchlist string columns)
    ' Watchlist strings overflow naturally - user copies from formula bar
    rankWs.Columns("A").AutoFit
    rankWs.Columns("C:I").AutoFit
    rankWs.Columns("K").AutoFit
    rankWs.Columns("M:S").AutoFit
    rankWs.Columns("J").ColumnWidth = 3  ' Spacer column
   
    ' Move Ranking sheet to position 3 (after Data, OrderFlow)
    On Error Resume Next
    rankWs.Move Before:=ThisWorkbook.Sheets(3)
    On Error GoTo 0
   
    ' Activate ranking sheet and scroll to latest batch
    rankWs.Activate
    rankWs.Cells(rankWs.Cells(rankWs.Rows.count, 1).End(xlUp).Row, 1).Select
   
    Application.ScreenUpdating = True
   
    ' Show summary
    Dim activeBull As Long, activeBear As Long
    activeBull = 0: activeBear = 0
   
    For i = 1 To bullCount
        If InStr(bullData(i, 4), "Active") > 0 Then activeBull = activeBull + 1
    Next i
    For i = 1 To bearCount
        If InStr(bearData(i, 4), "Active") > 0 Then activeBear = activeBear + 1
    Next i
   
    MsgBox "RANKING TABLE UPDATED" & vbCrLf & vbCrLf & _
           "Batch: " & Format(Now, "DD-MMM-YYYY HH:MM") & vbCrLf & vbCrLf & _
           "Active BULL signals: " & activeBull & vbCrLf & _
           "Active BEAR signals: " & activeBear & vbCrLf & vbCrLf & _
           "Top 3 in each section highlighted in bold", vbInformation, "Ranking Complete"
End Sub

Sub SortSignalArray(ByRef dataArr() As Variant, ByVal dataCount As Long)
    ' Sort signal array: Active first, then by Accel_Count descending
    ' Simple bubble sort - fine for <100 items
   
    Dim i As Long, j As Long, k As Long
    Dim temp As Variant
    Dim status1 As String, status2 As String
    Dim accel1 As Variant, accel2 As Variant
    Dim swap As Boolean
   
    For i = 1 To dataCount - 1
        For j = i + 1 To dataCount
            swap = False
            status1 = CStr(dataArr(i, 4))
            status2 = CStr(dataArr(j, 4))
            accel1 = dataArr(i, 5)
            accel2 = dataArr(j, 5)
           
            ' Priority: Active > Success > Failed
            If InStr(status2, "Active") > 0 And InStr(status1, "Active") = 0 Then
                swap = True
            ElseIf InStr(status1, "Active") > 0 And InStr(status2, "Active") > 0 Then
                ' Both Active - sort by Accel_Count descending
                If val(Nz(accel2, 0)) > val(Nz(accel1, 0)) Then swap = True
            ElseIf InStr(status1, "Active") = 0 And InStr(status2, "Active") = 0 Then
                ' Neither Active - sort by Accel_Count descending
                If val(Nz(accel2, 0)) > val(Nz(accel1, 0)) Then swap = True
            End If
           
            If swap Then
                ' Swap rows
                For k = 1 To 8
                    temp = dataArr(i, k)
                    dataArr(i, k) = dataArr(j, k)
                    dataArr(j, k) = temp
                Next k
            End If
        Next j
    Next i
End Sub

Function Nz(val As Variant, defaultVal As Variant) As Variant
    ' Null-to-zero helper
    If IsEmpty(val) Or IsNull(val) Or val = "" Then
        Nz = defaultVal
    Else
        Nz = val
    End If
End Function

Function BuildTradingViewString(signalArr() As Variant, signalCount As Long) As String
    '=========================================
    ' Build TradingView watchlist string from signal array
    ' Format: SGX:TICKER1,SGX:TICKER2,SGX:TICKER3,...
    ' Array must already be sorted (highest Accel_Count first)
    ' Looks up stock names in Watchlist sheet (Column C) and returns tickers (Column D)
    '=========================================
    
    Dim result As String
    Dim i As Long
    Dim stockName As String
    Dim ticker As String
    Dim watchlistWs As Worksheet
    Dim lastRow As Long
    Dim j As Long
    Dim found As Boolean
    
    ' Get Watchlist sheet
    On Error Resume Next
    Set watchlistWs = ThisWorkbook.Sheets("Watchlist")
    On Error GoTo 0
    
    If watchlistWs Is Nothing Then
        ' No Watchlist sheet - use stock names as tickers
        result = ""
        For i = 1 To signalCount
            stockName = CStr(signalArr(i, 2))
            If stockName <> "" Then
                If result = "" Then
                    result = "SGX:" & stockName
                Else
                    result = result & ",SGX:" & stockName
                End If
            End If
        Next i
        BuildTradingViewString = result
        Exit Function
    End If
    
    ' Find last row in Watchlist
    lastRow = watchlistWs.Cells(watchlistWs.Rows.Count, 3).End(xlUp).Row
    
    result = ""
    
    For i = 1 To signalCount
        stockName = CStr(signalArr(i, 2))  ' Column 2 is stock name
        ticker = ""
        found = False
        
        If stockName <> "" Then
            ' Look up stock name in Watchlist Column C
            For j = 2 To lastRow  ' Start from row 2 (skip header)
                If UCase(Trim(watchlistWs.Cells(j, 3).Value)) = UCase(Trim(stockName)) Then
                    ticker = Trim(watchlistWs.Cells(j, 4).Value)  ' Column D
                    found = True
                    Exit For
                End If
            Next j
            
            ' If not found, use stock name as ticker
            If Not found Or ticker = "" Then
                ticker = stockName
            End If
            
            ' Add to result string
            If result = "" Then
                result = "SGX:" & ticker
            Else
                result = result & ",SGX:" & ticker
            End If
        End If
    Next i
    
    BuildTradingViewString = result
End Function

Sub HighlightSignalRow(ws As Worksheet, rowNum As Long, startCol As Long, _
                                endCol As Long, status As String, rank As Long, signalType As String)
    ' Apply conditional highlighting to a signal row

    Dim bullishCol As Long
    Dim bearishCol As Long

    ' Determine Bullish/Bearish column positions based on section
    If startCol = 1 Then
        ' BULL section (A-I): Bullish in col 6, Bearish in col 7
        bullishCol = 6
        bearishCol = 7
    Else
        ' BEAR section (K-S): Bearish in col 16, Bullish in col 17
        bullishCol = 17
        bearishCol = 16
    End If

    With ws.Range(ws.Cells(rowNum, startCol), ws.Cells(rowNum, endCol))
        If InStr(status, "Active") > 0 Then
            .Interior.Color = RGB(220, 230, 241)  ' Light blue
            If rank <= 3 Then
                .Font.Bold = True
                If signalType = "BULL" Then
                    .Font.Color = RGB(0, 0, 139)    ' Dark blue
                Else
                    .Font.Color = RGB(139, 0, 0)   ' Dark red
                End If
            End If
        ElseIf InStr(status, "Success") > 0 Then
            .Interior.Color = RGB(226, 239, 218)  ' Light green
        ElseIf InStr(status, "Failed") > 0 Then
            .Interior.Color = RGB(248, 215, 215)  ' Light red
        End If
    End With

    ' Apply specific highlighting to Bullish/Bearish columns
    If ws.Cells(rowNum, bullishCol).Value = "Bullish" Then
        ws.Cells(rowNum, bullishCol).Interior.Color = RGB(198, 224, 180)  ' Light green (matches BULL timestamp)
    End If

    If ws.Cells(rowNum, bearishCol).Value = "Bearish" Then
        ws.Cells(rowNum, bearishCol).Interior.Color = RGB(244, 204, 204)  ' Light red (matches BEAR timestamp)
    End If
End Sub

Sub DeleteBatchProcessedSheets()
    '
    ' Delete all batch-processed ticker sheets
    ' Keeps only: Data, OrderFlow, Ranking
    '
    Dim ws As Worksheet
    Dim sheetName As String
    Dim coreSheets As String
    Dim deleteCount As Integer
   
    ' Core sheets to ALWAYS keep
    coreSheets = ",Data,OrderFlow,Ranking,"
   
    Application.DisplayAlerts = False
    deleteCount = 0
   
    ' Loop backwards through all sheets (safer when deleting)
    For i = ActiveWorkbook.Worksheets.count To 1 Step -1
        Set ws = ActiveWorkbook.Worksheets(i)
        sheetName = ws.Name
       
        ' If sheet is NOT in core list, delete it
        If InStr(1, coreSheets, "," & sheetName & ",", vbTextCompare) = 0 Then
            ws.Delete
            deleteCount = deleteCount + 1
        End If
    Next i
   
    Application.DisplayAlerts = True
   
    MsgBox "Deleted " & deleteCount & " ticker sheets.", vbInformation, "Cleanup Complete"

End Sub

Function LookupTickerCode(stockName As String) As String
    ' Look up ticker code from Watchlist sheet
    ' Returns ticker code if found, otherwise returns stock name

    Dim watchlistWs As Worksheet
    Dim lastRow As Long
    Dim j As Long
    Dim ticker As String

    ' Default to stock name if lookup fails
    LookupTickerCode = stockName

    ' Get Watchlist sheet
    On Error Resume Next
    Set watchlistWs = ThisWorkbook.Sheets("Watchlist")
    On Error GoTo 0

    If watchlistWs Is Nothing Then
        Exit Function
    End If

    ' Find last row in Watchlist
    lastRow = watchlistWs.Cells(watchlistWs.Rows.Count, 3).End(xlUp).Row

    ' Look up stock name in Watchlist Column C
    For j = 2 To lastRow  ' Start from row 2 (skip header)
        If UCase(Trim(watchlistWs.Cells(j, 3).Value)) = UCase(Trim(stockName)) Then
            ticker = Trim(watchlistWs.Cells(j, 4).Value)  ' Column D
            If ticker <> "" Then
                LookupTickerCode = ticker
                Exit Function
            End If
        End If
    Next j
End Function
