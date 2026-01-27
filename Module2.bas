Attribute VB_Name = "Module2"
' ===========================================
' COMPLETE MODULE2.BAS - ALL FUNCTIONS RESTORED + FIXES APPLIED
' ===========================================
' This version includes ALL original functions plus critical fixes:
' 1. Removed cumDelta(0) = 0 line that caused "Subscript out of range"
' 2. Added folder picker dialog (replaces hardcoded DATA_FOLDER)
' 3. Added delimiter auto-detection for tab/comma CSV files
' 4. Enhanced bounds checking in array operations
' 5. PERFORMANCE OPTIMIZATIONS: Binary search, bulk arrays, cached lookups

' ===========================================
' CONFIGURATION - DEFAULT DATA FOLDER
' ===========================================
' Change this to your preferred data folder location
' If the folder exists, QuickRankingUpdate will use it automatically
' If not found, it will show the folder picker dialog
Public Const DEFAULT_DATA_FOLDER As String = "C:\Users\siycm1.CGSCIMB\Desktop\Data\TS\"

Function ParseDateTime(cellValue As Variant) As Date
    ' Universal date/time parser for Order Flow System
    ' Handles both text format (office) and Excel date format (home)

    On Error GoTo ErrorHandler

    Dim result As Date
    Dim strValue As String

    ' Check if already a proper date
    If IsDate(cellValue) Then
        result = CDate(cellValue)
        ParseDateTime = result
        Exit Function
    End If

    ' Handle text format: "DD-MM-YYYY HH:MM:SS" or "DD-MM-YYYY HH:MM"
    strValue = Trim(CStr(cellValue))

    ' Validate minimum length
    If Len(strValue) < 16 Then
        GoTo ErrorHandler
    End If

    ' Parse components using MID
    Dim dayPart As Integer
    Dim monthPart As Integer
    Dim yearPart As Integer
    Dim hourPart As Integer
    Dim minutePart As Integer
    Dim secondPart As Integer

    ' Extract date parts (DD-MM-YYYY)
    dayPart = CInt(Mid(strValue, 1, 2))
    monthPart = CInt(Mid(strValue, 4, 2))
    yearPart = CInt(Mid(strValue, 7, 4))

    ' Extract time parts (HH:MM or HH:MM:SS)
    hourPart = CInt(Mid(strValue, 12, 2))
    minutePart = CInt(Mid(strValue, 15, 2))

    ' Check if seconds exist
    If Len(strValue) >= 19 Then
        secondPart = CInt(Mid(strValue, 18, 2))
    Else
        secondPart = 0
    End If

    ' Construct date/time
    result = DateSerial(yearPart, monthPart, dayPart) + TimeSerial(hourPart, minutePart, secondPart)
    ParseDateTime = result
    Exit Function

ErrorHandler:
    ' Return a recognizable error value (e.g., #N/A equivalent)
    ParseDateTime = 0
End Function

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

Function CheckTickerInSheet(ticker As String, sheetName As String) As Boolean
    ' Check if ticker exists in comma-separated string in target sheet (cell A1)
    ' Returns True if found (case-insensitive)
    ' NOTE: For batch operations, use cached version LoadTickerLists() + CheckTickerInCache()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim tickerString As String
    Dim tickerArray() As String
    Dim i As Integer
    Dim trimmedTicker As String

    ' Default return value
    CheckTickerInSheet = False

    ' Try to get the sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler

    ' If sheet doesn't exist, return False
    If ws Is Nothing Then
        Exit Function
    End If

    ' Get ticker string from A1
    tickerString = Trim(ws.Range("A1").Value)

    ' If empty, return False
    If tickerString = "" Then
        Exit Function
    End If

    ' Split by comma
    tickerArray = Split(tickerString, ",")

    ' Search for ticker (case-insensitive, trimmed)
    For i = LBound(tickerArray) To UBound(tickerArray)
        trimmedTicker = Trim(tickerArray(i))
        If UCase(trimmedTicker) = UCase(Trim(ticker)) Then
            CheckTickerInSheet = True
            Exit Function
        End If
    Next i

    Exit Function

ErrorHandler:
    CheckTickerInSheet = False
End Function

' ===========================================
' LOOKUP CACHING SYSTEM
' Reduces O(n) lookups to O(1) by pre-loading into Dictionaries
' ===========================================

' Module-level cache variables (persist within batch operation)
Private mTickerCache As Object      ' stockName -> ticker
Private mBullishCache As Object     ' ticker -> True
Private mBearishCache As Object     ' ticker -> True
Private mCacheInitialized As Boolean

Sub InitializeLookupCaches()
    ' Initialize all lookup caches for batch operations
    ' Call once at start of GenerateRankingTable or QuickRankingUpdate

    Dim watchlistWs As Worksheet
    Dim bullishWs As Worksheet
    Dim bearishWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim stockName As String
    Dim ticker As String
    Dim tickerString As String
    Dim tickerArray() As String

    ' Create Dictionary objects (late-binding - no reference required)
    Set mTickerCache = CreateObject("Scripting.Dictionary")
    Set mBullishCache = CreateObject("Scripting.Dictionary")
    Set mBearishCache = CreateObject("Scripting.Dictionary")

    mTickerCache.CompareMode = vbTextCompare  ' Case-insensitive
    mBullishCache.CompareMode = vbTextCompare
    mBearishCache.CompareMode = vbTextCompare

    ' ---------------------------------------------------------
    ' Load Watchlist (stockName -> ticker) into cache
    ' ---------------------------------------------------------
    On Error Resume Next
    Set watchlistWs = ThisWorkbook.Sheets("Watchlist")
    On Error GoTo 0

    If Not watchlistWs Is Nothing Then
        lastRow = watchlistWs.Cells(watchlistWs.Rows.count, 3).End(xlUp).Row
        For i = 2 To lastRow
            stockName = Trim(CStr(watchlistWs.Cells(i, 3).Value))
            ticker = Trim(CStr(watchlistWs.Cells(i, 4).Value))
            If stockName <> "" And ticker <> "" Then
                If Not mTickerCache.Exists(stockName) Then
                    mTickerCache.Add stockName, ticker
                End If
            End If
        Next i
    End If

    ' ---------------------------------------------------------
    ' Load Bullish tickers into cache
    ' ---------------------------------------------------------
    On Error Resume Next
    Set bullishWs = ThisWorkbook.Sheets("Bullish")
    On Error GoTo 0

    If Not bullishWs Is Nothing Then
        tickerString = Trim(CStr(bullishWs.Range("A1").Value))
        If tickerString <> "" Then
            tickerArray = Split(tickerString, ",")
            For i = LBound(tickerArray) To UBound(tickerArray)
                ticker = Trim(tickerArray(i))
                If ticker <> "" Then
                    If Not mBullishCache.Exists(ticker) Then
                        mBullishCache.Add ticker, True
                    End If
                End If
            Next i
        End If
    End If

    ' ---------------------------------------------------------
    ' Load Bearish tickers into cache
    ' ---------------------------------------------------------
    On Error Resume Next
    Set bearishWs = ThisWorkbook.Sheets("Bearish")
    On Error GoTo 0

    If Not bearishWs Is Nothing Then
        tickerString = Trim(CStr(bearishWs.Range("A1").Value))
        If tickerString <> "" Then
            tickerArray = Split(tickerString, ",")
            For i = LBound(tickerArray) To UBound(tickerArray)
                ticker = Trim(tickerArray(i))
                If ticker <> "" Then
                    If Not mBearishCache.Exists(ticker) Then
                        mBearishCache.Add ticker, True
                    End If
                End If
            Next i
        End If
    End If

    mCacheInitialized = True
End Sub

Sub ClearLookupCaches()
    ' Clear caches to free memory
    Set mTickerCache = Nothing
    Set mBullishCache = Nothing
    Set mBearishCache = Nothing
    mCacheInitialized = False
End Sub

Function LookupTickerCodeCached(stockName As String) As String
    ' O(1) lookup using cached Dictionary
    ' Falls back to stockName if not found

    If Not mCacheInitialized Then
        ' Fallback to slow version if cache not initialized
        LookupTickerCodeCached = LookupTickerCode(stockName)
        Exit Function
    End If

    If mTickerCache.Exists(stockName) Then
        LookupTickerCodeCached = mTickerCache(stockName)
    Else
        LookupTickerCodeCached = stockName  ' Fallback to stock name
    End If
End Function

Function IsBullishCached(ticker As String) As Boolean
    ' O(1) lookup using cached Dictionary

    If Not mCacheInitialized Then
        ' Fallback to slow version if cache not initialized
        IsBullishCached = CheckTickerInSheet(ticker, "Bullish")
        Exit Function
    End If

    IsBullishCached = mBearishCache.Exists(ticker)
    ' Correction: Check bullish cache, not bearish
    IsBullishCached = mBullishCache.Exists(ticker)
End Function

Function IsBearishCached(ticker As String) As Boolean
    ' O(1) lookup using cached Dictionary

    If Not mCacheInitialized Then
        ' Fallback to slow version if cache not initialized
        IsBearishCached = CheckTickerInSheet(ticker, "Bearish")
        Exit Function
    End If

    IsBearishCached = mBearishCache.Exists(ticker)
End Function

' -------------------------------------------------------------------
' CALCULATION ENGINE - Order Flow Velocity System
' -------------------------------------------------------------------

Function CalcSignedVol(aggressor As String, volume As Double) As Double
    ' Column E: Signed_Vol
    If aggressor = "s" Then
        CalcSignedVol = volume
    ElseIf aggressor = "b" Then
        CalcSignedVol = -volume
    Else
        CalcSignedVol = 0
    End If
End Function

Function CalcElapsedSec(currentTime As Date, startTime As Date) As Double
    ' Column G: Elapsed_Sec
    CalcElapsedSec = (currentTime - startTime) * 86400
End Function

' ===========================================
' OPTIMIZED VELOCITY CALCULATION
' Uses binary search instead of linear backward walk
' Reduces O(n²) to O(n log n)
' ===========================================

Function BinarySearchTime(elapsedArr() As Double, targetTime As Double, maxIndex As Long) As Long
    ' Binary search to find largest index where elapsedArr(index) <= targetTime
    ' Returns 0 if no such index exists

    Dim lo As Long, hi As Long, mid As Long

    If maxIndex < 1 Then
        BinarySearchTime = 0
        Exit Function
    End If

    ' If first element is already > target, no valid index
    If elapsedArr(1) > targetTime Then
        BinarySearchTime = 0
        Exit Function
    End If

    lo = 1
    hi = maxIndex

    Do While lo < hi
        mid = (lo + hi + 1) \ 2
        If elapsedArr(mid) <= targetTime Then
            lo = mid
        Else
            hi = mid - 1
        End If
    Loop

    BinarySearchTime = lo
End Function

Function CalcVelocityFromArrays(currentRow As Long, elapsedArr() As Double, cumDeltaArr() As Double) As Variant
    ' Column H: Velocity (5-minute rolling window)
    ' OPTIMIZED: Uses pre-loaded arrays + binary search
    ' Formula: (Current Cum_Delta - Cum_Delta at t-300s) / 300

    Dim elapsedSec As Double
    Dim targetTime As Double
    Dim currentCumDelta As Double
    Dim pastCumDelta As Double
    Dim foundIndex As Long

    ' Get current values from arrays
    elapsedSec = elapsedArr(currentRow)
    currentCumDelta = cumDeltaArr(currentRow)

    ' Need at least 300 seconds of data
    If elapsedSec < 300 Then
        CalcVelocityFromArrays = Empty
        Exit Function
    End If

    ' Target time is 300 seconds ago
    targetTime = elapsedSec - 300

    ' Binary search for the index (O(log n) instead of O(n))
    foundIndex = BinarySearchTime(elapsedArr, targetTime, currentRow)

    ' If we found a valid index
    If foundIndex > 0 Then
        pastCumDelta = cumDeltaArr(foundIndex)
        If pastCumDelta = 0 Then
            CalcVelocityFromArrays = Empty
        Else
            CalcVelocityFromArrays = (currentCumDelta - pastCumDelta) / 300
        End If
    Else
        CalcVelocityFromArrays = Empty
    End If
End Function

' Legacy function for backward compatibility (deprecated - use CalcVelocityFromArrays)
Function CalcVelocity(currentRow As Long, ws As Worksheet) As Variant
    ' Column H: Velocity (5-minute rolling window)
    ' NOTE: This function is kept for compatibility but is slow
    ' New code should use CalcVelocityFromArrays with pre-loaded arrays

    Dim elapsedSec As Double
    Dim targetTime As Double
    Dim currentCumDelta As Double
    Dim pastCumDelta As Variant
    Dim i As Long

    ' Get current values
    elapsedSec = ws.Cells(currentRow, 7).Value ' Column G
    currentCumDelta = ws.Cells(currentRow, 6).Value ' Column F

    ' Need at least 300 seconds of data
    If elapsedSec < 300 Then
        CalcVelocity = ""
        Exit Function
    End If

    ' Target time is 300 seconds ago
    targetTime = elapsedSec - 300

    ' Find the Cum_Delta value at or before target time (XLOOKUP behavior with match mode 1)
    pastCumDelta = Empty
    For i = currentRow To 2 Step -1
        If ws.Cells(i, 7).Value <= targetTime Then
            pastCumDelta = ws.Cells(i, 6).Value
            Exit For
        End If
    Next i

    ' If we found a past value
    If Not IsEmpty(pastCumDelta) Then
        If pastCumDelta = 0 Then
            CalcVelocity = ""
        Else
            CalcVelocity = (currentCumDelta - pastCumDelta) / 300
        End If
    Else
        CalcVelocity = ""
    End If
End Function

Function CalcZeroCross(currentRow As Long, ws As Worksheet) As String
    ' Column I: Zero_Cross (regime change detection)
    ' BULL: Velocity crosses from negative to non-negative
    ' BEAR: Velocity crosses from non-negative to negative

    If currentRow < 3 Then
        CalcZeroCross = ""
        Exit Function
    End If

    Dim currVel As Variant
    Dim prevVel As Variant

    currVel = ws.Cells(currentRow, 8).Value ' Column H
    prevVel = ws.Cells(currentRow - 1, 8).Value ' Column H

    ' Both must have values
    If currVel = "" Or prevVel = "" Then
        CalcZeroCross = ""
        Exit Function
    End If

    ' Check for crossovers
    If prevVel < 0 And currVel >= 0 Then
        CalcZeroCross = "BULL"
    ElseIf prevVel >= 0 And currVel < 0 Then
        CalcZeroCross = "BEAR"
    Else
        CalcZeroCross = ""
    End If
End Function

Function CalcAccel(currentRow As Long, ws As Worksheet) As Variant
    ' Column J: Accel (second derivative)

    If currentRow < 3 Then
        CalcAccel = ""
        Exit Function
    End If

    Dim currVel As Variant
    Dim prevVel As Variant

    currVel = ws.Cells(currentRow, 8).Value
    prevVel = ws.Cells(currentRow - 1, 8).Value

    If currVel = "" Or prevVel = "" Then
        CalcAccel = ""
    Else
        CalcAccel = currVel - prevVel
    End If
End Function

Function CalcAccelSignal(currentRow As Long, ws As Worksheet) As String
    ' Column K: Accel_Signal (combined state)

    If currentRow < 3 Then
        CalcAccelSignal = ""
        Exit Function
    End If

    Dim vel As Variant
    Dim accel As Variant

    vel = ws.Cells(currentRow, 8).Value ' Column H
    accel = ws.Cells(currentRow, 10).Value ' Column J

    If vel = "" Or accel = "" Then
        CalcAccelSignal = ""
        Exit Function
    End If

    ' Determine state
    If vel > 0 And accel > 0 Then
        CalcAccelSignal = "BULL ACCEL"
    ElseIf vel > 0 And accel <= 0 Then
        CalcAccelSignal = "BULL DECEL"
    ElseIf vel <= 0 And accel < 0 Then
        CalcAccelSignal = "BEAR ACCEL"
    ElseIf vel <= 0 And accel >= 0 Then
        CalcAccelSignal = "BEAR DECEL"
    Else
        CalcAccelSignal = ""
    End If
End Function

Function GetTickSize(price As Double) As Double
    ' SGX tick size rules
    If price >= 1 Then
        GetTickSize = 0.01
    ElseIf price >= 0.2 Then
        GetTickSize = 0.005
    Else
        GetTickSize = 0.001
    End If
End Function

' -------------------------------------------------------------------
' SIGNAL TRACKING ENGINE (Columns L-R)
' -------------------------------------------------------------------

Sub CalculateSignalTracking(ws As Worksheet, lastRow As Long)
    ' This calculates columns L-R for the entire sheet
    ' Matches the exact Signal_Status logic from orderflow template

    Dim i As Long
    Dim zeroCross As String
    Dim currentSignalType As String
    Dim currentStatus As String
    Dim entryPrice As Double
    Dim entryTick As Double
    Dim accelCount As Long
    Dim price As Double
    Dim accelSignal As String

    ' Initialize tracking variables
    currentSignalType = ""
    currentStatus = ""
    entryPrice = 0
    entryTick = 0
    accelCount = 0

    ' Start from row 2 (first data row)
    For i = 2 To lastRow
        ' Get current row values
        zeroCross = ws.Cells(i, 9).Value ' Column I
        price = ws.Cells(i, 2).Value ' Column B
        accelSignal = ws.Cells(i, 11).Value ' Column K

        ' ---------------------------------------------------------
        ' COLUMN L: Signal_Type (carry forward or new signal)
        ' ---------------------------------------------------------
        If zeroCross = "BULL" Or zeroCross = "BEAR" Then
            ' New signal detected
            currentSignalType = zeroCross
            If zeroCross = "BULL" Then
                currentStatus = "Active Bullish"
            Else
                currentStatus = "Active Bearish"
            End If
            entryPrice = price
            entryTick = GetTickSize(entryPrice)
            accelCount = 0
        End If

        ws.Cells(i, 12).Value = currentSignalType ' Column L

        ' ---------------------------------------------------------
        ' COLUMN M: Entry_Price
        ' ---------------------------------------------------------
        If currentSignalType <> "" Then
            ws.Cells(i, 13).Value = entryPrice
        Else
            ws.Cells(i, 13).Value = ""
        End If

        ' ---------------------------------------------------------
        ' COLUMN N: Entry_Tick
        ' ---------------------------------------------------------
        If currentSignalType <> "" Then
            ws.Cells(i, 14).Value = entryTick
        Else
            ws.Cells(i, 14).Value = ""
        End If

        ' ---------------------------------------------------------
        ' COLUMN O: Success_Price (±1 tick)
        ' ---------------------------------------------------------
        If currentSignalType = "BULL" Then
            ws.Cells(i, 15).Value = entryPrice + entryTick
        ElseIf currentSignalType = "BEAR" Then
            ws.Cells(i, 15).Value = entryPrice - entryTick
        Else
            ws.Cells(i, 15).Value = ""
        End If

        ' ---------------------------------------------------------
        ' COLUMN P: Fail_Price (±2 ticks opposite direction)
        ' ---------------------------------------------------------
        If currentSignalType = "BULL" Then
            ws.Cells(i, 16).Value = entryPrice - (2 * entryTick)
        ElseIf currentSignalType = "BEAR" Then
            ws.Cells(i, 16).Value = entryPrice + (2 * entryTick)
        Else
            ws.Cells(i, 16).Value = ""
        End If

        ' ---------------------------------------------------------
        ' COLUMN Q: Signal_Status
        ' Matches Excel formula exactly:
        ' - Once SUCCESS or FAILED, status is locked in
        ' - Only check price while in ACTIVE state
        ' ---------------------------------------------------------
        If currentStatus <> "" Then
            ' Check if already completed
            If InStr(currentStatus, "Success") > 0 Or InStr(currentStatus, "Failed") > 0 Then
                ' Already completed - carry forward
                ws.Cells(i, 17).Value = currentStatus
            Else
                ' Still ACTIVE - check for success or failure
                If currentSignalType = "BULL" Then
                    If price >= ws.Cells(i, 15).Value Then
                        currentStatus = "Success Bullish"
                    ElseIf price <= ws.Cells(i, 16).Value Then
                        currentStatus = "Failed Bullish"
                    ' Else remains "Active Bullish"
                    End If
                ElseIf currentSignalType = "BEAR" Then
                    If price <= ws.Cells(i, 15).Value Then
                        currentStatus = "Success Bearish"
                    ElseIf price >= ws.Cells(i, 16).Value Then
                        currentStatus = "Failed Bearish"
                    ' Else remains "Active Bearish"
                    End If
                End If
                ws.Cells(i, 17).Value = currentStatus
            End If
        Else
            ws.Cells(i, 17).Value = ""
        End If

        ' ---------------------------------------------------------
        ' COLUMN R: Accel_Count (only increment while ACTIVE)
        ' ---------------------------------------------------------
        If InStr(currentStatus, "Active") > 0 And currentSignalType <> "" Then
            ' Check if accelSignal matches signal type
            If currentSignalType = "BULL" And accelSignal = "BULL ACCEL" Then
                accelCount = accelCount + 1
            ElseIf currentSignalType = "BEAR" And accelSignal = "BEAR ACCEL" Then
                accelCount = accelCount + 1
            End If
            ws.Cells(i, 18).Value = accelCount
        ElseIf currentStatus <> "" Then
            ' Keep final count after signal completes (Success or Failed)
            ws.Cells(i, 18).Value = accelCount
        Else
            ws.Cells(i, 18).Value = ""
        End If
    Next i
End Sub

' -------------------------------------------------------------------
' VOLUME SIGNIFICANCE (Column S)
' -------------------------------------------------------------------

Function CalcVolFlag(volume As Double, medianVol As Double) As String
    ' Column S: Vol_Flag
    If medianVol = 0 Then
        CalcVolFlag = ""
        Exit Function
    End If

    If volume > 20 * medianVol Then
        CalcVolFlag = "BLOCK"
    ElseIf volume > 10 * medianVol Then
        CalcVolFlag = "LARGE"
    ElseIf volume > 5 * medianVol Then
        CalcVolFlag = "NOTABLE"
    Else
        CalcVolFlag = ""
    End If
End Function

Function CalculateMedianVolume(ws As Worksheet, lastRow As Long) As Double
    ' LEGACY VERSION - Uses bubble sort (slow for large datasets)
    ' Kept for backward compatibility
    ' New code should use CalculateMedianVolumeOptimized

    Dim volumes() As Double
    Dim count As Long
    Dim i As Long

    ' Count non-zero volumes
    count = 0
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value > 0 Then
            count = count + 1
        End If
    Next i

    If count = 0 Then
        CalculateMedianVolume = 0
        Exit Function
    End If

    ' Fill array
    ReDim volumes(1 To count)
    Dim idx As Long
    idx = 1
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value > 0 Then
            volumes(idx) = ws.Cells(i, 3).Value
            idx = idx + 1
        End If
    Next i

    ' Sort array (simple bubble sort - good enough for this use)
    Dim temp As Double
    Dim j As Long
    For i = 1 To count - 1
        For j = i + 1 To count
            If volumes(i) > volumes(j) Then
                temp = volumes(i)
                volumes(i) = volumes(j)
                volumes(j) = temp
            End If
        Next j
    Next i

    ' Return median
    If count Mod 2 = 1 Then
        CalculateMedianVolume = volumes((count + 1) / 2)
    Else
        CalculateMedianVolume = (volumes(count / 2) + volumes(count / 2 + 1)) / 2
    End If
End Function

Function CalculateMedianVolumeOptimized(volArr() As Double, dataCount As Long) As Double
    ' ===========================================
    ' OPTIMIZED VERSION - Works with in-memory array
    ' Uses Excel's built-in Median function (O(n) instead of O(n²))
    ' ===========================================

    Dim volumes() As Double
    Dim count As Long
    Dim i As Long

    ' Count non-zero volumes
    count = 0
    For i = 1 To dataCount
        If volArr(i) > 0 Then
            count = count + 1
        End If
    Next i

    If count = 0 Then
        CalculateMedianVolumeOptimized = 0
        Exit Function
    End If

    ' Fill array with non-zero volumes
    ReDim volumes(1 To count)
    Dim idx As Long
    idx = 1
    For i = 1 To dataCount
        If volArr(i) > 0 Then
            volumes(idx) = volArr(i)
            idx = idx + 1
        End If
    Next i

    ' Use Excel's built-in Median (much faster than bubble sort)
    On Error Resume Next
    CalculateMedianVolumeOptimized = Application.WorksheetFunction.Median(volumes)
    If Err.Number <> 0 Then
        ' Fallback: return simple average
        Dim total As Double
        total = 0
        For i = 1 To count
            total = total + volumes(i)
        Next i
        CalculateMedianVolumeOptimized = total / count
        Err.Clear
    End If
    On Error GoTo 0
End Function

' -------------------------------------------------------------------
' MAIN PROCESSING ROUTINE
' -------------------------------------------------------------------

Sub ProcessSingleStock(ws As Worksheet, lastRow As Long)
    ' ===========================================
    ' OPTIMIZED VERSION - Bulk Array Processing
    ' Reads all data into memory, processes in arrays, writes back in bulk
    ' ~10-20x faster than cell-by-cell version
    ' ===========================================

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim i As Long
    Dim startTime As Date
    Dim dataCount As Long

    dataCount = lastRow - 1  ' Rows 2 to lastRow

    If dataCount < 1 Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' ---------------------------------------------------------
    ' STEP 1: BULK READ - Load source data into arrays (ONE read)
    ' ---------------------------------------------------------
    Dim sourceData As Variant
    sourceData = ws.Range("A2:D" & lastRow).Value  ' Columns A-D

    ' ---------------------------------------------------------
    ' STEP 2: INITIALIZE CALCULATION ARRAYS
    ' ---------------------------------------------------------
    Dim timeArr() As Date
    Dim priceArr() As Double
    Dim volArr() As Double
    Dim aggressorArr() As String

    Dim signedVolArr() As Double
    Dim cumDeltaArr() As Double
    Dim elapsedSecArr() As Double
    Dim velocityArr() As Variant
    Dim zeroCrossArr() As String
    Dim accelArr() As Variant
    Dim accelSignalArr() As String
    Dim volFlagArr() As String

    ' Signal tracking arrays (L-R)
    Dim signalTypeArr() As String
    Dim entryPriceArr() As Variant
    Dim entryTickArr() As Variant
    Dim successPriceArr() As Variant
    Dim failPriceArr() As Variant
    Dim signalStatusArr() As String
    Dim accelCountArr() As Variant

    ReDim timeArr(1 To dataCount)
    ReDim priceArr(1 To dataCount)
    ReDim volArr(1 To dataCount)
    ReDim aggressorArr(1 To dataCount)
    ReDim signedVolArr(1 To dataCount)
    ReDim cumDeltaArr(1 To dataCount)
    ReDim elapsedSecArr(1 To dataCount)
    ReDim velocityArr(1 To dataCount)
    ReDim zeroCrossArr(1 To dataCount)
    ReDim accelArr(1 To dataCount)
    ReDim accelSignalArr(1 To dataCount)
    ReDim volFlagArr(1 To dataCount)
    ReDim signalTypeArr(1 To dataCount)
    ReDim entryPriceArr(1 To dataCount)
    ReDim entryTickArr(1 To dataCount)
    ReDim successPriceArr(1 To dataCount)
    ReDim failPriceArr(1 To dataCount)
    ReDim signalStatusArr(1 To dataCount)
    ReDim accelCountArr(1 To dataCount)

    ' ---------------------------------------------------------
    ' STEP 3: PARSE SOURCE DATA INTO TYPED ARRAYS
    ' ---------------------------------------------------------
    For i = 1 To dataCount
        timeArr(i) = ParseDateTime(sourceData(i, 1))
        priceArr(i) = CDbl(sourceData(i, 2))
        volArr(i) = CDbl(sourceData(i, 3))
        aggressorArr(i) = CStr(sourceData(i, 4))
    Next i

    startTime = timeArr(1)

    ' ---------------------------------------------------------
    ' STEP 4: CALCULATE E, F, G (Signed_Vol, Cum_Delta, Elapsed_Sec)
    ' ---------------------------------------------------------
    Dim cumDelta As Double
    cumDelta = 0

    For i = 1 To dataCount
        ' Signed_Vol
        If aggressorArr(i) = "s" Then
            signedVolArr(i) = volArr(i)
        ElseIf aggressorArr(i) = "b" Then
            signedVolArr(i) = -volArr(i)
        Else
            signedVolArr(i) = 0
        End If

        ' Cum_Delta (running sum)
        cumDelta = cumDelta + signedVolArr(i)
        cumDeltaArr(i) = cumDelta

        ' Elapsed_Sec
        elapsedSecArr(i) = (timeArr(i) - startTime) * 86400
    Next i

    ' ---------------------------------------------------------
    ' STEP 5: CALCULATE H (Velocity) - OPTIMIZED with binary search
    ' ---------------------------------------------------------
    Dim currVel As Variant, prevVel As Variant
    Dim currAccel As Variant

    For i = 1 To dataCount
        velocityArr(i) = CalcVelocityFromArrays(i, elapsedSecArr, cumDeltaArr)
    Next i

    ' ---------------------------------------------------------
    ' STEP 6: CALCULATE I, J, K (Zero_Cross, Accel, Accel_Signal)
    ' ---------------------------------------------------------
    For i = 1 To dataCount
        ' Zero_Cross
        If i < 2 Or IsEmpty(velocityArr(i)) Then
            zeroCrossArr(i) = ""
        Else
            prevVel = velocityArr(i - 1)
            currVel = velocityArr(i)
            If IsEmpty(prevVel) Then
                zeroCrossArr(i) = ""
            ElseIf prevVel < 0 And currVel >= 0 Then
                zeroCrossArr(i) = "BULL"
            ElseIf prevVel >= 0 And currVel < 0 Then
                zeroCrossArr(i) = "BEAR"
            Else
                zeroCrossArr(i) = ""
            End If
        End If

        ' Accel
        If i < 2 Or IsEmpty(velocityArr(i)) Or IsEmpty(velocityArr(i - 1)) Then
            accelArr(i) = Empty
        Else
            accelArr(i) = velocityArr(i) - velocityArr(i - 1)
        End If

        ' Accel_Signal
        If IsEmpty(velocityArr(i)) Or IsEmpty(accelArr(i)) Then
            accelSignalArr(i) = ""
        ElseIf velocityArr(i) > 0 And accelArr(i) > 0 Then
            accelSignalArr(i) = "BULL ACCEL"
        ElseIf velocityArr(i) > 0 And accelArr(i) <= 0 Then
            accelSignalArr(i) = "BULL DECEL"
        ElseIf velocityArr(i) <= 0 And accelArr(i) < 0 Then
            accelSignalArr(i) = "BEAR ACCEL"
        ElseIf velocityArr(i) <= 0 And accelArr(i) >= 0 Then
            accelSignalArr(i) = "BEAR DECEL"
        Else
            accelSignalArr(i) = ""
        End If
    Next i

    ' ---------------------------------------------------------
    ' STEP 7: CALCULATE L-R (Signal Tracking) - in memory
    ' ---------------------------------------------------------
    Dim currentSignalType As String
    Dim currentStatus As String
    Dim entryPrice As Double
    Dim entryTick As Double
    Dim accelCount As Long
    Dim successPrice As Double
    Dim failPrice As Double

    currentSignalType = ""
    currentStatus = ""
    entryPrice = 0
    entryTick = 0
    accelCount = 0

    For i = 1 To dataCount
        ' Check for new signal
        If zeroCrossArr(i) = "BULL" Or zeroCrossArr(i) = "BEAR" Then
            currentSignalType = zeroCrossArr(i)
            If zeroCrossArr(i) = "BULL" Then
                currentStatus = "Active Bullish"
            Else
                currentStatus = "Active Bearish"
            End If
            entryPrice = priceArr(i)
            entryTick = GetTickSize(entryPrice)
            accelCount = 0
        End If

        signalTypeArr(i) = currentSignalType

        ' Entry_Price, Entry_Tick, Success_Price, Fail_Price
        If currentSignalType <> "" Then
            entryPriceArr(i) = entryPrice
            entryTickArr(i) = entryTick
            If currentSignalType = "BULL" Then
                successPriceArr(i) = entryPrice + entryTick
                failPriceArr(i) = entryPrice - (2 * entryTick)
            Else
                successPriceArr(i) = entryPrice - entryTick
                failPriceArr(i) = entryPrice + (2 * entryTick)
            End If
        Else
            entryPriceArr(i) = Empty
            entryTickArr(i) = Empty
            successPriceArr(i) = Empty
            failPriceArr(i) = Empty
        End If

        ' Signal_Status
        If currentStatus <> "" Then
            If InStr(currentStatus, "Success") > 0 Or InStr(currentStatus, "Failed") > 0 Then
                ' Already completed - carry forward
            Else
                ' Still Active - check for success or failure
                If currentSignalType = "BULL" Then
                    If priceArr(i) >= successPriceArr(i) Then
                        currentStatus = "Success Bullish"
                    ElseIf priceArr(i) <= failPriceArr(i) Then
                        currentStatus = "Failed Bullish"
                    End If
                ElseIf currentSignalType = "BEAR" Then
                    If priceArr(i) <= successPriceArr(i) Then
                        currentStatus = "Success Bearish"
                    ElseIf priceArr(i) >= failPriceArr(i) Then
                        currentStatus = "Failed Bearish"
                    End If
                End If
            End If
            signalStatusArr(i) = currentStatus
        Else
            signalStatusArr(i) = ""
        End If

        ' Accel_Count
        If InStr(currentStatus, "Active") > 0 And currentSignalType <> "" Then
            If currentSignalType = "BULL" And accelSignalArr(i) = "BULL ACCEL" Then
                accelCount = accelCount + 1
            ElseIf currentSignalType = "BEAR" And accelSignalArr(i) = "BEAR ACCEL" Then
                accelCount = accelCount + 1
            End If
            accelCountArr(i) = accelCount
        ElseIf currentStatus <> "" Then
            accelCountArr(i) = accelCount
        Else
            accelCountArr(i) = Empty
        End If
    Next i

    ' ---------------------------------------------------------
    ' STEP 8: CALCULATE S (Vol_Flag) - use built-in Median
    ' ---------------------------------------------------------
    Dim medianVol As Double
    medianVol = CalculateMedianVolumeOptimized(volArr, dataCount)

    For i = 1 To dataCount
        If medianVol > 0 Then
            If volArr(i) > 20 * medianVol Then
                volFlagArr(i) = "BLOCK"
            ElseIf volArr(i) > 10 * medianVol Then
                volFlagArr(i) = "LARGE"
            ElseIf volArr(i) > 5 * medianVol Then
                volFlagArr(i) = "NOTABLE"
            Else
                volFlagArr(i) = ""
            End If
        Else
            volFlagArr(i) = ""
        End If
    Next i

    ' ---------------------------------------------------------
    ' STEP 9: BULK WRITE - Write all results back (ONE write per section)
    ' ---------------------------------------------------------
    Dim outputEG() As Variant  ' Columns E-G
    Dim outputHK() As Variant  ' Columns H-K
    Dim outputLR() As Variant  ' Columns L-R
    Dim outputS() As Variant   ' Column S

    ReDim outputEG(1 To dataCount, 1 To 3)
    ReDim outputHK(1 To dataCount, 1 To 4)
    ReDim outputLR(1 To dataCount, 1 To 7)
    ReDim outputS(1 To dataCount, 1 To 1)

    For i = 1 To dataCount
        ' E-G: Signed_Vol, Cum_Delta, Elapsed_Sec
        outputEG(i, 1) = signedVolArr(i)
        outputEG(i, 2) = cumDeltaArr(i)
        outputEG(i, 3) = elapsedSecArr(i)

        ' H-K: Velocity, Zero_Cross, Accel, Accel_Signal
        If IsEmpty(velocityArr(i)) Then
            outputHK(i, 1) = ""
        Else
            outputHK(i, 1) = velocityArr(i)
        End If
        outputHK(i, 2) = zeroCrossArr(i)
        If IsEmpty(accelArr(i)) Then
            outputHK(i, 3) = ""
        Else
            outputHK(i, 3) = accelArr(i)
        End If
        outputHK(i, 4) = accelSignalArr(i)

        ' L-R: Signal_Type, Entry_Price, Entry_Tick, Success_Price, Fail_Price, Signal_Status, Accel_Count
        outputLR(i, 1) = signalTypeArr(i)
        If IsEmpty(entryPriceArr(i)) Then
            outputLR(i, 2) = ""
        Else
            outputLR(i, 2) = entryPriceArr(i)
        End If
        If IsEmpty(entryTickArr(i)) Then
            outputLR(i, 3) = ""
        Else
            outputLR(i, 3) = entryTickArr(i)
        End If
        If IsEmpty(successPriceArr(i)) Then
            outputLR(i, 4) = ""
        Else
            outputLR(i, 4) = successPriceArr(i)
        End If
        If IsEmpty(failPriceArr(i)) Then
            outputLR(i, 5) = ""
        Else
            outputLR(i, 5) = failPriceArr(i)
        End If
        outputLR(i, 6) = signalStatusArr(i)
        If IsEmpty(accelCountArr(i)) Then
            outputLR(i, 7) = ""
        Else
            outputLR(i, 7) = accelCountArr(i)
        End If

        ' S: Vol_Flag
        outputS(i, 1) = volFlagArr(i)
    Next i

    ' BULK WRITE to worksheet (4 writes instead of ~60,000 individual writes for 10K rows)
    ws.Range("E2:G" & lastRow).Value = outputEG
    ws.Range("H2:K" & lastRow).Value = outputHK
    ws.Range("L2:R" & lastRow).Value = outputLR
    ws.Range("S2:S" & lastRow).Value = outputS

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' -------------------------------------------------------------------
' CSV IMPORT AND SHEET CREATION
' -------------------------------------------------------------------

Sub ImportCSVToSheet(filePath As String, sheetName As String)
    ' Import CSV file and create/update sheet with calculations
    ' FIXED: Added delimiter auto-detection for tab/comma files

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim templateWs As Worksheet
    Dim i As Long

    ' Check if sheet exists, if not create it
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create new sheet from template
        Set templateWs = ThisWorkbook.Sheets("orderflow")
        templateWs.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        Set ws = ActiveSheet
        ws.Name = sheetName

        ' Clear template data (keep headers and structure)
        ws.Range("A2:S10000").ClearContents
    Else
        ' Sheet exists - clear old data
        ws.Range("A2:S10000").ClearContents
    End If

    ' ---------------------------------------------------------
    ' IMPORT CSV DATA (Columns A-D only)
    ' FIXED: Added delimiter auto-detection
    ' ---------------------------------------------------------
    Dim fso As Object
    Dim textFile As Object
    Dim lineText As String
    Dim lineData() As String
    Dim rowNum As Long
    Dim delimiter As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set textFile = fso.OpenTextFile(filePath, 1) ' 1 = ForReading

    ' Skip header row and detect delimiter
    If Not textFile.AtEndOfStream Then
        lineText = textFile.ReadLine
        ' Detect delimiter from header
        If InStr(lineText, vbTab) > 0 Then
            delimiter = vbTab
        Else
            delimiter = ","
        End If
    End If

    ' Read data rows
    rowNum = 2
    Do While Not textFile.AtEndOfStream
        lineText = textFile.ReadLine
        lineData = Split(lineText, delimiter)

        If UBound(lineData) >= 3 Then
            ' Column A: Time (raw text, will be parsed later)
            ws.Cells(rowNum, 1).Value = Trim(lineData(0))

            ' Column B: Price
            ws.Cells(rowNum, 2).Value = CDbl(lineData(1))

            ' Column C: Vol(k)
            ws.Cells(rowNum, 3).Value = CDbl(lineData(2))

            ' Column D: W (aggressor)
            ws.Cells(rowNum, 4).Value = Trim(lineData(3))

            rowNum = rowNum + 1
        End If
    Loop

    textFile.Close
    Set textFile = Nothing
    Set fso = Nothing

    lastRow = rowNum - 1

    If lastRow < 2 Then
        MsgBox "No data found in " & sheetName, vbExclamation
        Exit Sub
    End If

    ' ---------------------------------------------------------
    ' SORT BY TIME
    ' ---------------------------------------------------------
    ' Parse all times first (convert text to dates)
    Dim parsedTimes() As Date
    ReDim parsedTimes(2 To lastRow)

    For i = 2 To lastRow
        parsedTimes(i) = ParseDateTime(ws.Cells(i, 1).Value)
    Next i

    ' Sort data by parsed time
    Dim sortRange As Range
    Set sortRange = ws.Range("A2:D" & lastRow)

    ' Create temporary column for sorting
    For i = 2 To lastRow
        ws.Cells(i, 20).Value = parsedTimes(i) ' Temp column T
    Next i

    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("T2:T" & lastRow), Order:=xlAscending

    With ws.Sort
        .SetRange ws.Range("A2:T" & lastRow)
        .Header = xlNo
        .Apply
    End With

    ' Clear temp column
    ws.Range("T2:T" & lastRow).ClearContents

    ' ---------------------------------------------------------
    ' CALCULATE ALL COLUMNS E-S
    ' ---------------------------------------------------------
    Call ProcessSingleStock(ws, lastRow)

    ' ---------------------------------------------------------
    ' SCROLL TO LAST ROW
    ' ---------------------------------------------------------
    ws.Activate
    ws.Cells(lastRow, 1).Select
    ActiveWindow.ScrollRow = Application.Max(1, lastRow - 10)
End Sub

' -------------------------------------------------------------------
' BATCH PROCESSOR
' FIXED: Added folder picker dialog (replaces hardcoded DATA_FOLDER)
' -------------------------------------------------------------------

Sub BatchProcessFolder()
    ' Main batch processing routine - process all CSV files in selected folder
    ' AUTO-GENERATES ranking table after completion
    ' Uses DEFAULT_DATA_FOLDER if it exists, otherwise shows folder picker

    Dim folderPath As String
    Dim fileName As String
    Dim fileCount As Long
    Dim successCount As Long
    Dim errorCount As Long
    Dim errorList As String
    Dim sheetName As String
    Dim startTime As Double

    ' Use default folder if it exists, otherwise show folder picker
    If Dir(DEFAULT_DATA_FOLDER, vbDirectory) <> "" Then
        folderPath = DEFAULT_DATA_FOLDER
    Else
        ' Default folder not found - show folder picker
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Select Folder Containing CSV Files (Default not found: " & DEFAULT_DATA_FOLDER & ")"
            .AllowMultiSelect = False
            If .Show = -1 Then
                folderPath = .SelectedItems(1)
            Else
                Exit Sub ' User cancelled
            End If
        End With

        ' Ensure folder path ends with backslash
        If Right(folderPath, 1) <> "\" Then
            folderPath = folderPath & "\"
        End If
    End If

    ' Initialize counters
    fileCount = 0
    successCount = 0
    errorCount = 0
    errorList = ""
    startTime = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Loop through all CSV files
    fileName = Dir(folderPath & "*.csv")

    Do While fileName <> ""
        fileCount = fileCount + 1

        ' Extract sheet name from filename (remove .csv extension)
        sheetName = Left(fileName, Len(fileName) - 4)

        ' Try to import and process
        On Error Resume Next
        Call ImportCSVToSheet(folderPath & fileName, sheetName)

        If Err.Number <> 0 Then
            errorCount = errorCount + 1
            errorList = errorList & vbCrLf & "  • " & fileName & " - " & Err.Description
            Err.Clear
        Else
            successCount = successCount + 1
        End If
        On Error GoTo 0

        ' Get next file
        fileName = Dir
    Loop

    ' ============================================================
    ' AUTO-GENERATE RANKING TABLE
    ' ============================================================
    If successCount > 0 Then
        Call GenerateRankingTable
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' Show summary
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime

    Dim msg As String
    msg = "BATCH PROCESSING COMPLETE" & vbCrLf & vbCrLf
    msg = msg & "Files found: " & fileCount & vbCrLf
    msg = msg & "Successful: " & successCount & vbCrLf
    msg = msg & "Errors: " & errorCount & vbCrLf
    msg = msg & "Time: " & Format(elapsedTime, "0.0") & " seconds"
    msg = msg & "Folder: " & folderPath

    If errorCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "ERRORS:" & errorList
    End If

    If successCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "✓ Ranking table generated automatically"
    End If

    MsgBox msg, vbInformation, "Batch Processing Summary"
End Sub

' ===========================================
' QUICK RANKING UPDATE - Memory-only processing
' FIXED: Added folder picker and removed cumDelta(0)=0 error
' ===========================================

Sub QuickRankingUpdate()
    '=========================================
    ' OPTIMIZED VERSION - Uses cached lookups
    ' Fast batch processing - no sheet creation
    ' Reads CSVs, calculates in memory, updates Ranking only
    ' Uses DEFAULT_DATA_FOLDER if it exists, otherwise shows folder picker
    '=========================================

    Dim folderPath As String
    Dim fileName As String
    Dim ticker As String
    Dim fileCount As Long
    Dim successCount As Long
    Dim errorCount As Long
    Dim errorList As String
    Dim startTime As Double

    ' Signal collection arrays
    Dim bullSignals() As Variant
    Dim bearSignals() As Variant
    Dim bullCount As Long
    Dim bearCount As Long
    Dim signalResult As Variant

    ' Use default folder if it exists, otherwise show folder picker
    If Dir(DEFAULT_DATA_FOLDER, vbDirectory) <> "" Then
        folderPath = DEFAULT_DATA_FOLDER
    Else
        ' Default folder not found - show folder picker
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Select Folder Containing CSV Files (Default not found: " & DEFAULT_DATA_FOLDER & ")"
            .AllowMultiSelect = False
            If .Show = -1 Then
                folderPath = .SelectedItems(1)
            Else
                Exit Sub ' User cancelled
            End If
        End With

        ' Ensure folder path ends with backslash
        If Right(folderPath, 1) <> "\" Then
            folderPath = folderPath & "\"
        End If
    End If

    ' Initialize lookup caches for O(1) lookups
    Call InitializeLookupCaches

    ' Initialize
    ReDim bullSignals(1 To 100, 1 To 9)
    ReDim bearSignals(1 To 100, 1 To 9)
    bullCount = 0
    bearCount = 0
    fileCount = 0
    successCount = 0
    errorCount = 0
    errorList = ""
    startTime = Timer

    ' Verify folder exists
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Data folder not found:" & vbCrLf & folderPath, vbCritical, "Folder Error"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ---------------------------------------------------------
    ' LOOP THROUGH ALL CSV FILES
    ' ---------------------------------------------------------
    fileName = Dir(folderPath & "*.csv")

    Do While fileName <> ""
        fileCount = fileCount + 1
        ticker = Left(fileName, Len(fileName) - 4)  ' Remove .csv

        On Error Resume Next
        signalResult = ProcessCSVToSignal(folderPath & fileName)

        If Err.Number <> 0 Then
            errorCount = errorCount + 1
            errorList = errorList & vbCrLf & "  • " & fileName & " - " & Err.Description
            Err.Clear
        ElseIf Not IsEmpty(signalResult) Then
            successCount = successCount + 1

            ' Add to appropriate array based on signal type
            If signalResult(1) = "BULL" Then
                bullCount = bullCount + 1
                ' Look up ticker code using CACHED lookup (O(1))
                Dim tickerCodeBull As String
                tickerCodeBull = LookupTickerCodeCached(ticker)

                bullSignals(bullCount, 1) = ""  ' Rank
                bullSignals(bullCount, 2) = ticker  ' Stock Name
                bullSignals(bullCount, 3) = tickerCodeBull  ' Ticker
                bullSignals(bullCount, 4) = signalResult(1)  ' Signal_Type
                bullSignals(bullCount, 5) = signalResult(2)  ' Signal_Status
                bullSignals(bullCount, 6) = signalResult(3)  ' Accel_Count
                bullSignals(bullCount, 7) = signalResult(4)  ' Entry_Price

                ' Check Bullish/Bearish flags using CACHED lookup (O(1))
                If IsBullishCached(tickerCodeBull) Then
                    bullSignals(bullCount, 8) = "Bullish"
                Else
                    bullSignals(bullCount, 8) = ""
                End If
                If IsBearishCached(tickerCodeBull) Then
                    bullSignals(bullCount, 9) = "Bearish"
                Else
                    bullSignals(bullCount, 9) = ""
                End If
            ElseIf signalResult(1) = "BEAR" Then
                bearCount = bearCount + 1
                ' Look up ticker code using CACHED lookup (O(1))
                Dim tickerCodeBear As String
                tickerCodeBear = LookupTickerCodeCached(ticker)

                bearSignals(bearCount, 1) = ""  ' Rank
                bearSignals(bearCount, 2) = ticker  ' Stock Name
                bearSignals(bearCount, 3) = tickerCodeBear  ' Ticker
                bearSignals(bearCount, 4) = signalResult(1)  ' Signal_Type
                bearSignals(bearCount, 5) = signalResult(2)  ' Signal_Status
                bearSignals(bearCount, 6) = signalResult(3)  ' Accel_Count
                bearSignals(bearCount, 7) = signalResult(4)  ' Entry_Price

                ' Check Bullish/Bearish flags using CACHED lookup (O(1))
                If IsBullishCached(tickerCodeBear) Then
                    bearSignals(bearCount, 8) = "Bullish"
                Else
                    bearSignals(bearCount, 8) = ""
                End If
                If IsBearishCached(tickerCodeBear) Then
                    bearSignals(bearCount, 9) = "Bearish"
                Else
                    bearSignals(bearCount, 9) = ""
                End If
            End If
        Else
            successCount = successCount + 1  ' Processed but no active signal
        End If
        On Error GoTo 0

        fileName = Dir
    Loop

    ' ---------------------------------------------------------
    ' WRITE TO RANKING SHEET
    ' ---------------------------------------------------------
    Call WriteQuickRanking(bullSignals, bullCount, bearSignals, bearCount)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' ---------------------------------------------------------
    ' SHOW SUMMARY
    ' ---------------------------------------------------------
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime

    Dim activeBull As Long, activeBear As Long
    activeBull = 0: activeBear = 0

    Dim i As Long
    For i = 1 To bullCount
        If InStr(bullSignals(i, 4), "Active") > 0 Then activeBull = activeBull + 1
    Next i
    For i = 1 To bearCount
        If InStr(bearSignals(i, 4), "Active") > 0 Then activeBear = activeBear + 1
    Next i

    ' Cleanup caches
    Call ClearLookupCaches

    Dim msg As String
    msg = "QUICK RANKING COMPLETE" & vbCrLf & vbCrLf
    msg = msg & "Files processed: " & fileCount & vbCrLf
    msg = msg & "Successful: " & successCount & vbCrLf
    msg = msg & "Errors: " & errorCount & vbCrLf
    msg = msg & "Time: " & Format(elapsedTime, "0.00") & " seconds" & vbCrLf & vbCrLf
    msg = "Active BULL: " & activeBull & vbCrLf
    msg = msg & "Active BEAR: " & activeBear
    msg = msg & "Folder: " & folderPath

    If errorCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "ERRORS:" & errorList
    End If

    MsgBox msg, vbInformation, "Quick Ranking"
End Sub

Function ProcessCSVToSignal(filePath As String) As Variant
    '=========================================
    ' FIXED VERSION: Process single CSV entirely in memory
    ' Returns array: (Signal_Type, Signal_Status, Accel_Count, Entry_Price, Current_Price, Success_Price, Fail_Price)
    ' Returns Empty if no signal
    ' FIXED: Removed cumDelta(0)=0 line, added delimiter detection, enhanced bounds checking
    '=========================================

    Dim fso As Object
    Dim textFile As Object
    Dim lineText As String
    Dim lineData() As String
    Dim delimiter As String

    ' Data arrays
    Dim timeArr() As Date
    Dim priceArr() As Double
    Dim volArr() As Double
    Dim aggressorArr() As String
    Dim rowCount As Long
    Dim maxRows As Long

    ' Calculation arrays
    Dim signedVol() As Double
    Dim cumDelta() As Double
    Dim elapsedSec() As Double
    Dim velocity() As Variant
    Dim zeroCross() As String
    Dim accel() As Variant
    Dim accelSignal() As String

    ' Signal tracking
    Dim signalType As String
    Dim signalStatus As String
    Dim entryPrice As Double
    Dim entryTick As Double
    Dim accelCount As Long
    Dim successPrice As Double
    Dim failPrice As Double

    Dim i As Long, j As Long
    Dim startTime As Date
    Dim targetTime As Double
    Dim pastCumDelta As Variant
    Dim currVel As Variant, prevVel As Variant
    Dim tempDate As Date, tempDbl As Double, tempStr As String

    On Error GoTo ErrorHandler

    ' Initialize
    maxRows = 50000
    ReDim timeArr(1 To maxRows)
    ReDim priceArr(1 To maxRows)
    ReDim volArr(1 To maxRows)
    ReDim aggressorArr(1 To maxRows)
    rowCount = 0

    ' ---------------------------------------------------------
    ' READ CSV INTO ARRAYS WITH DELIMITER DETECTION
    ' FIXED: Auto-detect tab vs comma delimiters
    ' ---------------------------------------------------------
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set textFile = fso.OpenTextFile(filePath, 1)

    ' Skip header and detect delimiter
    If Not textFile.AtEndOfStream Then
        lineText = textFile.ReadLine
        ' Detect delimiter from header line
        If InStr(lineText, vbTab) > 0 Then
            delimiter = vbTab
        Else
            delimiter = ","
        End If
    End If

    ' Read data rows
    Do While Not textFile.AtEndOfStream
        lineText = textFile.ReadLine
        lineData = Split(lineText, delimiter)

        If UBound(lineData) >= 3 Then
            rowCount = rowCount + 1
            timeArr(rowCount) = ParseDateTime(Trim(lineData(0)))
            priceArr(rowCount) = CDbl(lineData(1))
            volArr(rowCount) = CDbl(lineData(2))
            aggressorArr(rowCount) = Trim(lineData(3))
        End If
    Loop

    textFile.Close
    Set textFile = Nothing
    Set fso = Nothing

    If rowCount < 2 Then
        ProcessCSVToSignal = Empty
        Exit Function
    End If

    ' Resize arrays to actual size
    ReDim Preserve timeArr(1 To rowCount)
    ReDim Preserve priceArr(1 To rowCount)
    ReDim Preserve volArr(1 To rowCount)
    ReDim Preserve aggressorArr(1 To rowCount)

    ' ---------------------------------------------------------
    ' SORT BY TIME (simple bubble sort)
    ' ---------------------------------------------------------
    For i = 1 To rowCount - 1
        For j = i + 1 To rowCount
            If timeArr(i) > timeArr(j) Then
                ' Swap all arrays
                tempDate = timeArr(i): timeArr(i) = timeArr(j): timeArr(j) = tempDate
                tempDbl = priceArr(i): priceArr(i) = priceArr(j): priceArr(j) = tempDbl
                tempDbl = volArr(i): volArr(i) = volArr(j): volArr(j) = tempDbl
                tempStr = aggressorArr(i): aggressorArr(i) = aggressorArr(j): aggressorArr(j) = tempStr
            End If
        Next j
    Next i

    ' ---------------------------------------------------------
    ' CALCULATE COLUMNS E-K (FIXED VERSION)
    ' FIXED: Removed cumDelta(0) = 0 line that caused array bounds error
    ' ---------------------------------------------------------
    ReDim signedVol(1 To rowCount)
    ReDim cumDelta(1 To rowCount)
    ReDim elapsedSec(1 To rowCount)
    ReDim velocity(1 To rowCount)
    ReDim zeroCross(1 To rowCount)
    ReDim accel(1 To rowCount)
    ReDim accelSignal(1 To rowCount)

    startTime = timeArr(1)
    ' FIXED: Removed cumDelta(0) = 0 line that caused "Subscript out of range"

    ' First pass: Signed_Vol, Cum_Delta, Elapsed_Sec
    For i = 1 To rowCount
        ' Signed_Vol
        If aggressorArr(i) = "s" Then
            signedVol(i) = volArr(i)
        ElseIf aggressorArr(i) = "b" Then
            signedVol(i) = -volArr(i)
        Else
            signedVol(i) = 0
        End If

        ' Cum_Delta - FIXED: Proper initialization without index 0
        If i = 1 Then
            cumDelta(i) = signedVol(i)
        Else
            cumDelta(i) = cumDelta(i - 1) + signedVol(i)
        End If

        ' Elapsed_Sec
        elapsedSec(i) = (timeArr(i) - startTime) * 86400
    Next i

    ' Second pass: Velocity, Zero_Cross, Accel, Accel_Signal
    For i = 1 To rowCount
        ' Velocity (5-min rolling) - FIXED: Enhanced bounds checking
        If elapsedSec(i) < 300 Then
            velocity(i) = Empty
        Else
            targetTime = elapsedSec(i) - 300
            pastCumDelta = Empty

            ' Walk backwards safely
            For j = i To 1 Step -1
                If j < 1 Then Exit For  ' Safety check
                If elapsedSec(j) <= targetTime Then
                    pastCumDelta = cumDelta(j)
                    Exit For
                End If
            Next j

            If IsEmpty(pastCumDelta) Or pastCumDelta = 0 Then
                velocity(i) = Empty
            Else
                velocity(i) = (cumDelta(i) - pastCumDelta) / 300
            End If
        End If

        ' Zero_Cross - FIXED: Enhanced bounds checking
        If i < 2 Or IsEmpty(velocity(i)) Then
            zeroCross(i) = ""
        Else
            If i - 1 >= LBound(velocity) And i - 1 <= UBound(velocity) Then
                prevVel = velocity(i - 1)
            Else
                prevVel = Empty
            End If
            currVel = velocity(i)

            If IsEmpty(prevVel) Then
                zeroCross(i) = ""
            ElseIf prevVel < 0 And currVel >= 0 Then
                zeroCross(i) = "BULL"
            ElseIf prevVel >= 0 And currVel < 0 Then
                zeroCross(i) = "BEAR"
            Else
                zeroCross(i) = ""
            End If
        End If

        ' Accel - FIXED: Enhanced bounds checking
        If i < 2 Or IsEmpty(velocity(i)) Or i - 1 < LBound(velocity) Then
            accel(i) = Empty
        Else
            accel(i) = velocity(i) - velocity(i - 1)
        End If

        ' Accel_Signal
        If IsEmpty(velocity(i)) Or IsEmpty(accel(i)) Then
            accelSignal(i) = ""
        ElseIf velocity(i) > 0 And accel(i) > 0 Then
            accelSignal(i) = "BULL ACCEL"
        ElseIf velocity(i) > 0 And accel(i) <= 0 Then
            accelSignal(i) = "BULL DECEL"
        ElseIf velocity(i) <= 0 And accel(i) < 0 Then
            accelSignal(i) = "BEAR ACCEL"
        ElseIf velocity(i) <= 0 And accel(i) >= 0 Then
            accelSignal(i) = "BEAR DECEL"
        Else
            accelSignal(i) = ""
        End If
    Next i

    ' ---------------------------------------------------------
    ' SIGNAL TRACKING (L-R logic)
    ' ---------------------------------------------------------
    signalType = ""
    signalStatus = ""
    entryPrice = 0
    entryTick = 0
    accelCount = 0
    successPrice = 0
    failPrice = 0

    For i = 1 To rowCount
        ' New signal detected
        If zeroCross(i) = "BULL" Or zeroCross(i) = "BEAR" Then
            signalType = zeroCross(i)
            entryPrice = priceArr(i)
            entryTick = GetTickSize(entryPrice)
            accelCount = 0

            If signalType = "BULL" Then
                signalStatus = "Active Bullish"
                successPrice = entryPrice + entryTick
                failPrice = entryPrice - (2 * entryTick)
            Else
                signalStatus = "Active Bearish"
                successPrice = entryPrice - entryTick
                failPrice = entryPrice + (2 * entryTick)
            End If
        End If

        ' Check for success/failure (only while Active)
        If InStr(signalStatus, "Active") > 0 Then
            If signalType = "BULL" Then
                If priceArr(i) >= successPrice Then
                    signalStatus = "Success Bullish"
                ElseIf priceArr(i) <= failPrice Then
                    signalStatus = "Failed Bullish"
                End If
            ElseIf signalType = "BEAR" Then
                If priceArr(i) <= successPrice Then
                    signalStatus = "Success Bearish"
                ElseIf priceArr(i) >= failPrice Then
                    signalStatus = "Failed Bearish"
                End If
            End If
        End If

        ' Count accelerations (only while Active)
        If InStr(signalStatus, "Active") > 0 Then
            If signalType = "BULL" And accelSignal(i) = "BULL ACCEL" Then
                accelCount = accelCount + 1
            ElseIf signalType = "BEAR" And accelSignal(i) = "BEAR ACCEL" Then
                accelCount = accelCount + 1
            End If
        End If
    Next i

    ' ---------------------------------------------------------
    ' RETURN RESULT
    ' ---------------------------------------------------------
    If signalType = "" Or signalStatus = "" Then
        ProcessCSVToSignal = Empty
    Else
        Dim result(1 To 7) As Variant
        result(1) = signalType
        result(2) = signalStatus
        result(3) = accelCount
        result(4) = entryPrice
        result(5) = priceArr(rowCount)  ' Current price (last row)
        result(6) = successPrice
        result(7) = failPrice
        ProcessCSVToSignal = result
    End If

    Exit Function

ErrorHandler:
    ' Enhanced error reporting
    Dim errorMsg As String
    errorMsg = "Error in ProcessCSVToSignal:" & vbCrLf
    errorMsg = errorMsg & "File: " & filePath & vbCrLf
    errorMsg = errorMsg & "Error #" & Err.Number & ": " & Err.Description & vbCrLf
    errorMsg = errorMsg & "RowCount: " & rowCount & vbCrLf
    errorMsg = errorMsg & "Current i: " & i & vbCrLf

    Debug.Print errorMsg
    MsgBox errorMsg, vbCritical, "ProcessCSVToSignal Error"

    ProcessCSVToSignal = Empty
End Function

Sub WriteQuickRanking(bullSignals() As Variant, bullCount As Long, _
                       bearSignals() As Variant, bearCount As Long)
    '=========================================
    ' Write collected signals to Ranking sheet
    ' Same format as GenerateRankingTable
    '=========================================

    Dim rankWs As Worksheet
    Dim bullStartRow As Long
    Dim bearStartRow As Long
    Dim batchTimestamp As String
    Dim i As Long

    batchTimestamp = "Batch: " & Format(Now, "DD-MMM-YYYY HH:MM")

    ' ---------------------------------------------------------
    ' GET OR CREATE RANKING SHEET
    ' ---------------------------------------------------------
    On Error Resume Next
    Set rankWs = ThisWorkbook.Sheets("Ranking")
    On Error GoTo 0

    Dim bullLastRow As Long
    Dim bearLastRow As Long
    Dim maxLastRow As Long
    
    If rankWs Is Nothing Then
        Set rankWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(2))
        rankWs.Name = "Ranking"
        bullStartRow = 1
        bearStartRow = 1
    Else
        ' Find last used row in BULL (column A)
        bullLastRow = rankWs.Cells(rankWs.Rows.count, 1).End(xlUp).Row
        If bullLastRow = 1 And rankWs.Cells(1, 1).Value = "" Then bullLastRow = 0
        
        ' Find last used row in BEAR (column K)
        bearLastRow = rankWs.Cells(rankWs.Rows.count, 11).End(xlUp).Row
        If bearLastRow = 1 And rankWs.Cells(1, 11).Value = "" Then bearLastRow = 0
        
        ' Use maximum of both to align timestamps
        maxLastRow = Application.WorksheetFunction.Max(bullLastRow, bearLastRow)
        
        ' Both start on same row (max + 2 for blank separator)
        If maxLastRow > 0 Then
            bullStartRow = maxLastRow + 2
            bearStartRow = maxLastRow + 2
        Else
            bullStartRow = 1
            bearStartRow = 1
        End If
    End If

    ' ---------------------------------------------------------
    ' WRITE BULL SECTION (Columns A-I)
    ' ---------------------------------------------------------
    ' Sort bull data FIRST (needed for correct watchlist order)
    If bullCount > 0 Then
        Call SortSignalArray(bullSignals, bullCount)
    End If

    With rankWs
        ' Timestamp row
        .Cells(bullStartRow, 1).Value = batchTimestamp
        .Cells(bullStartRow, 1).Font.Bold = True
        .Cells(bullStartRow, 1).Font.Italic = True
        .Cells(bullStartRow, 1).Interior.Color = RGB(198, 224, 180)
        .Range(.Cells(bullStartRow, 1), .Cells(bullStartRow, 9)).Merge
        bullStartRow = bullStartRow + 1

        ' TradingView watchlist string row
        .Cells(bullStartRow, 2).Value = BuildTradingViewString(bullSignals, bullCount)
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

    ' Write bull data (already sorted above) - BULK WRITE OPTIMIZED
    If bullCount > 0 Then
        ' Prepare output array for bulk write
        Dim bullOutput() As Variant
        ReDim bullOutput(1 To bullCount, 1 To 9)

        For i = 1 To bullCount
            bullSignals(i, 1) = i  ' Rank
            bullOutput(i, 1) = i                    ' Rank
            bullOutput(i, 2) = bullSignals(i, 2)   ' Stock
            bullOutput(i, 3) = bullSignals(i, 3)   ' Ticker
            bullOutput(i, 4) = bullSignals(i, 7)   ' Entry_Price
            bullOutput(i, 5) = bullSignals(i, 6)   ' Accel_Count
            bullOutput(i, 6) = bullSignals(i, 8)   ' Bullish
            bullOutput(i, 7) = bullSignals(i, 9)   ' Bearish
            bullOutput(i, 8) = bullSignals(i, 4)   ' Signal_Type
            bullOutput(i, 9) = bullSignals(i, 5)   ' Signal_Status
        Next i

        ' BULK WRITE (single operation instead of 9 × bullCount operations)
        rankWs.Range(rankWs.Cells(bullStartRow, 1), rankWs.Cells(bullStartRow + bullCount - 1, 9)).Value = bullOutput

        ' Apply highlighting (still row by row, but this is fast)
        For i = 1 To bullCount
            Call HighlightSignalRow(rankWs, bullStartRow + i - 1, 1, 9, _
                                    CStr(bullSignals(i, 5)), i, "BULL")
        Next i
    End If

    ' ---------------------------------------------------------
    ' WRITE BEAR SECTION (Columns K-S)
    ' ---------------------------------------------------------
    ' Sort bear data FIRST (needed for correct watchlist order)
    If bearCount > 0 Then
        Call SortSignalArray(bearSignals, bearCount)
    End If

    With rankWs
        ' Timestamp row
        .Cells(bearStartRow, 11).Value = batchTimestamp
        .Cells(bearStartRow, 11).Font.Bold = True
        .Cells(bearStartRow, 11).Font.Italic = True
        .Cells(bearStartRow, 11).Interior.Color = RGB(244, 204, 204)
        .Range(.Cells(bearStartRow, 11), .Cells(bearStartRow, 19)).Merge
        bearStartRow = bearStartRow + 1

        ' TradingView watchlist string row
        .Cells(bearStartRow, 12).Value = BuildTradingViewString(bearSignals, bearCount)
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

    ' Write bear data (already sorted above) - BULK WRITE OPTIMIZED
    If bearCount > 0 Then
        ' Prepare output array for bulk write
        Dim bearOutput() As Variant
        ReDim bearOutput(1 To bearCount, 1 To 9)

        For i = 1 To bearCount
            bearSignals(i, 1) = i  ' Rank
            bearOutput(i, 1) = i                    ' Rank
            bearOutput(i, 2) = bearSignals(i, 2)   ' Stock
            bearOutput(i, 3) = bearSignals(i, 3)   ' Ticker
            bearOutput(i, 4) = bearSignals(i, 7)   ' Entry_Price
            bearOutput(i, 5) = bearSignals(i, 6)   ' Accel_Count
            bearOutput(i, 6) = bearSignals(i, 9)   ' Bearish (note: swapped for BEAR section)
            bearOutput(i, 7) = bearSignals(i, 8)   ' Bullish
            bearOutput(i, 8) = bearSignals(i, 4)   ' Signal_Type
            bearOutput(i, 9) = bearSignals(i, 5)   ' Signal_Status
        Next i

        ' BULK WRITE (single operation instead of 9 × bearCount operations)
        rankWs.Range(rankWs.Cells(bearStartRow, 11), rankWs.Cells(bearStartRow + bearCount - 1, 19)).Value = bearOutput

        ' Apply highlighting (still row by row, but this is fast)
        For i = 1 To bearCount
            Call HighlightSignalRow(rankWs, bearStartRow + i - 1, 11, 19, _
                                    CStr(bearSignals(i, 5)), i, "BEAR")
        Next i
    End If

    ' ---------------------------------------------------------
    ' FORMAT AND POSITION
    ' ---------------------------------------------------------
    ' AutoFit columns except B and L (watchlist string columns)
    ' Watchlist strings overflow naturally - user copies from formula bar
    rankWs.Columns("A").AutoFit
    rankWs.Columns("C:I").AutoFit
    rankWs.Columns("K").AutoFit
    rankWs.Columns("M:S").AutoFit
    rankWs.Columns("J").ColumnWidth = 3

    ' Move to position 3 if needed
    On Error Resume Next
    If rankWs.Index > 3 Then
        rankWs.Move Before:=ThisWorkbook.Sheets(3)
    End If
    On Error GoTo 0

    ' Activate and scroll to latest
    rankWs.Activate
    rankWs.Cells(rankWs.Cells(rankWs.Rows.count, 1).End(xlUp).Row, 1).Select
End Sub

' -------------------------------------------------------------------
' SUPPORTING FUNCTIONS FROM Module3.bas
' -------------------------------------------------------------------

Sub GenerateRankingTable()
    ' ===========================================
    ' OPTIMIZED VERSION - Uses cached lookups
    ' Scan all sheets, extract last-row signals, compile ranking table
    ' Side-by-side layout: BULL (A-I) | BEAR (K-R)
    ' Appends to existing data with timestamp separator
    ' ===========================================

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
    Dim lastRow As Long

    Application.ScreenUpdating = False

    ' ---------------------------------------------------------
    ' INITIALIZE LOOKUP CACHES (O(1) lookups instead of O(n))
    ' ---------------------------------------------------------
    Call InitializeLookupCaches

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
           And LCase(sheetName) <> "data" And LCase(sheetName) <> "orderflow" _
           And LCase(sheetName) <> "watchlist" And LCase(sheetName) <> "bullish" _
           And LCase(sheetName) <> "bearish" Then

            ' Find last row with data
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

            If lastRow >= 2 Then
                ' Extract last row data
                signalType = ws.Cells(lastRow, 12).Value   ' Column L
                signalStatus = ws.Cells(lastRow, 17).Value ' Column Q
                accelCount = ws.Cells(lastRow, 18).Value   ' Column R
                entryPrice = ws.Cells(lastRow, 13).Value   ' Column M

                ' Look up ticker code using CACHED lookup (O(1) instead of O(n))
                Dim tickerCode As String
                tickerCode = LookupTickerCodeCached(sheetName)

                ' Check Bullish/Bearish using CACHED lookup (O(1) instead of O(n))
                If IsBullishCached(tickerCode) Then
                    bullishFlag = "Bullish"
                Else
                    bullishFlag = ""
                End If

                If IsBearishCached(tickerCode) Then
                    bearishFlag = "Bearish"
                Else
                    bearishFlag = ""
                End If

                ' Collect BULL signals
                If signalType = "BULL" And signalStatus <> "" Then
                    bullCount = bullCount + 1
                    bullData(bullCount, 1) = ""  ' Rank (fill after sorting)
                    bullData(bullCount, 2) = sheetName  ' Stock Name
                    bullData(bullCount, 3) = tickerCode  ' Ticker
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
                    bearData(bearCount, 2) = sheetName  ' Stock Name
                    bearData(bearCount, 3) = tickerCode  ' Ticker
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

    Dim bullLastRow As Long
    Dim bearLastRow As Long
    Dim maxLastRow As Long
    
    If rankWs Is Nothing Then
        ' Create new Ranking sheet at position 3
        Set rankWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(2))
        rankWs.Name = "Ranking"
        bullStartRow = 1
        bearStartRow = 1
    Else
        ' Find last used row in BULL (column A)
        bullLastRow = rankWs.Cells(rankWs.Rows.count, 1).End(xlUp).Row
        If bullLastRow = 1 And rankWs.Cells(1, 1).Value = "" Then bullLastRow = 0
        
        ' Find last used row in BEAR (column K)
        bearLastRow = rankWs.Cells(rankWs.Rows.count, 11).End(xlUp).Row
        If bearLastRow = 1 And rankWs.Cells(1, 11).Value = "" Then bearLastRow = 0
        
        ' Use maximum of both to align timestamps
        maxLastRow = Application.WorksheetFunction.Max(bullLastRow, bearLastRow)
        
        ' Both start on same row (max + 2 for blank separator)
        If maxLastRow > 0 Then
            bullStartRow = maxLastRow + 2
            bearStartRow = maxLastRow + 2
        Else
            bullStartRow = 1
            bearStartRow = 1
        End If
    End If

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
        .Cells(bullStartRow, 4).Value = "Signal_Type"
        .Cells(bullStartRow, 5).Value = "Signal_Status"
        .Cells(bullStartRow, 6).Value = "Accel_Count"
        .Cells(bullStartRow, 7).Value = "Entry_Price"
        .Cells(bullStartRow, 8).Value = "Bullish"
        .Cells(bullStartRow, 9).Value = "Bearish"

        .Range(.Cells(bullStartRow, 1), .Cells(bullStartRow, 9)).Font.Bold = True
        .Range(.Cells(bullStartRow, 1), .Cells(bullStartRow, 9)).Interior.Color = RGB(169, 208, 142)
        .Range(.Cells(bullStartRow, 1), .Cells(bullStartRow, 9)).HorizontalAlignment = xlCenter
        bullStartRow = bullStartRow + 1
    End With

    ' Write bull data (already sorted above) - BULK WRITE OPTIMIZED
    If bullCount > 0 Then
        ' Prepare output array for bulk write
        Dim bullOutput() As Variant
        ReDim bullOutput(1 To bullCount, 1 To 9)

        For i = 1 To bullCount
            bullData(i, 1) = i  ' Rank
            bullOutput(i, 1) = i              ' Rank
            bullOutput(i, 2) = bullData(i, 2) ' Stock
            bullOutput(i, 3) = bullData(i, 3) ' Ticker
            bullOutput(i, 4) = bullData(i, 4) ' Signal_Type
            bullOutput(i, 5) = bullData(i, 5) ' Signal_Status
            bullOutput(i, 6) = bullData(i, 6) ' Accel_Count
            bullOutput(i, 7) = bullData(i, 7) ' Entry_Price
            bullOutput(i, 8) = bullData(i, 8) ' Bullish
            bullOutput(i, 9) = bullData(i, 9) ' Bearish
        Next i

        ' BULK WRITE (single operation instead of 9 × bullCount operations)
        rankWs.Range(rankWs.Cells(bullStartRow, 1), rankWs.Cells(bullStartRow + bullCount - 1, 9)).Value = bullOutput

        ' Apply highlighting (still row by row, but formatting is fast)
        For i = 1 To bullCount
            Call HighlightSignalRow(rankWs, bullStartRow + i - 1, 1, 9, _
                                    CStr(bullData(i, 5)), i, "BULL")
        Next i
    End If

    ' ---------------------------------------------------------
    ' WRITE BEAR SECTION (Columns K-R)
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
        .Cells(bearStartRow, 14).Value = "Signal_Type"
        .Cells(bearStartRow, 15).Value = "Signal_Status"
        .Cells(bearStartRow, 16).Value = "Accel_Count"
        .Cells(bearStartRow, 17).Value = "Entry_Price"
        .Cells(bearStartRow, 18).Value = "Bullish"
        .Cells(bearStartRow, 19).Value = "Bearish"

        .Range(.Cells(bearStartRow, 11), .Cells(bearStartRow, 19)).Font.Bold = True
        .Range(.Cells(bearStartRow, 11), .Cells(bearStartRow, 19)).Interior.Color = RGB(230, 145, 145)
        .Range(.Cells(bearStartRow, 11), .Cells(bearStartRow, 19)).HorizontalAlignment = xlCenter
        bearStartRow = bearStartRow + 1
    End With

    ' Write bear data (already sorted above) - BULK WRITE OPTIMIZED
    If bearCount > 0 Then
        ' Prepare output array for bulk write
        Dim bearOutput() As Variant
        ReDim bearOutput(1 To bearCount, 1 To 9)

        For i = 1 To bearCount
            bearData(i, 1) = i  ' Rank
            bearOutput(i, 1) = i              ' Rank
            bearOutput(i, 2) = bearData(i, 2) ' Stock
            bearOutput(i, 3) = bearData(i, 3) ' Ticker
            bearOutput(i, 4) = bearData(i, 4) ' Signal_Type
            bearOutput(i, 5) = bearData(i, 5) ' Signal_Status
            bearOutput(i, 6) = bearData(i, 6) ' Accel_Count
            bearOutput(i, 7) = bearData(i, 7) ' Entry_Price
            bearOutput(i, 8) = bearData(i, 8) ' Bullish
            bearOutput(i, 9) = bearData(i, 9) ' Bearish
        Next i

        ' BULK WRITE (single operation instead of 9 × bearCount operations)
        rankWs.Range(rankWs.Cells(bearStartRow, 11), rankWs.Cells(bearStartRow + bearCount - 1, 19)).Value = bearOutput

        ' Apply highlighting (still row by row, but formatting is fast)
        For i = 1 To bearCount
            Call HighlightSignalRow(rankWs, bearStartRow + i - 1, 11, 19, _
                                    CStr(bearData(i, 5)), i, "BEAR")
        Next i
    End If

    ' ---------------------------------------------------------
    ' AUTO-FIT AND POSITION
    ' ---------------------------------------------------------
    ' AutoFit all columns except watchlist string columns
    rankWs.Columns("A:I").AutoFit
    rankWs.Columns("K:S").AutoFit
    rankWs.Columns("J").ColumnWidth = 3  ' Spacer column

    ' Move Ranking sheet to position 3 (after Data, OrderFlow)
    On Error Resume Next
    rankWs.Move Before:=ThisWorkbook.Sheets(3)
    On Error GoTo 0

    ' Activate ranking sheet and scroll to latest batch
    rankWs.Activate
    rankWs.Cells(rankWs.Cells(rankWs.Rows.count, 1).End(xlUp).Row, 1).Select

    Application.ScreenUpdating = True

    ' ---------------------------------------------------------
    ' CLEANUP CACHES
    ' ---------------------------------------------------------
    Call ClearLookupCaches

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
            status1 = CStr(dataArr(i, 5))  ' Column 5: Signal_Status
            status2 = CStr(dataArr(j, 5))
            accel1 = dataArr(i, 6)  ' Column 6: Accel_Count
            accel2 = dataArr(j, 6)

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
                For k = 1 To 9
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
    ' Ticker is already in Column 3 of the array
    '=========================================

    Dim result As String
    Dim i As Long
    Dim ticker As String

    result = ""

    For i = 1 To signalCount
        ticker = CStr(signalArr(i, 3))  ' Column 3 is ticker
        If ticker <> "" Then
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
