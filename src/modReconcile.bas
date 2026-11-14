Attribute VB_Name = "modReconcile"
'==============================================================================
' Excel Data Reconciliation Tool
' Author : Swaminathan Narayanan
' Repo   : https://github.com/Swami529796/excel-data-reconciliation
'
' Compares two datasets (Source A and Source B) on a key and reports, row by row:
'   Matched        - key found in both, all compared columns agree
'   Mismatch       - key found in both, at least one compared column differs
'   Missing in B   - key only exists in A
'   Missing in A   - key only exists in B
'   Duplicate key  - the key appears more than once in the same source
'   Blank key      - the key cell(s) are empty
'
'  - Key can be one column or several (for example: Invoice No, Line No)
'  - Numbers are compared with a tolerance, so rounding differences can be ignored
'  - Uses arrays and a dictionary, so 100,000 rows take seconds, not minutes
'  - Everything is driven by the Settings sheet; no code changes are needed
'
' Run: Reconcile
'==============================================================================
Option Explicit

'---- Output sheets -----------------------------------------------------------
Private Const RESULTS_SHEET As String = "Results"
Private Const SUMMARY_SHEET As String = "Summary"
Private Const KEY_SEPARATOR As String = " | "
Private Const FIXED_COLUMNS As Long = 5   ' Key, Status, Differences, Row in A, Row in B

'---- Status labels -----------------------------------------------------------
Private Const ST_MATCHED As String = "Matched"
Private Const ST_MISMATCH As String = "Mismatch"
Private Const ST_MISSING_B As String = "Missing in B"
Private Const ST_MISSING_A As String = "Missing in A"
Private Const ST_DUPLICATE_A As String = "Duplicate key in A"
Private Const ST_DUPLICATE_B As String = "Duplicate key in B"
Private Const ST_BLANK_A As String = "Blank key in A"
Private Const ST_BLANK_B As String = "Blank key in B"

'---- Settings (read from the Settings sheet) ---------------------------------
Private Type RunSettings
    SheetA As String
    SheetB As String
    KeyColumns As String
    CompareColumns As String
    Tolerance As Double
    IgnoreCase As Boolean
End Type

Private mCalcMode As XlCalculation

'==============================================================================
' Entry point
'==============================================================================
Public Sub Reconcile()
    Dim startTime As Double
    startTime = Timer

    On Error GoTo CleanFail
    SpeedUp True

    Dim cfg As RunSettings
    cfg = ReadSettings()

    ' read both datasets into memory in one go
    Dim dataA As Variant, dataB As Variant
    dataA = ReadSheet(cfg.SheetA)
    dataB = ReadSheet(cfg.SheetB)

    ' locate the key and compare columns in each dataset
    Dim keyA() As Long, keyB() As Long
    keyA = FindColumns(dataA, SplitList(cfg.KeyColumns), cfg.SheetA)
    keyB = FindColumns(dataB, SplitList(cfg.KeyColumns), cfg.SheetB)

    Dim compareNames As Collection
    Set compareNames = CompareColumnNames(dataA, dataB, cfg)

    Dim cmpA() As Long, cmpB() As Long
    cmpA = FindColumns(dataA, compareNames, cfg.SheetA)
    cmpB = FindColumns(dataB, compareNames, cfg.SheetB)

    Dim cmpCount As Long
    cmpCount = compareNames.Count

    ' index Source B by key
    Dim indexB As Object, usedB As Object
    Set indexB = NewDictionary()
    Set usedB = NewDictionary()

    Dim rowsA As Long, rowsB As Long
    rowsA = UBound(dataA, 1)
    rowsB = UBound(dataB, 1)

    Dim outArr() As Variant
    ReDim outArr(1 To rowsA + rowsB, 1 To FIXED_COLUMNS + 2 * cmpCount)

    Dim counts As Object
    Set counts = NewDictionary()

    Dim diffCells As New Collection   ' (row, column) pairs to highlight
    Dim n As Long, r As Long, c As Long
    Dim key As String

    For r = 2 To rowsB
        key = BuildKey(dataB, r, keyB)
        If IsBlankKey(key) Then
            n = n + 1
            WriteRow outArr, n, key, ST_BLANK_B, "", 0, r
            FillValues outArr, n, dataB, r, cmpB, 1
            AddCount counts, ST_BLANK_B
        ElseIf indexB.Exists(key) Then
            n = n + 1
            WriteRow outArr, n, key, ST_DUPLICATE_B, "First seen in row " & indexB(key), 0, r
            FillValues outArr, n, dataB, r, cmpB, 1
            AddCount counts, ST_DUPLICATE_B
        Else
            indexB.Add key, r
        End If
    Next r

    ' walk through Source A and look each key up in B
    Dim seenA As Object
    Set seenA = NewDictionary()

    Dim rowB As Long, differences As String, valueA As Variant, valueB As Variant

    For r = 2 To rowsA
        key = BuildKey(dataA, r, keyA)
        n = n + 1

        If IsBlankKey(key) Then
            WriteRow outArr, n, key, ST_BLANK_A, "", r, 0
            FillValues outArr, n, dataA, r, cmpA, 0
            AddCount counts, ST_BLANK_A

        ElseIf seenA.Exists(key) Then
            WriteRow outArr, n, key, ST_DUPLICATE_A, "First seen in row " & seenA(key), r, 0
            FillValues outArr, n, dataA, r, cmpA, 0
            AddCount counts, ST_DUPLICATE_A

        ElseIf Not indexB.Exists(key) Then
            seenA.Add key, r
            WriteRow outArr, n, key, ST_MISSING_B, "", r, 0
            FillValues outArr, n, dataA, r, cmpA, 0
            AddCount counts, ST_MISSING_B

        Else
            seenA.Add key, r
            rowB = indexB(key)
            usedB(key) = True
            differences = ""

            For c = 1 To cmpCount
                valueA = dataA(r, cmpA(c))
                valueB = dataB(rowB, cmpB(c))
                If Not ValuesMatch(valueA, valueB, cfg) Then
                    If Len(differences) > 0 Then differences = differences & "; "
                    differences = differences & compareNames(c) & ": " & CellText(valueA) & " vs " & CellText(valueB)
                    diffCells.Add Array(n, FIXED_COLUMNS + 2 * c - 1)
                End If
            Next c

            If Len(differences) = 0 Then
                WriteRow outArr, n, key, ST_MATCHED, "", r, rowB
                AddCount counts, ST_MATCHED
            Else
                WriteRow outArr, n, key, ST_MISMATCH, differences, r, rowB
                AddCount counts, ST_MISMATCH
            End If
            FillValues outArr, n, dataA, r, cmpA, 0
            FillValues outArr, n, dataB, rowB, cmpB, 1
        End If
    Next r

    ' anything left in B was never matched
    Dim k As Variant
    For Each k In indexB.Keys
        If Not usedB.Exists(k) Then
            n = n + 1
            WriteRow outArr, n, CStr(k), ST_MISSING_A, "", 0, indexB(k)
            FillValues outArr, n, dataB, indexB(k), cmpB, 1
            AddCount counts, ST_MISSING_A
        End If
    Next k

    WriteResults outArr, n, compareNames, diffCells
    WriteSummary counts, cfg, compareNames, rowsA - 1, rowsB - 1, Timer - startTime

    SpeedUp False
    ThisWorkbook.Worksheets(SUMMARY_SHEET).Activate

    MsgBox "Reconciliation complete." & vbLf & vbLf & _
           "Rows in A:     " & Format(rowsA - 1, "#,##0") & vbLf & _
           "Rows in B:     " & Format(rowsB - 1, "#,##0") & vbLf & _
           "Matched:       " & Format(GetCount(counts, ST_MATCHED), "#,##0") & vbLf & _
           "Mismatches:    " & Format(GetCount(counts, ST_MISMATCH), "#,##0") & vbLf & _
           "Missing in B:  " & Format(GetCount(counts, ST_MISSING_B), "#,##0") & vbLf & _
           "Missing in A:  " & Format(GetCount(counts, ST_MISSING_A), "#,##0") & vbLf & _
           "Time taken:    " & Format(Timer - startTime, "0.0") & " sec" & vbLf & vbLf & _
           "See the '" & RESULTS_SHEET & "' sheet for the row-by-row detail.", vbInformation
    Exit Sub

CleanFail:
    SpeedUp False
    MsgBox "Reconciliation stopped:" & vbLf & Err.Description, vbCritical
End Sub

'==============================================================================
' Settings and input
'==============================================================================
Private Function ReadSettings() As RunSettings
    Dim cfg As RunSettings
    cfg.SheetA = SettingText("SheetA", "Source A")
    cfg.SheetB = SettingText("SheetB", "Source B")
    cfg.KeyColumns = SettingText("KeyColumns", "")
    cfg.CompareColumns = SettingText("CompareColumns", "")
    cfg.IgnoreCase = (UCase$(Left$(SettingText("IgnoreCase", "Yes"), 1)) = "Y")

    Dim toleranceText As String
    toleranceText = SettingText("NumberTolerance", "0")
    If IsNumeric(toleranceText) Then cfg.Tolerance = Abs(CDbl(toleranceText))

    If Len(cfg.KeyColumns) = 0 Then
        Err.Raise vbObjectError + 1, , "Key Column(s) is empty on the Settings sheet."
    End If
    ReadSettings = cfg
End Function

' Value of a named range on the Settings sheet, or the default if it is blank
Private Function SettingText(ByVal rangeName As String, ByVal defaultValue As String) As String
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Names(rangeName).RefersToRange.Value
    On Error GoTo 0
    If IsError(v) Or IsEmpty(v) Then
        SettingText = defaultValue
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        SettingText = defaultValue
    Else
        SettingText = Trim$(CStr(v))
    End If
End Function

' Whole used block of a sheet as a 2-D array (row 1 = headers)
Private Function ReadSheet(ByVal sheetName As String) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 2, , "Sheet '" & sheetName & "' was not found. Check the Settings sheet."
    End If

    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        Err.Raise vbObjectError + 3, , "Sheet '" & sheetName & "' is empty."
    End If

    Dim lastRow As Long, lastCol As Long
    lastRow = lastCell.Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then
        Err.Raise vbObjectError + 4, , "Sheet '" & sheetName & "' has headers but no data rows."
    End If

    ReadSheet = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
End Function

' "Invoice No, Line No" -> collection of trimmed names
Private Function SplitList(ByVal text As String) As Collection
    Dim result As New Collection
    Dim part As Variant
    For Each part In Split(text, ",")
        If Len(Trim$(CStr(part))) > 0 Then result.Add Trim$(CStr(part))
    Next part
    Set SplitList = result
End Function

' Column numbers for a list of header names; stops with a clear message if one is missing
Private Function FindColumns(ByRef data As Variant, ByVal names As Collection, ByVal sheetName As String) As Long()
    Dim result() As Long
    ReDim result(1 To IIf(names.Count = 0, 1, names.Count))

    Dim i As Long
    For i = 1 To names.Count
        result(i) = HeaderIndex(data, CStr(names(i)))
        If result(i) = 0 Then
            Err.Raise vbObjectError + 5, , "Column '" & names(i) & "' was not found on sheet '" & sheetName & "'."
        End If
    Next i
    FindColumns = result
End Function

Private Function HeaderIndex(ByRef data As Variant, ByVal headerName As String) As Long
    Dim c As Long
    For c = 1 To UBound(data, 2)
        If Not IsError(data(1, c)) Then
            If StrComp(Trim$(CStr(data(1, c))), headerName, vbTextCompare) = 0 Then
                HeaderIndex = c
                Exit Function
            End If
        End If
    Next c
End Function

' Columns to compare: the list on the Settings sheet, or every non-key column
' that exists in both sources when the setting is blank
Private Function CompareColumnNames(ByRef dataA As Variant, ByRef dataB As Variant, _
                                    ByRef cfg As RunSettings) As Collection
    If Len(cfg.CompareColumns) > 0 Then
        Set CompareColumnNames = SplitList(cfg.CompareColumns)
        Exit Function
    End If

    Dim keys As Collection, result As New Collection
    Set keys = SplitList(cfg.KeyColumns)

    Dim c As Long, i As Long, headerName As String, isKey As Boolean
    For c = 1 To UBound(dataA, 2)
        If Not IsError(dataA(1, c)) Then
            headerName = Trim$(CStr(dataA(1, c)))
            isKey = False
            For i = 1 To keys.Count
                If StrComp(CStr(keys(i)), headerName, vbTextCompare) = 0 Then isKey = True
            Next i
            If Len(headerName) > 0 And Not isKey Then
                If HeaderIndex(dataB, headerName) > 0 Then result.Add headerName
            End If
        End If
    Next c
    Set CompareColumnNames = result
End Function

'==============================================================================
' Comparison helpers
'==============================================================================
' Key for one row: the key column values joined together, upper-cased and trimmed
Private Function BuildKey(ByRef data As Variant, ByVal r As Long, ByRef cols() As Long) As String
    Dim i As Long, result As String
    For i = LBound(cols) To UBound(cols)
        If i > LBound(cols) Then result = result & KEY_SEPARATOR
        result = result & UCase$(Trim$(CellText(data(r, cols(i)))))
    Next i
    BuildKey = result
End Function

' True when every key column of the row is empty
Private Function IsBlankKey(ByVal key As String) As Boolean
    IsBlankKey = (Len(Trim$(Replace(key, Trim$(KEY_SEPARATOR), ""))) = 0)
End Function

Private Function ValuesMatch(ByVal a As Variant, ByVal b As Variant, ByRef cfg As RunSettings) As Boolean
    If IsError(a) Or IsError(b) Then
        ValuesMatch = (CellText(a) = CellText(b))
    ElseIf IsBlank(a) And IsBlank(b) Then
        ValuesMatch = True
    ElseIf IsBlank(a) Or IsBlank(b) Then
        ValuesMatch = False
    ElseIf IsNumber(a) And IsNumber(b) Then
        ValuesMatch = (Abs(CDbl(a) - CDbl(b)) <= cfg.Tolerance + 0.0000001)
    ElseIf IsDate(a) And IsDate(b) Then
        ValuesMatch = (CDate(a) = CDate(b))
    Else
        ValuesMatch = (StrComp(Trim$(CStr(a)), Trim$(CStr(b)), _
                               IIf(cfg.IgnoreCase, vbTextCompare, vbBinaryCompare)) = 0)
    End If
End Function

Private Function IsBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlank = True
    ElseIf VarType(v) = vbString Then
        IsBlank = (Len(Trim$(v)) = 0)
    End If
End Function

' True for real numbers and for text that looks like a number; dates are not numbers here
Private Function IsNumber(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            IsNumber = True
        Case vbString
            IsNumber = IsNumeric(v)
    End Select
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then
        CellText = "#ERROR"
    ElseIf IsEmpty(v) Then
        CellText = ""
    ElseIf VarType(v) = vbDate Then
        CellText = Format$(v, "yyyy-mm-dd")
    Else
        CellText = CStr(v)
    End If
End Function

'==============================================================================
' Output
'==============================================================================
Private Sub WriteRow(ByRef outArr() As Variant, ByVal n As Long, ByVal key As String, ByVal status As String, _
                     ByVal differences As String, ByVal rowA As Long, ByVal rowB As Long)
    outArr(n, 1) = "'" & key      ' keep keys such as 00123 as text
    outArr(n, 2) = status
    outArr(n, 3) = differences
    If rowA > 0 Then outArr(n, 4) = rowA
    If rowB > 0 Then outArr(n, 5) = rowB
End Sub

' Copy the compared columns of one source row into the output (offset 0 = A side, 1 = B side)
Private Sub FillValues(ByRef outArr() As Variant, ByVal n As Long, ByRef data As Variant, ByVal r As Long, _
                       ByRef cols() As Long, ByVal sideOffset As Long)
    Dim c As Long
    For c = 1 To (UBound(outArr, 2) - FIXED_COLUMNS) \ 2
        outArr(n, FIXED_COLUMNS + 2 * c - 1 + sideOffset) = data(r, cols(c))
    Next c
End Sub

Private Sub WriteResults(ByRef outArr() As Variant, ByVal n As Long, ByVal compareNames As Collection, _
                         ByVal diffCells As Collection)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(RESULTS_SHEET)
    ws.Cells.Clear
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    Dim colCount As Long, c As Long
    colCount = UBound(outArr, 2)

    ws.Range("A1:E1").Value = Array("Key", "Status", "Differences", "Row in A", "Row in B")
    For c = 1 To compareNames.Count
        ws.Cells(1, FIXED_COLUMNS + 2 * c - 1).Value = compareNames(c) & " (A)"
        ws.Cells(1, FIXED_COLUMNS + 2 * c).Value = compareNames(c) & " (B)"
    Next c

    If n > 0 Then ws.Cells(2, 1).Resize(n, colCount).Value = outArr

    ' colour the status column
    Dim r As Long
    For r = 2 To n + 1
        Select Case ws.Cells(r, 2).Value
            Case ST_MATCHED
                ws.Cells(r, 2).Interior.Color = RGB(198, 239, 206)
            Case ST_MISMATCH
                ws.Cells(r, 2).Interior.Color = RGB(255, 235, 156)
            Case Else
                ws.Cells(r, 2).Interior.Color = RGB(255, 199, 206)
        End Select
    Next r

    ' highlight the cells that differ (both the A and the B value)
    Dim cell As Variant
    For Each cell In diffCells
        ws.Cells(cell(0) + 1, cell(1)).Resize(1, 2).Interior.Color = RGB(255, 235, 156)
        ws.Cells(cell(0) + 1, cell(1)).Resize(1, 2).Font.Bold = True
    Next cell

    StyleHeader ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
    ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount)).AutoFilter
    ws.UsedRange.Columns.AutoFit
    If ws.Columns(3).ColumnWidth > 60 Then ws.Columns(3).ColumnWidth = 60

    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 2
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

Private Sub WriteSummary(ByVal counts As Object, ByRef cfg As RunSettings, ByVal compareNames As Collection, _
                         ByVal rowsA As Long, ByVal rowsB As Long, ByVal seconds As Double)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SUMMARY_SHEET)
    ws.Cells.Clear

    ws.Range("A1").Value = "Reconciliation Summary"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Run on " & Format$(Now, "dd mmm yyyy hh:nn") & "  |  " & Format$(seconds, "0.0") & " sec"
    ws.Range("A2").Font.Color = RGB(89, 89, 89)

    ws.Range("A4:C4").Value = Array("Status", "Rows", "% of rows")
    Dim statuses As Variant, i As Long, total As Long, r As Long
    statuses = Array(ST_MATCHED, ST_MISMATCH, ST_MISSING_B, ST_MISSING_A, _
                     ST_DUPLICATE_A, ST_DUPLICATE_B, ST_BLANK_A, ST_BLANK_B)
    For i = LBound(statuses) To UBound(statuses)
        total = total + GetCount(counts, CStr(statuses(i)))
    Next i

    r = 5
    For i = LBound(statuses) To UBound(statuses)
        ws.Cells(r, 1).Value = statuses(i)
        ws.Cells(r, 2).Value = GetCount(counts, CStr(statuses(i)))
        If total > 0 Then ws.Cells(r, 3).Value = GetCount(counts, CStr(statuses(i))) / total
        Select Case statuses(i)
            Case ST_MATCHED:  ws.Cells(r, 1).Interior.Color = RGB(198, 239, 206)
            Case ST_MISMATCH: ws.Cells(r, 1).Interior.Color = RGB(255, 235, 156)
            Case Else:        ws.Cells(r, 1).Interior.Color = RGB(255, 199, 206)
        End Select
        r = r + 1
    Next i
    ws.Cells(r, 1).Value = "Total"
    ws.Cells(r, 2).Value = total
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = True
    ws.Range(ws.Cells(5, 3), ws.Cells(r, 3)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(5, 2), ws.Cells(r, 2)).NumberFormat = "#,##0"

    Dim compareList As String
    For i = 1 To compareNames.Count
        If i > 1 Then compareList = compareList & ", "
        compareList = compareList & compareNames(i)
    Next i

    r = r + 2
    ws.Cells(r, 1).Resize(1, 2).Value = Array("Setting", "Value")
    ws.Cells(r + 1, 1).Resize(1, 2).Value = Array("Source A", cfg.SheetA & " (" & Format$(rowsA, "#,##0") & " rows)")
    ws.Cells(r + 2, 1).Resize(1, 2).Value = Array("Source B", cfg.SheetB & " (" & Format$(rowsB, "#,##0") & " rows)")
    ws.Cells(r + 3, 1).Resize(1, 2).Value = Array("Key column(s)", cfg.KeyColumns)
    ws.Cells(r + 4, 1).Resize(1, 2).Value = Array("Compared columns", compareList)
    ws.Cells(r + 5, 1).Resize(1, 2).Value = Array("Number tolerance", cfg.Tolerance)
    ws.Cells(r + 6, 1).Resize(1, 2).Value = Array("Ignore case", IIf(cfg.IgnoreCase, "Yes", "No"))

    StyleHeader ws.Range("A4:C4")
    StyleHeader ws.Cells(r, 1).Resize(1, 2)

    ws.Columns("A:C").AutoFit
    ws.Columns(2).HorizontalAlignment = xlLeft
End Sub

'==============================================================================
' Small utilities
'==============================================================================
Private Sub StyleHeader(ByVal target As Range)
    target.Font.Bold = True
    target.Font.Color = vbWhite
    target.Interior.Color = RGB(31, 78, 120)
End Sub

Private Function NewDictionary() As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
    NewDictionary.CompareMode = vbTextCompare
End Function

Private Sub AddCount(ByVal counts As Object, ByVal status As String)
    counts(status) = GetCount(counts, status) + 1
End Sub

Private Function GetCount(ByVal counts As Object, ByVal status As String) As Long
    If counts.Exists(status) Then GetCount = counts(status)
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Sub SpeedUp(ByVal turnOn As Boolean)
    With Application
        If turnOn Then
            mCalcMode = .Calculation
            .Calculation = xlCalculationManual
        ElseIf mCalcMode <> 0 Then
            .Calculation = mCalcMode
        End If
        .ScreenUpdating = Not turnOn
        .EnableEvents = Not turnOn
    End With
End Sub
