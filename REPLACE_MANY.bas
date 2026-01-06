Option Explicit

' ==============================
'  REPLACE_MANY Toolkit (v2)
' ==============================
' - UDF: REPLACE_MANY()  -> spill-safe, array-aware, full-word, mapping-driven replace
' - Macro: REPLACE_MANY_POPUP -> in-place replace across Selection / Sheet / Workbook (popup flow)
'
' Mapping rules:
'   * Provide a 2-column range: [From, To]. Blank keys are ignored. Duplicates keep first.
'   * Matching is case-insensitive by default; optional case-sensitive mode in the popup.
'   * Full-word only via delimiter padding & tokenization. Punctuation is handled.
'
' Safety/perf:
'   * For in-place runs, formulas are skipped by default. You can opt-in to convert
'     only text-result formulas to values before replacement (safer than converting all).
'   * Uses array processing per Area for speed. ScreenUpdating, Events, Calculation managed.
'
' -----------------------------------------
' UDF: REPLACE_MANY (array / spill function)
' -----------------------------------------
Public Function REPLACE_MANY(ByVal InputRange As Variant, _
                             ByVal MapRange As Variant, _
                             Optional ByVal delims As String = vbNullString) As Variant
    Dim inArr As Variant, outArr As Variant, mapArr As Variant
    Dim dict As Object
    Dim keys() As String, vals() As Variant, lens() As Long
    Dim n As Long, i As Long, r As Long, c As Long
    Dim rowsCount As Long, colsCount As Long
    Dim rowsToProcess As Long
    Dim inputIsFullColumn As Boolean

    On Error GoTo CleanFail

    ' Coerce inputs to 2D arrays
    inArr = To2D(InputRange)
    rowsCount = UBound(inArr, 1)
    colsCount = UBound(inArr, 2)

    mapArr = To2D(MapRange)

    ' Build mapping (ignore blank keys)
    n = 0
    ReDim keys(1 To UBound(mapArr, 1))
    ReDim vals(1 To UBound(mapArr, 1))
    ReDim lens(1 To UBound(mapArr, 1))

    For i = 1 To UBound(mapArr, 1)
        Dim k As String
        k = NzString(mapArr(i, 1))
        If Len(k) > 0 Then
            n = n + 1
            keys(n) = k
            vals(n) = mapArr(i, 2)
            lens(n) = Len(k)
        End If
    Next i

    If n = 0 Then
        ' No map provided; return original input
        If rowsCount = 1 And colsCount = 1 Then
            REPLACE_MANY = inArr(1, 1)
        Else
            REPLACE_MANY = inArr
        End If
        Exit Function
    End If

    ReDim Preserve keys(1 To n)
    ReDim Preserve vals(1 To n)
    ReDim Preserve lens(1 To n)

    ' Sort by key length descending (satisfies decreasing length requirement)
    QuickSortByLength keys, vals, lens, 1, n

    ' Case-insensitive dictionary lookup
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' vbTextCompare
    For i = 1 To n
        If Not dict.Exists(keys(i)) Then dict.Add keys(i), vals(i)
    Next i

    ' Default delimiters if not provided
    If Len(delims) = 0 Then delims = DefaultDelims()

    ' Decide how many rows to process (optimize for full-column input)
    inputIsFullColumn = IsFullColumn(InputRange)
    If colsCount = 1 And inputIsFullColumn Then
        rowsToProcess = LastNonBlankInColumnArray(inArr)
        If rowsToProcess = 0 Then rowsToProcess = 0 ' all blank; keep empty result
    Else
        rowsToProcess = rowsCount
    End If

    ReDim outArr(1 To rowsCount, 1 To colsCount)

    If rowsToProcess > 0 Then
        For r = 1 To rowsToProcess
            For c = 1 To colsCount
                outArr(r, c) = ReplaceInCell(inArr(r, c), dict, delims)
            Next c
        Next r
    End If

    ' Fill remaining (if any) with empty strings
    If rowsToProcess < rowsCount Then
        Dim rr As Long, cc As Long
        For rr = rowsToProcess + 1 To rowsCount
            For cc = 1 To colsCount
                outArr(rr, cc) = vbNullString
            Next cc
        Next rr
    End If

    ' Trim to last non-blank when input is a full column
    If colsCount = 1 And inputIsFullColumn Then
        Dim lastRes As Long
        lastRes = LastNonBlankInColumnArray(outArr)
        If lastRes < rowsCount Then
            REPLACE_MANY = Resize2D(outArr, lastRes, 1)
            Exit Function
        End If
    End If

    If rowsCount = 1 And colsCount = 1 Then
        REPLACE_MANY = outArr(1, 1)
    Else
        REPLACE_MANY = outArr
    End If
    Exit Function

CleanFail:
    ' On unexpected error, return the original input
    If rowsCount > 0 And colsCount > 0 Then
        If rowsCount = 1 And colsCount = 1 Then
            REPLACE_MANY = inArr(1, 1)
        Else
            REPLACE_MANY = inArr
        End If
    Else
        REPLACE_MANY = InputRange
    End If
End Function

' -----------------------------------------
' POPUP macro: in-place replace (Selection/Sheet/Workbook)
' -----------------------------------------

Public Sub REPLACE_MANY_POPUP()
    Dim mapRng As Range
    Dim scopeChoice As Variant
    Dim includeFormulas As VbMsgBoxResult
    Dim caseSensitive As VbMsgBoxResult
    Dim dict As Object
    Dim delims As String
    Dim ws As Worksheet
    Dim tgt As Range
    Dim startCalc As XlCalculation
    Dim shCount As Long, shIndex As Long
    Dim targetRng As Range

    On Error GoTo Bail

    ' --- Ask for map range ---
    Set mapRng = PromptForMapRange()
    If mapRng Is Nothing Then Exit Sub
    If mapRng.Columns.Count < 2 Then
        MsgBox "Please select a 2-column map (From, To)", vbExclamation, "REPLACE_MANY"
        Exit Sub
    End If

    ' --- Scope ---
    scopeChoice = Application.InputBox( _
        Prompt:="Scope:" & vbCrLf & _
                "1 = Choose a specific range" & vbCrLf & _
                "2 = ActiveSheet UsedRange" & vbCrLf & _
                "3 = All Sheets in this workbook", _
        Title:="REPLACE_MANY - Scope", Type:=1)
    If scopeChoice = False Then Exit Sub
    If scopeChoice < 1 Or scopeChoice > 3 Then
        MsgBox "Invalid choice. Use 1, 2, or 3.", vbExclamation
        Exit Sub
    End If

    ' --- Include formulas? ---
    includeFormulas = MsgBox( _
        "Include formulas (only formulas whose result is text will be converted to values first)?" & vbCrLf & _
        "Yes = include such formulas" & vbCrLf & _
        "No = skip formulas" & vbCrLf & _
        "Cancel = abort", _
        vbYesNoCancel + vbQuestion, "REPLACE_MANY - Formulas")
    If includeFormulas = vbCancel Then Exit Sub

    ' --- Case sensitivity ---
    caseSensitive = MsgBox("Case-sensitive matching? (Default: No)", vbYesNo + vbQuestion, "REPLACE_MANY - Case")

    ' --- Delimiters & dictionary ---
    delims = DefaultDelims()
    Set dict = BuildDict(mapRng, (caseSensitive = vbYes))
    If dict Is Nothing Or dict.Count = 0 Then
        MsgBox "No valid mapping keys found.", vbInformation
        Exit Sub
    End If

    ' --- Perf guards ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    startCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "REPLACE_MANY: preparing..."

    Select Case CLng(scopeChoice)
        Case 1 ' Ask for target range (instead of using current Selection)
            On Error Resume Next
            Set targetRng = Application.InputBox( _
                Prompt:="Select the target range to process." & vbCrLf & _
                        "Tip: you can select a filtered/visible range too.", _
                Title:="REPLACE_MANY - Target Range", _
                Default:=IIf(TypeName(Selection) = "Range", Selection.Address, ""), Type:=8)
            On Error GoTo 0
            If targetRng Is Nothing Then
                If TypeName(Selection) = "Range" Then
                    Set targetRng = Selection
                Else
                    MsgBox "No range selected. Aborting.", vbExclamation, "REPLACE_MANY"
                    GoTo FinallyClean
                End If
            End If
            Set tgt = targetRng
            ReplaceMany_ApplyToRange tgt, dict, delims, (includeFormulas = vbYes)

        Case 2 ' ActiveSheet
            Set tgt = ActiveSheet.UsedRange
            ReplaceMany_ApplyToRange tgt, dict, delims, (includeFormulas = vbYes)

        Case 3 ' All sheets
            shCount = ThisWorkbook.Worksheets.Count
            For shIndex = 1 To shCount
                Set ws = ThisWorkbook.Worksheets(shIndex)
                Application.StatusBar = "REPLACE_MANY: " & ws.Name & " (" & shIndex & "/" & shCount & ")"
                If Application.WorksheetFunction.CountA(ws.Cells) > 0 Then
                    ReplaceMany_ApplyToRange ws.UsedRange, dict, delims, (includeFormulas = vbYes)
                End If
            Next shIndex
    End Select

    MsgBox "REPLACE_MANY completed.", vbInformation

FinallyClean:
    Application.StatusBar = False
    Application.Calculation = startCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Bail:
    Application.StatusBar = False
    On Error Resume Next
    Application.Calculation = startCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Operation cancelled or failed: " & Err.Description, vbExclamation, "REPLACE_MANY"
End Sub


Private Sub ReplaceMany_ApplyToRange(ByVal rng As Range, ByVal dict As Object, _
                                     ByVal delims As String, ByVal includeFormulas As Boolean)
    Dim r As Range, area As Range
    Dim textConsts As Range, textFormulas As Range

    On Error Resume Next
    If includeFormulas Then
        Set textFormulas = rng.SpecialCells(xlCellTypeFormulas, xlTextValues)
        On Error GoTo 0
        If Not textFormulas Is Nothing Then
            ' Convert only text-result formulas to values
            textFormulas.Value2 = textFormulas.Value2
        End If
    End If

    On Error Resume Next
    Set textConsts = rng.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    If textConsts Is Nothing Then Exit Sub

    ' Process per Area using arrays for speed
    For Each area In textConsts.Areas
        Dim arr As Variant
        Dim i As Long, j As Long
        arr = area.Value2
        If IsArray(arr) Then
            For i = LBound(arr, 1) To UBound(arr, 1)
                For j = LBound(arr, 2) To UBound(arr, 2)
                    arr(i, j) = ReplaceInCell(arr(i, j), dict, delims)
                Next j
            Next i
            area.Value2 = arr
        Else
            ' Single cell area
            area.Value2 = ReplaceInCell(arr, dict, delims)
        End If
    Next area
End Sub

Private Function PromptForMapRange() As Range
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Select the 2-column mapping (From, To). Example: Map!A:B", _
        Title:="REPLACE_MANY - Mapping", Type:=8)
    On Error GoTo 0

    If rng Is Nothing Then
        ' Try default Map!A:B if present
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("Map")
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set rng = Intersect(ws.UsedRange, ws.Columns("A:B"))
        End If
    End If

    Set PromptForMapRange = rng
End Function

Private Function BuildDict(ByVal mapRng As Range, ByVal caseSensitive As Boolean) As Object
    Dim arr As Variant
    Dim dict As Object
    Dim i As Long
    Dim n As Long
    Dim keys() As String, vals() As Variant, lens() As Long

    arr = mapRng.Value2
    If Not IsArray(arr) Then
        ' Single cell or single row
        If mapRng.Columns.Count < 2 Then Exit Function
        ReDim keys(1 To 1)
        ReDim vals(1 To 1)
        ReDim lens(1 To 1)
        If LenB(CStr(mapRng.Cells(1, 1).Value2)) <> 0 Then
            keys(1) = CStr(mapRng.Cells(1, 1).Value2)
            vals(1) = mapRng.Cells(1, 2).Value2
            lens(1) = Len(keys(1))
            n = 1
        End If
    Else
        ReDim keys(1 To UBound(arr, 1))
        ReDim vals(1 To UBound(arr, 1))
        ReDim lens(1 To UBound(arr, 1))
        For i = 1 To UBound(arr, 1)
            Dim k As String
            k = NzString(arr(i, 1))
            If Len(k) > 0 Then
                n = n + 1
                keys(n) = k
                vals(n) = arr(i, 2)
                lens(n) = Len(k)
            End If
        Next i
        If n > 0 Then
            ReDim Preserve keys(1 To n)
            ReDim Preserve vals(1 To n)
            ReDim Preserve lens(1 To n)
        End If
    End If

    If n = 0 Then Exit Function

    ' Sort by length descending
    QuickSortByLength keys, vals, lens, 1, n

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = IIf(caseSensitive, 0, 1) ' 0=vbBinaryCompare, 1=vbTextCompare
    For i = 1 To n
        If Not dict.Exists(keys(i)) Then dict.Add keys(i), vals(i)
    Next i

    Set BuildDict = dict
End Function

' =================
' Shared Helpers
' =================
Private Function DefaultDelims() As String
    ' Common punctuation + whitespace (tab/newline)
    DefaultDelims = " ,.;:!?""'()[]{}<>/\|-_+=*&^%$#@~`" & vbLf & vbTab
End Function

Private Function To2D(ByVal v As Variant) As Variant
    Dim arr As Variant
    If IsObject(v) Then
        If TypeName(v) = "Range" Then
            If v.Cells.Count = 1 Then
                ReDim arr(1 To 1, 1 To 1)
                arr(1, 1) = v.Value2
                To2D = arr
            Else
                arr = v.Value2
                If IsArray(arr) Then
                    To2D = arr
                Else
                    ReDim arr(1 To 1, 1 To 1)
                    arr(1, 1) = v.Value2
                    To2D = arr
                End If
            End If
            Exit Function
        End If
    End If

    If IsArray(v) Then
        On Error GoTo AsScalar
        Dim l1 As Long, u1 As Long, l2 As Long, u2 As Long
        l1 = LBound(v, 1): u1 = UBound(v, 1)
        l2 = LBound(v, 2): u2 = UBound(v, 2)
        To2D = v
        Exit Function
AsScalar:
        ' 1D array -> convert to 2D column
        Dim i As Long
        ReDim arr(1 To UBound(v) - LBound(v) + 1, 1 To 1)
        For i = LBound(v) To UBound(v)
            arr(i - LBound(v) + 1, 1) = v(i)
        Next i
        To2D = arr
        Exit Function
    End If

    ' Scalar -> 1x1 array
    ReDim arr(1 To 1, 1 To 1)
    arr(1, 1) = v
    To2D = arr
End Function

Private Function Resize2D(ByRef arr As Variant, ByVal newRows As Long, ByVal newCols As Long) As Variant
    Dim r As Long, c As Long
    Dim res As Variant
    If newRows < 1 Then
        ReDim res(1 To 1, 1 To newCols)
        For c = 1 To newCols
            res(1, c) = vbNullString
        Next c
        Resize2D = res
        Exit Function
    End If
    ReDim res(1 To newRows, 1 To newCols)
    For r = 1 To newRows
        For c = 1 To newCols
            res(r, c) = arr(r, c)
        Next c
    Next r
    Resize2D = res
End Function

Private Function NzString(ByVal v As Variant) As String
    If IsError(v) Then
        NzString = vbNullString
    ElseIf IsNull(v) Or IsEmpty(v) Then
        NzString = vbNullString
    Else
        NzString = CStr(v)
    End If
End Function

Private Function IsFullColumn(ByVal v As Variant) As Boolean
    On Error GoTo NotRange
    If TypeName(v) = "Range" Then
        Dim rng As Range
        Set rng = v
        IsFullColumn = (rng.Columns.Count = 1 And rng.Rows.Count = rng.Parent.Rows.Count)
        Exit Function
    End If
NotRange:
    IsFullColumn = False
End Function

Private Function LastNonBlankInColumnArray(ByRef arr As Variant) As Long
    Dim r As Long
    For r = UBound(arr, 1) To 1 Step -1
        If Not IsEmpty(arr(r, 1)) Then
            If NzString(arr(r, 1)) <> vbNullString Then
                LastNonBlankInColumnArray = r
                Exit Function
            End If
        End If
    Next r
    LastNonBlankInColumnArray = 0
End Function

Private Function ReplaceInCell(ByVal v As Variant, ByRef dict As Object, ByVal delims As String) As Variant
    If IsError(v) Then
        ReplaceInCell = v
        Exit Function
    End If

    Dim s As String
    s = NzString(v)
    If Len(s) = 0 Then
        ReplaceInCell = vbNullString
        Exit Function
    End If

    Dim i As Long
    Dim ch As String

    ' Pad delimiters with spaces and guard ends
    s = " " & s & " "
    For i = 1 To Len(delims)
        ch = Mid$(delims, i, 1)
        s = Replace$(s, ch, " " & ch & " ")
    Next i

    ' Tokenize on spaces
    Dim toks() As String
    toks = Split(s, " ")

    ' Replace tokens (full-word only)
    Dim out() As String, m As Long
    ReDim out(0 To UBound(toks))
    m = -1
    For i = LBound(toks) To UBound(toks)
        If LenB(toks(i)) <> 0 Then
            m = m + 1
            If dict.Exists(toks(i)) Then
                out(m) = CStr(dict(toks(i)))
            Else
                out(m) = toks(i)
            End If
        End If
    Next i

    If m >= 0 Then
        ReDim Preserve out(0 To m)
        s = Join(out, " ")
    Else
        s = vbNullString
    End If

    ' Remove spaces before delimiters
    For i = 1 To Len(delims)
        ch = Mid$(delims, i, 1)
        s = Replace$(s, " " & ch, ch)
    Next i

    ' Tidy space after opening brackets
    Dim openers As String
    openers = "([{<"
    For i = 1 To Len(openers)
        ch = Mid$(openers, i, 1)
        s = Replace$(s, ch & " ", ch)
    Next i

    ReplaceInCell = Trim$(s)
End Function

' Quicksort helper to sort keys/vals by lens (descending)
Private Sub QuickSortByLength(ByRef keys() As String, _
                              ByRef vals() As Variant, _
                              ByRef lens() As Long, _
                              ByVal lo As Long, _
                              ByVal hi As Long)
    Dim i As Long, j As Long, p As Long
    Dim tmpL As Long, tmpK As String, tmpV As Variant

    i = lo: j = hi
    p = lens((lo + hi) \ 2)

    Do While i <= j
        Do While lens(i) > p: i = i + 1: Loop
        Do While lens(j) < p: j = j - 1: Loop
        If i <= j Then
            tmpL = lens(i): lens(i) = lens(j): lens(j) = tmpL
            tmpK = keys(i): keys(i) = keys(j): keys(j) = tmpK
            tmpV = vals(i): vals(i) = vals(j): vals(j) = tmpV
            i = i + 1: j = j - 1
        End If
    Loop

    If lo < j Then QuickSortByLength keys, vals, lens, lo, j
    If i < hi Then QuickSortByLength keys, vals, lens, i, hi
End Sub
