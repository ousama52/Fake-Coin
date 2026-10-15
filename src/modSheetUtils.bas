Attribute VB_Name = "modSheetUtils"
Option Explicit

' ---- Cell helpers ----------------------------------------------------------

Public Function CellText(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    CellText = CStr(v)
End Function

' Whole numbers only. Works for real numbers and for numbers stored as text
' (several starter columns are formatted as Text).
Public Function CellLong(ByVal v As Variant) As Long
    Dim s As String

    s = Trim$(CellText(v))
    If Len(s) = 0 Or Len(s) > 10 Then Exit Function
    If s Like "*[!0-9]*" Then Exit Function
    On Error Resume Next
    CellLong = CLng(s)
End Function

' Cell value as a Decimal number (0 when blank or not numeric).
Public Function CellDecimal(ByVal v As Variant) As Variant
    CellDecimal = CDec(0)
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then CellDecimal = CDec(v)
End Function

' True when text is a whole number that fits a Long (1 to 10 digits, at most 2147483647).
' No sign, no separators, no decimals.
Public Function TryParseLong(ByVal text As String, ByRef value As Long) As Boolean
    text = Trim$(text)
    If Len(text) < 1 Or Len(text) > 10 Then Exit Function
    If text Like "*[!0-9]*" Then Exit Function
    ' Ten digits (Unix timestamps) must still fit a Long; same-length text compares in order
    If Len(text) = 10 And text > "2147483647" Then Exit Function
    value = CLng(text)
    TryParseLong = True
End Function

' UTC-style text written into transactions, blocks and logs. Never validated.
Public Function NowText() As String
    NowText = Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss") & "Z"
End Function

' ---- Sheets ----------------------------------------------------------------

Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
End Function

Public Function GetSheet(ByVal sheetName As String) As Worksheet
    If Not SheetExists(sheetName) Then
        Err.Raise modConstants.ERR_APP + 1, "modSheetUtils.GetSheet", _
                  modMessages.FormatText(modMessages.MSG_SHEET_MISSING, sheetName)
    End If
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    If SheetExists(sheetName) Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        Set GetOrCreateSheet = ws
    End If
End Function

Public Function ControlValue(ByVal addr As String) As String
    If Not SheetExists(modConstants.SHEET_CONTROL) Then Exit Function
    ControlValue = Trim$(CellText(GetSheet(modConstants.SHEET_CONTROL).Range(addr).Value))
End Function

' Shows the last run status in Control!C11.
Public Sub SetStatus(ByVal text As String)
    If Not SheetExists(modConstants.SHEET_CONTROL) Then Exit Sub
    GetSheet(modConstants.SHEET_CONTROL).Range(modConstants.CTL_STATUS).Value = _
        "'" & NowText() & "  " & Replace(Left$(text, 250), vbLf, " ")
End Sub

' ---- Tables ----------------------------------------------------------------

Public Function TableExists(ByVal tableName As String) As Boolean
    Dim lo As ListObject

    If Not SheetExists(Mid$(tableName, 4)) Then Exit Function
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(Mid$(tableName, 4)).ListObjects(tableName)
    TableExists = Not lo Is Nothing
End Function

Public Function GetTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    Set ws = GetSheet(Mid$(tableName, 4))
    On Error Resume Next
    Set GetTable = ws.ListObjects(tableName)
    On Error GoTo 0
    If GetTable Is Nothing Then
        Err.Raise modConstants.ERR_APP + 2, "modSheetUtils.GetTable", _
                  modMessages.FormatText(modMessages.MSG_TABLE_MISSING, tableName)
    End If
End Function

' An Excel Table always keeps one row, so "no data" means a single row whose
' key (first) column is empty.
Public Function TableIsEmpty(ByVal lo As ListObject) As Boolean
    If lo.DataBodyRange Is Nothing Then
        TableIsEmpty = True
    ElseIf lo.ListRows.Count = 1 Then
        TableIsEmpty = (Len(CellText(lo.DataBodyRange.Cells(1, 1).Value)) = 0)
    End If
End Function

Public Function RowCount(ByRef data As Variant) As Long
    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function
    RowCount = UBound(data, 1) - LBound(data, 1) + 1
End Function

' Returns a 1-based 2D array, or Empty when the table has no data.
' Callers must test IsEmpty(result) before indexing.
Public Function ReadTable(ByVal lo As ListObject) As Variant
    Dim v As Variant
    Dim one(1 To 1, 1 To 1) As Variant

    If TableIsEmpty(lo) Then Exit Function
    v = lo.DataBodyRange.Value
    If Not IsArray(v) Then
        one(1, 1) = v
        ReadTable = one
    Else
        ReadTable = v
    End If
End Function

' Replaces the whole table body with data in one assignment.
' Pass Empty to clear the table (it keeps one blank row).
Public Sub WriteTable(ByVal lo As ListObject, ByRef data As Variant)
    Dim n As Long

    n = RowCount(data)
    CheckColumns lo, data
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    ' Values go in first, by position below the header: after a resize Excel can
    ' report DataBodyRange as Nothing for a table with a single (blank) row.
    If n > 0 Then BodyCells(lo, 0, n).Value = data
    ResizeBody lo, n
End Sub

' Adds rows below the existing ones in one assignment. When the table only
' holds its blank row, the new rows overwrite it.
Public Sub AppendRows(ByVal lo As ListObject, ByRef data As Variant)
    Dim n As Long
    Dim oldRows As Long

    n = RowCount(data)
    If n = 0 Then Exit Sub
    CheckColumns lo, data
    If Not TableIsEmpty(lo) Then oldRows = lo.ListRows.Count
    BodyCells(lo, oldRows, n).Value = data
    ResizeBody lo, oldRows + n
End Sub

Private Sub CheckColumns(ByVal lo As ListObject, ByRef data As Variant)
    Dim c As Long

    If RowCount(data) = 0 Then Exit Sub
    c = UBound(data, 2) - LBound(data, 2) + 1
    If c <> lo.ListColumns.Count Then
        Err.Raise modConstants.ERR_APP + 3, "modSheetUtils.CheckColumns", _
                  modMessages.FormatText(modMessages.MSG_COLUMN_MISMATCH, c, lo.Name, lo.ListColumns.Count)
    End If
End Sub

' n rows of the table body starting after skipRows rows, found from the header.
Private Function BodyCells(ByVal lo As ListObject, ByVal skipRows As Long, ByVal n As Long) As Range
    Set BodyCells = lo.HeaderRowRange.Offset(1 + skipRows).Resize(n, lo.ListColumns.Count)
End Function

Private Sub ResizeBody(ByVal lo As ListObject, ByVal n As Long)
    If n < 1 Then n = 1
    lo.Resize lo.HeaderRowRange.Resize(n + 1)
End Sub

' Position of a header in the table, 0 when the column is missing.
Public Function ColIndex(ByVal lo As ListObject, ByVal header As String) As Long
    On Error Resume Next
    ColIndex = lo.ListColumns(header).Index
End Function

' Creates the table on ws if missing. An existing table is returned as is.
Public Function EnsureTable(ByVal ws As Worksheet, ByVal tableName As String, _
                            ByVal headers As Variant) As ListObject
    Dim lo As ListObject
    Dim rng As Range
    Dim n As Long

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If lo Is Nothing Then
        n = UBound(headers) - LBound(headers) + 1
        Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(1, n))
        If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
            ' Text format first so Excel never reinterprets a header
            rng.NumberFormat = "@"
            rng.Value = headers
        Else
            Set rng = ws.Range("A1").CurrentRegion
        End If
        Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        lo.Name = tableName
        lo.TableStyle = "TableStyleMedium2"
    End If
    Set EnsureTable = lo
End Function

' Appends headers the table lacks. Returns how many were added.
Public Function AddMissingColumns(ByVal lo As ListObject, ByVal headers As Variant) As Long
    Dim i As Long
    Dim col As ListColumn

    For i = LBound(headers) To UBound(headers)
        If ColIndex(lo, CStr(headers(i))) = 0 Then
            Set col = lo.ListColumns.Add
            col.Name = CStr(headers(i))
            AddMissingColumns = AddMissingColumns + 1
        End If
    Next i
End Function

' ---- Formatting ------------------------------------------------------------

' Whole-column format so rows added later inherit it.
Public Sub SetColumnFormat(ByVal lo As ListObject, ByVal idx As Long, ByVal fmt As String)
    If Len(fmt) = 0 Then Exit Sub
    lo.Parent.Columns(lo.ListColumns(idx).Range.Column).NumberFormat = fmt
End Sub

' The cells below a table header in one column, whether or not the table has
' data rows. A table with no rows has no DataBodyRange (it is Nothing).
Public Function ColumnBelowHeader(ByVal lo As ListObject, ByVal idx As Long) As Range
    Dim ws As Worksheet
    Dim c As Long

    Set ws = lo.Parent
    c = lo.ListColumns(idx).Range.Column
    Set ColumnBelowHeader = ws.Range(ws.Cells(lo.HeaderRowRange.Row + 1, c), ws.Cells(ws.Rows.Count, c))
End Function

Public Sub AddListValidation(ByVal rng As Range, ByVal csvList As String)
    ' The list literal must use the regional list separator
    csvList = Replace(csvList, ",", Application.International(xlListSeparator))
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=csvList
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

' ---- Array helpers ---------------------------------------------------------

' Turns Array(Array(...), Array(...)) into a 1-based 2D array. Inner arrays
' must be 0-based (the VBA Array() default) and equally long.
Public Function JaggedToTable(ByVal jagged As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim nR As Long
    Dim nC As Long
    Dim result() As Variant

    nR = UBound(jagged) + 1
    nC = UBound(jagged(0)) + 1
    ReDim result(1 To nR, 1 To nC)
    For r = 1 To nR
        For c = 1 To nC
            result(r, c) = jagged(r - 1)(c - 1)
        Next c
    Next r
    JaggedToTable = result
End Function

' First n rows of a 2D array as a new exact-size array (Empty when n = 0).
' ReDim Preserve cannot shrink the first dimension, so this copies.
Public Function TakeRows(ByRef data As Variant, ByVal n As Long) As Variant
    Dim r As Long
    Dim c As Long
    Dim nC As Long
    Dim result() As Variant

    If n < 1 Then Exit Function
    nC = UBound(data, 2)
    ReDim result(1 To n, 1 To nC)
    For r = 1 To n
        For c = 1 To nC
            result(r, c) = data(r, c)
        Next c
    Next r
    TakeRows = result
End Function
