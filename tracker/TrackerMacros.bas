Attribute VB_Name = "TrackerMacros"
' Activity & Weight Tracker - macros
' Import this module into the workbook (Alt+F11 > File > Import File...),
' save the workbook as .xlsm, then run SetupButtons once (Alt+F8).
Option Explicit

Private Const ACT_SHEET As String = "Activities"
Private Const WGT_SHEET As String = "Weight"
Private Const DASH_SHEET As String = "Dashboard"
Private Const LAST_DATA_ROW As Long = 1000   ' formulas are prefilled down to this row

' ---------------------------------------------------------------------------
' Add one physical activity to the Activities sheet
' ---------------------------------------------------------------------------
Public Sub AddActivity()
    Dim ws As Worksheet
    Dim d As Variant, act As String, dur As Variant, dist As Variant, notes As String
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(ACT_SHEET)

    d = AskDate("Date of the activity", Date)
    If IsEmpty(d) Then Exit Sub

    act = AskActivity()
    If act = "" Then Exit Sub

    dur = AskNumber("Duration in minutes (e.g. 45)", 1, 1440, False)
    If IsEmpty(dur) Then Exit Sub

    dist = AskNumber("Distance in km (e.g. 8,5) - leave empty if not relevant", 0, 1000, True)
    If IsNull(dist) Then Exit Sub   ' Cancel pressed

    notes = InputBox("Notes (optional)", "Add activity")

    r = NextRow(ws)
    If r > LAST_DATA_ROW Then
        MsgBox "The Activities sheet is full (row " & LAST_DATA_ROW & ").", vbExclamation
        Exit Sub
    End If

    ws.Cells(r, 1).Value = CDate(d)
    ws.Cells(r, 2).Value = act
    ws.Cells(r, 3).Value = dur
    If Not IsEmpty(dist) Then ws.Cells(r, 4).Value = dist
    ws.Cells(r, 5).Value = notes

    SortByDate ws, 5
    MsgBox "Activity added: " & act & ", " & dur & " min on " & Format(d, "Short Date") & ".", vbInformation
End Sub

' ---------------------------------------------------------------------------
' Add (or replace) a weight measurement on the Weight sheet
' ---------------------------------------------------------------------------
Public Sub AddWeight()
    Dim ws As Worksheet
    Dim d As Variant, w As Variant
    Dim r As Long, found As Range

    Set ws = ThisWorkbook.Worksheets(WGT_SHEET)

    d = AskDate("Date of the measurement", Date)
    If IsEmpty(d) Then Exit Sub

    w = AskNumber("Weight in kg (e.g. 72,4)", 20, 400, False)
    If IsEmpty(w) Then Exit Sub

    Set found = FindDate(ws, CDate(d))
    If Not found Is Nothing Then
        If MsgBox("A weight already exists for " & Format(d, "Short Date") & " (" & _
                  found.Offset(0, 1).Value & " kg)." & vbCrLf & "Replace it?", _
                  vbYesNo + vbQuestion) = vbNo Then Exit Sub
        found.Offset(0, 1).Value = w
    Else
        r = NextRow(ws)
        If r > LAST_DATA_ROW Then
            MsgBox "The Weight sheet is full (row " & LAST_DATA_ROW & ").", vbExclamation
            Exit Sub
        End If
        ws.Cells(r, 1).Value = CDate(d)
        ws.Cells(r, 2).Value = w
        SortByDate ws, 2
    End If

    MsgBox "Weight saved: " & w & " kg on " & Format(d, "Short Date") & ".", vbInformation
End Sub

' ---------------------------------------------------------------------------
' Put "Add activity" / "Add weight" buttons on the Dashboard (run once)
' ---------------------------------------------------------------------------
Public Sub SetupButtons()
    Dim ws As Worksheet, b As Object, cell As Range

    Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    On Error Resume Next
    ws.Buttons("btnAddActivity").Delete
    ws.Buttons("btnAddWeight").Delete
    On Error GoTo 0

    Set cell = ws.Range("H2")
    Set b = ws.Buttons.Add(cell.Left, cell.Top, 130, 28)
    b.Name = "btnAddActivity"
    b.Caption = "+ Add activity"
    b.OnAction = "AddActivity"

    Set b = ws.Buttons.Add(cell.Left + 140, cell.Top, 130, 28)
    b.Name = "btnAddWeight"
    b.Caption = "+ Add weight"
    b.OnAction = "AddWeight"

    MsgBox "Buttons added to the Dashboard.", vbInformation
End Sub

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------
Private Function NextRow(ws As Worksheet) As Long
    NextRow = ws.Cells(LAST_DATA_ROW + 1, 1).End(xlUp).Row + 1
    If NextRow < 2 Then NextRow = 2
End Function

Private Sub SortByDate(ws As Worksheet, lastInputCol As Long)
    Dim lastRow As Long
    lastRow = NextRow(ws) - 1
    If lastRow < 3 Then Exit Sub
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastInputCol)).Sort _
        Key1:=ws.Cells(2, 1), Order1:=xlAscending, Header:=xlNo
End Sub

Private Function FindDate(ws As Worksheet, d As Date) As Range
    Dim r As Long
    For r = 2 To NextRow(ws) - 1
        If IsDate(ws.Cells(r, 1).Value) Then
            If CLng(ws.Cells(r, 1).Value) = CLng(d) Then
                Set FindDate = ws.Cells(r, 1)
                Exit Function
            End If
        End If
    Next r
End Function

' Returns a Date, or Empty if the user cancels.
Private Function AskDate(prompt As String, defaultDate As Date) As Variant
    Dim s As String
    Do
        s = InputBox(prompt & " (" & Format(defaultDate, "Short Date") & " = today)", _
                     "Tracker", Format(defaultDate, "Short Date"))
        If StrPtr(s) = 0 Or Trim$(s) = "" Then AskDate = Empty: Exit Function
        If IsDate(s) Then AskDate = CDate(s): Exit Function
        MsgBox "'" & s & "' is not a valid date.", vbExclamation
    Loop
End Function

' Returns a number, Empty (blank answer, only if allowBlank) or Null (Cancel).
' When allowBlank is False, Cancel or blank both return Empty.
Private Function AskNumber(prompt As String, minV As Double, maxV As Double, _
                           allowBlank As Boolean) As Variant
    Dim s As String, v As Double
    Do
        s = InputBox(prompt, "Tracker")
        If StrPtr(s) = 0 Then
            If allowBlank Then AskNumber = Null Else AskNumber = Empty
            Exit Function
        End If
        s = Trim$(Replace(s, ",", "."))   ' accept both 72,5 and 72.5
        If s = "" Then AskNumber = Empty: Exit Function
        If IsNumeric(Replace(s, ".", Application.DecimalSeparator)) Then
            v = Val(s)
            If v >= minV And v <= maxV Then AskNumber = v: Exit Function
        End If
        MsgBox "Please enter a number between " & minV & " and " & maxV & ".", vbExclamation
    Loop
End Function

' Shows the list from Settings and returns the chosen activity ("" if cancelled).
Private Function AskActivity() As String
    Dim list As Range, c As Range, msg As String, s As String, i As Long, n As Long

    Set list = ThisWorkbook.Worksheets("Settings").Range("ActivityTable").Columns(1)
    For Each c In list.Cells
        If Trim$(c.Value) <> "" Then
            n = n + 1
            msg = msg & n & " - " & c.Value & vbCrLf
        End If
    Next c

    Do
        s = InputBox("Type the number of the activity:" & vbCrLf & vbCrLf & msg, "Add activity")
        If StrPtr(s) = 0 Or Trim$(s) = "" Then AskActivity = "": Exit Function
        If IsNumeric(s) Then
            i = 0
            For Each c In list.Cells
                If Trim$(c.Value) <> "" Then
                    i = i + 1
                    If i = Val(s) Then AskActivity = c.Value: Exit Function
                End If
            Next c
        Else
            For Each c In list.Cells
                If StrComp(Trim$(c.Value), Trim$(s), vbTextCompare) = 0 Then
                    AskActivity = c.Value: Exit Function
                End If
            Next c
        End If
        MsgBox "Unknown activity. Type a number from 1 to " & n & ".", vbExclamation
    Loop
End Function
