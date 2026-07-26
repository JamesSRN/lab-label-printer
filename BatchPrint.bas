Attribute VB_Name = "BatchPrint"
Option Explicit

Public Sub PrintBatch()
    Dim wsB As Worksheet, wsL As Worksheet
    Set wsB = ThisWorkbook.Worksheets("Batch")
    Set wsL = ThisWorkbook.Worksheets("Labels")

    Const firstRow As Long = 6
    Dim lastRow As Long
    lastRow = wsB.Cells(wsB.Rows.Count, "B").End(xlUp).Row
    If lastRow > 105 Then lastRow = 105
    If lastRow < firstRow Then
        MsgBox "The list is empty. Add at least one patient row.", vbExclamation, "Nothing to print"
        Exit Sub
    End If

    Dim r As Long, totalLabels As Long, patientCount As Long, problems As String
    Dim ln As String, fn As String, dob As String, sx As String
    Dim qtyV As Variant, qty As Long
    For r = firstRow To lastRow
        ln = Trim(wsB.Cells(r, "B").Value & "")
        fn = Trim(wsB.Cells(r, "C").Value & "")
        dob = Trim(wsB.Cells(r, "D").Value & "")
        sx = Trim(wsB.Cells(r, "E").Value & "")
        If ln = "" And fn = "" And dob = "" And sx = "" Then GoTo NextCount
        If ln = "" Or fn = "" Or dob = "" Or sx = "" Then
            problems = problems & "  Row " & r & ": missing name / DOB / sex" & vbCrLf
            GoTo NextCount
        End If
        qtyV = wsB.Cells(r, "G").Value
        If Trim(qtyV & "") = "" Then
            qty = 1
        ElseIf IsNumeric(qtyV) Then
            qty = CLng(qtyV)
        Else
            qty = 0
        End If
        If qty < 1 Then
            problems = problems & "  Row " & r & ": invalid # labels" & vbCrLf
            GoTo NextCount
        End If
        totalLabels = totalLabels + qty
        patientCount = patientCount + 1
NextCount:
    Next r

    If patientCount = 0 Then
        MsgBox "No valid rows to print." & vbCrLf & vbCrLf & problems, vbExclamation, "Nothing to print"
        Exit Sub
    End If

    Dim msg As String
    msg = "Print " & totalLabels & " label(s) for " & patientCount & " patient(s)?"
    If Len(problems) > 0 Then msg = msg & vbCrLf & vbCrLf & "These rows will be skipped:" & vbCrLf & problems
    msg = msg & vbCrLf & "Printer: " & Application.ActivePrinter
    If MsgBox(msg, vbOKCancel + vbInformation, "Confirm batch print") = vbCancel Then Exit Sub

    Dim s3, s4, s5, s6, s7
    s3 = wsL.Range("C3").Value: s4 = wsL.Range("C4").Value: s5 = wsL.Range("C5").Value
    s6 = wsL.Range("C6").Value: s7 = wsL.Range("C7").Value

    On Error GoTo PrintErr
    Application.ScreenUpdating = False
    For r = firstRow To lastRow
        ln = Trim(wsB.Cells(r, "B").Value & "")
        fn = Trim(wsB.Cells(r, "C").Value & "")
        dob = Trim(wsB.Cells(r, "D").Value & "")
        sx = Trim(wsB.Cells(r, "E").Value & "")
        If ln = "" Or fn = "" Or dob = "" Or sx = "" Then GoTo NextPrint
        qtyV = wsB.Cells(r, "G").Value
        If Trim(qtyV & "") = "" Then
            qty = 1
        ElseIf IsNumeric(qtyV) Then
            qty = CLng(qtyV)
        Else
            qty = 0
        End If
        If qty < 1 Then GoTo NextPrint

        wsL.Range("C3").Value = ln
        wsL.Range("C4").Value = fn
        wsL.Range("C5").Value = wsB.Cells(r, "D").Value
        wsL.Range("C6").Value = sx
        If Trim(wsB.Cells(r, "F").Value & "") = "" Then
            wsL.Range("C7").Value = Date
        Else
            wsL.Range("C7").Value = wsB.Cells(r, "F").Value
        End If

        wsL.PrintOut Copies:=qty, Collate:=True
NextPrint:
    Next r
    Application.ScreenUpdating = True

    wsL.Range("C3").Value = s3: wsL.Range("C4").Value = s4: wsL.Range("C5").Value = s5
    wsL.Range("C6").Value = s6: wsL.Range("C7").Value = s7

    MsgBox "Done - sent " & totalLabels & " label(s) for " & patientCount & " patient(s).", vbInformation, "Batch complete"
    Exit Sub
PrintErr:
    Application.ScreenUpdating = True
    wsL.Range("C3").Value = s3: wsL.Range("C4").Value = s4: wsL.Range("C5").Value = s5
    wsL.Range("C6").Value = s6: wsL.Range("C7").Value = s7
    MsgBox "Printing stopped due to an error." & vbCrLf & Err.Description, vbCritical, "Print error"
End Sub

Public Sub ClearBatch()
    ThisWorkbook.Worksheets("Batch").Range("B6:G105").ClearContents
End Sub

Public Sub SetupBatch()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Batch")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Batch"
    End If

    Application.ScreenUpdating = False
    ws.Cells.Clear
    Dim b As Object
    For Each b In ws.Buttons: b.Delete: Next b

    With ws.Range("B1")
        .Value = "Batch Label Printing"
        .Font.Bold = True
        .Font.Size = 14
    End With
    ws.Range("B2").Value = "Type or paste one patient per row, then click PRINT ALL LABELS."

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 8
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 10

    ws.Range("B5").Value = "Last Name"
    ws.Range("C5").Value = "First Name"
    ws.Range("D5").Value = "DOB"
    ws.Range("E5").Value = "Sex"
    ws.Range("F5").Value = "Date"
    ws.Range("G5").Value = "# Labels"
    With ws.Range("B5:G5")
        .Font.Bold = True
        .Interior.Color = RGB(210, 210, 210)
        .Borders.LineStyle = xlContinuous
    End With

    With ws.Range("B6:G105")
        .Interior.Color = RGB(255, 255, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With

    With ws.Range("E6:E105").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="M,F,Other"
        .InCellDropdown = True
        .ShowError = False
    End With

    ws.Range("B107").Value = "Date blank = today.   # Labels blank = 1."
    ws.Range("B107").Font.Italic = True

    Dim btnP As Object, btnC As Object
    Set btnP = ws.Buttons.Add(ws.Range("B3").Left, ws.Range("B3").Top, 150, 26)
    btnP.OnAction = "PrintBatch"
    btnP.Caption = "PRINT ALL LABELS"
    btnP.Font.Bold = True

    Set btnC = ws.Buttons.Add(ws.Range("B3").Left + 160, ws.Range("B3").Top, 90, 26)
    btnC.OnAction = "ClearBatch"
    btnC.Caption = "Clear List"

    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("B6").Select
    MsgBox "Batch tab ready.", vbInformation
End Sub
