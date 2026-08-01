Attribute VB_Name = "Module1"
Option Explicit

' Set True by SetupWorkbook so the setup routines rebuild SILENTLY (no pop-ups,
' no sheet-switching) during an automated Build-Release run. Left False for normal
' manual use, so running a setup macro by hand still shows its confirmation.
Public gBuilding As Boolean

' ============================================================
'  SetupWorkbook - one-shot rebuild used by Build-Release.vbs.
'  Refreshes every sheet, the emblem, and the Developer controls
'  from the current code, forces a compile, and leaves the Labels
'  sheet active. Safe to run anytime to self-heal the workbook.
' ============================================================
Public Sub SetupWorkbook()
    gBuilding = True
    On Error Resume Next
    SetupSheet                      ' Labels layout + emblem + Developer controls
    SetupBatch                      ' Batch tab
    ThisWorkbook.Worksheets("Labels").Activate
    ThisWorkbook.Worksheets("Labels").Range("C3").Select
    On Error GoTo 0
    gBuilding = False
End Sub

Public Sub PrintLabels()
    Dim ws As Worksheet
    Set ws = LabelSheet()

    Dim missing As String
    If Trim(ws.Range("C3").Value & "") = "" Then missing = missing & "  - Last Name" & vbCrLf
    If Trim(ws.Range("C4").Value & "") = "" Then missing = missing & "  - First Name" & vbCrLf
    If Trim(ws.Range("C5").Value & "") = "" Then missing = missing & "  - DOB" & vbCrLf
    If Trim(ws.Range("C6").Value & "") = "" Then missing = missing & "  - Sex" & vbCrLf
    If Trim(ws.Range("C7").Value & "") = "" Then missing = missing & "  - Date" & vbCrLf
    If Len(missing) > 0 Then
        MsgBox "Please fill in these fields first:" & vbCrLf & vbCrLf & missing, vbExclamation, "Missing information"
        Exit Sub
    End If

    ComposeLabel ws

    Dim resp As Variant
    resp = Application.InputBox("How many labels to print for " & ws.Range("C3").Value & ", " & ws.Range("C4").Value & "?", "Number of labels", 1, Type:=1)
    If VarType(resp) = vbBoolean Then Exit Sub
    If Not IsNumeric(resp) Then
        MsgBox "Please enter a whole number.", vbExclamation, "Invalid number"
        Exit Sub
    End If
    Dim qty As Long
    qty = CLng(resp)
    If qty < 1 Then
        MsgBox "Please enter at least 1 label.", vbExclamation, "Invalid number"
        Exit Sub
    End If
    If qty > 50 Then
        If MsgBox(qty & " labels is a lot. Print anyway?", vbYesNo + vbQuestion, "Confirm") = vbNo Then Exit Sub
    End If

    Dim msg As String
    msg = "Print " & qty & " label(s):" & vbCrLf & vbCrLf
    msg = msg & ws.Range("E3").Value & vbCrLf
    msg = msg & ws.Range("E4").Value & vbCrLf
    msg = msg & ws.Range("E5").Value & vbCrLf
    msg = msg & ws.Range("E6").Value & vbCrLf
    msg = msg & ws.Range("E7").Value & vbCrLf & vbCrLf
    msg = msg & "Printer: " & Application.ActivePrinter
    If MsgBox(msg, vbOKCancel + vbInformation, "Confirm print") = vbCancel Then Exit Sub

    On Error GoTo PrintErr
    ws.PrintOut Copies:=qty, Collate:=True
    Exit Sub
PrintErr:
    MsgBox "Could not print. Please check:" & vbCrLf & _
           "  - the label printer is on and connected," & vbCrLf & _
           "  - it is the printer shown in the confirm box," & vbCrLf & _
           "  - the roll is loaded correctly (thermal side up), and" & vbCrLf & _
           "  - the driver's media type / size matches the loaded roll." & vbCrLf & vbCrLf & _
           "Details: " & Err.Description, vbCritical, "Print error"
End Sub

Public Sub ClearForm()
    Dim ws As Worksheet
    Set ws = LabelSheet()
    ws.Range("C3:C6").ClearContents
    ws.Range("C7").Value = Date
    Application.GoTo ws.Range("C3")
End Sub

Private Function LabelSheet() As Worksheet
    On Error Resume Next
    Set LabelSheet = ThisWorkbook.Worksheets("Labels")
    On Error GoTo 0
    If LabelSheet Is Nothing Then Set LabelSheet = ActiveSheet
End Function

Private Sub ComposeLabel(ws As Worksheet)
    ws.Range("E3").Formula = "=IF(C4="""",C3,C3&"", ""&C4)"
    ws.Range("E4").Formula = "=""DOB: ""&IF(ISNUMBER(C5),TEXT(C5,""mm/dd/yyyy""),C5)"
    ws.Range("E5").Formula = "=""Sex: ""&C6"
    ws.Range("E6").Formula = "=""Date: ""&IF(ISNUMBER(C7),TEXT(C7,""m/d/yyyy""),C7)"
    ws.Range("E7").Value = "Saturday Clinic for the Uninsured"
End Sub

Public Sub SetupSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    On Error Resume Next
    ws.Name = "Labels"
    On Error GoTo 0

    Application.ScreenUpdating = False

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 24
    ws.Columns("D").ColumnWidth = 2
    ws.Columns("E").ColumnWidth = 34

    With ws.Range("B1")
        .Value = "Lab Label Printer"
        .Font.Bold = True
        .Font.Size = 14
    End With
    ws.Range("B2").Value = "Fill in the fields, then click PRINT LABELS."

    ws.Range("B3").Value = "Last Name"
    ws.Range("B4").Value = "First Name"
    ws.Range("B5").Value = "DOB"
    ws.Range("B6").Value = "Sex"
    ws.Range("B7").Value = "Date"
    ws.Range("B3:B7").Font.Bold = True

    With ws.Range("C3:C7")
        .Interior.Color = RGB(255, 255, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
    End With
    ws.Range("C7").Value = Date

    With ws.Range("C6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="M,F,Other"
        .InCellDropdown = True
        .ShowError = False
    End With

    With ws.Range("E3:E7")
        .Font.Name = "Arial"
        .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Range("E6").Font.Bold = True
    ws.Range("E2").Value = "Label preview:"
    ws.Range("E2").Font.Italic = True

    ComposeLabel ws

    Dim b As Object
    For Each b In ws.Buttons
        b.Delete
    Next b

    Dim btnP As Object, btnC As Object
    Set btnP = ws.Buttons.Add(ws.Range("B9").Left, ws.Range("B9").Top, 130, 34)
    btnP.OnAction = "PrintLabels"
    btnP.caption = "PRINT LABELS"
    btnP.Font.Bold = True

    Set btnC = ws.Buttons.Add(ws.Range("B9").Left + 140, ws.Range("B9").Top, 80, 34)
    btnC.OnAction = "ClearForm"
    btnC.caption = "Clear"

    With ws.PageSetup
        .PrintArea = "$E$3:$F$7"       ' includes column F for the emblem
        .Orientation = xlLandscape
        .LeftMargin = Application.InchesToPoints(0.06)
        .RightMargin = Application.InchesToPoints(0.06)
        .TopMargin = Application.InchesToPoints(0.06)
        .BottomMargin = Application.InchesToPoints(0.06)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .CenterHorizontally = True
        .CenterVertically = True
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    Application.ScreenUpdating = True

    ' Rebuild the emblem and the Developer-sheet controls too, so SetupSheet is a
    ' complete one-shot rebuild. Wrapped so it never fails if those modules are absent.
    On Error Resume Next
    Application.Run "AddEmblem"
    Application.Run "SetupLabelToggle"
    On Error GoTo 0

    ws.Activate
    If Not gBuilding Then
        ws.Range("C3").Select
        MsgBox "Setup complete.", vbInformation
    End If
End Sub

