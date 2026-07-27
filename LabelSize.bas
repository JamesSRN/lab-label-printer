Attribute VB_Name = "LabelSize"
Option Explicit

'=================================================================
'  LABEL PLACEMENT TOGGLE  -  Lab Label Printer
'
'  IMPORTANT: this ALWAYS prints the lab label onto the SMALL roll.
'  It never prints on a large label. The two buttons just change
'  WHERE the content is placed, to match what the printer is
'  currently detecting - because non-Brother rolls have no spool
'  sensor, so the driver doesn't always recognize the small roll.
'
'   "Small roll"     -> use when the printer CORRECTLY detects the
'                       small roll. The content is centered and fit
'                       to exactly one label (no size mismatch, and
'                       no spilling onto a second label).
'
'   "Large / corner" -> use when the driver is STUCK thinking a big
'                       2.4 x 3.9" label is loaded. We keep the big
'                       page and place the content in the TOP-LEFT
'                       CORNER, so it still lands cleanly on the small
'                       roll that's really loaded (avoids the
'                       "label size mismatch" error).
'
'  The buttons live on a separate "Developer" sheet so the main
'  Labels sheet stays clean for volunteers.
'
'  NOTE: the driver's paper SIZE is set in the printer driver, not
'  here (Excel can't reliably pick the Brother custom size). These
'  buttons control PLACEMENT to suit whichever state the driver is in.
'
'  TO TUNE FOR A DIFFERENT ROLL: edit the constants just below.
'=================================================================

' ---- "Small roll" : printer detects the small roll -------------
Private Const SMALL_TOP    As Double = 0.06    ' inches from top
Private Const SMALL_LEFT   As Double = 0.06    ' inches from left
Private Const SMALL_CENTER As Boolean = True   ' center to fill the small label

' ---- "Large / corner" : driver stuck on the big 2.4x3.9 size ----
Private Const LARGE_TOP    As Double = 0.06    ' top-left corner of the big page
Private Const LARGE_LEFT   As Double = 0.06
Private Const LARGE_CENTER As Boolean = False  ' NOT centered - sits in the corner

' ---- roll dimensions shown in the status line (edit to match your stock) ----
Private Const SMALL_DIMS   As String = "1.1"" x 3.5"""   ' the small lab roll
Private Const LARGE_DIMS   As String = "2.4"" x 3.9"""   ' the big page the driver sticks on

' ---- shared ----------------------------------------------------
Private Const PRINT_AREA   As String = "E3:F7"
Private Const SHEET_NAME   As String = "Labels"       ' the sheet we actually print
Private Const DEV_SHEET    As String = "Developer"    ' where the control buttons live
Private Const STATUS_CELL  As String = "C12"          ' current-mode readout (Developer sheet)

Public Sub UseSmallLabel()
    ' fitToOne = True: squeeze the content onto exactly ONE small label so a
    ' long line (e.g. the clinic name) can't spill onto a second label.
    ApplyProfile SMALL_TOP, SMALL_LEFT, SMALL_CENTER, True
    ShowSize "SMALL roll  " & SMALL_DIMS & "   (printer detects small; centered, fit to 1 label)"
End Sub

Public Sub UseLargeLabel()
    ' No scaling here - the content sits at true size in the top-left corner
    ' of the big 2.4 x 3.9" page and lands on the small roll.
    ApplyProfile LARGE_TOP, LARGE_LEFT, LARGE_CENTER, False
    ShowSize "LARGE page  " & LARGE_DIMS & "   (driver stuck on big; corner placement onto the small roll)"
End Sub

Private Sub ApplyProfile(ByVal topIn As Double, ByVal leftIn As Double, _
                         ByVal centerOnPage As Boolean, ByVal fitToOne As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Application.ScreenUpdating = False
    With ws.PageSetup
        .PrintArea = ws.Range(PRINT_AREA).Address
        .Orientation = xlLandscape
        If fitToOne Then
            ' fit everything onto a single label (prevents overflow to label #2)
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        Else
            .Zoom = 100          ' No Scaling - true size, corner placement
        End If
        .CenterHorizontally = centerOnPage
        .CenterVertically = centerOnPage
        .TopMargin = Application.InchesToPoints(topIn)
        .LeftMargin = Application.InchesToPoints(leftIn)
        .RightMargin = Application.InchesToPoints(0.06)
        .BottomMargin = Application.InchesToPoints(0.06)
        .HeaderMargin = 0
        .FooterMargin = 0
    End With
    Application.ScreenUpdating = True
End Sub

Private Sub ShowSize(ByVal txt As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(DEV_SHEET).Range(STATUS_CELL).Value = txt
    On Error GoTo 0
End Sub

' ================================================================
'  Run ONCE to (re)build the Developer sheet with the two styled
'  placement buttons + status readout.
' ================================================================
Public Sub SetupLabelToggle()
    ' remove any old form-control buttons left on the Labels sheet
    On Error Resume Next
    Dim lab As Worksheet: Set lab = ThisWorkbook.Worksheets(SHEET_NAME)
    Dim s As Shape
    For Each s In lab.Shapes
        If s.Name = "btnSmall" Or s.Name = "btnLarge" Then s.Delete
    Next s
    lab.Range("B11").ClearContents   ' old status text used to live here
    On Error GoTo 0

    Dim dev As Worksheet
    Set dev = EnsureDevSheet()

    ' clear the dev sheet controls + contents, then rebuild
    Dim shp As Shape
    For Each shp In dev.Shapes
        shp.Delete
    Next shp
    dev.Cells.Clear

    dev.Range("B2").Value = "Label Placement - Developer Controls"
    dev.Range("B2").Font.Size = 16
    dev.Range("B2").Font.Bold = True

    dev.Range("B4").Value = "Pick the mode that matches what the printer is detecting."
    dev.Range("B5").Value = "Both modes print the lab label onto the SMALL roll."

    ' two styled buttons
    MakeButton dev, "btnSmall", "Small roll" & vbCr & "printer detects small", _
               "UseSmallLabel", dev.Range("B7").Left, dev.Range("B7").Top, RGB(46, 125, 50)
    MakeButton dev, "btnLarge", "Large / corner" & vbCr & "driver stuck on big", _
               "UseLargeLabel", dev.Range("B7").Left + 240, dev.Range("B7").Top, RGB(21, 101, 192)

    dev.Range("B12").Value = "Current mode:"
    dev.Range("B12").Font.Bold = True
    dev.Range(STATUS_CELL).Font.Italic = True

    UseSmallLabel      ' default to the fit-to-one small mode
    dev.Activate
End Sub

Private Function EnsureDevSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DEV_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = DEV_SHEET
    End If
    Set EnsureDevSheet = ws
End Function

Private Sub MakeButton(ws As Worksheet, ByVal nm As String, ByVal caption As String, _
                       ByVal action As String, ByVal L As Double, ByVal T As Double, _
                       ByVal clr As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, T, 220, 60)
    shp.Name = nm
    shp.Fill.ForeColor.RGB = clr
    shp.Line.Visible = msoFalse
    With shp.TextFrame2.TextRange
        .Text = caption
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Font.Size = 13
        .Font.Bold = msoTrue
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.OnAction = action
End Sub
