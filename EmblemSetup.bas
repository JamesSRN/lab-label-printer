Attribute VB_Name = "EmblemSetup"
Option Explicit

' One-time: inserts the SCU emblem on the right side of the label and
' extends the print area to include it. The image is embedded in the
' workbook, so it travels with the file even after copying elsewhere.
Public Sub AddEmblem()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Labels")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' Resolve the emblem PNG FIRST. If we can't find it, leave the emblem that is
    ' already embedded in the workbook untouched (do NOT delete it).
    Dim imgPath As String
    imgPath = ResolveEmblemPath()
    If imgPath = "" Then Exit Sub

    On Error Resume Next
    ws.Shapes("SCUEmblem").Delete
    On Error GoTo 0

    ' The printable width is capped, so the emblem sits inside the text
    ' area on the right side (upper rows, clear of the long clinic line).
    ws.Columns("F").ColumnWidth = 8.43

    Dim rng As Range
    Set rng = ws.Range("E3:E6")

    Dim w As Single, h As Single
    h = 46                      ' ~0.64" tall (bigger)
    w = h * 1024 / 683          ' keep the emblem's native aspect ratio

    Dim L As Single, T As Single
    ' Push the emblem to the right edge of the print area (column F), away from the text
    ' column, so a longer patient name on the top line no longer overlaps it. The small
    ' pad keeps it just off the very edge.
    L = ws.Range("F7").Left + ws.Range("F7").Width - w - 2
    T = rng.Top + (rng.Height - h) / 2                       ' centered over rows 3-6

    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture( _
        imgPath, _
        msoFalse, msoTrue, L, T, w, h)
    pic.Name = "SCUEmblem"
    pic.Placement = xlMoveAndSize

    ' include column F so the shifted emblem still prints
    ws.PageSetup.PrintArea = "$E$3:$F$7"
End Sub

' Find scu_emblem.png robustly. Prefer the copy next to the workbook, but ONLY when
' that is a real LOCAL path - OneDrive can report an https:// path that Dir() cannot
' read (it raises run-time error 52 "Bad file name or number"), so every Dir() call is
' guarded with On Error. Returns "" if the file cannot be located, in which case the
' caller keeps the emblem already embedded in the workbook.
Private Function ResolveEmblemPath() As String
    Dim p As String, wbPath As String
    ResolveEmblemPath = ""
    On Error Resume Next
    wbPath = ThisWorkbook.Path
    If Len(wbPath) > 0 And InStr(wbPath, "://") = 0 Then
        p = wbPath & "\scu_emblem.png"
        If Dir(p) <> "" Then
            ResolveEmblemPath = p
            On Error GoTo 0
            Exit Function
        End If
    End If
    ' Fallback: the known repo location on the clinic PC.
    p = "C:\Users\ringo\OneDrive\Documents\GitHub\lab-label-printer\scu_emblem.png"
    If Dir(p) <> "" Then ResolveEmblemPath = p
    On Error GoTo 0
End Function
