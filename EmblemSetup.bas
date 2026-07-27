Attribute VB_Name = "EmblemSetup"
Option Explicit

' One-time: inserts the SCU emblem on the right side of the label and
' extends the print area to include it. The image is embedded in the
' workbook, so it travels with the file even after copying elsewhere.
Public Sub AddEmblem()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Labels")

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
    ' shift right so the emblem sits toward column F (print area extended below)
    L = ws.Range("E7").Left + ws.Range("E7").Width - w + 14
    T = rng.Top + (rng.Height - h) / 2                       ' centered over rows 3-6

    Dim imgPath As String
    imgPath = "C:\Users\ringo\OneDrive\Documents\GitHub\lab-label-printer\scu_emblem.png"

    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture( _
        imgPath, _
        msoFalse, msoTrue, L, T, w, h)
    pic.Name = "SCUEmblem"
    pic.Placement = xlMoveAndSize

    ' include column F so the shifted emblem still prints
    ws.PageSetup.PrintArea = "$E$3:$F$7"
End Sub
