' ============================================================
'  Build-Release.vbs  (Lab Label Printer)  -  robust: never leaves
'  a stray Excel process.
'  Produces the definitive "Lab Label Printer.xlsm":
'    1. opens Lab Label Printer.xlsm
'    2. imports the newest Module1/BatchPrint/EmblemSetup/LabelSize .bas
'    3. runs SetupWorkbook (compiles the project + rebuilds the sheets,
'       emblem, and Developer buttons)
'    4. saves the workbook and LEAVES IT OPEN for immediate use
'  On ANY failure it shows a message and quits Excel cleanly, so no
'  background Excel process is left holding the file.
'
'  RUN: close Lab Label Printer.xlsm first, then double-click this
'  (or the "OPEN LAB TOOL (double-click me).cmd" launcher).
'  One-time on each PC: Excel > File > Options > Trust Center > Trust
'  Center Settings > Macro Settings > check "Trust access to the VBA
'  project object model".
' ============================================================
Option Explicit

Dim fso, scriptDir, xlsmPath
Dim xl, wb, proj, macroName, setupErr, saveErr
Dim modNames, basFiles, i

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
xlsmPath = fso.BuildPath(scriptDir, "Lab Label Printer.xlsm")

' The four code modules that are rebuilt from source on every launch.
modNames = Array("Module1", "BatchPrint", "EmblemSetup", "LabelSize")
basFiles = Array("Module1.bas", "BatchPrint.bas", "EmblemSetup.bas", "LabelSize.bas")

If Not fso.FileExists(xlsmPath) Then
    MsgBox "Cannot find the workbook:" & vbCrLf & xlsmPath & vbCrLf & vbCrLf & _
        "Make sure this script sits in the same folder as Lab Label Printer.xlsm.", _
        vbCritical, "Build-Release"
    WScript.Quit 1
End If
For i = 0 To UBound(basFiles)
    If Not fso.FileExists(fso.BuildPath(scriptDir, basFiles(i))) Then
        MsgBox "Cannot find source file:" & vbCrLf & basFiles(i), vbCritical, "Build-Release"
        WScript.Quit 1
    End If
Next

Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
xl.AutomationSecurity = 1
Prog 10, "Starting Excel..."

Prog 25, "Opening workbook..."
On Error Resume Next
Set wb = xl.Workbooks.Open(xlsmPath, 0, False)
If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    MsgBox "Could not open Lab Label Printer.xlsm (it may be open elsewhere or locked)." & vbCrLf & vbCrLf & _
        "Close it and end any stray Microsoft Excel in Task Manager, then run again.", _
        vbExclamation, "Build-Release"
    CleanupQuit
    WScript.Quit 1
End If
On Error GoTo 0
wb.Activate

' Guard: read-only means it is locked/open elsewhere - cannot save. Abort cleanly.
If wb.ReadOnly Then
    MsgBox "Lab Label Printer.xlsm opened READ-ONLY, so it cannot be saved." & vbCrLf & vbCrLf & _
        "It is most likely still open in another Excel window, or a stray Excel process " & _
        "is holding it (end them in Task Manager), or the file is marked read-only " & _
        "(right-click > Properties > uncheck Read-only)." & vbCrLf & vbCrLf & _
        "Nothing was changed. Excel has been closed. Fix that and run again.", _
        vbExclamation, "Build-Release"
    CleanupQuit
    WScript.Quit 1
End If

' Verify programmatic access to the VBA project.
On Error Resume Next
Set proj = wb.VBProject
If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    MsgBox "Cannot access the VBA project." & vbCrLf & vbCrLf & _
        "Enable it once in Excel:" & vbCrLf & _
        "File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
        "Macro Settings > check 'Trust access to the VBA project object model'," & vbCrLf & _
        "click OK, then run this script again.", vbExclamation, "Build-Release"
    CleanupQuit
    WScript.Quit 1
End If
On Error GoTo 0

' Remove each existing module (and any accidental "Name1" duplicate), then import fresh.
Prog 55, "Importing the latest code..."
For i = 0 To UBound(modNames)
    On Error Resume Next
    proj.VBComponents.Remove proj.VBComponents(modNames(i))
    proj.VBComponents.Remove proj.VBComponents(modNames(i) & "1")
    Err.Clear
    On Error GoTo 0
    proj.VBComponents.Import fso.BuildPath(scriptDir, basFiles(i))
Next

' Run SetupWorkbook (forces a full compile + rebuilds sheets/emblem/buttons).
Prog 85, "Compiling + rebuilding the label tool..."
macroName = "'" & wb.Name & "'!SetupWorkbook"
On Error Resume Next
xl.Run macroName
If Err.Number <> 0 Then
    setupErr = Err.Description
    Err.Clear
    On Error GoTo 0
    MsgBox "SetupWorkbook failed:" & vbCrLf & vbCrLf & setupErr & vbCrLf & vbCrLf & _
        "Excel is LEFT OPEN so you can inspect it: press Alt+F11, then Debug > Compile " & _
        "VBAProject to see the exact error and line. Nothing was saved.", _
        vbCritical, "Build-Release"
    WScript.Quit 1
End If
On Error GoTo 0

' Save; on failure clean up so no Excel process is left behind.
Prog 95, "Saving Lab Label Printer.xlsm..."
On Error Resume Next
wb.Save
If Err.Number <> 0 Then
    saveErr = Err.Description
    Err.Clear
    On Error GoTo 0
    MsgBox "Could not save Lab Label Printer.xlsm:" & vbCrLf & vbCrLf & saveErr & vbCrLf & vbCrLf & _
        "The file is open elsewhere or read-only. Close it (and end stray Excel in " & _
        "Task Manager), then run again. Nothing was saved; Excel has been closed cleanly.", _
        vbCritical, "Build-Release"
    CleanupQuit
    WScript.Quit 1
End If
On Error GoTo 0

' Success: workbook is saved and LEFT OPEN for immediate use.
Prog 100, "Build complete."
On Error Resume Next
xl.DisplayAlerts = True
xl.StatusBar = False
wb.Activate
On Error GoTo 0

MsgBox "Lab Label Printer is ready." & vbCrLf & vbCrLf & _
    "The workbook is open on the Labels tab - go ahead and use it." & vbCrLf & _
    "(The Developer tab has the Small/Large placement buttons.)" & vbCrLf & vbCrLf & _
    "When you are done for the day, just close it.", vbInformation, "Lab Label Printer"

' -------- helpers --------
Sub Prog(pct, msg)
    On Error Resume Next
    Dim total, filled, ii, bar
    total = 22
    filled = Int(total * pct / 100)
    bar = ""
    For ii = 1 To total
        If ii <= filled Then
            bar = bar & ChrW(9608)
        Else
            bar = bar & ChrW(9618)
        End If
    Next
    xl.StatusBar = "Building Lab Label Printer.xlsm   [" & bar & "]  " & pct & "%   -   " & msg
    On Error GoTo 0
End Sub

Sub CleanupQuit
    On Error Resume Next
    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set proj = Nothing
    Set xl = Nothing
    On Error GoTo 0
End Sub
