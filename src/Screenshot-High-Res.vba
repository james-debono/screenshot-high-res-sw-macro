' ===========================================================================
' Export PNG 0.2.0 - Module1
'
' Paste-ready: open Module1 in the VBA editor, select all, paste this in.
' (The "Attribute VB_Name" line is deliberately not here - it is managed by
' VBA and pasting it into the code window is a syntax error.)
' ===========================================================================

Option Explicit

' Tools > Options > Export > PNG is stored in the image capture preferences that
' SOLIDWORKS shares between TIFF / PNG / JPEG, so those are the ones the form drives:
'
'   Screen Capture / Print Capture  ->  swTiffScreenOrPrintCapture
'   DPI                             ->  swTiffPrintDPI
'   Remove Background               ->  "TIF Remove background" (registry - see below)
'
' There is no swUserPreference* constant for Remove Background in any installed
' version of swconst, so that one is read and written directly from the registry
' value SOLIDWORKS keeps it in.
Public Const CAPTURE_SCREEN As Long = 0
Public Const CAPTURE_PRINT As Long = 1

Public Const MIN_DPI As Long = 10
Public Const MAX_DPI As Long = 5000

Private Const REG_REMOVE_BACKGROUND As String = "TIF Remove background"

Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub ShowSaveAsForm()
    ' This sub will display your form
    Load UserForm1
    UserForm1.Show
End Sub

Sub SaveAsPNG(ByVal FN As String, ByVal bScreenCapture As Boolean, _
              ByVal lDPI As Long, ByVal bRemoveBackground As Boolean)
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    If Part Is Nothing Then
        MsgBox "No document is open. Please open a document before running the macro.", vbCritical
        Exit Sub
    End If

    ' Push the form settings into the SOLIDWORKS export options before saving
    ApplyExportSettings swApp, bScreenCapture, lDPI, bRemoveBackground

    Dim myModelView As Object
    Set myModelView = Part.ActiveView
    myModelView.FrameState = swWindowState_e.swWindowMaximized

    ' Construct the save path to the Downloads folder
    Dim sDownloadsPath As String
    sDownloadsPath = Environ("USERPROFILE") & "\Downloads\" & FN & ".PNG"

    ' Save As PNG
    ' The last argument (2) is swSaveAsOptions_e.swSaveAsOptions_Copy, so the
    ' export does not become the active document.
    longstatus = Part.SaveAs3(sDownloadsPath, 0, 2) ' 0 = swSaveAsVersion_e.swSaveAsCurrentVersion

    If longstatus = 0 Then ' If save was successful
        MsgBox "Document successfully saved as: " & sDownloadsPath, vbInformation
    Else
        MsgBox "Failed to save document. Status code: " & longstatus, vbExclamation
    End If

    ' The form will be closed by the CommandButton1_Click event after calling this sub
End Sub

' Writes the form values into the PNG export settings.
Private Sub ApplyExportSettings(ByVal swAppIn As Object, ByVal bScreenCapture As Boolean, _
                                ByVal lDPI As Long, ByVal bRemoveBackground As Boolean)
    Dim lCapture As Long

    If bScreenCapture Then
        lCapture = CAPTURE_SCREEN
    Else
        lCapture = CAPTURE_PRINT
    End If

    swAppIn.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture, lCapture

    ' Only written when the form handed us a usable value - DPI applies to print capture
    If lDPI >= MIN_DPI And lDPI <= MAX_DPI Then
        swAppIn.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swTiffPrintDPI, lDPI
    End If

    SetRemoveBackground swAppIn, bRemoveBackground
End Sub

' Reads the current capture mode. Returns CAPTURE_SCREEN or CAPTURE_PRINT.
Public Function GetCaptureMode(ByVal swAppIn As Object) As Long
    GetCaptureMode = swAppIn.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture)
End Function

' Reads the current export DPI.
Public Function GetExportDPI(ByVal swAppIn As Object) As Long
    GetExportDPI = swAppIn.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swTiffPrintDPI)
End Function

Public Function GetRemoveBackground(ByVal swAppIn As Object) As Boolean
    Dim wsh As Object
    Dim vValue As Variant

    Set wsh = CreateObject("WScript.Shell")

    On Error Resume Next
    vValue = wsh.RegRead(ExportSettingsKey(swAppIn) & REG_REMOVE_BACKGROUND)
    If Err.Number <> 0 Then
        Err.Clear
        GetRemoveBackground = False
    Else
        GetRemoveBackground = (CLng(vValue) <> 0)
    End If
    On Error GoTo 0
End Function

Public Sub SetRemoveBackground(ByVal swAppIn As Object, ByVal bRemoveBackground As Boolean)
    Dim wsh As Object
    Dim lValue As Long

    Set wsh = CreateObject("WScript.Shell")

    If bRemoveBackground Then
        lValue = 1
    Else
        lValue = 0
    End If

    On Error Resume Next
    wsh.RegWrite ExportSettingsKey(swAppIn) & REG_REMOVE_BACKGROUND, lValue, "REG_DWORD"
    Err.Clear
    On Error GoTo 0
End Sub

' Builds the version specific export settings key, e.g.
' "HKEY_CURRENT_USER\Software\SolidWorks\SOLIDWORKS 2025\Export Settings\"
' RevisionNumber returns "33.3.0" for SOLIDWORKS 2025, and the release year is
' always 1992 + the major number.
Private Function ExportSettingsKey(ByVal swAppIn As Object) As String
    Dim vParts As Variant
    Dim lYear As Long

    vParts = Split(swAppIn.RevisionNumber, ".")
    lYear = 1992 + CLng(vParts(0))

    ExportSettingsKey = "HKEY_CURRENT_USER\Software\SolidWorks\SOLIDWORKS " & _
                        CStr(lYear) & "\Export Settings\"
End Function
