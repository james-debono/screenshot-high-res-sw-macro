' ===========================================================================
' Export PNG 0.2.1 - Module1
'
' Paste-ready: open Module1 in the VBA editor, select all, paste this in.
' (The "Attribute VB_Name" line is deliberately not here - it is managed by
' VBA and pasting it into the code window is a syntax error.)
' ===========================================================================

Option Explicit

' The form is a shortcut to Tools > Options > System Options > Export >
' TIF/PSD/JPG/PNG. Mapping is per the SOLIDWORKS API help topic
' "System Options > Export > TIF/PSD/JPG/PNG":
'
'   Screen Capture / Print Capture -> swTiffScreenOrPrintCapture (0 / 1)
'   DPI                            -> swTiffPrintDPI
'                                     ("For drawings only; For Output as -
'                                      Print capture only")
'   Remove Background              -> no API. The help topic lists this one as
'                                     "Currently not available in the SOLIDWORKS
'                                     API", so it is read and written directly
'                                     from the registry value SOLIDWORKS keeps
'                                     it in. See SetRemoveBackground.
'
' Every setting is applied the moment it is changed on the form, so closing the
' form without exporting still leaves the settings changed.
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

' Returns True if the file was exported, False if it failed or the user backed
' out of overwriting an existing file, so the form can stay open.
Function SaveAsPNG(ByVal FN As String) As Boolean
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    If Part Is Nothing Then
        MsgBox "No document is open. Please open a document before running the macro.", vbCritical
        Exit Function
    End If

    ' Construct the save path to the Downloads folder
    Dim sDownloadsPath As String
    sDownloadsPath = Environ("USERPROFILE") & "\Downloads\" & FN & ".PNG"

    ' Never overwrite silently
    If FileExists(sDownloadsPath) Then
        If MsgBox(FN & ".PNG already exists in your Downloads folder." & vbCrLf & vbCrLf & _
                  "Do you want to overwrite it?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "File Already Exists") <> vbYes Then
            Exit Function
        End If
    End If

    Dim myModelView As Object
    Set myModelView = Part.ActiveView
    myModelView.FrameState = swWindowState_e.swWindowMaximized

    ' Save As PNG. The export settings are whatever the form has already applied.
    ' The last argument (2) is swSaveAsOptions_e.swSaveAsOptions_Copy, so the
    ' export does not become the active document.
    longstatus = Part.SaveAs3(sDownloadsPath, 0, 2) ' 0 = swSaveAsVersion_e.swSaveAsCurrentVersion

    If longstatus = 0 Then ' If save was successful
        MsgBox "Document successfully saved as: " & sDownloadsPath, vbInformation
        SaveAsPNG = True
    Else
        MsgBox "Failed to save document. Status code: " & longstatus, vbExclamation
    End If
End Function

' Dir raises an error rather than returning "" on a malformed name, so a bad
' file name falls through to SaveAs3 and gets reported as a save failure.
Private Function FileExists(ByVal sPath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(sPath) <> "")
    If Err.Number <> 0 Then
        Err.Clear
        FileExists = False
    End If
    On Error GoTo 0
End Function

' --- Screen capture / print capture -----------------------------------------

Public Function GetCaptureMode() As Long
    Set swApp = Application.SldWorks
    GetCaptureMode = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture)
End Function

Public Sub SetCaptureMode(ByVal bScreenCapture As Boolean)
    Dim lCapture As Long

    If bScreenCapture Then lCapture = CAPTURE_SCREEN Else lCapture = CAPTURE_PRINT

    Set swApp = Application.SldWorks
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture, lCapture
End Sub

' --- DPI ---------------------------------------------------------------------

Public Function GetExportDPI() As Long
    Set swApp = Application.SldWorks
    GetExportDPI = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swTiffPrintDPI)
End Function

Public Sub SetExportDPI(ByVal lDPI As Long)
    If lDPI < MIN_DPI Or lDPI > MAX_DPI Then Exit Sub

    Set swApp = Application.SldWorks
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swTiffPrintDPI, lDPI
End Sub

' --- Remove background -------------------------------------------------------
' No API exists for this one, so it goes straight to the registry value that
' backs the checkbox in the export options dialog.

Public Function GetRemoveBackground() As Boolean
    Dim wsh As Object
    Dim vValue As Variant

    Set wsh = CreateObject("WScript.Shell")

    On Error Resume Next
    vValue = wsh.RegRead(ExportSettingsKey() & REG_REMOVE_BACKGROUND)
    If Err.Number <> 0 Then
        Err.Clear ' value not written yet - SOLIDWORKS treats that as off
    Else
        GetRemoveBackground = (CLng(vValue) <> 0)
    End If
    On Error GoTo 0
End Function

' Returns True if the value was written and read back as expected.
Public Function SetRemoveBackground(ByVal bRemoveBackground As Boolean) As Boolean
    Dim wsh As Object
    Dim lValue As Long

    Set wsh = CreateObject("WScript.Shell")

    If bRemoveBackground Then lValue = 1 Else lValue = 0

    On Error Resume Next
    wsh.RegWrite ExportSettingsKey() & REG_REMOVE_BACKGROUND, lValue, "REG_DWORD"
    Err.Clear
    On Error GoTo 0

    SetRemoveBackground = (GetRemoveBackground() = bRemoveBackground)
End Function

' Builds the version specific export settings key, e.g.
' "HKEY_CURRENT_USER\Software\SolidWorks\SOLIDWORKS 2024\Export Settings\"
' RevisionNumber returns "32.5.0" for SOLIDWORKS 2024, and the release year is
' always 1992 + the major number.
Private Function ExportSettingsKey() As String
    Dim vParts As Variant

    Set swApp = Application.SldWorks
    vParts = Split(swApp.RevisionNumber, ".")

    ExportSettingsKey = "HKEY_CURRENT_USER\Software\SolidWorks\SOLIDWORKS " & _
                        CStr(1992 + CLng(vParts(0))) & "\Export Settings\"
End Function
