' ===========================================================================
' Export PNG 0.2.1 - UserForm1
'
' Paste-ready: right-click UserForm1 > View Code, select all, paste this in.
' Do NOT import this as a .frm - that would overwrite the layout you built.
'
' Controls this code expects (the names already on the 0.2.0 form):
'   TextBox1        file name
'   Label1          "File Name"
'   CommandButton1  "Export"
'   OptionButton1   "Screen Capture"
'   OptionButton2   "Print Capture"
'   ComboBox1       DPI
'   Label2          "DPI"
'   CheckBox1       "Remove Background"
'
' The three setting controls are a shortcut to Tools > Options > Export >
' TIF/PSD/JPG/PNG. They show what SOLIDWORKS is currently set to when the form
' opens, and each one is written back the moment it is changed - so closing the
' form without exporting still leaves the setting changed.
' ===========================================================================

Option Explicit

' Set while the form is being populated, so the control events below do not
' write the settings straight back while they are still being loaded.
Private m_bLoading As Boolean

Private Sub UserForm_Initialize()
    Dim swApp As Object
    Dim swModel As Object

    m_bLoading = True

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Fill the form from the current PNG export settings
    PopulateDpiList
    LoadExportSettings

    If Not swModel Is Nothing Then
        Dim sFullPath As String
        sFullPath = swModel.GetPathName

        If sFullPath <> "" Then
            ' Extract only the filename from the full path
            Dim sFileNameNoExtension As String
            sFileNameNoExtension = Left(sFullPath, InStrRev(sFullPath, ".") - 1) ' Remove extension first
            sFileNameNoExtension = Mid(sFileNameNoExtension, InStrRev(sFileNameNoExtension, "\") + 1) ' Get name after last backslash

            Me.TextBox1.Text = sFileNameNoExtension
        Else
            Me.TextBox1.Text = "NewDocument" ' Default if not saved yet
        End If
    Else
        Me.TextBox1.Text = "NoDocumentOpen"
        Me.CommandButton1.Enabled = False ' Disable button if no doc is open
    End If

    m_bLoading = False

    UpdateCaptureControls
End Sub

Private Sub CommandButton1_Click()
    If Trim$(Me.TextBox1.Text) = "" Then
        MsgBox "Please enter a file name.", vbExclamation
        Me.TextBox1.SetFocus
        Exit Sub
    End If

    ' Print capture is the only mode that uses the DPI, so make sure it is usable
    If Me.OptionButton2.Value Then
        If Not IsValidDpi(Me.ComboBox1.Value) Then
            MsgBox "DPI must be a number between " & MIN_DPI & " and " & MAX_DPI & ".", vbExclamation
            Me.ComboBox1.SetFocus
            Exit Sub
        End If
    End If

    ' The settings have already been applied as they were changed, so this only
    ' has to do the export. The form stays open if it did not happen.
    If SaveAsPNG(Trim$(Me.TextBox1.Text)) Then Unload Me
End Sub

' --- Setting controls - each one applies immediately --------------------------

Private Sub OptionButton1_Click()
    If m_bLoading Then Exit Sub

    SetCaptureMode True ' Screen capture
    UpdateCaptureControls
End Sub

Private Sub OptionButton2_Click()
    If m_bLoading Then Exit Sub

    SetCaptureMode False ' Print capture
    UpdateCaptureControls
End Sub

Private Sub ComboBox1_Change()
    If m_bLoading Then Exit Sub

    ' Ignore half typed values rather than pushing nonsense into the settings
    If IsValidDpi(Me.ComboBox1.Value) Then SetExportDPI CLng(Me.ComboBox1.Value)
End Sub

Private Sub CheckBox1_Click()
    If m_bLoading Then Exit Sub

    If Not SetRemoveBackground(Me.CheckBox1.Value = True) Then
        MsgBox "Could not write the Remove Background setting." & vbCrLf & vbCrLf & _
               "This option has no SOLIDWORKS API, so the macro writes it straight " & _
               "to the registry, and that write was refused.", vbExclamation
    End If
End Sub

' --- Helpers ------------------------------------------------------------------

' The DPI values offered by the SOLIDWORKS export options dialog.
Private Sub PopulateDpiList()
    Dim vDPI As Variant
    Dim i As Long

    vDPI = Array(72, 96, 150, 200, 300, 400, 600)

    Me.ComboBox1.Clear
    For i = LBound(vDPI) To UBound(vDPI)
        Me.ComboBox1.AddItem CStr(vDPI(i))
    Next i
End Sub

' Reads the current export settings and shows them on the form.
Private Sub LoadExportSettings()
    Dim lDPI As Long

    If GetCaptureMode() = CAPTURE_SCREEN Then
        Me.OptionButton1.Value = True
    Else
        Me.OptionButton2.Value = True
    End If

    lDPI = GetExportDPI()
    If lDPI < MIN_DPI Or lDPI > MAX_DPI Then lDPI = 300

    If Not DpiInList(lDPI) Then Me.ComboBox1.AddItem CStr(lDPI)
    Me.ComboBox1.Value = CStr(lDPI)

    Me.CheckBox1.Value = GetRemoveBackground()
End Sub

Private Function IsValidDpi(ByVal vValue As Variant) As Boolean
    If Not IsNumeric(vValue) Then Exit Function

    IsValidDpi = (CLng(vValue) >= MIN_DPI And CLng(vValue) <= MAX_DPI)
End Function

Private Function DpiInList(ByVal lDPI As Long) As Boolean
    Dim i As Long

    For i = 0 To Me.ComboBox1.ListCount - 1
        If Me.ComboBox1.List(i) = CStr(lDPI) Then
            DpiInList = True
            Exit Function
        End If
    Next i
End Function

' DPI only applies to a print capture, so grey it out for a screen capture.
Private Sub UpdateCaptureControls()
    Dim bPrintCapture As Boolean

    bPrintCapture = Me.OptionButton2.Value

    Me.ComboBox1.Enabled = bPrintCapture
    Me.Label2.Enabled = bPrintCapture
End Sub
