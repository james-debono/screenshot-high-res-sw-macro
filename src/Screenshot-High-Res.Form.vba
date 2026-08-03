' ===========================================================================
' Export PNG 0.2.0 - UserForm1
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
' ===========================================================================

Option Explicit

Private Sub CommandButton1_Click()
    Dim lDPI As Long

    If Trim$(Me.TextBox1.Text) = "" Then
        MsgBox "Please enter a file name.", vbExclamation
        Me.TextBox1.SetFocus
        Exit Sub
    End If

    lDPI = 0
    If IsNumeric(Me.ComboBox1.Value) Then lDPI = CLng(Me.ComboBox1.Value)

    If Me.OptionButton2.Value Then
        ' Print capture - the DPI is actually used, so it has to be valid
        If lDPI < MIN_DPI Or lDPI > MAX_DPI Then
            MsgBox "DPI must be a number between " & MIN_DPI & " and " & MAX_DPI & ".", vbExclamation
            Me.ComboBox1.SetFocus
            Exit Sub
        End If
    ElseIf lDPI < MIN_DPI Or lDPI > MAX_DPI Then
        ' Screen capture - an unusable DPI is left alone rather than blocking the export
        lDPI = 0
    End If

    ' Pass the form values to the main macro and close the form
    Call SaveAsPNG(Me.TextBox1.Text, Me.OptionButton1.Value, lDPI, (Me.CheckBox1.Value = True))
    Unload Me
End Sub

Private Sub OptionButton1_Click()
    UpdateCaptureControls
End Sub

Private Sub OptionButton2_Click()
    UpdateCaptureControls
End Sub

Private Sub UserForm_Initialize()
    Dim swApp As Object
    Dim swModel As Object

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Fill the form from the current PNG export settings
    PopulateDpiList
    LoadExportSettings swApp
    UpdateCaptureControls

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
End Sub

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
Private Sub LoadExportSettings(ByVal swApp As Object)
    Dim lDPI As Long

    If GetCaptureMode(swApp) = CAPTURE_SCREEN Then
        Me.OptionButton1.Value = True
    Else
        Me.OptionButton2.Value = True
    End If

    lDPI = GetExportDPI(swApp)
    If lDPI < MIN_DPI Or lDPI > MAX_DPI Then lDPI = 300

    If Not DpiInList(lDPI) Then Me.ComboBox1.AddItem CStr(lDPI)
    Me.ComboBox1.Value = CStr(lDPI)

    Me.CheckBox1.Value = GetRemoveBackground(swApp)
End Sub

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
