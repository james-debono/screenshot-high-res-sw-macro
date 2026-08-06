' ===========================================================================
' Export PNG 0.3.0 - UserForm1
'
' Paste-ready: right-click UserForm1 > View Code, select all, paste this in.
' Do NOT import this as a .frm - that would overwrite the layout you built.
'
' Controls this code expects:
'   Image1          preview
'   CommandButton2  "Refresh Image"
'   TextBox1        file name
'   OptionButton1   "Transparent Background"
'   OptionButton2   "White Background"
'   TextBox2        Width  (pixels)
'   TextBox3        Height (pixels)
'   CommandButton1  "Export"
'
' Width and Height default to 1920 x 1080 on open. If those land in the wrong
' boxes then TextBox2/TextBox3 are the other way round on your form - swap the
' two lines marked SIZE BOXES below and everything else follows.
' ===========================================================================

Option Explicit

' Set while the form is being populated, so the control events below do not
' start rendering previews before everything is in place.
Private m_bLoading As Boolean

Private Sub UserForm_Initialize()
    Dim swApp As Object
    Dim swModel As Object

    m_bLoading = True

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Zoom keeps any output aspect ratio correct inside the preview box instead
    ' of stretching it to fit
    Me.Image1.PictureSizeMode = fmPictureSizeModeZoom

    ' SIZE BOXES
    Me.TextBox2.Text = CStr(DEFAULT_WIDTH)
    Me.TextBox3.Text = CStr(DEFAULT_HEIGHT)

    Me.OptionButton1.Value = True ' Transparent Background

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
        Me.CommandButton1.Enabled = False ' Disable buttons if no doc is open
        Me.CommandButton2.Enabled = False
    End If

    m_bLoading = False

    ' Comment this line out if opening the form ever feels slow - the Refresh
    ' Image button does exactly the same thing on demand.
    RefreshPreview True
End Sub

Private Sub UserForm_Terminate()
    ClearPreviewCache
End Sub

' --- Buttons ------------------------------------------------------------------

Private Sub CommandButton2_Click() ' Refresh Image
    RefreshPreview True
End Sub

Private Sub CommandButton1_Click() ' Export
    Dim pxW As Long, pxH As Long
    Dim sName As String
    Dim lResult As Long

    sName = Trim$(Me.TextBox1.Text)

    If sName = "" Then
        MsgBox "Please enter a file name.", vbExclamation
        Me.TextBox1.SetFocus
        Exit Sub
    End If

    If Not TryGetSize(pxW, pxH, True) Then
        Me.TextBox2.SetFocus
        Exit Sub
    End If

    Me.MousePointer = fmMousePointerHourGlass
    lResult = ExportToDownloads(sName, pxW, pxH, Me.OptionButton1.Value)
    Me.MousePointer = fmMousePointerDefault

    Select Case lResult
        Case 0
            MsgBox "Saved to:" & vbCrLf & _
                   Environ("USERPROFILE") & "\Downloads\" & sName & ".PNG", vbInformation
            Unload Me

        Case 1
            ' Exported fine, but Transparent was asked for and the file is opaque
            MsgBox "Exported, but the image has no transparent background." & vbCrLf & vbCrLf & _
                   "Tick 'Remove background' in Tools > Options > System Options > " & _
                   "Export > TIF/PSD/JPG/PNG." & vbCrLf & vbCrLf & _
                   "That option is not in the SOLIDWORKS API, so it cannot be set from " & _
                   "this macro - it has to be ticked there once and left on.", vbExclamation
            Unload Me

        Case 2
            ' Cancelled at the overwrite prompt - leave the form open

        Case Else
            MsgBox "Export failed.", vbExclamation
    End Select
End Sub

' --- Background choice --------------------------------------------------------
' Only changes how the preview is composited, so no re-render is needed.

Private Sub OptionButton1_Click() ' Transparent Background
    RefreshPreview False
End Sub

Private Sub OptionButton2_Click() ' White Background
    RefreshPreview False
End Sub

' --- Preview ------------------------------------------------------------------

Private Sub RefreshPreview(ByVal bForce As Boolean)
    Dim pxW As Long, pxH As Long
    Dim sBmp As String

    If m_bLoading Then Exit Sub
    If Not Me.CommandButton2.Enabled Then Exit Sub
    If Not TryGetSize(pxW, pxH, bForce) Then Exit Sub

    Me.MousePointer = fmMousePointerHourGlass
    sBmp = PreviewToBmp(pxW, pxH, BackgroundMode(), bForce)
    Me.MousePointer = fmMousePointerDefault

    If sBmp = "" Then
        Me.Image1.Picture = LoadPicture("")
        Exit Sub
    End If

    Me.Image1.Picture = LoadPicture(sBmp)
End Sub

' Checkerboard behind a transparent export so you can see what is see-through,
' plain white when that is what you asked for.
Private Function BackgroundMode() As Long
    If Me.OptionButton1.Value Then
        BackgroundMode = BG_CHECKER
    Else
        BackgroundMode = BG_WHITE
    End If
End Function

' --- Validation ---------------------------------------------------------------

Private Function TryGetSize(ByRef pxW As Long, ByRef pxH As Long, ByVal bReport As Boolean) As Boolean
    If Not IsValidPx(Me.TextBox2.Text) Or Not IsValidPx(Me.TextBox3.Text) Then
        If bReport Then
            MsgBox "Width and Height must be whole numbers between " & _
                   MIN_PX & " and " & MAX_PX & " pixels.", vbExclamation
        End If
        Exit Function
    End If

    pxW = CLng(Me.TextBox2.Text)
    pxH = CLng(Me.TextBox3.Text)

    TryGetSize = True
End Function

Private Function IsValidPx(ByVal s As String) As Boolean
    s = Trim$(s)

    If s = "" Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    If InStr(s, ".") > 0 Or InStr(s, ",") > 0 Then Exit Function

    IsValidPx = (CLng(s) >= MIN_PX And CLng(s) <= MAX_PX)
End Function
