' ===========================================================================
' Export PNG 0.3.1 - UserForm1
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
' The form is modeless, so you can orbit and zoom with it open and then hit
' Refresh Image to see the new framing. Nothing about the document is cached at
' load - open, close or switch documents freely while it is up.
' ===========================================================================

Option Explicit

' Set while the form is being populated, so the control events below do not
' start rendering previews before everything is in place.
Private m_bLoading As Boolean

Private Sub UserForm_Initialize()
    m_bLoading = True

    ' Zoom keeps any output aspect ratio correct inside the preview box instead
    ' of stretching it to fit
    Me.Image1.PictureSizeMode = fmPictureSizeModeZoom

    ' SIZE BOXES
    Me.TextBox2.Text = CStr(DEFAULT_WIDTH)
    Me.TextBox3.Text = CStr(DEFAULT_HEIGHT)

    Me.OptionButton1.Value = True ' Transparent Background
    Me.TextBox1.Text = ActiveDocName()

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

    If Not HasActiveDoc() Then
        MsgBox "No document is open.", vbExclamation
        Exit Sub
    End If

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

    ' The form stays open on purpose now that it is modeless - reframe and
    ' export again without reopening it.
    Select Case lResult
        Case EXPORT_OK
            MsgBox "Saved to:" & vbCrLf & _
                   Environ("USERPROFILE") & "\Downloads\" & sName & ".PNG", vbInformation

        Case EXPORT_NO_ALPHA
            MsgBox "Exported, but the image has no transparent background." & vbCrLf & vbCrLf & _
                   "Tick 'Remove background' in Tools > Options > System Options > " & _
                   "Export > TIF/PSD/JPG/PNG." & vbCrLf & vbCrLf & _
                   "That option is not in the SOLIDWORKS API, so it cannot be set from " & _
                   "this macro - it has to be ticked there once and left on.", vbExclamation

        Case EXPORT_CANCELLED
            ' Declined the overwrite prompt - nothing to say

        Case EXPORT_NO_DOC
            MsgBox "No document is open.", vbExclamation

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

    If Not HasActiveDoc() Then
        Me.Image1.Picture = LoadPicture("")
        If bForce Then MsgBox "No document is open.", vbExclamation
        Exit Sub
    End If

    ' Picked up here as well as at load, so opening a document after the form
    ' still fills the name in
    If Trim$(Me.TextBox1.Text) = "" Then Me.TextBox1.Text = ActiveDocName()

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
