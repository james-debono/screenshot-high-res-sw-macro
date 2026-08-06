' ===========================================================================
' Export PNG 0.4.0 - UserForm1
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
' The form is modeless, and the macro's idle loop calls AutoRefreshTick below
' about every 10 ms. That is where the camera gets watched, so the preview
' refreshes itself shortly after you stop moving the view. Refresh Image is
' still there for when you want it now.
' ===========================================================================

Option Explicit

' Set while the form is being populated, so the control events below do not
' start rendering previews before everything is in place.
Private m_bLoading As Boolean

' The macro idles in a DoEvents loop to stay alive while this form is open, and
' the export itself can pump messages. Without this guard a second click part
' way through an export would start a second one on top of the first.
Private m_bBusy As Boolean

' Camera watching
Private m_sLastCamera As String
Private m_sngChangedAt As Single
Private m_sngLastPoll As Single
Private m_bPreviewStale As Boolean

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

    RefreshPreview True, False
End Sub

Private Sub UserForm_Terminate()
    ClearPreviewCache
End Sub

' --- Auto refresh -------------------------------------------------------------
' Called from the macro's idle loop. Notices that the view has moved, waits for
' it to settle, then refreshes once. Polling beats hooking DModelViewEvents
' here: it survives switching documents without re-hooking, and it cannot be
' swamped by a notification per frame during an orbit.

Public Sub AutoRefreshTick()
    Dim sNow As String
    Dim sngElapsed As Single
    Dim sngSincePoll As Single

    If Not AUTO_REFRESH Then Exit Sub
    If m_bLoading Or m_bBusy Then Exit Sub

    ' The idle loop ticks every 10 ms; reading the camera that often would be
    ' hundreds of COM calls a second for nothing
    sngSincePoll = Timer - m_sngLastPoll
    If sngSincePoll >= 0 And sngSincePoll < POLL_SECONDS Then Exit Sub
    m_sngLastPoll = Timer

    sNow = CameraSignature()
    If sNow = "" Then Exit Sub ' no document, or the view is not readable

    If sNow <> m_sLastCamera Then
        ' Still moving - restart the settle clock and wait
        m_sLastCamera = sNow
        m_sngChangedAt = Timer
        m_bPreviewStale = True
        Exit Sub
    End If

    If Not m_bPreviewStale Then Exit Sub

    sngElapsed = Timer - m_sngChangedAt
    If sngElapsed < 0 Then sngElapsed = SETTLE_SECONDS ' Timer wrapped at midnight
    If sngElapsed < SETTLE_SECONDS Then Exit Sub

    RefreshPreview True, False
End Sub

' Take the current camera as the new baseline. Called after every render,
' because the print capture can nudge the view - without this the render would
' look like another camera move and refresh forever.
Private Sub BaselineCamera()
    m_sLastCamera = CameraSignature()
    m_bPreviewStale = False
End Sub

' --- Buttons ------------------------------------------------------------------

Private Sub CommandButton2_Click() ' Refresh Image
    RefreshPreview True, True
End Sub

Private Sub CommandButton1_Click() ' Export
    Dim pxW As Long, pxH As Long
    Dim sName As String
    Dim lResult As Long

    If m_bBusy Then Exit Sub

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

    m_bBusy = True
    Me.MousePointer = fmMousePointerHourGlass
    lResult = ExportToDownloads(sName, pxW, pxH, Me.OptionButton1.Value)
    Me.MousePointer = fmMousePointerDefault
    m_bBusy = False

    BaselineCamera

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
    RefreshPreview False, False
End Sub

Private Sub OptionButton2_Click() ' White Background
    RefreshPreview False, False
End Sub

' --- Size boxes ---------------------------------------------------------------
' Marked stale rather than refreshed on the spot, so the settle delay lets you
' finish typing before anything renders.

Private Sub TextBox2_Change()
    MarkStale
End Sub

Private Sub TextBox3_Change()
    MarkStale
End Sub

Private Sub MarkStale()
    If m_bLoading Then Exit Sub

    m_sngChangedAt = Timer
    m_bPreviewStale = True
End Sub

' --- Preview ------------------------------------------------------------------

' bForce  - re-render rather than just re-compositing the cached render
' bReport - show a message box on a problem. Off for automatic refreshes, so a
'           half typed width cannot pop a dialog while you are still typing.
Private Sub RefreshPreview(ByVal bForce As Boolean, ByVal bReport As Boolean)
    Dim pxW As Long, pxH As Long
    Dim sBmp As String

    If m_bLoading Then Exit Sub
    If m_bBusy Then Exit Sub

    If Not HasActiveDoc() Then
        Me.Image1.Picture = LoadPicture("")
        If bReport Then MsgBox "No document is open.", vbExclamation
        Exit Sub
    End If

    ' Picked up here as well as at load, so opening a document after the form
    ' still fills the name in
    If Trim$(Me.TextBox1.Text) = "" Then Me.TextBox1.Text = ActiveDocName()

    If Not TryGetSize(pxW, pxH, bReport) Then Exit Sub

    m_bBusy = True
    Me.MousePointer = fmMousePointerHourGlass
    sBmp = PreviewToBmp(pxW, pxH, BackgroundMode(), bForce)
    Me.MousePointer = fmMousePointerDefault
    m_bBusy = False

    BaselineCamera

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
