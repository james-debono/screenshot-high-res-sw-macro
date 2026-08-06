' ============================================================================
' Export PNG
' UserForm1 - user interface
'
' Collects the save location, file name, output pixel size and background
' choice, displays a preview of the framing, and triggers the export. All
' SOLIDWORKS work is done in Module1.
'
'   Version   0.5.0
'   Author    James Debono
'   Updated   2026-08-06
'   License   (to be added)
'
' Controls
'   Image1          preview
'   CommandButton1  Export
'   CommandButton2  Refresh Image
'   CommandButton3  Browse
'   ComboBox1       save location
'   TextBox1        file name
'   TextBox2        output width in pixels
'   TextBox3        output height in pixels
'   OptionButton1   transparent background
'   OptionButton2   white background
'
' The form is modeless, and Module1's idle loop calls AutoRefreshTick roughly
' every 10 ms. That is where the camera is watched, so the preview refreshes
' itself shortly after the view stops moving.
' ============================================================================

Option Explicit

' True while the form is being populated, so that control events do not render
' previews before initialisation has finished.
Private m_bLoading As Boolean

' True while an export, preview or dialog is in progress. The idle loop pumps
' messages and the export can do the same, so without this guard a second click
' part way through would start another operation on top of the first - and an
' automatic refresh could fire underneath the Browse dialog.
Private m_bBusy As Boolean

' Camera watching state
Private m_sLastCamera As String
Private m_sngChangedAt As Single
Private m_sngLastPoll As Single
Private m_bPreviewStale As Boolean

Private Sub UserForm_Initialize()
    m_bLoading = True

    ' Zoom preserves the output aspect ratio inside the preview box rather than
    ' stretching the image to fill it
    Me.Image1.PictureSizeMode = fmPictureSizeModeZoom

    ' Output size defaults. TextBox2 is width, TextBox3 is height.
    Me.TextBox2.Text = CStr(DEFAULT_WIDTH)
    Me.TextBox3.Text = CStr(DEFAULT_HEIGHT)

    Me.OptionButton1.Value = True ' transparent background
    Me.TextBox1.Text = ActiveDocName()

    PopulateLocations

    m_bLoading = False

    RefreshPreview True, False
End Sub

Private Sub UserForm_Terminate()
    ClearPreviewCache
End Sub

' --- Save location ----------------------------------------------------------

' The reserved model folder entry first, then recently used folders, then the
' Downloads default if it is not already present.
Private Sub PopulateLocations()
    Dim vRecent As Variant
    Dim sLast As String
    Dim i As Long

    ' Editable, so a path can be typed or pasted as well as picked
    Me.ComboBox1.Style = fmStyleDropDownCombo
    Me.ComboBox1.Clear

    Me.ComboBox1.AddItem MODEL_FOLDER_TOKEN

    vRecent = Split(RecentFolders(), "|")
    For i = LBound(vRecent) To UBound(vRecent)
        If vRecent(i) <> "" Then Me.ComboBox1.AddItem vRecent(i)
    Next i

    If Not LocationInList(DefaultSaveFolder()) Then Me.ComboBox1.AddItem DefaultSaveFolder()

    sLast = LastLocationChoice()
    If Trim$(sLast) = "" Then sLast = DefaultSaveFolder()
    Me.ComboBox1.Text = sLast
End Sub

Private Function LocationInList(ByVal sPath As String) As Boolean
    Dim i As Long

    For i = 0 To Me.ComboBox1.ListCount - 1
        If StrComp(Me.ComboBox1.List(i), sPath, vbTextCompare) = 0 Then
            LocationInList = True
            Exit Function
        End If
    Next i
End Function

Private Sub CommandButton3_Click() ' Browse
    Dim sPicked As String

    If m_bBusy Then Exit Sub

    ' Held for the duration of the dialog, otherwise an automatic refresh can
    ' fire underneath it - the idle loop keeps running while it is open
    m_bBusy = True
    sPicked = BrowseFolder("Choose where to save the exported image")
    m_bBusy = False

    If sPicked = "" Then Exit Sub ' cancelled

    If Not LocationInList(sPicked) Then Me.ComboBox1.AddItem sPicked
    Me.ComboBox1.Text = sPicked
End Sub

' Resolves the chosen location to a real folder, offering to create it if it
' does not exist. Returns "" if it could not be settled, having already
' explained why.
Private Function SettleSaveFolder() As String
    Dim sChoice As String
    Dim sFolder As String

    sChoice = Trim$(Me.ComboBox1.Text)

    If sChoice = "" Then
        MsgBox "Please choose a save location.", vbExclamation
        Me.ComboBox1.SetFocus
        Exit Function
    End If

    sFolder = ResolveSaveFolder(sChoice)

    If sFolder = "" Then
        If StrComp(sChoice, MODEL_FOLDER_TOKEN, vbTextCompare) = 0 Then
            MsgBox "The active document has not been saved yet, so it has no " & _
                   "folder to export beside." & vbCrLf & vbCrLf & _
                   "Save the document first, or choose another location.", vbExclamation
        Else
            MsgBox "That save location could not be understood.", vbExclamation
            Me.ComboBox1.SetFocus
        End If
        Exit Function
    End If

    If Not FolderExists(sFolder) Then
        If MsgBox("This folder does not exist:" & vbCrLf & sFolder & vbCrLf & vbCrLf & _
                  "Create it?", vbQuestion + vbYesNo, "Folder Not Found") <> vbYes Then
            Exit Function
        End If

        If Not CreateFolderTree(sFolder) Then
            MsgBox "That folder could not be created:" & vbCrLf & sFolder, vbExclamation
            Exit Function
        End If
    End If

    SettleSaveFolder = sFolder
End Function

' --- Auto refresh -----------------------------------------------------------

' Called from Module1's idle loop. Detects that the view has moved, waits for it
' to settle, then refreshes once.
'
' Polling is used in preference to hooking DModelViewEvents: it survives a
' change of document without re-hooking, and cannot be swamped by one
' notification per frame during an orbit.
Public Sub AutoRefreshTick()
    Dim sNow As String
    Dim sngElapsed As Single
    Dim sngSincePoll As Single

    If Not AUTO_REFRESH Then Exit Sub
    If m_bLoading Or m_bBusy Then Exit Sub

    ' The idle loop ticks far more often than the camera needs reading
    sngSincePoll = Timer - m_sngLastPoll
    If sngSincePoll >= 0 And sngSincePoll < POLL_SECONDS Then Exit Sub
    m_sngLastPoll = Timer

    sNow = CameraSignature()
    If sNow = "" Then Exit Sub ' no document, or the view is not readable

    If sNow <> m_sLastCamera Then
        ' Still moving: restart the settle clock and wait
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

' Adopts the current camera as the baseline. Called after every render, because
' the print capture can shift the view; without this the render itself would
' register as a further camera movement and refresh indefinitely.
Private Sub BaselineCamera()
    m_sLastCamera = CameraSignature()
    m_bPreviewStale = False
End Sub

' --- Buttons ----------------------------------------------------------------

Private Sub CommandButton2_Click() ' Refresh Image
    RefreshPreview True, True
End Sub

Private Sub CommandButton1_Click() ' Export
    Dim pxW As Long, pxH As Long
    Dim sName As String
    Dim sFolder As String
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

    ' Settled before any rendering, so a bad path fails immediately rather than
    ' after a full size export
    m_bBusy = True
    sFolder = SettleSaveFolder()
    m_bBusy = False
    If sFolder = "" Then Exit Sub

    m_bBusy = True
    Me.MousePointer = fmMousePointerHourGlass
    lResult = ExportToFolder(sFolder, sName, pxW, pxH, Me.OptionButton1.Value)
    Me.MousePointer = fmMousePointerDefault
    m_bBusy = False

    BaselineCamera

    ' The form is left open so that the view can be reframed and exported again
    ' without reopening it.
    Select Case lResult
        Case EXPORT_OK, EXPORT_NO_ALPHA
            RememberLocation sFolder

            If lResult = EXPORT_OK Then
                MsgBox "Saved to:" & vbCrLf & sFolder & "\" & sName & ".PNG", vbInformation
            Else
                MsgBox "Exported, but the image has no transparent background." & vbCrLf & vbCrLf & _
                       "Tick 'Remove background' in Tools > Options > System Options > " & _
                       "Export > TIF/PSD/JPG/PNG." & vbCrLf & vbCrLf & _
                       "That option is not in the SOLIDWORKS API, so it cannot be set from " & _
                       "this macro - it has to be ticked there once and left on.", vbExclamation
            End If

        Case EXPORT_CANCELLED
            ' Overwrite prompt declined; nothing to report

        Case EXPORT_NO_DOC
            MsgBox "No document is open.", vbExclamation

        Case Else
            MsgBox "Export failed.", vbExclamation
    End Select
End Sub

' Stores the choice as typed, so the model folder entry is remembered as such,
' and the resolved folder in the recent list.
Private Sub RememberLocation(ByVal sFolder As String)
    Dim sChoice As String

    sChoice = Trim$(Me.ComboBox1.Text)
    SaveLastLocationChoice sChoice

    If StrComp(sChoice, MODEL_FOLDER_TOKEN, vbTextCompare) <> 0 Then
        PushRecentFolder sFolder
    End If
End Sub

' --- Background choice ------------------------------------------------------
' Affects only how the preview is composited, so no re-render is required.

Private Sub OptionButton1_Click() ' transparent background
    RefreshPreview False, False
End Sub

Private Sub OptionButton2_Click() ' white background
    RefreshPreview False, False
End Sub

' --- Size boxes -------------------------------------------------------------
' Marked stale rather than refreshed immediately, so that the settle delay
' allows typing to finish before anything is rendered.

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

' --- Preview ----------------------------------------------------------------

' bForce  - re-render rather than re-compositing the cached render
' bReport - report problems in a message box. Suppressed for automatic
'           refreshes, so a partially typed width cannot raise a dialog
'           mid-keystroke.
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

    ' Also handled here, not only at load, so that opening a document after the
    ' form still populates the name
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

' A checkerboard behind a transparent export makes the transparent areas
' visible; a white background previews on white.
Private Function BackgroundMode() As Long
    If Me.OptionButton1.Value Then
        BackgroundMode = BG_CHECKER
    Else
        BackgroundMode = BG_WHITE
    End If
End Function

' --- Validation -------------------------------------------------------------

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
