' ===========================================================================
' Export PNG 0.3.2 - Module1
'
' Paste-ready: open Module1 in the VBA editor, select all, paste this in.
' (The "Attribute VB_Name" line is deliberately not here - it is managed by
' VBA and pasting it into the code window is a syntax error.)
' ===========================================================================

Option Explicit

' --- How this works ----------------------------------------------------------
' Everything goes through the normal SOLIDWORKS PNG exporter in Print capture
' mode, which is the only way to exceed screen resolution.
'
' You give a pixel width and height. Paper size is set to User Defined and
' worked out from those pixels at EXPORT_DPI, so paper x DPI lands exactly on
' the size you asked for:
'
'     paper inches = pixels / EXPORT_DPI
'
' Framing depends only on the paper ASPECT RATIO, not its absolute size. The
' preview exploits that: it uses the same paper at a much lower DPI, so it is
' framed identically to the real export but renders a fraction of the pixels.
'
' Transparency comes from the "Remove background" tickbox in
' Tools > Options > Export > TIF/PSD/JPG/PNG. That setting is not in the
' SOLIDWORKS API and cannot be changed from a macro, so leave it TICKED. The
' export then always comes out transparent, and White Background is produced by
' flattening it onto white afterwards. If you pick Transparent and the file
' comes back opaque, the macro tells you that tickbox has been turned off
' rather than silently handing you a white rectangle.
'
' The form is shown MODELESS so you can orbit and zoom with it open, then hit
' Refresh Image to see the new framing. Everything therefore re-reads the
' active document each time rather than caching it at load.
' -----------------------------------------------------------------------------

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Const CAPTURE_SCREEN As Long = 0
Public Const CAPTURE_PRINT As Long = 1

' DPI used for the real export. Paper size is derived from it, so this only
' decides how big the "paper" is - the output pixel size is whatever you type.
Public Const EXPORT_DPI As Long = 300

Public Const DEFAULT_WIDTH As Long = 1920
Public Const DEFAULT_HEIGHT As Long = 1080
Public Const MIN_PX As Long = 16
Public Const MAX_PX As Long = 10000

' Roughly how many pixels wide the preview render should be
Private Const PREVIEW_PX As Long = 520
Private Const MIN_PREVIEW_DPI As Long = 6

' How long the idle loop naps between message pumps. Small enough to stay
' imperceptible while orbiting, large enough not to spin a core flat.
Private Const IDLE_MS As Long = 10

' Background handling passed to the image helper
Public Const BG_KEEP As Long = 0
Public Const BG_WHITE As Long = 1
Public Const BG_CHECKER As Long = 2

' ExportToDownloads results
Public Const EXPORT_OK As Long = 0
Public Const EXPORT_NO_ALPHA As Long = 1
Public Const EXPORT_CANCELLED As Long = 2
Public Const EXPORT_FAILED As Long = 3
Public Const EXPORT_NO_DOC As Long = 4

Private Const METRES_PER_INCH As Double = 0.0254

Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

' Cached preview render, so changing Transparent/White only re-composites
Private m_sPreviewPng As String
Private m_lPreviewW As Long
Private m_lPreviewH As Long

Sub ShowSaveAsForm()
    Load UserForm1
    UserForm1.Show vbModeless

    ' A SOLIDWORKS macro ends the moment this sub returns, and that takes any
    ' modeless form down with it - which looks exactly like the form never
    ' opening, with no error. Idling here keeps the macro alive until the form
    ' is closed.
    '
    ' DoEvents hands control back to SOLIDWORKS so the view stays fully
    ' interactive; the short Sleep stops the loop spinning a CPU core flat.
    Do While UserForms.Count > 0
        DoEvents
        Sleep IDLE_MS
    Loop
End Sub

' --- Active document ----------------------------------------------------------
' Re-read every time. With a modeless form the document can be opened, closed or
' switched while the form is sitting there.

Public Function HasActiveDoc() As Boolean
    Set swApp = Application.SldWorks
    HasActiveDoc = Not (swApp.ActiveDoc Is Nothing)
End Function

Public Function ActiveDocName() As String
    Dim swModel As Object
    Dim sFullPath As String
    Dim sName As String
    Dim lDot As Long

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Function

    sFullPath = swModel.GetPathName
    If sFullPath = "" Then
        ActiveDocName = "NewDocument" ' not saved yet
        Exit Function
    End If

    ' Strip the extension, then everything up to the last backslash
    lDot = InStrRev(sFullPath, ".")
    If lDot > 0 Then sName = Left(sFullPath, lDot - 1) Else sName = sFullPath
    sName = Mid(sName, InStrRev(sName, "\") + 1)

    ActiveDocName = sName
End Function

' --- Export ------------------------------------------------------------------

Public Function ExportToDownloads(ByVal FN As String, ByVal pxW As Long, ByVal pxH As Long, _
                                  ByVal bTransparent As Boolean) As Long
    Dim sPath As String
    Dim lPost As Long

    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    If Part Is Nothing Then
        ExportToDownloads = EXPORT_NO_DOC
        Exit Function
    End If

    sPath = Environ("USERPROFILE") & "\Downloads\" & FN & ".PNG"

    ' Never overwrite silently
    If FileExists(sPath) Then
        If MsgBox(FN & ".PNG already exists in your Downloads folder." & vbCrLf & vbCrLf & _
                  "Do you want to overwrite it?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "File Already Exists") <> vbYes Then
            ExportToDownloads = EXPORT_CANCELLED
            Exit Function
        End If
    End If

    If Not RenderToPng(sPath, pxW, pxH, EXPORT_DPI) Then
        ExportToDownloads = EXPORT_FAILED
        Exit Function
    End If

    ' White: flatten onto white. Transparent: leave alone, but still check that
    ' the export actually came back with an alpha channel.
    If bTransparent Then
        lPost = RunImageHelper(sPath, sPath, BG_KEEP, 0)
    Else
        lPost = RunImageHelper(sPath, sPath, BG_WHITE, 0)
    End If

    Select Case lPost
        Case 0 ' helper ok, source had transparency
            ExportToDownloads = EXPORT_OK
        Case 1 ' helper ok, source was fully opaque
            If bTransparent Then
                ExportToDownloads = EXPORT_NO_ALPHA
            Else
                ExportToDownloads = EXPORT_OK
            End If
        Case Else
            ExportToDownloads = EXPORT_FAILED
    End Select
End Function

' --- Preview -----------------------------------------------------------------

' Renders (or re-uses) a preview and returns the path of a BMP ready for the
' form's Image control. Returns "" if it could not be produced.
' VBA's LoadPicture cannot read PNG, which is why the helper emits a BMP.
Public Function PreviewToBmp(ByVal pxW As Long, ByVal pxH As Long, ByVal lBg As Long, _
                             ByVal bForce As Boolean) As String
    Dim sBmp As String
    Dim lDPI As Long

    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc
    If Part Is Nothing Then Exit Function

    If m_sPreviewPng = "" Then m_sPreviewPng = Environ("TEMP") & "\ExportPNG_preview.png"
    sBmp = Environ("TEMP") & "\ExportPNG_preview.bmp"

    ' Only re-render when the size changed or a refresh was asked for. Changing
    ' the background just re-composites the render we already have.
    If bForce Or pxW <> m_lPreviewW Or pxH <> m_lPreviewH Or Not FileExists(m_sPreviewPng) Then
        lDPI = PreviewDpi(pxW)
        If Not RenderToPng(m_sPreviewPng, pxW, pxH, lDPI) Then Exit Function
        m_lPreviewW = pxW
        m_lPreviewH = pxH
    End If

    If RunImageHelper(m_sPreviewPng, sBmp, lBg, 1) > 1 Then Exit Function

    PreviewToBmp = sBmp
End Function

Public Sub ClearPreviewCache()
    On Error Resume Next
    If m_sPreviewPng <> "" Then Kill m_sPreviewPng
    Kill Environ("TEMP") & "\ExportPNG_preview.bmp"
    On Error GoTo 0

    m_sPreviewPng = ""
    m_lPreviewW = 0
    m_lPreviewH = 0
End Sub

' Same paper, lower DPI - identical framing, far fewer pixels.
Private Function PreviewDpi(ByVal pxW As Long) As Long
    Dim lDPI As Long

    lDPI = CLng((PREVIEW_PX * CDbl(EXPORT_DPI)) / pxW)
    If lDPI < MIN_PREVIEW_DPI Then lDPI = MIN_PREVIEW_DPI
    If lDPI > EXPORT_DPI Then lDPI = EXPORT_DPI

    PreviewDpi = lDPI
End Function

' --- The actual SOLIDWORKS export --------------------------------------------

Private Function RenderToPng(ByVal sPath As String, ByVal pxW As Long, ByVal pxH As Long, _
                             ByVal lDPI As Long) As Boolean
    ApplyPrintCapture pxW, pxH, lDPI

    ' Deliberately no FrameState change here. The framing follows the graphics
    ' area aspect, so resizing or maximising the window between the preview and
    ' the export would make them disagree.
    longstatus = Part.SaveAs3(sPath, 0, 2) ' 0 = swSaveAsCurrentVersion, 2 = swSaveAsOptions_Copy

    RenderToPng = (longstatus = 0)
End Function

' Paper size always comes from the requested pixels at EXPORT_DPI so the aspect
' ratio - and therefore the framing - is the same for a preview and a full
' export. Only the DPI differs.
Private Sub ApplyPrintCapture(ByVal pxW As Long, ByVal pxH As Long, ByVal lDPI As Long)
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture, CAPTURE_PRINT
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swTiffPrintPaperSize, _
                                        swDwgPaperSizes_e.swDwgPapersUserDefined

    ' NOTE: swconst gives values 8 and 9 two names each, with width and height
    ' swapped between the pairs. These are the names the API help topic
    ' "System Options > Export > TIF/PSD/JPG/PNG" prescribes. If an export comes
    ' out transposed, swap these two lines - that is the whole fix.
    swApp.SetUserPreferenceDoubleValue swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperWidth, _
                                       PixelsToMetres(pxW, EXPORT_DPI)
    swApp.SetUserPreferenceDoubleValue swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperHeight, _
                                       PixelsToMetres(pxH, EXPORT_DPI)

    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swTiffPrintDPI, lDPI
End Sub

Private Function PixelsToMetres(ByVal px As Long, ByVal lDPI As Long) As Double
    PixelsToMetres = (CDbl(px) / lDPI) * METRES_PER_INCH
End Function

' --- Image helper ------------------------------------------------------------
' Flattens transparency, builds the checkerboard preview, converts to BMP and
' reports whether the source had an alpha channel at all.
'
' Returns 0 done and the source had transparency
'         1 done and the source was fully opaque
'         2 helper failed

Private Function RunImageHelper(ByVal sIn As String, ByVal sOut As String, _
                                ByVal lBg As Long, ByVal lBmp As Long) As Long
    Dim sScript As String

    sScript = Environ("TEMP") & "\ExportPNG_post.ps1"
    WriteImageHelper sScript

    RunImageHelper = RunHidden("powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & _
                               Quote(sScript) & " -In " & Quote(sIn) & " -Out " & Quote(sOut) & _
                               " -Bg " & lBg & " -Bmp " & lBmp)
End Function

' Written out at run time so the macro stays self contained. The inner loop is
' compiled C# via Add-Type - a per pixel PowerShell loop is far too slow at
' 3000x3000. Deliberately contains no double quotes, so it needs no escaping.
Private Sub WriteImageHelper(ByVal sPath As String)
    Dim iFile As Integer

    iFile = FreeFile
    Open sPath For Output As #iFile

    Print #iFile, "param([string]$In,[string]$Out,[int]$Bg,[int]$Bmp)"
    Print #iFile, "$ErrorActionPreference = 'Stop'"
    Print #iFile, "Add-Type -AssemblyName System.Drawing"
    Print #iFile, "$cs = @'"
    Print #iFile, "public class Post {"
    Print #iFile, "  public static bool HasAlpha(byte[] b) {"
    Print #iFile, "    for (int i = 3; i < b.Length; i += 4) if (b[i] != 255) return true;"
    Print #iFile, "    return false;"
    Print #iFile, "  }"
    Print #iFile, "  public static void Composite(byte[] b, int w, int h, int stride, int mode) {"
    Print #iFile, "    for (int y = 0; y < h; y++) {"
    Print #iFile, "      for (int x = 0; x < w; x++) {"
    Print #iFile, "        int i = y * stride + x * 4;"
    Print #iFile, "        int a = b[i+3];"
    Print #iFile, "        if (a == 255) continue;"
    Print #iFile, "        int bg = 255;"
    Print #iFile, "        if (mode == 2) { bg = ((((x >> 3) + (y >> 3)) & 1) == 0) ? 255 : 205; }"
    Print #iFile, "        b[i]   = (byte)((b[i]   * a + bg * (255 - a)) / 255);"
    Print #iFile, "        b[i+1] = (byte)((b[i+1] * a + bg * (255 - a)) / 255);"
    Print #iFile, "        b[i+2] = (byte)((b[i+2] * a + bg * (255 - a)) / 255);"
    Print #iFile, "        b[i+3] = 255;"
    Print #iFile, "      }"
    Print #iFile, "    }"
    Print #iFile, "  }"
    Print #iFile, "}"
    Print #iFile, "'@"
    Print #iFile, "Add-Type -TypeDefinition $cs"
    Print #iFile, "try {"
    Print #iFile, "  $src = New-Object System.Drawing.Bitmap($In)"
    Print #iFile, "  $w = $src.Width; $h = $src.Height"
    Print #iFile, "  $rect = New-Object System.Drawing.Rectangle(0,0,$w,$h)"
    Print #iFile, "  $fmt = [System.Drawing.Imaging.PixelFormat]::Format32bppArgb"
    Print #iFile, "  $work = New-Object System.Drawing.Bitmap($w,$h,$fmt)"
    Print #iFile, "  $g = [System.Drawing.Graphics]::FromImage($work)"
    Print #iFile, "  $g.Clear([System.Drawing.Color]::Transparent)"
    Print #iFile, "  $g.DrawImage($src,0,0,$w,$h)"
    Print #iFile, "  $g.Dispose(); $src.Dispose()"
    Print #iFile, "  $d = $work.LockBits($rect,[System.Drawing.Imaging.ImageLockMode]::ReadWrite,$fmt)"
    Print #iFile, "  $len = [Math]::Abs($d.Stride) * $h"
    Print #iFile, "  $buf = New-Object byte[] $len"
    Print #iFile, "  [System.Runtime.InteropServices.Marshal]::Copy($d.Scan0,$buf,0,$len)"
    Print #iFile, "  $had = [Post]::HasAlpha($buf)"
    Print #iFile, "  if ($Bg -gt 0) {"
    Print #iFile, "    [Post]::Composite($buf,$w,$h,[Math]::Abs($d.Stride),$Bg)"
    Print #iFile, "    [System.Runtime.InteropServices.Marshal]::Copy($buf,0,$d.Scan0,$len)"
    Print #iFile, "  }"
    Print #iFile, "  $work.UnlockBits($d)"
    Print #iFile, "  if ($Bmp -eq 1) {"
    Print #iFile, "    $flat = New-Object System.Drawing.Bitmap($w,$h,[System.Drawing.Imaging.PixelFormat]::Format24bppRgb)"
    Print #iFile, "    $g2 = [System.Drawing.Graphics]::FromImage($flat)"
    Print #iFile, "    $g2.Clear([System.Drawing.Color]::White)"
    Print #iFile, "    $g2.DrawImage($work,0,0,$w,$h)"
    Print #iFile, "    $g2.Dispose()"
    Print #iFile, "    $flat.Save($Out,[System.Drawing.Imaging.ImageFormat]::Bmp)"
    Print #iFile, "    $flat.Dispose()"
    Print #iFile, "  } elseif ($Bg -gt 0 -or $In -ne $Out) {"
    Print #iFile, "    $work.Save($Out,[System.Drawing.Imaging.ImageFormat]::Png)"
    Print #iFile, "  }"
    Print #iFile, "  $work.Dispose()"
    Print #iFile, "  if ($had) { exit 0 } else { exit 1 }"
    Print #iFile, "} catch { exit 2 }"

    Close #iFile
End Sub

' --- Small helpers -----------------------------------------------------------

Private Function RunHidden(ByVal sCmd As String) As Long
    Dim wsh As Object

    Set wsh = CreateObject("WScript.Shell")
    RunHidden = wsh.Run(sCmd, 0, True) ' 0 = hidden window, True = wait for it
End Function

Private Function Quote(ByVal s As String) As String
    Quote = Chr(34) & s & Chr(34)
End Function

' Dir raises an error rather than returning "" on a malformed name, so a bad
' file name falls through to the export and gets reported as a failure.
Private Function FileExists(ByVal sPath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(sPath) <> "")
    If Err.Number <> 0 Then
        Err.Clear
        FileExists = False
    End If
    On Error GoTo 0
End Function
