'==============================================================================
' Screenshot HD
'
' Exports the current view of a part or assembly to a PNG at any pixel size you
' ask for, with a live preview of exactly what will be captured.
'
' SOLIDWORKS can already export images, but not comfortably. Screen capture is
' capped at your monitor resolution. Print capture beats that, but is expressed
' in DPI and paper sizes rather than pixels. You cannot see what area will be
' captured before exporting, and the transparent background option cannot be
' changed per export.
'
' This gives a width and height in pixels, a preview of the real framing, a
' transparent or white background per export, and a save location that remembers
' where you have been.
'
' The form is modeless, so you can orbit and zoom while it is open and the
' preview catches up a moment after the view settles. It does not close after
' exporting - close it with the X, which is what ends the macro.
'
' Requires "Remove background" to be ticked in Tools > Options > System Options >
' Export > TIF/PSD/JPG/PNG, and left ticked. That option is absent from the
' SOLIDWORKS API, so the macro cannot set it. Leaving it on costs nothing: a
' white background is produced by flattening afterwards, and the macro warns if
' it finds the option has been switched off.
'
' To use, open a part or assembly and run the macro.
'
'   Version   0.6.2
'   Date      2026-08-20
'   Author    James Debono
'   Licence   MIT - full text below
'   Source    https://github.com/james-debono/screenshot-hd-sw-macro
'
' Requires the SOLIDWORKS and SOLIDWORKS Constant type libraries, which a
' SOLIDWORKS macro project references by default.
'
'------------------------------------------------------------------------------
' CHANGELOG (summary - see CHANGELOG.md for the full history)
'
'   0.6.2   Version shown in brackets in the title bar. Source URL updated.
'   0.6.1   Version shown in the form's title bar.
'   0.6.0   Renamed from "Export PNG" to "Screenshot HD". Licence and header.
'   0.5.x   Save location with recent folders; Windows folder picker; fixed the
'           macro never ending after the form was closed.
'   0.4.x   Pixel-size output, live preview, transparent or white background.
'   0.3.x   Modeless form and the idle loop that keeps it alive.
'   0.2.x   Export settings wired to the form.
'   0.1.0   Initial version. File name and Export.
'
'------------------------------------------------------------------------------
' MIT Licence
' SPDX-License-Identifier: MIT
'
' Copyright (c) 2026 James Debono
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'==============================================================================
'
' Design notes
'
' Export path
'   Everything goes through the standard SOLIDWORKS PNG exporter in Print
'   capture mode, which is the only mode that exceeds screen resolution.
'   Screen capture is documented as capturing "at the resolution of the screen
'   display" and offers no image size setting.
'
' Pixel size
'   The caller supplies a pixel width and height. Paper size is set to User
'   Defined and derived from those pixels at EXPORT_DPI, so paper x DPI resolves
'   exactly to the requested size:
'
'       paper inches = pixels / EXPORT_DPI
'
' Preview
'   Framing depends only on the paper aspect ratio, not its absolute size. The
'   preview uses the same paper at a much lower DPI, so it is framed identically
'   to the full export while rendering a fraction of the pixels. It is a real
'   render down the same path rather than a crop of the viewport, because
'   changing the paper aspect rescales the model in the view.
'
' Transparency
'   Produced by the "Remove background" option in Tools > Options > Export >
'   TIF/PSD/JPG/PNG, which must be left ticked. That option is absent from the
'   SOLIDWORKS API - the API help topic for that dialog lists it as "Currently
'   not available in the SOLIDWORKS API" - and writing its registry value has no
'   effect while SOLIDWORKS is running, as options are cached at startup and
'   flushed on exit. Exports are therefore always transparent, and an opaque
'   result is produced by flattening onto white afterwards.
'
' Save location
'   A folder is chosen on the form and remembered between sessions, along with a
'   short list of recently used folders. MODEL_FOLDER_TOKEN is a reserved list
'   entry meaning "beside the active document". Browse shows the standard
'   Windows folder picker, raised out of process because VBA cannot reach the
'   interface it needs - see BrowseFolder.
'
' Form lifetime
'   The form is modeless so the view can be manipulated while it is open. A
'   SOLIDWORKS macro ends as soon as its entry point returns, which would
'   destroy a modeless form immediately, so ShowSaveAsForm idles until the form
'   closes. That idle loop also drives the camera polling behind auto refresh.
'==============================================================================

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Must match the Version line in the header block above. build-library.ps1 checks
' that they agree and fails the build if they drift. Shown in the form's title bar.
Public Const MACRO_VERSION As String = "0.6.2"

Public Const CAPTURE_SCREEN As Long = 0
Public Const CAPTURE_PRINT As Long = 1

' DPI used for the full export. Paper size is derived from it, so this governs
' the size of the notional paper only - output pixels are set by the caller.
Public Const EXPORT_DPI As Long = 300

Public Const DEFAULT_WIDTH As Long = 1920
Public Const DEFAULT_HEIGHT As Long = 1080
Public Const MIN_PX As Long = 16
Public Const MAX_PX As Long = 10000

' Reserved entry in the location list meaning "the active document's folder"
Public Const MODEL_FOLDER_TOKEN As String = "<Model folder>"

' Auto refresh. AUTO_REFRESH False reverts to refreshing on the button only.
' SETTLE_SECONDS is how long the view must be still before a refresh fires, so
' that an orbit does not trigger a render per frame.
Public Const AUTO_REFRESH As Boolean = True
Public Const SETTLE_SECONDS As Single = 0.4

' How often the camera is read. The idle loop ticks every IDLE_MS, far more
' often than this needs checking, and each read costs several COM calls.
Public Const POLL_SECONDS As Single = 0.15

' Approximate width in pixels of the preview render
Private Const PREVIEW_PX As Long = 520
Private Const MIN_PREVIEW_DPI As Long = 6

' Idle loop period. Short enough to keep the view responsive, long enough that
' the loop does not saturate a CPU core.
Private Const IDLE_MS As Long = 10

' Stored settings, written under the standard VB and VBA program settings key
Private Const SETTINGS_APP As String = "Screenshot HD"
Private Const SETTINGS_SECTION As String = "Locations"
Private Const MAX_RECENT As Long = 8

' Background handling passed to the image helper
Public Const BG_KEEP As Long = 0
Public Const BG_WHITE As Long = 1
Public Const BG_CHECKER As Long = 2

' ExportToFolder results
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

' Cached preview render, so a change of background only re-composites
Private m_sPreviewPng As String
Private m_lPreviewW As Long
Private m_lPreviewH As Long

' Set by UserForm1 as it opens and closes. The idle loop below must key off this
' rather than off UserForm1 itself - see ShowSaveAsForm for why.
Public FormIsOpen As Boolean

' Entry point.
Sub ShowSaveAsForm()
    ' Clear anything left loaded by a previous run that was stopped rather than
    ' closed. Guarded by the count so that an unnecessary Unload cannot itself
    ' create a form.
    If UserForms.Count > 0 Then
        On Error Resume Next
        Unload UserForm1
        On Error GoTo 0
    End If

    FormIsOpen = True

    Load UserForm1
    UserForm1.Show vbModeless

    ' A SOLIDWORKS macro ends as soon as this procedure returns, which would
    ' take the modeless form down with it before it ever painted. Idling here
    ' keeps the macro alive until the form is closed.
    '
    ' DoEvents returns control to SOLIDWORKS so the view stays interactive; the
    ' Sleep prevents the loop from saturating a CPU core.
    '
    ' The exit condition is the flag the form sets, and the flag is re-tested
    ' before each call below. Referring to UserForm1 is what ends the loop's
    ' life, not what sustains it: touching a default instance re-creates it, so
    ' calling AutoRefreshTick after the form has closed silently loads a fresh
    ' hidden copy, puts UserForms.Count back to 1 and leaves this loop running
    ' forever - with an invisible form still polling the camera and rendering.
    ' That is what happened in 0.4.0 to 0.5.2, and it left the macro running
    ' after the window was closed, locking the file against editing.
    Do While FormIsOpen And UserForms.Count > 0
        DoEvents
        If Not FormIsOpen Then Exit Do

        Sleep IDLE_MS
        If Not FormIsOpen Then Exit Do

        UserForm1.AutoRefreshTick
    Loop

    FormIsOpen = False
End Sub

' --- Save location ----------------------------------------------------------

Public Function DefaultSaveFolder() As String
    DefaultSaveFolder = Environ("USERPROFILE") & "\Downloads"
End Function

' Turns a list entry into an actual folder path. Returns an empty string if it
' cannot be resolved - either nothing was chosen, or MODEL_FOLDER_TOKEN was
' chosen with no document open or a document that has never been saved.
Public Function ResolveSaveFolder(ByVal sChoice As String) As String
    Dim swModel As Object
    Dim sPath As String
    Dim lSlash As Long

    sChoice = Trim$(sChoice)
    If sChoice = "" Then Exit Function

    If StrComp(sChoice, MODEL_FOLDER_TOKEN, vbTextCompare) = 0 Then
        Set swApp = Application.SldWorks
        Set swModel = swApp.ActiveDoc
        If swModel Is Nothing Then Exit Function

        sPath = swModel.GetPathName
        If sPath = "" Then Exit Function ' never saved, so it has no folder yet

        lSlash = InStrRev(sPath, "\")
        If lSlash > 1 Then ResolveSaveFolder = Left$(sPath, lSlash - 1)
        Exit Function
    End If

    ' Drop a trailing separator so callers can append one unconditionally.
    ' Length 3 is a drive root such as C:\, which keeps its separator.
    If Right$(sChoice, 1) = "\" And Len(sChoice) > 3 Then
        sChoice = Left$(sChoice, Len(sChoice) - 1)
    End If

    ResolveSaveFolder = sChoice
End Function

Public Function FolderExists(ByVal sPath As String) As Boolean
    Dim fso As Object

    If Trim$(sPath) = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(sPath)
    If Err.Number <> 0 Then
        Err.Clear
        FolderExists = False
    End If
    On Error GoTo 0
End Function

' Creates a folder and any missing parents. Returns True if it exists afterwards.
Public Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim fso As Object
    Dim sParent As String

    If Trim$(sPath) = "" Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(sPath) Then
        CreateFolderTree = True
        Exit Function
    End If

    On Error Resume Next
    sParent = fso.GetParentFolderName(sPath)
    On Error GoTo 0

    If sParent = "" Or StrComp(sParent, sPath, vbTextCompare) = 0 Then Exit Function
    If Not CreateFolderTree(sParent) Then Exit Function

    On Error Resume Next
    fso.CreateFolder sPath
    Err.Clear
    On Error GoTo 0

    CreateFolderTree = fso.FolderExists(sPath)
End Function

' Shows the standard Windows folder picker - the Explorer style common item
' dialog with FOS_PICKFOLDERS - and returns the chosen folder, or an empty
' string if cancelled.
'
' Why it goes out to PowerShell. That dialog is IFileOpenDialog, a vtable only
' COM interface with no IDispatch, so VBA cannot reach it with CreateObject.
' SOLIDWORKS offers file dialogs but no folder picker, and the shell's
' BrowseForFolder is the older tree. A .NET host can declare the interface
' directly, and this macro already shells out to PowerShell for image work, so
' the dialog is raised there.
'
' Running it out of process is also a safety property worth keeping: a mistake
' in the interop takes down a throwaway powershell.exe rather than SOLIDWORKS.
'
' The result comes back through a file rather than stdout, because WScript.Shell
' Run cannot capture output and Exec cannot hide its console window.
Public Function BrowseFolder(ByVal sTitle As String, ByVal sStartFolder As String) As String
    Dim sScript As String
    Dim sResult As String
    Dim lRc As Long

    sScript = Environ("TEMP") & "\ScreenshotHD_pick.ps1"
    sResult = Environ("TEMP") & "\ScreenshotHD_pick.txt"

    On Error Resume Next
    Kill sResult
    On Error GoTo 0

    WriteFolderPicker sScript

    If sStartFolder = "" Then sStartFolder = DefaultSaveFolder()

    lRc = RunHidden("powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & _
                    Quote(sScript) & " -Start " & Quote(sStartFolder) & _
                    " -Title " & Quote(sTitle) & " -Out " & Quote(sResult))

    ' 0 picked, 1 cancelled, anything else failed
    If lRc = 0 Then BrowseFolder = ReadUnicodeLine(sResult)

    On Error Resume Next
    Kill sScript
    Kill sResult
    On Error GoTo 0
End Function

' Unlike the image helper, this script does contain double quotes - C# attribute
' arguments cannot be written without them - so they are escaped as "" below.
Private Sub WriteFolderPicker(ByVal sPath As String)
    Dim iFile As Integer

    iFile = FreeFile
    Open sPath For Output As #iFile

    Print #iFile, "param([string]$Start = '',[string]$Title = '',[string]$Out = '')"
    Print #iFile, "$ErrorActionPreference = 'Stop'"
    Print #iFile, "$cs = @'"
    Print #iFile, "using System;"
    Print #iFile, "using System.Runtime.InteropServices;"
    Print #iFile, "public static class Picker {"
    Print #iFile, "  [ComImport, Guid(""43826d1e-e718-42ee-bc55-a1e261c37bfe""), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]"
    Print #iFile, "  public interface IShellItem {"
    Print #iFile, "    void BindToHandler(IntPtr b, ref Guid s, ref Guid r, out IntPtr v);"
    Print #iFile, "    void GetParent(out IShellItem p);"
    Print #iFile, "    void GetDisplayName(uint sigdn, out IntPtr name);"
    Print #iFile, "    void GetAttributes(uint mask, out uint attribs);"
    Print #iFile, "    void Compare(IShellItem psi, uint hint, out int order);"
    Print #iFile, "  }"
    ' Every IFileDialog method must be declared, in order, or the vtable slots
    ' shift and the wrong function is called
    Print #iFile, "  [ComImport, Guid(""d57c7288-d4ad-4768-be02-9d969532d960""), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]"
    Print #iFile, "  public interface IFileOpenDialog {"
    Print #iFile, "    [PreserveSig] int Show(IntPtr parent);"
    Print #iFile, "    void SetFileTypes(uint c, IntPtr specs);"
    Print #iFile, "    void SetFileTypeIndex(uint i);"
    Print #iFile, "    void GetFileTypeIndex(out uint i);"
    Print #iFile, "    void Advise(IntPtr de, out uint cookie);"
    Print #iFile, "    void Unadvise(uint cookie);"
    Print #iFile, "    void SetOptions(uint fos);"
    Print #iFile, "    void GetOptions(out uint fos);"
    Print #iFile, "    void SetDefaultFolder(IShellItem si);"
    Print #iFile, "    void SetFolder(IShellItem si);"
    Print #iFile, "    void GetFolder(out IShellItem si);"
    Print #iFile, "    void GetCurrentSelection(out IShellItem si);"
    Print #iFile, "    void SetFileName(string name);"
    Print #iFile, "    void GetFileName(out IntPtr name);"
    Print #iFile, "    void SetTitle(string title);"
    Print #iFile, "    void SetOkButtonLabel(string text);"
    Print #iFile, "    void SetFileNameLabel(string label);"
    Print #iFile, "    void GetResult(out IShellItem si);"
    Print #iFile, "    void AddPlace(IShellItem si, int fdap);"
    Print #iFile, "    void SetDefaultExtension(string ext);"
    Print #iFile, "    void Close(int hr);"
    Print #iFile, "    void SetClientGuid(ref Guid g);"
    Print #iFile, "    void ClearClientData();"
    Print #iFile, "    void SetFilter(IntPtr filter);"
    Print #iFile, "    void GetResults(out IntPtr items);"
    Print #iFile, "    void GetSelectedItems(out IntPtr items);"
    Print #iFile, "  }"
    Print #iFile, "  [DllImport(""shell32.dll"", CharSet = CharSet.Unicode, PreserveSig = false)]"
    Print #iFile, "  private static extern void SHCreateItemFromParsingName("
    Print #iFile, "      string path, IntPtr bc, ref Guid iid,"
    Print #iFile, "      [MarshalAs(UnmanagedType.Interface)] out IShellItem item);"
    Print #iFile, "  private const uint FOS_PICKFOLDERS     = 0x20;"
    Print #iFile, "  private const uint FOS_FORCEFILESYSTEM = 0x40;"
    Print #iFile, "  private const uint SIGDN_FILESYSPATH   = 0x80058000;"
    Print #iFile, "  public static string Pick(string start, string title) {"
    Print #iFile, "    Type t = Type.GetTypeFromCLSID(new Guid(""DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7""));"
    Print #iFile, "    IFileOpenDialog d = (IFileOpenDialog)Activator.CreateInstance(t);"
    Print #iFile, "    uint opts;"
    Print #iFile, "    d.GetOptions(out opts);"
    Print #iFile, "    d.SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);"
    Print #iFile, "    if (title != null && title.Length > 0) d.SetTitle(title);"
    Print #iFile, "    if (start != null && start.Length > 0) {"
    Print #iFile, "      try {"
    Print #iFile, "        Guid iid = new Guid(""43826d1e-e718-42ee-bc55-a1e261c37bfe"");"
    Print #iFile, "        IShellItem si;"
    Print #iFile, "        SHCreateItemFromParsingName(start, IntPtr.Zero, ref iid, out si);"
    Print #iFile, "        if (si != null) d.SetFolder(si);"
    Print #iFile, "      } catch { }"
    Print #iFile, "    }"
    Print #iFile, "    int hr = d.Show(IntPtr.Zero);"
    Print #iFile, "    if (hr != 0) { Marshal.ReleaseComObject(d); return String.Empty; }"
    Print #iFile, "    IShellItem res;"
    Print #iFile, "    d.GetResult(out res);"
    Print #iFile, "    IntPtr p;"
    Print #iFile, "    res.GetDisplayName(SIGDN_FILESYSPATH, out p);"
    Print #iFile, "    string path = Marshal.PtrToStringUni(p);"
    Print #iFile, "    Marshal.FreeCoTaskMem(p);"
    Print #iFile, "    Marshal.ReleaseComObject(d);"
    Print #iFile, "    return path;"
    Print #iFile, "  }"
    Print #iFile, "}"
    Print #iFile, "'@"
    Print #iFile, "Add-Type -TypeDefinition $cs"
    Print #iFile, "try {"
    Print #iFile, "  $p = [Picker]::Pick($Start, $Title)"
    Print #iFile, "  if ($p -eq '') { exit 1 }"
    ' Unicode so that folder names outside the ANSI code page survive
    Print #iFile, "  Set-Content -LiteralPath $Out -Value $p -Encoding Unicode"
    Print #iFile, "  exit 0"
    Print #iFile, "} catch { exit 2 }"

    Close #iFile
End Sub

' Read back as UTF-16 to match what the picker writes, so folder names outside
' the ANSI code page are not mangled.
Private Function ReadUnicodeLine(ByVal sPath As String) As String
    Dim fso As Object
    Dim ts As Object

    On Error GoTo Bail

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(sPath) Then Exit Function

    Set ts = fso.OpenTextFile(sPath, 1, False, -1) ' ForReading, TristateTrue
    If Not ts.AtEndOfStream Then ReadUnicodeLine = Trim$(ts.ReadLine)
    ts.Close

    Exit Function

Bail:
    ReadUnicodeLine = ""
End Function

' --- Remembered locations ---------------------------------------------------
' Stored with SaveSetting/GetSetting, which needs no registry code. "Last" is
' the raw list entry last exported to, so the token is remembered as such;
' "Recent" is a bar separated list of real folder paths.

Public Function LastLocationChoice() As String
    LastLocationChoice = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Last", "")
End Function

Public Sub SaveLastLocationChoice(ByVal sChoice As String)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Last", sChoice
End Sub

Public Function RecentFolders() As String
    RecentFolders = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Recent", "")
End Function

' Moves a folder to the front of the recent list, de-duplicating and trimming.
Public Sub PushRecentFolder(ByVal sFolder As String)
    Dim vParts As Variant
    Dim sOut As String
    Dim lCount As Long
    Dim i As Long

    If Trim$(sFolder) = "" Then Exit Sub

    sOut = sFolder
    lCount = 1

    vParts = Split(RecentFolders(), "|")
    For i = LBound(vParts) To UBound(vParts)
        If lCount >= MAX_RECENT Then Exit For
        If vParts(i) <> "" And StrComp(vParts(i), sFolder, vbTextCompare) <> 0 Then
            sOut = sOut & "|" & vParts(i)
            lCount = lCount + 1
        End If
    Next i

    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Recent", sOut
End Sub

' --- Active document --------------------------------------------------------
' Re-read on every call. The form is modeless, so the document can be opened,
' closed or switched while it is displayed.

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
        ActiveDocName = "NewDocument" ' document has never been saved
        Exit Function
    End If

    ' Strip the extension, then everything up to the last backslash
    lDot = InStrRev(sFullPath, ".")
    If lDot > 0 Then sName = Left(sFullPath, lDot - 1) Else sName = sFullPath
    sName = Mid(sName, InStrRev(sName, "\") + 1)

    ActiveDocName = sName
End Function

' --- Export -----------------------------------------------------------------

' sFolder must already be resolved and known to exist; the form does that so it
' can offer to create a missing folder before any rendering happens.
' Returns one of the EXPORT_* constants.
Public Function ExportToFolder(ByVal sFolder As String, ByVal FN As String, _
                               ByVal pxW As Long, ByVal pxH As Long, _
                               ByVal bTransparent As Boolean) As Long
    Dim sPath As String
    Dim lPost As Long

    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc

    If Part Is Nothing Then
        ExportToFolder = EXPORT_NO_DOC
        Exit Function
    End If

    sPath = sFolder & "\" & FN & ".PNG"

    If FileExists(sPath) Then
        If MsgBox(FN & ".PNG already exists in:" & vbCrLf & sFolder & vbCrLf & vbCrLf & _
                  "Do you want to overwrite it?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "File Already Exists") <> vbYes Then
            ExportToFolder = EXPORT_CANCELLED
            Exit Function
        End If
    End If

    If Not RenderToPng(sPath, pxW, pxH, EXPORT_DPI) Then
        ExportToFolder = EXPORT_FAILED
        Exit Function
    End If

    ' An opaque result is flattened onto white. A transparent one is left as it
    ' is, but the file is still inspected to confirm it has an alpha channel -
    ' its absence means "Remove background" has been switched off.
    If bTransparent Then
        lPost = RunImageHelper(sPath, sPath, BG_KEEP, 0)
    Else
        lPost = RunImageHelper(sPath, sPath, BG_WHITE, 0)
    End If

    Select Case lPost
        Case 0 ' helper succeeded, source had transparency
            ExportToFolder = EXPORT_OK
        Case 1 ' helper succeeded, source was fully opaque
            If bTransparent Then
                ExportToFolder = EXPORT_NO_ALPHA
            Else
                ExportToFolder = EXPORT_OK
            End If
        Case Else
            ExportToFolder = EXPORT_FAILED
    End Select
End Function

' --- Preview ----------------------------------------------------------------

' Renders, or re-uses, a preview and returns the path of a BMP suitable for a
' form Image control. Returns an empty string on failure.
'
' Previews always go to the temp folder, so the chosen save location has no
' bearing on them and a bad path cannot break the preview.
'
' BMP rather than PNG because LoadPicture cannot read PNG.
Public Function PreviewToBmp(ByVal pxW As Long, ByVal pxH As Long, ByVal lBg As Long, _
                             ByVal bForce As Boolean) As String
    Dim sBmp As String
    Dim lDPI As Long

    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc
    If Part Is Nothing Then Exit Function

    If m_sPreviewPng = "" Then m_sPreviewPng = Environ("TEMP") & "\ScreenshotHD_preview.png"
    sBmp = Environ("TEMP") & "\ScreenshotHD_preview.bmp"

    ' Re-render only when the size changed or a refresh was requested. A change
    ' of background re-composites the existing render instead.
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
    Kill Environ("TEMP") & "\ScreenshotHD_preview.bmp"
    On Error GoTo 0

    m_sPreviewPng = ""
    m_lPreviewW = 0
    m_lPreviewH = 0
End Sub

' Same paper, lower DPI: identical framing, far fewer pixels.
Private Function PreviewDpi(ByVal pxW As Long) As Long
    Dim lDPI As Long

    lDPI = CLng((PREVIEW_PX * CDbl(EXPORT_DPI)) / pxW)
    If lDPI < MIN_PREVIEW_DPI Then lDPI = MIN_PREVIEW_DPI
    If lDPI > EXPORT_DPI Then lDPI = EXPORT_DPI

    PreviewDpi = lDPI
End Function

' --- SOLIDWORKS export ------------------------------------------------------

Private Function RenderToPng(ByVal sPath As String, ByVal pxW As Long, ByVal pxH As Long, _
                             ByVal lDPI As Long) As Boolean
    ApplyPrintCapture pxW, pxH, lDPI

    ' The window state is deliberately left alone. Framing follows the graphics
    ' area aspect, so resizing the window between a preview and an export would
    ' make the two disagree.
    longstatus = Part.SaveAs3(sPath, 0, 2) ' swSaveAsCurrentVersion, swSaveAsOptions_Copy

    RenderToPng = (longstatus = 0)
End Function

' Paper size is always derived from the requested pixels at EXPORT_DPI, so the
' aspect ratio - and therefore the framing - is identical for a preview and a
' full export. Only the DPI differs.
Private Sub ApplyPrintCapture(ByVal pxW As Long, ByVal pxH As Long, ByVal lDPI As Long)
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture, CAPTURE_PRINT
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swTiffPrintPaperSize, _
                                        swDwgPaperSizes_e.swDwgPapersUserDefined

    ' swconst assigns preference values 8 and 9 two names each, with width and
    ' height transposed between the pairs:
    '
    '   swTiffPrintPaperWidth  = 8       swTiffPrintDrawingPaperHeight = 8
    '   swTiffPrintPaperHeight = 9       swTiffPrintDrawingPaperWidth  = 9
    '
    ' Only one pairing can be correct. The names below are those prescribed by
    ' the API help topic "System Options > Export > TIF/PSD/JPG/PNG". If exports
    ' emerge transposed, exchanging these two statements is the entire fix.
    swApp.SetUserPreferenceDoubleValue swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperWidth, _
                                       PixelsToMetres(pxW, EXPORT_DPI)
    swApp.SetUserPreferenceDoubleValue swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperHeight, _
                                       PixelsToMetres(pxH, EXPORT_DPI)

    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swTiffPrintDPI, lDPI
End Sub

Private Function PixelsToMetres(ByVal px As Long, ByVal lDPI As Long) As Double
    PixelsToMetres = (CDbl(px) / lDPI) * METRES_PER_INCH
End Function

' --- Camera -----------------------------------------------------------------

' Returns a fingerprint of everything affecting the framing, for comparison on
' each idle tick. Orientation and translation cover orbit and pan, Scale2 covers
' zoom, and the frame size covers a window resize - which alters the graphics
' area aspect and so the framing, without the camera having moved.
'
' Returns an empty string when there is no readable view, which callers treat as
' "nothing to compare".
Public Function CameraSignature() As String
    Dim swModel As Object
    Dim swView As Object
    Dim vOrient As Variant
    Dim vTrans As Variant
    Dim i As Long
    Dim s As String

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Exit Function

    Set swView = swModel.ActiveView
    If swView Is Nothing Then Exit Function

    On Error GoTo Bail

    vOrient = swView.Orientation2
    vTrans = swView.Translation2

    For i = LBound(vOrient) To UBound(vOrient)
        s = s & Format$(vOrient(i), "0.000000") & ";"
    Next i

    For i = LBound(vTrans) To UBound(vTrans)
        s = s & Format$(vTrans(i), "0.000000") & ";"
    Next i

    s = s & Format$(swView.Scale2, "0.000000") & ";"
    s = s & swView.FrameWidth & "x" & swView.FrameHeight

    CameraSignature = s
    Exit Function

Bail:
    CameraSignature = ""
End Function

' --- Image helper -----------------------------------------------------------
' Flattens transparency, builds the checkerboard preview backdrop, converts to
' BMP, and reports whether the source carried an alpha channel.
'
' Returns 0 succeeded, source had transparency
'         1 succeeded, source was fully opaque
'         2 failed

Private Function RunImageHelper(ByVal sIn As String, ByVal sOut As String, _
                                ByVal lBg As Long, ByVal lBmp As Long) As Long
    Dim sScript As String

    sScript = Environ("TEMP") & "\ScreenshotHD_post.ps1"
    WriteImageHelper sScript

    RunImageHelper = RunHidden("powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & _
                               Quote(sScript) & " -In " & Quote(sIn) & " -Out " & Quote(sOut) & _
                               " -Bg " & lBg & " -Bmp " & lBmp)
End Function

' Emitted at run time so the macro remains self contained, with no companion
' files to deploy. The per pixel loop is compiled C# via Add-Type, as an
' equivalent PowerShell loop is prohibitively slow at 3000x3000.
'
' The script contains no double quote characters, so it requires no escaping
' here. Preserve that property when editing it.
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

' --- Utilities --------------------------------------------------------------

Private Function RunHidden(ByVal sCmd As String) As Long
    Dim wsh As Object

    Set wsh = CreateObject("WScript.Shell")
    RunHidden = wsh.Run(sCmd, 0, True) ' hidden window, wait for completion
End Function

Private Function Quote(ByVal s As String) As String
    Quote = Chr(34) & s & Chr(34)
End Function

' Dir raises an error rather than returning an empty string when given a
' malformed name, so an invalid file name falls through to the export and is
' reported there as a failure.
Private Function FileExists(ByVal sPath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(sPath) <> "")
    If Err.Number <> 0 Then
        Err.Clear
        FileExists = False
    End If
    On Error GoTo 0
End Function
