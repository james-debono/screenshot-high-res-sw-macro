# Changelog

Semantic Versioning. `MAJOR` reaches 1 when the behaviour is settled enough
to promise not to break it; 0.x is an honest statement that it may still move.

---

Not every version has a `.swp` — **0.4.0** and **0.5.2** were superseded before
being built. Source-only versions are normal here.

This macro has been renamed twice. It was **Export PNG** up to and including
0.5.3, **Screenshot HD** from 0.6.0 to 0.6.2, and **Screenshot High Res** from
0.7.0 — "HD" named one specific resolution, and the macro exports at any pixel
size well past it.

## 0.7.0 — 2026-08-21

- **Renamed from "Screenshot HD" to "Screenshot High Res."** "HD" names one
  specific resolution, and this macro exports at any pixel size well past it.
- Moved to its own repository, with the `Source` URL updated to match.
- **Saved settings start fresh.** The registry key holding the recent-folders
  list is named after the macro, so renaming it leaves the old list behind.
- No functional change.

## 0.6.2 — 2026-08-20

- The version in the form's title bar is now shown in brackets — `Screenshot HD
  (0.6.2)` — so it reads as metadata rather than part of the name.
- The `Source` URL in the header now points at the renamed repository,
  `screenshot-hd-sw-macro`.
- No functional change.

---

## 0.6.1 — 2026-08-13

### Added
- The version now appears in the form's title bar, so it's visible whenever the
  macro is used rather than only by opening the code. Set from a `MACRO_VERSION`
  constant that `build-library.ps1` checks against the header.

  This also corrects the title bar, which still read "Export PNG" — the caption
  came from the form designer, which the 0.6.0 rename couldn't reach because the
  layout lives only inside the `.swp`.

No functional change.

---

## 0.6.0 — 2026-08-09

### Changed
- **Renamed from "Export PNG" to "Screenshot HD".** The macro name, the filename,
  the folder name and the toolbar button now all agree. A minor rather than a
  patch release because the name is user-visible.
- The saved-settings registry key moved from `Export PNG` to `Screenshot HD`.
  **Your remembered save location and recent folders reset once** on first run
  after upgrading; everything works normally from then on.
- Temporary helper scripts renamed `ScreenshotHD_*` to match. They are written to
  `%TEMP%` and deleted after each run, so nothing is left behind by the old names.

### Added
- Released under the **MIT licence**, with the full text carried in the code
  itself so a `.swp` passed on by itself still carries its licence.
- Header rewritten to explain what the macro does and why, including the
  "Remove background" requirement. Design notes kept below it.

No change to how the export works.

---

## 0.5.3 — 2026-08-06

### Fixed
- The macro kept running after the form was closed, holding the `.swp` open
  against editing and making the next launch fail with
  `swRunMacroError_UserInterrupt` (11).

  The idle loop called `UserForm1.AutoRefreshTick` without first checking that the
  form still existed. Referring to a VBA default instance re-creates it, so
  closing the window loaded a fresh hidden copy, put `UserForms.Count` back to 1
  and left the loop running forever — with an invisible form still polling the
  camera and firing renders in the background.

  The loop now keys off a `FormIsOpen` flag that the form clears in `QueryClose`
  and `Terminate`, and re-tests it before every call back into the form.
- `ShowSaveAsForm` unloads anything left over from a run that was stopped rather
  than closed, guarded by `UserForms.Count` so the cleanup cannot itself create a
  form.

### Affects
- 0.4.0 through 0.5.2. Introduced in 0.4.0 with auto refresh, which added the
  first reference to `UserForm1` inside the loop. 0.3.2 and earlier are clean —
  that loop touched nothing but `DoEvents` and `Sleep`.

---

## 0.5.2 — 2026-08-06

### Changed
- Browse now shows the standard Windows folder picker — the Explorer-style common
  item dialog with `FOS_PICKFOLDERS` — opened at whatever the location box
  currently points to. This replaces 0.5.1's file dialog, and with it the
  awkwardness of selecting a file in order to choose a folder, including the
  empty-folder problem that came with it.

### How
- That dialog is `IFileOpenDialog`, a vtable-only COM interface with no
  `IDispatch`, so VBA cannot reach it via `CreateObject`. A .NET host can declare
  it directly, and the macro already shells out to PowerShell for image work, so
  the dialog is raised there and the path returns through a UTF-16 file.
- Out of process is a deliberate safety property, not just convenience: a mistake
  in the interop takes down a throwaway `powershell.exe` rather than SOLIDWORKS.
- The interop was verified before shipping by reading the dialog options back
  after setting them — `0x1868`, with `FOS_PICKFOLDERS` and `FOS_FORCEFILESYSTEM`
  both present — which confirms the vtable slots line up.

---

## 0.5.1 — 2026-08-06

### Changed
- Browse now uses `ISldWorks::GetOpenFileName`, the same dialog SOLIDWORKS opens
  files with, instead of the shell's folder tree. It looks native, and it is
  seeded with the current location and file name so it opens in the right place
  rather than at the desktop root as the old one did.

### Trade-off
- SOLIDWORKS exposes no folder picker, only file dialogs, so the folder is taken
  from whichever file is selected. Picking from a completely empty folder
  therefore depends on the dialog accepting the seeded file name rather than
  insisting the file exists. If it does not, the path can still be typed or
  pasted into the box, and the recent list covers repeat use.

---

## 0.5.0 — 2026-08-06

### Added
- Save location on the form. An editable dropdown with a Browse button, replacing
  the hard-coded Downloads folder.
- `<Model folder>` as a reserved list entry, exporting beside the active document.
  Reported clearly when the document has never been saved and so has no folder
  yet.
- Recent folders, the last eight, most recent first. The last choice is remembered
  between sessions and stored as typed, so the model folder entry comes back as
  itself rather than as whatever it resolved to.
- Missing folders are offered for creation, including missing parents.

### Changed
- `ExportToDownloads` is now `ExportToFolder` and takes the folder as an argument.
- The location is resolved and checked before any rendering, so a bad path fails
  immediately rather than after a full-size export.

### Details worth knowing
- SOLIDWORKS exposes file dialogs but no folder picker, so Browse uses
  `Shell.Application`. That dialog always opens at the desktop root, which is why
  the recent list carries the weight for repeat use.
- The busy guard is held for the duration of the Browse dialog. The idle loop
  keeps running while it is open, so without that an automatic refresh would fire
  underneath it.
- Previews still go to the temp folder, so the save location cannot affect them
  and a bad path cannot break the preview.

---

## 0.4.1 — 2026-08-06

### Changed
- Both modules now open with a proper header block — title, purpose, version,
  author, updated date and a Licence line — instead of instructions for pasting
  the file into the VBA editor. The code reads as source now, not as a delivery
  note.
- Installation instructions moved to the README, which is where they belong.
- Comments reworded throughout to read as maintainer documentation rather than as
  a conversation. No functional change; the code itself is identical to 0.4.0.

### Note
- Auto refresh polls only while the form is open. The idle loop that drives it
  exits when the form closes, so nothing runs in the background.

---

## 0.4.0 — 2026-08-06

*Source only — superseded by 0.4.1 before a `.swp` was built.*

### Added
- The preview refreshes itself a moment after the view stops moving. The macro's
  existing idle loop reads a camera fingerprint — orientation, translation, zoom
  and graphics area size — every 150 ms, restarts a settle clock on any change,
  and fires one refresh once the view has been still for 400 ms. Orbit, pan, zoom
  and window resizes are all picked up.
- Typing in Width or Height uses the same settle clock, so it waits until you stop
  typing rather than rendering at "19" on the way to 1920.
- `AUTO_REFRESH`, `SETTLE_SECONDS` and `POLL_SECONDS` at the top of the module to
  tune it or switch it off.

### Why polling rather than DModelViewEvents
- The event route works — `DModelViewEvents` has `ViewChangeNotify` — but it needs
  re-hooking whenever the document changes, and it fires per frame during a drag,
  so it would need the same settle logic on top anyway. Polling has neither
  problem and reuses a loop that already exists.

### Details worth knowing
- The camera fingerprint is re-baselined after every render. The print capture can
  nudge the view, and without that the render would look like another camera move
  and refresh forever.
- Automatic refreshes never show a message box, so a half-typed width cannot pop a
  dialog mid-keystroke. `RefreshPreview` now takes separate "re-render" and
  "report problems" flags rather than conflating the two.

---

## 0.3.2 — 2026-08-06

### Fixed
- The form did not appear at all in 0.3.1, with no error. A SOLIDWORKS macro ends
  as soon as its entry sub returns; `Show vbModeless` returns immediately, so the
  macro ended and the VBA project was torn down, destroying the form before it
  painted. `ShowSaveAsForm` now idles in a `DoEvents` loop until the form is
  closed. `DoEvents` keeps SOLIDWORKS interactive; a 10 ms `Sleep` stops the loop
  spinning a CPU core flat.
- Added a busy guard on Export and Refresh. The idle loop pumps messages and the
  export can too, so without it a second click part way through would have started
  a second export on top of the first.

### Notes
- SOLIDWORKS considers the macro to be running while the form is open. Stopping
  the macro closes the form, which is a safe way out.

---

## 0.3.1 — 2026-08-06

### Fixed
- The form is now modeless, so the model can be orbited and zoomed while it is
  open. It was modal, which locked the view — the whole point of the preview is to
  reframe and re-check, and that was impossible.

### Changed, as consequences of going modeless
- The form no longer closes itself after exporting, so you can reframe and export
  again without reopening it. Close it with the X.
- Nothing about the document is cached when the form opens. Open, close or switch
  documents while it is up and the export follows the active one.
- The Export and Refresh buttons are no longer disabled when no document is open
  at load time — they check when clicked instead, since a document can now be
  opened after the form.
- The file name fills in from the active document on refresh as well as at load,
  if the box is empty.
- `ActiveDocName` guards against a path with no extension, which would have thrown
  on `Left(path, -1)`.

---

## 0.3.0 — 2026-08-06

### Added
- Output size in pixels. Width and Height boxes drive a User Defined paper size
  worked out from the pixels at 300 DPI, so paper × DPI lands exactly on what you
  typed. 1920×1080 through 3000×3000 and anything between.
- Live preview. Refresh Image renders through the same Print capture path at low
  DPI — identical framing, about 1/13 the pixels for a 1920-wide export — and
  shows it in the form. Transparent areas appear as a checkerboard.
- Transparent Background / White Background choice. Applied after the export, so
  switching between them re-composites the existing preview instead of
  re-rendering.
- Warning when Transparent is picked but the exported file has no alpha channel,
  which means Remove background has been unticked in Tools > Options. The macro
  checks the finished file rather than trusting a setting it cannot read reliably.

### Changed
- Print capture is now always used — it is the only way to beat screen resolution.
  The Screen/Print radio buttons and the DPI dropdown are gone from the form; DPI
  is derived, not chosen.
- The macro no longer maximises the model window. Framing follows the graphics
  area aspect, so changing the window between preview and export would make them
  disagree.

### Why this shape
- Remove background is not in the API and a registry write is ignored while
  SOLIDWORKS is running, so it is left ticked permanently and the macro flattens
  for the opaque case instead. The gap stops mattering.
- Framing depends only on the paper aspect ratio, not its size — A3-Landscape and
  A4-Landscape frame identically. But changing the aspect visibly rescales the
  model, so the preview has to be a real render; a computed crop of a screen grab
  would have been confidently wrong.

---

## 0.2.1 — 2026-08-03

### Fixed
- Remove Background, Screen/Print Capture and DPI now apply the moment they are
  changed on the form instead of only on export, so closing the form without
  exporting still leaves the settings changed.
- Export no longer overwrites an existing PNG silently. It prompts Yes/No
  (defaulting to No) and leaves the form open if you decline, so you can change
  the name.
- `SaveAsPNG` is now a Function returning True/False, so the form stays open when
  an export is cancelled or fails.
- Remove Background writes are now read back and verified; the form reports a
  failed write instead of failing quietly.
- Existence check goes through a `FileExists` helper, because `Dir()` raises an
  error rather than returning `""` on a malformed file name.

### Corrected
- DPI is documented as "For drawings only; For Output as - Print capture only" —
  it does nothing when exporting a part or assembly.
- Remove Background sits in the "Output as" group, not the print capture group, so
  it stays enabled for both capture modes.
- Confirmed against the SOLIDWORKS API help topic "System Options > Export >
  TIF/PSD/JPG/PNG" that Remove Background is "Currently not available in the
  SOLIDWORKS API", so the registry is the only route. Ruled out
  `swTIFExportIncludeDrawingsPaperColor`, which looks like a match by elimination
  but is a different setting.

---

## 0.2.0 — 2026-08-03

### Added
- Wired the new form controls to the PNG export settings:

  | Control | Setting |
  |---|---|
  | Screen Capture / Print Capture | `swTiffScreenOrPrintCapture` |
  | DPI | `swTiffPrintDPI` |
  | Remove Background | `"TIF Remove background"` registry value |

- Form loads the current export settings when it opens and writes them back on
  export.
- DPI list populated with the values the SOLIDWORKS export dialog offers, and
  greyed out for a screen capture.
- File name is validated as non-empty before exporting.

### Note
- The form itself — the extra controls — was added by hand in the `.swp` before
  this; 0.2.0 is where the code behind them arrived.

---

## 0.1.0 — 2026-06-04

- Form with file name and Export. Pre-fills the active document's name and saves
  it to `%USERPROFILE%\Downloads\<name>.PNG` using whatever was already set in
  Tools > Options > Export > PNG.
