# Screenshot High Res — development notes

The API findings behind this macro. `README.md` covers what it does and how to
use it.

## How the export works

Screen capture is capped at monitor resolution, so **print capture is the only
route past it**. The macro sets User Defined paper and a DPI that multiply out to
the pixel size asked for, exports, then restores nothing — see limitations.

**Framing depends only on aspect ratio, not on absolute size.** That is what makes
the live preview accurate: the preview is a real render at a small size, and the
export is the same framing at the requested size. It also means resizing the
SOLIDWORKS window between previewing and exporting makes the two disagree, because
framing follows the graphics area's shape.

## API findings

- **"Remove background" is not in the API at all.** There is no user preference
  for it, and writing its registry value has no effect while SOLIDWORKS is
  running. The macro cannot set it, which is why the README asks the user to tick
  it once by hand. The macro flattens onto white when a white background is
  requested, and warns if it notices the option is off.
- **There is no folder picker in the SOLIDWORKS API.** The Windows dialog is
  called through COM instead.
- The preview must be a **real render**, not a crop of the viewport — the
  captured region genuinely is not the same as what is on screen.

## Traps

- **Auto-refresh only watches the camera.** Changing display style, hiding a
  component, switching configuration or editing the model does not trigger it,
  which is why there is a manual **Refresh Image** button.
- **The form is modeless and its lifetime ends with the X.** An earlier version
  never ended the macro after the form closed.
- **Editing a `.swp` is a paste, not an import.** Importing a `.frm` overwrites
  the form layout, which exists only inside the `.swp` and has no text source.

## Known limitations

- Exporting leaves print capture, User Defined paper and the DPI set in the
  export options. A later manual **File > Save As** may behave unexpectedly.
- Drawings are out of scope; sheet print options are untouched.
- No photorealistic rendering and no batch export.

## Verification status

Confirmed working in SOLIDWORKS, including the transparency path and the preview
matching the exported framing.

## There is no build step

A `.swp` is a binary VBA project. Editing the `.vba` in `src\` changes nothing
that runs until the source is pasted into the SOLIDWORKS VBA editor and saved.
Treat the `.vba` as the source of truth and re-paste after every change.

Do not patch a `.swp` directly: it stores compiled p-code ahead of the source
text and VBA runs the p-code, so a patched file shows new code in the editor
while still running the old.
