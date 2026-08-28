# Screenshot High Res

A SOLIDWORKS macro that exports the current view of a part or assembly to a PNG at
any pixel size you ask for, with a live preview of exactly what will be captured.

Works with SOLIDWORKS 2022, 2024 and 2025.

> **Want all ten macros at once?** This one is part of the [MacroShelf
> Collection](https://github.com/james-debono/macroshelf-collection-sw-macro-library) — a single
> download, already arranged as a library for the [MacroShelf](https://github.com/james-debono/macroshelf-sw-addin)
> add-in, so every macro lands on a SOLIDWORKS toolbar tab with its icon and
> hover text already set up.

## Why

SOLIDWORKS can already export images, but not comfortably:

- **Screen capture is capped at your monitor resolution.** No good for a slide or a
  printed document.
- **Print capture beats that**, but you specify it in DPI and paper sizes rather
  than pixels — so getting 1920 × 1080 means doing arithmetic.
- **You cannot see what will be captured** before you export. The framing is not
  the same as what's on screen.
- **Transparent background** is buried in system options, so it can't be changed
  per export.

This macro gives you a width and height in pixels, a live preview of the actual
framing, a transparent-or-white choice per export, and a save location that
remembers where you've been.

## Install

**The macro on its own:** download `Screenshot-High-Res.swp` from the
[latest release](../../releases/latest), then run it with **Tools > Macro > Run**,
or add it to a toolbar with **Tools > Customize > Commands > Macro**.

**With [MacroShelf](https://github.com/james-debono/macroshelf-sw-addin):** get the
[MacroShelf Collection](https://github.com/james-debono/macroshelf-collection-sw-macro-library/releases/latest),
which packages this macro with its icon and hover text alongside every other macro
in the set. Point MacroShelf at the unzipped folder and it appears as a button.

## One-time setup — required

**Tools > Options > System Options > Export > TIF/PSD/JPG/PNG**

> Tick **Remove background**, and leave it ticked permanently.

This is where the transparency comes from, and the macro genuinely cannot set it —
the option is absent from the SOLIDWORKS API, and writing its registry value has no
effect while SOLIDWORKS is running. Leaving it on costs you nothing: the macro
flattens onto white when you ask for a white background, and it warns you if it
notices the option has been switched off.

## Using it

1. Open a part or assembly and frame the view roughly.
2. Run the macro. The form opens and stays open.
3. **Orbit and zoom freely** — the form is modeless, and the preview catches up a
   moment after the view settles.
4. Set the width and height in pixels, pick transparent or white, choose where to
   save, and click **Export**.

The form doesn't close after exporting, so you can reframe and export again.
**Close it with the X — that's what ends the macro.**

### Things worth knowing

- The preview is a **real render**, not a crop of your screen. That's why it's
  accurate — the captured region genuinely isn't the same as the viewport.
- The **Refresh Image** button exists because the automatic refresh only watches the
  camera. Changing display style, hiding a component, switching configuration or
  editing the model won't trigger it.
- Don't resize the SOLIDWORKS window between previewing and exporting — framing
  follows the graphics area's shape, so the two would disagree.
- Exporting leaves print capture, User Defined paper and the DPI set in your export
  options. Worth knowing if a later manual File > Save As behaves unexpectedly.

## It shells out to PowerShell, and why

Worth stating plainly, because a macro that writes a `.ps1` and runs it is a
pattern corporate antivirus and EDR tools flag on sight.

This macro calls **PowerShell twice**, writing a short script to your `%TEMP%`
folder and running it hidden:

- **The folder picker.** The SOLIDWORKS API has no folder-browse dialog, and
  raising the Windows one in-process is unreliable from VBA.
- **Image conversion**, for the transparent-background path.

Both scripts are written by the macro, run once, and do nothing but the job named
above — no network access, no downloads, nothing persistent. Running out of
process is also a safety property: a mistake there takes down a throwaway
`powershell.exe` rather than SOLIDWORKS.

`%TEMP%` is per-user and access-controlled, so nothing here is exposed to other
users of the machine. If your IT department asks, this section is the answer.

## Not in scope

Drawings (sheet print options are untouched), photorealistic rendering, and batch
export.

## Building from source

`src\Screenshot-High-Res.vba` is the standard module and
`src\Screenshot-High-Res.Form.vba` is the form's code-behind. A `.swp` is a binary
VBA project, so it can only be produced from inside SOLIDWORKS — there is no build
step:

1. Open the `.swp` via **Tools > Macro > Edit**.
2. Paste the module source into the module, and the form source into the form's
   code window (right-click the form > View Code).
3. Save.

The **form layout** exists only inside the `.swp` and has no text source. Don't
import the form source as a `.frm` — that would overwrite the layout. And don't try
to patch the `.swp` directly: it stores compiled p-code ahead of the source, so a
patched file shows new code in the editor while still running the old.

Technical detail is in [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

## Licence

MIT — see [LICENSE](LICENSE). Free to use, modify and share. The full licence text
is also carried inside the macro itself, so a `.swp` passed on by itself still
carries its licence.

Created by James Debono, with AI assistance. Everything here was tested by
hand in SOLIDWORKS — nothing that touches the API can be verified any other way.

## Trademarks

SOLIDWORKS is a registered trademark of Dassault Systèmes SolidWorks Corporation.
This project is independent: it is not affiliated with, endorsed by, or sponsored
by Dassault Systèmes, and uses only the published SOLIDWORKS API.
