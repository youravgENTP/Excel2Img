# Excel2Img

Excel2Img renders bordered tables from the workbook currently being edited in
Microsoft Excel for Mac as high-resolution PNGs. Rendering is triggered manually,
so an ordinary Cmd+S only saves the workbook and never starts a conversion. It is
a small, local command-line utility; it does not upload workbooks or execute
workbook macros.

## Install (macOS)

Requirements: Python 3.10+ and Microsoft Excel for Mac.

```bash
cd Excel2Img
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp excelDirectoryList.example.json excelDirectoryList.json
```

`excelDirectoryList.json` is needed only for the optional `--once` and `--watch`
modes. Edit it and replace the example paths with the folders that contain your
workbooks. `~`, spaces, Korean text, and multiple folders are supported.

On the first render, macOS may ask Terminal (or your terminal application) for
permission to control Microsoft Excel. Allow it in **System Settings → Privacy &
Security → Automation**. Excel must be installed. The manual command requires
Excel to be frontmost and a previously saved workbook to be active.

## Run

Render only the active Excel workbook:

```bash
python main.py --active
```

Excel2Img confirms that Excel is frontmost without bringing it forward. If the
workbook has changes, it performs a normal save of that workbook; it never uses
Save As. An unchanged workbook is not saved again. A never-saved or read-only
modified workbook produces an error without prompting, discarding changes, or
processing another workbook. The saved file is allowed to settle before it is
read. Concurrent `--active` invocations are safely coalesced: only the first
runs, and the kernel releases the lock even if that process crashes.

Check the complete Excel automation path first. If a workbook is already open,
this copies its active cell directly without activating or navigating anything.
If Excel has no open workbook, the check creates and removes an isolated tiny
workbook. Both paths verify PNG conversion without modifying a document:

```bash
python main.py --check
```

Diagnose or render one workbook without reading the directory configuration:

```bash
python main.py --file "/absolute/path/to/workbook.xlsx"
```

Process all existing eligible workbooks once:

```bash
python main.py --once
```

Legacy automatic watching is available only when explicitly requested (stop
with Ctrl+C):

```bash
python main.py --watch
```

Running `python main.py` without a mode prints help and does not start a watcher.
Batch and watcher startup run the same preflight automatically. If Excel stops
responding, a batch aborts instead of sending every remaining workbook to a
broken automation session. Resolve any visible Excel dialog and run `--check`
again; Excel2Img never force-quits Excel or closes the user's workbooks.

## macOS keyboard shortcut

Create a Quick Action in the Shortcuts app, add **Run Shell Script**, and use
this exact script:

```bash
cd /Users/younyung.gene/Excel2Img
./.venv/bin/python main.py --active
```

The command uses absolute project and interpreter paths, needs no activated
virtual environment or interactive shell, writes failures to standard error,
and does not open Terminal or another application window. Assign the desired
keyboard shortcut in macOS. If using Cmd+Shift+N, remove or reassign Finder's
existing shortcut in macOS settings so Finder does not intercept it; the CLI
cannot override application keyboard shortcuts.

For `Notes.xlsx`, images are written beside it in `Notes_Img/`, for example
`Sheet1_01.png`. Files and worksheets whose names begin with `!`, Excel lock
files beginning with `~$`, and generated `_Img` folders are ignored. Repeated
saves replace the same images; obsolete generated images are removed using a
private manifest, while unrelated files are left alone.

## How detection and rendering work

The detector treats visible top/bottom/left/right cell borders as barriers. It
finds regions fully enclosed by those barriers and groups edge-adjacent enclosed
cells. This preserves unbordered and blank cells inside an outlined table and
keeps diagonally touching tables separate. A lone decorative line is not an
enclosure and is ignored.

For active-workbook rendering, Excel2Img finds the source workbook that is
already open in Excel and sends `copy picture` directly to each range.
Microsoft Excel for Mac supports this without activating the workbook or sheet
and without selecting or navigating to the range. The active workbook, active
sheet, selected cell, scroll position, and keyboard focus therefore remain
unchanged.

When Excel has no workbook open (for example, a standalone `--once` run),
Excel2Img places a uniquely named copy inside Excel's private macOS sandbox
temporary directory and opens it read-only. This avoids the macOS “Grant
Access” dialog that blocks AppleScript for files in the system temporary
directory. It asks Excel for printer-quality pictures, reads the vector PDF
flavor from the macOS pasteboard, and rasterizes it at 300 DPI. The source is
never edited or saved. Output files are atomically replaced.

If the requested source is closed while some other workbook is open,
Excel2Img deliberately refuses to open a temporary copy because doing so would
change the user's active Excel context. Open and save that source workbook, or
close all Excel workbooks before running `--file`/`--once`.

Rendering an already-open workbook does not change Excel application settings.
For a closed `.xlsm` opened only when Excel is otherwise empty, the renderer
reads the current automation-security value, temporarily disables macros while
opening the copy, then restores that exact observed value in Python. It never
changes `display alerts`. Timeout recovery closes a workbook only when its
random temporary name is present; it cannot target a user workbook.

## Known limitations

- A table needs a closed outer border. A gap in the outline makes inside versus
  outside geometrically ambiguous and the table may not be detected.
- Two outlined tables sharing a full cell edge are indistinguishable from one
  divided table and are exported together. Corner-only contact stays separate.
- Rendering uses the macOS pasteboard briefly, so it replaces current clipboard
  contents. Excel automation requires a logged-in graphical macOS session.
- `--active` requires Excel to remain frontmost while the command starts. It
  does not activate Excel if another application is frontmost.
- Password prompts, external-link prompts, protected content, and extremely
  large ranges can prevent Excel from copying a picture. Macros are not run by
  Excel2Img; temporary copies are opened read-only with link updates disabled.
- A modal dialog already open in Excel can block AppleScript. Excel2Img reports
  the failed phase and performs bounded recovery, but it does not dismiss user
  dialogs or terminate Excel automatically.
- A closed source is not rendered while another workbook is open. This is an
  intentional nonintrusion guarantee rather than a rendering limitation.
- Rich text fidelity is determined by Excel's own Copy Picture output.

## Tests and verification

```bash
pytest
```

The automated suite verifies active-workbook selection and save decisions,
locking, detection, exclusions, Unicode paths, stable output/cleanup,
configuration, and opt-in watcher behavior without needing Excel. The
Excel automation integration can only be fully verified on a logged-in Mac that
has granted Automation permission; failures are reported without stopping the
watcher.

Current implementation validation (2026-10-11): direct `copy picture` from a
non-active worksheet and from a workbook other than the active workbook both
succeeded in the installed Microsoft Excel for Mac. Before/after checks showed
identical active workbook, active worksheet, selected cell, scroll row, and
scroll column. Ten real tables were then rendered from the workbook being
edited with the same state preserved. The frontmost application was Microsoft
Word before and after a full `--check`, confirming that keyboard focus was not
stolen. The `--active` frontmost guard was also exercised with Code in front: it
exited nonzero, did not render, and left Code frontmost. A live read of Excel's
active workbook returned its exact saved path and state without bringing Excel
forward, and the ordinary-save AppleScript compiled successfully without being
executed. A successful `--active` render was not forced during that check because
Excel was not frontmost; the command deliberately refused to change that state.
The private workbooks are not committed.
