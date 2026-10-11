# Excel2Img

Excel2Img watches folders for saved `.xlsx` and `.xlsm` workbooks, finds tables
enclosed by cell borders, and asks Microsoft Excel for Mac to render each table
as a high-resolution PNG. It is a small, local command-line utility; it does not
upload workbooks or execute workbook macros.

## Install (macOS)

Requirements: Python 3.10+ and Microsoft Excel for Mac.

```bash
cd Excel2Img
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp excelDirectoryList.example.json excelDirectoryList.json
```

Edit `excelDirectoryList.json` and replace the example paths with the folders
that contain your workbooks. `~`, spaces, Korean text, and multiple folders are
supported. Missing folders produce a warning. Nested and duplicate paths are
collapsed automatically.

On the first render, macOS may ask Terminal (or your terminal application) for
permission to control Microsoft Excel. Allow it in **System Settings → Privacy &
Security → Automation**. Excel must be installed, but it does not have to be
open before Excel2Img starts.

## Run

Process all existing eligible workbooks once:

```bash
python main.py --once
```

Then watch continuously (stop with Ctrl+C):

```bash
python main.py
```

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

For rendering, Excel2Img opens a temporary copy read-only, asks Excel to copy
each detected range as a printer-quality picture, reads the vector PDF flavor
from the macOS pasteboard, and rasterizes it at 300 DPI. The source workbook is
never edited or saved. Output files are atomically replaced.

## Known limitations

- A table needs a closed outer border. A gap in the outline makes inside versus
  outside geometrically ambiguous and the table may not be detected.
- Two outlined tables sharing a full cell edge are indistinguishable from one
  divided table and are exported together. Corner-only contact stays separate.
- Rendering uses the macOS pasteboard briefly, so it replaces current clipboard
  contents. Excel automation requires a logged-in graphical macOS session.
- Password prompts, external-link prompts, protected content, and extremely
  large ranges can prevent Excel from copying a picture. Macros are not run by
  Excel2Img; temporary copies are opened read-only with link updates disabled.
- Rich text fidelity is determined by Excel's own Copy Picture output.

## Tests and verification

```bash
pytest
```

The automated suite verifies detection, exclusions, Unicode paths, stable
output/cleanup, configuration, and watcher behavior without needing Excel. The
Excel automation integration can only be fully verified on a logged-in Mac that
has granted Automation permission; failures are reported without stopping the
watcher.

Current implementation validation (2026-10-11): the full Excel pipeline was
run successfully on the installed Microsoft Excel for Mac using a synthetic
workbook with Korean text, a merged cell, font/fill formatting, and Greek and
superscript characters. It also succeeded while the source workbook remained
open, and Excel's automation-security setting was restored afterward. The
private study workbook named in the original requirements was not present and
was therefore not tested or committed.
