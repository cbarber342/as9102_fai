# AS9102 FAI (GUI)

This is a PySide6 desktop application for building AS9102 First Article Inspection (FAI) packages.

It includes:
- A main window for loading a Calypso `.chr` file, an Excel AS9102 template, and a Drawing PDF.
- A Drawing Viewer (PDF) with “bubble” annotations that can be saved/reopened and exported as real PDF annotations.

## Run

From the workspace root:

```bash
python -m as9102_fai
```

## Python versions

This app supports Python 3.8+ (tested with 3.8–3.12). Dependencies are pinned with
version markers to install compatible GUI wheels on both older and newer Python versions.

## Drawing Viewer notes

- Bubbles are editable overlays in the app and can be exported to the PDF as annotations.
- An edit sidecar is written next to your saved drawing PDF:
	- `<drawing>.pdf.as9102_bubbles.json`
	- This stores bubble locations/sizes, page rotations, and next bubble number.
	- It prevents “duplicate bubbles” when reopening a saved PDF.
- If a sidecar is not present, the viewer can import bubbles from existing PDF FreeText annotations (best-effort).

## Environment variables

- `AS9102_FAI_SAMPLE_DIR`: optional folder containing `sample_chr.txt`, `template.xlsx`, and `drawing.pdf` for auto-loading defaults.
- `AS9102_DEBUG_PDF=1`: enables extra PDF viewer debug logging.
