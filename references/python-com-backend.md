# Optional Python COM Backend

Use this backend when a PowerShell Visio session hangs, leaks references, or must
run for a long time. It is optional; ordinary short edits can use the bundled
PowerShell scripts.

## Requirements

- Windows with Microsoft Visio installed and registered for COM automation.
- Python 3.10 or newer in a dedicated environment.
- `pywin32` installed in that environment.

Do not install packages into the skill repository. Use the caller's approved
runtime or virtual environment and report any installation before running it.

## Session pattern

```python
from pathlib import Path
import win32com.client as win32

target = Path(r"C:\path\figure.vsdx").resolve()
visio = win32.gencache.EnsureDispatch("Visio.Application")
visio.Visible = False
doc = None
page = None
try:
    doc = visio.Documents.Open(str(target))
    page = doc.Pages.Item(1)
    # Draw or edit native Visio shapes here.
    doc.Save()
finally:
    if doc is not None:
        try:
            doc.Saved = True
            doc.Close()
        except Exception:
            pass
    page = None
    doc = None
    try:
        visio.Quit()
    except Exception:
        pass
    visio = None
```

Keep the COM session in one process and do not mix it with an already-open
interactive Visio instance. Suppress method return values, save before closing,
and inspect the target timestamp and lock state if the process times out.

## Image reconstruction boundary

OpenCV can detect coarse rectangles, lines, grids, or sampled colors. Use those
results as a geometry inventory only. For text-heavy scientific figures, supply
agent-authored annotations for labels, semantic groups, arrows, and image-like
motifs, then draw those elements as native Visio shapes. Do not deliver an
automatic parser output without visual inspection and package verification.
