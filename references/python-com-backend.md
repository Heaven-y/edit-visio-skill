# Optional Python COM Backend

Use Python when integrating Visio into an existing Python workflow, not as an
automatic cure for a blocked or hung PowerShell process. Diagnose the specific
failure before starting another application instance.

Requires Windows, licensed Visio and `pywin32` in the caller's approved environment.
Report missing dependencies and the proposed installation location first; do not
install packages into the skill repository.

## Session Contract

Use `win32com.client.DispatchEx("Visio.InvisibleApp")` for an owned session, not
`Dispatch` or `EnsureDispatch("Visio.Application")`, which can attach to the user's
interactive application. Keep COM work in one initialized apartment/thread.

Follow the same lifecycle as the shared PowerShell runtime:

1. Resolve paths and fail on an unavailable/locked target without closing user files.
2. Copy the target to a task-owned staging path before a destructive rebuild.
3. Draw native content; save only after successful drawing, never in `finally`.
4. Export a temporary preview, close owned documents, release references and quit.
5. Validate the staged VSDX and compare the rendered page against the reference.
6. Replace the target only after successful checks; remove owned temporary files.

Derive the real Windows PID from the application's window handle through
`win32process.GetWindowThreadProcessId`. Visio's `Application.ProcessID` is not the
Windows PID. Never terminate unrelated VISIO processes.

## Reconstruction Boundary

Read image dimensions through a raster API. OpenCV can help detect rectangles,
lines, grids and sampled colors, but does not determine scientific meaning.
Preserve editable labels, semantic groups and user-specified relationships.
Do not deliver automatic tracing without visual comparison and package checks.
