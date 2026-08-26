# Geometry Construction Board component fix

## Problem
The previous implementation used `components.html(...)`, which creates a static iframe.
That iframe can display HTML but is not a true two-way Streamlit component. As a result:
- drawing was unreliable;
- Apple Pencil/touch handling was inconsistent;
- tool icons were ambiguous;
- `Save to notes` could not reliably return the canvas image to Python.

## Fix
The board is now a proper Streamlit custom component using the component-v1 protocol.

### Visible labelled tools
- Pencil
- Line / straightedge
- Compass
- Protractor
- Ruler
- Eraser
- Undo
- Clear
- optional Grid

### Input
Uses Pointer Events, so it supports:
- Apple Pencil pressure
- iPad touch
- mouse

### Construction behaviour
- Compass draws a circle from centre + radius.
- Protractor creates a baseline, second ray, angle arc and measured angle.
- Ruler draws a measured segment with endpoint ticks.
- Straightedge draws clean line segments.

### Saving
`Save to notes` returns a PNG through the Streamlit component channel. Python decodes the
PNG and appends it to the student's persistent lesson notes, where the existing Word/PDF
notes export can include it.

No additional Python dependency is required.
