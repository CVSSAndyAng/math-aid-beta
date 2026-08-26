# Geometry board runtime component fix

The repeated Streamlit Cloud error was caused by the custom component being registered
from a repository folder that was not present in the deployed app:

`No such component directory`

This version removes that deployment dependency.

The complete component frontend is embedded directly inside `app.py`. At runtime the app:
1. creates a writable temporary directory;
2. writes the embedded `index.html` into it;
3. registers that temporary directory as the Streamlit component frontend.

Therefore, for this fix, replacing `app.py` is sufficient. The app does not require a
`geometry_board_component/` directory to exist in the deployed repository.

The board still retains:
- two-way Streamlit communication;
- Apple Pencil/touch/mouse pointer handling;
- Pencil, Line, Compass, Protractor, Ruler, Eraser, Undo and Clear;
- Save to notes as a PNG.
