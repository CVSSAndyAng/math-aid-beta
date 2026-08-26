# Exam-style 3D diagram construction update

The Paper Setter now prioritises deterministic dimensioned diagrams for standard
mensuration solids instead of generic oversized wireframe scenes.

Supported clean diagram types include:
- cylinder
- cone
- cuboid
- cube
- triangular prism
- cylinder + cuboid
- prism + cuboid
- cylinder + prism
- cylinder + cuboid + triangular prism
- cone with a smaller similar cone cut from the top

The renderer follows the supplied SVG logic:
- clean perspective faces;
- ellipse-based circular faces;
- radius lines;
- arrowed dimension lines;
- labels beside the correct measurements;
- restrained exam-style colours/outlines.

Generic `diagram_scene_3d` rendering remains only as a fallback for unusual spatial
geometry that the deterministic solid engine cannot represent.

PWA/iPad Home Screen assets are included in this package.
