# Examination-standard composite solid diagrams

The cylinder-on-cuboid renderer has been upgraded from a generic generated illustration
to a cleaner technical/examination style.

Changes:
- cylinder is proportioned more naturally relative to the cuboid;
- the cylinder is centred on the actual top face;
- the hidden lower circular face is no longer drawn through the cuboid;
- only the visible front contact arc is shown;
- perspective is consistent across the cuboid faces;
- dimension arrows are moved outside the object;
- short extension lines connect dimensions to the measured edges;
- radius uses a light leader from the top circular face;
- labels are kept away from the solid;
- the overall drawing is more compact with less unused white space.

The Gemini generation prompt now also contains an EXAM-DIAGRAM QUALITY CONTRACT so future
generated visual specifications follow the same design standard.
