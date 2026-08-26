# MathIO / question structure fix

This update addresses the formatting problems shown in the generated papers.

## Fixed rules
- English sentences remain normal text.
- Only mathematical expressions are rendered through MathIO.
- Raw LaTeX/backslash commands must not appear visibly.
- Roots/fractions/powers use proper mathematical notation.
- Units in prose display normally, e.g. `7 cm`, not `{cm}` or `text{{cm}}`.
- Survey response options are displayed as a table instead of one italic equation line.
- Construction instructions stay as prose.
- Redundant 'given' equation lines are suppressed when the same information is already in the question text.
- The generator is instructed not to duplicate an instruction in both the stem and a subpart.

The regenerative engine remains enabled, so uploaded/reference questions still generate
fresh variants with different values, variables, dimensions or contexts.
