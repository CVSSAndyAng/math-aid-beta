# Question–diagram relevance fix

The example showed a polygon-angle question with an unrelated Cartesian line graph.

This build adds semantic diagram validation.

## Polygon questions
Interior/exterior-angle and regular-polygon questions now use:
- a polygon diagram matching the stated number of sides, or
- no diagram if none is required.

They cannot use an unrelated coordinate graph.

## Validation
Before displaying the generated paper, Math Teacher Aid audits whether the diagram type
matches the question topic. Common mismatches are flagged for regeneration.

## Together with the previous formatting fixes
The generator is also reminded that:
- prose must stay as prose;
- maths uses MathIO;
- raw backslashes/LaTeX commands must not be visible;
- duplicate givens/instructions should be removed;
- survey/frequency information should use tables.

The regenerative generation behavior remains enabled.
