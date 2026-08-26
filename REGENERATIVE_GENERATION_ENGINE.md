# Regenerative generation engine

The project now treats uploaded papers and questions as exemplars of the underlying
mathematics rather than text templates to copy.

## Teacher uploads / reference papers
Math Teacher Aid infers:
- topic and learning outcome;
- assessment style and difficulty;
- representation/diagram/table structure.

It then regenerates fresh questions by changing at least three meaningful features where
possible: values, variables, dimensions, context, information order, diagram measurements,
sub-question structure or representation.

## Student typed/uploaded questions
Ask Math Advisor can generate a **fresh similar question** from:
- typed questions;
- uploaded images/screenshots/PDFs;
- photographed questions;
- handwritten/iPad question input.

The regenerated question keeps the same core mathematical skill while changing values,
variables, dimensions or context. The answer is not revealed.

## Topic-specific regeneration
- Algebra: coefficients/constants/variable letters change.
- Geometry/mensuration: dimensions change and diagrams are rebuilt consistently.
- Statistics: new coherent datasets/tables/graphs are generated.
- Trigonometry: graph/function parameters change while preserving the target concept.
- Similarity: new valid scale relationships are used.
- Real-life problems: contexts/names are varied rather than copied.

The existing deterministic table, graph and 3D diagram engines are preserved.
