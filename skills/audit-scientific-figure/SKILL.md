---
name: audit-scientific-figure
description: Review and score an existing scientific illustration in visible draw.io, Microsoft PowerPoint, or WPS Presentation without hiding defects through flattening. Use for publication-quality aesthetic review, style-direction comparison, reference fidelity, layout and whitespace cleanup, connector review, text fit, deep editability, raster atomicity, local-region gates, or repeated verification in a Designer-Drawer-Reviewer-Corrector loop.
---

# Audit Scientific Figure

Act as the Reviewer. Review read-only evidence and issue findings; do not draw during the review phase. A successful MCP call is not evidence that the figure is visually or structurally correct.

## Collect both evidence channels

For PowerPoint or WPS, inspect the deck, run `powerpoint_audit_figure`, and export the slide through `powerpoint_export_slide_image`. Record whether the renderer is Office.js PowerPoint, COM PowerPoint, or the OOXML fallback. Treat renderer differences as application-specific evidence, not permission to flatten editable content. In Office.js, treat `connector_mode=geometry_backed` and `implementation=officejs_editable_shape_composite` as declared limitations that still require visual routing and editability review.

For draw.io, inspect the live model, run `drawio_live_audit_figure`, and capture the current renderer through `drawio_live_screenshot`.

When a reference exists, inspect the full reference and the crop matching the current region. Compare at readable resolution.

When the figure is derived from a manuscript, paper PDF, method section, equations, or an evidence pack, read [Manuscript-to-Figure Workflow](../design-scientific-figure/references/manuscript-to-figure-workflow.md) completely. Obtain the frozen Figure Claim, Paper Figure Signature, required-node and required-edge ledgers, equation-operand ledger, evidence ledger, publication display width, and permitted omissions. Review the source contract and the artifact; do not infer a missing manuscript fact from the Drawer's visual intent.

## Apply the publication aesthetic gate

For a paper overview, graphical abstract, teaser, final whole-figure review, or comparison of candidate style directions, read [references/publication-aesthetic-review.md](references/publication-aesthetic-review.md) completely and apply it to the fresh render. This gate is mandatory even when deterministic structure audit reports zero hard failures.

Review the figure at thumbnail scale for silhouette and focal hierarchy, fit-to-slide scale for composition and rhythm, and readable scale for typography, spacing, connectors, and local finish. When multiple style candidates are presented, compare them together; palette, font, corner-radius, or border changes alone do not constitute different directions.

For the final whole-figure publication gate, use an independent fresh-eye Reviewer when available. Give it the semantic contract, selected reference and Style DNA, but do not provide the Drawer's defense of the current composition or a list of suspected defects.

When the declared style is hand-drawn, sketchnote, pencil, doodle, whiteboard, or Excalidraw-like, also read [Hand-drawn Technical Style](../design-scientific-figure/references/hand-drawn-technical-style.md) completely. Review scientific clarity and hand-drawn fidelity as separate questions. A clean diagram can pass semantics yet miss the requested style; a visibly rough diagram fails when its jitter, texture, font, or doodles reduce publication-width legibility or make topology ambiguous. Do not reward noise as authenticity.

## Review taxonomy

Review every region and the whole figure for:

1. scientific semantics, exact readable text, topology, and arrow direction;
2. editability coverage and meaningful object hierarchy;
3. raster irreducibility and decomposition metadata;
4. geometry, repeated alignment, equal spacing, margins, and whitespace;
5. text fit, wrapping, font hierarchy, and color consistency;
6. clipping, unintended overlap, z-order, and object bounds;
7. arrowhead clearance, connector path-through-object, label intersection, backtracking, and route crossings;
8. reference correspondence or no-reference design consistency.

For manuscript-derived figures, additionally require 100% coverage of non-omissible required nodes and edges. From a fresh render at the declared publication display width, reconstruct the positive edge table and every declared negative path. Check operands, axis scope, inverse/alignment operations, producer/model scope, training versus inference, and update/freeze ownership wherever they affect the scientific argument. Object ids and connector metadata are supporting evidence only; fail any route whose visible geometry changes or obscures the relation.

Inspect a title-hidden grayscale thumbnail at the declared review size. Confirm that the Figure Claim's mechanism or result is the first focal point and that overview/detail hierarchy survives without color. Re-run the affected render-scale checks after each structural correction.

Use the same categories and thresholds for draw.io and PowerPoint.

## Deep editability audit

Inspect pictures semantically, not only by object count. Fail the region when one image still contains separable content such as:

- a row or grid of predictions, masks, heatmaps, or microscopy fields;
- a before/after or method comparison;
- multiple independent photographs or channels;
- editable titles, labels, borders, arrows, legends, axes, tables, or regular plots.

Require one image object per irreducible visual field. Require the surrounding frame, grid, heading, legend, connector, and annotation to be separate editable objects. A crop that merely removes the outer panel border is not sufficient when the remaining crop is still composite.

Every retained image must have a precise reason, tight crop, `atomic_raster_unit=true`, `contains_reconstructable_content=false`, and a useful decomposition note.

## Finding format

Emit one record per defect:

```text
region: stable region id
objects: exact names/ids
category: shared defect category
severity: hard | warning
evidence: measurable structure or renderer observation
correction: required outcome, not vague advice
acceptance: condition the next audit can verify
confidence: 0..1 with evidence basis
```

Hard failures override averages. Treat wrong text, wrong direction, reconstructable content inside a picture, a non-atomic picture, clipping, arrow intrusion, a route through a label/object, and an unrelated connector crossing as hard failures.

For aesthetic findings, use the A/B/C classification and the location, defect, cognitive impact, severity, exact correction, and expected-effect fields required by the publication aesthetic reference. Do not hide a concrete aesthetic defect behind an averaged score.

## Scorecard

Score the affected region and whole figure from 0 to 1 for:

- semantic/text accuracy;
- editability coverage;
- geometry/alignment;
- spacing/whitespace;
- connector clarity;
- typography/color consistency;
- clipping/overlap safety;
- reference correspondence when applicable.

Pass only when readable semantics, reconstructable editability, and clipping/overlap safety equal 1.00; geometry and connector clarity are at least 0.95; reference correspondence is at least 0.90; deterministic audit has zero hard failures; and no warning remains except an explicitly documented source ambiguity. For manuscript-derived work, all non-omissible contract items and declared negative-path checks must also pass at the stated publication scale.

These numeric checks cover measurable structure, not publication beauty. A publication-aesthetic review passes only when no class-A issue remains, no class-B issue blocks the intended reading path, every large whitespace region has a clear compositional function, the primary content is not visibly underscaled, and the figure has one dominant focal hierarchy. Do not claim top-tier visual quality from numeric scores alone.

## Review loop

1. Review one completed region in whole-slide/canvas context.
2. Send findings to `$correct-scientific-figure`.
3. Let the backend Drawer execute the correction plan.
4. Collect a fresh structure audit and fresh render.
5. Review again from new evidence.
6. Approve the region only after the pass gate is satisfied.
7. After all regions pass, repeat for the whole figure.

Never approve based on the Drawer or Corrector claiming success. Never reuse a stale screenshot.

## Review report

Return region scores, whole-figure scores, all findings, native/composite/raster counts, every raster declaration, resolved finding ids, unresolved source ambiguities, and the final pass/fail verdict. For the publication aesthetic gate, also answer the four mandatory closing questions from the reference and identify the three changes with the highest visual impact.
