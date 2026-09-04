---
name: design-scientific-figure
description: Design a new scientific illustration, graphical abstract, workflow, architecture figure, or mechanism diagram from a brief or manuscript without a reference image, for visible draw.io, Microsoft PowerPoint, or WPS Presentation. Use when a clean editable publication layout, manuscript-to-figure translation, structurally distinct style directions, planned connector lanes, consistent visual grammar, and Designer-to-Drawer-to-Reviewer-to-Corrector quality gates are required.
---

# Design Scientific Figure

Act as the Designer in the four-role You-Only-Figure-Once protocol. Produce a backend-neutral design specification before the Drawer adds any object. The selected backend affects object mapping, not the design quality or acceptance gate.

## Detect constraints

Read the selected backend's capabilities first. Use `powerpoint_status` then `powerpoint_get_capabilities` for PowerPoint/WPS, or `drawio_live_get_capabilities` for draw.io. When live Mac PowerPoint is requested, also require `powerpoint_officejs_status` to report a connected task pane. Design only with semantic objects the selected adapter can create editably; use declared editable composites when a native monolithic object is unavailable.

## Define the message

Record:

- one sentence stating the scientific message;
- intended audience and reading order;
- required stages, entities, evidence, comparisons, and causal links;
- target aspect ratio and output size;
- labels or facts that must remain exact;
- uncertainty or content still needing user input.

## Derive from a manuscript

When the source is a manuscript, paper PDF, method section, equations, supplementary text, or an evidence pack, read [Manuscript-to-Figure Workflow](references/manuscript-to-figure-workflow.md) completely before freezing the design. Use it for ordinary manuscript-driven design and for blind Figure Gym runs.

Freeze the Figure Claim, Paper Figure Signature, required-node ledger, required-edge ledger, equation-operand ledger, and evidence ledger before choosing a direction. Derive the publication slot and composition family from that contract before reference retrieval. In blinded independent design, also require a target-paper and derivative-asset exclusion receipt before retrieved assets are materialized.

Do not enter publication drawing until the selected direction maps every required node and edge or records a source-grounded permitted omission. A missing operand, condition, inverse/alignment step, evidence producer, or update owner is a contract failure rather than a styling choice.

## Build the layout system

Specify before drawing:

- outer margins, panel grid, gutters, and shared alignment anchors;
- panel ids, bounds, hierarchy, and construction order;
- standard node sizes, corner radii, stroke widths, and spacing tokens;
- title, section, label, annotation, and caption typography;
- a limited semantic palette with accessible contrast;
- reserved connector lanes and permitted entry/exit sides;
- legend, annotation, table, chart, and raster-evidence locations;
- z-order and meaningful grouping.

Prefer one clear reading path: left-to-right, top-to-bottom, or an explicitly labeled cycle. Use no more than three primary hierarchy levels.

## Define publication aesthetics before drawing

For a publication-facing figure, graphical abstract, style comparison, or any request involving aesthetics, visual quality, or beautification, read [Publication Aesthetic Review](../audit-scientific-figure/references/publication-aesthetic-review.md) completely before freezing the `design_spec`.

For each direction, define a `style_dna` that changes drawing decisions, not just surface decoration:

- composition skeleton and reading path;
- dominant focal zone and primary silhouette at thumbnail scale;
- relative visual mass of scientific novelty, process, and real evidence;
- shape grammar, line behavior, arrow convention, icon language, and depth treatment;
- typography roles, semantic palette, and contrast hierarchy;
- whitespace strategy, including the intended function of every large empty region;
- forbidden motifs that would create a generic PPT, dashboard, poster, or AI-generated appearance.

Treat whitespace as functional only when it separates hierarchy, protects a focal point, reserves a connector lane, or improves grouping. Do not shrink the scientific content to preserve empty bands that have no stated function. Plan the meaningful-content hull inside the usable canvas and reject a layout when the main content can be enlarged substantially without harming hierarchy or connector clarity.

When proposing multiple style directions, compare their `style_dna` before drawing. Two directions are cosmetic variants if they preserve the same composition skeleton, reading path, focal zone, major module proportions, and evidence placement while changing only palette, font, border, or corner radius. A distinct direction must change the composition skeleton or reading path and at least one of focal strategy, evidence-to-mechanism weighting, or shape/line grammar, with a scientific communication reason for the change. Reject cosmetic variants before the Drawer creates slides. Also compare equal-size grayscale thumbnails with style titles hidden; if their large-scale mass distribution and first focal point remain interchangeable, return the directions for redesign.

## Design connectors

- Route process flow through reserved orthogonal lanes with few bends.
- Connect from the side facing the destination and avoid immediate backtracking.
- Keep arrows outside unrelated shapes and labels.
- Separate parallel routes by a consistent lane gap.
- Avoid crossings; if scientific topology makes one unavoidable, redesign the node placement before accepting it.
- Distinguish process, inhibition, feedback, grouping, and association with consistent conventions.
- Reserve endpoint clearance so arrowheads touch a boundary without covering the target fill or text.

## Plan editability

Classify every planned element as editable text, shape, line, connector, table/chart, repeated motif, or irreducible raster evidence. Split any multi-image evidence block into one atomic image per field and plan its title, border, grid, legend, arrows, and annotations as editable objects.

## Produce the design handoff

Return a `design_spec` containing:

- canvas/slide geometry and layout tokens;
- region and object ids with bounds and styles;
- exact text and scientific topology;
- connector source, target, sites, waypoints, lanes, and arrow convention;
- grouping and z-order;
- raster decomposition declarations;
- artifact mode (`direction_review` or `publication`), final display size, and figure archetype;
- `style_dna`, focal hierarchy, evidence weighting, and forbidden motifs;
- usable canvas bounds, meaningful-content hull, intended canvas-utilization range, and the function of each planned large whitespace region;
- an evidence budget for overview figures that states the placement and relative visual mass of representative real results, or the scientific reason no real evidence is available;
- for manuscript-derived work, source locations, Figure Claim, Paper Figure Signature, complete node/edge/operand/evidence ledgers, publication display width, permitted omissions, and any blind-exclusion receipt;
- when alternatives are requested, a style-divergence comparison showing why each direction is structurally distinct;
- local construction order and acceptance conditions.

Do not improvise geometry one object at a time after drawing begins.

## Enter the construction loop

Hand the design to `$edit-powerpoint-live` or `$recreate-scientific-figure-in-drawio`. After each region, require `$audit-scientific-figure`; when it finds a defect, require `$correct-scientific-figure`, return the object-level plan to the Drawer, rerender, and review again. For an overview, graphical abstract, teaser, final whole figure, or style-direction comparison, require the Reviewer to apply its publication aesthetic reference before approval.

Finish only when every local region and the whole figure have exact readable semantics, full reconstructable editability, no clipping or unintended overlap, layout and connector confidence of at least 0.95, and no unresolved audit finding except documented content ambiguity. For manuscript-derived work, also require complete contract coverage plus positive- and negative-edge reconstruction from a fresh render at the declared publication width and a title-hidden grayscale hierarchy check.

## Delivery

Save the editable source and requested exports. Report the selected backend, design tokens, object counts, raster declarations, local gates, whole-figure gate, and remaining content ambiguity.

In `publication` mode, remove style numbers, reference-DNA labels, selection instructions, reviewer notes, and other direction-review scaffolding from the final figure.
