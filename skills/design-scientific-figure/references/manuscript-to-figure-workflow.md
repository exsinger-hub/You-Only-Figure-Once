# Manuscript-to-Figure Workflow

Use this reference when the source of a new scientific figure is a manuscript, paper PDF, method section, equations, supplementary text, or an evidence pack rather than a short free-form brief. It also governs blind Figure Gym runs in which a target figure is sealed until after an independent design is frozen.

The purpose is not to make the generated figure resemble an existing one. The purpose is to preserve the manuscript's scientific argument in a figure whose visible topology can be independently reconstructed at publication size.

## 1. Declare source authority and mode

Before extracting content, record:

- exact manuscript/PDF version and source locations;
- whether the task is ordinary design, reference-led reconstruction, or blinded independent design;
- which captions, figures, supplementary files, code, and evidence assets are allowed;
- which target figures or derivative assets are sealed or excluded;
- intended publication slot, final display width, and output formats.

In blinded independent design, do not inspect the target figure, its crop, its alt text, or layout-descriptive caption passages before the design is frozen. If reference retrieval is used, exclude the target paper and derivative assets before candidates are materialized. Use stable identifiers available for the paper, such as DOI, arXiv/OpenReview id, exact title, and canonical source URL; do not rely on title substring matching alone when a stable id exists.

Do not apply the exclusion rule to an ordinary redesign in which the user explicitly supplied the target figure as a reference.

## 2. Build the scientific contract

Freeze these artifacts before selecting a visual direction:

### Figure Claim

Write one sentence stating what a reader must understand after viewing the figure. A list of modules is not a Figure Claim.

### Paper Figure Signature

Record:

- causal and data-flow topology;
- scientific novelty and intended first focal point;
- input, output, condition, supervision, state, operator, and evidence roles;
- training-only, inference-only, shared, updated, and frozen boundaries;
- overview/detail hierarchy;
- required real evidence and its scientific scope;
- publication slot and reading order.

Derive the needed composition family from this signature before retrieving references. Search for role-specific anchors for structure, mechanism, evidence, and visual language; do not choose a whole template first and force the manuscript into it.

### Required-node ledger

For every node, record a stable id, exact label or notation, semantic type, source location, condition membership, visual priority, and whether omission is allowed. If omission is allowed, give a source-grounded reason.

### Required-edge ledger

For every relation, record a stable id, source node, target node, relation type, direction, condition, source location, and visual encoding. Include scientifically meaningful absent relations as negative-path checks when a false connection would change interpretation.

### Equation-operand ledger

For each equation that affects topology, conditioning, alignment, optimization, or update ownership, record:

- output quantity;
- all operands and operators;
- index or axis scope when it changes the operation;
- parameter sharing or independence;
- train/inference status;
- planned visual representation or an explicit reason it remains textual.

Do not silently reduce an operand-bearing relation to a module label. A missing conditioning operand, inverse mapping, alignment step, weighted contribution, or update target is a source-contract failure.

### Evidence ledger

For every proposed raster or empirical result, record producer/model stage, sample-level or paper-level scope, allowed claim, source binding, crop/atomicity declaration, and whether it is representative, quantitative, or schematic. Proximity, a shared frame, color, or a leader line must not imply a stronger claim than the ledger authorizes.

## 3. Freeze a direction before publication drawing

When alternatives are useful, keep them low fidelity. Compare equal-size, title-hidden grayscale skeletons. Directions must differ structurally as required by the main design skill; do not build full presentation pages for candidates that can be rejected from a skeleton.

For every candidate considered selectable, map every required node and edge to a visible object or route. The selected direction must reach complete required-node and required-edge coverage before the Drawer begins publication-fidelity work. Do not replace a missing edge with nearby prose.

Derive the canvas ratio and type/evidence budgets from the intended publication slot. Do not default to 16:9. Record the final display width used for review. Any minimum text or evidence size is task-specific and must be declared in the design spec; do not reuse a pixel threshold from another paper without justification.

Build only the selected direction at publication fidelity unless the user explicitly requests multiple finished alternatives.

## 4. Review the visible artifact, not only its object graph

The final review needs both current structure evidence and a fresh renderer export.

At the declared publication display width:

1. identify every required node from the render;
2. reconstruct every required positive edge from visible endpoints, direction, and route;
3. test declared negative paths, including misleading shared containment or leader lines;
4. verify that text and evidence meet their declared minimum size;
5. verify that training/inference and update/freeze semantics do not depend on color alone.

Then inspect an equal-size title-hidden grayscale thumbnail for silhouette, first focal point, hierarchy, and overview/detail separation. A connector id, correct source/target metadata, or zero deterministic findings is supporting evidence, not a substitute for visible-edge reconstruction.

Re-run the affected scale tests after every structural correction. Do not upgrade a verdict from a local geometry fix without refreshing the same evidence that originally failed.

## 5. Attribute and correct the earliest deviation

Classify a failure at the earliest stage where the correct information or relation was lost:

- `parse`: manuscript/equation/evidence extraction omitted or altered a fact;
- `spec`: the frozen contract contradicts the source or itself;
- `retrieval`: references introduced an incompatible topology or leaked the target;
- `direction`: the chosen composition flattened or merged required semantics;
- `draw`: implementation changed visible geometry, routing, text, or evidence binding;
- `review`: available evidence contained the defect but the Reviewer approved it.

Fix the earliest responsible artifact first, then regenerate only the affected downstream objects. Do not conceal an upstream contract defect with a terminal label, decorative cue, or raster replacement.

## 6. Optional blinded comparison

After a blind artifact is frozen and passes its pre-reveal gate, reveal the source figure and compare:

- scientific semantics and causal topology;
- composition, hierarchy, and evidence integration;
- publication legibility, consistency, and editability.

Do not use pixel similarity as the primary verdict. Treat the published figure as a strong reference, not an infallible ground truth. Preserve both generated and original advantages, and attribute each failure to its earliest stage.

Promote a workflow rule only when it recurs across cases or prevents a severe scientific error, and only in a generalized form. Never promote case-specific node counts, edge counts, aspect ratios, pixel thresholds, palettes, or silhouettes as universal rules.
