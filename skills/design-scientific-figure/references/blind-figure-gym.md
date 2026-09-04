# Blind Figure Gym Protocol

Use this reference only when the user requests an independently designed scientific figure whose existing target figure must remain unseen until after the new artifact is frozen, or asks for a blind-design comparison.

The goal is to test scientific-communication decisions without visual leakage. It is not a reconstruction task, a memory test, or a pixel-similarity benchmark.

## 1. Seal the target before retrieval

Record a human-readable exclusion receipt containing:

- a stable paper identity such as DOI, arXiv/OpenReview id, exact title, and canonical source;
- the allowed pre-freeze source set;
- the sealed target figure, caption, alt text, layout-descriptive prose, slides, blog diagrams, thumbnails, and derivative assets;
- the publication slot, audience, backend, and requested deliverables.

Choose the allowed source set to match the claim being tested. A paper-level design may use manuscript text and equations; an implementation-level design may use official code. Label the resulting scope explicitly. Do not infer that a code-scoped artifact should reproduce conceptual branches shown only in the paper.

If retrieval cannot exclude the target or derivatives before results are materialized, stop calling the run blind. Continue as an ordinary reference-aware redesign and say why.

## 2. Freeze the scientific contract

Before selecting a layout, write and freeze:

- one Figure Claim;
- a Paper Figure Signature describing topology, novelty, evidence scope, reading order, and scientific boundaries;
- required-node and required-edge ledgers with source locations;
- equation operands, conditions, inverse/alignment operations, and update ownership;
- evidence producers and the claims each evidence item may support;
- negative paths whose accidental appearance would change the science.

Resolve conflicts from the allowed sources before drawing. Record any uncertainty that affects topology instead of hiding it in a generic label.

## 3. Compare directions without target leakage

Develop at least two low-fidelity directions only when the request benefits from alternatives. They must differ in composition skeleton or reading path and in at least one of focal strategy, evidence-to-mechanism weighting, or shape/line grammar.

Select a direction by its fit to the Figure Claim and publication slot. Reject a direction that makes an optional path appear mandatory, collapses distinct operands, obscures an update owner, or needs unreadably small text to fit.

## 4. Pass the pre-reveal gate

The artifact may be frozen only when all of the following are true:

- every required node and relation is visible or has a source-grounded permitted omission;
- positive relations and negative paths can be reconstructed from a fresh renderer export at the declared display width;
- no hard semantic, clipping, text-fit, connector, or raster-atomicity defect remains;
- the editable source exists and its object structure has been inspected;
- the latest render and the editable source name the same frozen design.

Record the freeze date, file paths, backend, object counts, raster declarations, and audit verdict. A human-readable receipt is sufficient; do not add hashes unless they replace a materially more expensive check and change a later decision.

Do not open the sealed target until this receipt exists. After reveal, never overwrite the blind source or render to make the result look retrospectively closer.

## 5. Reveal and compare decisions

Attribute the target figure and record the exact source viewed. Compare the two artifacts along these dimensions:

- scientific claim and causal topology;
- source scope and abstraction level;
- first focal point and reading path;
- mechanism detail versus communication compression;
- real evidence, schematic evidence, and what each may claim;
- ambiguity, conditions, feedback, and training/inference boundaries;
- legibility at the intended publication size;
- editability and reuse.

State what the target does better, what the blind design does better, and whether any difference is caused by source scope rather than design quality. Treat the published target as a strong authored solution, not an infallible ground truth. Pixel resemblance is not a primary success criterion.

Give an explicit verdict in prose. Do not hide judgment in one composite score. If the blind design contains a scientific error, attribute it to the earliest responsible artifact: source scope, contract, direction, draw, or review.

## 6. Preserve learning without overfitting

Keep the frozen artifact and the post-reveal comparison as separate deliverables. If the user asks for an improved design after comparison, save it as a clearly named post-reveal version.

Promote a lesson into the skill only when it generalizes across cases or prevents a severe scientific error. Promote the decision rule, not the example's node count, dimensions, palette, typography, or silhouette.

## Deliverables

A complete blind comparison includes:

1. source authority and exclusion receipt;
2. frozen scientific contract and selected design specification;
3. frozen editable source and fresh review render;
4. pre-reveal audit verdict;
5. attributed target source;
6. post-reveal comparison and explicit verdict;
7. any generalized skill change, separately identified from the case artifact.
