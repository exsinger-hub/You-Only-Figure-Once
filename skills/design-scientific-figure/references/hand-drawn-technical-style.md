# Hand-drawn Technical Style

Use this reference when a scientific figure is requested in a hand-drawn, sketchnote, pencil, doodle, whiteboard, or Excalidraw-like style. The goal is controlled human character without sacrificing scientific topology, publication-width readability, or native editability.

## The governing split

Treat hand-drawn character as a visual layer, not as permission to make scientific structure imprecise.

- **Exact layer:** required nodes, labels, equations, connector endpoints, arrow direction, alignment anchors, panel bounds, evidence attribution, and training/inference boundaries.
- **Expressive layer:** title lettering, section marks, native doodle icons, highlighter swashes, limited secondary outlines, and small asymmetries that do not alter topology.

If the expressive layer makes an exact relation harder to recover at the declared publication size, reduce or remove the expressive treatment. Do not add random jitter merely to increase visible roughness.

## Choose one style family

### `restrained_technical_handdrawn`

Default for paper overviews, method figures, architecture diagrams, and README examples that must remain publication-like.

- strict macro grid with one dominant focal module;
- off-white canvas, dark ink, one cool accent, and one warm accent;
- handwritten display face only for the title, section marks, and short callouts;
- technical labels, formulas, and dense annotations in a clean sans serif;
- simple grouped doodles made from native primitives;
- exact single-stroke semantic connectors;
- no global texture, shadow, glow, or arbitrary rotation.

### `sketchnote_map`

Use for teaching posters, talk summaries, broad concept maps, or a large display surface where exploration matters more than compact architecture reconstruction.

- freer regions, annotations, and visual metaphors;
- a clearly declared reading path despite looser grouping;
- handwriting may be more prominent, but exact technical labels remain typeset when density rises;
- reject this family when the final paper slot forces dense handwriting or ambiguous crossing routes.

### `pencil_illustrative`

Use only when the user explicitly values an illustrative pencil treatment and the backend can preserve legibility and editability. It is not the default for a technical paper figure.

- pencil character must come from native strokes or a tightly scoped irreducible illustration;
- never place a whole-slide paper or pencil raster over reconstructable content;
- formulas, metrics, axes, legends, and semantic arrows remain native and crisp;
- fall back to `restrained_technical_handdrawn` when pencil texture becomes the dominant signal.

## Style-source firewall

External examples may teach visual language; they do not become scientific authority.

For every style source, record:

- source title, author, date, and URL;
- the transferable traits, such as highlighter hierarchy, two-color ink, doodle density, or line restraint;
- the non-transferable content: labels, topology, characters, composition-specific metaphors, and source-paper claims;
- whether the source is human-drawn, digitally hand-drawn, or merely rough-styled when that distinction is known.

Do not call a figure independently designed if the target overview, caption, alt text, layout description, or close derivative was inspected before design freeze. When the task is style adaptation rather than recreation, copy no source-specific arrangement.

## Design-spec additions

Add these fields to `style_dna`:

```text
style_family: restrained_technical_handdrawn | sketchnote_map | pencil_illustrative
imperfection_budget:
  expressive_regions: named titles, callouts, doodles, or focal outlines
  exact_regions: named connectors, formulas, dense labels, evidence frames
  maximum_rotation_or_offset: explicit small bounds, or none
typography_split:
  display: handwritten face and fallback
  technical: sans-serif face and fallback
line_split:
  semantic: exact connector convention
  expressive: permitted native accent strokes
native_doodle_inventory: stable object groups and their scientific roles
style_source_firewall: sources, transferable traits, prohibited transfers
rollback_signals: concrete evidence that the style layer harms reading
```

Do not use the same roughness level everywhere. Reserve the strongest hand-drawn cue for the title, focal mechanism, or one narrative transition; otherwise the slide loses hierarchy.

## Publication-oriented visual grammar

Useful starting tokens, to be adapted rather than treated as a fixed template:

- canvas: warm off-white near `#F7F3E8`;
- ink: charcoal near `#20252B`;
- cool stream: muted blue near `#4F7CAC` with a pale fill;
- warm stream: amber or coral near `#E09F3E` / `#D96855` with a pale fill;
- semantic connector: 1.4–2.0 pt dark ink, one arrowhead, no echo stroke;
- module outline: 1.3–1.8 pt; a secondary offset outline is optional only on one or two focal objects;
- highlighter: pale native shape behind text, low saturation, no glow;
- display typography: short phrases only; technical typography carries all exact labels.

Maintain a stable macro alignment system even if icons and annotations are asymmetric. Let shape choice, doodles, highlighter marks, and typography produce the hand-drawn character; do not intentionally damage equal spacing or connector clearance.

## PowerPoint-native mapping

First inspect backend capabilities. Use only object families the selected backend can create and re-audit.

- When PowerPoint COM is connected but Office interop enum metadata is unavailable, use the capability-reported built-in stable AutoShape subset or its numeric `shape_type_id`; do not treat missing enum reflection as proof that native shapes are unavailable.
- Build paper, panels, labels, highlighters, and glyphs with native shapes and text boxes.
- Build each doodle from a small named group of native shapes and lines. Keep its semantic role obvious without relying on decoration.
- Prefer attached connectors for scientific relations when the backend and route geometry support them. For a deliberately segmented orthogonal lane, exact named coordinate lines are acceptable when their endpoints, arrowhead, clearance, and semantic ownership are audited after every node movement. Expressive accents never carry a scientific relation.
- A limited double-ink effect may use one no-fill duplicate outline behind a focal shape, offset slightly and with lower opacity. Name it as an expressive accent, keep it out of labels and connectors, and remove it if the renderer makes it look like a shadow.
- Use small rotations only for non-semantic highlighters or callout tags. Do not rotate dense labels, evidence panels, or nodes that serve as connector anchors.
- If freeform or polyline creation is unavailable, do not import a broad SVG/PNG to fake rough outlines. Prefer a cleaner native result and disclose the constrained style mapping.
- Do not use Comic Sans everywhere, global paper textures, drop shadows, gradients, 3-D objects, clipart packs, decorative speech bubbles, or unrelated cartoon characters.

## LLM and agent overviews

LLM diagrams have recurring semantic confusions that a loose sketch can amplify. Keep these distinctions visible:

- context tokens are not model weights;
- retrieval evidence enters the assembled context, not the answer directly;
- memory is external state unless the source explicitly defines another mechanism;
- an LLM may propose a tool call, while an agent/controller executes it and returns an observation;
- tool observations re-enter the context or agent state; they do not update model weights during ordinary inference;
- training, fine-tuning, and inference occupy separate regions or receive an explicit boundary;
- attention, feed-forward processing, and next-token prediction should not be mislabeled as factual verification;
- citations or grounding indicators require an actual evidence path in the diagram.

For a compact LLM-agent overview, a strong default composition is:

```text
question + sources -> context builder -> focal LLM core -> answer
                                      -> tool request -> controller -> tool
                         updated context <- observation <-+
```

Use one clear return lane for the observation loop. Avoid an orbit of crossing arrows around the model merely because circular sketches look hand-drawn.

## Adversarial rejection gate

Reject or revise the direction when any of these is true:

- the same diagram with a handwritten font removed would reveal no other intentional hand-drawn language;
- roughness is applied to semantic connectors or makes endpoints uncertain;
- random rotations and uneven gaps replace an actual composition system;
- doodles compete with the scientific focal point or imply unsupported mechanisms;
- a paper texture, SVG, or screenshot contains reconstructable labels, borders, arrows, tables, or regular plots;
- handwriting is smaller or less legible than the clean technical fallback at publication width;
- the output resembles a classroom poster when the declared slot is a paper method overview;
- the output resembles a standard dashboard with a few squiggles added afterward;
- style references have silently transferred source topology or scientific content.

Positive hand-drawn fidelity does not require maximum roughness. A restrained result passes when the expressive cues are unmistakable at whole-slide scale, the exact layer survives grayscale and publication-size review, and removing the expressive layer would materially change the intended visual voice.

## Reviewer evidence and verdict

Use a fresh renderer export at thumbnail, whole-slide, and readable scales, plus a title-hidden grayscale view.

Answer separately:

1. Can every required scientific relation be reconstructed without relying on color or decorative strokes?
2. Is the requested hand-drawn family visibly present beyond the display font?
3. Are the expressive cues concentrated enough to create hierarchy rather than noise?
4. Does the result still read as a publication figure at the declared final size?
5. Which expressive objects would be removed first if legibility regressed?

When a correction is needed, change the smallest responsible layer: reduce a highlighter, remove an echo outline, replace one handwritten technical label, simplify a doodle, or reroute one exact connector. Do not globally randomize, recolor, or redraw already approved geometry.
