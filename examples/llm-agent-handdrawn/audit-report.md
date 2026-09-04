# LLM agent hand-drawn overview audit report

Status: PASS

## Delivered evidence

- Editable source: [`assets/examples/llm-agent-handdrawn-overview.pptx`](../../assets/examples/llm-agent-handdrawn-overview.pptx)
- Fresh renderer export: [`assets/examples/llm-agent-handdrawn-overview.png`](../../assets/examples/llm-agent-handdrawn-overview.png)
- Renderer: Microsoft PowerPoint 16.0 through Windows COM
- Slide: 960 × 540 pt; review export: 1600 × 900 px
- Object graph: 99 native PowerPoint objects, 0 picture objects
- Saved-file reopen check: 1 slide, 99 shapes, 0 pictures, expected title present
- Deterministic whole-slide audit: 0 hard failures, 0 warnings
- Thumbnail and grayscale review: PASS at 520 px width

## Contract reconstruction

All 11 required nodes are visibly recoverable:

| contract node | visible implementation |
| --- | --- |
| Question | native stick figure, question bubble, and `Question` label |
| Retrieved sources | three offset native document shapes with source IDs 1–3 |
| Context builder | blue panel with retrieve, tokenize, and assemble stages |
| Context tokens | eight native token cells |
| LLM core | dominant dark panel with self-attention, feed-forward, and next-token probabilities |
| proposal | `answer or tool?` decision diamond |
| Agent controller | explicit validate-and-execute node |
| External tool | search/API/calculator node |
| Observation | tool-result node |
| Grounded answer | separate answer panel |
| Evidence links | S1–S3 chips retained inside the answer panel |

All 12 required relations are reconstructable. Ten use exact named line segments with visible arrowheads. The two provenance relations use matched source IDs (`1–3` → `S1–S3`) and containment inside the grounded-answer panel, so they do not add a visually misleading bypass connector.

## Adversarial semantic checks

| attempted misreading | visible counter-evidence | verdict |
| --- | --- | --- |
| Retrieved evidence becomes prose without entering context | both question and source paths terminate at the context builder | PASS |
| The model itself executes external tools | proposal routes to a separate agent controller before the tool | PASS |
| Tool use is mandatory | action route and return lane are dashed and labelled optional | PASS |
| Observation updates model weights | return label explicitly states that model weights stay frozen | PASS |
| Token prediction verifies factual truth | LLM panel states `prediction ≠ verification`; answer retains source IDs | PASS |

## Style-fidelity checks

- Hand-drawn character comes from native doodles, restrained rotations, one focal echo outline, highlighter tags, and a display/technical typography split.
- Every semantic route remains geometrically exact. No random endpoint jitter, raster paper texture, copied icon, or reference screenshot is present.
- The central LLM remains the first focal point in color and grayscale. Context is second; grounded answer is third; the optional tool loop remains subordinate.
- The figure reads as a technical paper illustration rather than a generic software architecture diagram or a decorative sketchnote.

## Correction history

1. The first context-to-token route crossed the `Context tokens` label. It was replaced with three exact segments routed around the label, then rerendered and re-audited.
2. The intentionally offset evidence sheets were initially named like a numbered alignment series. They were renamed by semantic layer (`back_coral`, `mid_blue`, `front`) so the audit rule for accidentally misaligned repeated data objects would not misclassify a declared doodle stack.
3. The LLM echo outline was initially treated as an unrelated semantic node. It was renamed `hd_llm_background_accent`, matching its actual decorative role, while its geometry remained independently visible for review.
4. Live construction exposed a backend defect: friendly AutoShape names failed when Office interop enum metadata was unavailable even though COM drawing worked. The bridge now has a stable native-shape subset fallback, and the COM smoke test uses `rectangle` by name to cover the regression.

After the corrections, a fresh 1600 × 900 PowerPoint export and whole-slide audit returned 99 native objects, 0 pictures, 0 hard failures, and 0 warnings.

## Scope boundary

This is a generic explanatory LLM-agent overview. It makes no benchmark, model-vendor, performance, safety, or empirical claim. A journal-specific submission should still rerun the final-width typography gate for that journal's exact figure slot.
