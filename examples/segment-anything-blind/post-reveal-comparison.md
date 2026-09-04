# Segment Anything post-reveal comparison

Reveal date: 2026-09-04

Blind artifact status: frozen before reveal; not retroactively edited

## Sources

- Paper: [*Segment Anything*, ICCV 2023](https://openaccess.thecvf.com/content/ICCV2023/html/Kirillov_Segment_Anything_ICCV_2023_paper.html)
- Official target: [`assets/model_diagram.png`](https://github.com/facebookresearch/segment-anything/blob/main/assets/model_diagram.png)
- Official repository and license: [facebookresearch/segment-anything](https://github.com/facebookresearch/segment-anything) · [Apache License 2.0](https://github.com/facebookresearch/segment-anything/blob/main/LICENSE)
- Frozen blind artifact: [`segment-anything-blind-overview.drawio`](../../assets/examples/segment-anything-blind-overview.drawio) · [PNG render](../../assets/examples/segment-anything-blind-overview.png)

The official image was first viewed only after the source contract, editable draw.io file, PNG render, and pre-reveal audit receipt had been frozen.

## What matched independently

- Both designs use a left-to-right image-encoding path, a separate prompt path, a prompt-conditioned mask decoder, and multiple masks with quality scores.
- Both distinguish an image-sized representation from prompt information and converge them before mask prediction.
- Both communicate ambiguity by showing multiple valid output masks rather than one deterministic answer.

These matches come from the method's scientific topology, not from layout copying.

## What the official overview does better

- It explains the product-level idea in one glance: a real image and user prompt produce several plausible masks.
- Real scissors photographs make the input/output relationship immediately concrete.
- Its ultra-wide single chain uses very few labels and is efficient for a paper reader who has not yet learned the implementation.
- The mask prompt is shown as a convolutional feature added to the image embedding, while points, box, and text enter the prompt encoder; this is a compact visual distinction.

## What the blind design does better

- It exposes the two-way token/image exchange rather than presenting the decoder as a black box.
- It separates sparse and dense prompts, mask tokens and the IoU token, upscaled image features and mask-token hypernetworks.
- It makes predicted IoU visibly originate from a quality-token path rather than from the mask pixels.
- It marks low-resolution-logit feedback as optional and uses a separate dashed lane, preventing it from being read as a mandatory one-pass cycle.
- Every visible label, glyph, container, and connector remains independently editable in draw.io.

## Source-scope difference

The blind contract used the public implementation as its detailed authority, so it includes point, box, and mask inputs and omits a text-prompt branch. The official paper overview includes text as a conceptual prompt mode. This is not evidence that either drawing copied or missed the other; it shows why a figure must state whether it describes the paper's conceptual interface or the released implementation.

## Verdict

The official figure is stronger as a compact, evidence-led introduction. The blind design is stronger as an editable architecture explanation and as a teaching artifact. The blind run succeeded because it recovered the central topology and ambiguity-aware output without target leakage, not because its composition resembles the official PNG.

The reusable lesson is to freeze source scope and communication purpose before drawing, then compare claim, abstraction, evidence, hierarchy, legibility, and editability after reveal. No case-specific palette, aspect ratio, or node count should become a universal rule.
