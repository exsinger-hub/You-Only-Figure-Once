# Segment Anything blind-design specification

Status: selected direction and blind artifact frozen before target-figure reveal on 2026-09-04

## Canvas and publication slot

- Logical design canvas: 960 × 540 units, scaled to an 816 × 459 native draw.io content hull.
- Export: 1600 px-wide PNG (1600 × 903 after content cropping).
- Review width: 980 px in GitHub README; 520 px grayscale thumbnail.
- Outer margin: 34 pt; title band: 56 pt; footer: 30 pt.
- Usable content hull: x=34..926, y=74..498.

## Direction decision

Two low-fidelity structures were considered without seeing the target figure:

1. **Converge-and-fan** — image and prompt lanes converge on one central decoder and fan out to candidate masks.
2. **Interaction loop** — the image occupies the center while prompts and iterative masks orbit it in a clockwise loop.

Selected: **Converge-and-fan**. It makes the paper's efficiency claim visible at thumbnail scale: the expensive image lane is executed once, while lightweight prompts can be changed repeatedly. The interaction-loop direction was rejected because it overstates feedback as mandatory and weakens the one-pass topology.

## Figure archetype and Style DNA

- Archetype: asymmetric dual-lane convergence with output fan.
- Reading path: upper-left image lane + lower-left prompt lane → central two-way decoder → right-side outputs.
- First focal point: the dark central `Two-way Transformer` aperture.
- Second focal point: the separation between `ENCODE ONCE` and `PROMPT MANY`.
- Third focal point: three candidate masks and their quality estimates.
- Shape language: flat native shapes, restrained corner radius, token pills, one grid-based image embedding, three mask silhouettes.
- Line language: orthogonal or direct attached routes; teal for image flow, amber for prompt flow, coral for outputs, dashed gray only for optional feedback.
- Whitespace: central focal protection and a dedicated bottom feedback lane; no decorative empty band.
- Forbidden motifs: gradients, shadows, glass effects, dashboard cards, clipart, full-panel rasterization, decorative background texture.

## Palette and typography

- canvas: `#F7F8F5`
- ink: `#172431`
- muted: `#60717D`
- image stream: `#147D89` / light `#DCEFF0`
- prompt stream: `#D8891C` / light `#FFF0D4`
- decoder focal: `#173F4F`
- output stream: `#D65F52` / light `#FBE5E1`
- optional feedback: `#8A98A3`
- font: Aptos / Arial fallback.
- title: 25 pt semibold; section labels: 10.5 pt bold; node labels: 13–16 pt; annotations: 9.5–11 pt.

## Region geometry

| region | bounds (x, y, w, h) | role |
| --- | --- | --- |
| title | 34, 20, 892, 46 | claim and provenance |
| image_lane | 34, 92, 500, 168 | image encoded once |
| prompt_lane | 34, 294, 500, 158 | reusable prompt path |
| decoder | 552, 132, 178, 272 | dominant two-way fusion |
| outputs | 758, 92, 168, 360 | ambiguity-aware outputs |
| feedback | 270, 468, 488, 28 | optional low-resolution-logit loop |

## Object map

Every required node and edge in `source-contract.md` maps to a named native draw.io cell or attached edge. No raster object is present.

- `n_image` → `sam_node_image`
- `n_image_encoder` → `sam_node_image_encoder`
- `n_image_embedding` → `sam_image_embedding_container`
- `n_point`, `n_box`, `n_mask_prompt` → `sam_node_point_prompt`, `sam_node_box_prompt`, `sam_node_mask_prompt`
- `n_prompt_encoder` → `sam_prompt_encoder_container`
- `n_sparse`, `n_dense` → `sam_node_sparse`, `sam_node_dense`
- `n_mask_tokens`, `n_iou_token` → `sam_node_mask_tokens`, `sam_node_iou_token`
- `n_two_way` → `sam_two_way_container`
- `n_upscale`, `n_hypernet` → `sam_node_upscale`, `sam_node_hypernet`
- `n_masks` → `sam_candidate_1_container..sam_candidate_3_container`
- `n_quality` → `sam_candidate_1_q..sam_candidate_3_q`
- `n_feedback` → `sam_edge_feedback`

## Construction order

1. Canvas, title, and lane labels.
2. Image schematic, encoder, and embedding grid.
3. Point/box/mask prompt glyphs, prompt encoder, sparse/dense outputs.
4. Decoder aperture, two-way arrows, mask/IoU token pills.
5. Upscale/hypernetwork split and three output masks with quality bars.
6. Optional feedback route and legend.
7. Structure audit, 1600 px render, 980 px review, 520 px grayscale review.

## Acceptance conditions before reveal

- 17/17 required nodes and 18/18 required relations are visible.
- The image encoder receives no prompt edge.
- The prompt encoder outputs sparse and dense representations distinctly.
- Token → image and image → token exchange are both visible.
- Three candidate masks have matched predicted-IoU labels without implying ground truth.
- Optional feedback is visually subordinate and explicitly dashed.
- No clipping, text overflow, unrelated connector crossing, or raster picture.
- The central decoder remains the first focus at 520 px grayscale.
