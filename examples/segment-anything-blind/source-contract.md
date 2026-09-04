# Segment Anything blind-design source contract

Status: blind artifact frozen before target-figure reveal on 2026-09-04

## Source authority and exclusion receipt

- Mode: blinded independent design / Figure Gym.
- Stable paper id: arXiv:2304.02643.
- Publication: *Segment Anything*, ICCV 2023.
- Allowed before design freeze:
  - the arXiv abstract;
  - official implementation files `sam.py`, `image_encoder.py`, `prompt_encoder.py`, `mask_decoder.py`, and `transformer.py`;
  - no paper figure, caption, alt text, slide, blog diagram, or derivative architecture illustration.
- Sealed before design freeze:
  - every figure and figure caption in the paper;
  - `facebookresearch/segment-anything/assets/model_diagram.png`;
  - README passages that display or describe that diagram's layout.
- Target slot: GitHub README case-study hero, reviewed at 980 px width.
- Deliverables: editable draw.io source, 1600 px PNG, source contract, design spec, post-reveal comparison.

## Figure Claim

SAM encodes an image once, converts point/box/mask prompts into sparse or dense embeddings, lets prompt tokens and image features exchange information in a lightweight two-way Transformer, and returns ambiguity-aware masks together with predicted mask quality.

## Paper Figure Signature

- Primary topology: image → ViT image encoder → image embedding → two-way mask decoder → upscaled image embedding → masks.
- Prompt topology: point/box → sparse embeddings; mask prompt → dense embedding; both condition the mask decoder.
- Decoder state: IoU token and mask tokens join prompt tokens; token-to-image and image-to-token cross-attention update both representations.
- Outputs: one or multiple masks and corresponding predicted IoU values; low-resolution logits can be reused as a mask prompt in a later interaction.
- Novelty focus: decoupled expensive image encoding from lightweight prompt-conditioned decoding.
- Reading order: upper image lane and lower prompt lane converge at the central decoder, then fan out to candidate masks.
- Evidence scope: architecture-only schematic; no empirical image or benchmark claim.

## Required-node ledger

| id | exact label | type | source | priority |
| --- | --- | --- | --- | --- |
| n_image | Image | input | `sam.py:90-92` | primary |
| n_image_encoder | ViT image encoder | encoder | `image_encoder.py:14-108` | primary |
| n_image_embedding | image embedding | dense feature | `sam.py:90-92` | primary |
| n_point | point prompt | prompt | `sam.py:66-70` | secondary |
| n_box | box prompt | prompt | `sam.py:71-72` | secondary |
| n_mask_prompt | mask prompt | prompt | `sam.py:73-74` | secondary |
| n_prompt_encoder | prompt encoder | encoder | `prompt_encoder.py:13-153` | primary |
| n_sparse | sparse prompt embeddings | token feature | `prompt_encoder.py:123-145` | primary |
| n_dense | dense prompt embedding | dense feature | `prompt_encoder.py:123-153` | primary |
| n_mask_tokens | mask tokens | decoder state | `mask_decoder.py:44-46` | secondary |
| n_iou_token | IoU token | decoder state | `mask_decoder.py:44-46` | secondary |
| n_two_way | two-way Transformer | decoder | `transformer.py:13-96` | focal |
| n_upscale | upscaled image embedding | dense feature | `mask_decoder.py:123-131` | secondary |
| n_hypernet | mask hypernetwork MLPs | predictor | `mask_decoder.py:54-58,126-131` | secondary |
| n_masks | candidate masks | output | `sam.py:75-85` | primary |
| n_quality | predicted IoU | output quality | `sam.py:84-85` | primary |
| n_feedback | low-resolution logits | iterative state | `sam.py:86-88` | tertiary |

## Required-edge ledger

| id | source → target | relation |
| --- | --- | --- |
| e_image_encoder | n_image → n_image_encoder | encode once |
| e_encoder_embedding | n_image_encoder → n_image_embedding | produces dense image features |
| e_embedding_decoder | n_image_embedding → n_two_way | image keys/values |
| e_point_prompt | n_point → n_prompt_encoder | sparse prompt input |
| e_box_prompt | n_box → n_prompt_encoder | sparse prompt input |
| e_mask_prompt | n_mask_prompt → n_prompt_encoder | dense prompt input |
| e_prompt_sparse | n_prompt_encoder → n_sparse | produces sparse tokens |
| e_prompt_dense | n_prompt_encoder → n_dense | produces dense conditioning |
| e_sparse_decoder | n_sparse → n_two_way | prompt queries |
| e_dense_decoder | n_dense → n_two_way | adds to image features |
| e_mask_tokens_decoder | n_mask_tokens → n_two_way | output queries |
| e_iou_token_decoder | n_iou_token → n_two_way | quality query |
| e_two_way_exchange | n_two_way ↔ n_image_embedding | token/image exchange |
| e_decoder_upscale | n_two_way → n_upscale | updates and upscales image representation |
| e_decoder_hypernet | n_two_way → n_hypernet | emits mask-token weights |
| e_mask_product | n_upscale + n_hypernet → n_masks | dynamic mask prediction |
| e_quality_head | n_two_way → n_quality | IoU-token MLP prediction |
| e_feedback | n_masks → n_feedback → n_mask_prompt | optional next interaction |

## Negative-path checks

- A point or box must not be drawn as a dense image-sized prompt.
- A mask prompt must not be drawn as a sparse token-only input.
- Predicted IoU is a quality estimate, not a ground-truth metric or supervision edge.
- The prompt encoder must not feed the image encoder.
- Feedback logits are optional iterative input, not a mandatory cycle for one-pass inference.

## Equation/operand ledger

- Mask prediction visibly requires both an upscaled image embedding and mask-token-derived hypernetwork weights.
- Predicted IoU visibly originates from the IoU token output, independent of the mask raster itself.
- The two-way block must show both token → image and image → token exchange; a single one-way arrow is insufficient.

## Evidence ledger

- All image and mask glyphs in the independent design are editable vector schematics.
- They claim only input/output roles and ambiguity, not qualitative superiority or dataset evidence.
- No raster evidence is permitted in the blind artifact.

## Freeze receipt

- Frozen editable artifact: `assets/examples/segment-anything-blind-overview.drawio`.
- Frozen review render: `assets/examples/segment-anything-blind-overview.png`.
- Backend deviation before reveal: the live PowerPoint COM session was being switched by another active drawing task, so the clean design was rebuilt with native draw.io cells. No target figure, caption, README diagram, or derivative overview was viewed during that change.
- Final pre-reveal audit: 102 editable objects, 80 vertices, 22 edges, 0 pictures, 0 hard failures, and 0 warnings.
