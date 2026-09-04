# LLM agent hand-drawn overview design specification

Status: frozen before PowerPoint construction

## Canvas and message

- Backend: Windows Microsoft PowerPoint through native COM.
- Canvas: 960 × 540 pt, 16:9 README demonstration slide.
- Review exports: 1600 × 900 full render and 520 × 293 grayscale/thumbnail judgment.
- Usable bounds: x=34..926, y=82..500.
- Reading path: question and evidence → context → focal LLM → answer; an isolated lower return lane handles optional tool use.
- Figure Claim: the sentence frozen in `source-contract.md`.

## Composition

- Archetype: asymmetric left-to-right spine with one lower feedback lane.
- First focal point: central dark `LLM CORE` block.
- Second focal point: the blue `CONTEXT` assembly block.
- Third focal point: the warm `GROUNDED ANSWER` block.
- Visual mass: context and LLM occupy about half the usable hull; input and output remain compact; the lower tool loop is subordinate.
- Whitespace functions:
  - title-to-content gap: hierarchy;
  - space around the LLM core: focal protection;
  - bottom band: connector lane for tool observations;
  - right margin: output isolation.

## Hand-drawn Style DNA

```text
style_family: restrained_technical_handdrawn
imperfection_budget:
  expressive_regions: title highlighter, five section tags, input/source doodles,
    focal LLM echo outline, answer underline
  exact_regions: all semantic connectors, technical labels, token sequence,
    controller/tool nodes, citation chips
  maximum_rotation_or_offset: 1.2 degrees or 3 pt, expressive objects only
typography_split:
  display: Segoe Print; fallback Aptos
  technical: Aptos; fallback Arial
line_split:
  semantic: exact attached connectors, charcoal, 1.7 pt, triangle arrowheads
  expressive: native short lines or one low-opacity echo outline
native_doodle_inventory:
  doodle_questioner: head, body, arms, question bubble
  doodle_sources: three offset native document shapes
  doodle_answer: native response strokes and evidence chips
style_source_firewall: source-contract.md
rollback_signals: connector ambiguity, technical handwriting, double outline read as shadow,
  doodle competition, or loss of readability at 520 px
```

## Palette and typography

- canvas: `#F7F3E8`
- ink: `#20252B`
- muted ink: `#687078`
- context blue: `#4F7CAC`; pale blue: `#E5EEF6`
- answer coral: `#D96855`; pale coral: `#F8E3DC`
- tool amber: `#E09F3E`; pale amber: `#F8EDD2`
- focal LLM: `#26343D`; focal text: `#FFFDF7`
- title: Segoe Print 27 pt bold; section tags: Segoe Print 12–13 pt bold.
- technical labels: Aptos 10–15 pt; annotations: Aptos 9–10 pt.

## Region geometry

| region | bounds (x, y, w, h) | role |
| --- | --- | --- |
| title | 38, 20, 884, 54 | title, subtitle, highlighter |
| inputs | 38, 116, 126, 266 | question and retrieved sources |
| context | 196, 126, 170, 238 | context assembly and token output |
| llm | 405, 105, 226, 280 | dominant model core and proposal split |
| answer | 780, 137, 144, 224 | grounded response and evidence links |
| tool_loop | 414, 424, 360, 72 | controller, external tool, observation |
| return_lane | 180, 492, 594, 18 | observation feedback into context |

## Stable object plan

- Canvas and title: `hd_canvas`, `hd_title_highlight`, `hd_title`, `hd_subtitle`.
- Inputs: `hd_question_head`, `hd_question_body`, `hd_question_bubble`, `hd_sources_back_coral`, `hd_sources_mid_blue`, `hd_sources_front`, and clean labels.
- Context: `hd_context_panel`, three native assembly rows, `hd_token_1..8`, and `hd_context_tokens_label`.
- LLM: `hd_llm_background_accent`, `hd_llm_panel`, `hd_llm_title`, three transformer rows, and `hd_proposal`.
- Answer: `hd_answer_panel`, native response lines, `hd_citation_1..3`, and a short expressive underline.
- Tool loop: `hd_controller`, `hd_tool`, `hd_observation` with exact labels.
- Every required relation uses a named connector or named coordinate line; expressive strokes never carry scientific meaning.

## PowerPoint-native implementation

- Use only text boxes, AutoShapes, lines, attached connectors, groups, and z-order operations.
- Use no SVG, PNG, screenshot, paper texture, chart, or broad visual overlay.
- Keep the main spine on one horizontal alignment band.
- Use one no-fill echo outline behind `hd_llm_panel`, offset by 3 pt; remove it if PowerPoint renders it as a drop shadow.
- Rotate only highlighter/tag shapes, never connector anchors or technical labels.
- Doodle icons remain composed of individually editable native objects and may be grouped only after local review.

## Construction and review order

1. New isolated PowerPoint presentation, focus policy `preserve`.
2. Canvas and title region.
3. Inputs and context region; export and local review.
4. LLM and answer region; export and local review.
5. Tool loop and feedback lane; export and local review.
6. Whole-slide deterministic audit, 1600 px render, 520 px thumbnail and grayscale review.
7. At least one object-level correction followed by a fresh render and repeat audit.

## Pass conditions

- 11/11 required nodes and 12/12 required relations visible.
- No connector crosses text or unrelated objects; the feedback lane does not look mandatory.
- Zero raster objects, clipping, text overflow, or unrelated overlaps.
- Hand-drawn cues remain visible at whole-slide scale but technical labels stay crisp at 520 px.
- The slide reads as a publication-oriented illustration, not a dashboard, marketing poster, or classroom worksheet.

Final evidence and the correction history are recorded in [audit-report.md](audit-report.md).
