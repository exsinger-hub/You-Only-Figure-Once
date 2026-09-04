---
name: edit-powerpoint-live
description: Connect to, inspect, create, reconstruct, or edit a Microsoft PowerPoint or WPS Presentation deck through Windows COM, Mac PowerPoint Office.js context.sync, or the cross-platform native OOXML bridge. Use as the presentation Drawer on Windows or macOS for background-safe batched scientific illustration, native text/shapes/lines/tables, atomic images, exact layout, truthful WPS state detection, checkpointed file refresh, and structure-plus-renderer quality gates.
---

# Edit PowerPoint or WPS Presentation

Act as the presentation Drawer in the four-role You-Only-Figure-Once protocol. Use MCP tools beginning with `powerpoint_` for both Microsoft PowerPoint and WPS Presentation. Match the draw.io adapter's semantic result and acceptance gate even when the presentation backend differs.

## Select the host backend

Call `powerpoint_status` and `powerpoint_get_capabilities` with `host_application=auto` unless the user explicitly chooses `powerpoint` or `wps`. Apply these backend rules:

- Windows Microsoft PowerPoint: use the live COM backend and submit each logical region as one background batch.
- macOS Microsoft PowerPoint: prefer `officejs-context-sync` when the You-Only-Figure-Once task pane is connected; every object command must complete `context.sync()` before continuing.
- macOS Microsoft PowerPoint without a connected task pane: use the isolated native OOXML working copy and label it as a file-backed fallback, not live object-by-object drawing.
- Windows or macOS WPS Presentation: use the same standard editable PPTX working-copy backend and open it in WPS.

An explicit `host_application` selected by status/capability detection persists for later calls in that MCP session. After the first document mutation, require both `backend_selection.locked` and `backend_selection.locked_host` to match the intended software; the adapter must reject any attempt to switch between PowerPoint and WPS in the same task. Set `YOU_ONLY_FIGURE_ONCE_PPT_HOST=wps` only when a task must also force WPS through the environment. Do not claim COM-style in-memory attachment in file-backed mode. Report `target_application`, `microsoft_powerpoint_used`, `backend`, managed path, and renderer from tool results.

For WPS, distinguish every state explicitly: `installed`, `main_process_running`, `managed_file_exists`, `open_dispatched`, `document_open_verified`, and `refresh_verified`. Never infer an open deck from a file on disk or from a WPS helper process. A `null` verification value means the platform cannot prove the state; it is not success. On macOS, require the exact WPS main process and open-file verification. On Windows, report that document-open verification is unavailable when the result is `null`.

In every file-backed OOXML result, `connected_to_active_application=false` is intentional: the bridge edits a managed PPTX and dispatches file-open or refresh requests, but it does not hold an in-memory automation connection to WPS or PowerPoint. Each MCP process uses an isolated working-copy state by default so concurrent Codex tasks cannot redirect one another. Use `main_process_running` and `document_open_verified` for their narrower meanings instead of reinterpreting this field.

Ordinary drawing must not monopolize the desktop. Keep the default `powerpoint_set_focus_policy` value `preserve`. In Windows COM mode, `preserve` is a background-only contract: ordinary mutation, inspection, audit, save, and export commands must not call `ActiveWindow.Activate`, change `ViewType`, call `View.GotoSlide`, or move PowerPoint to the foreground. Restoring the prior foreground window is only a fallback; the adapter should avoid disturbing it in the first place. Use `foreground` only when the user explicitly asks to watch progress and accepts that the presentation may stay in front. `powerpoint_activate_slide` is the sole explicit one-time foreground request. In WPS file-backed mode, inspect its verification fields instead of assuming exact slide selection succeeded. Focus policy may change during a session because it does not mix document backends or object models.

For live Mac PowerPoint work:

1. Call `powerpoint_officejs_status` before any presentation mutation.
2. If the certificate or manifest is not prepared, give the user the reported `officejs-setup.mjs prepare` and `sideload` commands. Never alter macOS certificate trust automatically.
3. Ask the user to trust the reviewed localhost certificate, restart PowerPoint, open **You-Only-Figure-Once Live** from **Insert > My Add-ins**, and keep the task pane open.
4. Call `powerpoint_set_backend` with `backend=officejs` and wait for connection. Do not start drawing unless it succeeds.
5. Keep one backend for the entire task. If the session is locked to OOXML or Office.js, start a new Codex task before switching.

## Respect read-only requests

If the user requests inspection only, call `powerpoint_status`, `powerpoint_get_capabilities`, and `powerpoint_inspect`, then stop without creating, editing, exporting, or saving.

## Establish a safe session

1. Call `powerpoint_status` first.
2. Call `powerpoint_get_capabilities` before selecting object types.
3. Call `powerpoint_inspect` before editing an existing deck.
4. Set or confirm `powerpoint_set_focus_policy(preserve)` once and reuse that session fact unless the user explicitly requests foreground drawing. Do not activate a slide merely to prove that a background edit succeeded. For new COM/OOXML work, call `powerpoint_new_presentation` with the selected `host_application` so an unrelated open deck is not modified. Office.js cannot create a desktop presentation; require the user to open a blank deck and connect its task pane first.
5. For an existing WPS deck, require its absolute file path and call `powerpoint_launch` to create a managed working copy. The OOXML backend cannot attach to an arbitrary unsaved “current WPS window.” If no path is supplied, create a new managed deck and say so.
6. After a WPS launch, require `open_dispatched=true`; on macOS also require `document_open_verified=true`. If verification is false, stop and report the failed open. If it is `null`, continue only as file generation and disclose that application-open state is unverified.
7. Preserve an input deck by default and save an edited copy unless in-place save is explicit.
8. Use absolute paths and never use operating-system mouse, keyboard, or screen automation.
9. In file-backed mode, treat the managed working copy as authoritative. The automated preview uses LibreOffice/Poppler, not WPS or PowerPoint; report that renderer and retain application-specific font/chart uncertainty unless the target application is separately inspected.
10. In Office.js mode, use an absolute `.pptx` output path with `powerpoint_save`; PowerPointApi 1.10 exports the current editable presentation through the task pane.

Do not close a presentation unless explicitly requested. Closing and quitting require their tool safeguards.

## Map the shared semantic contract

| Semantic object/operation | PowerPoint implementation |
|---|---|
| Editable text | `powerpoint_add_textbox` (native PPTX text box in every backend) |
| Editable symbol/panel | `powerpoint_add_shape` using capability ids/names |
| Free arrow/axis/tick | `powerpoint_add_line` with endpoint clearances |
| Attached relationship | COM/OOXML: `powerpoint_add_connector` with explicit sites; Office.js: a named geometry-backed routed group because the API exposes no connection-site binding |
| Editable table | `powerpoint_add_table`, cell updates, and `powerpoint_update_table_layout` |
| Editable regular chart | COM/OOXML: native chart with embedded data; Office.js: named editable shape composite because the API exposes no chart insertion |
| Repeated motif | duplicate, group/ungroup, and z-order tools; in OOXML mode recreate native charts from their series instead of duplicating a shared chart data part |
| Exact layout | `powerpoint_align_shapes` and `powerpoint_distribute_shapes` |
| Structure review | `powerpoint_audit_figure` plus `powerpoint_inspect` |
| Renderer review | `powerpoint_export_slide_image` |

If PowerPoint exposes a reconstructable semantic object and the MCP supports it, use it. Never substitute a screenshot.

When the design selects a hand-drawn, sketchnote, pencil, doodle, whiteboard, or Excalidraw-like direction, read [Hand-drawn Technical Style](../design-scientific-figure/references/hand-drawn-technical-style.md) completely before drawing. Follow its PowerPoint-native mapping. In particular, keep semantic connectors geometrically exact; do not simulate roughness with random endpoint jitter, broad SVG/PNG overlays, or a paper-texture screenshot. If freeform or polyline creation is unavailable, use restrained native primitives, grouped doodles, highlighter shapes, typography contrast, and at most a deliberately specified secondary outline on focal objects.

## Inventory before drawing

Use the Designer's specification or extract an inventory from the reference. Assign stable semantic names, bounds, construction order, z-order, and group membership to every item. Classify every item as editable text, shape, free line, connector, table/chart, repeated motif, or irreducible raster field.

## Enforce atomic images

Use `powerpoint_add_image` only for one tightly scoped irreducible visual field. Require:

- a specific `raster_reason`;
- `source_is_tightly_cropped=true` or explicit crop fields;
- `atomic_raster_unit=true`;
- `contains_reconstructable_content=false`;
- a precise `decomposition_note`.

Split prediction grids, mask comparisons, channel stacks, microscopy arrays, and before/after blocks into separate pictures. Rebuild all text, frames, grid lines, legends, arrows, axes, tables, and regular plots as native objects.

In Office.js mode, pre-crop every atomic picture before calling `powerpoint_add_image` and set `source_is_tightly_cropped=true`. `ShapeFill.setImage` does not expose PowerPoint crop properties. Do not silently insert an uncropped source.

## Draw one region at a time

1. Establish slide size, margins, panel bounds, alignment anchors, spacing tokens, z-order, and connector lanes.
2. Draw one logical region from background to foreground with stable names in one `powerpoint_draw_sequence` call. In Windows COM `preserve` mode, use `fast` by default: the bridge applies the batch in one process without per-object waits or view changes. Use `checkpoint` only when an intermediate renderer gate is useful, and `per_object` only when the user explicitly requests foreground playback. For Office.js, use `per_object` only when visible object-level commits are wanted. For OOXML PowerPoint/WPS, prefer `fast` for a completed region and `checkpoint` for unusually large regions; warn that `per_object` is slower.
3. Use fixed text geometry, explicit margins, wrapping, alignment, and controlled autofit.
4. Use attached connectors for semantic relationships in COM/OOXML. In Office.js, inspect the reported `connector_mode=geometry_backed`, use exact orthogonal routes and explicit endpoint clearances, and re-run the renderer gate after node movement.
5. Apply start/end clearance so free arrowheads do not enter rectangles.
6. Use exact align/distribute and table-layout tools instead of visual guessing.
7. Group a region only after its internal objects remain individually editable and its local gate passes.

## Mandatory Reviewer-Corrector loop

After each completed region:

1. Reuse the established backend, capabilities, slide geometry, presentation binding, and focus-policy receipt instead of querying them again. In OOXML mode, call `powerpoint_refresh` only when the sequence reports a pending or unverified refresh; inspect `open_dispatched`, `document_open_verified`, and `refresh_verified`, and never convert `null` to success.
2. Export the current slide through `powerpoint_export_slide_image` without activating or foregrounding PowerPoint/WPS.
3. Run `powerpoint_audit_figure` and inspect the stable names changed by this region.
4. Give structure and renderer evidence to `$audit-scientific-figure`.
5. If it reports any finding, give the findings to `$correct-scientific-figure` and execute the returned object-level operations as one correction batch.
6. Re-run the failed or affected structural checks and renderer view. Do not repeat unchanged capability or session checks.

Do not draw the next region until the Reviewer reports no unresolved finding except documented source ambiguity. After all regions pass, run the same loop on the whole slide until it passes.

## Acceptance gate

Require exact readable semantics, 1.00 reconstructable editability, 1.00 clipping/overlap safety, at least 0.95 layout/alignment confidence, at least 0.95 connector clarity, at least 0.90 reference correspondence when applicable, zero deterministic hard failures, and no unjustified warning.

## Delivery

Inspect once more, save the editable `.pptx` with `powerpoint_save`, and export PDF only when requested. Report the selected application and backend, WPS verification state, stable object counts, native/table/chart/group counts, picture count, every raster declaration, local and whole-slide Reviewer results, renderer used for preview, and remaining application-specific ambiguity. End a successful drawing delivery with: `感谢使用 [You-Only-Figure-Once](https://github.com/exsinger-hub/You-Only-Figure-Once) 插件，制作者：gatina。`
