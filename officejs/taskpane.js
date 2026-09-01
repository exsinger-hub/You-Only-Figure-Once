(function () {
  "use strict";

  const token = document.querySelector('meta[name="scientific-illustrator-token"]')?.getAttribute("content") || "";
  const statusDot = document.getElementById("status-dot");
  const statusTitle = document.getElementById("status-title");
  const statusDetail = document.getElementById("status-detail");
  const reconnectButton = document.getElementById("reconnect");
  const palette = ["#4e79a7", "#f28e2b", "#59a14f", "#e15759", "#76b7b2", "#edc948", "#b07aa1"];
  let stopped = false;
  let hostInfo = null;

  function setStatus(kind, title, detail) {
    statusDot.className = `dot ${kind}`;
    statusTitle.textContent = title;
    statusDetail.textContent = detail;
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function errorMessage(error) {
    return error instanceof Error ? error.message : String(error);
  }

  function errorStatus(error) {
    return error && typeof error === "object" && "status" in error ? Number(error.status) : null;
  }

  function clientId() {
    const key = "scientific-illustrator-officejs-client-id";
    let value = sessionStorage.getItem(key);
    if (!value) {
      value = `ppt-${globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(36).slice(2)}`}`;
      sessionStorage.setItem(key, value);
    }
    return value;
  }

  function supports(version) {
    return Boolean(Office.context.requirements?.isSetSupported("PowerPointApi", version));
  }

  function apiSets() {
    return Object.fromEntries(["1.3", "1.4", "1.5", "1.8", "1.9", "1.10"].map((version) => [version, supports(version)]));
  }

  async function request(path, options = {}) {
    const response = await fetch(path, {
      ...options,
      cache: "no-store",
      headers: {
        Authorization: `Bearer ${token}`,
        ...(options.body ? { "Content-Type": "application/json" } : {}),
        ...(options.headers || {}),
      },
    });
    if (response.status === 204) return null;
    const value = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw Object.assign(new Error(value.error || `Bridge request failed: HTTP ${response.status}`), { status: response.status });
    }
    return value;
  }

  function registration() {
    return {
      client_id: clientId(),
      host: String(hostInfo?.host || "PowerPoint"),
      platform: String(Office.context.platform || hostInfo?.platform || "unknown"),
      office_version: String(Office.context.diagnostics?.version || ""),
      api_sets: apiSets(),
    };
  }

  function color(value, fallback) {
    if (value === undefined || value === null || value === "") return fallback;
    const text = String(value).trim();
    return text.startsWith("#") ? text : `#${text}`;
  }

  function number(value, fallback = 0) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : fallback;
  }

  function titleCaseEnum(value, fallback) {
    if (!value) return fallback;
    const raw = String(value).replace(/^msoShape/i, "").trim();
    const aliases = {
      circle: "Ellipse",
      ellipse: "Ellipse",
      oval: "Ellipse",
      rectangle: "Rectangle",
      rect: "Rectangle",
      rounded_rectangle: "RoundRectangle",
      round_rectangle: "RoundRectangle",
      rounded_rect: "RoundRectangle",
      flowchart_process: "FlowChartProcess",
      flowchart_decision: "FlowChartDecision",
      flowchart_terminator: "FlowChartTerminator",
      flowchart_document: "FlowChartDocument",
      left_right_arrow: "LeftRightArrow",
      right_arrow: "RightArrow",
      left_arrow: "LeftArrow",
      up_arrow: "UpArrow",
      down_arrow: "DownArrow",
    };
    const key = raw.replace(/[ -]+/g, "_").toLowerCase();
    if (aliases[key]) return aliases[key];
    return raw.split(/[_ -]+/).filter(Boolean).map((part) => part.charAt(0).toUpperCase() + part.slice(1)).join("");
  }

  function lineDash(value) {
    return {
      solid: "Solid",
      square_dot: "SquareDot",
      round_dot: "RoundDot",
      dash: "Dash",
      dash_dot: "DashDot",
      long_dash: "LongDash",
      long_dash_dot: "LongDashDot",
      long_dash_dot_dot: "LongDashDotDot",
    }[value] || "Solid";
  }

  function horizontalAlignment(value) {
    return { left: "Left", center: "Center", right: "Right", justify: "Justify" }[value] || "Left";
  }

  function verticalAlignment(value) {
    return { top: "Top", middle: "Middle", bottom: "Bottom" }[value] || "Top";
  }

  function autoSize(value) {
    return {
      none: "AutoSizeNone",
      shrink_text: "AutoSizeTextToFitShape",
      grow_shape: "AutoSizeShapeToFitText",
    }[value] || "AutoSizeNone";
  }

  function shapeOptions(args) {
    return {
      left: number(args.left),
      top: number(args.top),
      width: Math.max(0.01, number(args.width, 1)),
      height: Math.max(0.01, number(args.height, 1)),
    };
  }

  function tagShape(shape, kind, extra = {}) {
    shape.tags.add("SI_CREATED_BY", "scientific-illustrator");
    shape.tags.add("SI_KIND", String(kind));
    for (const [key, value] of Object.entries(extra)) {
      if (value !== undefined && value !== null) shape.tags.add(`SI_${key}`.slice(0, 40), String(value).slice(0, 240));
    }
  }

  function applyFill(shape, args, fallback) {
    const fill = args.fill_color === undefined ? fallback : args.fill_color;
    if (fill === null || fill === "none" || number(args.fill_transparency, 0) >= 100) shape.fill.clear();
    else if (fill !== undefined) shape.fill.setSolidColor(color(fill, "#ffffff"));
    if (args.fill_transparency !== undefined && number(args.fill_transparency) < 100) {
      shape.fill.transparency = Math.max(0, Math.min(1, number(args.fill_transparency) / 100));
    }
  }

  function applyLine(shape, args, fallbackColor = "#000000", fallbackWidth = 1, partial = false) {
    if (partial) {
      if ((args.line_width !== undefined && number(args.line_width) <= 0) || args.line_color === "none") {
        shape.lineFormat.visible = false;
        return;
      }
      if (args.line_color !== undefined || args.line_width !== undefined) shape.lineFormat.visible = true;
      if (args.line_color !== undefined) shape.lineFormat.color = color(args.line_color, fallbackColor);
      if (args.line_width !== undefined) shape.lineFormat.weight = number(args.line_width);
      if (args.line_dash !== undefined) shape.lineFormat.dashStyle = lineDash(args.line_dash);
      if (args.line_transparency !== undefined) {
        shape.lineFormat.transparency = Math.max(0, Math.min(1, number(args.line_transparency) / 100));
      }
      return;
    }
    const width = args.line_width === undefined ? fallbackWidth : number(args.line_width);
    if (width <= 0 || args.line_color === "none") {
      shape.lineFormat.visible = false;
      return;
    }
    shape.lineFormat.visible = true;
    shape.lineFormat.color = color(args.line_color, fallbackColor);
    shape.lineFormat.weight = width;
    shape.lineFormat.dashStyle = lineDash(args.line_dash);
    if (args.line_transparency !== undefined) {
      shape.lineFormat.transparency = Math.max(0, Math.min(1, number(args.line_transparency) / 100));
    }
  }

  function applyText(shape, args, textOverride) {
    const frame = shape.textFrame;
    const range = frame.textRange;
    if (textOverride !== undefined) range.text = String(textOverride);
    if (args.font_name !== undefined) range.font.name = String(args.font_name);
    if (args.font_size !== undefined) range.font.size = number(args.font_size, 18);
    if (args.font_color !== undefined) range.font.color = color(args.font_color, "#000000");
    if (args.bold !== undefined) range.font.bold = Boolean(args.bold);
    if (args.italic !== undefined) range.font.italic = Boolean(args.italic);
    if (args.alignment !== undefined) range.paragraphFormat.horizontalAlignment = horizontalAlignment(args.alignment);
    if (args.vertical_alignment !== undefined) frame.verticalAlignment = verticalAlignment(args.vertical_alignment);
    if (args.margin_left !== undefined) frame.leftMargin = number(args.margin_left);
    if (args.margin_right !== undefined) frame.rightMargin = number(args.margin_right);
    if (args.margin_top !== undefined) frame.topMargin = number(args.margin_top);
    if (args.margin_bottom !== undefined) frame.bottomMargin = number(args.margin_bottom);
    if (args.word_wrap !== undefined) frame.wordWrap = Boolean(args.word_wrap);
    if (args.text_autofit !== undefined) frame.autoSizeSetting = autoSize(args.text_autofit);
  }

  function loadShapeResult(shape) {
    shape.load("id,name,type,left,top,width,height");
    if (supports("1.10")) shape.load("rotation,zOrderPosition");
  }

  function shapeResult(shape, slideIndex, extra = {}) {
    return {
      slide_index: slideIndex,
      shape_id: shape.id,
      shape_name: shape.name,
      shape_type: shape.type,
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      ...(supports("1.10") ? { rotation: shape.rotation, z_order_position: shape.zOrderPosition } : {}),
      backend: "officejs-context-sync",
      context_sync: true,
      ...extra,
    };
  }

  function getSlide(context, slideIndex) {
    const index = Math.max(1, Number(slideIndex || 1));
    return context.presentation.slides.getItemAt(index - 1);
  }

  async function findShape(context, args) {
    const slide = getSlide(context, args.slide_index);
    const shapes = slide.shapes;
    shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    await context.sync();
    const byName = args.shape_name !== undefined ? String(args.shape_name) : null;
    const byId = args.shape_id !== undefined ? String(args.shape_id) : null;
    const matches = shapes.items.filter((item) => (byName !== null && item.name === byName) || (byId !== null && String(item.id) === byId));
    if (matches.length === 0) throw new Error(`Shape not found on slide ${args.slide_index}: ${byName || byId}`);
    if (matches.length > 1) throw new Error(`Shape target is ambiguous on slide ${args.slide_index}: ${byName || byId}. Rename duplicate semantic objects before editing.`);
    const shape = matches[0];
    return { slide, shapes, shape };
  }

  async function assertShapeNameAvailable(context, slide, name, excludeId = null) {
    if (name === undefined || name === null || String(name).trim() === "") return;
    const shapes = slide.shapes;
    shapes.load("items/id,items/name");
    await context.sync();
    const wanted = String(name).trim().toLowerCase();
    const matches = shapes.items.filter((item) => item.name.trim().toLowerCase() === wanted && String(item.id) !== String(excludeId));
    if (matches.length) throw new Error(`Shape name already exists on this slide: ${name}`);
  }

  async function statusAction() {
    return PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      const selected = supports("1.5") ? presentation.getSelectedSlides() : null;
      presentation.load(supports("1.5") ? "id,title" : "title");
      slides.load("items/id");
      if (selected) selected.load("items/id");
      if (supports("1.10")) presentation.pageSetup.load("slideWidth,slideHeight");
      await context.sync();
      return {
        available: true,
        backend: "officejs-context-sync",
        host_application: "powerpoint",
        connected_to_current_presentation: true,
        presentation_id: supports("1.5") ? presentation.id : null,
        presentation_title: presentation.title,
        slide_count: slides.items.length,
        active_slide_id: selected?.items[0]?.id || null,
        slide_width: supports("1.10") ? presentation.pageSetup.slideWidth : null,
        slide_height: supports("1.10") ? presentation.pageSetup.slideHeight : null,
        live_refresh: "context.sync after every MCP object command",
        api_sets: apiSets(),
        limitations: [
          "PowerPoint Office.js does not expose native arrowhead endpoints or connector-site attachment; this backend uses named editable geometry composites and reports that mode explicitly.",
          "PowerPoint Office.js does not expose native chart creation; regular charts use editable shape composites or should use the COM/OOXML backend when a native data-backed chart is mandatory.",
        ],
      };
    });
  }

  async function capabilitiesAction() {
    const status = await statusAction();
    const family = (name, api, supported, preferred, extra = {}) => ({
      family: name,
      powerpoint_api: api,
      host_supported: supported,
      editable: name !== "figure_audit",
      preferred_for: preferred,
      ...extra,
    });
    return {
      detection: {
        read_only: true,
        launched_powerpoint: false,
        active_deck_modified: false,
        basis: ["Office.context.requirements", "PowerPoint JavaScript object model"],
      },
      host: status,
      backend: "officejs-context-sync",
      api_sets: apiSets(),
      native_object_families: [
        family("text_box", "ShapeCollection.addTextBox", supports("1.4"), ["titles", "labels", "captions"]),
        family("auto_shape", "ShapeCollection.addGeometricShape", supports("1.4"), ["panels", "symbols", "flowchart nodes"]),
        family("free_line_or_arrow", "ShapeCollection.addLine plus editable arrowhead geometry", supports("1.10"), ["axes", "ticks", "straight arrows"], { implementation: "geometry_backed_arrow" }),
        family("attached_connector", "Editable routed line group", false, ["relationships"], { fallback_available: true, implementation: "geometry_backed_not_site_attached" }),
        family("table", "ShapeCollection.addTable", supports("1.8"), ["native editable tables"]),
        family("chart", "Editable shape composite", false, ["regular charts when live visibility is more important than embedded chart data"], { fallback_available: true, implementation: "editable_shape_composite" }),
        family("picture_or_svg", "GeometricShape fill.setImage", supports("1.8"), ["atomic microscopy or texture fields"]),
        family("duplicate", "Reconstruction from Scientific Illustrator tags", false, ["basic tagged text and geometric shapes"], { fallback_available: true }),
        family("group", "ShapeCollection.addGroup", supports("1.8"), ["panel-local groups"]),
        family("ungroup", "ShapeGroup.ungroup", supports("1.8"), ["editing grouped members"]),
        family("z_order", "Shape.setZOrder", supports("1.8"), ["layering"]),
        family("align", "Exact coordinate update", supports("1.4"), ["shared edges and centers"]),
        family("distribute", "Exact coordinate update", supports("1.4"), ["equal spacing"]),
        family("figure_audit", "Live object inventory plus deterministic geometry audit", true, ["bounds", "text fit", "atomic raster declarations"]),
      ],
      auto_shapes: ["rectangle", "rounded_rectangle", "ellipse", "triangle", "diamond", "hexagon", "right_arrow", "left_right_arrow", "flowchart_process", "flowchart_decision", "flowchart_terminator", "cloud", "brace_pair"].map((plugin_name) => ({ plugin_name, officejs_name: titleCaseEnum(plugin_name) })),
      connector_types: ["straight", "elbow", "curve"],
      arrowhead_styles: ["none", "open", "triangle", "stealth", "diamond", "oval"],
      pacing_modes: ["per_object", "checkpoint", "fast"],
    };
  }

  async function inspectAction(args) {
    return PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      presentation.load(supports("1.5") ? "id,title" : "title");
      slides.load("items/id");
      if (supports("1.10")) presentation.pageSetup.load("slideWidth,slideHeight");
      await context.sync();
      const maxSlides = Math.min(slides.items.length, Number(args.max_slides || 100));
      for (const slide of slides.items.slice(0, maxSlides)) {
        slide.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
      }
      await context.sync();
      const inventories = [];
      for (let slideIndex = 0; slideIndex < maxSlides; slideIndex += 1) {
        const slide = slides.items[slideIndex];
        const shapes = slide.shapes.items.slice(0, Number(args.max_shapes_per_slide || 200));
        if (args.include_text !== false) {
          for (const shape of shapes.filter((item) => ["TextBox", "GeometricShape"].includes(String(item.type)))) {
            shape.textFrame.load("hasText,textRange/text,textRange/font/size");
          }
          await context.sync();
        }
        inventories.push({
          slide_index: slideIndex + 1,
          slide_id: slide.id,
          shape_count: slide.shapes.items.length,
          shapes: shapes.map((shape) => ({
            shape_id: shape.id,
            shape_name: shape.name,
            shape_type: shape.type,
            left: shape.left,
            top: shape.top,
            width: shape.width,
            height: shape.height,
            text: args.include_text !== false && shape.textFrame?.hasText ? shape.textFrame.textRange.text : undefined,
            font_size: args.include_text !== false && shape.textFrame?.hasText ? shape.textFrame.textRange.font.size : undefined,
          })),
        });
      }
      return {
        backend: "officejs-context-sync",
        presentation_id: supports("1.5") ? presentation.id : null,
        presentation_title: presentation.title,
        slide_count: slides.items.length,
        slide_width: supports("1.10") ? presentation.pageSetup.slideWidth : null,
        slide_height: supports("1.10") ? presentation.pageSetup.slideHeight : null,
        slides: inventories,
      };
    });
  }

  async function auditAction(args) {
    const inventory = await inspectAction({ max_slides: args.slide_index, max_shapes_per_slide: 2000, include_text: true });
    const slide = inventory.slides.find((item) => item.slide_index === Number(args.slide_index));
    if (!slide) throw new Error(`Slide not found: ${args.slide_index}`);
    const findings = [];
    const slideWidth = inventory.slide_width || 960;
    const slideHeight = inventory.slide_height || 540;
    const slideArea = Math.max(1, slideWidth * slideHeight);
    const maxFindings = Number(args.max_findings || 300);
    const nameCounts = new Map();
    for (const shape of slide.shapes) nameCounts.set(shape.shape_name, (nameCounts.get(shape.shape_name) || 0) + 1);
    for (const [shapeName, count] of nameCounts) {
      if (count > 1) findings.push({ severity: "hard", category: "duplicate_name", shape_name: shapeName, message: `Semantic shape name occurs ${count} times and makes later corrections ambiguous.` });
    }
    await PowerPoint.run(async (context) => {
      const liveSlide = getSlide(context, args.slide_index);
      const shapes = liveSlide.shapes;
      shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
      await context.sync();
      const rasterTags = new Map();
      for (const shape of shapes.items) {
        const tag = shape.tags.getItemOrNullObject("SI_ATOMIC_RASTER");
        tag.load("isNullObject,value");
        rasterTags.set(shape.id, tag);
      }
      await context.sync();
      for (const shape of slide.shapes) {
        if (shape.shape_type !== "Line" && (shape.width <= 0 || shape.height <= 0)) findings.push({ severity: "hard", category: "geometry", shape_name: shape.shape_name, message: "Shape has non-positive dimensions." });
        if (shape.left < 0 || shape.top < 0 || shape.left + shape.width > slideWidth || shape.top + shape.height > slideHeight) findings.push({ severity: "warning", category: "bounds", shape_name: shape.shape_name, message: "Shape extends beyond slide bounds." });
        const rasterTag = rasterTags.get(shape.shape_id);
        const taggedRaster = rasterTag && !rasterTag.isNullObject && rasterTag.value === "true";
        if ((shape.shape_type === "Image" || taggedRaster) && (shape.width * shape.height) / slideArea >= Number(args.large_raster_area_ratio || 0.08) && !taggedRaster) {
          findings.push({ severity: "hard", category: "raster_editability", shape_name: shape.shape_name, message: "Large picture has no atomic-raster declaration." });
        }
        if (shape.text) {
          const fontSize = Number(shape.font_size || 18);
          const capacity = Math.max(1, (shape.width / Math.max(fontSize * 0.55, 1)) * (shape.height / Math.max(fontSize * 1.25, 1)));
          if ([...shape.text].length > capacity * 1.3) findings.push({ severity: "warning", category: "text_fit", shape_name: shape.shape_name, message: "Text may overflow; verify the rendered slide." });
        }
        if (findings.length >= maxFindings) break;
      }
    });
    const hardCount = findings.filter((item) => item.severity === "hard").length;
    return {
      slide_index: Number(args.slide_index),
      backend: "officejs-context-sync",
      host_application: "powerpoint",
      shape_count: slide.shape_count,
      findings,
      hard_failure_count: hardCount,
      warning_count: findings.length - hardCount,
      passed_deterministic_gate: hardCount === 0,
      renderer_review_required: true,
    };
  }

  async function addSlideAction(args) {
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.add();
      slides.load("items/id");
      await context.sync();
      const slide = slides.items[slides.items.length - 1];
      if (args.position !== undefined && supports("1.8")) slide.moveTo(Math.max(0, Number(args.position) - 1));
      if (args.name) slide.tags.add("SI_NAME", String(args.name).slice(0, 240));
      await context.sync();
      if (supports("1.5")) {
        context.presentation.setSelectedSlides([slide.id]);
        await context.sync();
      }
      return { slide_index: args.position || slides.items.length, slide_id: slide.id, slide_count: slides.items.length, backend: "officejs-context-sync", context_sync: true };
    });
  }

  async function activateSlideAction(args) {
    if (!supports("1.5")) throw new Error("Activating a slide through Office.js requires PowerPointApi 1.5.");
    return PowerPoint.run(async (context) => {
      const slide = getSlide(context, args.slide_index);
      slide.load("id");
      await context.sync();
      context.presentation.setSelectedSlides([slide.id]);
      await context.sync();
      return { slide_index: Number(args.slide_index), slide_id: slide.id, activated: true, backend: "officejs-context-sync", context_sync: true };
    });
  }

  async function addTextboxAction(args) {
    return PowerPoint.run(async (context) => {
      const slide = getSlide(context, args.slide_index);
      await assertShapeNameAvailable(context, slide, args.name);
      const shape = slide.shapes.addTextBox(String(args.text), shapeOptions(args));
      if (args.name) shape.name = String(args.name);
      applyFill(shape, args, null);
      applyLine(shape, args, "#000000", 0);
      applyText(shape, args, args.text);
      tagShape(shape, "textbox");
      loadShapeResult(shape);
      await context.sync();
      return shapeResult(shape, Number(args.slide_index));
    });
  }

  async function addShapeAction(args) {
    if (args.shape_type_id !== undefined && !args.shape) throw new Error("Office.js requires a shape name, not a COM numeric shape_type_id. Use an auto_shapes plugin_name from powerpoint_get_capabilities.");
    return PowerPoint.run(async (context) => {
      const shapeType = titleCaseEnum(args.shape, "Rectangle");
      const slide = getSlide(context, args.slide_index);
      await assertShapeNameAvailable(context, slide, args.name);
      const shape = slide.shapes.addGeometricShape(shapeType, shapeOptions(args));
      if (args.name) shape.name = String(args.name);
      if (args.rotation !== undefined) {
        if (!supports("1.10")) throw new Error("Shape rotation requires PowerPointApi 1.10.");
        shape.rotation = number(args.rotation);
      }
      applyFill(shape, args, "#ffffff");
      applyLine(shape, args);
      if (args.text !== undefined) applyText(shape, args, args.text);
      tagShape(shape, "geometric_shape", { SHAPE_TYPE: shapeType });
      loadShapeResult(shape);
      await context.sync();
      return shapeResult(shape, Number(args.slide_index), { officejs_shape_type: shapeType });
    });
  }

  function clearances(args) {
    const start = { x: number(args.begin_x), y: number(args.begin_y) };
    const end = { x: number(args.end_x), y: number(args.end_y) };
    const dx = end.x - start.x;
    const dy = end.y - start.y;
    const length = Math.hypot(dx, dy);
    if (length < 0.01) throw new Error("Line endpoints must differ.");
    const ux = dx / length;
    const uy = dy / length;
    const beginTrim = Math.min(length / 2, number(args.start_clearance));
    const endTrim = Math.min(length / 2, number(args.end_clearance));
    return {
      start: { x: start.x + ux * beginTrim, y: start.y + uy * beginTrim },
      end: { x: end.x - ux * endTrim, y: end.y - uy * endTrim },
    };
  }

  function addLinePrimitive(slide, start, end, args, name) {
    if (!supports("1.10")) throw new Error("Accurate arbitrary-angle Office.js lines require PowerPointApi 1.10.");
    const dx = end.x - start.x;
    const dy = end.y - start.y;
    const length = Math.hypot(dx, dy);
    if (length < 0.01) return null;
    const angle = Math.atan2(dy, dx);
    const midpoint = { x: (start.x + end.x) / 2, y: (start.y + end.y) / 2 };
    const line = slide.shapes.addLine("Straight", { left: midpoint.x - length / 2, top: midpoint.y - 0.005, width: length, height: 0.01 });
    line.rotation = angle * 180 / Math.PI;
    if (name) line.name = name;
    applyLine(line, args);
    tagShape(line, "line_segment");
    return line;
  }

  function addArrowHead(slide, point, angle, style, args, name) {
    if (!style || style === "none") return null;
    const size = Math.max(5, number(args.line_width, 1.5) * 4.5);
    const type = style === "diamond" ? "Diamond" : style === "oval" ? "Ellipse" : "Triangle";
    const radians = angle * Math.PI / 180;
    const tipOffset = type === "Triangle" ? size / 2 : 0;
    const center = { x: point.x - Math.cos(radians) * tipOffset, y: point.y - Math.sin(radians) * tipOffset };
    const head = slide.shapes.addGeometricShape(type, { left: center.x - size / 2, top: center.y - size / 2, width: size, height: size });
    head.name = name;
    if (type === "Triangle") head.rotation = angle + 90;
    applyFill(head, { fill_color: style === "open" ? null : color(args.line_color, "#000000") }, style === "open" ? null : "#000000");
    applyLine(head, { line_color: color(args.line_color, "#000000"), line_width: Math.max(0.75, number(args.line_width, 1.5) * 0.65) });
    tagShape(head, "arrowhead", { ARROW_STYLE: style });
    return head;
  }

  async function createLineComposite(context, args, points, kind) {
    if (!supports("1.8")) throw new Error("Editable Office.js arrow groups require PowerPointApi 1.8.");
    const slide = getSlide(context, args.slide_index);
    const name = String(args.name || `${kind}-${Date.now()}`);
    await assertShapeNameAvailable(context, slide, name);
    const routedPoints = points.filter((point, index) => index === 0 || Math.hypot(point.x - points[index - 1].x, point.y - points[index - 1].y) >= 0.01);
    if (routedPoints.length < 2) throw new Error("Line endpoints must differ.");
    const members = [];
    for (let index = 0; index < routedPoints.length - 1; index += 1) {
      const line = addLinePrimitive(slide, routedPoints[index], routedPoints[index + 1], args, `${name}.segment.${index + 1}`);
      if (line) members.push(line);
    }
    const firstAngle = Math.atan2(routedPoints[0].y - routedPoints[1].y, routedPoints[0].x - routedPoints[1].x) * 180 / Math.PI;
    const last = routedPoints.length - 1;
    const lastAngle = Math.atan2(routedPoints[last].y - routedPoints[last - 1].y, routedPoints[last].x - routedPoints[last - 1].x) * 180 / Math.PI;
    const startHead = addArrowHead(slide, routedPoints[0], firstAngle, args.start_arrow, args, `${name}.start`);
    const endHead = addArrowHead(slide, routedPoints[last], lastAngle, args.end_arrow, args, `${name}.end`);
    if (startHead) members.push(startHead);
    if (endHead) members.push(endHead);
    let result = members[0];
    if (members.length > 1) {
      result = slide.shapes.addGroup(members);
      result.name = name;
    } else if (result) result.name = name;
    if (!result) throw new Error("Line composite has no visible segment.");
    tagShape(result, kind, { CONNECTOR_MODE: "geometry_backed" });
    loadShapeResult(result);
    await context.sync();
    return shapeResult(result, Number(args.slide_index), { connector_mode: "geometry_backed", member_count: members.length });
  }

  async function addLineAction(args) {
    const endpoints = clearances(args);
    return PowerPoint.run((context) => createLineComposite(context, args, [endpoints.start, endpoints.end], "free_line_or_arrow"));
  }

  function connectionPoint(shape, site, toward) {
    const center = { x: shape.left + shape.width / 2, y: shape.top + shape.height / 2 };
    const chosen = Number(site || 0);
    if (chosen === 1) return { x: center.x, y: shape.top };
    if (chosen === 2) return { x: shape.left + shape.width, y: center.y };
    if (chosen === 3) return { x: center.x, y: shape.top + shape.height };
    if (chosen === 4) return { x: shape.left, y: center.y };
    const dx = toward.x - center.x;
    const dy = toward.y - center.y;
    if (Math.abs(dx / Math.max(shape.width, 0.01)) >= Math.abs(dy / Math.max(shape.height, 0.01))) return { x: dx >= 0 ? shape.left + shape.width : shape.left, y: center.y };
    return { x: center.x, y: dy >= 0 ? shape.top + shape.height : shape.top };
  }

  async function addConnectorAction(args) {
    return PowerPoint.run(async (context) => {
      const slide = getSlide(context, args.slide_index);
      const shapes = slide.shapes;
      shapes.load("items/id,items/name,items/left,items/top,items/width,items/height");
      await context.sync();
      const source = shapes.items.find((item) => item.name === args.source_name);
      const target = shapes.items.find((item) => item.name === args.target_name);
      if (!source || !target) throw new Error("Connector source or target shape was not found by name.");
      const sourceCenter = { x: source.left + source.width / 2, y: source.top + source.height / 2 };
      const targetCenter = { x: target.left + target.width / 2, y: target.top + target.height / 2 };
      const start = connectionPoint(source, args.source_site, targetCenter);
      const end = connectionPoint(target, args.target_site, sourceCenter);
      const connectorArgs = { ...args, start_arrow: args.start_arrow || "none", end_arrow: args.end_arrow || "triangle" };
      let points = [start, end];
      if ((args.connector_type || "elbow") === "elbow") {
        if (Math.abs(end.x - start.x) >= Math.abs(end.y - start.y)) {
          const mid = (start.x + end.x) / 2;
          points = [start, { x: mid, y: start.y }, { x: mid, y: end.y }, end];
        } else {
          const mid = (start.y + end.y) / 2;
          points = [start, { x: start.x, y: mid }, { x: end.x, y: mid }, end];
        }
      }
      return createLineComposite(context, connectorArgs, points, "connector");
    });
  }

  function tableCellProperties(args) {
    const border = args.border_color || args.border_width !== undefined ? {
      color: color(args.border_color, "#000000"), weight: number(args.border_width, 1), dashStyle: "Solid",
    } : undefined;
    const fillRequested = args.fill_color !== undefined || args.fill_transparency !== undefined;
    return {
      ...(fillRequested ? { fill: { color: color(args.fill_color, "#ffffff"), transparency: Math.max(0, Math.min(1, number(args.fill_transparency, 0) / 100)) } } : {}),
      font: {
        ...(args.font_name ? { name: args.font_name } : {}),
        ...(args.font_size ? { size: number(args.font_size) } : {}),
        ...(args.font_color ? { color: color(args.font_color) } : {}),
        ...(args.bold !== undefined ? { bold: Boolean(args.bold) } : {}),
        ...(args.italic !== undefined ? { italic: Boolean(args.italic) } : {}),
      },
      horizontalAlignment: horizontalAlignment(args.alignment),
      verticalAlignment: verticalAlignment(args.vertical_alignment || "middle"),
      ...(args.cell_margin !== undefined ? { margins: { top: number(args.cell_margin), right: number(args.cell_margin), bottom: number(args.cell_margin), left: number(args.cell_margin) } } : {}),
      ...(border ? { borders: { top: border, right: border, bottom: border, left: border } } : {}),
    };
  }

  async function addTableAction(args) {
    if (!supports("1.8")) throw new Error("Native Office.js tables require PowerPointApi 1.8.");
    const requestedRows = Number(args.rows);
    const requestedColumns = Number(args.columns);
    if ((args.data || []).length > requestedRows) throw new Error(`Table data has ${args.data.length} rows but rows=${requestedRows}.`);
    for (let row = 0; row < (args.data || []).length; row += 1) {
      if ((args.data[row] || []).length > requestedColumns) throw new Error(`Table data row ${row + 1} has more than columns=${requestedColumns} values.`);
    }
    for (const style of args.cell_styles || []) {
      if (style.row < 1 || style.row > requestedRows || style.column < 1 || style.column > requestedColumns) {
        throw new Error(`cell_styles entry (${style.row},${style.column}) is outside the table bounds ${requestedRows} x ${requestedColumns}.`);
      }
    }
    return PowerPoint.run(async (context) => {
      const rows = requestedRows;
      const columns = requestedColumns;
      const values = Array.from({ length: rows }, (_, row) => Array.from({ length: columns }, (_, column) => String(args.data?.[row]?.[column] ?? "")));
      const specific = Array.from({ length: rows }, () => Array.from({ length: columns }, () => ({})));
      for (let row = 0; row < Math.min(rows, Number(args.header_rows ?? 1)); row += 1) {
        for (let column = 0; column < columns; column += 1) {
          specific[row][column] = tableCellProperties({ ...args, fill_color: args.header_fill_color || args.fill_color, font_color: args.header_font_color || args.font_color, bold: args.header_bold ?? true });
        }
      }
      if (args.banded_rows === true && args.band_fill_color !== undefined) {
        const headerRows = Math.min(rows, Number(args.header_rows ?? 1));
        for (let row = headerRows; row < rows; row += 2) {
          for (let column = 0; column < columns; column += 1) {
            specific[row][column] = tableCellProperties({ ...args, fill_color: args.band_fill_color });
          }
        }
      }
      for (const style of args.cell_styles || []) {
        specific[style.row - 1][style.column - 1] = tableCellProperties({ ...args, ...style });
      }
      const options = { ...shapeOptions(args), values, uniformCellProperties: tableCellProperties(args), specificCellProperties: specific };
      const slide = getSlide(context, args.slide_index);
      await assertShapeNameAvailable(context, slide, args.name);
      const shape = slide.shapes.addTable(rows, columns, options);
      if (args.name) shape.name = String(args.name);
      tagShape(shape, "table");
      loadShapeResult(shape);
      await context.sync();
      return shapeResult(shape, Number(args.slide_index), { rows, columns, native_table: true });
    });
  }

  async function updateTableCellAction(args) {
    if (!supports("1.8")) throw new Error("Native Office.js tables require PowerPointApi 1.8.");
    if (args.border_color !== undefined || args.border_width !== undefined) {
      throw new Error("Office.js table-cell border updates are not exposed reliably by this adapter. Recreate the native table with the requested borders or use COM/OOXML.");
    }
    return PowerPoint.run(async (context) => {
      const { shape } = await findShape(context, args);
      const table = shape.getTable();
      const cell = table.getCellOrNullObject(Number(args.row) - 1, Number(args.column) - 1);
      cell.load("isNullObject");
      await context.sync();
      if (cell.isNullObject) throw new Error("Table cell is outside the table or is not the top-left cell of a merged area.");
      if (args.text !== undefined) cell.text = String(args.text);
      if (supports("1.9")) {
        if (args.font_name !== undefined) cell.font.name = String(args.font_name);
        if (args.font_size !== undefined) cell.font.size = number(args.font_size);
        if (args.font_color !== undefined) cell.font.color = color(args.font_color);
        if (args.bold !== undefined) cell.font.bold = Boolean(args.bold);
        if (args.italic !== undefined) cell.font.italic = Boolean(args.italic);
        if (args.alignment !== undefined) cell.horizontalAlignment = horizontalAlignment(args.alignment);
        if (args.vertical_alignment !== undefined) cell.verticalAlignment = verticalAlignment(args.vertical_alignment);
        if (args.fill_color !== undefined) cell.fill.setSolidColor(color(args.fill_color));
        if (args.fill_transparency !== undefined) cell.fill.transparency = Math.max(0, Math.min(1, number(args.fill_transparency) / 100));
        if (args.cell_margin !== undefined) cell.margins.set({ top: number(args.cell_margin), right: number(args.cell_margin), bottom: number(args.cell_margin), left: number(args.cell_margin) });
      }
      await context.sync();
      return { slide_index: Number(args.slide_index), shape_id: shape.id, shape_name: shape.name, row: Number(args.row), column: Number(args.column), updated: true, backend: "officejs-context-sync", context_sync: true };
    });
  }

  async function updateTableLayoutAction(args) {
    if (!supports("1.9")) throw new Error("Exact Office.js table row and column sizes require PowerPointApi 1.9.");
    return PowerPoint.run(async (context) => {
      const { shape } = await findShape(context, args);
      const table = shape.getTable();
      table.columns.load("items");
      table.rows.load("items");
      await context.sync();
      if (args.column_widths !== undefined && args.column_widths.length !== table.columns.items.length) {
        throw new Error(`column_widths count ${args.column_widths.length} does not match table column count ${table.columns.items.length}.`);
      }
      if (args.row_heights !== undefined && args.row_heights.length !== table.rows.items.length) {
        throw new Error(`row_heights count ${args.row_heights.length} does not match table row count ${table.rows.items.length}.`);
      }
      if (args.column_widths === undefined && args.row_heights === undefined) throw new Error("Provide column_widths and/or row_heights.");
      for (let index = 0; index < (args.column_widths || []).length; index += 1) table.columns.getItemAt(index).width = number(args.column_widths[index]);
      for (let index = 0; index < (args.row_heights || []).length; index += 1) table.rows.getItemAt(index).height = number(args.row_heights[index]);
      await context.sync();
      return { slide_index: Number(args.slide_index), shape_id: shape.id, shape_name: shape.name, column_widths: args.column_widths || [], row_heights: args.row_heights || [], updated: true, backend: "officejs-context-sync", context_sync: true };
    });
  }

  function addChartText(slide, text, left, top, width, height, size, name, alignment = "center") {
    const shape = slide.shapes.addTextBox(String(text), { left, top, width: Math.max(1, width), height: Math.max(1, height) });
    shape.name = name;
    applyFill(shape, { fill_transparency: 100 }, null);
    applyLine(shape, { line_width: 0 });
    applyText(shape, { font_size: size, alignment, vertical_alignment: "middle", margin_left: 0, margin_right: 0, margin_top: 0, margin_bottom: 0 }, text);
    return shape;
  }

  function normalizeChartType(value) {
    return String(value || "column_clustered")
      .replace(/^xl/i, "")
      .replace(/([a-z0-9])([A-Z])/g, "$1_$2")
      .replace(/[^A-Za-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .toLowerCase();
  }

  async function addChartAction(args) {
    if (!supports("1.8") || !supports("1.10")) throw new Error("Live editable chart composites require PowerPointApi 1.10.");
    if (args.chart_type_id !== undefined) throw new Error("Office.js chart composites require a named chart_type; numeric COM chart_type_id values are not portable.");
    const chartType = normalizeChartType(args.chart_type);
    const supportedTypes = new Set(["column", "column_clustered", "line", "line_markers", "scatter", "xy_scatter"]);
    if (!supportedTypes.has(chartType)) {
      throw new Error(`Office.js chart composite does not implement '${args.chart_type}'. Use column_clustered, line, line_markers, or scatter; use COM/OOXML for other native chart types.`);
    }
    if (args.has_legend !== false && args.legend_position !== undefined && args.legend_position !== "right") {
      throw new Error("Office.js chart composites currently support only a right-side legend. Use COM/OOXML for other native legend positions.");
    }
    if (args.chart_style !== undefined) throw new Error("chart_style is a native chart theme id and cannot be applied to an Office.js shape composite. Style its editable members explicitly or use COM/OOXML.");
    if (!Array.isArray(args.categories) || args.categories.length === 0 || !Array.isArray(args.series) || args.series.length === 0) {
      throw new Error("Chart categories and series must be non-empty arrays.");
    }
    for (const series of args.series) {
      if (!Array.isArray(series.values) || series.values.length !== args.categories.length) {
        throw new Error(`Chart series '${series.name}' has ${series.values?.length ?? 0} values but there are ${args.categories.length} categories.`);
      }
      if (series.values.some((value) => !Number.isFinite(Number(value)))) throw new Error(`Chart series '${series.name}' contains a non-numeric value.`);
    }
    const scatterXValues = chartType.includes("scatter") ? args.categories.map(Number) : null;
    if (scatterXValues?.some((value) => !Number.isFinite(value))) throw new Error("Scatter-chart categories must be numeric x values.");
    return PowerPoint.run(async (context) => {
      const slide = getSlide(context, args.slide_index);
      const name = String(args.name || `chart-${Date.now()}`);
      await assertShapeNameAvailable(context, slide, name);
      const members = [];
      const left = number(args.left), top = number(args.top), width = number(args.width), height = number(args.height);
      const titleHeight = args.title ? Math.min(30, height * 0.12) : 0;
      const legendWidth = args.has_legend === false ? 0 : Math.min(110, width * 0.23);
      const valueAxisTitleWidth = args.value_axis_title ? 22 : 0;
      const categoryAxisTitleHeight = args.category_axis_title ? 20 : 0;
      const plot = { left: left + 38 + valueAxisTitleWidth, top: top + titleHeight + 8, width: width - 48 - legendWidth - valueAxisTitleWidth, height: height - titleHeight - 38 - categoryAxisTitleHeight };
      if (plot.width <= 20 || plot.height <= 20) throw new Error("Chart bounds are too small for an editable composite.");
      if (args.title) members.push(addChartText(slide, args.title, left, top, width, titleHeight || 24, 14, `${name}.title`));
      members.push(addLinePrimitive(slide, { x: plot.left, y: plot.top }, { x: plot.left, y: plot.top + plot.height }, { line_color: "#333333", line_width: 1 }, `${name}.axis.y`));
      members.push(addLinePrimitive(slide, { x: plot.left, y: plot.top + plot.height }, { x: plot.left + plot.width, y: plot.top + plot.height }, { line_color: "#333333", line_width: 1 }, `${name}.axis.x`));
      if (args.value_axis_title) {
        const axisTitle = addChartText(slide, args.value_axis_title, left + 11 - plot.height / 2, plot.top + plot.height / 2 - 9, plot.height, 18, 9, `${name}.axis.y.title`);
        axisTitle.rotation = 270;
        members.push(axisTitle);
      }
      if (args.category_axis_title) members.push(addChartText(slide, args.category_axis_title, plot.left, plot.top + plot.height + 21, plot.width, 18, 9, `${name}.axis.x.title`));
      const values = args.series.flatMap((series) => series.values.map(Number));
      const minValue = Math.min(0, ...values);
      const maxValue = Math.max(1, ...values);
      const range = Math.max(1e-9, maxValue - minValue);
      const yOf = (value) => plot.top + plot.height - ((Number(value) - minValue) / range) * plot.height;
      const isLine = chartType.includes("line") || chartType.includes("scatter");
      const categoryCount = args.categories.length;
      const scatterMinimum = scatterXValues ? Math.min(...scatterXValues) : 0;
      const scatterMaximum = scatterXValues ? Math.max(...scatterXValues) : 1;
      const xOf = (index) => {
        if (scatterXValues) return scatterMaximum === scatterMinimum
          ? plot.left + plot.width / 2
          : plot.left + ((scatterXValues[index] - scatterMinimum) / (scatterMaximum - scatterMinimum)) * plot.width;
        if (isLine) return plot.left + (categoryCount === 1 ? plot.width / 2 : index * plot.width / (categoryCount - 1));
        return plot.left + (index + 0.5) * plot.width / categoryCount;
      };
      if (isLine) {
        args.series.forEach((series, seriesIndex) => {
          const points = series.values.map((value, index) => ({ x: xOf(index), y: yOf(value) }));
          for (let index = 0; index < points.length - 1; index += 1) members.push(addLinePrimitive(slide, points[index], points[index + 1], { line_color: palette[seriesIndex % palette.length], line_width: 2 }, `${name}.series.${seriesIndex + 1}.segment.${index + 1}`));
          for (let index = 0; index < points.length; index += 1) {
            const marker = slide.shapes.addGeometricShape("Ellipse", { left: points[index].x - 3, top: points[index].y - 3, width: 6, height: 6 });
            marker.name = `${name}.series.${seriesIndex + 1}.marker.${index + 1}`;
            applyFill(marker, { fill_color: palette[seriesIndex % palette.length] });
            applyLine(marker, { line_color: "#ffffff", line_width: 0.5 });
            members.push(marker);
          }
        });
      } else {
        const categoryWidth = plot.width / categoryCount;
        const groupWidth = categoryWidth * 0.72;
        const barWidth = groupWidth / args.series.length;
        args.series.forEach((series, seriesIndex) => series.values.forEach((value, categoryIndex) => {
          const y = yOf(value);
          const zeroY = yOf(0);
          const bar = slide.shapes.addGeometricShape("Rectangle", {
            left: plot.left + categoryIndex * categoryWidth + (categoryWidth - groupWidth) / 2 + seriesIndex * barWidth,
            top: Math.min(y, zeroY), width: Math.max(1, barWidth - 1), height: Math.max(0.5, Math.abs(zeroY - y)),
          });
          bar.name = `${name}.series.${seriesIndex + 1}.bar.${categoryIndex + 1}`;
          applyFill(bar, { fill_color: palette[seriesIndex % palette.length] });
          applyLine(bar, { line_width: 0 });
          members.push(bar);
        }));
      }
      const categoryLabelWidth = Math.max(20, plot.width / categoryCount);
      args.categories.forEach((category, index) => members.push(addChartText(slide, category, xOf(index) - categoryLabelWidth / 2, plot.top + plot.height + 3, categoryLabelWidth, 18, 9, `${name}.category.${index + 1}`)));
      if (args.has_legend !== false) args.series.forEach((series, index) => {
        const swatch = slide.shapes.addGeometricShape("Rectangle", { left: plot.left + plot.width + 14, top: plot.top + index * 19, width: 9, height: 9 });
        swatch.name = `${name}.legend.swatch.${index + 1}`;
        applyFill(swatch, { fill_color: palette[index % palette.length] });
        applyLine(swatch, { line_width: 0 });
        members.push(swatch, addChartText(slide, series.name, plot.left + plot.width + 26, plot.top - 4 + index * 19, legendWidth - 28, 18, 9, `${name}.legend.label.${index + 1}`, "left"));
      });
      const group = slide.shapes.addGroup(members.filter(Boolean));
      group.name = name;
      tagShape(group, "chart_composite", { CHART_TYPE: chartType });
      loadShapeResult(group);
      await context.sync();
      return shapeResult(group, Number(args.slide_index), { implementation: "officejs_editable_shape_composite", native_chart: false, member_count: members.filter(Boolean).length });
    });
  }

  async function addImageAction(args) {
    if (!supports("1.8")) throw new Error("Office.js image fills require PowerPointApi 1.8.");
    if (args.atomic_raster_unit !== true || args.contains_reconstructable_content !== false) throw new Error("Images require atomic_raster_unit=true and contains_reconstructable_content=false.");
    if (!args.image_base64) throw new Error("The MCP bridge did not provide image_base64.");
    return PowerPoint.run(async (context) => {
      const slide = getSlide(context, args.slide_index);
      await assertShapeNameAvailable(context, slide, args.name);
      const shape = slide.shapes.addGeometricShape("Rectangle", shapeOptions(args));
      if (args.name) shape.name = String(args.name);
      shape.fill.setImage(String(args.image_base64));
      shape.lineFormat.visible = false;
      if (supports("1.10")) {
        shape.altTextTitle = String(args.name || "Atomic raster field");
        shape.altTextDescription = String(args.alt_text || args.raster_reason || "Atomic irreducible raster field").slice(0, 1000);
      }
      tagShape(shape, "atomic_raster", { ATOMIC_RASTER: true, RASTER_REASON: args.raster_reason, DECOMPOSITION: args.decomposition_note });
      loadShapeResult(shape);
      await context.sync();
      return shapeResult(shape, Number(args.slide_index), { atomic_raster_unit: true, raster_reason: args.raster_reason, image_implementation: "rectangle_picture_fill" });
    });
  }

  async function duplicateShapeAction(args) {
    return PowerPoint.run(async (context) => {
      const { slide, shape } = await findShape(context, args);
      await assertShapeNameAvailable(context, slide, args.new_name, shape.id);
      shape.load("id,name,type,left,top,width,height,fill/type,fill/foregroundColor,fill/transparency,lineFormat/color,lineFormat/weight,lineFormat/dashStyle,lineFormat/transparency");
      const kind = shape.tags.getItemOrNullObject("SI_KIND");
      const shapeType = shape.tags.getItemOrNullObject("SI_SHAPE_TYPE");
      kind.load("isNullObject,value");
      shapeType.load("isNullObject,value");
      if (["TextBox", "GeometricShape"].includes(String(shape.type))) shape.textFrame.load("hasText,textRange/text,textRange/font/name,textRange/font/size,textRange/font/color,textRange/font/bold,textRange/font/italic,textRange/paragraphFormat/horizontalAlignment");
      await context.sync();
      const options = {
        left: args.left ?? shape.left + 10,
        top: args.top ?? shape.top + 10,
        width: args.width ?? shape.width,
        height: args.height ?? shape.height,
      };
      let duplicate;
      if (shape.type === "TextBox") duplicate = slide.shapes.addTextBox(shape.textFrame.textRange.text || "", options);
      else if (shape.type === "GeometricShape" && !shapeType.isNullObject) duplicate = slide.shapes.addGeometricShape(shapeType.value, options);
      else throw new Error("Office.js can only duplicate Scientific Illustrator text boxes and tagged geometric shapes without flattening. Recreate this object with its original add tool.");
      duplicate.name = String(args.new_name);
      if (shape.fill.type === "Solid") {
        duplicate.fill.setSolidColor(shape.fill.foregroundColor);
        duplicate.fill.transparency = shape.fill.transparency;
      } else duplicate.fill.clear();
      duplicate.lineFormat.visible = shape.lineFormat.weight > 0;
      duplicate.lineFormat.color = shape.lineFormat.color;
      duplicate.lineFormat.weight = shape.lineFormat.weight;
      duplicate.lineFormat.dashStyle = shape.lineFormat.dashStyle;
      duplicate.lineFormat.transparency = shape.lineFormat.transparency;
      if (["TextBox", "GeometricShape"].includes(String(shape.type)) && shape.textFrame.hasText) {
        const source = shape.textFrame.textRange;
        applyText(duplicate, { font_name: source.font.name, font_size: source.font.size, font_color: source.font.color, bold: source.font.bold, italic: source.font.italic, alignment: String(source.paragraphFormat.horizontalAlignment).toLowerCase() }, source.text);
      }
      if (args.rotation !== undefined) duplicate.rotation = number(args.rotation);
      tagShape(duplicate, kind.isNullObject ? "duplicate" : kind.value, shapeType.isNullObject ? {} : { SHAPE_TYPE: shapeType.value });
      loadShapeResult(duplicate);
      await context.sync();
      return shapeResult(duplicate, Number(args.slide_index), { duplicated_from: shape.name });
    });
  }

  async function groupShapesAction(args) {
    if (!supports("1.8")) throw new Error("Office.js grouping requires PowerPointApi 1.8.");
    return PowerPoint.run(async (context) => {
      const slide = getSlide(context, args.slide_index);
      await assertShapeNameAvailable(context, slide, args.name);
      const shapes = slide.shapes;
      shapes.load("items/id,items/name");
      await context.sync();
      const names = new Set(args.shape_names.map(String));
      const members = shapes.items.filter((shape) => names.has(shape.name));
      if (members.length !== names.size) throw new Error("One or more shapes to group were not found by name.");
      const group = shapes.addGroup(members);
      if (args.name) group.name = String(args.name);
      tagShape(group, "group");
      loadShapeResult(group);
      await context.sync();
      return shapeResult(group, Number(args.slide_index), { member_count: members.length });
    });
  }

  async function ungroupShapeAction(args) {
    if (!supports("1.8")) throw new Error("Office.js ungrouping requires PowerPointApi 1.8.");
    return PowerPoint.run(async (context) => {
      const { slide, shape } = await findShape(context, args);
      if (shape.type !== "Group") throw new Error("Target shape is not a group.");
      const shapeName = shape.name;
      shape.group.ungroup();
      await context.sync();
      slide.shapes.load("items/id,items/name,items/type");
      await context.sync();
      return { slide_index: Number(args.slide_index), ungrouped_shape_name: shapeName, shape_count_after: slide.shapes.items.length, backend: "officejs-context-sync", context_sync: true };
    });
  }

  async function zOrderAction(args) {
    if (!supports("1.8")) throw new Error("Office.js z-order operations require PowerPointApi 1.8.");
    return PowerPoint.run(async (context) => {
      const { shape } = await findShape(context, args);
      const command = { bring_to_front: "BringToFront", send_to_back: "SendToBack", bring_forward: "BringForward", send_backward: "SendBackward" }[args.command];
      for (let index = 0; index < Number(args.repeat || 1); index += 1) shape.setZOrder(command);
      await context.sync();
      return { slide_index: Number(args.slide_index), shape_id: shape.id, shape_name: shape.name, command: args.command, repeat: Number(args.repeat || 1), backend: "officejs-context-sync", context_sync: true };
    });
  }

  async function exactLayoutAction(args, distribute) {
    return PowerPoint.run(async (context) => {
      const slide = getSlide(context, args.slide_index);
      const shapes = slide.shapes;
      shapes.load("items/id,items/name,items/left,items/top,items/width,items/height");
      if (args.relative_to === "slide") {
        if (!supports("1.10")) throw new Error("Layout relative to the slide requires PowerPointApi 1.10.");
        context.presentation.pageSetup.load("slideWidth,slideHeight");
      }
      await context.sync();
      const names = new Set(args.shape_names.map(String));
      const selected = shapes.items.filter((shape) => names.has(shape.name));
      if (selected.length !== names.size) throw new Error("One or more layout target shapes were not found by name.");
      if (distribute) {
        const horizontal = args.direction === "horizontal";
        selected.sort((a, b) => (horizontal ? a.left - b.left : a.top - b.top));
        const first = selected[0], last = selected[selected.length - 1];
        const totalSize = selected.reduce((sum, shape) => sum + (horizontal ? shape.width : shape.height), 0);
        const relativeToSlide = args.relative_to === "slide";
        const span = relativeToSlide
          ? (horizontal ? context.presentation.pageSetup.slideWidth : context.presentation.pageSetup.slideHeight)
          : (horizontal ? last.left + last.width - first.left : last.top + last.height - first.top);
        const gap = (span - totalSize) / (selected.length - 1);
        let cursor = relativeToSlide ? 0 : (horizontal ? first.left : first.top);
        for (const shape of selected) {
          if (horizontal) shape.left = cursor;
          else shape.top = cursor;
          cursor += (horizontal ? shape.width : shape.height) + gap;
        }
      } else {
        const relativeToSlide = args.relative_to === "slide";
        const minLeft = relativeToSlide ? 0 : Math.min(...selected.map((shape) => shape.left));
        const maxRight = relativeToSlide ? context.presentation.pageSetup.slideWidth : Math.max(...selected.map((shape) => shape.left + shape.width));
        const minTop = relativeToSlide ? 0 : Math.min(...selected.map((shape) => shape.top));
        const maxBottom = relativeToSlide ? context.presentation.pageSetup.slideHeight : Math.max(...selected.map((shape) => shape.top + shape.height));
        const target = args.alignment;
        for (const shape of selected) {
          if (target === "left") shape.left = minLeft;
          if (target === "center") shape.left = (minLeft + maxRight - shape.width) / 2;
          if (target === "right") shape.left = maxRight - shape.width;
          if (target === "top") shape.top = minTop;
          if (target === "middle") shape.top = (minTop + maxBottom - shape.height) / 2;
          if (target === "bottom") shape.top = maxBottom - shape.height;
        }
      }
      await context.sync();
      return { slide_index: Number(args.slide_index), shape_names: [...names], operation: distribute ? `distribute_${args.direction}` : `align_${args.alignment}`, exact_coordinates: true, backend: "officejs-context-sync", context_sync: true };
    });
  }

  async function updateShapeAction(args) {
    if (args.start_arrow !== undefined || args.end_arrow !== undefined) {
      throw new Error("Office.js cannot safely retarget arrowhead geometry in place. Delete and recreate the named line or connector with the requested arrowheads.");
    }
    return PowerPoint.run(async (context) => {
      const { slide, shape } = await findShape(context, args);
      if (args.new_name !== undefined) {
        await assertShapeNameAvailable(context, slide, args.new_name, shape.id);
        shape.name = String(args.new_name);
      }
      for (const property of ["left", "top", "width", "height", "rotation"]) if (args[property] !== undefined) shape[property] = number(args[property]);
      if (args.fill_color !== undefined || args.fill_transparency !== undefined) applyFill(shape, args);
      if (args.line_color !== undefined || args.line_width !== undefined || args.line_dash !== undefined || args.line_transparency !== undefined) applyLine(shape, args, "#000000", 1, true);
      const textSettings = [
        "text", "font_name", "font_size", "font_color", "bold", "italic", "alignment", "vertical_alignment",
        "margin_left", "margin_right", "margin_top", "margin_bottom", "word_wrap", "text_autofit",
      ];
      if (textSettings.some((key) => args[key] !== undefined)) {
        if (!["TextBox", "GeometricShape"].includes(String(shape.type))) throw new Error("Target shape does not support editable text through this Office.js adapter.");
        applyText(shape, args, args.text);
      }
      loadShapeResult(shape);
      await context.sync();
      return shapeResult(shape, Number(args.slide_index), { updated: true });
    });
  }

  async function deleteShapeAction(args) {
    if (args.confirm !== true) throw new Error("confirm=true is required to delete a shape.");
    return PowerPoint.run(async (context) => {
      const { shape } = await findShape(context, args);
      const deleted = { shape_id: shape.id, shape_name: shape.name };
      shape.delete();
      await context.sync();
      return { slide_index: Number(args.slide_index), ...deleted, deleted: true, backend: "officejs-context-sync", context_sync: true };
    });
  }

  function imageToJpeg(base64, width, height) {
    return new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = width || image.naturalWidth;
        canvas.height = height || image.naturalHeight;
        const context = canvas.getContext("2d");
        context.fillStyle = "#ffffff";
        context.fillRect(0, 0, canvas.width, canvas.height);
        context.drawImage(image, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", 0.94).split(",")[1]);
      };
      image.onerror = () => reject(new Error("Could not convert the rendered slide to JPEG."));
      image.src = `data:image/png;base64,${base64}`;
    });
  }

  async function exportSlideAction(args) {
    if (!supports("1.8")) throw new Error("Live slide rendering requires PowerPointApi 1.8.");
    if (!/\.(?:png|jpe?g)$/i.test(String(args.output_path || ""))) throw new Error("output_path must end with .png, .jpg, or .jpeg.");
    const rendered = await PowerPoint.run(async (context) => {
      const slide = getSlide(context, args.slide_index);
      const requestedWidth = Math.max(1, Math.round(number(args.width, 1920)));
      const heightWasExplicit = args.height !== undefined;
      const requestedHeight = Math.max(1, Math.round(heightWasExplicit ? number(args.height) : 1080));
      const preserveAspectRatio = args.preserve_aspect_ratio !== false;
      if (preserveAspectRatio && !supports("1.10")) {
        throw new Error("Aspect-ratio-preserving Office.js export requires PowerPointApi 1.10 page-size metadata. Upgrade PowerPoint or pass preserve_aspect_ratio=false with explicit width and height.");
      }
      if (preserveAspectRatio) context.presentation.pageSetup.load("slideWidth,slideHeight");
      await context.sync();
      let width = requestedWidth;
      let height = requestedHeight;
      if (preserveAspectRatio) {
        const slideWidth = number(context.presentation.pageSetup.slideWidth);
        const slideHeight = number(context.presentation.pageSetup.slideHeight);
        if (slideWidth <= 0 || slideHeight <= 0) throw new Error("PowerPoint returned invalid slide dimensions for image export.");
        if (heightWasExplicit) {
          const scale = Math.min(requestedWidth / slideWidth, requestedHeight / slideHeight);
          width = Math.max(1, Math.round(slideWidth * scale));
          height = Math.max(1, Math.round(slideHeight * scale));
        } else {
          height = Math.max(1, Math.round(requestedWidth * slideHeight / slideWidth));
        }
      }
      const image = slide.getImageAsBase64({ width, height });
      await context.sync();
      return { image_base64: image.value, requested_width: requestedWidth, requested_height: requestedHeight, width, height, aspect_ratio_preserved: preserveAspectRatio };
    });
    const jpeg = /\.jpe?g$/i.test(String(args.output_path || ""));
    return {
      slide_index: Number(args.slide_index),
      image_base64: jpeg ? await imageToJpeg(rendered.image_base64, rendered.width, rendered.height) : rendered.image_base64,
      requested_width: rendered.requested_width,
      requested_height: rendered.requested_height,
      width: rendered.width,
      height: rendered.height,
      aspect_ratio_preserved: rendered.aspect_ratio_preserved,
      mime_type: jpeg ? "image/jpeg" : "image/png",
      renderer: "PowerPoint Office.js Slide.getImageAsBase64",
      backend: "officejs-context-sync",
    };
  }

  async function saveAction(args) {
    if (!args.output_path) return { saved_in_place: false, backend: "officejs-context-sync", note: "Office.js cannot force a desktop Save command. Use output_path to export the current editable presentation, or save normally in PowerPoint." };
    if (String(args.format || "pptx").toLowerCase() === "pdf" || /\.pdf$/i.test(args.output_path)) throw new Error("Office.js cannot export PDF. Save PPTX first, then use PowerPoint or the OOXML renderer for PDF.");
    if (!/\.pptx$/i.test(String(args.output_path))) throw new Error("Office.js editable presentation export requires a .pptx output_path.");
    if (!supports("1.10")) throw new Error("Editable PPTX export requires PowerPointApi 1.10.");
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      const result = slides.exportAsBase64Presentation(slides.items);
      await context.sync();
      return { file_base64: result.value, mime_type: "application/vnd.openxmlformats-officedocument.presentationml.presentation", backend: "officejs-context-sync", exported_editable_presentation: true };
    });
  }

  async function dispatchAction(action, args) {
    const actions = {
      status: statusAction,
      capabilities: capabilitiesAction,
      launch: statusAction,
      inspect: inspectAction,
      audit_figure: auditAction,
      activate_slide: activateSlideAction,
      add_slide: addSlideAction,
      add_textbox: addTextboxAction,
      add_shape: addShapeAction,
      add_image: addImageAction,
      add_line: addLineAction,
      add_connector: addConnectorAction,
      add_table: addTableAction,
      update_table_cell: updateTableCellAction,
      update_table_layout: updateTableLayoutAction,
      add_chart: addChartAction,
      duplicate_shape: duplicateShapeAction,
      group_shapes: groupShapesAction,
      ungroup_shape: ungroupShapeAction,
      set_z_order: zOrderAction,
      align_shapes: (value) => exactLayoutAction(value, false),
      distribute_shapes: (value) => exactLayoutAction(value, true),
      update_shape: updateShapeAction,
      delete_shape: deleteShapeAction,
      export_slide_image: exportSlideAction,
      save: saveAction,
    };
    if (action === "new_presentation") throw new Error("Office.js controls the current PowerPoint presentation and cannot create a new desktop presentation. Open a blank deck, load this task pane, and retry.");
    if (action === "close_presentation" || action === "quit_application") throw new Error("Office.js deliberately does not close the current deck or quit PowerPoint.");
    const handler = actions[action];
    if (!handler) throw new Error(`Unsupported Office.js action: ${action}`);
    return handler(args || {});
  }

  function errorValue(error) {
    return {
      message: error?.message || String(error),
      name: error?.name || "Error",
      code: error?.code,
      debug_info: error?.debugInfo,
    };
  }

  async function sendResult(command, ok, value) {
    const payload = {
      client_id: clientId(),
      command_id: command.id,
      ok,
      ...(ok ? { result: value } : { error: errorValue(value) }),
    };
    let lastError;
    for (let attempt = 0; attempt < 3; attempt += 1) {
      try {
        await request("/api/result", { method: "POST", body: JSON.stringify(payload) });
        return;
      } catch (error) {
        lastError = error;
        await sleep(300 * (attempt + 1));
      }
    }
    throw lastError;
  }

  async function commandLoop() {
    let retryMs = 500;
    while (!stopped) {
      try {
        const response = await request(`/api/command?client_id=${encodeURIComponent(clientId())}`);
        retryMs = 500;
        if (!response?.command) continue;
        const command = response.command;
        setStatus("busy", "正在绘制", `${command.action} · context.sync()`);
        try {
          const result = await dispatchAction(command.action, command.arguments || {});
          await sendResult(command, true, result);
          setStatus("connected", "已连接，等待绘制命令", "每个对象都会在 context.sync() 后立即显示。");
        } catch (error) {
          await sendResult(command, false, error);
          setStatus("error", "上一条命令未完成", errorMessage(error));
        }
      } catch (error) {
        if (errorStatus(error) === 401) {
          setStatus("waiting", "会话已更新，正在重新加载", "本机桥接器重新启动后需要刷新任务窗格令牌。");
          await sleep(700);
          location.reload();
          return;
        }
        setStatus("error", "本机桥接器暂时不可用", errorMessage(error));
        await sleep(retryMs);
        retryMs = Math.min(5000, retryMs * 1.6);
      }
    }
  }

  async function connect() {
    stopped = false;
    setStatus("waiting", "正在连接本机桥接器…", "正在注册当前 PowerPoint 任务窗格。");
    await request("/api/register", { method: "POST", body: JSON.stringify(registration()) });
    setStatus("connected", "已连接，等待绘制命令", "Codex 现在可以直接控制当前幻灯片的 Office.js 对象模型。");
    await commandLoop();
  }

  reconnectButton.addEventListener("click", () => {
    stopped = true;
    location.reload();
  });

  Office.onReady((info) => {
    hostInfo = info;
    if (info.host !== Office.HostType.PowerPoint) {
      setStatus("error", "请在 Microsoft PowerPoint 中使用", `当前宿主：${info.host || "unknown"}`);
      return;
    }
    if (!supports("1.4")) {
      setStatus("error", "PowerPoint 版本过旧", "Office.js 实时绘制至少需要 PowerPointApi 1.4。");
      return;
    }
    connect().catch((error) => setStatus("error", "连接失败", error.message));
  });
}());
