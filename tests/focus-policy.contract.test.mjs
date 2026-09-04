import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const testDirectory = path.dirname(fileURLToPath(import.meta.url));
const repositoryRoot = path.resolve(testDirectory, "..");
const comBridgePath = path.join(repositoryRoot, "scripts", "powerpoint-bridge.ps1");
const ooxmlBridgePath = path.join(repositoryRoot, "scripts", "powerpoint-mac-bridge.py");
const serverPath = path.join(repositoryRoot, "scripts", "powerpoint-server.mjs");

const comBridge = readFileSync(comBridgePath, "utf8");
const ooxmlBridge = readFileSync(ooxmlBridgePath, "utf8");
const server = readFileSync(serverPath, "utf8");

function extractPowerShellFunction(source, name) {
  const startPattern = new RegExp(`^function\\s+${name.replaceAll("-", "\\-")}\\s*\\{`, "mi");
  const match = startPattern.exec(source);
  assert.ok(match, `PowerShell function ${name} must exist`);
  const start = match.index;
  const remainder = source.slice(start + match[0].length);
  const nextFunction = /^function\s+[A-Za-z0-9-]+\s*\{/mi.exec(remainder);
  return source.slice(start, nextFunction ? start + match[0].length + nextFunction.index : source.length).trim();
}

function extractJavaScriptFunction(source, declaration, nextDeclaration) {
  const start = source.indexOf(declaration);
  assert.notEqual(start, -1, `${declaration} must exist`);
  const end = source.indexOf(nextDeclaration, start + declaration.length);
  assert.notEqual(end, -1, `${nextDeclaration} must follow ${declaration}`);
  return source.slice(start, end);
}

function extractPythonFunction(source, name, nextName) {
  const declaration = `def ${name}(`;
  const nextDeclaration = `def ${nextName}(`;
  const start = source.indexOf(declaration);
  assert.notEqual(start, -1, `${declaration} must exist`);
  const end = source.indexOf(nextDeclaration, start + declaration.length);
  assert.notEqual(end, -1, `${nextDeclaration} must follow ${declaration}`);
  return source.slice(start, end);
}

function runPowerShellJson(program) {
  const encodedProgram = Buffer.from(program, "utf16le").toString("base64");
  const result = spawnSync("pwsh", ["-NoProfile", "-NonInteractive", "-EncodedCommand", encodedProgram], {
    cwd: repositoryRoot,
    encoding: "utf8",
  });
  assert.equal(result.status, 0, result.stderr || result.stdout);
  return JSON.parse(result.stdout || "[]");
}

function runShowSlide({ policy, forceForeground = false }) {
  const showSlide = extractPowerShellFunction(comBridge, "Show-Slide");
  return runPowerShellJson(String.raw`
$ErrorActionPreference = "Stop"
$script:Events = [System.Collections.Generic.List[string]]::new()
class MockSlideView {
    [System.Collections.Generic.List[string]] $Events
    MockSlideView([System.Collections.Generic.List[string]] $Events) {
        $this.Events = $Events
    }
    [void] GotoSlide([int] $Index) {
        $this.Events.Add("GotoSlide:$Index")
    }
}
class MockPresentationWindow {
    [int] $ViewType = 0
    [MockSlideView] $View
    [System.Collections.Generic.List[string]] $Events
    MockPresentationWindow([System.Collections.Generic.List[string]] $Events) {
        $this.Events = $Events
        $this.View = [MockSlideView]::new($Events)
    }
    [void] Activate() {
        $this.Events.Add("Activate")
    }
}
$window = [MockPresentationWindow]::new($script:Events)
$application = [pscustomobject]@{ ActiveWindow = $window }
${showSlide}
$script:FocusPolicy = "${policy}"
Show-Slide $application 7 $${forceForeground ? "true" : "false"}
$json = @($script:Events) | ConvertTo-Json -Compress -AsArray
[Console]::Out.Write($json)
`);
}

function runPreserveRestore({
  currentHandle,
  preservedHandle,
  controlledHandle = 99,
  policy = "preserve",
  explicit = false,
  force = false,
}) {
  const restoreIfChanged = extractPowerShellFunction(comBridge, "Restore-PreservedForegroundWindowIfChanged");
  return runPowerShellJson(String.raw`
$ErrorActionPreference = "Stop"
$script:Events = [System.Collections.Generic.List[string]]::new()
function Get-ForegroundWindowHandle { return [IntPtr]${currentHandle} }
function Restore-ForegroundWindow {
    param([IntPtr] $WindowHandle)
    $script:Events.Add("SetForegroundWindow:$([int64] $WindowHandle)")
}
${restoreIfChanged}
$script:FocusPolicy = "${policy}"
$script:ExplicitActivationPerformed = $${explicit ? "true" : "false"}
$script:PreservedForegroundWindow = [IntPtr]${preservedHandle}
$script:ControlledPowerPointWindow = [IntPtr]${controlledHandle}
Restore-PreservedForegroundWindowIfChanged $${force ? "true" : "false"}
$json = @($script:Events) | ConvertTo-Json -Compress -AsArray
[Console]::Out.Write($json)
`);
}

test("preserve Show-Slide performs no navigation or activation", () => {
  assert.deepEqual(
    runShowSlide({ policy: "preserve" }),
    [],
    "A preserve-mode mutation must not call GotoSlide or Activate. Selecting a slide in a background PowerPoint window is still a UI side effect.",
  );
});

test("foreground Show-Slide navigates and activates exactly once", () => {
  assert.deepEqual(runShowSlide({ policy: "foreground" }), ["GotoSlide:7", "Activate"]);
});

test("a standalone explicit activate_slide request may foreground under preserve policy", () => {
  assert.deepEqual(
    runShowSlide({ policy: "preserve", forceForeground: true }),
    ["GotoSlide:7", "Activate"],
  );
  const activateSlide = extractPowerShellFunction(comBridge, "Invoke-ActivateSlide");
  assert.match(
    activateSlide,
    /Show-Slide\s+\$application\s+\$index\s+\$true/,
    "Invoke-ActivateSlide must remain the explicit force-foreground path.",
  );
});

test("new/open presentation activation remains gated by foreground policy", () => {
  for (const name of ["Invoke-NewPresentation", "Invoke-Launch"]) {
    const body = extractPowerShellFunction(comBridge, name);
    assert.match(
      body,
      /if\s*\(\$script:FocusPolicy\s+-eq\s+"foreground"\)\s*\{[\s\S]*?\.Windows\.Item\(1\)\.Activate\(\)/,
      `${name} must not activate a presentation window outside the foreground-policy guard.`,
    );
  }
});

test("PowerPoint window-focus primitives remain behind their named choke points", () => {
  assert.equal((comBridge.match(/ActiveWindow\.View\.GotoSlide\s*\(/g) || []).length, 1);
  assert.equal((comBridge.match(/ActiveWindow\.Activate\s*\(/g) || []).length, 1);
  assert.equal((comBridge.match(/\.Windows\.Item\(1\)\.Activate\s*\(/g) || []).length, 2);
  assert.equal((comBridge.match(/\]::SetForegroundWindow\s*\(/g) || []).length, 1);

  const showSlide = extractPowerShellFunction(comBridge, "Show-Slide");
  assert.match(showSlide, /ActiveWindow\.View\.GotoSlide\s*\(/);
  assert.match(showSlide, /ActiveWindow\.Activate\s*\(/);
  const restore = extractPowerShellFunction(comBridge, "Restore-ForegroundWindow");
  assert.match(restore, /\]::SetForegroundWindow\s*\(/);
});

test("native chart-data activation immediately invokes conditional focus recovery", () => {
  const addChart = extractPowerShellFunction(comBridge, "Invoke-AddChart");
  assert.match(
    addChart,
    /\$chartData\.Activate\(\)[\s\S]{0,320}?Restore-PreservedForegroundWindowIfChanged/,
    "ChartData.Activate is not presentation-window activation, but it can surface the embedded workbook and therefore requires immediate compare-before-restore recovery.",
  );
});

test("ordinary COM preserve calls avoid SetForegroundWindow when focus did not change", () => {
  assert.deepEqual(runPreserveRestore({ currentHandle: 41, preservedHandle: 41 }), []);
});

test("preserve fallback restores only when controlled PowerPoint took foreground", () => {
  assert.deepEqual(
    runPreserveRestore({ currentHandle: 99, preservedHandle: 41 }),
    ["SetForegroundWindow:41"],
  );
  assert.deepEqual(
    runPreserveRestore({ currentHandle: 88, preservedHandle: 41, controlledHandle: 99 }),
    [],
    "A user-initiated switch to a third app must not be overwritten by batch cleanup.",
  );
  assert.deepEqual(
    runPreserveRestore({ currentHandle: 88, preservedHandle: 41, controlledHandle: 99, force: true }),
    ["SetForegroundWindow:41"],
    "A known focus-stealing Office API may request immediate recovery.",
  );
  assert.deepEqual(
    runPreserveRestore({ currentHandle: 99, preservedHandle: 41, policy: "foreground" }),
    [],
  );
  assert.deepEqual(
    runPreserveRestore({ currentHandle: 99, preservedHandle: 41, explicit: true }),
    [],
  );
});

test("the COM entry point uses the conditional restore helper", () => {
  const entryStart = comBridge.search(/^try\s*\{\s*\r?\n\s*\$json/m);
  assert.notEqual(entryStart, -1);
  const entryPoint = comBridge.slice(entryStart);
  assert.match(
    entryPoint,
    /Restore-PreservedForegroundWindowIfChanged/,
    "The entry point must not call SetForegroundWindow directly or restore an unchanged foreground handle.",
  );
  assert.doesNotMatch(entryPoint, /Restore-ForegroundWindow\s+\$previousForegroundWindow/);
});

test("Windows OOXML preserve refresh avoids SetForegroundWindow when focus is unchanged", () => {
  const openWindowsPresentation = extractPythonFunction(
    ooxmlBridge,
    "_open_windows_presentation",
    "_parse_macos_process_ids",
  );
  assert.match(
    openWindowsPresentation,
    /current_foreground\s*=\s*user32\.GetForegroundWindow\(\)[\s\S]{0,240}?if\s+current_foreground\s*!=\s*previous_foreground\s*:[\s\S]{0,160}?user32\.SetForegroundWindow\(previous_foreground\)/,
    "SW_SHOWNOACTIVATE should normally leave ownership untouched. If a host ignores it, compare handles before recovery instead of calling SetForegroundWindow at every sequence checkpoint.",
  );
  assert.match(openWindowsPresentation, /SW_SHOWNOACTIVATE/);
});

test("OOXML activate_slide explicitly selects foreground policy", () => {
  const activateSlide = extractPythonFunction(ooxmlBridge, "action_activate_slide", "action_refresh");
  assert.match(activateSlide, /focus_policy\s*=\s*"foreground"/);
});

test("COM apply_sequence suppresses ordinary slide-view navigation", () => {
  const applySequence = extractPowerShellFunction(comBridge, "Invoke-ApplySequence");
  assert.match(
    applySequence,
    /\$script:SuppressSlideView\s*=\s*\$true[\s\S]*?for\s*\(\$index\s*=\s*0/,
    "The batch must enable slide-view suppression before dispatching its first operation.",
  );
  assert.match(
    applySequence,
    /if\s*\(\$script:FocusPolicy\s+-eq\s+"foreground"[\s\S]{0,180}?Show-Slide/,
    "Only a foreground-policy checkpoint may temporarily show a slide during a batch.",
  );
});

test("preserve apply_sequence never performs an unconditional foreground restore inside the operation loop", () => {
  const applySequence = extractPowerShellFunction(comBridge, "Invoke-ApplySequence");
  const loopStart = applySequence.indexOf("for ($index = 0");
  const finalizerStart = applySequence.indexOf("\n    finally {", loopStart);
  assert.notEqual(loopStart, -1);
  assert.notEqual(finalizerStart, -1);
  const loopAndSuccessTail = applySequence.slice(loopStart, finalizerStart);
  assert.doesNotMatch(
    loopAndSuccessTail,
    /Restore-ForegroundWindow/,
    "The loop may call the compare-before-restore helper after an unavoidable host focus change, but must never call SetForegroundWindow unconditionally.",
  );
});

test("COM apply_sequence includes an error-safe final foreground restore", () => {
  const applySequence = extractPowerShellFunction(comBridge, "Invoke-ApplySequence");
  assert.match(
    applySequence,
    /finally\s*\{[\s\S]*?Restore-PreservedForegroundWindowIfChanged/,
    "The batch must restore the original foreground in an error-safe finalizer.",
  );
});

test("the server sends a COM sequence through one bridge process", () => {
  const runSequence = extractJavaScriptFunction(
    server,
    "async function runSequence(args)",
    "async function handleTool(name, args = {})",
  );
  assert.match(runSequence, /if\s*\(sequenceBackend\s*===\s*"com"\)/);
  assert.match(runSequence, /runBridge\("apply_sequence"[\s\S]*?,\s*"com"\)/);
  assert.match(runSequence, /delete\s+operation\.focus_policy/);
  assert.match(runSequence, /bridge_process_count:\s*1/);
});

test("a default preserve COM sequence uses fast pacing with zero inter-step delay", () => {
  const runSequence = extractJavaScriptFunction(
    server,
    "async function runSequence(args)",
    "async function handleTool(name, args = {})",
  );
  assert.match(
    runSequence,
    /const\s+comPacingMode\s*=\s*requestedFocusPolicy\s*===\s*"preserve"\s*&&\s*args\.pacing_mode\s*===\s*undefined\s*\?\s*"fast"\s*:\s*pacingMode/,
  );
  assert.match(
    runSequence,
    /const\s+comStepDelay\s*=\s*requestedFocusPolicy\s*===\s*"preserve"\s*&&\s*args\.pacing_mode\s*===\s*undefined\s*\?\s*0\s*:\s*requestedDelay/,
  );
  assert.match(runSequence, /step_delay_ms:\s*comStepDelay[\s\S]*?pacing_mode:\s*comPacingMode/);
});

test("a preserve COM batch rejects activate_slide before opening the bridge", () => {
  const runSequence = extractJavaScriptFunction(
    server,
    "async function runSequence(args)",
    "async function handleTool(name, args = {})",
  );
  const rejection = runSequence.indexOf('type === "activate_slide" && requestedFocusPolicy === "preserve"');
  const dispatch = runSequence.indexOf('runBridge("apply_sequence"');
  assert.notEqual(rejection, -1);
  assert.notEqual(dispatch, -1);
  assert.ok(rejection < dispatch, "The full operation list must reject preserve+activate_slide before any COM bridge process is opened.");
  assert.match(
    runSequence.slice(rejection, dispatch),
    /focus_policy to foreground[\s\S]*?No object was dispatched/,
  );
});

test("draw_sequence never implicitly upgrades an ordinary operation to foreground", () => {
  const runSequence = extractJavaScriptFunction(
    server,
    "async function runSequence(args)",
    "async function handleTool(name, args = {})",
  );
  assert.doesNotMatch(runSequence, /focus_policy\s*[:=]\s*["']foreground["']/);
  assert.match(runSequence, /activate_slide:\s*"activate_slide"/);
});
