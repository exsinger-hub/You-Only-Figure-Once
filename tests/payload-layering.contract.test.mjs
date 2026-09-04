import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const testDir = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(testDir, "..");
const bridge = readFileSync(path.join(root, "scripts", "powerpoint-bridge.ps1"), "utf8");
const server = readFileSync(path.join(root, "scripts", "powerpoint-server.mjs"), "utf8");

function extractPowerShellFunction(source, name) {
  const start = source.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `Missing PowerShell function ${name}.`);
  const next = source.indexOf("\nfunction ", start + 1);
  return source.slice(start, next === -1 ? source.length : next);
}

test("large COM batches use a payload file instead of the Windows command line", () => {
  assert.match(bridge, /\[string\]\$PayloadFile/);
  assert.match(bridge, /ReadAllText\(\$payloadPath, \[Text\.Encoding\]::UTF8\)/);
  assert.match(server, /payload\.length <= 12000/);
  assert.match(server, /\["-PayloadFile", payloadFile\]/);
  assert.match(server, /fs\.unlink\(payloadFile\)/);
});

test("native connectors remain above a full-slide background", () => {
  const addConnector = extractPowerShellFunction(bridge, "Invoke-AddConnector");
  assert.doesNotMatch(addConnector, /\$shape\.ZOrder\(1\)/, "SendToBack hides connectors below a slide-sized background.");
  assert.match(addConnector, /\$shape\.ZOrder\(3\)/, "A single SendBackward step keeps the route behind its endpoints.");
});

test("raster decomposition notes cannot trigger an affirmative composite-raster failure", () => {
  assert.match(bridge, /\$compositeText = "\$\(\$shape\.Name\) \$reason"/);
  assert.doesNotMatch(bridge, /\$compositeText = "\$\(\$shape\.Name\) \$reason \$decomposition"/);
});

test("polyline strokes are excluded from repeated-box alignment checks", () => {
  assert.match(bridge, /\$summaries \| Where-Object \{ \[int\]\$_\.type -ne 9 \}/);
});
