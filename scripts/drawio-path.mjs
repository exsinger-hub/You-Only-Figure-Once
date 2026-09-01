import { existsSync } from "node:fs";
import os from "node:os";
import path from "node:path";

function pathApi(platform) {
  return platform === "win32" ? path.win32 : path.posix;
}

function expandHome(value, platform, homeDir) {
  if (!value) return value;
  if (value === "~") return homeDir;
  if (value.startsWith("~/") || value.startsWith("~\\")) {
    const relative = value.slice(2).replace(/[\\/]+/g, pathApi(platform).sep);
    return pathApi(platform).join(homeDir, relative);
  }
  return value;
}

export function drawioExecutableCandidates({ platform = process.platform, env = process.env, homeDir = os.homedir() } = {}) {
  const platformPath = pathApi(platform);
  if (platform === "win32") {
    return [
      env.ProgramFiles && platformPath.join(env.ProgramFiles, "draw.io", "draw.io.exe"),
      env["ProgramFiles(x86)"] && platformPath.join(env["ProgramFiles(x86)"], "draw.io", "draw.io.exe"),
      env.LOCALAPPDATA && platformPath.join(env.LOCALAPPDATA, "Programs", "draw.io", "draw.io.exe"),
      env.LOCALAPPDATA && platformPath.join(env.LOCALAPPDATA, "draw.io", "draw.io.exe"),
    ].filter(Boolean);
  }

  if (platform === "darwin") {
    return [
      "/Applications/draw.io.app/Contents/MacOS/draw.io",
      platformPath.join(homeDir, "Applications", "draw.io.app", "Contents", "MacOS", "draw.io"),
    ];
  }

  return ["/usr/bin/drawio", "/usr/local/bin/drawio", "/snap/bin/drawio", "/opt/drawio/drawio"];
}

export function resolveDrawioExecutable({
  platform = process.platform,
  env = process.env,
  homeDir = os.homedir(),
  exists = existsSync,
} = {}) {
  const configured = expandHome(env.DRAWIO_PATH?.trim(), platform, homeDir);
  if (configured) return { executable: configured, source: "DRAWIO_PATH" };

  for (const candidate of drawioExecutableCandidates({ platform, env, homeDir })) {
    if (exists(candidate)) return { executable: candidate, source: "auto-detected" };
  }

  return {
    executable: platform === "win32" ? "draw.io.exe" : "drawio",
    source: "PATH fallback",
  };
}

export function drawioInstallHint() {
  return "Install the draw.io desktop app from https://www.drawio.com/ or set DRAWIO_PATH to its executable.";
}
