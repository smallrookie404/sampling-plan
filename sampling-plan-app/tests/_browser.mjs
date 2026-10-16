// 测试公用：定位可用的 Chrome / Edge 浏览器可执行文件。
// 之前各冒烟测试各自硬编码 "C:/Program Files/Google/Chrome/..."，只装了 Edge 的机器上会直接失败；
// 这里统一改成「按候选路径 + 环境变量」依次探测，找不到就明确跳过（不误报失败）。
import fs from "node:fs";

const CANDIDATES = [
  process.env.CHROME_PATH,
  process.env.EDGE_PATH,
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
  "/usr/bin/google-chrome",
  "/usr/bin/chromium",
  "/usr/bin/chromium-browser",
];

// 返回第一个存在的浏览器路径；找不到返回 null（调用方自行决定跳过）。
export function findBrowser() {
  for (const p of CANDIDATES) {
    if (p && fs.existsSync(p)) return p;
  }
  return null;
}
