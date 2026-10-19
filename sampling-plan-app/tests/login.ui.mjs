// 临时验证：登录界面默认不显示验证码（captchaField hidden），且「改用验证码登录」入口默认隐藏。
import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright-core";
import { findBrowser } from "./_browser.mjs";

const here = path.dirname(fileURLToPath(import.meta.url));
const APP = path.resolve(here, "..");
const exe = findBrowser();
if (!exe) { console.log("跳过：未找到浏览器"); process.exit(0); }

const srv = spawn(process.execPath, ["server.mjs"], { cwd: APP, stdio: "ignore", env: { ...process.env, NO_PROXY: "*", no_proxy: "*" } });
await new Promise((r) => setTimeout(r, 1500));

let pass = 0, fail = 0;
const check = (n, ok, extra) => { if (ok) { pass++; console.log("  ok  " + n); } else { fail++; console.log("  FAIL " + n + (extra ? " :: " + extra : "")); } };

try {
  const browser = await chromium.launch({ executablePath: exe, headless: true });
  const page = await browser.newPage();
  await page.goto("http://127.0.0.1:8017/index.html", { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(800);

  check("登录遮罩可见", await page.isVisible("#loginPanel"));
  check("账号输入框可见", await page.isVisible("#username"));
  check("密码输入框可见", await page.isVisible("#password"));

  const capVisible = await page.isVisible("#captcha");
  check("验证码输入框不可见", capVisible === false, "isVisible=" + capVisible);
  const imgVisible = await page.isVisible("#captchaImg");
  check("验证码图片不可见", imgVisible === false, "isVisible=" + imgVisible);
  const rowVisible = await page.isVisible("#captchaField");
  check("验证码整行不可见", rowVisible === false, "isVisible=" + rowVisible);
  check("captchaField 带 hidden 属性", await page.getAttribute("#captchaField", "hidden") !== null);
  check("「改用验证码登录」入口默认隐藏", (await page.isVisible("#btnCaptchaFallback")) === false);

  // 兜底入口能正常展开验证码
  await page.evaluate(() => { document.getElementById("btnCaptchaFallback").hidden = false; });
  await page.click("#btnCaptchaFallback");
  await page.waitForTimeout(300);
  check("点击兜底入口后验证码输入框出现", await page.isVisible("#captcha"));
  check("点击兜底入口后兜底链接消失", (await page.isVisible("#btnCaptchaFallback")) === false);

  await browser.close();
} finally {
  srv.kill();
}

console.log("\n" + pass + " 通过 / " + fail + " 失败");
process.exit(fail ? 1 : 0);
