// 登录界面：按运行环境区分两种登录方式
//   · 云端（桥 ping 返回 cloud:true）→ 云端代登录，界面不显示验证码
//   · 本地（file:// / 本地服务）      → 沿用原步骤：显示验证码，手输
// 另含密码框「眼睛」切换明文/密文的检查。
import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright-core";
import { findBrowser } from "./_browser.mjs";

const here = path.dirname(fileURLToPath(import.meta.url));
const APP = path.resolve(here, "..");
const exe = findBrowser();
if (!exe) { console.log("跳过：未找到浏览器"); process.exit(0); }

const srv = spawn(process.execPath, ["server.mjs"], {
  cwd: APP, stdio: "ignore",
  env: { ...process.env, NO_PROXY: "*", no_proxy: "*" },
});
await new Promise((r) => setTimeout(r, 1500));

let pass = 0, fail = 0;
const check = (n, ok, extra) => { if (ok) { pass++; console.log("  ok  " + n); } else { fail++; console.log("  FAIL " + n + (extra ? " :: " + extra : "")); } };
const URL_APP = "http://127.0.0.1:8017/index.html";

try {
  const browser = await chromium.launch({ executablePath: exe, headless: true });

  // ========== A. 云端模式（把 /api/bridge/ping 打桩成 cloud:true）==========
  console.log("\n[A] 云端部署：界面不显示验证码");
  {
    const page = await browser.newPage();
    await page.route("**/api/bridge/ping", (route) =>
      route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ ok: true, autoStart: false, up: true, cloud: true }) }));
    await page.goto(URL_APP, { waitUntil: "domcontentloaded" });
    await page.waitForTimeout(800);

    check("登录遮罩可见", await page.isVisible("#loginPanel"));
    check("账号输入框可见", await page.isVisible("#username"));
    check("密码输入框可见", await page.isVisible("#password"));
    check("验证码输入框不可见", (await page.isVisible("#captcha")) === false);
    check("验证码图片不可见", (await page.isVisible("#captchaImg")) === false);
    check("验证码整行不可见", (await page.isVisible("#captchaField")) === false);
    check("captchaField 带 hidden 属性", (await page.getAttribute("#captchaField", "hidden")) !== null);
    check("「改用验证码登录」入口默认隐藏", (await page.isVisible("#btnCaptchaFallback")) === false);

    // 兜底入口能展开验证码（云端登录失败时的逃生路径）
    await page.evaluate(() => { document.getElementById("btnCaptchaFallback").hidden = false; });
    await page.click("#btnCaptchaFallback");
    await page.waitForTimeout(300);
    check("点击兜底入口后验证码输入框出现", await page.isVisible("#captcha"));
    check("点击兜底入口后兜底链接消失", (await page.isVisible("#btnCaptchaFallback")) === false);
    await page.close();
  }

  // ========== B. 本地模式（不打桩，本地服务 ping 无 cloud 字段）==========
  console.log("\n[B] 本地打开：沿用原来的手输验证码");
  const page = await browser.newPage();
  await page.goto(URL_APP, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(1200);

  check("登录遮罩可见", await page.isVisible("#loginPanel"));
  check("验证码输入框可见", await page.isVisible("#captcha"));
  check("验证码图片可见", await page.isVisible("#captchaImg"));
  check("验证码整行可见", await page.isVisible("#captchaField"));
  check("captchaField 无 hidden 属性", (await page.getAttribute("#captchaField", "hidden")) === null);
  check("本地不显示「改用验证码登录」入口", (await page.isVisible("#btnCaptchaFallback")) === false);

  // 图片是否真的取到了（本地直连平台；取不到则 loginMsg 会有错误提示）
  const imgOk = await page.waitForFunction(() => {
    const el = document.getElementById("captchaImg");
    return !!el && /^data:image\//.test(el.getAttribute("src") || "");
  }, { timeout: 15000 }).then(() => true).catch(() => false);
  const msg = ((await page.textContent("#loginMsg")) || "").trim();
  check("验证码图片已取到（data URL）", imgOk, "loginMsg=" + JSON.stringify(msg));
  check("验证码输入框为空（不自动填）", (await page.inputValue("#captcha")) === "");

  // ========== C. 密码框「眼睛」==========
  console.log("\n[C] 密码框眼睛切换明文/密文");
  const eye = "#btnPwdEye";
  check("眼睛按钮可见", await page.isVisible(eye));
  check("密码框默认 type=password", (await page.getAttribute("#password", "type")) === "password");
  check("眼睛默认 data-shown=false", (await page.getAttribute(eye, "data-shown")) === "false");
  check("眼睛默认提示「显示密码」", (await page.getAttribute(eye, "title")) === "显示密码");
  check("眼睛默认 aria-pressed=false", (await page.getAttribute(eye, "aria-pressed")) === "false");
  check("默认只显示 open 图标（off 隐藏）", await page.evaluate(() => {
    const b = document.getElementById("btnPwdEye");
    return getComputedStyle(b.querySelector(".ico-on")).display !== "none"
      && getComputedStyle(b.querySelector(".ico-off")).display === "none";
  }));

  await page.fill("#password", "abc123");
  check("可正常输入密码", (await page.inputValue("#password")) === "abc123");

  await page.click(eye);
  check("点击后 type=text", (await page.getAttribute("#password", "type")) === "text");
  check("点击后 data-shown=true", (await page.getAttribute(eye, "data-shown")) === "true");
  check("点击后提示「隐藏密码」", (await page.getAttribute(eye, "title")) === "隐藏密码");
  check("点击后 aria-pressed=true", (await page.getAttribute(eye, "aria-pressed")) === "true");
  check("明文时值不变", (await page.inputValue("#password")) === "abc123");
  check("明文时改显 off 图标", await page.evaluate(() => {
    const b = document.getElementById("btnPwdEye");
    return getComputedStyle(b.querySelector(".ico-on")).display === "none"
      && getComputedStyle(b.querySelector(".ico-off")).display !== "none";
  }));

  await page.click(eye);
  check("再次点击回到 type=password", (await page.getAttribute("#password", "type")) === "password");
  check("再次点击 data-shown=false", (await page.getAttribute(eye, "data-shown")) === "false");

  // 点眼睛不应触发登录
  const msgBefore = (await page.textContent("#loginMsg") || "").trim();
  await page.click(eye);
  await page.waitForTimeout(400);
  const msgAfter = (await page.textContent("#loginMsg") || "").trim();
  check("点击眼睛不会触发登录", msgBefore === msgAfter, JSON.stringify({ msgBefore, msgAfter }));

  // ========== D. 本地模式登录不走云端代登录 ==========
  console.log("\n[D] 本地登录：浏览器直接调平台，不经过 /api/bridge/login");
  let bridgeLoginCalled = false;
  let platformLoginCalled = false;
  page.on("request", (r) => {
    const u = r.url();
    if (u.includes("/api/bridge/login")) bridgeLoginCalled = true;
    if (/\/auth\/login$/.test(u)) platformLoginCalled = true;
  });
  await page.fill("#username", "btc-wf");
  await page.fill("#password", "123456");
  await page.fill("#captcha", "1");        // 故意填错，只为观察走哪条链路
  await page.click("#btnLogin");
  await page.waitForTimeout(8000);
  check("本地登录不调用 /api/bridge/login", bridgeLoginCalled === false, "bridgeLoginCalled=" + bridgeLoginCalled);
  check("本地登录由浏览器直调平台 /auth/login", platformLoginCalled === true, "platformLoginCalled=" + platformLoginCalled);

  await browser.close();
} finally {
  srv.kill();
}

console.log("\n" + pass + " 通过 / " + fail + " 失败");
process.exit(fail ? 1 : 0);
