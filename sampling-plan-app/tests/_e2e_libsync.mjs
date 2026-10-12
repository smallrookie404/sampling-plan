// 验证：① 打开页面不再写参考库（原来 server 模式下每次开页都会 PUT /api/library）
//       ② 真正改动参考库后仍会写一次
import { chromium } from 'playwright-core';
import fs from 'node:fs';
const LIB = 'data/library.json';
const st = () => { const s = fs.statSync(LIB); return { mtime: s.mtimeMs, size: s.size }; };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const b = await chromium.launch({ executablePath: 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', headless: true });
const p = await b.newPage();
const errs = []; p.on('pageerror', (e) => errs.push(String(e)));
p.on('dialog', (d) => d.accept());   // 自动确认「恢复默认库」
const libPuts = [];
const libGets = [];
p.on('request', (r) => { if (r.url().includes('/api/library')) { if (r.method() === 'PUT') libPuts.push(Date.now()); else libGets.push(Date.now()); } });

await sleep(300);
const before = st();
console.log('=== 打开页面之前 ===', JSON.stringify(before));

await p.goto('http://127.0.0.1:8019/index.html', { waitUntil: 'load' });
await p.waitForTimeout(4000);          // 等 loadLibrary 跑完
await p.evaluate(() => { const e = document.getElementById('xcdc-login'); if (e) e.style.display = 'none'; });
const afterOpen = st();
console.log('=== 打开页面 4s 后 ===', JSON.stringify(afterOpen));
console.log('  ① 开页未写参考库：', afterOpen.mtime === before.mtime ? '✅ 通过' : '❌ 仍被写入');
console.log('  开页期间 /api/library GET 次数：', libGets.length, libGets.length >= 1 ? '✅ 说明 loadLibrary 确实跑了' : '⚠ 未触发');
console.log('  开页期间 /api/library PUT 次数：', libPuts.length, libPuts.length === 0 ? '✅' : '❌');

// ② 真正改一次参考库
await p.click('.tab[data-tab="hazard"]');
await p.waitForTimeout(600);
libPuts.length = 0;
const t0 = Date.now();
await p.click('#hazard-restore');
await p.waitForTimeout(3500);          // 防抖 1.2s + 写入
const afterEdit = st();
console.log('\n=== 点「恢复默认库」之后 ===');
console.log('  ② 改动后写入了参考库：', afterEdit.mtime > afterOpen.mtime ? '✅ 通过' : '❌ 未写入');
console.log('  /api/library PUT 次数：', libPuts.length, '（应为 1）', libPuts.length === 1 ? '✅' : '❌');
console.log('  耗时', ((Date.now() - t0) / 1000).toFixed(2) + 's');
console.log('\n页面错误：', errs.length ? errs.join(' | ') : '(无)');
await b.close();
