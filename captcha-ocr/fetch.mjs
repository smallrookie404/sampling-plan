// 采集平台登录验证码样本（只读 GET /auth/code，不涉及登录）
// 用法: node fetch.mjs [数量] [起始序号]
import fs from 'node:fs';
import path from 'node:path';

const API = 'http://223.93.144.122:27800';
const N = Number(process.argv[2] || 60);
const START = Number(process.argv[3] || 0);
const OUT = path.join(import.meta.dirname, 'samples');

fs.mkdirSync(OUT, { recursive: true });

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let ok = 0, fail = 0;
for (let i = 0; i < N; i++) {
  const id = String(START + i).padStart(4, '0');
  try {
    const r = await fetch(API + '/auth/code', { signal: AbortSignal.timeout(15000) });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const d = await r.json();
    if (!d || !d.img || !d.uuid) throw new Error('返回缺少 img/uuid');
    const b64 = String(d.img).replace(/^data:image\/\w+;base64,/, '');
    fs.writeFileSync(path.join(OUT, id + '.png'), Buffer.from(b64, 'base64'));
    fs.writeFileSync(path.join(OUT, id + '.uuid'), d.uuid);
    ok++;
    process.stdout.write('.');
  } catch (e) {
    fail++;
    process.stdout.write('x');
    if (fail <= 3) console.error('\n失败: ' + e.message);
  }
  await sleep(120);
}
console.log('\n完成: 成功 ' + ok + ' 张，失败 ' + fail + ' 张 → ' + OUT);
