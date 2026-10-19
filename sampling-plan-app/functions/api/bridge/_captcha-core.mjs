/*
 * 验证码 OCR + 登录核心（纯 ESM，浏览器 / Node / Cloudflare Worker 通用）
 *
 * 从 captcha-ocr/*.py 移植，用于「数据报告回退」在 Cloudflare 上也能零人工运行：
 *   1) decodePng      —— 极简 PNG 解码（本项目验证码恒为 111×36 / 8bit / colorType 2 / 非交错）
 *   2) segment        —— 墨迹列投影切分，强制归 5 段（同 Python segment3.py）
 *   3) feature        —— 裁剪 → 等比缩放 → 居中 24×24 → 高斯模糊 → 8×8 灰度 → 去均值 + L2 归一
 *   4) recognize      —— kNN（模板由 scripts/build-captcha-templates.mjs 用同一套特征生成）
 *   5) rsaEncrypt     —— RSA PKCS#1 v1.5（BigInt 手写，与 js/upload.js / pool.py 一致）
 *   6) loginAccount   —— 取验证码 → OCR（置信度门控换图）→ 登录 → 返回 token
 *
 * ⚠ 特征提取与 Python 版不必逐位相同：只要「模板」与「待识别图」用**同一套** JS 特征，
 *   kNN 就是自洽的（见 tests/captcha.test.mjs 的留一法准确率）。
 */

// ---------------------------------------------------------------- PNG

function readU32(b, i) {
  return ((b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]) >>> 0;
}

function paeth(a, b, c) {
  const p = a + b - c;
  const pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  if (pb <= pc) return b;
  return c;
}

async function inflate(bytes) {
  const ds = new DecompressionStream('deflate');
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/**
 * 解码 PNG → { w, h, px }，px 为 RGB 数组（Uint8Array，长度 w*h*3）
 * 仅支持 8bit、非交错、colorType 2(RGB) / 6(RGBA) / 0(灰度)。
 */
export async function decodePng(bytes) {
  const sig = [137, 80, 78, 71, 13, 10, 26, 10];
  for (let i = 0; i < 8; i++) {
    if (bytes[i] !== sig[i]) throw new Error('不是 PNG');
  }
  let pos = 8;
  let w = 0, h = 0, bitDepth = 8, colorType = 2, interlace = 0;
  const idat = [];
  while (pos + 8 <= bytes.length) {
    const len = readU32(bytes, pos);
    const type = String.fromCharCode(bytes[pos + 4], bytes[pos + 5], bytes[pos + 6], bytes[pos + 7]);
    const data = bytes.subarray(pos + 8, pos + 8 + len);
    if (type === 'IHDR') {
      w = readU32(data, 0); h = readU32(data, 4);
      bitDepth = data[8]; colorType = data[9]; interlace = data[12];
    } else if (type === 'IDAT') {
      idat.push(data);
    } else if (type === 'IEND') {
      break;
    }
    pos += 12 + len;
  }
  if (!w || !h) throw new Error('PNG 缺少 IHDR');
  if (bitDepth !== 8) throw new Error('仅支持 8bit PNG（实际 ' + bitDepth + '）');
  if (interlace) throw new Error('不支持交错 PNG');
  const ch = colorType === 6 ? 4 : colorType === 2 ? 3 : colorType === 0 ? 1 : 0;
  if (!ch) throw new Error('不支持的 PNG colorType ' + colorType);

  let merged;
  if (idat.length === 1) merged = idat[0];
  else {
    let total = 0;
    for (const d of idat) total += d.length;
    merged = new Uint8Array(total);
    let o = 0;
    for (const d of idat) { merged.set(d, o); o += d.length; }
  }
  const raw = await inflate(merged);

  const stride = w * ch;
  const out = new Uint8Array(h * stride);
  let p = 0;
  for (let y = 0; y < h; y++) {
    const ft = raw[p++];
    const line = raw.subarray(p, p + stride);
    p += stride;
    const rowOff = y * stride;
    const prevOff = rowOff - stride;
    for (let x = 0; x < stride; x++) {
      const a = x >= ch ? out[rowOff + x - ch] : 0;
      const b = y > 0 ? out[prevOff + x] : 0;
      const c = (y > 0 && x >= ch) ? out[prevOff + x - ch] : 0;
      let v = line[x];
      if (ft === 1) v += a;
      else if (ft === 2) v += b;
      else if (ft === 3) v += (a + b) >> 1;
      else if (ft === 4) v += paeth(a, b, c);
      out[rowOff + x] = v & 0xff;
    }
  }

  const px = new Uint8Array(w * h * 3);
  for (let i = 0, n = w * h; i < n; i++) {
    if (ch === 3) {
      px[i * 3] = out[i * 3]; px[i * 3 + 1] = out[i * 3 + 1]; px[i * 3 + 2] = out[i * 3 + 2];
    } else if (ch === 4) {
      px[i * 3] = out[i * 4]; px[i * 3 + 1] = out[i * 4 + 1]; px[i * 3 + 2] = out[i * 4 + 2];
    } else {
      const g = out[i];
      px[i * 3] = g; px[i * 3 + 1] = g; px[i * 3 + 2] = g;
    }
  }
  return { w, h, px };
}

// ---------------------------------------------------------------- 分割

export const INK_THRESHOLD = 90;
export const N_CHARS = 5;

export function maskOf(decoded) {
  const { w, h, px } = decoded;
  const m = new Uint8Array(w * h);          // 1 = 墨迹
  for (let i = 0, n = w * h; i < n; i++) {
    const r = px[i * 3], g = px[i * 3 + 1], b = px[i * 3 + 2];
    if ((255 - r) + (255 - g) + (255 - b) > INK_THRESHOLD) m[i] = 1;
  }
  return m;
}

export function colInk(m, w, h) {
  const ci = new Int32Array(w);
  for (let x = 0; x < w; x++) {
    let s = 0;
    for (let y = 0; y < h; y++) s += m[y * w + x];
    ci[x] = s;
  }
  return ci;
}

/** 返回 [[x0,x1], ...] 共 5 段 */
export function segment(m, w, h, n = N_CHARS) {
  const ci = colInk(m, w, h);
  let segs = [];
  let cur = null;
  for (let x = 0; x < w; x++) {
    if (ci[x] > 0) {
      if (cur === null) cur = [x, x]; else cur[1] = x;
    } else if (cur) {
      segs.push(cur); cur = null;
    }
  }
  if (cur) segs.push(cur);
  const inkSum = (s) => { let t = 0; for (let x = s[0]; x <= s[1]; x++) t += ci[x]; return t; };
  segs = segs.filter((s) => inkSum(s) >= 8);

  // 段太少 → 劈最宽的（在墨迹最少处）
  let guard = 0;
  while (segs.length < n && guard < 12) {
    guard++;
    let i = 0;
    for (let k = 1; k < segs.length; k++) {
      if ((segs[k][1] - segs[k][0]) > (segs[i][1] - segs[i][0])) i = k;
    }
    const [x0, x1] = segs[i];
    if (x1 - x0 < 4) break;
    const lo = x0 + Math.max(2, Math.floor((x1 - x0) / 4));
    const hi = x1 - Math.max(2, Math.floor((x1 - x0) / 4));
    if (hi <= lo) break;
    let best = lo;
    for (let x = lo + 1; x <= hi; x++) if (ci[x] < ci[best]) best = x;
    segs = segs.slice(0, i).concat([[x0, best - 1], [best + 1, x1]], segs.slice(i + 1));
  }
  // 段太多 → 合并间隙最小的相邻段
  while (segs.length > n) {
    let k = 0, bestGap = Infinity;
    for (let i = 0; i < segs.length - 1; i++) {
      const gap = segs[i + 1][0] - segs[i][1];
      if (gap < bestGap) { bestGap = gap; k = i; }
    }
    segs = segs.slice(0, k).concat([[segs[k][0], segs[k + 1][1]]], segs.slice(k + 2));
  }
  return segs;
}

// ---------------------------------------------------------------- 特征

export const FEAT = 8;
export const CANVAS = 24;

/** 双线性缩放（src 为 Float32Array，w×h） */
function resizeBilinear(src, sw, sh, dw, dh) {
  const out = new Float32Array(dw * dh);
  if (sw <= 0 || sh <= 0 || dw <= 0 || dh <= 0) return out;
  for (let y = 0; y < dh; y++) {
    const sy = dh === 1 ? 0 : (y * (sh - 1)) / (dh - 1);
    const y0 = Math.floor(sy), y1 = Math.min(sh - 1, y0 + 1), fy = sy - y0;
    for (let x = 0; x < dw; x++) {
      const sx = dw === 1 ? 0 : (x * (sw - 1)) / (dw - 1);
      const x0 = Math.floor(sx), x1 = Math.min(sw - 1, x0 + 1), fx = sx - x0;
      const v00 = src[y0 * sw + x0], v01 = src[y0 * sw + x1];
      const v10 = src[y1 * sw + x0], v11 = src[y1 * sw + x1];
      out[y * dw + x] = (v00 * (1 - fx) + v01 * fx) * (1 - fy) + (v10 * (1 - fx) + v11 * fx) * fy;
    }
  }
  return out;
}

/** 可分离高斯模糊（sigma 用 PIL GaussianBlur 的口径：半径 ≈ 2*sigma） */
function gaussianBlur(src, w, h, sigma) {
  if (sigma <= 0) return src;
  const radius = Math.max(1, Math.round(sigma * 2));
  const k = new Float32Array(radius * 2 + 1);
  let sum = 0;
  for (let i = -radius; i <= radius; i++) {
    const v = Math.exp(-(i * i) / (2 * sigma * sigma));
    k[i + radius] = v; sum += v;
  }
  for (let i = 0; i < k.length; i++) k[i] /= sum;
  const tmp = new Float32Array(w * h);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      let s = 0;
      for (let i = -radius; i <= radius; i++) {
        const xx = Math.min(w - 1, Math.max(0, x + i));
        s += src[y * w + xx] * k[i + radius];
      }
      tmp[y * w + x] = s;
    }
  }
  const out = new Float32Array(w * h);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      let s = 0;
      for (let i = -radius; i <= radius; i++) {
        const yy = Math.min(h - 1, Math.max(0, y + i));
        s += tmp[yy * w + x] * k[i + radius];
      }
      out[y * w + x] = s;
    }
  }
  return out;
}

/** 单字符特征：裁剪 → 等比缩放 → 居中 24×24 → 高斯模糊 → 8×8 → 去均值 + L2 归一 */
export function feature(m, w, h, x0, x1, y0, y1) {
  const cw = x1 - x0 + 1, chh = y1 - y0 + 1;
  const crop = new Float32Array(cw * chh);
  for (let y = 0; y < chh; y++) {
    for (let x = 0; x < cw; x++) {
      crop[y * cw + x] = m[(y0 + y) * w + (x0 + x)] ? 255 : 0;
    }
  }
  const scale = (CANVAS - 4) / Math.max(cw, chh);
  const nw = Math.max(1, Math.round(cw * scale));
  const nh = Math.max(1, Math.round(chh * scale));
  const small = resizeBilinear(crop, cw, chh, nw, nh);

  const canvas = new Float32Array(CANVAS * CANVAS);
  const ox = Math.floor((CANVAS - nw) / 2), oy = Math.floor((CANVAS - nh) / 2);
  for (let y = 0; y < nh; y++) {
    for (let x = 0; x < nw; x++) canvas[(oy + y) * CANVAS + (ox + x)] = small[y * nw + x];
  }
  const blurred = gaussianBlur(canvas, CANVAS, CANVAS, 1.0);
  const feat = resizeBilinear(blurred, CANVAS, CANVAS, FEAT, FEAT);

  const v = new Float64Array(FEAT * FEAT);
  let mean = 0;
  for (let i = 0; i < v.length; i++) { v[i] = feat[i] / 255; mean += v[i]; }
  mean /= v.length;
  let norm = 0;
  for (let i = 0; i < v.length; i++) { v[i] -= mean; norm += v[i] * v[i]; }
  norm = Math.sqrt(norm) || 1;
  const out = new Array(v.length);
  for (let i = 0; i < v.length; i++) out[i] = v[i] / norm;
  return out;
}

/** 提取整图 5 个字符的特征（失败的位置为 null） */
export function extractFeatures(decoded) {
  const { w, h } = decoded;
  const m = maskOf(decoded);
  const segs = segment(m, w, h);
  return segs.map(([x0, x1]) => {
    let y0 = -1, y1 = -1;
    for (let y = 0; y < h; y++) {
      let has = false;
      for (let x = x0; x <= x1; x++) if (m[y * w + x]) { has = true; break; }
      if (has) { if (y0 < 0) y0 = y; y1 = y; }
    }
    if (y0 < 0) return null;
    return feature(m, w, h, x0, x1, y0, y1);
  });
}

export function dist2(a, b) {
  let s = 0;
  for (let i = 0; i < a.length; i++) { const d = a[i] - b[i]; s += d * d; }
  return s;
}

// ---------------------------------------------------------------- 识别

const OPS = { '+': (a, b) => a + b, '-': (a, b) => a - b, '×': (a, b) => a * b };

/** templates: { digits: [[label, feat], ...], ops: [[label, feat], ...] } */
export function recognize(decoded, templates) {
  const feats = extractFeatures(decoded);
  if (!feats || feats.length < 3 || feats.slice(0, 3).some((f) => !f)) return null;
  const picks = [];
  const confs = [];
  for (const pos of [0, 1, 2]) {
    const pool = pos === 1 ? templates.ops : templates.digits;
    const fv = feats[pos];
    const cand = pool
      .map((t) => ({ label: t[0], d: dist2(fv, t[1]) }))
      .sort((a, b) => a.d - b.d);
    const best = cand[0];
    const second = cand.find((c) => c.label !== best.label);
    if (!second) confs.push(1);
    else confs.push(second.d > 0 ? Math.max(0, (second.d - best.d) / second.d) : 1);
    picks.push(best.label);
  }
  const [d1, op, d2] = picks;
  const answer = OPS[op] ? OPS[op](Number(d1), Number(d2)) : null;
  return { d1, op, d2, expr: d1 + op + d2, answer, conf: Math.min.apply(null, confs) };
}

// ---------------------------------------------------------------- RSA PKCS#1 v1.5

export const RSA_PUB_B64 = 'MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBANL378k3RiZHWx5AfJqdH9xRNBmD9wGD2iRe41HdTNF8RUhNnHit5NpMNtGL0NPTSSpPjjI1kJfVorRvaQerUgkCAwEAAQ==';

export function b64ToBytes(b64) {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export function bytesToB64(bytes) {
  let bin = '';
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}

function bytesToBigInt(bytes) {
  let hex = '';
  for (let i = 0; i < bytes.length; i++) hex += bytes[i].toString(16).padStart(2, '0');
  return BigInt('0x' + (hex || '0'));
}

function bigIntToBytes(v, len) {
  let hex = v.toString(16);
  if (hex.length % 2) hex = '0' + hex;
  const out = new Uint8Array(len);
  let off = len - hex.length / 2;
  for (let i = 0; i < hex.length; i += 2) out[off++] = parseInt(hex.substr(i, 2), 16);
  return out;
}

function modPow(base, exp, mod) {
  let result = 1n;
  base %= mod;
  while (exp > 0n) {
    if (exp & 1n) result = (result * base) % mod;
    base = (base * base) % mod;
    exp >>= 1n;
  }
  return result;
}

let _key = null;
function publicKey() {
  if (_key) return _key;
  const der = b64ToBytes(RSA_PUB_B64);
  // 与 pool.py 相同的固定偏移：64 字节模数 + e=65537
  _key = { n: bytesToBigInt(der.subarray(25, 89)), e: bytesToBigInt(der.subarray(91, 94)), k: 64 };
  return _key;
}

export function rsaEncrypt(plain) {
  const { n, e, k } = publicKey();
  const data = new TextEncoder().encode(plain);
  const padLen = k - data.length - 3;
  if (padLen < 8) throw new Error('密码过长');
  const pad = new Uint8Array(padLen);
  crypto.getRandomValues(pad);
  for (let i = 0; i < padLen; i++) if (pad[i] === 0) pad[i] = 1;   // 非零填充
  const em = new Uint8Array(k);
  em[0] = 0; em[1] = 2;
  em.set(pad, 2);
  em[2 + padLen] = 0;
  em.set(data, 3 + padLen);
  const c = modPow(bytesToBigInt(em), e, n);
  return bytesToB64(bigIntToBytes(c, k));
}

// ---------------------------------------------------------------- 登录

/**
 * 只做「取验证码 → OCR（置信度门控换图）」，**不登录、不涉及任何账号口令**。
 * 供登录页「验证码自动填写」使用：后端给出 uuid + 识别结果，前端仍用用户手输的
 * 账号密码自己调 /auth/login（密码不离开浏览器）。
 * @param {object} opts
 *   apiBase    平台根地址（如 http://223.93.144.122:27800 或同源代理 /api/platform）
 *   fetchImpl  自定义 fetch（Worker 里走 socket 转发）
 *   templates  字符模板（不传则懒加载内置模板）
 *   confThreshold / maxTries / onLog / onCaptcha
 * @returns {object|null} { uuid, answer, expr, conf, img, attempt } 或 null
 */
export async function solveCaptcha(opts) {
  const {
    apiBase, fetchImpl = fetch, templates,
    confThreshold = 0.25, maxTries = 4, onLog = () => {}, onCaptcha = null,
  } = opts;
  const tpl = templates || (await import('./_captcha-templates.mjs')).CAPTCHA_TEMPLATES;

  for (let attempt = 1; attempt <= maxTries; attempt++) {
    let code;
    try {
      const r = await fetchImpl(apiBase + '/auth/code', { method: 'GET' });
      code = await r.json();
    } catch (e) {
      onLog('取验证码失败：' + (e && e.message ? e.message : e));
      continue;
    }
    if (!code || !code.img || !code.uuid) {
      onLog('取验证码返回异常');
      continue;
    }
    if (onCaptcha) { try { await onCaptcha(); } catch (e) {} }
    const b64 = String(code.img).split(',').pop();
    let rec;
    try {
      rec = recognize(await decodePng(b64ToBytes(b64)), tpl);
    } catch (e) {
      onLog('验证码解码失败：' + (e && e.message ? e.message : e));
      continue;
    }
    if (!rec || rec.answer === null || rec.conf < confThreshold) {
      onLog('第 ' + attempt + ' 张置信度 ' + (rec ? rec.conf.toFixed(2) : 'n/a') + ' 低于门限，换图');
      continue;
    }
    return {
      uuid: code.uuid,
      answer: rec.answer,
      expr: rec.expr,
      conf: rec.conf,
      img: code.img,          // 原样回传（data:image/png;base64,…），前端仍可显示供用户核对
      attempt: attempt,
    };
  }
  return null;
}

/**
 * 取验证码 → OCR（置信度门控换图）→ 登录。
 * @param {object} opts
 *   apiBase    平台根地址（如 http://223.93.144.122:27800 或同源代理 /api/platform）
 *   username/password
 *   fetchImpl  自定义 fetch（Worker 里走 socket 转发）
 *   confThreshold / maxTries / onLog
 * @returns {object|null} { token, via, conf, user } 或 null
 */
export async function loginAccount(opts) {
  const {
    apiBase, username, password, fetchImpl = fetch,
    templates, confThreshold = 0.25, maxTries = 4, onLog = () => {}, onCaptcha = null,
  } = opts;
  const tpl = templates || (await import('./_captcha-templates.mjs')).CAPTCHA_TEMPLATES;

  for (let attempt = 1; attempt <= maxTries; attempt++) {
    // 每轮只取 1 张验证码：被平台拒（OCR 猜错）就换新图重试，故 maxTries 传 1
    const sol = await solveCaptcha({
      apiBase, fetchImpl, templates: tpl, confThreshold, maxTries: 1, onLog, onCaptcha,
    });
    if (!sol) continue;
    const rec = { answer: sol.answer, expr: sol.expr, conf: sol.conf };
    let d = null;
    try {
      const r = await fetchImpl(apiBase + '/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          username,
          password: rsaEncrypt(password),
          code: rec.answer,
          uuid: sol.uuid,
        }),
      });
      d = await r.json();
    } catch (e) {
      onLog('登录请求失败：' + (e && e.message ? e.message : e));
      continue;
    }
    if (d && d.token) {
      const u = d.user && typeof d.user === 'object' ? d.user : {};
      // 平台登录响应的 user 可能是 {user:{…}}，也可能直接就是用户对象，两种都兼容
      const uo = (u.user && typeof u.user === 'object') ? u.user : u;
      const orgId = (uo && uo.organizationId)
        || (uo && uo.organization && uo.organization.organizationId)
        || (u && u.organizationId) || null;
      return {
        token: d.token,
        via: rec.expr,
        conf: Math.round(rec.conf * 100) / 100,
        user: u.userName || (uo && uo.userName) || null,
        // 只回传最小画像：完整用户对象含身份证号/手机号等，没必要经桥再传一遍；
        // 浏览器拿到 token 后可自己调 /auth/info 取权威信息。
        profile: {
          userName: (uo && uo.userName) || u.userName || null,
          userCode: (uo && uo.userCode) || u.userCode || null,
          organizationId: orgId,
        },
        orgId: orgId,
      };
    }
    onLog('第 ' + attempt + ' 张被拒：' + String((d && (d.message || d.msg)) || '').slice(0, 80));
  }
  return null;
}
