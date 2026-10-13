// Cloudflare Pages Function：平台接口同源代理（原始 TCP 实现）
// 背景：Cloudflare Workers/Pages Functions 的 fetch() 不允许直接请求 IP 地址
// （错误 1003 Direct IP Access Not Allowed），而平台接口是裸 IP + HTTP
// （http://223.93.144.122:27800）。因此这里用 cloudflare:sockets 发起原始
// TCP 连接，按 HTTP/1.1 协议收发，绕开该限制，解决 HTTPS 页面上
// 验证码 / 登录 / 上传不可用的问题。
import { connect } from "cloudflare:sockets";

const TARGET_HOST = "223.93.144.122";
const TARGET_PORT = 27800;

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, organizationId, belongProject, Content-Type",
  "Access-Control-Max-Age": "86400",
};

function concatBytes(chunks) {
  let total = 0;
  for (const c of chunks) total += c.length;
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.length;
  }
  return out;
}

function indexOfBytes(data, start, byte) {
  for (let i = start; i < data.length; i++) {
    if (data[i] === byte) return i;
  }
  return -1;
}

// 解析 HTTP/1.1 chunked 编码响应体
function dechunk(data) {
  const out = [];
  let i = 0;
  const dec = new TextDecoder();
  while (i < data.length) {
    const nl = indexOfBytes(data, i, 0x0a);
    if (nl < 0) break;
    const sizeLine = dec.decode(data.subarray(i, nl)).trim();
    const size = parseInt(sizeLine.split(";")[0].trim(), 16);
    if (isNaN(size) || size === 0) break;
    out.push(data.subarray(nl + 1, nl + 1 + size));
    i = nl + 1 + size + 2; // 跳过数据块与结尾 CRLF
  }
  return concatBytes(out);
}

// 原始 TCP + HTTP/1.1 请求，返回 { status, headers, body(Uint8Array) }
async function rawHttpRequest(host, port, method, pathAndQuery, headerLines, body, timeoutMs) {
  const socket = connect({ hostname: host, port: port });
  const writer = socket.writable.getWriter();
  const reader = socket.readable.getReader();

  let timer = null;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => {
      socket.close();
      reject(new Error("请求超时"));
    }, timeoutMs || 60000);
  });

  try {
    const head =
      `${method} ${pathAndQuery} HTTP/1.1\r\n` +
      `Host: ${host}:${port}\r\n` +
      `Connection: close\r\n` +
      (body && body.length ? `Content-Length: ${body.length}\r\n` : "") +
      headerLines.join("\r\n") +
      `\r\n` +
      `\r\n`;
    await Promise.race([
      writer.write(new TextEncoder().encode(head)),
      timeout,
    ]);
    if (body && body.length) {
      await Promise.race([writer.write(body), timeout]);
    }

    // 不主动关闭写端：服务器按 Connection: close 返回后自行关闭连接
    const chunks = [];
    while (true) {
      const r = await Promise.race([reader.read(), timeout]);
      if (r.done) break;
      chunks.push(r.value);
    }

    const raw = concatBytes(chunks);
    // 找 \r\n\r\n；找不到时兼容仅 \n\n 的响应
    let headerEnd = -1;
    let crlf = true;
    for (let i = 0; i + 3 < raw.length; i++) {
      if (raw[i] === 13 && raw[i + 1] === 10 && raw[i + 2] === 13 && raw[i + 3] === 10) {
        headerEnd = i;
        break;
      }
    }
    if (headerEnd < 0) {
      for (let i = 0; i + 1 < raw.length; i++) {
        if (raw[i] === 10 && raw[i + 1] === 10) {
          headerEnd = i;
          crlf = false;
          break;
        }
      }
    }
    if (headerEnd < 0) {
      throw new Error("平台响应格式错误（原始响应前200字节：" + bytesPreview(raw) + "）");
    }

    const headText = new TextDecoder().decode(raw.subarray(0, headerEnd));
    const lines = headText.split(crlf ? "\r\n" : "\n");
    const status = parseInt((lines[0] || "").split(" ")[1], 10);
    const headers = {};
    for (let i = 1; i < lines.length; i++) {
      const idx = lines[i].indexOf(":");
      if (idx > 0) {
        headers[lines[i].slice(0, idx).trim().toLowerCase()] = lines[i].slice(idx + 1).trim();
      }
    }
    let bodyBytes = raw.subarray(headerEnd + (crlf ? 4 : 2));
    const te = (headers["transfer-encoding"] || "").toLowerCase();
    if (te.includes("chunked")) {
      bodyBytes = dechunk(bodyBytes);
    } else if (headers["content-length"]) {
      const len = parseInt(headers["content-length"], 10);
      bodyBytes = bodyBytes.subarray(0, len);
    }
    return { status, headers, body: bodyBytes };
  } finally {
    clearTimeout(timer);
    try {
      reader.releaseLock();
    } catch {}
    try {
      socket.close();
    } catch {}
  }
}

function bytesPreview(data) {
  const n = Math.min(data.length, 200);
  let s = "";
  for (let i = 0; i < n; i++) {
    const b = data[i];
    if (b >= 32 && b < 127) s += String.fromCharCode(b);
    else s += "\\x" + b.toString(16).padStart(2, "0");
  }
  return s;
}

export async function onRequest(context) {
  const { request } = context;
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS_HEADERS });
  }

  const url = new URL(request.url);
  const parts = (context.params && context.params.path) || [];
  const path = Array.isArray(parts) ? parts.join("/") : String(parts || "");

  const pathAndQuery = "/" + path + url.search;

  // 透传平台需要的关键请求头（统一为规范大小写）
  const headerMap = [
    ["authorization", "Authorization"],
    ["organizationid", "organizationId"],
    ["belongproject", "belongProject"],
    ["content-type", "Content-Type"],
    ["user-agent", "User-Agent"],
    ["accept", "Accept"],
    ["cookie", "Cookie"],
  ];
  const headerLines = [];
  for (const [src, dst] of headerMap) {
    const v = request.headers.get(src);
    if (v) headerLines.push(`${dst}: ${v}`);
  }

  let body = null;
  if (request.method !== "GET" && request.method !== "HEAD") {
    body = new Uint8Array(await request.arrayBuffer());
  }

  let result;
  try {
    result = await rawHttpRequest(TARGET_HOST, TARGET_PORT, request.method, pathAndQuery, headerLines, body, 60000);
  } catch (e) {
    return new Response(JSON.stringify({ error: String((e && e.message) || e) }), {
      status: 502,
      headers: { "Content-Type": "application/json", ...CORS_HEADERS },
    });
  }

  // 透传平台响应头（跳过 hop-by-hop 头），确保 Set-Cookie / Content-Disposition 等不被丢弃
  const hopByHop = new Set([
    "connection",
    "keep-alive",
    "proxy-authenticate",
    "proxy-authorization",
    "te",
    "trailer",
    "transfer-encoding",
    "upgrade",
  ]);
  let bodyBytes = result.body;
  const respHeaders = {};
  for (const [k, v] of Object.entries(result.headers)) {
    if (hopByHop.has(k)) continue;
    if (k === "content-length" || k === "content-encoding") continue; // 后面按实际内容处理
    respHeaders[k] = v;
  }
  const enc = (result.headers["content-encoding"] || "").toLowerCase();
  if (enc.includes("gzip") && typeof DecompressionStream !== "undefined") {
    try {
      const ds = new DecompressionStream("gzip");
      const stream = new Blob([bodyBytes]).stream().pipeThrough(ds);
      bodyBytes = new Uint8Array(await new Response(stream).arrayBuffer());
    } catch {
      // 解压失败时保留原样并去掉编码声明，避免浏览器解析错误
      respHeaders["content-encoding"] = result.headers["content-encoding"];
    }
  }
  respHeaders["content-type"] = result.headers["content-type"] || "application/octet-stream";
  for (const [k, v] of Object.entries(CORS_HEADERS)) respHeaders[k] = v;
  return new Response(bodyBytes, { status: result.status || 502, headers: respHeaders });
}
