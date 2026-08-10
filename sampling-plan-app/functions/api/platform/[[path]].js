// Cloudflare Pages Function：平台接口同源代理
// HTTPS 页面（GitHub Pages / Cloudflare Pages）被浏览器禁止调用 http:// 平台接口（混合内容），
// 通过本代理在 Cloudflare Pages 同源转发，解决验证码/登录/上传不可用的问题。
const TARGET = "http://223.93.144.122:27800";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, organizationId, belongProject, Content-Type",
  "Access-Control-Max-Age": "86400",
};

export async function onRequest(context) {
  const { request } = context;
  const url = new URL(request.url);
  const parts = (context.params && context.params.path) || [];
  const path = Array.isArray(parts) ? parts.join("/") : String(parts || "");
  const targetUrl = TARGET + "/" + path + url.search;

  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS_HEADERS });
  }

  const headers = new Headers(request.headers);
  headers.delete("host");
  let body = undefined;
  if (request.method !== "GET" && request.method !== "HEAD") {
    body = await request.arrayBuffer();
  }

  const resp = await fetch(targetUrl, {
    method: request.method,
    headers: headers,
    body: body,
    redirect: "manual",
  });

  const out = new Response(resp.body, resp);
  for (const [k, v] of Object.entries(CORS_HEADERS)) out.headers.set(k, v);
  return out;
}
