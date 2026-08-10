// Cloudflare Pages Function：健康检查（供软件识别“服务模式”）
export function onRequestGet() {
  return new Response(JSON.stringify({ ok: true }), {
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
  });
}
