/*
 * 采样计划软件 - 本地服务
 * 职责：
 *  1. 以 http://127.0.0.1:8017 提供软件页面与静态资源；
 *  2. 提供 /api/records 读写接口，数据保存到本文件同目录下的 data/records.json。
 * 纯 Node.js 标准库实现，无需安装任何依赖。
 */
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_PORT = 8017;
const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".md": "text/plain; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
};

export function createServer(opts = {}) {
  const appDir = opts.dir || __dirname;
  const dataFile = opts.dataFile || path.join(appDir, "data", "records.json");

  function readRecords() {
    try {
      const list = JSON.parse(fs.readFileSync(dataFile, "utf8"));
      return Array.isArray(list) ? list : [];
    } catch {
      return [];
    }
  }

  function writeRecords(list) {
    fs.mkdirSync(path.dirname(dataFile), { recursive: true });
    const tmp = dataFile + ".tmp";
    fs.writeFileSync(tmp, JSON.stringify(list, null, 2), "utf8");
    fs.renameSync(tmp, dataFile);
  }

  function sendJson(res, code, obj) {
    res.writeHead(code, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
    res.end(JSON.stringify(obj));
  }

  const server = http.createServer(async (req, res) => {
    try {
      const url = new URL(req.url, "http://localhost");
      const pathname = decodeURIComponent(url.pathname);

      if (pathname === "/api/health") return sendJson(res, 200, { ok: true });

      if (pathname === "/api/records") {
        if (req.method === "GET") return sendJson(res, 200, readRecords());
        if (req.method === "PUT" || req.method === "POST") {
          let body = "";
          for await (const chunk of req) {
            body += chunk;
            if (body.length > 50 * 1024 * 1024) return sendJson(res, 413, { error: "数据过大" });
          }
          try {
            const list = JSON.parse(body);
            if (!Array.isArray(list)) return sendJson(res, 400, { error: "数据格式错误" });
            writeRecords(list);
            return sendJson(res, 200, { ok: true, count: list.length });
          } catch {
            return sendJson(res, 400, { error: "JSON 解析失败" });
          }
        }
        return sendJson(res, 405, { error: "不支持的方法" });
      }

      // 静态资源（仅限软件目录内，防目录穿越）
      const rel = pathname === "/" ? "index.html" : pathname.replace(/^\/+/, "");
      const filePath = path.resolve(appDir, rel);
      const rootOk = filePath === path.join(appDir, "index.html") || filePath.startsWith(appDir + path.sep);
      if (!rootOk) return sendJson(res, 403, { error: "禁止访问" });

      fs.stat(filePath, (err, st) => {
        if (err || !st.isFile()) {
          res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
          res.end("404 Not Found");
          return;
        }
        const ext = path.extname(filePath).toLowerCase();
        res.writeHead(200, {
          "Content-Type": MIME[ext] || "application/octet-stream",
          "Cache-Control": "no-store",
        });
        fs.createReadStream(filePath).pipe(res);
      });
    } catch (e) {
      sendJson(res, 500, { error: String(e && e.message ? e.message : e) });
    }
  });

  return server;
}

export function startServer(opts = {}) {
  return new Promise((resolve, reject) => {
    const server = createServer(opts);
    server.once("error", reject);
    const port = opts.port === undefined ? DEFAULT_PORT : opts.port;
    server.listen(port, "127.0.0.1", () => resolve(server));
  });
}

// 直接运行时启动服务（被“启动采样计划软件.bat”调用）
const isMain =
  process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMain) {
  startServer({ port: Number(process.env.PORT || DEFAULT_PORT) })
    .then((server) => {
      const { port } = server.address();
      console.log(`采样计划软件已启动：http://127.0.0.1:${port}`);
      console.log(`数据文件：${path.join(__dirname, "data", "records.json")}`);
    })
    .catch((e) => {
      if (e && e.code === "EADDRINUSE") process.exit(0); // 已有实例在运行，直接退出
      console.error(e.message || e);
      process.exit(1);
    });
}
