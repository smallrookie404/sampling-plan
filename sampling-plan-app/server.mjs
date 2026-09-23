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
  const recordsDir = path.join(path.dirname(dataFile), "records");

  function readJsonSafe(filePath) {
    try {
      return JSON.parse(fs.readFileSync(filePath, "utf8"));
    } catch {
      return null;
    }
  }

  // 记录索引：[{id,name,createdAt,updatedAt}, ...]
  function readIndex() {
    const idx = readJsonSafe(dataFile);
    if (Array.isArray(idx)) return idx.filter((r) => r && typeof r === "object" && r.id);
    // 旧格式迁移：records.json 是完整记录数组 → 拆分为单条文件 + 索引
    if (Array.isArray(idx) === false && idx) return [];
    return [];
  }

  // 兼容旧格式：若 records.json 含 rows 字段则视为旧结构，一次性拆分
  function migrateLegacy() {
    const raw = readJsonSafe(dataFile);
    if (!Array.isArray(raw)) return;
    const isLegacy = raw.some((r) => r && typeof r === "object" && Array.isArray(r.rows));
    if (!isLegacy) return; // 已是新格式索引
    const index = [];
    fs.mkdirSync(recordsDir, { recursive: true });
    for (const rec of raw) {
      if (!rec || !rec.id) continue;
      writeJson(path.join(recordsDir, encodeURIComponent(rec.id) + ".json"), rec);
      index.push({ id: rec.id, name: rec.name || "", createdAt: rec.createdAt || "", updatedAt: rec.updatedAt || "" });
    }
    writeJson(dataFile, index);
  }

  function readRecord(id) {
    return readJsonSafe(path.join(recordsDir, encodeURIComponent(id) + ".json"));
  }

  function writeRecord(rec) {
    writeJson(path.join(recordsDir, encodeURIComponent(rec.id) + ".json"), rec);
  }

  function deleteRecord(id) {
    const p = path.join(recordsDir, encodeURIComponent(id) + ".json");
    try { fs.unlinkSync(p); } catch {}
  }

  function writeJson(filePath, obj) {
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    const tmp = filePath + ".tmp";
    fs.writeFileSync(tmp, JSON.stringify(obj, null, 2), "utf8");
    fs.renameSync(tmp, filePath);
  }

  function writeRecords(list) {
    writeJson(dataFile, list);
  }

  function readLibrary() {
    try {
      const lib = JSON.parse(fs.readFileSync(path.join(appDir, "data", "library.json"), "utf8"));
      return lib && typeof lib === "object" ? lib : null;
    } catch {
      return null;
    }
  }

  function writeLibrary(lib) {
    writeJson(path.join(appDir, "data", "library.json"), lib);
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
        migrateLegacy();
        const qid = url.searchParams.get("id");
        if (req.method === "GET") {
          // GET /api/records           → 索引（不含 rows，轻量）
          // GET /api/records?id=<id>   → 单条完整记录
          if (qid) {
            const rec = readRecord(qid);
            if (!rec) return sendJson(res, 404, { error: "记录不存在" });
            return sendJson(res, 200, rec);
          }
          return sendJson(res, 200, readIndex());
        }
        if (req.method === "PUT" || req.method === "POST") {
          // PUT /api/records?id=<id>，body 为单条记录 → 写单条文件并更新索引
          if (!qid) return sendJson(res, 400, { error: "缺少记录 id 参数" });
          let body = "";
          for await (const chunk of req) {
            body += chunk;
            if (body.length > 50 * 1024 * 1024) return sendJson(res, 413, { error: "数据过大" });
          }
          try {
            const rec = JSON.parse(body);
            if (!rec || typeof rec !== "object" || rec.id !== qid) {
              return sendJson(res, 400, { error: "数据格式错误（需含与 id 参数一致的 id 字段）" });
            }
            const recs = readRecord(qid);
            if (recs) rec.updatedAt = rec.updatedAt || new Date().toISOString();
            writeRecord(rec);
            const index = readIndex();
            const meta = { id: rec.id, name: rec.name || "", createdAt: rec.createdAt || "", updatedAt: rec.updatedAt || "" };
            const pos = index.findIndex((r) => r.id === qid);
            if (pos >= 0) index[pos] = meta;
            else index.push(meta);
            writeJson(dataFile, index);
            return sendJson(res, 200, { ok: true, id: qid });
          } catch {
            return sendJson(res, 400, { error: "JSON 解析失败" });
          }
        }
        if (req.method === "DELETE") {
          // DELETE /api/records?id=<id> → 删除单条并更新索引
          if (!qid) return sendJson(res, 400, { error: "缺少记录 id 参数" });
          deleteRecord(qid);
          const index = readIndex().filter((r) => r.id !== qid);
          writeJson(dataFile, index);
          return sendJson(res, 200, { ok: true });
        }
        return sendJson(res, 405, { error: "不支持的方法" });
      }

      if (pathname === "/api/library") {
        if (req.method === "GET") {
          const lib = readLibrary();
          return sendJson(res, 200, lib || {});
        }
        if (req.method === "PUT" || req.method === "POST") {
          let body = "";
          for await (const chunk of req) {
            body += chunk;
            if (body.length > 50 * 1024 * 1024) return sendJson(res, 413, { error: "数据过大" });
          }
          try {
            const lib = JSON.parse(body);
            if (!lib || typeof lib !== "object") return sendJson(res, 400, { error: "数据格式错误" });
            writeLibrary(lib);
            return sendJson(res, 200, { ok: true });
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
