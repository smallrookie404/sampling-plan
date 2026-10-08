import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { startServer } from "../server.mjs";

const TEST_DATA = path.resolve("data/test-server.json");
const server = await startServer({ port: 0, dataFile: TEST_DATA });
const base = `http://127.0.0.1:${server.address().port}`;

try {
  // 1) 写入 / 读取记录（现行契约：PUT 单条 ?id=，索引+单条文件）
  const put = await fetch(base + "/api/records?id=1", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ id: "1", name: "服务测试记录", rows: [] }),
  });
  if (!put.ok) throw new Error("PUT 失败: " + put.status);
  const got = await (await fetch(base + "/api/records")).json();
  if (!Array.isArray(got) || got.length !== 1 || got[0].name !== "服务测试记录") {
    throw new Error("GET 结果不符");
  }

  // 2) 单条完整记录落盘到 data/records/<id>.json
  const recFile = JSON.parse(fs.readFileSync(path.resolve("data/records/1.json"), "utf8"));
  if (recFile.id !== "1" || recFile.name !== "服务测试记录") throw new Error("单条记录未写入数据文件");

  // 3) 静态首页可访问
  const idx = await (await fetch(base + "/")).text();
  if (!idx.includes("系统测点布局调查") || !idx.includes("js/app.js")) throw new Error("首页未正常提供");

  // 4) 路径穿越拦截（原始 HTTP 请求，不经 URL 归一化）
  function rawGet(pathname) {
    return new Promise((resolve, reject) => {
      const req = http.request(
        { host: "127.0.0.1", port: server.address().port, path: pathname, method: "GET" },
        (res) => {
          res.resume();
          res.on("end", () => resolve(res.statusCode));
        }
      );
      req.on("error", reject);
      req.end();
    });
  }
  const code = await rawGet("/..%2f..%2f..%2f..%2fWindows%2fwin.ini");
  if (code !== 403 && code !== 404) throw new Error("路径穿越未被拦截: " + code);

  // 5) 并发写不互相覆盖：多个 PUT 同时到达，全部落盘且索引无重复
  const N = 10;
  const putAll = [...Array(N)].map((_, i) =>
    fetch(base + "/api/records?id=c" + i, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id: "c" + i, name: "并发" + i, rows: [] }),
    })
  );
  const rs = await Promise.all(putAll);
  if (!rs.every((x) => x.ok)) throw new Error("并发 PUT 有失败");
  const idxAfter = await (await fetch(base + "/api/records")).json();
  const ids = idxAfter.map((r) => r.id);
  if (new Set(ids).size !== ids.length) throw new Error("索引出现重复条目");
  if (ids.filter((x) => String(x).startsWith("c")).length !== N) throw new Error("并发写入有丢失");

  console.log("server 读写 / 单条落盘 / 静态服务 / 防穿越 / 并发写 校验通过 ✔");
} finally {
  try {
    server.closeAllConnections();
    await server.close();
  } catch {}
  fs.rmSync(TEST_DATA, { force: true });
  fs.rmSync(path.resolve("data/records/1.json"), { force: true });
  fs.rmSync(path.resolve("data/records"), { recursive: true, force: true });
}
