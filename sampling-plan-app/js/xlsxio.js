/*
 * xlsx 读写模块：仅依赖 jszip，纯 JS 生成/解析标准 xlsx（Office Open XML）。
 * 浏览器与 Node 通用。
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.SamplingXlsx = factory(root);
})(typeof self !== "undefined" ? self : this, function (root) {
  "use strict";

  const JSZip = typeof require === "function" ? require("jszip") : root.JSZip;

  // ---------- XML 工具 ----------
  function esc(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function unesc(s) {
    return s
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
      .replace(/&#x([0-9a-fA-F]+);/g, (_, d) => String.fromCodePoint(parseInt(d, 16)))
      .replace(/&amp;/g, "&");
  }

  // ---------- 读取 ----------
  function colToIndex(ref) {
    const mm = /^([A-Z]+)(\d+)$/.exec(ref);
    if (!mm) return null;
    let c = 0;
    for (const ch of mm[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
    return { col: c - 1, row: Number(mm[2]) - 1 };
  }

  function indexToCol(index) {
    let s = "";
    let n = index + 1;
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function parseSharedStrings(xml) {
    return [...xml.matchAll(/<si>(.*?)<\/si>/gs)].map((m) =>
      unesc([...m[1].matchAll(/<t(?: [^>]*)?>(.*?)<\/t>/gs)].map((t) => t[1]).join(""))
    );
  }

  function parseSheetXml(xml, ss) {
    const grid = new Map();
    let maxRow = -1, maxCol = -1;
    for (const m of xml.matchAll(/<c\s+([^>]*?)(?:\/>|>(.*?)<\/c>)/gs)) {
      const attrs = m[1], body = m[2] || "";
      const ref = /r="([A-Z]+\d+)"/.exec(attrs);
      if (!ref) continue;
      const pos = colToIndex(ref[1]);
      if (!pos) continue;
      const t = /t="([^"]+)"/.exec(attrs);
      const type = t ? t[1] : "";
      const v = /<v>(.*?)<\/v>/s.exec(body);
      let val = null;
      if (type === "inlineStr") {
        const is = /<is>(.*?)<\/is>/s.exec(body);
        if (is) val = unesc([...is[1].matchAll(/<t(?: [^>]*)?>(.*?)<\/t>/gs)].map((x) => x[1]).join(""));
      } else if (type === "s" && v) {
        val = ss[Number(v[1])] ?? null;
      } else if (type === "str") {
        val = v ? unesc(v[1]) : "";
      } else if (type === "e") {
        val = v ? unesc(v[1]) : "#ERR";
      } else if (v) {
        const raw = v[1];
        val = raw === "" ? null : Number(raw);
        if (isNaN(val)) val = raw;
      }
      if (val !== null) {
        grid.set(`${pos.row}|${pos.col}`, val);
        if (pos.row > maxRow) maxRow = pos.row;
        if (pos.col > maxCol) maxCol = pos.col;
      }
    }
    return { grid, rows: maxRow + 1, cols: maxCol + 1 };
  }

  async function readWorkbook(blob) {
    const zip = await JSZip.loadAsync(blob);
    const wbXml = await zip.file("xl/workbook.xml").async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    const rels = {};
    for (const m of relsXml.matchAll(/<Relationship\s+([^>]*?)\/>/gs)) {
      const id = /Id="([^"]+)"/.exec(m[1])?.[1];
      const target = /Target="([^"]+)"/.exec(m[1])?.[1];
      if (id && target) rels[id] = "xl/" + target.replace(/^\//, "");
    }
    const ss = zip.file("xl/sharedStrings.xml")
      ? parseSharedStrings(await zip.file("xl/sharedStrings.xml").async("string"))
      : [];

    const sheets = [];
    for (const m of wbXml.matchAll(/<sheet\s+([^>]*?)\/>/gs)) {
      const name = /name="([^"]+)"/.exec(m[1])?.[1];
      const rid = /r:id="([^"]+)"/.exec(m[1])?.[1];
      if (!name || !rels[rid]) continue;
      const xml = await zip.file(rels[rid]).async("string");
      sheets.push({ name, ...parseSheetXml(xml, ss) });
    }
    return sheets;
  }

  function sheetToArray(sheet) {
    const arr = [];
    for (let r = 0; r < sheet.rows; r++) {
      const row = [];
      for (let c = 0; c < sheet.cols; c++) row.push(sheet.grid.get(`${r}|${c}`) ?? "");
      arr.push(row);
    }
    return arr;
  }

  // ---------- 写入 ----------
  function sheetXml(header, rows, colWidths) {
    const dim = { row: 0, col: 0 };
    const data = [header, ...rows];
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const v = data[r][c];
        if (v !== null && v !== undefined && v !== "") {
          if (r > dim.row) dim.row = r;
          if (c > dim.col) dim.col = c;
        }
      }
    }
    const ref = `A1:${indexToCol(dim.col)}${dim.row + 1}`;
    let colsXml = "";
    if (colWidths && colWidths.length) {
      const parts = colWidths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("");
      colsXml = `<cols>${parts}</cols>`;
    }
    let body = "";
    for (let r = 0; r < data.length; r++) {
      let cells = "";
      for (let c = 0; c < data[r].length; c++) {
        const v = data[r][c];
        if (v === null || v === undefined || v === "") continue;
        const refCell = `${indexToCol(c)}${r + 1}`;
        if (typeof v === "number") {
          cells += `<c r="${refCell}"${r === 0 ? ' s="1"' : ""}><v>${v}</v></c>`;
        } else {
          cells += `<c r="${refCell}" t="inlineStr"${r === 0 ? ' s="1"' : ""}><is><t>${esc(v)}</t></is></c>`;
        }
      }
      body += `<row r="${r + 1}">${cells}</row>`;
    }
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="${ref}"/>${colsXml}<sheetData>${body}</sheetData></worksheet>`
    );
  }

  const STYLES_XML =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts>` +
    `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill></fills>` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>` +
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
    `</styleSheet>`;

  function contentTypes(sheetCount) {
    const overrides = [
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ];
    for (let i = 1; i <= sheetCount; i++) {
      overrides.push(
        `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
      );
    }
    overrides.push(
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    );
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      overrides.join("") +
      `</Types>`
    );
  }

  const ROOT_RELS =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
    `</Relationships>`;

  function workbookXml(sheets) {
    const parts = sheets
      .map(
        (s, i) =>
          `<sheet name="${esc(s.name)}" sheetId="${i + 1}"${s.hidden ? ' state="hidden"' : ""} r:id="rId${i + 1}"/>`
      )
      .join("");
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<bookViews><workbookView activeTab="${sheets.length}"/></bookViews><sheets>${parts}</sheets></workbook>`
    );
  }

  function workbookRels(sheetCount) {
    const parts = [];
    for (let i = 1; i <= sheetCount; i++) {
      parts.push(
        `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`
      );
    }
    parts.push(
      `<Relationship Id="rId${sheetCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
    );
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${parts.join("")}</Relationships>`
    );
  }

  const CORE_XML =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:title>采样点布局情况调查</dc:title>` +
    `<dc:creator>采样计划软件</dc:creator>` +
    `</cp:coreProperties>`;

  const APP_XML =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    `<Application>Sampling Plan App</Application>` +
    `</Properties>`;

  /**
   * 生成 xlsx。
   *  mainRows / mainWidths：必填，主表（自动计算区）。
   *  hazardRows / itemRows：可选；提供则追加“危害因素”“检测项目”工作表（保持原表结构）。
   */
  async function writeWorkbook({ hazardRows, mainRows, itemRows, mainWidths }) {
    const sheets = [];
    if (hazardRows && hazardRows.length) sheets.push({ name: "危害因素", rows: hazardRows });
    sheets.push({ name: "测点布局情况调查", rows: mainRows, widths: mainWidths });
    if (itemRows && itemRows.length) sheets.push({ name: "检测项目", rows: itemRows, hidden: true });

    const zip = new JSZip();
    zip.file("[Content_Types].xml", contentTypes(sheets.length));
    zip.file("_rels/.rels", ROOT_RELS);
    zip.file("xl/workbook.xml", workbookXml(sheets));
    zip.file("xl/_rels/workbook.xml.rels", workbookRels(sheets.length));
    zip.file("xl/styles.xml", STYLES_XML);
    zip.file("docProps/core.xml", CORE_XML);
    zip.file("docProps/app.xml", APP_XML);
    sheets.forEach((s, i) => {
      zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s.rows[0], s.rows.slice(1), s.widths || null));
    });
    const out = await zip.generateAsync({
      type: "uint8array",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    return out;
  }

  function downloadBlob(blob, filename) {
    const a = document.createElement("a");
    const url = URL.createObjectURL(blob);
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(url);
      a.remove();
    }, 300);
  }

  return {
    readWorkbook,
    sheetToArray,
    writeWorkbook,
    downloadBlob,
    colToIndex,
    indexToCol,
  };
});
