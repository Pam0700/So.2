"use strict";

const APP_VERSION = "7.0.0";

const NS = {
  main: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  officeRel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  a16: "http://schemas.microsoft.com/office/drawing/2014/main",
  xml: "http://www.w3.org/XML/1998/namespace",
};

const state = {
  dgRows: [],
  productRows: [],
  groups: [],
  pages: [],
  warnings: [],
  outputBlob: null,
  outputName: "",
  templateFile: null,
  templateLayout: null,
};

const aliases = {
  code: ["code", "ma code", "ma", "ma don", "ma don hang", "order code", "product code", "item code", "sku"],
  customer: ["ten khach hang", "ten khach", "khach hang", "customer", "customer name", "ten nguoi nhan"],
  deliveryPoint: ["diem giao hang", "diem giao", "dia chi giao hang", "dia chi giao", "dia chi", "delivery point", "delivery address", "address"],
  ship: ["ship", "ma ship", "so ship", "shipment", "shipment code", "shipping code", "ma van don", "van don"],
  product: ["san pham", "ten san pham", "product", "product name", "item", "item name", "ten hang"],
};

const demoDgRows = [
  { code: "2450D171440", customer: "Công Ty Cổ Phần Vifon - Chi Nhánh Hải Dương", deliveryPoint: "Lô CN8, KCN Tân Trường, Hải Dương", ship: "24SPF47158" },
  { code: "2450D172050", customer: "Công ty TNHH BBQ Việt Nam", deliveryPoint: "BBQ Kon Tum - KHO NGUYÊN KHÊ", ship: "24SPF47158" },
  { code: "2450D172272", customer: "Công Ty TNHH Thương Mại Và Dịch Vụ The City Việt Nam", deliveryPoint: "Phú Vinh Tây, Hoằng Hóa, Thanh Hóa", ship: "24SPF47158" },
  { code: "2450D172275", customer: "Công Ty TNHH Thương Mại Và Dịch Vụ The City Việt Nam", deliveryPoint: "Chân cầu Đe, Minh Lộc, Hậu Lộc, Thanh Hóa", ship: "24SPF47158" },
];

const demoProductRows = [
  { code: "2450D171440", product: "Ức Phi Lê (5KG) VietGap", sourceFile: "50_demo.xlsx" },
  { code: "2450D171440", product: "Ức Phi Lê 500g VietGap", sourceFile: "50_demo.xlsx" },
  { code: "2450D172050", product: "Gà 9 miếng BBQ đông lạnh", sourceFile: "50_demo.xlsx" },
  { code: "2450D172050", product: "Ức trong BBQ đông lạnh (0.5KG)", sourceFile: "50_demo.xlsx" },
  { code: "2450D172272", product: "Đùi gà góc tư (5KG) VietGap", sourceFile: "54_demo.xlsx" },
  { code: "2450D172272", product: "Ức Phi Lê (2KG) VietGap", sourceFile: "54_demo.xlsx" },
  { code: "2450D172272", product: "Tim gà (500 gr) VietGap", sourceFile: "54_demo.xlsx" },
  { code: "2450D172275", product: "Cánh gà nguyên 5KG VietGap", sourceFile: "58_demo.xlsx" },
  { code: "2450D172275", product: "Cánh gà 1KG VietGap", sourceFile: "58_demo.xlsx" },
  { code: "2450D172275", product: "Ức gà TF VietGap", sourceFile: "58_demo.xlsx" },
  { code: "2450D172275", product: "Ức gà 500g VietGap", sourceFile: "58_demo.xlsx" },
];

const el = Object.fromEntries([
  "dgFiles", "productFiles", "templateFile", "dgStatus", "productStatus", "templateStatus",
  "shipSelect", "processAll", "removeDuplicates", "removeVietgap", "loadDemoButton", "processButton",
  "shareButton", "downloadButton", "message", "resultSection", "resultBody", "warningList",
  "shipMetric", "dgMetric", "productMetric", "groupMetric", "pageMetric", "warningMetric", "layoutInfo",
].map((id) => [id, document.getElementById(id)]));

function stripAccents(value) {
  return String(value ?? "")
    .replace(/Đ/g, "D")
    .replace(/đ/g, "d")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function normalize(value) {
  return stripAccents(value)
    .toLowerCase()
    .replace(/[\r\n\t]+/g, " ")
    .replace(/[_\-./\\:]+/g, " ")
    .replace(/[^a-z0-9& ]+/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function clean(value) {
  return String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function key(value) {
  return clean(value).toUpperCase();
}

function setMessage(text, type = "") {
  el.message.textContent = text;
  el.message.className = `message ${type}`.trim();
}

function fieldForHeader(value, allowedFields) {
  const n = normalize(value);
  if (!n) return null;
  for (const field of allowedFields) {
    if (aliases[field].includes(n)) return field;
  }
  if (allowedFields.includes("customer") && n.includes("ten khach")) return "customer";
  if (allowedFields.includes("deliveryPoint") && (n.includes("diem giao") || n.includes("dia chi"))) return "deliveryPoint";
  if (allowedFields.includes("product") && n.includes("san pham")) return "product";
  if (allowedFields.includes("ship") && n.includes("ship")) return "ship";
  if (allowedFields.includes("code") && (n === "code" || n.includes("ma code"))) return "code";
  return null;
}

function findHeaderRow(matrix, requiredFields) {
  const maxRows = Math.min(matrix.length, 80);
  for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
    const row = matrix[rowIndex] || [];
    const mapping = {};
    row.forEach((cell, colIndex) => {
      const field = fieldForHeader(cell, requiredFields);
      if (field && mapping[field] === undefined) mapping[field] = colIndex;
    });
    if (requiredFields.every((field) => mapping[field] !== undefined)) return { rowIndex, mapping };
  }
  return null;
}

function readMatrices(workbook) {
  return workbook.SheetNames.map((sheetName) => ({
    sheetName,
    matrix: XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: "",
      raw: false,
      blankrows: false,
    }),
  }));
}

async function parseDgFiles(files) {
  const output = [];
  for (const file of files) {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", raw: false, cellDates: false });
    let parsedAny = false;
    for (const { sheetName, matrix } of readMatrices(workbook)) {
      let header = findHeaderRow(matrix, ["code", "customer", "deliveryPoint", "ship"]);
      if (!header && matrix.some((row) => (row || []).filter((v) => clean(v)).length >= 4)) {
        header = { rowIndex: -1, mapping: { code: 0, customer: 1, deliveryPoint: 2, ship: 3 } };
      }
      if (!header) continue;
      const start = header.rowIndex + 1;
      for (let i = start; i < matrix.length; i += 1) {
        const row = matrix[i] || [];
        const record = {
          code: clean(row[header.mapping.code]),
          customer: clean(row[header.mapping.customer]),
          deliveryPoint: clean(row[header.mapping.deliveryPoint]),
          ship: clean(row[header.mapping.ship]),
          sourceFile: file.name,
          sourceSheet: sheetName,
          sourceRow: i + 1,
        };
        if (!record.code || (!record.customer && !record.deliveryPoint && !record.ship)) continue;
        output.push(record);
        parsedAny = true;
      }
      if (parsedAny) break;
    }
    if (!parsedAny) throw new Error(`${file.name}: chưa nhận diện được 4 cột Code, Tên khách hàng, Điểm giao hàng, Ship.`);
  }
  return output;
}

async function parseProductFiles(files) {
  const output = [];
  for (const file of files) {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", raw: false, cellDates: false });
    let parsedAny = false;
    for (const { sheetName, matrix } of readMatrices(workbook)) {
      let header = findHeaderRow(matrix, ["code", "product"]);
      if (!header && matrix.some((row) => (row || []).filter((v) => clean(v)).length >= 2)) {
        header = { rowIndex: -1, mapping: { code: 0, product: 1 } };
      }
      if (!header) continue;
      const start = header.rowIndex + 1;
      for (let i = start; i < matrix.length; i += 1) {
        const row = matrix[i] || [];
        const code = clean(row[header.mapping.code]);
        const product = clean(row[header.mapping.product]);
        if (!code || !product) continue;
        output.push({ code, product, sourceFile: file.name, sourceSheet: sheetName, sourceRow: i + 1 });
        parsedAny = true;
      }
      if (parsedAny) break;
    }
    if (!parsedAny) throw new Error(`${file.name}: chưa nhận diện được 2 cột Code và Sản phẩm.`);
  }
  return output;
}

function updateShipSelect() {
  const ships = [...new Set(state.dgRows.map((row) => clean(row.ship)).filter(Boolean))].sort((a, b) => a.localeCompare(b, "vi"));
  el.shipSelect.innerHTML = "";
  if (!ships.length) {
    el.shipSelect.add(new Option("Không có mã Ship", ""));
    el.shipSelect.disabled = true;
  } else {
    ships.forEach((ship) => el.shipSelect.add(new Option(ship, ship)));
    el.shipSelect.disabled = false;
  }
  el.shipMetric.textContent = String(ships.length);
}

function updateMetrics() {
  el.dgMetric.textContent = String(state.dgRows.length);
  el.productMetric.textContent = String(state.productRows.length);
  el.groupMetric.textContent = String(state.groups.length);
  el.pageMetric.textContent = String(state.pages.length);
  el.warningMetric.textContent = String(state.warnings.length);
  el.processButton.disabled = !(state.dgRows.length && state.productRows.length);
}

function productSortRank(product) {
  const n = normalize(product).replace(/\s+/g, "");
  if (/5kg/.test(n) && !/0[.,]?5kg/.test(n)) return 0;
  if (/2kg/.test(n)) return 1;
  if (/1kg/.test(n) && !/0[.,]?1kg/.test(n)) return 2;
  if (/(^|[^a-z])tf([^a-z]|$)/.test(normalize(product))) return 3;
  if (/500g|500gr|0[.,]?5kg/.test(n)) return 4;
  return 5;
}

function sortProducts(products) {
  return [...products].sort((a, b) => {
    const r = productSortRank(a) - productSortRank(b);
    return r || a.localeCompare(b, "vi", { sensitivity: "base", numeric: true });
  });
}

function removeVietgapText(product) {
  return clean(String(product).replace(/viet\s*gap/gi, ""));
}

function buildGroups() {
  state.warnings = [];
  const selectedShip = clean(el.shipSelect.value);
  const dgRows = state.dgRows.filter((row) => el.processAll.checked || key(row.ship) === key(selectedShip));
  if (!dgRows.length) throw new Error("Không có dòng DG phù hợp với mã Ship đã chọn.");

  const productMap = new Map();
  for (const row of state.productRows) {
    const k = key(row.code);
    if (!productMap.has(k)) productMap.set(k, []);
    productMap.get(k).push(row.product);
  }

  const groupMap = new Map();
  for (const dg of dgRows) {
    const gk = `${key(dg.customer)}\u0001${key(dg.deliveryPoint)}`;
    if (!groupMap.has(gk)) {
      groupMap.set(gk, { customer: dg.customer, deliveryPoint: dg.deliveryPoint, products: [], codes: [] });
    }
    const group = groupMap.get(gk);
    group.codes.push(dg.code);
    const products = productMap.get(key(dg.code)) || [];
    if (!products.length) state.warnings.push(`Code ${dg.code} không tìm thấy sản phẩm.`);
    group.products.push(...products);
  }

  const groups = [];
  for (const group of groupMap.values()) {
    let products = group.products.map((p) => (el.removeVietgap.checked ? removeVietgapText(p) : clean(p))).filter(Boolean);
    if (el.removeDuplicates.checked) {
      const seen = new Set();
      products = products.filter((p) => {
        const k = key(p);
        if (seen.has(k)) return false;
        seen.add(k);
        return true;
      });
    }
    products = sortProducts(products);
    if (!products.length) continue;
    groups.push({ ...group, products });
  }
  if (!groups.length) throw new Error("Không có sản phẩm nào để đưa vào biểu mẫu QC.");
  return groups;
}

function makePage(index, layout) {
  return { index, capacity: index === 0 ? layout.firstCapacity : layout.cloneCapacity, used: 0, entries: [] };
}

function packGroups(groups, layout) {
  const pages = [makePage(0, layout)];
  const maxCapacity = layout.cloneCapacity;

  const current = () => pages[pages.length - 1];
  const addPage = () => {
    pages.push(makePage(pages.length, layout));
    return current();
  };

  for (const group of groups) {
    const count = group.products.length;
    if (count <= maxCapacity) {
      const requiredRows = Math.max(2, count);
      let page = current();
      if (requiredRows > page.capacity - page.used) page = addPage();
      if (requiredRows > page.capacity) {
        throw new Error(`Nhóm ${group.customer} có ${count} sản phẩm nhưng không vừa biểu mẫu.`);
      }
      page.entries.push({ ...group, rowOffset: page.used, products: [...group.products], splitPart: 1 });
      page.used += requiredRows;
      continue;
    }

    let remaining = [...group.products];
    let part = 1;
    while (remaining.length) {
      let page = current();
      let available = page.capacity - page.used;
      if (available < 2) {
        page = addPage();
        available = page.capacity;
      }
      const take = Math.min(available, remaining.length);
      const chunk = remaining.splice(0, take);
      page.entries.push({ ...group, rowOffset: page.used, products: chunk, splitPart: part });
      page.used += Math.max(2, chunk.length);
      part += 1;
    }
  }

  return pages;
}

function parseXml(text, label) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error(`Không đọc được XML: ${label}`);
  return doc;
}

function serializeXml(doc) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${new XMLSerializer().serializeToString(doc.documentElement)}`;
}


function elementChildren(node) {
  return Array.from(node?.childNodes || []).filter((n) => n.nodeType === 1);
}

function descendantsByNs(node, ns, localName) {
  return Array.from(node.getElementsByTagNameNS(ns, localName));
}

function firstByNs(node, ns, localName) {
  return node.getElementsByTagNameNS(ns, localName)[0] || null;
}

function resolvePackagePath(baseFile, target) {
  const base = baseFile.split("/");
  base.pop();
  for (const part of target.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") base.pop();
    else base.push(part);
  }
  return base.join("/");
}

function base64ToArrayBuffer(base64) {
  const binary = atob(base64.replace(/\s+/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

async function getTemplateBuffer() {
  if (state.templateFile) return state.templateFile.arrayBuffer();
  if (!window.QC_TEMPLATE_BASE64) throw new Error("Không tìm thấy mẫu QC tích hợp.");
  return base64ToArrayBuffer(window.QC_TEMPLATE_BASE64);
}

function readSharedStrings(doc) {
  if (!doc) return [];
  return descendantsByNs(doc, NS.main, "si").map((si) => descendantsByNs(si, NS.main, "t").map((t) => t.textContent || "").join(""));
}

function cellValue(cell, sharedStrings) {
  const type = cell.getAttribute("t") || "";
  if (type === "s") {
    const v = firstByNs(cell, NS.main, "v");
    return v ? sharedStrings[Number(v.textContent)] || "" : "";
  }
  if (type === "inlineStr") return descendantsByNs(cell, NS.main, "t").map((t) => t.textContent || "").join("");
  const v = firstByNs(cell, NS.main, "v");
  return v ? v.textContent || "" : "";
}

function splitCellRef(ref) {
  const match = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!match) throw new Error(`Địa chỉ ô không hợp lệ: ${ref}`);
  return { col: match[1], row: Number(match[2]) };
}

function columnNumber(col) {
  let n = 0;
  for (const ch of col) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}

function shiftCellRef(ref, rowOffset) {
  const { col, row } = splitCellRef(ref);
  return `${col}${row + rowOffset}`;
}

function shiftRangeRef(ref, rowOffset) {
  const parts = ref.split(":");
  return parts.map((part) => shiftCellRef(part, rowOffset)).join(":");
}

function rangeRows(ref) {
  const parts = ref.split(":");
  const a = splitCellRef(parts[0]);
  const b = splitCellRef(parts[1] || parts[0]);
  return { start: Math.min(a.row, b.row), end: Math.max(a.row, b.row), startCol: a.col, endCol: b.col };
}

function buildCellMap(sheetDoc, sharedStrings) {
  const map = new Map();
  for (const cell of descendantsByNs(sheetDoc, NS.main, "c")) {
    const ref = cell.getAttribute("r");
    if (ref) map.set(ref, cellValue(cell, sharedStrings));
  }
  return map;
}

function findMergeForCell(mergeRefs, ref) {
  const target = splitCellRef(ref);
  const targetCol = columnNumber(target.col);
  return mergeRefs.find((merge) => {
    const parts = merge.split(":");
    const a = splitCellRef(parts[0]);
    const b = splitCellRef(parts[1] || parts[0]);
    const minCol = Math.min(columnNumber(a.col), columnNumber(b.col));
    const maxCol = Math.max(columnNumber(a.col), columnNumber(b.col));
    return target.row >= Math.min(a.row, b.row) && target.row <= Math.max(a.row, b.row) && targetCol >= minCol && targetCol <= maxCol;
  }) || null;
}

function detectTemplateLayout(sheetDoc, sharedStrings, drawingDoc) {
  const cellMap = buildCellMap(sheetDoc, sharedStrings);
  const mergeRefs = descendantsByNs(sheetDoc, NS.main, "mergeCell").map((m) => m.getAttribute("ref"));

  const customerHeaders = [...cellMap.entries()]
    .filter(([ref, value]) => splitCellRef(ref).col === "J" && normalize(value).includes("khach hang") && normalize(value).includes("customer"))
    .map(([ref]) => splitCellRef(ref).row)
    .sort((a, b) => a - b);

  if (customerHeaders.length < 2) throw new Error("Mẫu QC không có đủ hai form gốc để xác định khoảng lặp.");

  const dataStarts = customerHeaders.slice(0, 2).map((headerRow) => {
    const merge = findMergeForCell(mergeRefs, `J${headerRow}`);
    return merge ? rangeRows(merge).end + 1 : headerRow + 1;
  });

  const findCompanyRow = (headerRow) => {
    for (let r = headerRow; r >= Math.max(1, headerRow - 12); r -= 1) {
      if (normalize(cellMap.get(`A${r}`)).startsWith("chi nhanh cong ty")) return r;
    }
    return null;
  };

  const findNoteRow = (dataStart) => {
    for (let r = dataStart; r <= dataStart + 40; r += 1) {
      const rowText = normalize(
        [...cellMap.entries()]
          .filter(([ref, value]) => splitCellRef(ref).row === r && clean(value))
          .map(([, value]) => value)
          .join(" "),
      );
      if (rowText.includes("ghi chu") || rowText.includes("cac chi tieu kiem tra dat")) return r;
    }
    return null;
  };

  const firstStart = findCompanyRow(customerHeaders[0]);
  const sourceStart = findCompanyRow(customerHeaders[1]);
  const noteRows = dataStarts.map(findNoteRow);
  if (!firstStart || !sourceStart || !noteRows[0] || !noteRows[1]) throw new Error("Không xác định được đầu form hoặc dòng Ghi chú trong mẫu QC.");

  let sourceEnd = 0;
  for (const [ref, value] of cellMap.entries()) {
    const { row } = splitCellRef(ref);
    if (row >= sourceStart && clean(value)) sourceEnd = Math.max(sourceEnd, row);
  }
  for (const merge of mergeRefs) {
    const rr = rangeRows(merge);
    if (rr.start >= sourceStart) sourceEnd = Math.max(sourceEnd, rr.end);
  }
  if (sourceEnd <= sourceStart) throw new Error("Không xác định được cuối form thứ hai.");

  const blockHeight = sourceEnd - sourceStart + 1;
  const firstCapacity = noteRows[0] - dataStarts[0];
  const cloneCapacity = noteRows[1] - dataStarts[1];
  if (cloneCapacity !== 25) throw new Error(`Mẫu form lặp có ${cloneCapacity} dòng sản phẩm, không phải 25.`);

  const rowHeights = new Map();
  for (const rowEl of descendantsByNs(sheetDoc, NS.main, "row")) {
    rowHeights.set(Number(rowEl.getAttribute("r")), Number(rowEl.getAttribute("ht") || 15));
  }
  let pageHeightEmu = 0;
  for (let r = sourceStart; r <= sourceEnd; r += 1) pageHeightEmu += (rowHeights.get(r) || 15) * 12700;
  pageHeightEmu = Math.round(pageHeightEmu);

  const signatureAnchors = [];
  if (drawingDoc) {
    const candidates = [];
    for (const anchor of elementChildren(drawingDoc.documentElement)) {
      const from = elementChildren(anchor).find((n) => n.namespaceURI === NS.xdr && n.localName === "from");
      if (!from) continue;
      const rowNode = elementChildren(from).find((n) => n.namespaceURI === NS.xdr && n.localName === "row");
      const colNode = elementChildren(from).find((n) => n.namespaceURI === NS.xdr && n.localName === "col");
      const anchorRow = Number(rowNode?.textContent ?? -1);
      const anchorCol = Number(colNode?.textContent ?? -1);
      const text = descendantsByNs(anchor, NS.a, "t").map((t) => t.textContent || "").join("");
      const textNorm = normalize(text);
      if (anchorRow === sourceEnd - 3 && (textNorm.includes("reported by") || textNorm.includes("inspected by"))) {
        const off = descendantsByNs(anchor, NS.a, "off")[0];
        candidates.push({ anchor, anchorCol, textNorm, y: Number(off?.getAttribute("y") || 0) });
      }
    }
    for (const kind of ["reported by", "inspected by"]) {
      const matching = candidates.filter((c) => c.textNorm.includes(kind)).sort((a, b) => b.y - a.y);
      if (matching[0]) signatureAnchors.push(matching[0].anchor);
    }
  }

  return {
    firstStart,
    firstEnd: sourceStart - 1,
    firstDataStart: dataStarts[0],
    firstCapacity,
    sourceStart,
    sourceEnd,
    sourceDataStart: dataStarts[1],
    cloneCapacity,
    blockHeight,
    noteRows,
    mergeRefs,
    pageHeightEmu,
    signatureAnchors,
  };
}

async function openTemplatePackage(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const workbookPath = "xl/workbook.xml";
  const workbookRelsPath = "xl/_rels/workbook.xml.rels";
  const workbookDoc = parseXml(await zip.file(workbookPath).async("string"), workbookPath);
  const workbookRelsDoc = parseXml(await zip.file(workbookRelsPath).async("string"), workbookRelsPath);

  const sheet = descendantsByNs(workbookDoc, NS.main, "sheet").find((s) => s.getAttribute("name") === "Xe to");
  if (!sheet) throw new Error("Không tìm thấy sheet 'Xe to' trong mẫu QC.");
  const relationId = sheet.getAttributeNS(NS.officeRel, "id") || sheet.getAttribute("r:id");
  const relation = descendantsByNs(workbookRelsDoc, NS.rel, "Relationship").find((r) => r.getAttribute("Id") === relationId);
  if (!relation) throw new Error("Không tìm thấy liên kết XML của sheet 'Xe to'.");
  const sheetPath = resolvePackagePath(workbookPath, relation.getAttribute("Target"));
  const sheetDoc = parseXml(await zip.file(sheetPath).async("string"), sheetPath);

  let sharedStringsDoc = null;
  const sharedFile = zip.file("xl/sharedStrings.xml");
  if (sharedFile) sharedStringsDoc = parseXml(await sharedFile.async("string"), "xl/sharedStrings.xml");
  const sharedStrings = readSharedStrings(sharedStringsDoc);

  const sheetRelsPath = `${sheetPath.slice(0, sheetPath.lastIndexOf("/"))}/_rels/${sheetPath.slice(sheetPath.lastIndexOf("/") + 1)}.rels`;
  let sheetRelsDoc = null;
  let drawingPath = null;
  let drawingDoc = null;
  if (zip.file(sheetRelsPath)) {
    sheetRelsDoc = parseXml(await zip.file(sheetRelsPath).async("string"), sheetRelsPath);
    const drawingRel = descendantsByNs(sheetRelsDoc, NS.rel, "Relationship").find((r) => (r.getAttribute("Type") || "").endsWith("/drawing"));
    if (drawingRel) {
      drawingPath = resolvePackagePath(sheetPath, drawingRel.getAttribute("Target"));
      if (zip.file(drawingPath)) drawingDoc = parseXml(await zip.file(drawingPath).async("string"), drawingPath);
    }
  }

  const layout = detectTemplateLayout(sheetDoc, sharedStrings, drawingDoc);
  return { zip, workbookDoc, workbookPath, sheetDoc, sheetPath, drawingDoc, drawingPath, layout };
}

function directRows(sheetDoc) {
  const sheetData = firstByNs(sheetDoc, NS.main, "sheetData");
  return elementChildren(sheetData).filter((n) => n.namespaceURI === NS.main && n.localName === "row");
}

function rowMap(sheetDoc) {
  return new Map(directRows(sheetDoc).map((r) => [Number(r.getAttribute("r")), r]));
}

function cloneRowsForForms(ctx, formCount) {
  const { sheetDoc, layout } = ctx;
  const sheetData = firstByNs(sheetDoc, NS.main, "sheetData");
  const sourceRows = directRows(sheetDoc)
    .filter((row) => {
      const r = Number(row.getAttribute("r"));
      return r >= layout.sourceStart && r <= layout.sourceEnd;
    })
    .map((row) => row.cloneNode(true));

  for (const row of directRows(sheetDoc)) {
    if (Number(row.getAttribute("r")) > layout.sourceEnd) sheetData.removeChild(row);
  }

  const mergeCells = firstByNs(sheetDoc, NS.main, "mergeCells");
  const existingMergeNodes = mergeCells ? elementChildren(mergeCells).filter((n) => n.namespaceURI === NS.main && n.localName === "mergeCell") : [];
  if (!mergeCells) throw new Error("Mẫu QC không có vùng merge.");
  for (const node of existingMergeNodes) {
    const ref = node.getAttribute("ref");
    if (rangeRows(ref).start > layout.sourceEnd) mergeCells.removeChild(node);
  }
  const sourceMerges = elementChildren(mergeCells)
    .filter((n) => n.namespaceURI === NS.main && n.localName === "mergeCell")
    .map((n) => n.getAttribute("ref"))
    .filter((ref) => {
      const rr = rangeRows(ref);
      return rr.start >= layout.sourceStart && rr.end <= layout.sourceEnd;
    });

  for (let formIndex = 2; formIndex < formCount; formIndex += 1) {
    const offset = (formIndex - 1) * layout.blockHeight;
    for (const sourceRow of sourceRows) {
      const clone = sourceRow.cloneNode(true);
      clone.setAttribute("r", String(Number(clone.getAttribute("r")) + offset));
      for (const cell of descendantsByNs(clone, NS.main, "c")) {
        const ref = cell.getAttribute("r");
        if (ref) cell.setAttribute("r", shiftCellRef(ref, offset));
      }
      for (const formula of descendantsByNs(clone, NS.main, "f")) {
        if (formula.textContent) formula.textContent = formula.textContent.replace(/([A-Z]{1,3})(\d+)/g, (_, col, row) => `${col}${Number(row) + offset}`);
      }
      sheetData.appendChild(clone);
    }
    for (const sourceMerge of sourceMerges) {
      const merge = sheetDoc.createElementNS(NS.main, "mergeCell");
      merge.setAttribute("ref", shiftRangeRef(sourceMerge, offset));
      mergeCells.appendChild(merge);
    }
  }
  mergeCells.setAttribute("count", String(elementChildren(mergeCells).filter((n) => n.localName === "mergeCell").length));

  const lastRow = formCount === 1 ? layout.firstEnd : layout.sourceEnd + Math.max(0, formCount - 2) * layout.blockHeight;
  const dimension = firstByNs(sheetDoc, NS.main, "dimension");
  if (dimension) dimension.setAttribute("ref", `A2:X${lastRow}`);

  let rowBreaks = firstByNs(sheetDoc, NS.main, "rowBreaks");
  if (!rowBreaks) {
    rowBreaks = sheetDoc.createElementNS(NS.main, "rowBreaks");
    sheetDoc.documentElement.appendChild(rowBreaks);
  }
  while (rowBreaks.firstChild) rowBreaks.removeChild(rowBreaks.firstChild);
  const breaks = [];
  if (formCount > 1) breaks.push(layout.firstEnd);
  for (let i = 2; i < formCount; i += 1) breaks.push(layout.sourceEnd + (i - 2) * layout.blockHeight);
  for (const breakRow of breaks) {
    const brk = sheetDoc.createElementNS(NS.main, "brk");
    brk.setAttribute("id", String(breakRow));
    brk.setAttribute("max", "23");
    brk.setAttribute("man", "1");
    rowBreaks.appendChild(brk);
  }
  rowBreaks.setAttribute("count", String(breaks.length));
  rowBreaks.setAttribute("manualBreakCount", String(breaks.length));
  return lastRow;
}

function cloneSignatureDrawings(ctx, formCount) {
  const { drawingDoc, layout } = ctx;
  if (!drawingDoc || !layout.signatureAnchors.length || formCount <= 2) return;
  let maxId = 0;
  for (const cNvPr of descendantsByNs(drawingDoc, NS.xdr, "cNvPr")) maxId = Math.max(maxId, Number(cNvPr.getAttribute("id") || 0));

  for (let formIndex = 2; formIndex < formCount; formIndex += 1) {
    const rowOffset = (formIndex - 1) * layout.blockHeight;
    const yOffset = (formIndex - 1) * layout.pageHeightEmu;
    for (const sourceAnchor of layout.signatureAnchors) {
      const clone = sourceAnchor.cloneNode(true);
      const from = elementChildren(clone).find((n) => n.namespaceURI === NS.xdr && n.localName === "from");
      if (from) {
        const rowNode = elementChildren(from).find((n) => n.namespaceURI === NS.xdr && n.localName === "row");
        if (rowNode) rowNode.textContent = String(Number(rowNode.textContent) + rowOffset);
      }
      const to = elementChildren(clone).find((n) => n.namespaceURI === NS.xdr && n.localName === "to");
      if (to) {
        const rowNode = elementChildren(to).find((n) => n.namespaceURI === NS.xdr && n.localName === "row");
        if (rowNode) rowNode.textContent = String(Number(rowNode.textContent) + rowOffset);
      }
      const off = descendantsByNs(clone, NS.a, "off")[0];
      if (off) off.setAttribute("y", String(Number(off.getAttribute("y") || 0) + yOffset));
      const cNvPr = descendantsByNs(clone, NS.xdr, "cNvPr")[0];
      if (cNvPr) {
        maxId += 1;
        cNvPr.setAttribute("id", String(maxId));
        cNvPr.setAttribute("name", `${cNvPr.getAttribute("name") || "TextBox"} Form ${formIndex + 1}`);
      }
      for (const creationId of descendantsByNs(clone, NS.a16, "creationId")) {
        if (crypto.randomUUID) creationId.setAttribute("id", `{${crypto.randomUUID().toUpperCase()}}`);
      }
      drawingDoc.documentElement.appendChild(clone);
    }
  }
}

function ensureCell(sheetDoc, rowEl, ref) {
  let target = elementChildren(rowEl).find((n) => n.namespaceURI === NS.main && n.localName === "c" && n.getAttribute("r") === ref);
  if (target) return target;
  target = sheetDoc.createElementNS(NS.main, "c");
  target.setAttribute("r", ref);
  const targetCol = columnNumber(splitCellRef(ref).col);
  const cells = elementChildren(rowEl).filter((n) => n.namespaceURI === NS.main && n.localName === "c");
  const before = cells.find((c) => columnNumber(splitCellRef(c.getAttribute("r")).col) > targetCol);
  if (before) rowEl.insertBefore(target, before);
  else rowEl.appendChild(target);
  return target;
}

function setInlineString(sheetDoc, rows, ref, value) {
  const { row } = splitCellRef(ref);
  const rowEl = rows.get(row);
  if (!rowEl) throw new Error(`Mẫu thiếu dòng ${row}.`);
  const cell = ensureCell(sheetDoc, rowEl, ref);
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  if (!value) {
    cell.removeAttribute("t");
    return;
  }
  cell.setAttribute("t", "inlineStr");
  const is = sheetDoc.createElementNS(NS.main, "is");
  const text = sheetDoc.createElementNS(NS.main, "t");
  if (value !== value.trim() || value.includes("\n")) text.setAttributeNS(NS.xml, "xml:space", "preserve");
  text.textContent = value;
  is.appendChild(text);
  cell.appendChild(is);
}

function clearAndFillForms(ctx, pages, formCount) {
  const rows = rowMap(ctx.sheetDoc);
  const formInfo = [];
  formInfo.push({ dataStart: ctx.layout.firstDataStart, capacity: ctx.layout.firstCapacity });
  for (let i = 1; i < formCount; i += 1) {
    formInfo.push({
      dataStart: ctx.layout.sourceDataStart + Math.max(0, i - 1) * ctx.layout.blockHeight,
      capacity: ctx.layout.cloneCapacity,
    });
  }

  for (const form of formInfo) {
    for (let r = form.dataStart; r < form.dataStart + form.capacity; r += 1) {
      setInlineString(ctx.sheetDoc, rows, `J${r}`, "");
      setInlineString(ctx.sheetDoc, rows, `K${r}`, "");
    }
  }

  pages.forEach((page, pageIndex) => {
    const form = formInfo[pageIndex];
    page.entries.forEach((entry) => {
      const start = form.dataStart + entry.rowOffset;
      setInlineString(ctx.sheetDoc, rows, `J${start}`, entry.customer);
      if (start + 1 < form.dataStart + form.capacity) setInlineString(ctx.sheetDoc, rows, `J${start + 1}`, entry.deliveryPoint);
      entry.products.forEach((product, idx) => setInlineString(ctx.sheetDoc, rows, `K${start + idx}`, product));
    });
  });
}

function updatePrintArea(ctx, lastRow) {
  for (const definedName of descendantsByNs(ctx.workbookDoc, NS.main, "definedName")) {
    if (definedName.getAttribute("name") === "_xlnm.Print_Area") {
      const text = definedName.textContent || "";
      if (text.includes("'Xe to'")) definedName.textContent = `'Xe to'!$A$1:$X$${lastRow}`;
    }
  }
}

async function generateQcWorkbook(groups, pages) {
  const ctx = await openTemplatePackage(await getTemplateBuffer());
  state.templateLayout = ctx.layout;
  const formCount = Math.max(2, pages.length);
  const lastRow = cloneRowsForForms(ctx, formCount);
  cloneSignatureDrawings(ctx, formCount);
  clearAndFillForms(ctx, pages, formCount);
  updatePrintArea(ctx, lastRow);

  ctx.zip.file(ctx.sheetPath, serializeXml(ctx.sheetDoc));
  ctx.zip.file(ctx.workbookPath, serializeXml(ctx.workbookDoc));
  if (ctx.drawingDoc && ctx.drawingPath) ctx.zip.file(ctx.drawingPath, serializeXml(ctx.drawingDoc));

  const blob = await ctx.zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  return { blob, layout: ctx.layout };
}

function renderResult() {
  el.resultBody.innerHTML = "";
  state.pages.forEach((page, pageIndex) => {
    const pageRow = document.createElement("tr");
    pageRow.className = "page-row";
    pageRow.innerHTML = `<td colspan="4"><strong>Form ${pageIndex + 1}</strong> · ${page.used}/${page.capacity} dòng đã dùng</td>`;
    el.resultBody.appendChild(pageRow);
    page.entries.forEach((entry) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(entry.customer)}</td>
        <td>${escapeHtml(entry.deliveryPoint)}</td>
        <td>${entry.products.length}</td>
        <td>${escapeHtml(entry.products.slice(0, 4).join(" · "))}${entry.products.length > 4 ? " …" : ""}</td>
      `;
      el.resultBody.appendChild(tr);
    });
  });
  el.warningList.innerHTML = state.warnings.length
    ? state.warnings.map((w) => `<li>${escapeHtml(w)}</li>`).join("")
    : "<li>Không có cảnh báo.</li>";
  el.resultSection.classList.remove("hidden");
}

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[ch]));
}

function safeFilename(value) {
  return stripAccents(value).replace(/[^A-Za-z0-9_-]+/g, "_").replace(/^_+|_+$/g, "") || "Tat_ca_Ship";
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

async function processAllData() {
  try {
    setMessage("Đang ghép dữ liệu và tạo biểu mẫu QC…", "working");
    el.processButton.disabled = true;
    state.groups = buildGroups();

    const templatePreview = await openTemplatePackage(await getTemplateBuffer());
    state.templateLayout = templatePreview.layout;
    state.pages = packGroups(state.groups, templatePreview.layout);
    const { blob, layout } = await generateQcWorkbook(state.groups, state.pages);
    state.outputBlob = blob;

    const shipLabel = el.processAll.checked ? "Tat_ca_Ship" : safeFilename(el.shipSelect.value);
    const date = new Date().toISOString().slice(0, 10);
    state.outputName = `Bao_cao_QC_${shipLabel}_${date}.xlsx`;

    el.layoutInfo.textContent = `Mẫu nhận diện: Form 1 có ${layout.firstCapacity} dòng; Form lặp có ${layout.cloneCapacity} dòng; khoảng lặp ${layout.blockHeight} dòng; chữ ký textbox được sao chép.`;
    renderResult();
    updateMetrics();
    setMessage(`Đã tạo ${state.pages.length} form QC. File mới giữ phần Người báo cáo, Kho & vận chuyển và Người kiểm tra.`, "success");
  } catch (error) {
    console.error(error);
    state.outputBlob = null;
    setMessage(error.message || "Có lỗi khi xử lý dữ liệu.", "error");
  } finally {
    el.processButton.disabled = !(state.dgRows.length && state.productRows.length);
  }
}

el.dgFiles.addEventListener("change", async () => {
  try {
    setMessage("Đang đọc file DG…", "working");
    state.dgRows = await parseDgFiles([...el.dgFiles.files]);
    el.dgStatus.textContent = `Đã đọc ${state.dgRows.length} dòng từ ${el.dgFiles.files.length} file DG.`;
    updateShipSelect();
    updateMetrics();
    setMessage("Đã đọc file DG.", "success");
  } catch (error) {
    state.dgRows = [];
    el.dgStatus.textContent = error.message;
    updateShipSelect();
    updateMetrics();
    setMessage(error.message, "error");
  }
});

el.productFiles.addEventListener("change", async () => {
  try {
    setMessage("Đang đọc file sản phẩm…", "working");
    state.productRows = await parseProductFiles([...el.productFiles.files]);
    el.productStatus.textContent = `Đã đọc ${state.productRows.length} dòng từ ${el.productFiles.files.length} file sản phẩm.`;
    updateMetrics();
    setMessage("Đã đọc file sản phẩm.", "success");
  } catch (error) {
    state.productRows = [];
    el.productStatus.textContent = error.message;
    updateMetrics();
    setMessage(error.message, "error");
  }
});

el.templateFile.addEventListener("change", async () => {
  state.templateFile = el.templateFile.files[0] || null;
  if (!state.templateFile) {
    el.templateStatus.textContent = "Đang dùng mẫu QC tích hợp.";
    return;
  }
  try {
    const ctx = await openTemplatePackage(await state.templateFile.arrayBuffer());
    state.templateLayout = ctx.layout;
    el.templateStatus.textContent = `Mẫu hợp lệ: Form 1 ${ctx.layout.firstCapacity} dòng, form lặp ${ctx.layout.cloneCapacity} dòng.`;
    setMessage("Mẫu QC tải lên đã được kiểm tra.", "success");
  } catch (error) {
    state.templateFile = null;
    el.templateFile.value = "";
    el.templateStatus.textContent = "Mẫu không hợp lệ; đã quay lại mẫu tích hợp.";
    setMessage(error.message, "error");
  }
});

el.loadDemoButton.addEventListener("click", () => {
  state.dgRows = demoDgRows.map((row) => ({ ...row }));
  state.productRows = demoProductRows.map((row) => ({ ...row }));
  el.dgStatus.textContent = `Đã nạp ${state.dgRows.length} dòng DG demo.`;
  el.productStatus.textContent = `Đã nạp ${state.productRows.length} dòng sản phẩm demo.`;
  updateShipSelect();
  el.shipSelect.value = "24SPF47158";
  updateMetrics();
  setMessage("Đã nạp dữ liệu demo. Nhấn “Tạo báo cáo QC”.", "success");
});

el.processButton.addEventListener("click", processAllData);

el.downloadButton.addEventListener("click", () => {
  if (!state.outputBlob) return setMessage("Hãy tạo báo cáo QC trước.", "error");
  downloadBlob(state.outputBlob, state.outputName);
});

el.shareButton.addEventListener("click", async () => {
  if (!state.outputBlob) return setMessage("Hãy tạo báo cáo QC trước.", "error");
  const file = new File([state.outputBlob], state.outputName, { type: state.outputBlob.type });
  if (navigator.canShare && navigator.canShare({ files: [file] }) && navigator.share) {
    try {
      await navigator.share({ files: [file], title: "Báo cáo QC xuất hàng" });
      return;
    } catch (error) {
      if (error.name === "AbortError") return;
    }
  }
  downloadBlob(state.outputBlob, state.outputName);
});

el.processAll.addEventListener("change", () => {
  el.shipSelect.disabled = el.processAll.checked || !state.dgRows.length;
});

(async function init() {
  el.templateStatus.textContent = "Đang dùng mẫu QC tích hợp: BM-QC-26.xlsx.";
  try {
    const ctx = await openTemplatePackage(await getTemplateBuffer());
    state.templateLayout = ctx.layout;
    el.layoutInfo.textContent = `Mẫu tích hợp đã kiểm tra: Form 1 có ${ctx.layout.firstCapacity} dòng sản phẩm; form lặp có ${ctx.layout.cloneCapacity} dòng; khoảng lặp ${ctx.layout.blockHeight} dòng.`;
  } catch (error) {
    console.error(error);
    el.layoutInfo.textContent = `Không kiểm tra được mẫu tích hợp: ${error.message}`;
  }
  updateShipSelect();
  updateMetrics();
})();
