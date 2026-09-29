(() => {
  "use strict";

  const MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
  ];
  const MONTH_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const A4 = [595.303937, 841.889764];
  const PAGE_BOTTOM = 800;

  const elements = {
    themeToggle: document.getElementById("themeToggle"),
    themeIcon: document.getElementById("themeIcon"),
    themeLabel: document.getElementById("themeLabel"),
    fileInput: document.getElementById("fileInput"),
    selectFileButton: document.getElementById("selectFileButton"),
    generateButton: document.getElementById("generateButton"),
    dropZone: document.getElementById("dropZone"),
    fileName: document.getElementById("fileName"),
    fileStateChip: document.getElementById("fileStateChip"),
    status: document.getElementById("status"),
    summaryMonth: document.getElementById("summaryMonth"),
    summaryOutlets: document.getElementById("summaryOutlets"),
    summaryRecipients: document.getElementById("summaryRecipients"),
    summaryTotal: document.getElementById("summaryTotal"),
    approvalDate: document.getElementById("approvalDate"),
  };

  const state = {
    file: null,
    data: null,
    valid: false,
    busy: false,
  };

  const assetCache = new Map();

  function setTheme(theme) {
    const next = theme === "light" ? "light" : "dark";
    document.documentElement.dataset.theme = next;
    elements.themeIcon.textContent = next === "dark" ? "☀" : "☾";
    elements.themeLabel.textContent = next === "dark" ? "Light mode" : "Dark mode";
    try {
      localStorage.setItem("new-outlet-approval-theme", next);
    } catch (_error) {
      // The dashboard remains usable when storage is unavailable.
    }
  }

  function initTheme() {
    let saved = "dark";
    try {
      saved = localStorage.getItem("new-outlet-approval-theme") || "dark";
    } catch (_error) {
      saved = "dark";
    }
    setTheme(saved);
  }

  function setStatus(message, kind = "") {
    elements.status.textContent = message;
    elements.status.className = `status${kind ? ` ${kind}` : ""}`;
  }

  function setChip(text, kind = "idle") {
    elements.fileStateChip.textContent = text;
    elements.fileStateChip.className = `chip chip-${kind}`;
  }

  function setBusy(busy) {
    state.busy = busy;
    elements.selectFileButton.disabled = busy;
    elements.fileInput.disabled = busy;
    elements.generateButton.disabled = busy || !state.valid;
    elements.generateButton.textContent = busy
      ? "Generating approval PDF…"
      : "New outlet opening approval Generate";
  }

  function resetSummary() {
    elements.summaryMonth.textContent = "—";
    elements.summaryOutlets.textContent = "—";
    elements.summaryRecipients.textContent = "—";
    elements.summaryTotal.textContent = "—";
    elements.approvalDate.textContent = "Derived after file selection";
  }

  function normalizeText(value) {
    return String(value ?? "")
      .replace(/\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  }

  function asText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
  }

  function asNumber(value) {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    const parsed = Number(String(value ?? "").replace(/,/g, "").trim());
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function formatInteger(value) {
    return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(Math.round(value));
  }

  function ordinalSuffix(day) {
    const remainder100 = day % 100;
    if (remainder100 >= 11 && remainder100 <= 13) return "th";
    if (day % 10 === 1) return "st";
    if (day % 10 === 2) return "nd";
    if (day % 10 === 3) return "rd";
    return "th";
  }

  function approvalDateParts(year, monthIndex) {
    const date = new Date(Date.UTC(year, monthIndex + 1, 3));
    const day = date.getUTCDate();
    const month = MONTHS[date.getUTCMonth()];
    const full = `${String(day).padStart(2, "0")}${ordinalSuffix(day)} ${month} ${date.getUTCFullYear()}`;
    const reference = `${String(day).padStart(2, "0")}${String(date.getUTCMonth() + 1).padStart(2, "0")}${date.getUTCFullYear()}`;
    return { date, day, month, year: date.getUTCFullYear(), full, reference };
  }

  function parseXml(source, label) {
    const documentXml = new DOMParser().parseFromString(source, "application/xml");
    const error = documentXml.getElementsByTagName("parsererror")[0];
    if (error) throw new Error(`Unable to read ${label}.`);
    return documentXml;
  }

  function byLocalName(root, localName) {
    return Array.from(root.getElementsByTagNameNS("*", localName));
  }

  function firstByLocalName(root, localName) {
    return root.getElementsByTagNameNS("*", localName)[0] || null;
  }

  function normalizedZipPath(path) {
    const parts = [];
    String(path).replace(/^\//, "").split("/").forEach((part) => {
      if (!part || part === ".") return;
      if (part === "..") parts.pop();
      else parts.push(part);
    });
    return parts.join("/");
  }

  function columnIndexFromRef(reference) {
    const match = String(reference || "").match(/^([A-Z]+)/i);
    if (!match) return -1;
    return match[1].toUpperCase().split("").reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1;
  }

  async function zipText(zip, path, required = true) {
    const entry = zip.file(path);
    if (!entry) {
      if (!required) return "";
      throw new Error(`The workbook is missing ${path}.`);
    }
    return entry.async("string");
  }

  function parseSharedStrings(source) {
    if (!source) return [];
    const xml = parseXml(source, "shared strings");
    return byLocalName(xml, "si").map((item) => byLocalName(item, "t").map((node) => node.textContent || "").join(""));
  }

  function parseCellValue(cell, sharedStrings) {
    const type = cell.getAttribute("t") || "";
    if (type === "inlineStr") {
      const inline = firstByLocalName(cell, "is");
      return inline ? byLocalName(inline, "t").map((node) => node.textContent || "").join("") : "";
    }
    const valueNode = firstByLocalName(cell, "v");
    if (!valueNode) return null;
    const raw = valueNode.textContent || "";
    if (type === "s") return sharedStrings[Number(raw)] ?? "";
    if (type === "str" || type === "e") return raw;
    if (type === "b") return raw === "1";
    const numeric = Number(raw);
    return Number.isFinite(numeric) ? numeric : raw;
  }

  function parseWorksheet(source, sharedStrings, name) {
    const xml = parseXml(source, `worksheet ${name}`);
    const rows = byLocalName(xml, "row").map((rowNode) => {
      const number = Number(rowNode.getAttribute("r")) || 0;
      const values = [];
      byLocalName(rowNode, "c").forEach((cell) => {
        const column = columnIndexFromRef(cell.getAttribute("r"));
        if (column >= 0) values[column] = parseCellValue(cell, sharedStrings);
      });
      return { number, values };
    });
    return { name, rows };
  }

  async function parseWorkbookFile(file) {
    if (!window.JSZip) throw new Error("The Excel reader did not load.");
    const zip = await window.JSZip.loadAsync(await file.arrayBuffer());
    const workbookXml = parseXml(await zipText(zip, "xl/workbook.xml"), "workbook");
    const relationshipsXml = parseXml(await zipText(zip, "xl/_rels/workbook.xml.rels"), "workbook relationships");
    const sharedStrings = parseSharedStrings(await zipText(zip, "xl/sharedStrings.xml", false));

    const relationships = new Map();
    byLocalName(relationshipsXml, "Relationship").forEach((relationship) => {
      relationships.set(relationship.getAttribute("Id"), relationship.getAttribute("Target"));
    });

    const sheets = [];
    for (const sheetNode of byLocalName(workbookXml, "sheet")) {
      const name = sheetNode.getAttribute("name") || "Worksheet";
      const relationshipId = sheetNode.getAttribute("r:id")
        || sheetNode.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
      const target = relationships.get(relationshipId);
      if (!target) continue;
      const path = target.startsWith("/")
        ? normalizedZipPath(target)
        : normalizedZipPath(`xl/${target}`);
      const source = await zipText(zip, path);
      sheets.push(parseWorksheet(source, sharedStrings, name));
    }

    if (!sheets.length) throw new Error("No worksheets were found in this Excel file.");
    return { sheets };
  }

  function rowHasLabels(row, labels) {
    const values = row.values.map(normalizeText);
    return labels.every((label) => values.includes(normalizeText(label)));
  }

  function findHeader(sheet, labels) {
    return sheet.rows.find((row) => rowHasLabels(row, labels)) || null;
  }

  function firstIndex(row, label) {
    const target = normalizeText(label);
    return row.values.findIndex((value) => normalizeText(value) === target);
  }

  function allIndices(row, label) {
    const target = normalizeText(label);
    const indices = [];
    row.values.forEach((value, index) => {
      if (normalizeText(value) === target) indices.push(index);
    });
    return indices;
  }

  function excelDate(value) {
    if (value instanceof Date && !Number.isNaN(value.valueOf())) return value;
    if (typeof value === "number" && Number.isFinite(value)) {
      return new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
    }
    if (typeof value === "string" && value.trim()) {
      const parsed = new Date(value.trim());
      if (!Number.isNaN(parsed.valueOf())) return parsed;
      const compact = value.trim().match(/^(\d{1,2})[-\/]([A-Za-z]{3,9})[-\/](\d{2,4})$/);
      if (compact) {
        const month = MONTHS.findIndex((name) => name.toLowerCase().startsWith(compact[2].toLowerCase().slice(0, 3)));
        const year = Number(compact[3]) < 100 ? 2000 + Number(compact[3]) : Number(compact[3]);
        if (month >= 0) return new Date(Date.UTC(year, month, Number(compact[1])));
      }
    }
    return null;
  }

  function detectPeriod(workbook, detailRows) {
    const aliases = MONTHS.flatMap((month, index) => [
      { value: month, index },
      { value: MONTH_SHORT[index], index },
    ]).sort((a, b) => b.value.length - a.value.length);

    const monthPattern = aliases.map((entry) => entry.value).join("|");
    const regex = new RegExp(`(?:^|[^A-Za-z])(${monthPattern})[\\s_\\-]*(20\\d{2})(?:$|[^0-9])`, "i");
    for (const sheet of workbook.sheets) {
      const firstRows = sheet.rows.slice(0, 6).flatMap((row) => row.values).filter((value) => value !== null && value !== undefined).join(" ");
      const candidate = `${sheet.name} ${firstRows}`;
      const match = candidate.match(regex);
      if (match) {
        const monthIndex = aliases.find((entry) => entry.value.toLowerCase() === match[1].toLowerCase()).index;
        return { monthIndex, year: Number(match[2]) };
      }
    }

    const counts = new Map();
    detailRows.forEach((row) => {
      const date = excelDate(row.date);
      if (!date) return;
      const key = `${date.getUTCFullYear()}-${date.getUTCMonth()}`;
      counts.set(key, (counts.get(key) || 0) + 1);
    });
    const mostCommon = Array.from(counts.entries()).sort((a, b) => b[1] - a[1])[0];
    if (mostCommon) {
      const [year, monthIndex] = mostCommon[0].split("-").map(Number);
      return { year, monthIndex };
    }
    throw new Error("The reporting month could not be detected from the workbook title or inauguration dates.");
  }

  function extractDashboardData(workbook) {
    let detailSheet = null;
    let detailHeader = null;
    let summarySheet = null;
    let summaryHeader = null;

    workbook.sheets.forEach((sheet) => {
      const detailCandidate = findHeader(sheet, ["SL", "Inauguration Date", "Code", "Outlet Name", "Outlet Team Member", "Zonal Name", "Expansion By"]);
      if (detailCandidate && !detailSheet) {
        detailSheet = sheet;
        detailHeader = detailCandidate;
      }
      const summaryCandidate = findHeader(sheet, ["SL", "Expansion By", "ID", "Designation", "Incentive Amount"]);
      if (summaryCandidate && !summarySheet && !detailCandidate) {
        summarySheet = sheet;
        summaryHeader = summaryCandidate;
      }
    });

    if (!detailSheet || !detailHeader) {
      throw new Error("Appendix A source was not found. The workbook must contain the outlet detail headers.");
    }
    if (!summarySheet || !summaryHeader) {
      throw new Error("Appendix B source was not found. The workbook must contain a Summary sheet with recipient details.");
    }

    const detailIndices = {
      sl: firstIndex(detailHeader, "SL"),
      date: firstIndex(detailHeader, "Inauguration Date"),
      code: firstIndex(detailHeader, "Code"),
      outlet: firstIndex(detailHeader, "Outlet Name"),
      teamMember: firstIndex(detailHeader, "Outlet Team Member"),
      zonalName: firstIndex(detailHeader, "Zonal Name"),
      rhoName: firstIndex(detailHeader, "Expansion By"),
    };
    const incentiveIndices = allIndices(detailHeader, "Incentive Amount");
    const quantityIndices = allIndices(detailHeader, "Outlet Quantity");
    if (incentiveIndices.length < 4 || quantityIndices.length < 3) {
      throw new Error("The outlet detail sheet does not contain the expected incentive columns.");
    }
    Object.assign(detailIndices, {
      teamAmount: incentiveIndices[0],
      zonalAmount: incentiveIndices[1],
      rhoAmount: incentiveIndices[2],
      assistantAmount: incentiveIndices[3],
      rhoQuantity: quantityIndices[2],
    });

    const detailRows = [];
    for (const row of detailSheet.rows.filter((item) => item.number > detailHeader.number)) {
      if (row.values.some((value) => normalizeText(value) === "grand total")) break;
      const code = asText(row.values[detailIndices.code]);
      const outlet = asText(row.values[detailIndices.outlet]);
      const sl = row.values[detailIndices.sl];
      if (!code && !outlet && (sl === null || sl === undefined || sl === "")) continue;
      detailRows.push({
        sl: asNumber(sl) || detailRows.length + 1,
        date: row.values[detailIndices.date],
        code,
        outlet,
        teamMember: asText(row.values[detailIndices.teamMember]),
        teamAmount: asNumber(row.values[detailIndices.teamAmount]),
        zonalName: asText(row.values[detailIndices.zonalName]),
        zonalAmount: asNumber(row.values[detailIndices.zonalAmount]),
        rhoName: asText(row.values[detailIndices.rhoName]),
        rhoQuantity: asNumber(row.values[detailIndices.rhoQuantity]),
        rhoAmount: asNumber(row.values[detailIndices.rhoAmount]),
        assistantAmount: asNumber(row.values[detailIndices.assistantAmount]),
      });
    }

    if (!detailRows.length) throw new Error("No outlet rows were found for Appendix A.");

    let assistantLabel = "Assistant Director";
    detailSheet.rows
      .filter((row) => row.number < detailHeader.number)
      .flatMap((row) => row.values)
      .forEach((value) => {
        if (normalizeText(value).startsWith("assistant director")) assistantLabel = asText(value);
      });

    const summaryIndices = {
      sl: firstIndex(summaryHeader, "SL"),
      name: firstIndex(summaryHeader, "Expansion By"),
      id: firstIndex(summaryHeader, "ID"),
      designation: firstIndex(summaryHeader, "Designation"),
      amount: firstIndex(summaryHeader, "Incentive Amount"),
    };
    const summaryRows = [];
    for (const row of summarySheet.rows.filter((item) => item.number > summaryHeader.number)) {
      if (row.values.some((value) => normalizeText(value) === "grand total")) break;
      const name = asText(row.values[summaryIndices.name]);
      if (!name) continue;
      summaryRows.push({
        sl: asNumber(row.values[summaryIndices.sl]) || summaryRows.length + 1,
        name,
        id: asText(row.values[summaryIndices.id]),
        designation: asText(row.values[summaryIndices.designation]),
        amount: asNumber(row.values[summaryIndices.amount]),
      });
    }
    if (!summaryRows.length) throw new Error("No recipient rows were found for Appendix B.");

    const period = detectPeriod(workbook, detailRows);
    const monthName = MONTHS[period.monthIndex];
    const monthShort = MONTH_SHORT[period.monthIndex];
    const approval = approvalDateParts(period.year, period.monthIndex);
    const detailTotal = detailRows.reduce((sum, row) => sum + row.teamAmount + row.zonalAmount + row.rhoAmount + row.assistantAmount, 0);
    const summaryTotal = summaryRows.reduce((sum, row) => sum + row.amount, 0);
    const difference = Math.abs(detailTotal - summaryTotal);

    return {
      detailRows,
      summaryRows,
      assistantLabel,
      detailTotal,
      summaryTotal,
      difference,
      total: summaryTotal,
      monthIndex: period.monthIndex,
      monthName,
      monthShort,
      year: period.year,
      approval,
    };
  }

  function updateSummary(data) {
    elements.summaryMonth.textContent = `${data.monthName} ${data.year}`;
    elements.summaryOutlets.textContent = formatInteger(data.detailRows.length);
    elements.summaryRecipients.textContent = formatInteger(data.summaryRows.length);
    elements.summaryTotal.textContent = `${formatInteger(data.total)} BDT`;
    elements.approvalDate.textContent = `${data.approval.full} · Ref. ${data.approval.reference}`;
  }

  async function handleFile(file) {
    state.file = null;
    state.data = null;
    state.valid = false;
    elements.generateButton.disabled = true;
    resetSummary();

    if (!file) {
      elements.fileName.textContent = "No Excel file selected";
      setChip("No file selected", "idle");
      setStatus("Select an Excel file to begin.");
      return;
    }
    if (!/\.xlsx$/i.test(file.name)) {
      elements.fileName.textContent = file.name;
      setChip("Unsupported file", "bad");
      setStatus("Please select an Excel .xlsx file.", "bad");
      return;
    }

    elements.fileName.textContent = file.name;
    setChip("Reading workbook", "warn");
    setStatus("Reading the workbook and checking Appendix A against Appendix B…");
    setBusy(true);
    try {
      const workbook = await parseWorkbookFile(file);
      const data = extractDashboardData(workbook);
      state.file = file;
      state.data = data;
      updateSummary(data);
      if (data.difference > 0.5) {
        state.valid = false;
        setChip("Totals do not match", "bad");
        setStatus(`Appendix A totals ${formatInteger(data.detailTotal)} BDT, but Appendix B totals ${formatInteger(data.summaryTotal)} BDT. Correct the Excel file before generating the approval PDF.`, "bad");
      } else {
        state.valid = true;
        setChip("Workbook ready", "good");
        setStatus(`Workbook ready. ${formatInteger(data.detailRows.length)} outlets and ${formatInteger(data.summaryRows.length)} recipients reconcile to ${formatInteger(data.total)} BDT.`, "good");
      }
    } catch (error) {
      console.error(error);
      setChip("Workbook error", "bad");
      setStatus(error instanceof Error ? error.message : "The workbook could not be read.", "bad");
    } finally {
      setBusy(false);
    }
  }

  async function assetBytes(path) {
    if (!assetCache.has(path)) {
      assetCache.set(path, fetch(path).then(async (response) => {
        if (!response.ok) throw new Error(`Required dashboard asset is unavailable: ${path}`);
        return new Uint8Array(await response.arrayBuffer());
      }));
    }
    return assetCache.get(path);
  }

  function topToY(page, top, height = 0) {
    return page.getHeight() - top - height;
  }

  function coverTop(page, x, top, width, height, color) {
    page.drawRectangle({ x, y: topToY(page, top, height), width, height, color });
  }

  function textTop(page, text, x, top, size, font, color, options = {}) {
    const width = font.widthOfTextAtSize(text, size);
    let drawX = x;
    if (options.align === "center") drawX = x + (options.width - width) / 2;
    if (options.align === "right") drawX = x + options.width - width;
    page.drawText(text, { x: drawX, y: topToY(page, top, size), size, font, color });
    return width;
  }

  function drawRichWrappedText(page, segments, options) {
    const {
      x, top, maxWidth, size, lineHeight, color, highlightColor,
    } = options;
    let cursorX = x;
    let lineTop = top;
    let pendingSpace = 0;
    let pendingHighlight = false;

    const nextLine = () => {
      cursorX = x;
      lineTop += lineHeight;
      pendingSpace = 0;
      pendingHighlight = false;
    };

    segments.forEach((segment) => {
      const font = segment.font;
      const segmentSize = segment.size || size;
      const tokens = String(segment.text).split(/(\n|\s+)/).filter((token) => token !== "");
      tokens.forEach((token) => {
        if (token === "\n") {
          nextLine();
          return;
        }
        if (/^\s+$/.test(token)) {
          pendingSpace = font.widthOfTextAtSize(" ", segmentSize);
          pendingHighlight = Boolean(segment.highlight);
          return;
        }
        const tokenWidth = font.widthOfTextAtSize(token, segmentSize);
        if (cursorX > x && cursorX + pendingSpace + tokenWidth > x + maxWidth) nextLine();
        if (pendingSpace && cursorX > x) {
          if (pendingHighlight) {
            page.drawRectangle({
              x: cursorX,
              y: topToY(page, lineTop + 1, segmentSize + 2),
              width: pendingSpace,
              height: segmentSize + 2,
              color: highlightColor,
            });
          }
          cursorX += pendingSpace;
        }
        if (segment.highlight) {
          page.drawRectangle({
            x: cursorX - 0.7,
            y: topToY(page, lineTop + 1, segmentSize + 2),
            width: tokenWidth + 1.4,
            height: segmentSize + 2,
            color: highlightColor,
          });
        }
        page.drawText(token, {
          x: cursorX,
          y: topToY(page, lineTop, segmentSize),
          size: segmentSize,
          font,
          color,
        });
        cursorX += tokenWidth;
        pendingSpace = 0;
        pendingHighlight = false;
      });
    });

    return lineTop + lineHeight;
  }

  function drawDateField(page, data, fonts, colors) {
    const x = 67.2;
    const top = 74.1;
    const dayText = String(data.approval.day).padStart(2, "0");
    const suffix = ordinalSuffix(data.approval.day);
    const tail = ` ${data.approval.month} ${data.approval.year}`;
    const dayWidth = fonts.regular.widthOfTextAtSize(dayText, 11);
    const suffixWidth = fonts.regular.widthOfTextAtSize(suffix, 6.4);
    const tailWidth = fonts.regular.widthOfTextAtSize(tail, 11);
    const totalWidth = dayWidth + suffixWidth + tailWidth;
    coverTop(page, 66.3, 72.7, 111, 18.2, colors.white);
    page.drawRectangle({ x: x - 0.8, y: topToY(page, 73.9, 14.2), width: totalWidth + 1.6, height: 14.2, color: colors.yellow });
    textTop(page, dayText, x, top + 1.2, 11, fonts.regular, colors.black);
    textTop(page, suffix, x + dayWidth, top - 0.3, 6.4, fonts.regular, colors.black);
    textTop(page, tail, x + dayWidth + suffixWidth, top + 1.2, 11, fonts.regular, colors.black);
  }

  function patchApprovalPages(pdfDoc, data, fonts, signatures, colors) {
    const page1 = pdfDoc.getPage(0);
    const page2 = pdfDoc.getPage(1);
    const amount = `${formatInteger(data.total)}/=`;
    const monthYear = `${data.monthName} ${data.year}`;

    drawDateField(page1, data, fonts, colors);

    coverTop(page1, 357.8, 72.7, 47.4, 18.2, colors.white);
    const refWidth = fonts.regular.widthOfTextAtSize(data.approval.reference, 11);
    page1.drawRectangle({ x: 358.8, y: topToY(page1, 73.9, 14.2), width: refWidth + 1.6, height: 14.2, color: colors.yellow });
    textTop(page1, data.approval.reference, 359.4, 75.3, 11, fonts.regular, colors.black);

    coverTop(page1, 35.2, 134.4, 525, 34.8, colors.white);
    drawRichWrappedText(page1, [
      { text: "SUBJECT: Approval for Incentive Disbursement amount of ", font: fonts.bold },
      { text: amount, font: fonts.bold, highlight: true },
      { text: " BDT to Operations Team for New Outlet Opening ", font: fonts.bold },
      { text: `(${monthYear})`, font: fonts.bold, highlight: true },
      { text: ".", font: fonts.bold },
    ], { x: 36.1, top: 138.5, maxWidth: 523, size: 11, lineHeight: 14.5, color: colors.black, highlightColor: colors.yellow });

    coverTop(page1, 35.2, 177.7, 525, 59, colors.white);
    drawRichWrappedText(page1, [
      { text: "In accordance with the approved incentive policy for motivating and strengthening the ongoing expansion initiatives of the Operations Departments, the performance and eligibility for incentive disbursement for the month of ", font: fonts.regular },
      { text: monthYear, font: fonts.regular, highlight: true },
      { text: " have been reviewed, verified, and evaluated as per the approved guidelines and supporting details enclosed in Appendix-A & B.", font: fonts.regular },
    ], { x: 36.1, top: 181.8, maxWidth: 523, size: 11, lineHeight: 13.45, color: colors.black, highlightColor: colors.yellow });

    coverTop(page1, 35.2, 728.8, 525, 35, colors.white);
    drawRichWrappedText(page1, [
      { text: "The approved incentive amounts shall be disbursed among the eligible employees listed in Appendix-A and B including the associate team for ", font: fonts.regular },
      { text: `${data.monthName} -${data.year}`, font: fonts.regular, highlight: true },
      { text: ", through MFS/bank transfer.", font: fonts.regular },
    ], { x: 36.1, top: 733.2, maxWidth: 523, size: 11, lineHeight: 13.45, color: colors.black, highlightColor: colors.yellow });

    coverTop(page2, 35.2, 92.0, 525, 43.5, colors.white);
    drawRichWrappedText(page2, [
      { text: "In view of the above, ", font: fonts.regular },
      { text: "approval for Incentive Disbursement amount of ", font: fonts.bold },
      { text: amount, font: fonts.bold, highlight: true },
      { text: " BDT", font: fonts.bold },
      { text: " is hereby recommended for incentive disbursement as per the details stated above and the attached appendices.", font: fonts.regular },
    ], { x: 36.1, top: 96.0, maxWidth: 523, size: 11, lineHeight: 13.45, color: colors.black, highlightColor: colors.yellow });

    const aftabWidth = 69;
    const aftabHeight = aftabWidth * signatures.aftab.height / signatures.aftab.width;
    page2.drawImage(signatures.aftab, { x: 67, y: 572.5, width: aftabWidth, height: aftabHeight, opacity: 0.96 });

    const saifulWidth = 72;
    const saifulHeight = saifulWidth * signatures.saiful.height / signatures.saiful.width;
    page2.drawImage(signatures.saiful, { x: 194, y: 568.2, width: saifulWidth, height: saifulHeight, opacity: 0.96 });
  }

  function fitTextSize(text, font, preferred, maxWidth, minimum = 4.8) {
    let size = preferred;
    while (size > minimum && font.widthOfTextAtSize(String(text), size) > maxWidth) size -= 0.2;
    return Math.max(minimum, size);
  }

  function wrappedLines(text, font, size, maxWidth, maxLines = 3) {
    const words = String(text || "").trim().split(/\s+/).filter(Boolean);
    if (!words.length) return [""];
    const lines = [];
    let current = "";
    words.forEach((word) => {
      const next = current ? `${current} ${word}` : word;
      if (!current || font.widthOfTextAtSize(next, size) <= maxWidth) current = next;
      else {
        lines.push(current);
        current = word;
      }
    });
    if (current) lines.push(current);
    if (lines.length <= maxLines) return lines;
    const kept = lines.slice(0, maxLines);
    kept[maxLines - 1] = `${kept[maxLines - 1].replace(/\.*$/, "")}…`;
    return kept;
  }

  function drawCell(page, value, x, top, width, height, options) {
    const {
      font, size = 7, align = "left", color, borderColor, fillColor,
      padding = 2.2, wrap = false, maxLines = 3, borderWidth = 0.55,
    } = options;
    page.drawRectangle({
      x,
      y: topToY(page, top, height),
      width,
      height,
      color: fillColor,
      borderColor,
      borderWidth,
    });
    const text = value === null || value === undefined ? "" : String(value);
    if (!text) return;
    let actualSize = size;
    let lines;
    if (wrap) {
      lines = wrappedLines(text, font, actualSize, Math.max(1, width - padding * 2), maxLines);
      while (actualSize > 4.6 && lines.length > maxLines) {
        actualSize -= 0.2;
        lines = wrappedLines(text, font, actualSize, Math.max(1, width - padding * 2), maxLines);
      }
    } else {
      actualSize = fitTextSize(text, font, actualSize, Math.max(1, width - padding * 2));
      lines = [text];
    }
    const lineHeight = actualSize + 1.4;
    const blockHeight = lines.length * lineHeight - 1.4;
    let lineTop = top + (height - blockHeight) / 2 - 0.5;
    lines.forEach((line) => {
      const lineWidth = font.widthOfTextAtSize(line, actualSize);
      let textX = x + padding;
      if (align === "center") textX = x + (width - lineWidth) / 2;
      if (align === "right") textX = x + width - padding - lineWidth;
      page.drawText(line, {
        x: textX,
        y: topToY(page, lineTop, actualSize),
        size: actualSize,
        font,
        color,
      });
      lineTop += lineHeight;
    });
  }

  function drawAppendixHeading(page, label, continuation, fonts, colors, top = 36) {
    const text = continuation ? `${label} (continued)` : label;
    const size = 16;
    const width = fonts.bold.widthOfTextAtSize(text, size);
    page.drawRectangle({ x: 36, y: topToY(page, top - 1, size + 6), width: width + 3, height: size + 6, color: colors.yellow });
    textTop(page, text, 36, top, size, fonts.bold, colors.black);
    page.drawLine({ start: { x: 36, y: topToY(page, top + size + 1) }, end: { x: 36 + width, y: topToY(page, top + size + 1) }, thickness: 0.8, color: colors.black });
    return top + 28;
  }

  function formatAppendixDate(value) {
    const date = excelDate(value);
    if (!date) return asText(value);
    return `${date.getUTCDate()}-${MONTH_SHORT[date.getUTCMonth()]}-${String(date.getUTCFullYear()).slice(-2)}`;
  }

  function enrichRhoGroups(rows) {
    let current = null;
    let key = 0;
    return rows.map((row) => {
      if (row.rhoName || row.rhoQuantity || row.rhoAmount) {
        current = {
          key: key += 1,
          name: row.rhoName,
          quantity: row.rhoQuantity,
          amount: row.rhoAmount,
        };
      }
      return { ...row, rhoGroup: current || { key: `blank-${row.sl}`, name: "", quantity: 0, amount: 0 } };
    });
  }

  function drawAppendixAChunk(page, data, rows, top, isLast, fonts, colors) {
    const x = 21;
    const widths = [16, 48, 31, 89, 62, 36, 67, 36, 70, 24, 35, 39];
    const totalWidth = widths.reduce((sum, value) => sum + value, 0);
    const rowHeight = 15;
    const titleHeight = 20;
    const groupHeight = 38;
    const headerHeight = 32;
    const totalHeight = 17;
    let cursor = top;

    drawCell(page, `Total Incentive Summary For New Outlet Opening of Operations : ${data.monthShort}-${data.year}`, x, cursor, totalWidth, titleHeight, {
      font: fonts.bold, size: 9.2, align: "left", color: colors.black, borderColor: colors.black, fillColor: colors.white, padding: 2.4,
    });
    cursor += titleHeight;

    const groups = [
      { label: "New Outlet Details", start: 0, span: 4 },
      { label: "Outlet Team Member", start: 4, span: 2 },
      { label: "Zonal Manager", start: 6, span: 2 },
      { label: "RHO", start: 8, span: 3 },
      { label: data.assistantLabel, start: 11, span: 1 },
    ];
    groups.forEach((group) => {
      const groupX = x + widths.slice(0, group.start).reduce((sum, value) => sum + value, 0);
      const groupWidth = widths.slice(group.start, group.start + group.span).reduce((sum, value) => sum + value, 0);
      drawCell(page, group.label, groupX, cursor, groupWidth, groupHeight, {
        font: fonts.bold, size: group.start === 11 ? 5.6 : 8.2, align: "center", wrap: true, maxLines: group.start === 11 ? 5 : 3,
        color: colors.black, borderColor: colors.black, fillColor: colors.white,
      });
    });
    cursor += groupHeight;

    const headers = [
      "SL", "Inauguration Date", "Code", "Outlet Name", "Outlet Team Member", "Incentive Amount",
      "Zonal Name", "Incentive Amount", "Expansion By", "Outlet Quantity", "Incentive Amount", "Incentive Amount",
    ];
    let headerX = x;
    headers.forEach((header, index) => {
      drawCell(page, header, headerX, cursor, widths[index], headerHeight, {
        font: fonts.bold, size: 6.6, align: "center", wrap: true, maxLines: 3,
        color: colors.black, borderColor: colors.black, fillColor: colors.white,
      });
      headerX += widths[index];
    });
    cursor += headerHeight;

    rows.forEach((row) => {
      const values = [
        row.sl,
        formatAppendixDate(row.date),
        row.code,
        row.outlet,
        row.teamMember,
        formatInteger(row.teamAmount),
        row.zonalName,
        formatInteger(row.zonalAmount),
      ];
      let cellX = x;
      values.forEach((value, index) => {
        drawCell(page, value, cellX, cursor, widths[index], rowHeight, {
          font: fonts.regular,
          size: index === 3 || index === 4 || index === 6 ? 6.1 : 6.5,
          align: [0, 1, 2, 5, 7].includes(index) ? "center" : "left",
          color: colors.black,
          borderColor: colors.black,
          fillColor: colors.white,
        });
        cellX += widths[index];
      });
      const assistantX = x + widths.slice(0, 11).reduce((sum, value) => sum + value, 0);
      drawCell(page, formatInteger(row.assistantAmount), assistantX, cursor, widths[11], rowHeight, {
        font: fonts.regular, size: 6.5, align: "center", color: colors.black, borderColor: colors.black, fillColor: colors.white,
      });
      cursor += rowHeight;
    });

    const rhoX = x + widths.slice(0, 8).reduce((sum, value) => sum + value, 0);
    let groupStart = 0;
    while (groupStart < rows.length) {
      const groupKey = rows[groupStart].rhoGroup.key;
      let groupEnd = groupStart + 1;
      while (groupEnd < rows.length && rows[groupEnd].rhoGroup.key === groupKey) groupEnd += 1;
      const spanHeight = (groupEnd - groupStart) * rowHeight;
      const groupTop = top + titleHeight + groupHeight + headerHeight + groupStart * rowHeight;
      const group = rows[groupStart].rhoGroup;
      drawCell(page, group.name, rhoX, groupTop, widths[8], spanHeight, {
        font: fonts.regular, size: 6.2, align: "center", wrap: true, maxLines: 3,
        color: colors.black, borderColor: colors.black, fillColor: colors.white,
      });
      drawCell(page, group.quantity ? formatInteger(group.quantity) : "", rhoX + widths[8], groupTop, widths[9], spanHeight, {
        font: fonts.regular, size: 6.5, align: "center", color: colors.black, borderColor: colors.black, fillColor: colors.white,
      });
      drawCell(page, formatInteger(group.amount), rhoX + widths[8] + widths[9], groupTop, widths[10], spanHeight, {
        font: fonts.regular, size: 6.5, align: "center", color: colors.black, borderColor: colors.black, fillColor: colors.white,
      });
      groupStart = groupEnd;
    }

    if (isLast) {
      const totals = {
        team: data.detailRows.reduce((sum, row) => sum + row.teamAmount, 0),
        zonal: data.detailRows.reduce((sum, row) => sum + row.zonalAmount, 0),
        quantity: data.detailRows.reduce((sum, row) => sum + (row.rhoName ? row.rhoQuantity : 0), 0),
        rho: data.detailRows.reduce((sum, row) => sum + row.rhoAmount, 0),
        assistant: data.detailRows.reduce((sum, row) => sum + row.assistantAmount, 0),
      };
      let totalX = x;
      const totalCells = [
        { value: "Grand Total", span: 4 },
        { value: "", span: 1 },
        { value: formatInteger(totals.team), span: 1 },
        { value: "", span: 1 },
        { value: formatInteger(totals.zonal), span: 1 },
        { value: "", span: 1 },
        { value: formatInteger(totals.quantity), span: 1 },
        { value: formatInteger(totals.rho), span: 1 },
        { value: formatInteger(totals.assistant), span: 1 },
      ];
      let widthCursor = 0;
      totalCells.forEach((cell) => {
        const cellWidth = widths.slice(widthCursor, widthCursor + cell.span).reduce((sum, value) => sum + value, 0);
        drawCell(page, cell.value, totalX, cursor, cellWidth, totalHeight, {
          font: fonts.bold, size: 7, align: "center", color: colors.black, borderColor: colors.black, fillColor: colors.white,
        });
        totalX += cellWidth;
        widthCursor += cell.span;
      });
      cursor += totalHeight;
    }
    return cursor;
  }

  function drawAppendixA(pdfDoc, data, fonts, colors) {
    const rows = enrichRhoGroups(data.detailRows);
    const rowHeight = 15;
    const fixedNoTotal = 20 + 38 + 32;
    const totalHeight = 17;
    let cursor = 0;
    let page = null;
    let top = 36;
    let pageIndex = 0;

    while (cursor < rows.length) {
      page = pdfDoc.addPage(A4);
      top = drawAppendixHeading(page, "Appendix-A", pageIndex > 0, fonts, colors, 36);
      const remaining = rows.length - cursor;
      const capacityWithTotal = Math.max(1, Math.floor((PAGE_BOTTOM - top - fixedNoTotal - totalHeight) / rowHeight));
      const capacityWithoutTotal = Math.max(1, Math.floor((PAGE_BOTTOM - top - fixedNoTotal) / rowHeight));
      const isLast = remaining <= capacityWithTotal;
      const count = isLast ? remaining : Math.min(remaining, capacityWithoutTotal);
      top = drawAppendixAChunk(page, data, rows.slice(cursor, cursor + count), top, isLast, fonts, colors);
      cursor += count;
      pageIndex += 1;
    }
    return { page, top };
  }

  function estimatedAppendixBHeight(rowCount, includeHeading = true) {
    return (includeHeading ? 28 : 0) + 44 + 18 + rowCount * 14 + 18;
  }

  function drawAppendixBChunk(page, data, rows, top, isLast, fonts, colors) {
    const x = 36;
    const widths = [28, 195, 48, 170, 82];
    const totalWidth = widths.reduce((sum, value) => sum + value, 0);
    const titleHeight = 44;
    const headerHeight = 18;
    const rowHeight = 14;
    const totalHeight = 18;
    let cursor = top;

    drawCell(page, `Total Incentive Summary For New Outlet Opening of Operations :\n${data.monthName}-${data.year}`, x, cursor, totalWidth, titleHeight, {
      font: fonts.bold, size: 13, align: "center", wrap: true, maxLines: 2,
      color: colors.black, borderColor: colors.black, fillColor: colors.white,
    });
    cursor += titleHeight;

    const headers = ["SL", "Expansion By", "ID", "Designation", "Incentive Amount"];
    let headerX = x;
    headers.forEach((header, index) => {
      drawCell(page, header, headerX, cursor, widths[index], headerHeight, {
        font: fonts.bold, size: 9.2, align: "center", color: colors.black, borderColor: colors.black, fillColor: colors.white,
      });
      headerX += widths[index];
    });
    cursor += headerHeight;

    rows.forEach((row) => {
      const values = [row.sl, row.name, row.id, row.designation, formatInteger(row.amount)];
      let cellX = x;
      values.forEach((value, index) => {
        drawCell(page, value, cellX, cursor, widths[index], rowHeight, {
          font: fonts.regular,
          size: index === 1 || index === 3 ? 8.2 : 8.5,
          align: index === 1 ? "left" : "center",
          color: colors.black,
          borderColor: colors.black,
          fillColor: colors.white,
        });
        cellX += widths[index];
      });
      cursor += rowHeight;
    });

    if (isLast) {
      const labelWidth = widths.slice(0, 4).reduce((sum, value) => sum + value, 0);
      drawCell(page, "Grand Total", x, cursor, labelWidth, totalHeight, {
        font: fonts.bold, size: 13, align: "center", color: colors.black, borderColor: colors.black, fillColor: colors.white,
      });
      drawCell(page, formatInteger(data.total), x + labelWidth, cursor, widths[4], totalHeight, {
        font: fonts.bold, size: 13, align: "center", color: colors.black, borderColor: colors.black, fillColor: colors.white,
      });
      cursor += totalHeight;
    }
    return cursor;
  }

  function drawAppendixB(pdfDoc, data, fonts, colors, start) {
    const rowHeight = 14;
    const fixedNoTotal = 44 + 18;
    const totalHeight = 18;
    let page = start.page;
    let top = start.top + 14;
    let cursor = 0;
    let continuation = false;

    const fullHeight = estimatedAppendixBHeight(data.summaryRows.length, true);
    if (!page || top + fullHeight > PAGE_BOTTOM) {
      page = pdfDoc.addPage(A4);
      top = 36;
    }

    while (cursor < data.summaryRows.length) {
      top = drawAppendixHeading(page, "Appendix-B", continuation, fonts, colors, top);
      const remaining = data.summaryRows.length - cursor;
      const capacityWithTotal = Math.max(1, Math.floor((PAGE_BOTTOM - top - fixedNoTotal - totalHeight) / rowHeight));
      const capacityWithoutTotal = Math.max(1, Math.floor((PAGE_BOTTOM - top - fixedNoTotal) / rowHeight));
      const isLast = remaining <= capacityWithTotal;
      const count = isLast ? remaining : Math.min(remaining, capacityWithoutTotal);
      top = drawAppendixBChunk(page, data, data.summaryRows.slice(cursor, cursor + count), top, isLast, fonts, colors);
      cursor += count;
      if (cursor < data.summaryRows.length) {
        page = pdfDoc.addPage(A4);
        top = 36;
        continuation = true;
      }
    }
    return { page, top };
  }

  function patchPageNumbers(pdfDoc, fonts, colors) {
    const total = pdfDoc.getPageCount();
    const pages = pdfDoc.getPages();
    const totalText = String(total).padStart(2, "0");

    coverTop(pages[0], 349.8, 102.9, 15.5, 17.5, colors.white);
    textTop(pages[0], totalText, 350.5, 105.3, 11, fonts.regular, colors.black);

    pages.forEach((page, index) => {
      coverTop(page, 478, 777.5, 82, 24, colors.white);
      const label = `Page ${index + 1} of ${total}`;
      textTop(page, label, 478, 781.4, 11, fonts.regular, colors.black, { align: "right", width: 67 });
    });
  }

  const CANVAS_SCALE = 2.2;

  function createPdfCanvas() {
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(A4[0] * CANVAS_SCALE);
    canvas.height = Math.round(A4[1] * CANVAS_SCALE);
    const context = canvas.getContext("2d");
    if (!context) throw new Error("This browser cannot create the PDF page canvas.");
    context.scale(CANVAS_SCALE, CANVAS_SCALE);
    context.textBaseline = "top";
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, A4[0], A4[1]);
    return { canvas, context };
  }

  function canvasFont(context, size, bold = false) {
    context.font = `${bold ? 700 : 400} ${size}px "ApprovalSans", Arial, sans-serif`;
  }

  async function ensureCanvasFonts() {
    if (!document.fonts || typeof document.fonts.load !== "function") return;
    await Promise.all([
      document.fonts.load('400 12px "ApprovalSans"'),
      document.fonts.load('700 12px "ApprovalSans"'),
    ]);
  }

  function canvasTextWidth(context, text, size, bold = false) {
    canvasFont(context, size, bold);
    return context.measureText(String(text)).width;
  }

  function fillWhite(context, x, top, width, height) {
    context.save();
    context.fillStyle = "#ffffff";
    context.fillRect(x, top, width, height);
    context.restore();
  }

  function drawCanvasText(context, text, x, top, size, bold = false, align = "left", width = 0) {
    context.save();
    canvasFont(context, size, bold);
    context.fillStyle = "#000000";
    context.textAlign = align;
    let drawX = x;
    if (align === "center") drawX = x + width / 2;
    if (align === "right") drawX = x + width;
    context.fillText(String(text), drawX, top);
    context.restore();
  }

  function drawRichCanvas(context, segments, options) {
    const { x, top, maxWidth, size, lineHeight } = options;
    let cursorX = x;
    let lineTop = top;
    let pendingSpace = 0;
    let pendingHighlight = false;

    const nextLine = () => {
      cursorX = x;
      lineTop += lineHeight;
      pendingSpace = 0;
      pendingHighlight = false;
    };

    segments.forEach((segment) => {
      const segmentSize = segment.size || size;
      const tokens = String(segment.text).split(/(\n|\s+)/).filter((token) => token !== "");
      tokens.forEach((token) => {
        canvasFont(context, segmentSize, Boolean(segment.bold));
        if (token === "\n") {
          nextLine();
          return;
        }
        if (/^\s+$/.test(token)) {
          pendingSpace = context.measureText(" ").width;
          pendingHighlight = Boolean(segment.highlight);
          return;
        }
        const tokenWidth = context.measureText(token).width;
        if (cursorX > x && cursorX + pendingSpace + tokenWidth > x + maxWidth) nextLine();
        if (pendingSpace && cursorX > x) {
          if (pendingHighlight) {
            context.fillStyle = "#ffff00";
            context.fillRect(cursorX, lineTop - 0.5, pendingSpace, segmentSize + 2.5);
          }
          cursorX += pendingSpace;
        }
        if (segment.highlight) {
          context.fillStyle = "#ffff00";
          context.fillRect(cursorX - 0.7, lineTop - 0.5, tokenWidth + 1.4, segmentSize + 2.5);
        }
        context.fillStyle = "#000000";
        context.fillText(token, cursorX, lineTop);
        cursorX += tokenWidth;
        pendingSpace = 0;
        pendingHighlight = false;
      });
    });
    return lineTop + lineHeight;
  }

  function drawCanvasDateField(context, data) {
    const x = 67.2;
    const top = 75.1;
    const dayText = String(data.approval.day).padStart(2, "0");
    const suffix = ordinalSuffix(data.approval.day);
    const tail = ` ${data.approval.month} ${data.approval.year}`;
    fillWhite(context, 66.4, 73.6, 110, 15.4);
    const dayWidth = canvasTextWidth(context, dayText, 11);
    const suffixWidth = canvasTextWidth(context, suffix, 6.4);
    const tailWidth = canvasTextWidth(context, tail, 11);
    context.fillStyle = "#ffff00";
    context.fillRect(x - 0.7, top - 0.4, dayWidth + suffixWidth + tailWidth + 1.4, 13.8);
    drawCanvasText(context, dayText, x, top, 11);
    drawCanvasText(context, suffix, x + dayWidth, top - 1.5, 6.4);
    drawCanvasText(context, tail, x + dayWidth + suffixWidth, top, 11);
  }

  function drawCanvasFooter(context, pageNumber, totalPages) {
    fillWhite(context, 475, 778, 84, 19);
    drawCanvasText(context, `Page ${pageNumber} of ${totalPages}`, 478, 781.2, 11, false, "right", 67);
  }

  function renderApprovalCanvases(data, images, totalPages) {
    const amount = `${formatInteger(data.total)}/=`;
    const monthYear = `${data.monthName} ${data.year}`;

    const first = createPdfCanvas();
    first.context.drawImage(images.page1, 0, 0, A4[0], A4[1]);
    drawCanvasDateField(first.context, data);

    fillWhite(first.context, 357.8, 73.6, 47.4, 15.4);
    const refWidth = canvasTextWidth(first.context, data.approval.reference, 11);
    first.context.fillStyle = "#ffff00";
    first.context.fillRect(358.8, 74.8, refWidth + 1.4, 13.8);
    drawCanvasText(first.context, data.approval.reference, 359.4, 75.1, 11);

    fillWhite(first.context, 349.5, 103.7, 16.5, 15.2);
    drawCanvasText(first.context, String(totalPages).padStart(2, "0"), 350.6, 105.2, 11);

    fillWhite(first.context, 35.2, 135.4, 525, 32.7);
    drawRichCanvas(first.context, [
      { text: "SUBJECT: Approval for Incentive Disbursement amount of ", bold: true },
      { text: amount, bold: true, highlight: true },
      { text: " BDT to Operations Team for New Outlet Opening ", bold: true },
      { text: `(${monthYear})`, bold: true, highlight: true },
      { text: ".", bold: true },
    ], { x: 36.1, top: 138.2, maxWidth: 523, size: 11, lineHeight: 14.5 });

    fillWhite(first.context, 35.2, 178.6, 525, 57.3);
    drawRichCanvas(first.context, [
      { text: "In accordance with the approved incentive policy for motivating and strengthening the ongoing expansion initiatives of the Operations Departments, the performance and eligibility for incentive disbursement for the month of " },
      { text: monthYear, highlight: true },
      { text: " have been reviewed, verified, and evaluated as per the approved guidelines and supporting details enclosed in Appendix-A & B." },
    ], { x: 36.1, top: 181.4, maxWidth: 523, size: 11, lineHeight: 13.45 });

    fillWhite(first.context, 35.2, 729.5, 525, 33.1);
    drawRichCanvas(first.context, [
      { text: "The approved incentive amounts shall be disbursed among the eligible employees listed in Appendix-A and B including the associate team for " },
      { text: `${data.monthName} -${data.year}`, highlight: true },
      { text: ", through MFS/bank transfer." },
    ], { x: 36.1, top: 732.7, maxWidth: 523, size: 11, lineHeight: 13.45 });
    drawCanvasFooter(first.context, 1, totalPages);

    const second = createPdfCanvas();
    second.context.drawImage(images.page2, 0, 0, A4[0], A4[1]);
    fillWhite(second.context, 35.2, 91.4, 525, 43.7);
    drawRichCanvas(second.context, [
      { text: "In view of the above, " },
      { text: "approval for Incentive Disbursement amount of ", bold: true },
      { text: amount, bold: true, highlight: true },
      { text: " BDT", bold: true },
      { text: " is hereby recommended for incentive disbursement as per the details stated above and the attached appendices." },
    ], { x: 36.1, top: 95.5, maxWidth: 523, size: 11, lineHeight: 13.45 });

    const aftabWidth = 69;
    const aftabHeight = aftabWidth * images.aftab.height / images.aftab.width;
    second.context.drawImage(images.aftab, 67, 842 - 572.5 - aftabHeight, aftabWidth, aftabHeight);
    const saifulWidth = 72;
    const saifulHeight = saifulWidth * images.saiful.height / images.saiful.width;
    second.context.drawImage(images.saiful, 194, 842 - 568.2 - saifulHeight, saifulWidth, saifulHeight);
    drawCanvasFooter(second.context, 2, totalPages);

    return [first, second];
  }

  function fitCanvasTextSize(context, text, preferred, maxWidth, bold = false, minimum = 4.8) {
    let size = preferred;
    while (size > minimum && canvasTextWidth(context, text, size, bold) > maxWidth) size -= 0.2;
    return Math.max(minimum, size);
  }

  function canvasWrappedLines(context, text, size, maxWidth, bold = false, maxLines = 3) {
    canvasFont(context, size, bold);
    const explicit = String(text || "").split("\n");
    const lines = [];
    explicit.forEach((part) => {
      const words = part.trim().split(/\s+/).filter(Boolean);
      if (!words.length) {
        lines.push("");
        return;
      }
      let current = "";
      words.forEach((word) => {
        const next = current ? `${current} ${word}` : word;
        if (!current || context.measureText(next).width <= maxWidth) current = next;
        else {
          lines.push(current);
          current = word;
        }
      });
      if (current) lines.push(current);
    });
    if (lines.length <= maxLines) return lines;
    const kept = lines.slice(0, maxLines);
    let last = kept[maxLines - 1];
    while (last && canvasTextWidth(context, `${last}…`, size, bold) > maxWidth) last = last.slice(0, -1);
    kept[maxLines - 1] = `${last}…`;
    return kept;
  }

  function drawCanvasCell(context, value, x, top, width, height, options) {
    const {
      size = 7, bold = false, align = "left", padding = 2.2,
      wrap = false, maxLines = 3, lineWidth = 0.55,
    } = options;
    context.save();
    context.fillStyle = "#ffffff";
    context.fillRect(x, top, width, height);
    context.strokeStyle = "#000000";
    context.lineWidth = lineWidth;
    context.strokeRect(x, top, width, height);
    const text = value === null || value === undefined ? "" : String(value);
    if (!text) {
      context.restore();
      return;
    }

    let actualSize = size;
    let lines;
    if (wrap) {
      lines = canvasWrappedLines(context, text, actualSize, Math.max(1, width - padding * 2), bold, maxLines);
    } else {
      actualSize = fitCanvasTextSize(context, text, actualSize, Math.max(1, width - padding * 2), bold);
      lines = [text];
    }
    canvasFont(context, actualSize, bold);
    const lineHeight = actualSize + 1.4;
    const blockHeight = lines.length * lineHeight - 1.4;
    let lineTop = top + (height - blockHeight) / 2;
    context.beginPath();
    context.rect(x + 0.3, top + 0.3, width - 0.6, height - 0.6);
    context.clip();
    context.fillStyle = "#000000";
    lines.forEach((line) => {
      const lineWidthValue = context.measureText(line).width;
      let textX = x + padding;
      if (align === "center") textX = x + (width - lineWidthValue) / 2;
      if (align === "right") textX = x + width - padding - lineWidthValue;
      context.fillText(line, textX, lineTop);
      lineTop += lineHeight;
    });
    context.restore();
  }

  function drawCanvasAppendixHeading(context, label, continuation, top = 36) {
    const text = continuation ? `${label} (continued)` : label;
    const size = 16;
    const width = canvasTextWidth(context, text, size, true);
    context.fillStyle = "#ffff00";
    context.fillRect(36, top - 1, width + 3, size + 6);
    drawCanvasText(context, text, 36, top, size, true);
    context.strokeStyle = "#000000";
    context.lineWidth = 0.8;
    context.beginPath();
    context.moveTo(36, top + size + 1);
    context.lineTo(36 + width, top + size + 1);
    context.stroke();
    return top + 28;
  }

  function drawCanvasAppendixAChunk(context, data, rows, top, isLast) {
    const x = 21;
    const widths = [16, 48, 31, 89, 62, 36, 67, 36, 70, 24, 35, 39];
    const totalWidth = widths.reduce((sum, value) => sum + value, 0);
    const rowHeight = 15;
    const titleHeight = 20;
    const groupHeight = 38;
    const headerHeight = 32;
    const totalHeight = 17;
    let cursor = top;

    drawCanvasCell(context, `Total Incentive Summary For New Outlet Opening of Operations : ${data.monthShort}-${data.year}`, x, cursor, totalWidth, titleHeight, {
      bold: true, size: 9.2, align: "left", padding: 2.4,
    });
    cursor += titleHeight;

    const groups = [
      { label: "New Outlet Details", start: 0, span: 4 },
      { label: "Outlet Team Member", start: 4, span: 2 },
      { label: "Zonal Manager", start: 6, span: 2 },
      { label: "RHO", start: 8, span: 3 },
      { label: data.assistantLabel, start: 11, span: 1 },
    ];
    groups.forEach((group) => {
      const groupX = x + widths.slice(0, group.start).reduce((sum, value) => sum + value, 0);
      const groupWidth = widths.slice(group.start, group.start + group.span).reduce((sum, value) => sum + value, 0);
      drawCanvasCell(context, group.label, groupX, cursor, groupWidth, groupHeight, {
        bold: true, size: group.start === 11 ? 5.6 : 8.2, align: "center", wrap: true, maxLines: group.start === 11 ? 5 : 3,
      });
    });
    cursor += groupHeight;

    const headers = [
      "SL", "Inauguration Date", "Code", "Outlet Name", "Outlet Team Member", "Incentive Amount",
      "Zonal Name", "Incentive Amount", "Expansion By", "Outlet Quantity", "Incentive Amount", "Incentive Amount",
    ];
    let headerX = x;
    headers.forEach((header, index) => {
      drawCanvasCell(context, header, headerX, cursor, widths[index], headerHeight, {
        bold: true, size: 6.6, align: "center", wrap: true, maxLines: 3,
      });
      headerX += widths[index];
    });
    cursor += headerHeight;

    rows.forEach((row) => {
      const values = [
        row.sl,
        formatAppendixDate(row.date),
        row.code,
        row.outlet,
        row.teamMember,
        formatInteger(row.teamAmount),
        row.zonalName,
        formatInteger(row.zonalAmount),
      ];
      let cellX = x;
      values.forEach((value, index) => {
        drawCanvasCell(context, value, cellX, cursor, widths[index], rowHeight, {
          size: index === 3 || index === 4 || index === 6 ? 6.1 : 6.5,
          align: [0, 1, 2, 5, 7].includes(index) ? "center" : "left",
        });
        cellX += widths[index];
      });
      const assistantX = x + widths.slice(0, 11).reduce((sum, value) => sum + value, 0);
      drawCanvasCell(context, formatInteger(row.assistantAmount), assistantX, cursor, widths[11], rowHeight, { size: 6.5, align: "center" });
      cursor += rowHeight;
    });

    const rhoX = x + widths.slice(0, 8).reduce((sum, value) => sum + value, 0);
    let groupStart = 0;
    while (groupStart < rows.length) {
      const groupKey = rows[groupStart].rhoGroup.key;
      let groupEnd = groupStart + 1;
      while (groupEnd < rows.length && rows[groupEnd].rhoGroup.key === groupKey) groupEnd += 1;
      const spanHeight = (groupEnd - groupStart) * rowHeight;
      const groupTop = top + titleHeight + groupHeight + headerHeight + groupStart * rowHeight;
      const group = rows[groupStart].rhoGroup;
      drawCanvasCell(context, group.name, rhoX, groupTop, widths[8], spanHeight, { size: 6.2, align: "center", wrap: true, maxLines: 3 });
      drawCanvasCell(context, group.quantity ? formatInteger(group.quantity) : "", rhoX + widths[8], groupTop, widths[9], spanHeight, { size: 6.5, align: "center" });
      drawCanvasCell(context, formatInteger(group.amount), rhoX + widths[8] + widths[9], groupTop, widths[10], spanHeight, { size: 6.5, align: "center" });
      groupStart = groupEnd;
    }

    if (isLast) {
      const totals = {
        team: data.detailRows.reduce((sum, row) => sum + row.teamAmount, 0),
        zonal: data.detailRows.reduce((sum, row) => sum + row.zonalAmount, 0),
        quantity: data.detailRows.reduce((sum, row) => sum + (row.rhoName ? row.rhoQuantity : 0), 0),
        rho: data.detailRows.reduce((sum, row) => sum + row.rhoAmount, 0),
        assistant: data.detailRows.reduce((sum, row) => sum + row.assistantAmount, 0),
      };
      let totalX = x;
      const totalCells = [
        { value: "Grand Total", span: 4 }, { value: "", span: 1 }, { value: formatInteger(totals.team), span: 1 },
        { value: "", span: 1 }, { value: formatInteger(totals.zonal), span: 1 }, { value: "", span: 1 },
        { value: formatInteger(totals.quantity), span: 1 }, { value: formatInteger(totals.rho), span: 1 },
        { value: formatInteger(totals.assistant), span: 1 },
      ];
      let widthCursor = 0;
      totalCells.forEach((cell) => {
        const cellWidth = widths.slice(widthCursor, widthCursor + cell.span).reduce((sum, value) => sum + value, 0);
        drawCanvasCell(context, cell.value, totalX, cursor, cellWidth, totalHeight, { bold: true, size: 7, align: "center" });
        totalX += cellWidth;
        widthCursor += cell.span;
      });
      cursor += totalHeight;
    }
    return cursor;
  }

  function drawCanvasAppendixBChunk(context, data, rows, top, isLast) {
    const x = 36;
    const widths = [28, 195, 48, 170, 82];
    const totalWidth = widths.reduce((sum, value) => sum + value, 0);
    const titleHeight = 44;
    const headerHeight = 18;
    const rowHeight = 14;
    const totalHeight = 18;
    let cursor = top;
    drawCanvasCell(context, `Total Incentive Summary For New Outlet Opening of Operations :\n${data.monthName}-${data.year}`, x, cursor, totalWidth, titleHeight, {
      bold: true, size: 13, align: "center", wrap: true, maxLines: 2,
    });
    cursor += titleHeight;
    const headers = ["SL", "Expansion By", "ID", "Designation", "Incentive Amount"];
    let headerX = x;
    headers.forEach((header, index) => {
      drawCanvasCell(context, header, headerX, cursor, widths[index], headerHeight, { bold: true, size: 9.2, align: "center" });
      headerX += widths[index];
    });
    cursor += headerHeight;
    rows.forEach((row) => {
      const values = [row.sl, row.name, row.id, row.designation, formatInteger(row.amount)];
      let cellX = x;
      values.forEach((value, index) => {
        drawCanvasCell(context, value, cellX, cursor, widths[index], rowHeight, {
          size: index === 1 || index === 3 ? 8.2 : 8.5,
          align: index === 1 ? "left" : "center",
        });
        cellX += widths[index];
      });
      cursor += rowHeight;
    });
    if (isLast) {
      const labelWidth = widths.slice(0, 4).reduce((sum, value) => sum + value, 0);
      drawCanvasCell(context, "Grand Total", x, cursor, labelWidth, totalHeight, { bold: true, size: 13, align: "center" });
      drawCanvasCell(context, formatInteger(data.total), x + labelWidth, cursor, widths[4], totalHeight, { bold: true, size: 13, align: "center" });
      cursor += totalHeight;
    }
    return cursor;
  }

  function renderAppendixCanvases(data) {
    const pages = [];
    const rows = enrichRhoGroups(data.detailRows);
    const rowHeightA = 15;
    const fixedANoTotal = 20 + 38 + 32;
    const totalAHeight = 17;
    let rowCursor = 0;
    let currentPage = null;
    let top = 36;
    let pageIndex = 0;

    while (rowCursor < rows.length) {
      currentPage = createPdfCanvas();
      pages.push(currentPage);
      top = drawCanvasAppendixHeading(currentPage.context, "Appendix-A", pageIndex > 0, 36);
      const remaining = rows.length - rowCursor;
      const capacityWithTotal = Math.max(1, Math.floor((PAGE_BOTTOM - top - fixedANoTotal - totalAHeight) / rowHeightA));
      const capacityWithoutTotal = Math.max(1, Math.floor((PAGE_BOTTOM - top - fixedANoTotal) / rowHeightA));
      const isLast = remaining <= capacityWithTotal;
      const count = isLast ? remaining : Math.min(remaining, capacityWithoutTotal);
      top = drawCanvasAppendixAChunk(currentPage.context, data, rows.slice(rowCursor, rowCursor + count), top, isLast);
      rowCursor += count;
      pageIndex += 1;
    }

    const rowHeightB = 14;
    const fixedBNoTotal = 44 + 18;
    const totalBHeight = 18;
    let summaryCursor = 0;
    let continuation = false;
    top += 14;
    const fullBHeight = estimatedAppendixBHeight(data.summaryRows.length, true);
    if (!currentPage || top + fullBHeight > PAGE_BOTTOM) {
      currentPage = createPdfCanvas();
      pages.push(currentPage);
      top = 36;
    }

    while (summaryCursor < data.summaryRows.length) {
      top = drawCanvasAppendixHeading(currentPage.context, "Appendix-B", continuation, top);
      const remaining = data.summaryRows.length - summaryCursor;
      const capacityWithTotal = Math.max(1, Math.floor((PAGE_BOTTOM - top - fixedBNoTotal - totalBHeight) / rowHeightB));
      const capacityWithoutTotal = Math.max(1, Math.floor((PAGE_BOTTOM - top - fixedBNoTotal) / rowHeightB));
      const isLast = remaining <= capacityWithTotal;
      const count = isLast ? remaining : Math.min(remaining, capacityWithoutTotal);
      top = drawCanvasAppendixBChunk(currentPage.context, data, data.summaryRows.slice(summaryCursor, summaryCursor + count), top, isLast);
      summaryCursor += count;
      if (summaryCursor < data.summaryRows.length) {
        currentPage = createPdfCanvas();
        pages.push(currentPage);
        top = 36;
        continuation = true;
      }
    }
    return pages;
  }

  async function bytesToCanvasImage(bytes) {
    const blob = new Blob([bytes], { type: "image/png" });
    if (typeof createImageBitmap === "function") return createImageBitmap(blob);
    return new Promise((resolve, reject) => {
      const image = new Image();
      const url = URL.createObjectURL(blob);
      image.onload = () => {
        URL.revokeObjectURL(url);
        resolve(image);
      };
      image.onerror = () => {
        URL.revokeObjectURL(url);
        reject(new Error("A PDF image asset could not be loaded."));
      };
      image.src = url;
    });
  }

  async function canvasPngBytes(canvas) {
    if (typeof canvas.convertToBlob === "function") {
      const blob = await canvas.convertToBlob({ type: "image/png" });
      return new Uint8Array(await blob.arrayBuffer());
    }
    if (typeof canvas.toBlob === "function") {
      const blob = await new Promise((resolve, reject) => {
        canvas.toBlob((value) => value ? resolve(value) : reject(new Error("The PDF page image could not be created.")), "image/png");
      });
      return new Uint8Array(await blob.arrayBuffer());
    }
    if (typeof canvas.toBuffer === "function") {
      return new Uint8Array(canvas.toBuffer("image/png"));
    }
    throw new Error("This browser cannot export the PDF page image.");
  }

  async function buildPdf(data) {
    if (!window.PDFLib) throw new Error("The PDF generator did not load.");
    const { PDFDocument } = window.PDFLib;
    const [templatePage1Bytes, templatePage2Bytes, aftabBytes, saifulBytes] = await Promise.all([
      assetBytes("assets/approval-template-page-1.png"),
      assetBytes("assets/approval-template-page-2.png"),
      assetBytes("assets/signature-aftab.png"),
      assetBytes("assets/signature-saiful.png"),
    ]);

    await ensureCanvasFonts();
    const images = {
      page1: await bytesToCanvasImage(templatePage1Bytes),
      page2: await bytesToCanvasImage(templatePage2Bytes),
      aftab: await bytesToCanvasImage(aftabBytes),
      saiful: await bytesToCanvasImage(saifulBytes),
    };
    const appendixCanvases = renderAppendixCanvases(data);
    const totalPages = 2 + appendixCanvases.length;
    appendixCanvases.forEach((page, index) => drawCanvasFooter(page.context, index + 3, totalPages));
    const approvalCanvases = renderApprovalCanvases(data, images, totalPages);
    const allCanvases = [...approvalCanvases, ...appendixCanvases];

    const pdfDoc = await PDFDocument.create();
    for (const renderedPage of allCanvases) {
      const pngBytes = await canvasPngBytes(renderedPage.canvas);
      const png = await pdfDoc.embedPng(pngBytes);
      const page = pdfDoc.addPage(A4);
      page.drawImage(png, { x: 0, y: 0, width: A4[0], height: A4[1] });
    }

    pdfDoc.setTitle(`New Outlet Opening Approval - ${data.monthName} ${data.year}`);
    pdfDoc.setSubject("Operations incentive disbursement approval with Appendix A and Appendix B");
    pdfDoc.setCreator("New outlet opening approval format dashboard");
    return pdfDoc.save();
  }

  function downloadBytes(bytes, filename) {
    const blob = new Blob([bytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  async function generateApproval() {
    if (!state.valid || !state.data || state.busy) return;
    setBusy(true);
    setChip("Generating PDF", "warn");
    setStatus("Building the approval note, both appendices and the signed approval page…");
    try {
      const bytes = await buildPdf(state.data);
      const filename = `New_Outlet_Opening_Approval_${state.data.monthName}_${state.data.year}.pdf`;
      downloadBytes(bytes, filename);
      setChip("PDF generated", "good");
      setStatus(`${filename} has been generated and downloaded.`, "good");
    } catch (error) {
      console.error(error);
      setChip("Generation failed", "bad");
      setStatus(error instanceof Error ? error.message : "The PDF could not be generated.", "bad");
    } finally {
      setBusy(false);
    }
  }

  elements.themeToggle.addEventListener("click", () => {
    setTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark");
  });
  elements.selectFileButton.addEventListener("click", () => elements.fileInput.click());
  elements.fileInput.addEventListener("change", () => handleFile(elements.fileInput.files?.[0]));
  elements.generateButton.addEventListener("click", generateApproval);

  ["dragenter", "dragover"].forEach((eventName) => {
    elements.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      if (!state.busy) elements.dropZone.classList.add("is-dragging");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    elements.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropZone.classList.remove("is-dragging");
    });
  });
  elements.dropZone.addEventListener("drop", (event) => {
    if (state.busy) return;
    const file = event.dataTransfer?.files?.[0];
    if (file) handleFile(file);
  });

  initTheme();
  resetSummary();

  window.__approvalDashboard = {
    parseWorkbookFile,
    extractDashboardData,
    buildPdf,
  };
})();
