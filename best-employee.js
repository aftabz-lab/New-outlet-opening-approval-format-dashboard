(() => {
  "use strict";
  const api = window.__approvalDashboard;
  const tools = api.pdfTools;
  const A4 = [595.303937, 841.889764];
  const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const SHORT = MONTHS.map(month => month.slice(0, 3));
  const HEADERS = ["SL", "Criteria", "Employee ID", "Employee Name", "Designation", "Department", "Score %", "Final Amount"];
  const CRITERIA = [
    ["Best OM", "District wise"], ["Best ICMO", "Division Wise"], ["Best BSM", "Division Wise"],
    ["Employee of the month", "Zonal Wise"], ["Best Zonal Operations", "Region Wise"],
    ["Best Zonal ICMO", "Division Wise"], ["Best Stock Gap Reduce Outlet", "Division Wise"],
    ["Best Wastage Reduce Outlet", "Division Wise"], ["Best Merchandising Outlet (Display)", "Division Wise"],
    ["Best RHO of the month", "National"], ["Best Divisional ICMO of the month", "National"],
  ];
  const ALIASES = {
    criteria: ["Criteria", "Award Criteria", "Award Category"],
    id: ["Employee ID", "ID", "Staff ID"], name: ["Employee Name", "Name", "Staff Name"],
    designation: ["Designation"], department: ["Department", "Dept"],
    score: ["Score %", "Score%", "Score (%)", "Final Score %", "Total Score %"],
    amount: ["Final Amount", "Final Amount (BDT)", "Final Incentive Amount", "Final Incentive (BDT)"],
    source: ["Source File(s)", "Source File", "Source Files"],
    period: ["Reporting Month", "Award Month", "Month", "Reporting Period"],
  };
  const $ = id => document.getElementById(id);
  const ui = Object.fromEntries(["FileInput", "SelectFileButton", "GenerateButton", "DropZone", "FileName", "FileChip", "Status", "Period", "SummaryMonth", "SummaryRecipients", "SummaryCriteria", "SummaryTotal", "ApprovalDate"].map(key => [key, $("employee" + key)]));
  const state = { data: null, busy: false, request: 0 };
  const norm = value => String(value ?? "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
  const headerKey = value => norm(value).replace(/\(\s*bdt\s*\)/g, "").replace(/[().:%_\-\s]/g, "");
  const text = value => value === null || value === undefined ? "" : String(value).trim();
  const money = value => new Intl.NumberFormat("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 2 }).format(value);
  const cents = value => Math.round((value + Number.EPSILON) * 100);

  function numeric(value, label, row) {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    const source = text(value).replace(/,/g, "").replace(/\s*(?:BDT|TK\.?|৳)\s*/gi, "");
    if (!source || !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%?$/.test(source)) {
      throw new Error(`${label} is missing or invalid in worksheet row ${row}.`);
    }
    const result = Number(source.replace(/%$/, ""));
    if (!Number.isFinite(result)) throw new Error(`${label} is invalid in worksheet row ${row}.`);
    return result;
  }

  function locateHeaders(workbook) {
    const candidates = [];
    for (const sheet of workbook.sheets) {
      const limit = Math.min(40, sheet.rows.length);
      for (let start = 0; start < limit; start++) {
        for (let depth = 1; depth <= 4 && start + depth <= sheet.rows.length; depth++) {
          const windowRows = sheet.rows.slice(start, start + depth);
          const indices = {}, locations = {};
          for (const [key, names] of Object.entries(ALIASES)) {
            const allowed = new Set(names.map(headerKey));
            for (let offset = 0; offset < windowRows.length; offset++) {
              windowRows[offset].values.forEach((value, column) => {
                if (allowed.has(headerKey(value))) { indices[key] = column; locations[key] = start + offset; }
              });
            }
          }
          if (["criteria", "id", "name", "designation", "department", "score", "amount"].every(key => indices[key] !== undefined)
            && new Set([indices.criteria, indices.id, indices.name, indices.score, indices.amount]).size === 5) {
            candidates.push({ sheet, indices, end: Math.max(...Object.values(locations)), start, depth });
            break;
          }
        }
      }
    }
    candidates.sort((a, b) => a.start - b.start || a.depth - b.depth);
    if (!candidates.length) throw new Error("The workbook must contain Criteria, Employee ID, Employee Name, Designation, Department, Score % and Final Amount columns. Their order may vary.");
    return candidates[0];
  }

  function periodsIn(value) {
    const matches = [];
    const months = MONTHS.flatMap((month, index) => [[month, index], [SHORT[index], index]]).sort((a, b) => b[0].length - a[0].length);
    const pattern = new RegExp(`(?:^|[^A-Za-z])(${months.map(m => m[0]).join("|")})[\\s_'’./-]*(20\\d{2}|\\d{2})(?!\\d)`, "gi");
    for (const match of text(value).matchAll(pattern)) {
      const index = months.find(m => m[0].toLowerCase() === match[1].toLowerCase())[1];
      const year = match[2].length === 2 ? 2000 + Number(match[2]) : Number(match[2]);
      matches.push(`${year}-${String(index + 1).padStart(2, "0")}`);
    }
    for (const match of text(value).matchAll(/\b(20\d{2})[-/](0?[1-9]|1[0-2])(?:\b|[-/])/g)) matches.push(`${match[1]}-${match[2].padStart(2, "0")}`);
    return [...new Set(matches)];
  }

  function mostCommon(values) {
    const counts = new Map();
    values.forEach(value => counts.set(value, (counts.get(value) || 0) + 1));
    const sorted = [...counts].sort((a, b) => b[1] - a[1]);
    return sorted.length && (!sorted[1] || sorted[0][1] > sorted[1][1]) ? sorted[0][0] : "";
  }

  function detectPeriod(workbook, records, filename) {
    const explicit = records.flatMap(row => periodsIn(row.period));
    if (explicit.length) return mostCommon(explicit);
    const filePeriod = periodsIn(filename);
    if (filePeriod.length === 1) return filePeriod[0];
    const titles = workbook.sheets.flatMap(sheet => periodsIn(sheet.name));
    if (titles.length) return mostCommon(titles);
    const sources = records.flatMap(row => periodsIn(row.source));
    if (sources.length) return mostCommon(sources);
    const firstRows = workbook.sheets.flatMap(sheet => sheet.rows.slice(0, 3).flatMap(row => row.values).flatMap(periodsIn));
    return mostCommon(firstRows);
  }

  function extractData(workbook, filename = "") {
    const { sheet, indices, end } = locateHeaders(workbook);
    const canonical = new Map(CRITERIA.map(([name]) => [norm(name), name]));
    const records = [];
    for (const row of sheet.rows.slice(end + 1)) {
      const values = row.values;
      const criterion = text(values[indices.criteria]);
      const name = text(values[indices.name]);
      const id = text(values[indices.id]);
      if (!criterion && !name && !id) continue;
      if (/^(?:grand\s*)?total\b|^sub\s*total\b/i.test(criterion) && !name && !id) continue;
      if (headerKey(criterion) === "criteria" && headerKey(name) === "employeename") continue;
      if (!criterion || !name) throw new Error(`Criteria or Employee Name is missing in worksheet row ${row.number}.`);
      const amount = numeric(values[indices.amount], "Final Amount", row.number);
      if (amount < 0) throw new Error(`Final Amount cannot be negative in worksheet row ${row.number}.`);
      const rawScore = values[indices.score];
      const score = rawScore === null || rawScore === undefined || text(rawScore) === "" ? null : numeric(rawScore, "Score %", row.number);
      records.push({ criteria: canonical.get(norm(criterion)) || criterion, id, name,
        designation: text(values[indices.designation]), department: text(values[indices.department]), score,
        scoreText: typeof rawScore === "string" && /%/.test(rawScore) ? money(score) : null,
        amount, source: indices.source === undefined ? "" : text(values[indices.source]),
        period: indices.period === undefined ? "" : text(values[indices.period]), rowNumber: row.number });
    }
    if (!records.length) throw new Error("No employee records were found below the detected column headers.");
    const grouped = new Map();
    for (const record of records) {
      const key = norm(record.criteria);
      if (!grouped.has(key)) grouped.set(key, { criteria: record.criteria, records: [], amountCents: 0 });
      const group = grouped.get(key); group.records.push(record); group.amountCents += cents(record.amount);
    }
    const order = new Map(CRITERIA.map(([name], index) => [norm(name), index]));
    const groups = [...grouped.values()].sort((a, b) => (order.get(norm(a.criteria)) ?? 999) - (order.get(norm(b.criteria)) ?? 999));
    return { records, groups, total: groups.reduce((sum, group) => sum + group.amountCents, 0) / 100,
      period: detectPeriod(workbook, records, filename), filename, worksheet: sheet.name };
  }

  function today(now = new Date()) {
    const parts = Object.fromEntries(new Intl.DateTimeFormat("en-US", { timeZone: "Asia/Dhaka", day: "numeric", month: "numeric", year: "numeric" }).formatToParts(now).map(p => [p.type, p.value]));
    return { day: Number(parts.day), month: Number(parts.month) - 1, year: Number(parts.year), label: `${SHORT[Number(parts.month) - 1]} ${parts.day}, ${parts.year}` };
  }
  function periodParts(value) {
    const match = /^(20\d{2})-(0[1-9]|1[0-2])$/.exec(value || "");
    if (!match) throw new Error("Choose the reporting month before generating the PDF.");
    return { year: Number(match[1]), month: Number(match[2]) - 1, label: `${MONTHS[Number(match[2]) - 1]} ${match[1]}` };
  }

  const SMALL = ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  function words(number) {
    const n = Math.floor(Math.abs(number));
    if (n < 20) return SMALL[n];
    if (n < 100) return ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"][Math.floor(n / 10)] + (n % 10 ? " " + words(n % 10) : "");
    for (const [unit, label] of [[10000000, "Crore"], [100000, "Lac"], [1000, "Thousand"], [100, "Hundred"]]) {
      if (n >= unit) return words(Math.floor(n / unit)) + " " + label + (n % unit ? " " + words(n % unit) : "");
    }
    return "";
  }
  function amountWords(amount) {
    const value = cents(amount), taka = Math.floor(value / 100), paisa = value % 100;
    return `${words(taka)}${paisa ? ` and ${words(paisa)} Paisa` : ""} BDT only`;
  }

  function font(context, size, bold = false, family = "serif") {
    context.font = `${bold ? 700 : 400} ${size}px "${family === "serif" ? "AwardSerif" : family === "sans" ? "AwardSans" : "ApprovalSans"}", ${family === "serif" ? "serif" : "sans-serif"}`;
  }
  function white(context, x, y, width, height) { context.save(); context.fillStyle = "#ffffff"; context.fillRect(x, y, width, height); context.restore(); }
  function write(context, value, x, y, size = 12, bold = false, family = "serif", align = "left", width = 0) {
    context.save(); font(context, size, bold, family); context.fillStyle = "#000000"; context.textAlign = align;
    context.fillText(text(value), x + (align === "center" ? width / 2 : align === "right" ? width : 0), y); context.restore();
  }
  function wrapped(context, value, width, size, bold = false, family = "sans") {
    font(context, size, bold, family);
    const lines = [];
    for (const part of text(value).split("\n")) {
      let current = "";
      for (let word of part.split(/\s+/).filter(Boolean)) {
        if (current && context.measureText(current + " " + word).width > width) { lines.push(current); current = ""; }
        while (context.measureText(word).width > width && word.length > 1) {
          let at = word.length - 1;
          while (at > 1 && context.measureText(word.slice(0, at)).width > width) at--;
          if (current) { lines.push(current); current = ""; }
          lines.push(word.slice(0, at)); word = word.slice(at);
        }
        current = current ? current + " " + word : word;
      }
      if (current || !part) lines.push(current);
    }
    return lines.length ? lines : [""];
  }

  function paragraph(context, segments, { x = 36.1, y, width = 523, size = 12, leading = 13.8, family = "serif", justify = true, underline = false } = {}) {
    const tokens = [];
    for (const segment of segments) {
      font(context, size, !!segment.bold, family);
      for (const token of text(segment.text).split(/\s+/).filter(Boolean)) tokens.push({ text: token, bold: !!segment.bold, width: context.measureText(token).width });
    }
    font(context, size, false, family);
    const gap = context.measureText(" ").width;
    const lines = []; let line = [], used = 0;
    for (const token of tokens) {
      if (line.length && used + gap + token.width > width) { lines.push(line); line = []; used = 0; }
      used += (line.length ? gap : 0) + token.width; line.push(token);
    }
    if (line.length) lines.push(line);
    lines.forEach((line, index) => {
      const contentWidth = line.reduce((sum, token) => sum + token.width, 0);
      const space = justify && index < lines.length - 1 && line.length > 1 ? (width - contentWidth) / (line.length - 1) : gap;
      let left = x;
      for (const token of line) { write(context, token.text, left, y, size, token.bold, family); left += token.width + space; }
      if (underline) { context.save(); context.strokeStyle = "#000"; context.lineWidth = .5; context.beginPath(); context.moveTo(x, y + size + .5); context.lineTo(left - space, y + size + .5); context.stroke(); context.restore(); }
      y += leading;
    });
    return y;
  }

  function cell(context, value, x, y, width, height, { size = 7.4, bold = false, align = "left", family = "sans", padding = 2.4 } = {}) {
    white(context, x, y, width, height);
    context.save(); context.strokeStyle = "#000"; context.lineWidth = .5; context.strokeRect(x, y, width, height);
    const lines = wrapped(context, value, width - padding * 2, size, bold, family);
    context.textBaseline = "alphabetic";
    const metrics = context.measureText("Hg"), leading = size + 1.4;
    const inkHeight = metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent;
    let top = y + (height - inkHeight - (lines.length - 1) * leading) / 2 + metrics.actualBoundingBoxAscent;
    for (const line of lines) { write(context, line, x + padding, top, size, bold, family, align, width - padding * 2); top += leading; }
    context.restore();
  }

  function header(context, image, approval, totalPages) {
    white(context, 0, 0, A4[0], 122);
    context.drawImage(image, 0, 0, image.width, image.height * 120 / A4[1], 0, 0, A4[0], 120);
    white(context, 45, 72.2, 232, 13.2); write(context, `Date: ${approval.label}`, 47.1, 73, 11);
    white(context, 287, 72.2, 232, 13.2); write(context, `Dept Ref. No.   : Retail Operation/${approval.year}`, 288.8, 73.8, 11, false, "sans");
    white(context, 362, 101.8, 34, 12.4); write(context, String(totalPages).padStart(2, "0"), 363.5, 102, 12, false, "sans");
  }
  function footer(context, number, total) { white(context, 35, 800, 527, 40); write(context, `Page ${number} of ${total}`, 440, 815.3, 12, false, "serif", "right", 120); }

  function monthlySummary(context, data, period, totalPages) {
    white(context, 35, 422, 527, 348);
    const found = new Map(data.groups.map(group => [norm(group.criteria), group]));
    const entries = CRITERIA.map(([criteria, note]) => ({ criteria, note, count: found.get(norm(criteria))?.records.length || 0, amount: (found.get(norm(criteria))?.amountCents || 0) / 100 }));
    const known = new Set(CRITERIA.map(([criteria]) => norm(criteria)));
    data.groups.filter(g => !known.has(norm(g.criteria))).forEach(g => entries.push({ criteria: g.criteria, note: "", count: g.records.length, amount: g.amountCents / 100 }));
    const heights = Math.min(15, 180 / entries.length), widths = [232, 105, 96, 91];
    let y = 425.7;
    let x = 36.1;
    ["Criteria", "Employee QTY", "Total Amount", "Remarks"].forEach((value, index) => { cell(context, value, x, y, widths[index], 16, { size: 10.2, bold: true, align: "center", family: "approval", padding: 1 }); x += widths[index]; });
    y += 16;
    for (const entry of entries) {
      x = 36.1;
      [entry.criteria, money(entry.count), money(entry.amount), entry.note].forEach((value, index) => { cell(context, value, x, y, widths[index], heights, { size: Math.min(10.2, heights - 3), align: index === 0 ? "left" : "center", family: "approval", padding: 2.4 }); x += widths[index]; });
      y += heights;
    }
    x = 36.1;
    ["Total", money(data.records.length), money(data.total), ""].forEach((value, index) => { cell(context, value, x, y, widths[index], 16, { size: 10.2, bold: true, align: "center", family: "approval", padding: 1 }); x += widths[index]; });
    y += 30.5;
    y = paragraph(context, [{ text: "Note:", bold: true }, { text: " Employee QTY & Total Amount may vary due to expansion/closure of outlets monthly." }], { y, justify: false }) + 13.8;
    y = paragraph(context, [{ text: `Total ${SHORT[period.month]}-${String(period.year).slice(2)} Amount: BDT ${money(data.total)}`, bold: true }, { text: ` (${amountWords(data.total)})` }], { y, justify: false }) + 13.8;
    y = paragraph(context, [{ text: "Outlet Operating Procedure:", bold: true }], { y, justify: false });
    paragraph(context, [{ text: `To ensure transparency & fairness there are there are some KPI’s has been set for this Award program which attached in annexure page - 3 to ${totalPages}.` }], { y });
  }

  function employeeSignatories(context) {
    white(context, 35, 405, 527, 377);
    const width = 113;
    const draw = (person, x, y, gap) => {
      write(context, person[0] + ":", x, y, 12, true, "sans", "center", width);
      context.save(); context.strokeStyle = "#000"; context.lineWidth = .8; context.beginPath(); context.moveTo(x, y + gap); context.lineTo(x + width, y + gap); context.stroke(); context.restore();
      write(context, person[1], x, y + gap + 10, 11, true, "sans", "center", width);
      wrapped(context, person[2], width, 11, true, "sans").forEach((line, index) => write(context, line, x, y + gap + 26 + index * 14, 11, true, "sans", "center", width));
    };
    const first = [
      ["Prepared by", "Md. Aftab Ul Islam", "Sr. Executive,\nOperations"], ["Checked by", "Saiful Alam Rasel", "Assistant Director,\nOperations"],
      ["Reviewed by", "Rumana Jahan", "Head of HR"], ["Recommended by", "Abu Naser", "Director, Operations"],
    ];
    first.forEach((person, i) => draw(person, (A4[0] - 518) / 2 + 135 * i, 426, 55));
    const second = [
      ["Viewed by", "Tamal Paul", "Director - Business Planning & Investor Relations"],
      ["Endorsed by", "Md Shofiqul Islam", "Head of Finance"],
      ["Approved by", "Sabbir Hasan Nasir", "Managing Director\nACILL"],
    ];
    second.forEach((person, i) => draw(person, (A4[0] - 413) / 2 + 150 * i, 632, 46));
  }

  function renderApproval(data, images, period, approval, totalPages) {
    const first = tools.createPdfCanvas(); first.context.drawImage(images.first, 0, 0, ...A4);
    header(first.context, images.first, approval, totalPages);
    white(first.context, 35, 127, 527, 40);
    paragraph(first.context, [{ text: `Subject: Approval for Incentive Disbursement amount (Best Employee Award Program) of ${money(data.total)}/= BDT (${period.label}).`, bold: true }], { y: 131.1, size: 12, leading: 16.8, family: "sans", justify: false, underline: true });
    white(first.context, 35, 171, 527, 68);
    paragraph(first.context, [{ text: "Background:", bold: true }, { text: ` In accordance with the approved incentive policy (Best Employee Award Program) for motivating and strengthening the ongoing expansion initiatives of the Departments, the performance and eligibility for incentive disbursement for the month of ${period.label} have been reviewed, verified, and evaluated as per the approved guidelines and supporting details enclosed in Annexure.` }], { y: 173.8, leading: 15.85 });
    monthlySummary(first.context, data, period, totalPages);
    footer(first.context, 1, totalPages);
    const second = tools.createPdfCanvas(); second.context.drawImage(images.second, 0, 0, ...A4);
    header(second.context, images.first, approval, totalPages);
    white(second.context, 35, 252, 527, 89);
    paragraph(second.context, [{ text: "The approved incentive amounts shall be disbursed among the eligible employees listed in Annexure including the associate team for" }, { text: ` ${MONTHS[period.month]} -${period.year}`, bold: true }, { text: ", through name-wise Bank/MFS transfers. This incentive disbursement is a recognition of the commendable efforts, dedication, and operational commitment demonstrated by the teams in expanding Shwapno’s retail network during the stated period despite challenging conditions. Management sincerely appreciates their valuable contribution and encourages continued performance excellence in the upcoming quarters." }], { y: 256 });
    white(second.context, 35, 349, 527, 40);
    paragraph(second.context, [{ text: "In view of the above," }, { text: ` approval for Incentive Disbursement amount of ${money(data.total)}/= BDT`, bold: true }, { text: " is hereby recommended as per the details stated above and the attached annexure." }], { y: 352.8 });
    employeeSignatories(second.context); footer(second.context, 2, totalPages);
    return [first, second];
  }

  const ANNEX_WIDTHS = [22, 110, 53, 120, 64, 55, 36, 64];
  const SCORE = value => value === null ? "" : new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 }).format(value);
  function planAnnexure(data, context) {
    const pages = [[]]; let y = 191, serial = 0;
    const add = (item, height, following = 0) => {
      if (y + height + following > 790 && pages.at(-1).length) { pages.push([]); y = 191; }
      pages.at(-1).push({ ...item, y, height }); y += height;
    };
    for (const group of data.groups) {
      group.records.forEach((record, index) => {
        const values = [++serial, group.criteria, record.id, record.name, record.designation, record.department, SCORE(record.score), money(record.amount)];
        const count = Math.max(...values.map((value, column) => wrapped(context, value, ANNEX_WIDTHS[column] - 4.8, 7.4).length));
        const height = Math.max(13.4, count * 8.8 + 4.2);
        add({ type: "record", values, record }, height, index === group.records.length - 1 ? 16 : 0);
      });
      add({ type: "subtotal", label: `Subtotal - ${group.criteria} (${group.records.length})`, amount: group.amountCents / 100 }, 16);
    }
    add({ type: "total", label: `Grand Total (${data.records.length} recipients)`, amount: data.total }, 18);
    return pages;
  }

  function renderAnnexure(plan, image, approval, period, totalPages) {
    return plan.map((items, pageIndex) => {
      const page = tools.createPdfCanvas(), c = page.context;
      header(c, image, approval, totalPages);
      write(c, pageIndex ? "Annexure (continued):" : "Annexure:", 36.1, 130.8, 14, true);
      write(c, `Best Employee Award Program (${SHORT[period.month]}- ${period.year})`, 38, 157, 10, true, "sans");
      let x = 36.1;
      HEADERS.forEach((value, i) => { cell(c, value, x, 171, ANNEX_WIDTHS[i], 20, { size: 8.2, bold: true, align: "center" }); x += ANNEX_WIDTHS[i]; });
      for (const item of items) {
        if (item.type === "record") {
          x = 36.1;
          item.values.forEach((value, i) => { cell(c, value, x, item.y, ANNEX_WIDTHS[i], item.height, { align: [0, 2, 6].includes(i) ? "center" : i === 7 ? "right" : "left" }); x += ANNEX_WIDTHS[i]; });
        } else {
          cell(c, item.label, 36.1, item.y, 460, item.height, { size: 8.2, bold: true });
          cell(c, money(item.amount), 496.1, item.y, 64, item.height, { size: 8.2, bold: true, align: "right" });
        }
      }
      footer(c, pageIndex + 3, totalPages); return page;
    });
  }

  async function ensureFonts() {
    if (document.fonts) await Promise.all([document.fonts.load('400 12px "AwardSerif"'), document.fonts.load('700 12px "AwardSerif"'), document.fonts.load('400 12px "AwardSans"'), document.fonts.load('700 12px "AwardSans"'), document.fonts.load('400 12px "ApprovalSans"'), document.fonts.load('700 12px "ApprovalSans"')]);
  }
  async function buildPdf(data, options = {}) {
    const period = periodParts(options.period || data.period), approval = today(options.now);
    await ensureFonts();
    const [firstBytes, secondBytes] = await Promise.all([tools.assetBytes("assets/best-employee-template-page-1.png"), tools.assetBytes("assets/best-employee-template-page-2.png")]);
    const images = { first: await tools.bytesToCanvasImage(firstBytes), second: await tools.bytesToCanvasImage(secondBytes) };
    const plan = planAnnexure(data, tools.createPdfCanvas().context), total = plan.length + 2;
    const pages = [...renderApproval(data, images, period, approval, total), ...renderAnnexure(plan, images.first, approval, period, total)];
    const pdf = await window.PDFLib.PDFDocument.create();
    for (const item of pages) {
      const png = await pdf.embedPng(await tools.canvasPngBytes(item.canvas));
      pdf.addPage(A4).drawImage(png, { x: 0, y: 0, width: A4[0], height: A4[1] });
    }
    pdf.setTitle(`Best Employee Award Approval - ${period.label}`);
    pdf.setSubject("Best Employee Award incentive disbursement approval and criteria-wise annexure");
    pdf.setCreator("Approval Format Dashboard");
    return pdf.save();
  }

  function status(message, kind = "") { ui.Status.textContent = message; ui.Status.className = `status${kind ? " " + kind : ""}`; }
  function ready() { return !!state.data && /^(20\d{2})-(0[1-9]|1[0-2])$/.test(ui.Period.value); }
  function refresh() {
    ui.SummaryMonth.textContent = ui.Period.value ? periodParts(ui.Period.value).label : "—";
    ui.SummaryRecipients.textContent = state.data ? money(state.data.records.length) : "—";
    ui.SummaryCriteria.textContent = state.data ? money(state.data.groups.length) : "—";
    ui.SummaryTotal.textContent = state.data ? money(state.data.total) + " BDT" : "—";
    ui.ApprovalDate.textContent = state.data ? today().label : "Derived after file selection";
    ui.GenerateButton.disabled = state.busy || !ready();
  }
  function busy(value) {
    state.busy = value;
    for (const element of [ui.SelectFileButton, ui.FileInput, ui.Period, ...document.querySelectorAll("[data-approval-mode]")]) element.disabled = value;
    ui.GenerateButton.textContent = value ? "Generating approval PDF…" : "Best Employee Award Approval Generate";
    refresh();
  }
  async function handleFile(file) {
    if (!file || state.busy) return;
    const request = ++state.request; state.data = null; ui.FileName.textContent = file.name; ui.Period.value = ""; refresh();
    if (!/\.xlsx$/i.test(file.name)) { status("Select an .xlsx Excel workbook.", "bad"); return; }
    ui.FileChip.textContent = "Reading workbook"; ui.FileChip.className = "chip chip-warn"; status("Reading the award criteria and employee records…");
    try {
      const workbook = await api.parseWorkbookFile(file), data = extractData(workbook, file.name);
      if (request !== state.request) return;
      state.data = data; ui.Period.value = data.period;
      ui.FileChip.textContent = "Workbook ready"; ui.FileChip.className = "chip chip-good";
      status(data.period ? `${data.records.length} recipient rows, ${data.groups.length} criteria and ${money(data.total)} BDT loaded. Confirm the reporting month and generate the PDF.` : "Workbook loaded. Choose the reporting month, then generate the PDF.", data.period ? "good" : "warn");
      refresh();
    } catch (error) {
      if (request !== state.request) return;
      ui.FileChip.textContent = "Check workbook"; ui.FileChip.className = "chip chip-bad"; status(error.message || "The Excel file could not be read.", "bad");
    }
  }

  function selectMode(mode) {
    const employee = mode === "employee";
    $("openingWorkflow").hidden = employee; $("openingIncluded").hidden = employee; $("employeeWorkflow").hidden = !employee;
    for (const button of document.querySelectorAll("[data-approval-mode]")) {
      const active = button.dataset.approvalMode === mode; button.setAttribute("aria-pressed", String(active)); button.classList.toggle("btn-primary", active);
    }
  }
  for (const button of document.querySelectorAll("[data-approval-mode]")) button.addEventListener("click", () => selectMode(button.dataset.approvalMode));
  ui.SelectFileButton.addEventListener("click", () => ui.FileInput.click());
  ui.FileInput.addEventListener("change", () => handleFile(ui.FileInput.files?.[0]));
  ui.Period.addEventListener("change", () => { if (state.data) state.data.period = ui.Period.value; refresh(); });
  for (const eventName of ["dragenter", "dragover"]) ui.DropZone.addEventListener(eventName, event => { event.preventDefault(); if (!state.busy) ui.DropZone.classList.add("is-dragging"); });
  for (const eventName of ["dragleave", "drop"]) ui.DropZone.addEventListener(eventName, event => { event.preventDefault(); ui.DropZone.classList.remove("is-dragging"); });
  ui.DropZone.addEventListener("drop", event => { if (!state.busy) handleFile(event.dataTransfer?.files?.[0]); });
  ui.GenerateButton.addEventListener("click", async () => {
    if (!ready() || state.busy) return;
    busy(true); status("Building the approval note, signatories and criteria-wise annexure…");
    try {
      const data = { ...state.data, period: ui.Period.value }, period = periodParts(data.period);
      const bytes = await buildPdf(data);
      const blob = new Blob([bytes], { type: "application/pdf" }), url = URL.createObjectURL(blob);
      const link = document.createElement("a"); link.href = url; link.download = `Best_Employee_Award_Approval_${MONTHS[period.month]}_${period.year}.pdf`;
      document.body.appendChild(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 1500);
      status("The Best Employee Award approval PDF has been generated and downloaded.", "good");
    } catch (error) { status(error.message || "The PDF could not be generated.", "bad"); }
    finally { busy(false); }
  });
  selectMode("opening");
  window.__bestEmployeeDashboard = Object.freeze({ extractData, detectPeriod, buildPdf, planAnnexure, handleFile, amountWords, today, headers: HEADERS });
})();
