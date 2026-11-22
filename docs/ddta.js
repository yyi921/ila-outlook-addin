/* iLA DDTA core: read a DDTA file name, draw the AP purchase invoice page,
   and put it in front of the raw DDTA scan. No Outlook code here, so it also runs in Node. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("pdf-lib"));
  else root.ILADdta = factory(root.PDFLib);
})(typeof self !== "undefined" ? self : this, function (PDFLib) {
  const { PDFDocument, StandardFonts, rgb } = PDFLib;

  // What goes on the invoice for a DDTA. A cheque profile can be added beside it later.
  const DDTA_PROFILE = {
    entity: "Breakwater Island Trust",
    supplier: "SV0753--Gaming wins",
    location: "10--Breakwater Island Trust",
    currency: "AUD",
    amountAccount: "1054--Cash on Hand - Main Bank/Safe",
    gstAccount: "1080--GST Input Tax",
  };

  function isDdtaFile(name) {
    return /\.pdf$/i.test(name || "") && /ddta\s*#?\s*\d/i.test(name);
  }

  /* "DDTA 1234 - 260105 CITIZEN Jane $2,500 DDTA#1234.pdf"
     -> { number:"1234", date:"05/01/2026", name:"Jane Citizen", amount:2500, warnings:[] } */
  function parseDdtaFilename(filename) {
    const base = String(filename || "").replace(/\.pdf$/i, "").trim();
    const out = { number: "", date: "", name: "", amount: null, warnings: [] };

    const numbers = [...new Set([...base.matchAll(/DDTA\s*#?\s*(\d{3,})/gi)].map((m) => m[1]))];
    if (numbers.length) out.number = numbers[0];
    else out.warnings.push("No DDTA number in the file name.");
    if (numbers.length > 1) out.warnings.push(`File name has two DDTA numbers (${numbers.join(", ")}).`);

    const d = base.match(/-\s*(\d{6})(?!\d)/) || base.match(/(?:^|\s)(\d{6})(?!\d)/);
    if (d) {
      const yy = +d[1].slice(0, 2), mm = +d[1].slice(2, 4), dd = +d[1].slice(4, 6);
      const dt = new Date(2000 + yy, mm - 1, dd);
      if (dt.getMonth() === mm - 1 && dt.getDate() === dd) out.date = `${pad(dd)}/${pad(mm)}/${2000 + yy}`;
    }
    if (!out.date) out.warnings.push("No valid date (YYMMDD) in the file name.");

    const a = base.match(/\$\s*([\d,]+(?:\.\d{1,2})?)/);
    if (a) out.amount = parseFloat(a[1].replace(/,/g, ""));
    if (!(out.amount > 0)) { out.amount = null; out.warnings.push("No amount in the file name."); }

    const n = base.match(/\d{6}\s+(.+?)\s*(?:\$|DDTA\s*#|$)/i);
    if (n) out.name = tidyName(n[1]);
    if (!out.name) out.warnings.push("No patron name in the file name.");

    return out;
  }

  // "CITIZEN Jane" -> "Jane Citizen": upper-case words are the surname and go last.
  function tidyName(raw) {
    const words = raw.replace(/,/g, " ").split(/\s+/).filter(Boolean);
    const isUpper = (w) => /[A-Z]/.test(w) && w === w.toUpperCase();
    const surname = words.filter(isUpper), given = words.filter((w) => !isUpper(w));
    const ordered = surname.length && given.length ? given.concat(surname) : words;
    return ordered.map(titleCase).join(" ");
  }
  function titleCase(w) {
    return w.toLowerCase().replace(/(^|[-'])([a-z])/g, (m, p, c) => p + c.toUpperCase());
  }
  function pad(n) { return String(n).padStart(2, "0"); }

  function formatMoney(n) {
    return "$" + Number(n).toLocaleString("en-AU", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  // Returns a list of problems; empty means the fields are ready to use.
  function validateFields(f) {
    const errs = [];
    if (!/^\d+$/.test(f.number || "")) errs.push("DDTA number");
    if (!/^\d{2}\/\d{2}\/\d{4}$/.test(f.date || "")) errs.push("date (dd/mm/yyyy)");
    if (!(f.name || "").trim()) errs.push("patron name");
    if (!(Number(f.amount) > 0)) errs.push("amount");
    return errs;
  }

  function invoiceNumber(f) { return "DDTA" + f.number; }
  function description(f) { return `${invoiceNumber(f)} - ${f.name.trim()}`; }

  /* One Letter page laid out like the Sage Intacct "AP purchase invoice" print. */
  async function addInvoicePage(doc, f, profile) {
    const p = profile || DDTA_PROFILE;
    const page = doc.addPage([612, 792]);
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const bold = await doc.embedFont(StandardFonts.HelveticaBold);
    const H = 792, black = rgb(0, 0, 0), grey = rgb(0.72, 0.72, 0.72);
    // Standard PDF fonts only cover Latin-1.
    const safe = (s) => String(s == null ? "" : s).replace(/[^\x20-\x7E -ÿ]/g, "?");
    const width = (s, size, fnt) => (fnt || font).widthOfTextAtSize(safe(s), size);
    const text = (s, x, yTop, o = {}) => {
      const size = o.size || 8, fnt = o.font || font;
      let tx = x;
      if (o.align === "center") tx = x - width(s, size, fnt) / 2;
      if (o.align === "right") tx = x - width(s, size, fnt);
      page.drawText(safe(s), { x: tx, y: H - yTop, size, font: fnt, color: black });
    };
    const wrap = (s, maxW, size) => {
      const lines = []; let line = "";
      for (const word of safe(s).split(/\s+/).filter(Boolean)) {
        const next = line ? line + " " + word : word;
        if (line && width(next, size) > maxW) { lines.push(line); line = word; } else line = next;
      }
      if (line) lines.push(line);
      return lines.length ? lines : [""];
    };

    text(p.entity, 306, 17, { font: bold, size: 9, align: "center" });
    text(p.entity, 306, 29, { font: bold, size: 9, align: "center" });
    text("Accounts Payable purchase invoices", 306, 43, { size: 12, align: "center" });

    const inv = invoiceNumber(f), desc = description(f);
    const left = [
      ["Date", f.date], ["Supplier", p.supplier], ["AP purchase invoice", inv], ["number", ""],
      ["GL posting date", f.date], ["Reference number", inv], ["Description", desc], ["Term", ""],
      ["Due date", f.date],
    ];
    const right = [
      ["Recommended payment date", ""], ["Payment priority", "Normal"],
      ["Place AP purchase invoice on", "FALSE"], ["hold", ""], ["Transaction currency", p.currency],
      ["Exchange rate date", ""], ["Exchange rate type", "Intacct Daily Rate"], ["Exchange rate", "1"],
      ["Attachments", ""], ["Lightyear_Doc", ""], ["AccountsFlow_doc", ""], ["HL_Venue", ""],
      ["HL_period_start", ""], ["HL_period_end", ""],
    ];
    left.forEach(([k, v], i) => { text(k, 11, 57 + i * 12.5); text(v, 126, 57 + i * 12.5); });
    right.forEach(([k, v], i) => { text(k, 357, 57 + i * 12.5); text(v, 492, 57 + i * 12.5); });

    // Line-item table
    const cols = [11, 155, 300, 390, 480, 570];
    const top = 260, headH = 29, padX = 3, lineH = 9.5;
    page.drawRectangle({ x: cols[0], y: H - top - headH, width: cols[5] - cols[0], height: headH,
      color: grey, borderColor: black, borderWidth: 0.5 });
    ["GL account", "Memo", "Department", "Location", "Amount"].forEach((h, i) =>
      text(h, (cols[i] + cols[i + 1]) / 2, top + headH / 2 + 3, { align: "center" }));

    const rows = [
      [p.amountAccount, desc, "", p.location, formatMoney(f.amount)],
      [p.gstAccount, desc, "", p.location, formatMoney(0)],
    ];
    let y = top + headH;
    for (const row of rows) {
      const cells = row.map((c, i) => (i === 4 ? [c] : wrap(c, cols[i + 1] - cols[i] - 2 * padX, 8)));
      const rowH = Math.max(...cells.map((c) => c.length)) * lineH + 5;
      cells.forEach((lines, i) => lines.forEach((ln, j) => {
        if (i === 4) text(ln, cols[5] - padX, y + 9 + j * lineH, { align: "right" });
        else text(ln, cols[i] + padX, y + 9 + j * lineH);
      }));
      y += rowH;
    }
    const bodyTop = top + headH;
    page.drawRectangle({ x: cols[0], y: H - y, width: cols[5] - cols[0], height: y - bodyTop,
      borderColor: black, borderWidth: 0.5 });
    cols.slice(1, 5).forEach((x) => page.drawLine({ start: { x, y: H - top }, end: { x, y: H - y }, thickness: 0.5, color: black }));

    page.drawRectangle({ x: cols[4], y: H - y - 24, width: cols[5] - cols[4], height: 24,
      color: grey, borderColor: black, borderWidth: 0.5 });
    text("Total", cols[4] - 2, y + 11, { align: "right" });
    text(formatMoney(f.amount), cols[5] - padX, y + 15, { align: "right" });

    page.drawLine({ start: { x: 22, y: H - 768 }, end: { x: 600, y: H - 768 }, thickness: 0.5, color: black });
    text("Page 1", 592, 778, { align: "right", size: 7 });
    return page;
  }

  /* Invoice page first, then every page of the raw DDTA scan. Returns Uint8Array. */
  async function buildCombinedPdf(f, rawPdfBytes, profile) {
    const errs = validateFields(f);
    if (errs.length) throw new Error("Missing or invalid: " + errs.join(", "));
    const doc = await PDFDocument.create();
    doc.setTitle(description(f));
    await addInvoicePage(doc, f, profile);
    if (rawPdfBytes) {
      const raw = await PDFDocument.load(rawPdfBytes, { ignoreEncryption: true });
      (await doc.copyPages(raw, raw.getPageIndices())).forEach((pg) => doc.addPage(pg));
    }
    return doc.save();
  }

  function outputFilename(f) { return `${description(f).replace(/[\\/:*?"<>|]/g, "")}.pdf`; }

  return { DDTA_PROFILE, isDdtaFile, parseDdtaFilename, validateFields, formatMoney,
    invoiceNumber, description, outputFilename, buildCombinedPdf };
});
