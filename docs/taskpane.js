/* iLA Outlook add-in — reads the open email, turns each DDTA attachment into an
   AP invoice + DDTA PDF, and emails them one by one to Lightyear. */
(function () {
"use strict";
const CFG = window.ILA_CONFIG || {};
const D = window.ILADdta;
const MAX_ATTACH_BYTES = 3 * 1024 * 1024; // Graph sendMail accepts about 3 MB per attachment this way

let summaryText = "";
let rows = [];        // one per DDTA attachment: { name, getBytes, fields, warnings, el, state }
let inOutlook = false;
let busy = false;
let pca = null;

Office.onReady((info) => {
  inOutlook = info.host === Office.HostType.Outlook;
  document.getElementById("copy").addEventListener("click", copySummary);
  document.getElementById("send").addEventListener("click", sendAll);
  text("ly-address", CFG.lightyearEmail || "(not set)");

  if (!inOutlook) { startTestMode(); return; }
  loadItem();
  // Refresh when the user selects another email while the pane is pinned
  if (Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadItem);
  }
});

/* ---------- Reading the email ---------- */

function loadItem() {
  const item = Office.context.mailbox.item;
  if (!item) { setStatus("Select an email."); return; }

  const from = item.from ? `${item.from.displayName} <${item.from.emailAddress}>` : "—";
  const received = item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString("en-AU") : "—";
  const atts = (item.attachments || []).filter((a) => !a.isInline);

  text("from", from);
  text("subject", item.subject || "(no subject)");
  text("received", received);

  const ddta = atts.filter((a) => D.isDdtaFile(a.name));
  const other = atts.filter((a) => !D.isDdtaFile(a.name));
  setRows(ddta.map((a) => ({ name: a.name, getBytes: () => attachmentBytes(item, a) })));
  listOther(other);

  item.body.getAsync(Office.CoercionType.Text, (res) => {
    const body = res.status === Office.AsyncResultStatus.Succeeded ? res.value.trim() : "";
    text("preview", body.slice(0, 600) + (body.length > 600 ? "…" : ""));
    summaryText = [
      `From: ${from}`,
      `Subject: ${item.subject || ""}`,
      `Received: ${received}`,
      `Attachments: ${atts.map((a) => a.name).join(", ") || "none"}`,
    ].join("\n");
  });

  document.getElementById("status").hidden = true;
  document.getElementById("email").hidden = false;
}

function attachmentBytes(item, att) {
  return new Promise((resolve, reject) => {
    if (!item.getAttachmentContentAsync) return reject(new Error("This Outlook version can't read attachment content."));
    item.getAttachmentContentAsync(att.id, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(res.error.message));
      if (res.value.format !== "base64") return reject(new Error("Attachment is not a file."));
      resolve(base64ToBytes(res.value.content));
    });
  });
}

/* Outside Outlook the page still works on PDFs you pick, so the output can be checked. */
function startTestMode() {
  document.getElementById("status").hidden = true;
  document.getElementById("email").hidden = false;
  document.getElementById("meta").hidden = true;
  document.getElementById("testmode").hidden = false;
  document.getElementById("files").addEventListener("change", (e) => {
    const files = [...e.target.files];
    setRows(files.filter((f) => D.isDdtaFile(f.name))
      .map((f) => ({ name: f.name, getBytes: async () => new Uint8Array(await f.arrayBuffer()) })));
    listOther(files.filter((f) => !D.isDdtaFile(f.name)));
  });
  setRows([]);
}

/* ---------- DDTA cards ---------- */

function setRows(sources) {
  rows = sources.map((s) => {
    const parsed = D.parseDdtaFilename(s.name);
    return { ...s, fields: { number: parsed.number, date: parsed.date, name: parsed.name,
      amount: parsed.amount == null ? "" : parsed.amount.toFixed(2) }, warnings: parsed.warnings };
  });
  const box = document.getElementById("ddta-list");
  box.innerHTML = "";
  text("ddta-count", `(${rows.length})`);
  document.getElementById("ddta-empty").hidden = rows.length > 0;
  rows.forEach((r) => box.appendChild(renderRow(r)));
  refreshSendButton();
}

function renderRow(r) {
  const card = el("div", "card");
  const head = el("label", "card-head");
  r.check = el("input"); r.check.type = "checkbox";
  const sentOn = sentDate(r.fields.number);
  r.check.checked = !sentOn;
  r.check.addEventListener("change", refreshSendButton);
  head.append(r.check, el("span", "file", r.name));
  card.appendChild(head);

  const grid = el("div", "fields");
  [["number", "DDTA no."], ["date", "Date"], ["name", "Patron"], ["amount", "Amount $"]].forEach(([key, label]) => {
    const input = el("input"); input.value = r.fields[key];
    if (key === "amount") input.inputMode = "decimal";
    input.addEventListener("input", () => { r.fields[key] = input.value.trim(); r.done = false; refreshRow(r); refreshSendButton(); });
    const wrap = el("label", key === "name" ? "wide" : "", label);
    wrap.appendChild(input);
    grid.appendChild(wrap);
  });
  card.appendChild(grid);

  r.note = el("p", "note");
  card.appendChild(r.note);
  const dl = el("button", "link", "Download PDF");
  dl.type = "button";
  dl.addEventListener("click", () => downloadRow(r));
  card.appendChild(dl);

  r.sentOn = sentOn;
  refreshRow(r);
  return card;
}

function rowFields(r) {
  return { ...r.fields, amount: parseFloat(String(r.fields.amount).replace(/[$,\s]/g, "")) };
}

function refreshRow(r, message, kind) {
  const errs = D.validateFields(rowFields(r));
  r.valid = errs.length === 0;
  let msg = message, cls = kind || "";
  if (!msg && errs.length) { msg = "Check: " + errs.join(", "); cls = "bad"; }
  else if (!msg && r.warnings.length) { msg = r.warnings.join(" "); cls = "warn"; }
  else if (!msg && r.sentOn) { msg = `Already sent ${r.sentOn}. Tick to send again.`; cls = "warn"; }
  else if (!msg) { msg = `${D.description(rowFields(r))} · ${D.formatMoney(rowFields(r).amount)}`; }
  r.note.textContent = msg;
  r.note.className = "note " + cls;
}

async function buildRow(r) {
  const f = rowFields(r);
  const raw = await r.getBytes();
  return { f, bytes: await D.buildCombinedPdf(f, raw), filename: D.outputFilename(f) };
}

async function downloadRow(r) {
  try {
    refreshRow(r, "Building PDF…");
    const { bytes, filename } = await buildRow(r);
    const a = el("a");
    a.href = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" }));
    a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 60000);
    refreshRow(r);
  } catch (e) { refreshRow(r, e.message, "bad"); }
}

/* ---------- Sending to Lightyear ---------- */

function sendBlocker() {
  if (!inOutlook) return "Open the add-in from Outlook to send. Here you can only download.";
  if (!CFG.lightyearEmail) return "Set lightyearEmail in config.js.";
  if (!CFG.clientId) return "Sending is not set up yet: add the Entra client ID to config.js (see README). You can still download each PDF.";
  if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"))
    return "This Outlook version can't sign the add-in in. Use Outlook on the web or the new Outlook.";
  return "";
}

function refreshSendButton() {
  const btn = document.getElementById("send");
  const picked = rows.filter((r) => r.check.checked && !r.done);
  const blocker = sendBlocker();
  btn.textContent = picked.length ? `Send ${picked.length} to Lightyear` : "Send to Lightyear";
  btn.disabled = busy || !!blocker || !picked.length || picked.some((r) => !r.valid);
  text("send-note", rows.length ? blocker : "");
}

async function sendAll() {
  const picked = rows.filter((r) => r.check.checked && !r.done && r.valid);
  if (!picked.length || busy) return;
  busy = true; refreshSendButton();
  let ok = 0;
  try {
    const token = await getToken();
    for (const r of picked) {            // one email per DDTA, in order
      try {
        refreshRow(r, "Building PDF…");
        const { f, bytes, filename } = await buildRow(r);
        if (bytes.length > MAX_ATTACH_BYTES) throw new Error("PDF is over 3 MB. Download it and email it by hand.");
        refreshRow(r, "Sending…");
        await sendMail(token, D.description(f), filename, bytes);
        markSent(f.number);
        r.done = true; r.sentOn = sentDate(f.number); r.check.checked = false;
        refreshRow(r, `Sent to ${CFG.lightyearEmail}`, "good");
        ok++;
      } catch (e) { refreshRow(r, "Not sent: " + e.message, "bad"); }
    }
    toast(`${ok} of ${picked.length} sent.`);
  } catch (e) {
    toast("Could not sign in: " + (e.message || e));
  } finally {
    busy = false; refreshSendButton();
  }
}

async function getToken() {
  if (!pca) {
    pca = await msal.createNestablePublicClientApplication({
      auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/common" },
    });
  }
  const request = { scopes: ["https://graph.microsoft.com/Mail.Send"] };
  try { return (await pca.acquireTokenSilent(request)).accessToken; }
  catch { return (await pca.acquireTokenPopup(request)).accessToken; }
}

async function sendMail(token, subject, filename, bytes) {
  const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({
      message: {
        subject,
        body: { contentType: "Text", content: `${subject}\n\nSent by iLA.` },
        toRecipients: [{ emailAddress: { address: CFG.lightyearEmail } }],
        attachments: [{
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: filename, contentType: "application/pdf", contentBytes: bytesToBase64(bytes),
        }],
      },
      saveToSentItems: true,
    }),
  });
  if (!res.ok) {
    let detail = "";
    try { detail = (await res.json()).error.message; } catch { /* no JSON body */ }
    throw new Error(detail || `Mail service returned ${res.status}`);
  }
}

/* Remember what this browser has sent, so a second click doesn't create a duplicate bill. */
function sentDate(number) {
  try { return number ? localStorage.getItem("ila.sent.DDTA" + number) || "" : ""; } catch { return ""; }
}
function markSent(number) {
  try { localStorage.setItem("ila.sent.DDTA" + number, new Date().toLocaleDateString("en-AU")); } catch { /* storage blocked */ }
}

/* ---------- Small helpers ---------- */

function listOther(atts) {
  const ul = document.getElementById("attachments");
  ul.innerHTML = "";
  text("att-count", `(${atts.length})`);
  if (!atts.length) ul.innerHTML = '<li class="muted">None</li>';
  atts.forEach((a) => {
    const li = el("li", "", a.name);
    li.appendChild(el("small", "", formatBytes(a.size)));
    ul.appendChild(li);
  });
}

async function copySummary() {
  try {
    await navigator.clipboard.writeText(summaryText);
    toast("Copied.");
  } catch {
    toast("Clipboard blocked here — select the text manually.");
  }
}

function base64ToBytes(b64) {
  const bin = atob(b64), out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
function bytesToBase64(bytes) {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}
function el(tag, cls, txt) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (txt != null) e.textContent = txt;
  return e;
}
function text(id, v) { document.getElementById(id).textContent = v; }
function setStatus(msg) { text("status", msg); }
function toast(msg) { text("toast", msg); setTimeout(() => text("toast", ""), 6000); }
function formatBytes(n) {
  if (!n && n !== 0) return "";
  return n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(0)} KB` : `${(n / 1048576).toFixed(1)} MB`;
}
})();
