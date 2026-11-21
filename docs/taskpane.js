/* iLA Outlook add-in (test) — reads the open email and shows a summary. */
let summaryText = "";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    setStatus("Open this add-in from Outlook.");
    return;
  }
  loadItem();
  // Refresh when the user selects another email while the pane is pinned
  if (Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadItem);
  }
  document.getElementById("copy").addEventListener("click", copySummary);
});

function loadItem() {
  const item = Office.context.mailbox.item;
  if (!item) { setStatus("Select an email."); return; }

  const from = item.from ? `${item.from.displayName} <${item.from.emailAddress}>` : "—";
  const received = item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString("en-AU") : "—";
  const atts = (item.attachments || []).filter((a) => !a.isInline);

  text("from", from);
  text("subject", item.subject || "(no subject)");
  text("received", received);
  text("att-count", `(${atts.length})`);

  const ul = document.getElementById("attachments");
  ul.innerHTML = "";
  if (!atts.length) ul.innerHTML = '<li class="muted">None</li>';
  atts.forEach((a) => {
    const li = document.createElement("li");
    li.textContent = a.name;
    const s = document.createElement("small");
    s.textContent = formatBytes(a.size);
    li.appendChild(s);
    ul.appendChild(li);
  });

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

async function copySummary() {
  try {
    await navigator.clipboard.writeText(summaryText);
    toast("Copied.");
  } catch {
    toast("Clipboard blocked here — select the text manually.");
  }
}

function text(id, v) { document.getElementById(id).textContent = v; }
function setStatus(msg) { text("status", msg); }
function toast(msg) { text("toast", msg); setTimeout(() => text("toast", ""), 2500); }
function formatBytes(n) {
  if (!n && n !== 0) return "";
  return n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(0)} KB` : `${(n / 1048576).toFixed(1)} MB`;
}
