// Page wiring: form -> live preview -> .docx downloads. Filling logic is in fill.js.
(function () {
  "use strict";
  const F = window.DocFill;
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const form = document.getElementById("form");
  const tabs = document.getElementById("tabs");
  const page = document.getElementById("page");
  const status = document.getElementById("status");
  // Private-use characters mark filled / missing values inside the preview text.
  const OPEN = "", CLOSE = "", MISS_OPEN = "", MISS_CLOSE = "";

  let templates = [];   // { file, title, bytes, parts: { name: xml } }
  let current = 0;

  function base64ToBytes(b64) {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  async function loadTemplates() {
    templates = await Promise.all(window.LETTER_TEMPLATES.map(async (t) => {
      const bytes = base64ToBytes(t.data);
      const zip = await JSZip.loadAsync(bytes);
      const parts = {};
      for (const name of Object.keys(zip.files).sort()) {
        if (F.PART_RE.test(name)) parts[name] = await zip.file(name).async("string");
      }
      return { file: t.file, title: t.title, bytes, parts };
    }));
  }

  function readForm() {
    const data = {};
    for (const el of form.elements) {
      if (!el.name) continue;
      if (el.type === "radio") { if (el.checked) data[el.name] = el.value; }
      else if (el.type === "checkbox") data[el.name] = el.checked;
      else data[el.name] = el.value;
    }
    return data;
  }

  function valuesFor(t) {
    return F.valuesFor(t.file, F.computeValues(readForm()));
  }

  function missingFor(t, values) {
    const missing = new Set();
    for (const xml of Object.values(t.parts)) F.fillXml(xml, values, missing);
    return missing;
  }

  // ---- Preview -------------------------------------------------------------

  function escapeHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function markedValues(values) {
    const out = {};
    for (const [k, v] of Object.entries(values)) {
      if (typeof v !== "string") continue;
      if (v === "") out[k] = F.OPTIONAL.has(k) ? "" : MISS_OPEN + (F.LABELS[k] || k) + MISS_CLOSE;
      else out[k] = OPEN + v + CLOSE;
    }
    return out;
  }

  function runHtml(r) {
    let text = "";
    for (const c of r.childNodes) {
      if (c.namespaceURI !== W) continue;
      if (c.localName === "t") text += escapeHtml(c.textContent);
      else if (c.localName === "br") text += "\n";
      else if (c.localName === "tab") text += "    ";
    }
    const pr = r.getElementsByTagNameNS(W, "rPr")[0];
    const on = (tag) => {
      const el = pr && pr.getElementsByTagNameNS(W, tag)[0];
      const v = el && el.getAttributeNS(W, "val");
      return el && v !== "0" && v !== "false" && v !== "none";
    };
    const va = pr && pr.getElementsByTagNameNS(W, "vertAlign")[0];
    if (va && va.getAttributeNS(W, "val") === "superscript") text = `<sup>${text}</sup>`;
    if (on("u")) text = `<u>${text}</u>`;
    if (on("b")) text = `<b>${text}</b>`;
    return text;
  }

  function paragraphsHtml(xml) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    let html = "";
    for (const p of doc.getElementsByTagNameNS(W, "p")) {
      let inner = "";
      for (const r of p.getElementsByTagNameNS(W, "r")) inner += runHtml(r);
      inner = inner
        .replace(/ {6,}/g, '<span class="gap"></span>')
        .replace(/\n/g, "<br>")
        .replace(//g, "<mark>").replace(//g, "</mark>")
        .replace(//g, '<mark class="miss">').replace(//g, "</mark>");
      const listed = p.getElementsByTagNameNS(W, "numPr").length > 0;
      html += listed ? `<p class="li">•&nbsp;&nbsp;${inner}</p>` : `<p>${inner}</p>`;
    }
    return html;
  }

  function renderPreview() {
    const t = templates[current];
    if (!t) return;
    const values = valuesFor(t);
    const marked = markedValues(values);
    let html = "", foot = "";
    for (const [name, xml] of Object.entries(t.parts)) {
      const part = paragraphsHtml(F.fillXml(xml, marked));
      if (name === "word/document.xml") html = part;
      else if (part.replace(/<[^>]+>|\s/g, "")) foot += part;
    }
    page.innerHTML = html + (foot ? `<div class="foot">${foot}</div>` : "");

    const missing = new Set();
    for (const tt of templates) missingFor(tt, valuesFor(tt)).forEach((m) => missing.add(m));
    status.className = "status" + (missing.size ? " warn" : "");
    status.textContent = missing.size ? "Still empty: " + [...missing].join(", ") : "All fields filled — ready to download.";
  }

  function renderTabs() {
    tabs.innerHTML = "";
    templates.forEach((t, i) => {
      const b = document.createElement("button");
      b.type = "button";
      b.setAttribute("role", "tab");
      b.textContent = `${i + 1}. ${t.title}`;
      b.setAttribute("aria-selected", String(i === current));
      b.onclick = () => { current = i; renderTabs(); renderPreview(); };
      tabs.appendChild(b);
    });
  }

  function renderAmountWords() {
    for (const span of document.querySelectorAll(".words")) {
      const raw = form.elements[span.dataset.for].value.trim();
      const n = F.parseAmount(raw);
      span.classList.toggle("bad", raw !== "" && Number.isNaN(n));
      span.textContent = raw === "" ? "" : Number.isNaN(n) ? "Numbers only, e.g. 4500000 or 45,00,000"
        : `₹ ${F.formatIndian(n)}/- (${F.amountInWords(n)} Only)`;
    }
  }

  // ---- Downloads -----------------------------------------------------------

  function safeName(s) {
    return (s || "Student").replace(/[\\/:*?"<>|]+/g, " ").trim();
  }

  function save(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  function confirmMissing(list) {
    const missing = new Set();
    for (const t of list) missingFor(t, valuesFor(t)).forEach((m) => missing.add(m));
    return !missing.size || window.confirm("These are still empty and will be left blank:\n\n• " +
      [...missing].join("\n• ") + "\n\nDownload anyway?");
  }

  async function withBusy(fn) {
    const buttons = document.querySelectorAll("button");
    buttons.forEach((b) => (b.disabled = true));
    try { await fn(); } catch (e) { alert("Something went wrong: " + e.message); }
    finally { buttons.forEach((b) => (b.disabled = false)); }
  }

  async function downloadAll() {
    if (!confirmMissing(templates)) return;
    const student = safeName(readForm().sName);
    const out = new JSZip();
    for (const t of templates) {
      const { data } = await F.fillDocx(JSZip, t.bytes, valuesFor(t));
      out.file(`${student} - ${t.file[0]}. ${t.title}.docx`, data);
    }
    save(await out.generateAsync({ type: "blob" }), `${student} - Visa Letters.zip`);
  }

  async function downloadOne() {
    const t = templates[current];
    if (!confirmMissing([t])) return;
    const { data } = await F.fillDocx(JSZip, t.bytes, valuesFor(t));
    save(new Blob([data], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
      `${safeName(readForm().sName)} - ${t.title}.docx`);
  }

  // ---- Setup ---------------------------------------------------------------

  function today() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }

  function update() {
    form.elements.date.disabled = form.elements.dateBlank.checked;
    renderAmountWords();
    renderPreview();
  }

  form.addEventListener("input", update);
  form.addEventListener("change", update);
  form.addEventListener("submit", (e) => { e.preventDefault(); withBusy(downloadAll); });
  document.getElementById("downloadOne").addEventListener("click", () => withBusy(downloadOne));
  document.getElementById("clear").addEventListener("click", () => {
    if (!window.confirm("Clear the form for the next student?")) return;
    form.reset();
    form.elements.date.value = today();
    update();
    window.scrollTo(0, 0);
  });

  form.elements.date.value = today();
  loadTemplates().then(() => { renderTabs(); update(); })
    .catch((e) => { status.className = "status warn"; status.textContent = "Could not load the templates: " + e.message; });
})();
