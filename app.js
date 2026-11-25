// Page wiring: form -> live preview -> print to PDF. Filling logic is in fill.js, page layout in render.js.
(function () {
  "use strict";
  const F = window.DocFill;
  const form = document.getElementById("form");
  const tabs = document.getElementById("tabs");
  const page = document.getElementById("page");
  const status = document.getElementById("status");
  // Private-use characters mark filled / missing values inside the preview text.
  const OPEN = "", CLOSE = "", MISS_OPEN = "", MISS_CLOSE = "";

  let templates = [];   // { file, title, files: { name: xml }, parts: { name: xml } (the ones with placeholders) }
  let current = 0;
  const LAYOUT_RE = /^word\/(styles|numbering)\.xml$|^word\/_rels\/document\.xml\.rels$/;

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
      const files = {}, parts = {};
      for (const name of Object.keys(zip.files).sort()) {
        if (F.PART_RE.test(name)) files[name] = parts[name] = await zip.file(name).async("string");
        else if (LAYOUT_RE.test(name)) files[name] = await zip.file(name).async("string");
      }
      await DocRender.loadEmbeddedFonts(zip);
      return { file: t.file, title: t.title, files, parts };
    }));
  }

  function filledFiles(t, values) {
    const out = Object.assign({}, t.files);
    for (const [name, xml] of Object.entries(t.parts)) out[name] = F.fillXml(xml, values);
    return out;
  }

  // ---- Saved "Other" schools (this browser only) ----------------------------

  const SCHOOLS_KEY = "docfiller.schools";
  const schoolsBox = document.getElementById("schools");
  const otherLabel = schoolsBox.querySelector('input[value="other"]').parentElement;

  function loadSchools() {
    try { const list = JSON.parse(localStorage.getItem(SCHOOLS_KEY)); return Array.isArray(list) ? list : []; }
    catch (e) { return []; }
  }

  function storeSchools(list) {
    try { localStorage.setItem(SCHOOLS_KEY, JSON.stringify(list)); } catch (e) { /* private window: not saved */ }
  }

  function renderSavedSchools(select) {
    schoolsBox.querySelectorAll("[data-saved]").forEach((l) => l.remove());
    loadSchools().forEach((s, i) => {
      const label = document.createElement("label");
      label.dataset.saved = "";
      label.innerHTML = '<input type="radio" name="school"><span></span>';
      label.firstChild.value = "saved:" + i;
      label.firstChild.checked = select === i;
      label.lastChild.textContent = s.school;
      schoolsBox.insertBefore(label, otherLabel);
    });
  }

  // When an "Other" school is used for a download, remember it.
  function rememberOtherSchool() {
    const d = readForm(true);
    if (d.school !== "other") return;
    const s = { school: d.oSchool.trim(), town: d.oTown.trim(), address: d.oAddress.trim() };
    if (!s.school || !s.town || !s.address) return;
    const list = loadSchools().filter((x) => x.school.toLowerCase() !== s.school.toLowerCase());
    list.push(s);
    storeSchools(list);
    renderSavedSchools(list.length - 1);
    update();
  }

  function readForm(raw) {
    const data = {};
    for (const el of form.elements) {
      if (!el.name) continue;
      if (el.type === "radio") { if (el.checked) data[el.name] = el.value; }
      else if (el.type === "checkbox") data[el.name] = el.checked;
      else data[el.name] = el.value;
    }
    const saved = !raw && /^saved:(\d+)$/.exec(data.school || "");
    const s = saved && loadSchools()[saved[1]];
    if (s) Object.assign(data, { school: "other", oSchool: s.school, oTown: s.town, oAddress: s.address });
    return data;
  }

  function valuesFor() {
    return F.computeValues(readForm());
  }

  function missingFor(t, values) {
    const missing = new Set();
    for (const xml of Object.values(t.parts)) F.fillXml(xml, values, missing);
    return missing;
  }

  // ---- Preview -------------------------------------------------------------

  function markedValues(values) {
    const out = {};
    for (const [k, v] of Object.entries(values)) {
      if (typeof v !== "string") continue;
      if (v === "") out[k] = F.OPTIONAL.has(k) ? "" : MISS_OPEN + (F.LABELS[k] || k) + MISS_CLOSE;
      else out[k] = OPEN + v + CLOSE;
    }
    return out;
  }

  // Pages are laid out at real size (8.5in wide) and zoomed to fit the preview column.
  function fitPreview() {
    const pages = page.firstElementChild;
    if (!pages) return;
    pages.style.zoom = 1;
    pages.style.zoom = Math.min(1, (page.clientWidth - 24) / pages.offsetWidth);
  }

  function highlight(html) {
    return html.split(OPEN).join("<mark>").split(CLOSE).join("</mark>")
      .split(MISS_OPEN).join('<mark class="miss">').split(MISS_CLOSE).join("</mark>");
  }

  function renderPreview() {
    const t = templates[current];
    if (!t) return;
    page.innerHTML = '<div class="pages">' + DocRender.renderPages(filledFiles(t, markedValues(valuesFor(t))), highlight) + "</div>";
    fitPreview();

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

  function confirmMissing(list) {
    const missing = new Set();
    for (const t of list) missingFor(t, valuesFor(t)).forEach((m) => missing.add(m));
    return !missing.size || window.confirm("These are still empty and will be left blank:\n\n• " +
      [...missing].join("\n• ") + "\n\nContinue anyway?");
  }

  // The browser's print window turns the pages into a PDF ("Save as PDF");
  // the page title becomes the suggested file name.
  function printLetters(list, name) {
    if (!confirmMissing(list)) return;
    const area = document.getElementById("print");
    area.innerHTML = list.map((t) => DocRender.renderPages(filledFiles(t, valuesFor(t)))).join("");
    const title = document.title;
    document.title = name;
    window.addEventListener("afterprint", () => { document.title = title; area.innerHTML = ""; }, { once: true });
    rememberOtherSchool();
    window.print();
  }

  // ---- Setup ---------------------------------------------------------------

  function today() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }

  function update() {
    form.elements.date.disabled = form.elements.dateBlank.checked;
    const school = readForm(true).school || "";
    document.getElementById("otherSchool").hidden = school !== "other";
    document.getElementById("savedNote").hidden = !school.startsWith("saved:");
    renderAmountWords();
    renderPreview();
  }

  form.addEventListener("input", update);
  form.addEventListener("change", update);
  form.addEventListener("submit", (e) => {
    e.preventDefault();
    printLetters(templates, `${safeName(readForm().sName)} - Visa Letters`);
  });
  document.getElementById("downloadOne").addEventListener("click", () => {
    const t = templates[current];
    printLetters([t], `${safeName(readForm().sName)} - ${t.title}`);
  });
  window.addEventListener("resize", fitPreview);
  document.getElementById("forgetSchool").addEventListener("click", () => {
    const m = /^saved:(\d+)$/.exec(readForm(true).school || "");
    const list = loadSchools();
    if (!m || !list[m[1]] || !window.confirm(`Remove "${list[m[1]].school}" from the saved schools?`)) return;
    list.splice(Number(m[1]), 1);
    storeSchools(list);
    renderSavedSchools();
    form.elements.school.value = "mach1";
    update();
  });
  document.getElementById("clear").addEventListener("click", () => {
    if (!window.confirm("Clear the form for the next student?")) return;
    form.reset();
    form.elements.date.value = today();
    update();
    window.scrollTo(0, 0);
  });

  form.elements.date.value = today();
  renderSavedSchools();
  loadTemplates().then(() => { renderTabs(); update(); })
    .catch((e) => { status.className = "status warn"; status.textContent = "Could not load the templates: " + e.message; });
})();
