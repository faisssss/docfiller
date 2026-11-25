// Turns a filled letter (.docx parts) into an HTML page laid out like Word:
// page size and margins, alignment, indents, line spacing, lists, footer.
// Used for the on-screen preview and for printing to PDF.
(function (root) {
  "use strict";
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const LINE = 1.15;  // Word's "single" line height for Times New Roman / Arial, in ems

  const kids = (node, tag) => node ? [...node.childNodes].filter((c) => c.namespaceURI === W && (!tag || c.localName === tag)) : [];
  const kid = (node, tag) => kids(node, tag)[0];
  const val = (node, name = "val") => node ? node.getAttributeNS(W, name) : null;
  const num = (node, name) => { const v = val(node, name); return v == null || v === "" ? null : Number(v); };
  const on = (node) => !!node && !["0", "false", "none"].includes(val(node));
  const parse = (xml) => new DOMParser().parseFromString(xml, "application/xml");
  const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  // ---- Properties: docDefaults -> style chain -> direct ------------------------

  function readRPr(rPr, o) {
    if (!rPr) return o;
    const f = kid(rPr, "rFonts"); if (f && val(f, "ascii")) o.font = val(f, "ascii");
    const sz = kid(rPr, "sz"); if (sz) o.size = Number(val(sz)) / 2;
    for (const t of ["b", "i", "u"]) { const n = kid(rPr, t); if (n) o[t] = on(n); }
    const va = kid(rPr, "vertAlign"); if (va) o.vert = val(va);
    return o;
  }

  function readPPr(pPr, o) {
    if (!pPr) return o;
    const jc = kid(pPr, "jc"); if (jc) o.jc = val(jc);
    const ind = kid(pPr, "ind");
    if (ind) {
      const l = num(ind, "left") ?? num(ind, "start"); if (l != null) o.left = l;
      const r = num(ind, "right") ?? num(ind, "end"); if (r != null) o.right = r;
      if (num(ind, "hanging") != null) { o.hanging = num(ind, "hanging"); o.firstLine = 0; }
      if (num(ind, "firstLine") != null) { o.firstLine = num(ind, "firstLine"); o.hanging = 0; }
    }
    const sp = kid(pPr, "spacing");
    if (sp) {
      for (const k of ["before", "after", "line"]) if (num(sp, k) != null) o[k] = num(sp, k);
      if (val(sp, "lineRule")) o.lineRule = val(sp, "lineRule");
    }
    const np = kid(pPr, "numPr");
    if (np) { o.numId = val(kid(np, "numId")); o.ilvl = Number(val(kid(np, "ilvl")) || 0); }
    return o;
  }

  function styleSheet(stylesXml) {
    const doc = parse(stylesXml || "<w:styles xmlns:w='" + W + "'/>");
    const dd = kid(doc.documentElement, "docDefaults");
    const baseR = readRPr(kid(kid(dd, "rPrDefault"), "rPr"), { font: "Times New Roman", size: 10 });
    const baseP = readPPr(kid(kid(dd, "pPrDefault"), "pPr"), {});
    const byId = {};
    let defaultPara = null;
    for (const s of kids(doc.documentElement, "style")) {
      if (val(s, "type") !== "paragraph") continue;
      byId[val(s, "styleId")] = s;
      if (val(s, "default") === "1") defaultPara = val(s, "styleId");
    }
    function resolve(id, seen = new Set()) {
      const s = byId[id];
      if (!s || seen.has(id)) return { p: Object.assign({}, baseP), r: Object.assign({}, baseR) };
      seen.add(id);
      const parent = val(kid(s, "basedOn"));
      const out = parent ? resolve(parent, seen) : { p: Object.assign({}, baseP), r: Object.assign({}, baseR) };
      readPPr(kid(s, "pPr"), out.p);
      readRPr(kid(s, "rPr"), out.r);
      return out;
    }
    return (id) => resolve(id || defaultPara);
  }

  // ---- Lists -------------------------------------------------------------------

  function numbering(numberingXml) {
    if (!numberingXml) return () => "";
    const doc = parse(numberingXml);
    const abstract = {};
    for (const a of kids(doc.documentElement, "abstractNum")) abstract[val(a, "abstractNumId")] = a;
    const nums = {};
    for (const n of kids(doc.documentElement, "num")) nums[val(n, "numId")] = abstract[val(kid(n, "abstractNumId"))];
    const counters = {};
    const fmt = (n, f) => f === "lowerLetter" ? String.fromCharCode(96 + n) : f === "upperLetter" ? String.fromCharCode(64 + n)
      : f === "lowerRoman" ? roman(n).toLowerCase() : f === "upperRoman" ? roman(n) : String(n);
    return function label(numId, ilvl) {
      const a = nums[numId];
      if (!a || numId === "0") return "";
      const lvl = kids(a, "lvl").find((l) => Number(val(l, "ilvl")) === ilvl);
      if (!lvl) return "";
      const c = counters[numId] = counters[numId] || [];
      c[ilvl] = (c[ilvl] || num(kid(lvl, "start"), "val") || 1) + (c[ilvl] == null ? 0 : 1);
      c.length = ilvl + 1;
      const f = val(kid(lvl, "numFmt"));
      const text = val(kid(lvl, "lvlText")) || "";
      if (f === "bullet") return text;
      return text.replace(/%(\d)/g, (_, d) => fmt(c[Number(d) - 1] || 1, val(kid(kids(a, "lvl")[Number(d) - 1], "numFmt"))));
    };
  }

  function roman(n) {
    const map = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
    let s = "";
    for (const [v, r] of map) while (n >= v) { s += r; n -= v; }
    return s;
  }

  // ---- Paragraphs ----------------------------------------------------------------

  const pt = (twips) => (twips || 0) / 20 + "pt";

  function fontCss(r) {
    const css = [`font-family:'${r.font}',${/arial|helvetica|sans/i.test(r.font) ? "sans-serif" : "serif"}`,
      `font-size:${r.vert === "superscript" || r.vert === "subscript" ? r.size * 0.65 : r.size}pt`];
    if (r.b) css.push("font-weight:bold");
    if (r.i) css.push("font-style:italic");
    if (r.u) css.push("text-decoration:underline");
    if (r.vert === "superscript") css.push("vertical-align:super");
    if (r.vert === "subscript") css.push("vertical-align:sub");
    return css.join(";");
  }

  function runText(r) {
    let s = "";
    for (const c of kids(r)) {
      if (c.localName === "t") s += esc(c.textContent);
      else if (c.localName === "br") s += "<br>";
      else if (c.localName === "tab") s += '<span style="display:inline-block;width:36pt"></span>';
    }
    return s;
  }

  function paragraph(p, styles, label) {
    const pPr = kid(p, "pPr");
    const st = styles(val(kid(pPr, "pStyle")));
    const pp = readPPr(pPr, Object.assign({}, st.p));
    const mark = readRPr(kid(pPr, "rPr"), Object.assign({}, st.r));   // paragraph mark: empty-line height, list label
    let lbl = "";
    if (pp.numId) {
      const text = label(pp.numId, pp.ilvl);
      if (text) {
        const hang = pp.hanging || 360;
        lbl = `<span style="${fontCss(mark)};display:inline-block;min-width:${pt(hang)};text-indent:0">${esc(text)}</span>`;
      }
    }
    let inner = "";
    for (const r of p.getElementsByTagNameNS(W, "r")) {
      const html = runText(r);
      if (html) inner += `<span style="${fontCss(readRPr(kid(r, "rPr"), Object.assign({}, st.r)))}">${html}</span>`;
    }
    const lh = pp.lineRule === "exact" || pp.lineRule === "atLeast" ? pt(pp.line) : String(((pp.line || 240) / 240) * LINE);
    const align = { both: "justify", center: "center", right: "right", end: "right", distribute: "justify" }[pp.jc] || "left";
    const indent = pp.hanging ? -pp.hanging : (pp.firstLine || 0);
    const css = [`${fontCss(Object.assign({}, mark, { b: false, i: false, u: false, vert: null }))}`, `line-height:${lh}`,
      `text-align:${align}`, `padding:${pt(pp.before)} ${pt(pp.right)} ${pt(pp.after)} ${pt(pp.left)}`, `text-indent:${pt(indent)}`];
    return `<div class="p" style="${css.join(";")}">${lbl}${inner || "<br>"}</div>`;
  }

  function body(xml, styles, label, container) {
    const doc = parse(xml);
    const root = container ? doc.getElementsByTagNameNS(W, container)[0] : doc.documentElement;
    return kids(root, "p").map((p) => paragraph(p, styles, label)).join("");
  }

  // ---- Page --------------------------------------------------------------------------

  function rels(xml) {
    const out = {};
    if (!xml) return out;
    for (const r of parse(xml).documentElement.childNodes) if (r.getAttribute) out[r.getAttribute("Id")] = "word/" + r.getAttribute("Target");
    return out;
  }

  // ---- Pages -----------------------------------------------------------------------

  const px = (twips) => twips / 15;   // 1440 twips = 1 inch = 96 CSS px

  // Lines of a laid-out paragraph as [top, bottom] in px from `origin`.
  function lineBoxes(p, origin) {
    const range = document.createRange();
    range.selectNodeContents(p);
    const lines = [];
    for (const r of range.getClientRects()) {
      if (!r.height) continue;
      const top = r.top - origin, bottom = r.bottom - origin;
      const line = lines.find((l) => top < l[1] - 1 && bottom > l[0] + 1);
      if (line) { line[0] = Math.min(line[0], top); line[1] = Math.max(line[1], bottom); }
      else lines.push([top, bottom]);
    }
    lines.sort((a, b) => a[0] - b[0]);
    if (!lines.length) { const b = p.getBoundingClientRect(); lines.push([b.top - origin, b.bottom - origin]); }
    return lines;
  }

  // Where each page starts, in px down the letter's text. Breaks between lines;
  // like Word, never leaves a single line of a paragraph alone on a page.
  function pageBreaks(flow, limit) {
    const origin = flow.getBoundingClientRect().top;
    const cuts = [0];
    for (const p of flow.children) {
      const lines = lineBoxes(p, origin);
      const pTop = p.getBoundingClientRect().top - origin;
      for (;;) {
        const start = cuts[cuts.length - 1];
        const k = lines.findIndex((l) => l[1] - start > limit + 0.5);
        if (k < 0) break;
        let at = k;
        if (at === 1 && lines.length > 1) at = 0;                       // orphan: move the whole paragraph
        else if (at >= 2 && lines.length - at === 1) at -= 1;           // widow: take one more line along
        const cut = at === 0 ? Math.max(pTop, start) : (lines[at - 1][1] + lines[at][0]) / 2;
        if (cut <= start + 1) { cuts.push(lines[k][0]); break; }        // a line taller than a page: give up splitting
        cuts.push(cut);
        lines.splice(0, at);
        if (at === 0) break;
      }
    }
    return cuts;
  }

  // files: { "word/document.xml": xml, "word/styles.xml": ..., "word/numbering.xml": ...,
  //          "word/_rels/document.xml.rels": ..., "word/footerN.xml": ... } with values already filled in.
  // decorate(html) may change the text HTML (e.g. highlight filled values) before the pages are measured.
  // Returns the HTML of one .sheet per page.
  function renderPages(files, decorate = (h) => h) {
    const styles = styleSheet(files["word/styles.xml"]);
    const label = numbering(files["word/numbering.xml"]);
    const doc = parse(files["word/document.xml"]);
    const sect = doc.getElementsByTagNameNS(W, "sectPr")[0];
    const sz = kid(sect, "pgSz"), mar = kid(sect, "pgMar");
    const page = { w: px(num(sz, "w") || 12240), h: px(num(sz, "h") || 15840) };
    const m = {};
    for (const [k, d] of Object.entries({ top: 1440, bottom: 1440, left: 1440, right: 1440, header: 720, footer: 720 })) m[k] = px(num(mar, k) ?? d);
    const titlePg = on(kid(sect, "titlePg"));
    const link = rels(files["word/_rels/document.xml.rels"]);
    // Page 1 uses the "first" header/footer when the letter has a different first page.
    const part = (tag, first) => {
      const refs = kids(sect, tag);
      const ref = refs.find((r) => val(r, "type") === (first && titlePg ? "first" : "default"));
      const xml = ref && files[link[ref.getAttributeNS(R, "id")]];
      const html = xml ? decorate(body(xml, styles, label)) : "";
      return html.replace(/<[^>]+>|\s/g, "") ? html : "";
    };
    const flowHtml = decorate(body(files["word/document.xml"], styles, label, "body"));
    const width = page.w - m.left - m.right, limit = page.h - m.top - m.bottom;

    const probe = document.createElement("div");
    probe.className = "sheet";
    probe.style.cssText = `position:absolute;left:-10000px;top:0;visibility:hidden;zoom:1;width:${width}px`;
    probe.innerHTML = `<div class="flow">${flowHtml}</div>`;
    document.body.appendChild(probe);
    const cuts = pageBreaks(probe.firstChild, limit);
    const total = probe.firstChild.getBoundingClientRect().height;
    probe.remove();

    return cuts.map((cut, i) => {
      const height = Math.min((cuts[i + 1] ?? total) - cut, limit);
      const header = part("headerReference", i === 0), footer = part("footerReference", i === 0);
      const box = (style, html) => `<div class="hf" style="left:${m.left}px;right:${m.right}px;${style}">${html}</div>`;
      return `<div class="sheet" style="width:${page.w}px;height:${page.h}px">` +
        (header ? box(`top:${m.header}px`, header) : "") +
        box(`top:${m.top}px;height:${height}px;overflow:hidden`, `<div class="flow" style="margin-top:${-cut}px">${flowHtml}</div>`) +
        (footer ? box(`bottom:${m.footer}px`, footer) : "") +
        "</div>";
    }).join("");
  }

  // Fonts embedded in the .docx (e.g. Lora in the covering letter), so the page matches offline.
  async function loadEmbeddedFonts(zip) {
    const table = zip.file("word/fontTable.xml"), relsFile = zip.file("word/_rels/fontTable.xml.rels");
    if (!table || !relsFile || typeof FontFace === "undefined") return;
    const link = rels(await relsFile.async("string"));
    const variants = { embedRegular: {}, embedBold: { weight: "bold" }, embedItalic: { style: "italic" },
      embedBoldItalic: { weight: "bold", style: "italic" } };
    for (const f of kids(parse(await table.async("string")).documentElement, "font")) {
      for (const [tag, desc] of Object.entries(variants)) {
        const e = kid(f, tag);
        const file = e && zip.file(link[e.getAttributeNS(R, "id")]);
        // Obfuscated fonts (non-zero fontKey, from Word) can't be used as-is; the page falls back to serif.
        if (!file || !/^\{0{8}-/.test(val(e, "fontKey") || "{0")) continue;
        try {
          const face = new FontFace(val(f, "name"), await file.async("arraybuffer"), desc);
          document.fonts.add(await face.load());
        } catch (err) { /* fall back to a system serif */ }
      }
    }
  }

  root.DocRender = { renderPages, loadEmbeddedFonts };
})(window);
