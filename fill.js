// Core logic for filling the visa letter templates. Works in the browser
// (window.DocFill) and in Node (require('./fill.js')) for the tests.
(function (root) {
  "use strict";

  const SCHOOLS = {
    mach1: {
      school: "Mach1 Aviation Academy",
      airfield: "Springs",
      // Split in three so the "th" keeps its superscript in the letter.
      sa: ["9, 9", "th", " road, Vlakfontein, No-69-IR, Benoni 1501"],
    },
    worcester: {
      school: "Worcester Flying School",
      airfield: "Worcester",
      sa: ["Vreugde Avenue 6, Worcester West, 850", "", ""],
    },
  };

  // Shown in the preview and in the "missing" warning.
  const LABELS = {
    sName: "Student name", passport: "Passport no.", pName: "Parent's name",
    spName: "Other parent's name", pAge: "Parent's age", pProfession: "Profession", aProfession: "Profession",
    address: "Home address", addressBlock: "Home address",
    propCents: "Property (cents)", propPlace: "Property place", propDistrict: "Property district",
    propValue: "Property value", propWords: "Property value", goldSov: "Gold (sovereigns)",
    goldValue: "Gold value", goldWords: "Gold value", incomeValue: "Annual income",
    incomeWords: "Annual income", footMob: "Mobile no.", footEmail: "Email",
  };

  // ---- Numbers ---------------------------------------------------------

  const ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen",
    "Eighteen", "Nineteen"];
  const TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];

  function belowHundred(n) {
    return n < 20 ? ONES[n] : (TENS[Math.floor(n / 10)] + (n % 10 ? " " + ONES[n % 10] : ""));
  }

  function belowThousand(n) {
    const h = Math.floor(n / 100), rest = n % 100;
    return [h ? ONES[h] + " Hundred" : "", rest ? belowHundred(rest) : ""].filter(Boolean).join(" ");
  }

  // Indian numbering: 4500000 -> "Forty Five Lakh".
  function amountInWords(n) {
    if (!Number.isFinite(n) || n < 0) return "";
    n = Math.floor(n);
    if (n === 0) return "Zero";
    const crore = Math.floor(n / 1e7);
    const lakh = Math.floor((n % 1e7) / 1e5);
    const thousand = Math.floor((n % 1e5) / 1e3);
    const rest = n % 1e3;
    const parts = [];
    if (crore) parts.push(amountInWords(crore) + " Crore");
    if (lakh) parts.push(belowHundred(lakh) + " Lakh");
    if (thousand) parts.push(belowHundred(thousand) + " Thousand");
    if (rest) parts.push(belowThousand(rest));
    return parts.join(" ");
  }

  // 4500000 -> "45,00,000"
  function formatIndian(n) {
    const s = String(Math.floor(n));
    if (s.length <= 3) return s;
    const last3 = s.slice(-3);
    const head = s.slice(0, -3).replace(/\B(?=(\d{2})+(?!\d))/g, ",");
    return head + "," + last3;
  }

  // Accepts "4500000", "45,00,000", "45 00 000". Returns NaN if not a number.
  function parseAmount(text) {
    const clean = String(text || "").replace(/[,\s₹]|rs\.?|\/-/gi, "");
    return /^\d+(\.\d+)?$/.test(clean) ? Number(clean) : NaN;
  }

  // ---- Form -> placeholder values ----------------------------------------

  function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

  function listWithAnd(items) {
    return items.length <= 1 ? items.join("") : items.slice(0, -1).join(", ") + " and " + items[items.length - 1];
  }

  // Address typed over several lines is kept as lines; one long line is
  // split at the commas for the address block under the signature.
  function addressLines(text) {
    const byLine = String(text || "").split(/\n/).map((l) => l.trim().replace(/,+$/, "").trim()).filter(Boolean);
    if (byLine.length > 1) return byLine;
    return String(text || "").split(",").map((l) => l.trim()).filter(Boolean);
  }

  function formatDate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
    return m ? `${m[3]}/${m[2]}/${m[1]}` : "";
  }

  function amount(form, key) {
    const n = parseAmount(form[key]);
    return Number.isNaN(n) ? { value: String(form[key] || "").trim(), words: "" }
      : { value: formatIndian(n), words: amountInWords(n) };
  }

  function computeValues(form) {
    const t = (k) => String(form[k] || "").trim();
    const isSon = form.sGender !== "daughter";
    const isFather = form.sponsor !== "mother";
    const business = form.incomeType === "business";
    const school = SCHOOLS[form.school] || SCHOOLS.mach1;
    const his = isSon ? "his" : "her";
    const siblings = t("siblings").split(/,|\band\b|\n/).map((s) => s.trim()).filter(Boolean);
    const lines = addressLines(form.address);
    const prop = amount(form, "propValue"), gold = amount(form, "goldValue"), income = amount(form, "incomeValue");
    const year = (form.date || "").slice(0, 4) || String(new Date().getFullYear());

    return {
      date: form.dateBlank || !form.date ? `    /    /${year}` : formatDate(form.date),
      sName: t("sName"),
      sTitle: isSon ? "Mr" : "Ms",
      sdo: isSon ? "S/O" : "D/O",
      child: isSon ? "son" : "daughter",
      childs: isSon ? "son’s" : "daughter’s",
      his, he: isSon ? "he" : "she", He: isSon ? "He" : "She",
      passport: t("passport").toUpperCase(),
      pName: t("pName"),
      pTitle: isFather ? "Mr" : "Mrs",
      pRel: isFather ? "father" : "mother",
      pRelCap: isFather ? "Father" : "Mother",
      pAge: t("pAge"),
      pProfession: t("pProfession"),
      aProfession: t("pProfession") ? (/^[aeiou]/i.test(t("pProfession")) ? "an " : "a ") + t("pProfession") : "",
      spName: t("spName"),
      spTitle: isFather ? "Mrs" : "Mr",
      spouseRel: isFather ? "wife" : "husband",
      siblingClause: siblings.length
        ? ` and ${his} ${siblings.length > 1 ? "siblings" : "sibling"} ${listWithAnd(siblings)}` : "",
      incomeMain: business ? "revenue from the business" : "salary",
      incomeShort: business ? "business" : "salary",
      incomeVia: business ? "business as revenue" : "salary",
      address: lines.join(", "),
      addressBlock: lines.join("\n"),
      school: school.school, airfield: school.airfield,
      sa1: school.sa[0], sa2: school.sa[1], sa3: school.sa[2],
      propCents: t("propCents"), propPlace: t("propPlace"), propDistrict: t("propDistrict"),
      propValue: prop.value, propWords: prop.words,
      goldSov: t("goldSov"), goldValue: gold.value, goldWords: gold.words,
      incomeValue: income.value, incomeWords: income.words,
      // Footer contact: the student's on the Self Affidavit, the parent's elsewhere.
      parentContact: { footMob: t("pMobile"), footEmail: t("pEmail") },
      studentContact: { footMob: t("sMobile"), footEmail: t("sEmail") },
    };
  }

  function valuesFor(file, values) {
    const contact = file === "2_Self_Affidavit" ? values.studentContact : values.parentContact;
    return Object.assign({}, values, contact);
  }

  // ---- Filling the .docx -------------------------------------------------

  const PART_RE = /^word\/(document|header\d+|footer\d+)\.xml$/;
  // sa2/sa3 are legitimately empty for Worcester; siblings may be none.
  const OPTIONAL = new Set(["siblingClause", "sa1", "sa2", "sa3"]);

  function escapeXml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  // Placeholders sit whole inside one <w:t>, so a newline in a value can
  // close the text, add a line break and reopen it.
  function fillXml(xml, values, missing) {
    return xml.replace(/\{\{(\w+)\}\}/g, (_, key) => {
      const v = values[key];
      if (v == null || v === "") {
        if (!OPTIONAL.has(key) && missing) missing.add(LABELS[key] || key);
        return "";
      }
      return escapeXml(String(v)).replace(/\n/g, '</w:t><w:br/><w:t xml:space="preserve">');
    });
  }

  // Returns { data: Uint8Array of the filled .docx, missing: [labels] }.
  async function fillDocx(JSZip, templateBytes, values) {
    const zip = await JSZip.loadAsync(templateBytes);
    const missing = new Set();
    for (const name of Object.keys(zip.files)) {
      if (!PART_RE.test(name)) continue;
      const xml = await zip.file(name).async("string");
      zip.file(name, fillXml(xml, values, missing));
    }
    const data = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    return { data, missing: [...missing] };
  }

  const api = { SCHOOLS, LABELS, OPTIONAL, amountInWords, formatIndian, parseAmount, addressLines,
    computeValues, valuesFor, fillXml, fillDocx, PART_RE };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.DocFill = api;
})(typeof window !== "undefined" ? window : globalThis);
