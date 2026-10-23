// Run with: node --test tests/
const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const JSZip = require("../vendor/jszip.min.js");
const F = require("../fill.js");

const ROOT = path.join(__dirname, "..");
const FILES = ["1_Sponsorship_Letter", "2_Self_Affidavit", "3_Affidavit_of_Support", "4_Financial_Support_Letter"];

const SON_FATHER = {
  sName: "Arjun Nair", sGender: "son", passport: "z1234567", sMobile: "9000000001", sEmail: "arjun@example.com",
  sponsor: "father", pName: "Ravi Nair", spName: "Lekha Nair", pAge: "52", pProfession: "Engineer",
  incomeType: "salary", pMobile: "9000000002", pEmail: "ravi@example.com",
  siblings: "Anand Nair", address: "Thekkethil House, Kakkanad, Ernakulam, Kerala 682030",
  school: "mach1", propCents: "25", propPlace: "Kakkanad", propDistrict: "Ernakulam",
  propValue: "4500000", goldValue: "12,50,000", goldSov: "40", incomeValue: "1800000", date: "2026-10-23",
};

const DAUGHTER_MOTHER = Object.assign({}, SON_FATHER, {
  sName: "Meera Das", sGender: "daughter", sponsor: "mother", pName: "Sudha Das", spName: "Mohan Das",
  incomeType: "business", siblings: "", school: "worcester", dateBlank: true,
  address: "Rose Villa\nPattom\nThiruvananthapuram",
});

async function fillText(file, form) {
  const bytes = fs.readFileSync(path.join(ROOT, "templates", file + ".docx"));
  const { data, missing } = await F.fillDocx(JSZip, bytes, F.valuesFor(file, F.computeValues(form)));
  const zip = await JSZip.loadAsync(data);
  let text = "";
  for (const name of Object.keys(zip.files).filter((n) => F.PART_RE.test(n)).sort()) {
    const xml = await zip.file(name).async("string");
    text += xml.replace(/<w:br\/>/g, "\n").replace(/<\/w:p>/g, "\n").replace(/<[^>]+>/g, "") + "\n";
  }
  return { text: text.replace(/&amp;/g, "&"), missing };
}

test("amount in words uses Indian numbering", () => {
  assert.strictEqual(F.amountInWords(4500000), "Forty Five Lakh");
  assert.strictEqual(F.amountInWords(1250000), "Twelve Lakh Fifty Thousand");
  assert.strictEqual(F.amountInWords(25075300), "Two Crore Fifty Lakh Seventy Five Thousand Three Hundred");
  assert.strictEqual(F.amountInWords(101), "One Hundred One");
  assert.strictEqual(F.formatIndian(4500000), "45,00,000");
  assert.strictEqual(F.formatIndian(125000000), "12,50,00,000");
  assert.strictEqual(F.formatIndian(950), "950");
  assert.strictEqual(F.parseAmount("45,00,000"), 4500000);
  assert.ok(Number.isNaN(F.parseAmount("forty lakh")));
});

test("address splits at commas or lines", () => {
  assert.deepStrictEqual(F.addressLines("A House, B Road, C"), ["A House", "B Road", "C"]);
  assert.deepStrictEqual(F.addressLines("A House, B Road\nC"), ["A House, B Road", "C"]);
});

for (const file of FILES) {
  test(`${file}: son/father fills every placeholder`, async () => {
    const { text, missing } = await fillText(file, SON_FATHER);
    assert.deepStrictEqual(missing, []);
    assert.ok(!text.includes("{{"), text.match(/\{\{\w+\}\}/g));
    assert.ok(text.includes("Arjun Nair"));
    assert.ok(text.includes("Mach1 Aviation Academy"));
    assert.ok(!/\bher\b|\bshe\b|\bdaughter|Mrs\. Ravi/.test(text), "wrong gender wording");
    assert.ok(text.includes(file === "2_Self_Affidavit" ? "9000000001" : "9000000002"));
    assert.ok(text.includes("23/10/2026"));
  });

  test(`${file}: daughter/mother, Worcester, no siblings`, async () => {
    const { text, missing } = await fillText(file, DAUGHTER_MOTHER);
    assert.deepStrictEqual(missing, []);
    assert.ok(!text.includes("{{"));
    assert.ok(text.includes("Worcester Flying School"));
    assert.ok(!text.includes("Mach1") && !text.includes("Benoni"));
    assert.ok(!/\bhis\b|\bhe\b|\bson\b/.test(text), "wrong gender wording");
    assert.ok(text.includes("    /    /2026"));
  });
}

test("sponsorship letter wording", async () => {
  const a = (await fillText("1_Sponsorship_Letter", SON_FATHER)).text;
  assert.ok(a.includes("I Mr. Ravi Nair, aged 52 years, father of Mr. Arjun Nair, Permanent resident of Thekkethil House, Kakkanad, Ernakulam, Kerala 682030 solemnly"));
  assert.ok(a.includes("consisting my wife Mrs. Lekha Nair and my son Arjun Nair and his sibling Anand Nair. "));
  assert.ok(a.includes("Indian Rupees 45,00,000/- (Forty Five Lakh Only)."));
  assert.ok(a.includes("Rupees 12,50,000/- (Twelve Lakh Fifty Thousand Only) as its"));
  assert.ok(a.includes("Rupees 18,00,000/- (Eighteen Lakh Only) through my salary annually"));
  assert.ok(a.includes("in Kakkanad, Ernakulam, Kerala."));
  assert.ok(a.includes("I am working as an Engineer. My salary is"));
  assert.ok(a.includes("Ravi Nair (Father)\n\nThekkethil House\nKakkanad\nErnakulam\nKerala 682030"));

  const b = (await fillText("1_Sponsorship_Letter", DAUGHTER_MOTHER)).text;
  assert.ok(b.includes("consisting my husband Mr. Mohan Das and my daughter Meera Das. "));
  assert.ok(b.includes("My revenue from the business is the main source"));
  assert.ok(b.includes("income derived from my business."));
});

test("self affidavit and support letter wording", async () => {
  const s = (await fillText("2_Self_Affidavit", DAUGHTER_MOTHER)).text;
  assert.ok(s.includes("I, Ms. Meera Das D/O Mrs. Sudha Das, Rose Villa, Pattom, Thiruvananthapuram have applied in Worcester Flying School, Worcester Airfield"));
  assert.ok(s.includes("My Passport No. Z1234567"));
  const a = (await fillText("3_Affidavit_of_Support", SON_FATHER)).text;
  assert.ok(a.includes("during his stay at 9, 9th road, Vlakfontein, No-69-IR, Benoni 1501 and SOUTH AFRICA."));
  const w = (await fillText("3_Affidavit_of_Support", DAUGHTER_MOTHER)).text;
  assert.ok(w.includes("stay in Vreugde Avenue 6, Worcester West, 850 and SOUTH AFRICA,"));
  const f = (await fillText("4_Financial_Support_Letter", SON_FATHER)).text;
  assert.ok(f.includes("Financial support for Mr. Arjun Nair"));
  assert.ok(f.includes("I, Mr. Ravi Nair, father of Mr. Arjun Nair, that I will financially support my son for his studies"));
});

test("missing fields are reported", async () => {
  const { missing } = await fillText("2_Self_Affidavit", Object.assign({}, SON_FATHER, { passport: "", sEmail: "" }));
  assert.deepStrictEqual(missing.sort(), ["Email", "Passport no."]);
});

test("templates.js matches templates/", () => {
  const js = fs.readFileSync(path.join(ROOT, "templates.js"), "utf8");
  const list = JSON.parse(js.slice(js.indexOf("["), js.lastIndexOf("]") + 1));
  for (const t of list) {
    assert.strictEqual(t.data, fs.readFileSync(path.join(ROOT, "templates", t.file + ".docx")).toString("base64"));
  }
});
