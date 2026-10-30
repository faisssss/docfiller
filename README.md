# Visa Letter Filler

Fills the 5 South Africa visa letters for a student from one short form:

1. Sponsorship Letter
2. Self Affidavit
3. Affidavit of Support
4. Financial Support Letter
5. Covering Letter for Study Visa

You type about 20 details once. The page works out the rest: his/her, son/daughter,
Mr/Ms/Mrs, S/O or D/O, wife/husband, sibling wording, school addresses and amounts in words
(Indian format, e.g. `45,00,000/- (Forty Five Lakh Only)`). You get the letters as Word files
with the original layout, and the filled text is black.

Everything runs in your browser. Student details are not saved or sent anywhere.

## Using it

- **Online:** open the GitHub Pages link (see setup below) and bookmark it.
- **Offline:** download this repo and double-click `index.html`.

Fill the form, check the live preview (blue = filled, red = still empty), then press
**Download all 5 letters**. Press **Clear** before the next student.

## Built-in choices

| Field | Behaviour |
|---|---|
| Date | Today's date by default; tick "Leave blank" to keep `    /    /2026` for writing by hand |
| Footer phone/email | Student's on the Self Affidavit and Covering Letter, parent's on the other 3 |
| No siblings | The "and his/her sibling…" part is left out |
| Business income | "salary" wording becomes "revenue from the business" / "business" |
| S/O / D/O | Uses the sponsoring parent's name |
| Address | Written on one line in the letter body; split into lines (at commas, or as typed) under the signature |

## Setting up GitHub Pages (once)

1. On GitHub, open the repo's **Settings → Pages**.
2. Under **Build and deployment**, set **Source** to *Deploy from a branch*.
3. Pick the branch with this code (e.g. `main`) and folder `/ (root)`, then **Save**.

After a minute the page is at `https://<your-username>.github.io/docfiller/`.

## Changing a template

The original red-text letters are in `originals/`. To change the wording:

1. Edit the letter in `originals/` (keep the per-student parts in red).
2. If you added, removed or moved red parts, update that letter's list in
   `tools/build_templates.py` (one entry per red part, in order).
3. Run `npm run build` (needs Python 3 with `lxml`), then `npm test`.

To add a new letter, put it in `originals/` with a number prefix (e.g. `6_Name.docx`),
add an entry for it in `SPECS` in `tools/build_templates.py`, then rebuild and test. Any
detail the form doesn't ask for yet also needs a form box (`index.html`) and a value
(`computeValues` in `fill.js`).

Fixed text such as "Kerala", "2 years" or "Mumbai" is ordinary black text in the
templates. Edit it in `originals/` and rebuild.

## Files

| Path | What it is |
|---|---|
| `index.html`, `app.js` | The page: form, preview, downloads |
| `fill.js` | Wording rules, amounts in words, filling the .docx |
| `templates/`, `templates.js` | Generated placeholder templates (don't edit by hand) |
| `originals/` | Your original templates with red text |
| `tools/build_templates.py` | Turns `originals/` into `templates/` |
| `tests/` | `npm test` fills every letter for sample students and checks the wording |
| `vendor/jszip.min.js` | [JSZip](https://stuk.github.io/jszip/) 3.10.1 (MIT), for reading/writing .docx |
