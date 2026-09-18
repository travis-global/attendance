# EOCNS Hostel Attendance Report Builder

A one-page website that turns the hostel gate log into a ready-to-use
boarding fee absence report — no spreadsheet skills required. Built for
**Eastern Oak College of Nursing Sciences**, Abakaliki.

Upload the `.xls`/`.xlsx` file exported from the gate machine, set a date
range and the rate charged per absent day, and it generates a styled
`.xlsx` report (one row per student, one column per day, a Naira total
per student) and downloads it automatically.

Everything runs **entirely in the browser** — the file never leaves the
visitor's device, so it's safe to use on GitHub Pages with no backend
and no server costs.

## Live site

Once published (see below), the site will be at:

```
https://<your-github-username>.github.io/<repo-name>/
```

## What it expects in the uploaded file

The first row must contain these column headings (exactly as exported by
the gate machine):

| Column | Used for |
|---|---|
| `Department` | shown per student (not currently displayed in the report, kept for future use) |
| `Name` | student name |
| `No.` | student ID number |
| `Date/Time` | the check-in timestamp |
| `Status` | only rows marked `C/IN` count as present |
| `Location ID` | `1` → "Hostel 1 name", `2` → "Hostel 2 name", anything else → "Unknown" |

If a column is missing, the page tells the visitor exactly which one, in
plain language, instead of failing silently.

## Project structure

```
index.html      the page itself
css/style.css   Eastern Oak colours, type and layout
js/app.js       file parsing, the absence calculation, and the Excel export
```

No build step, no `npm install` — it's plain HTML/CSS/JS plus two
CDN-hosted libraries (SheetJS for reading the upload, ExcelJS for writing
the styled report).

## Publishing it on GitHub Pages

1. Create a new **public** repository on GitHub (e.g. `eocns-attendance`).
2. Push these files to the repository's default branch:
   ```bash
   cd eocns-attendance
   git init
   git add .
   git commit -m "Hostel attendance report builder"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git push -u origin main
   ```
3. On GitHub, open the repo → **Settings** → **Pages**.
4. Under "Build and deployment", set **Source** to "Deploy from a branch",
   branch **main**, folder **/ (root)**, then **Save**.
5. GitHub will publish the site at
   `https://<your-username>.github.io/<repo-name>/` within a minute or two.

No further configuration is needed — the page is fully static.

## Customising

- **Colours / logo**: `css/style.css` (`:root` variables) and the crest
  `<img>` in `index.html`, which currently links directly to Eastern
  Oak's own logo at `easternoak.org` so the file doesn't need to be
  duplicated in this repo.
- **Default rate per day**: the `value="15000"` on the rate field in
  `index.html`, and `RATE_DEFAULT` in `js/app.js`.
- **Required columns**: `REQUIRED_COLUMNS` in `js/app.js`.
