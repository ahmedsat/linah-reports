# Linah Reports

Public GitHub Pages site for Linah Farms reports.

**Site:** <https://ahmedsat.github.io/linah-reports/>

## Reports

- [Medjool Report — 2026](https://ahmedsat.github.io/linah-reports/2026/medjool-report/)
- [Medjool Outsource Report — 2026](https://ahmedsat.github.io/linah-reports/2026/medjool-outsource-report/)

## Local preview

From the repository root:

```sh
python3 -m http.server 8000 --directory site
```

Then open <http://localhost:8000/>.

## Updating report data

Open a report with `?admin` to upload an Excel file and download a replacement `data.json`. Replace the file under `site/2026/<report>/`, then commit and push to `main`. GitHub Pages publishes the update automatically.

## Layout

- `site/` — GitHub Pages source
- `site/vendor/` — shared local JS, CSS, and font assets
- `outsource-medjool/` — Excel tool for the outsource report
- `.github/workflows/deploy.yml` — Pages deployment workflow

The `transformer/` project is maintained separately at <https://github.com/ahmedsat/transformer>.
