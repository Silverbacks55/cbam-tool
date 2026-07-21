# CBAM Self-Assessment Tool (2026 definitive regime)

Single-page web tool: scope check (CN code / origin / purpose / 50-tonne rule)
plus a certificate cost estimator using the Commission's official benchmarks
and default values.

## Files
- `index.html` — the tool (all logic and styling)
- `data/assessment-data.json` — CN codes, reporting fields, country list
- `data/cost-data.json` — benchmarks (Reg 2025/2620) and default values (Reg 2025/2621)
- `build_data.py` — regenerates both JSON files from the Commission's Excel files

## Hosting (Squarespace can't host these files itself)
1. Go to https://app.netlify.com/drop (free, no account needed to start)
2. Drag this whole folder onto the page
3. You get a URL like `https://something.netlify.app` — the tool is live there

## Embedding in Squarespace
Add a **Code** block (requires Business plan or higher for scripts) with:

    <iframe id="cbam-tool" src="https://YOUR-SITE.netlify.app/"
            style="width:100%;border:0;min-height:900px"
            title="CBAM Self-Assessment"></iframe>
    <script>
      window.addEventListener('message', function (e) {
        if (e.data && e.data.height) {
          document.getElementById('cbam-tool').style.height = e.data.height + 'px';
        }
      });
    </script>

On a Personal plan (no scripts allowed), use just the iframe line with a fixed
height, e.g. style="width:100%;border:0;height:1600px".

## Updating the data
When the Commission publishes new files
(https://taxation-customs.ec.europa.eu/carbon-border-adjustment-mechanism/cbam-legislation-and-guidance_en):

    pip install openpyxl
    python3 build_data.py \
        --assessment "CBAM Self Assessment Tool Version 1.1.xlsx" \
        --benchmarks "CBAM Benchmarks_20260206.xlsx" \
        --dvs        "DVs as adopted_v20260204 .xlsx"

Then re-upload the folder to Netlify (drag it onto your site's Deploys page).
`index.html` never needs to change for a data update.

## Notes
- Testing locally: run `python3 -m http.server` in this folder and open
  http://localhost:8000 (opening index.html directly from disk blocks the
  data fetch).
- The tool is informational, not legal advice; only the Official Journal of
  the EU is authentic. Keep the disclaimer footer intact.
