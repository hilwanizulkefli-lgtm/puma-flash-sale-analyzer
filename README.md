# ⚡ Flash Sale Product Analyzer

A Streamlit app that analyzes which products qualify for a "Shocking Sale" /
flash sale, by joining 4 source files and applying 5 configurable rules.

## What it does

Upload these 4 files:

1. **zeCOM Tracker** — master product list with `Style#`, platform flags
   (Lazada/Shopee/TikTok), RRP, and a `Remark` column.
2. **Content / Master file** — maps each `Style#` to its EAN barcodes and
   sizes (used to expand a style into sellable units and detect cut sizes).
3. **Warehouse Inventory file** — stock quantity per EAN.
4. **Marketplace Price/Stock export** — current marketplace price, special
   price, and listed quantity per EAN (SellerSKU).

The app joins all 4 on `Style# ↔ Color_No ↔ EAN ↔ SellerSKU` and checks each
product against 5 rules (all adjustable from the sidebar):

| # | Rule | Default |
|---|------|---------|
| ① | Minimum warehouse stock | ≥ 50 units |
| ② | No cut sizes (S/M/L/XL etc.) | always on |
| ③ | Remark equals a required value | "Open for all" |
| ④ | Flash price discount vs RRP | ≥ 1% off |
| ⑤ | Marketplace stock as % of warehouse stock | ≥ 20% |

Results show pass/fail per product, fail reasons, charts, filters, and a CSV
export.

## Run locally

```bash
git clone <your-repo-url>
cd <your-repo-folder>
pip install -r requirements.txt
streamlit run app.py
```

The app opens at `http://localhost:8501`.

## Deploy on Streamlit Community Cloud (free)

1. Push this folder to a **public or private GitHub repo**.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with
   GitHub.
3. Click **"New app"**, select your repo, branch (`main`), and set the
   main file path to `app.py`.
4. Click **Deploy**. Your app gets a URL like
   `https://your-app-name.streamlit.app`.

Any time you push new commits to the repo, the deployed app auto-updates.

## File structure

```
.
├── app.py              # Main Streamlit app
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

## Notes on column matching

Each source file's columns are matched by **partial, case-insensitive name**
(e.g. any header containing "ean", "sellersku", "qty", "remark", etc.), so
minor header differences between export batches generally won't break the
app. If your file structure changes significantly (new platform export
template, renamed columns), check the `find_col()` candidate lists near the
top of `app.py` and add your column name variants there.
