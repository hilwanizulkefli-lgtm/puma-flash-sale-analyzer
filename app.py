"""
PUMA Flash Sale ("Shocking Sale") Product Analyzer
---------------------------------------------------
Joins 4 source files (zeCOM tracker, B2C inventory, Marketplace price/stock
export, Content/master file) and flags which products qualify for a
flash sale based on 5 configurable rules.

Run locally:
    streamlit run app.py

Deploy:
    Push this repo to GitHub, then deploy on https://share.streamlit.io
"""

import io
import re
from collections import Counter

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# Page setup
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Flash Sale Product Analyzer",
    page_icon="⚡",
    layout="wide",
)

CUT_SIZES = {
    "XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXXS",
    "2XL", "3XL", "4XL", "5XL", "6XL",
}

PLATFORM_COLORS = {"Lazada": "🟧", "Shopee": "🟥", "TikTok": "🟪"}


# --------------------------------------------------------------------------
# Helpers: flexible column finding (handles header name drift across files)
# --------------------------------------------------------------------------
def find_col(columns, candidates):
    """Return the first column name that contains any candidate substring."""
    cols_lower = {c: str(c).lower() for c in columns}
    for cand in candidates:
        for col, low in cols_lower.items():
            if cand.lower() in low:
                return col
    return None


def to_str_id(value):
    """Normalize EAN / SKU / Style values (handles floats like 123.0)."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def to_number(value, default=0.0):
    try:
        if pd.isna(value):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


# --------------------------------------------------------------------------
# File readers — each returns a tidy DataFrame regardless of source quirks
# --------------------------------------------------------------------------
def read_zecom_tracker(file, sheet_name=None):
    """
    zeCOM tracker: header row is NOT row 0 (it's usually row 3 / index 3),
    contains Style#, platform Y/N flags, RRP, and a Remark column
    (e.g. 'Open for all') that can appear in more than one place
    depending on the campaign section (Lazada/Zalora/TikTok vs Shopee).

    IMPORTANT: the actual "Remark" columns (containing values like "Open
    for all") sit under merged section headers one row ABOVE the column
    header row (e.g. a row-2 label "EXCLUSION" spanning the Lazada/Zalora/
    TikTok block and again for the Shopee Mega block). Because of the
    merge, row 3 (the column header row) is blank for these columns, so
    they show up as "Unnamed: N" after a normal pd.read_excel — they
    cannot be found by searching column names for "remark". Instead we
    scan the row directly above the header row for a label such as
    "EXCLUSION" or "REMARK" and use THOSE column positions.
    """
    xls = pd.ExcelFile(file)
    sheet = sheet_name or xls.sheet_names[0]
    raw = pd.read_excel(xls, sheet_name=sheet, header=None)

    # Auto-detect header row: the row that contains a cell equal to "Style#"
    header_row_idx = None
    for i in range(min(10, len(raw))):
        row_vals = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if any(v in ("style#", "style #", "style") for v in row_vals):
            header_row_idx = i
            break
    if header_row_idx is None:
        header_row_idx = 3  # fallback based on observed file structure

    # Find remark/exclusion columns from the section-label row directly
    # above the column header row (handles merged-cell blank headers).
    section_row_idx = max(header_row_idx - 1, 0)
    section_row = raw.iloc[section_row_idx]
    remark_col_positions = [
        i
        for i, v in enumerate(section_row)
        if pd.notna(v) and str(v).strip().upper() in ("EXCLUSION", "REMARK", "REMARKS")
    ]

    df = pd.read_excel(xls, sheet_name=sheet, header=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]

    style_col = find_col(df.columns, ["style#", "style #", "style"])
    desc_col = find_col(df.columns, ["description", "product name", "name"])
    rbu_col = find_col(df.columns, ["rbu"])
    div_col = find_col(df.columns, ["div"])
    gender_col = find_col(df.columns, ["gender"])
    article_col = find_col(df.columns, ["article type", "article_type"])
    rrp_col = find_col(df.columns, ["rrp"])
    lazada_col = find_col(df.columns, ["lazada"])
    shopee_col = find_col(df.columns, ["shopee"])
    tiktok_col = find_col(df.columns, ["tiktok", "tik tok"])

    # Also fall back to any column literally named "remark" in case a
    # future file version gives these columns a proper header.
    named_remark_cols = [c for c in df.columns if "remark" in str(c).lower()]
    remark_cols = (
        [df.columns[i] for i in remark_col_positions if i < len(df.columns)]
        or named_remark_cols
    )

    out = pd.DataFrame()
    out["style"] = df[style_col].apply(to_str_id) if style_col else ""
    out["description"] = df[desc_col].astype(str) if desc_col else ""
    out["rbu"] = df[rbu_col].astype(str) if rbu_col else ""
    out["division"] = df[div_col].astype(str) if div_col else ""
    out["gender"] = df[gender_col].astype(str) if gender_col else ""
    out["article_type"] = df[article_col].astype(str) if article_col else ""
    out["rrp"] = df[rrp_col].apply(to_number) if rrp_col else 0.0
    out["lazada_flag"] = (
        df[lazada_col].astype(str).str.upper().str.strip() if lazada_col else ""
    )
    out["shopee_flag"] = (
        df[shopee_col].astype(str).str.upper().str.strip() if shopee_col else ""
    )
    out["tiktok_flag"] = (
        df[tiktok_col].astype(str).str.upper().str.strip() if tiktok_col else ""
    )

    # Combine remark columns: take the first non-empty remark per row
    if remark_cols:
        remark_series = df[remark_cols].apply(
            lambda row: next(
                (str(v).strip() for v in row if pd.notna(v) and str(v).strip()), ""
            ),
            axis=1,
        )
        out["remark"] = remark_series
    else:
        out["remark"] = ""

    out = out[out["style"] != ""].reset_index(drop=True)
    return out


def read_content_file(file, sheet_name=None):
    """
    Content/master file: maps a Style/Color code -> list of EANs (one per
    size). Used to expand a Style into its EAN children, and to detect
    whether a style has "cut sizes" (apparel S/M/L/XL) vs straight sizes
    (footwear UK/EU numerics).

    IMPORTANT: this file typically has TWO size-related columns:
      - "Size No." — an internal numeric size code used for EVERY
        product including footwear (e.g. 110, 250, 90). This is NOT
        useful for cut-size detection since footwear also has numbers.
      - "Print Size Code (UK)" — the actual customer-facing size shown
        on the box/label. For apparel this is where letter sizes like
        XS/S/M/L/XL/XXL appear; for footwear it's a UK shoe size number.
    We must use the latter (the printed/customer-facing size) to detect
    cut sizes, not the generic internal size code.
    """
    xls = pd.ExcelFile(file)
    sheet = sheet_name or xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]

    style_col = find_col(df.columns, ["color_no", "color no", "style", "colorno"])
    ean_col = find_col(df.columns, ["ean"])
    # Prefer the customer-facing printed size code; fall back to a
    # generic "size" column only if the printed-size column is absent.
    size_col = find_col(
        df.columns,
        [
            "print size code",
            "print size",
            "size uk",
            "size_uk",
            "uk size",
            "size code",
        ],
    )
    if size_col is None:
        size_col = find_col(df.columns, ["size"])

    out = pd.DataFrame()
    out["style"] = df[style_col].apply(to_str_id) if style_col else ""
    out["ean"] = df[ean_col].apply(to_str_id) if ean_col else ""
    out["size"] = df[size_col].astype(str).str.strip().str.upper() if size_col else ""

    out = out[(out["style"] != "") & (out["ean"] != "")].reset_index(drop=True)
    return out


def read_inventory_file(file, sheet_name=None):
    """
    Warehouse / B2C channel inventory: one row per EAN (sometimes per
    EAN+location). We sum QtyAvailable per EAN.
    """
    xls = pd.ExcelFile(file)
    sheet = sheet_name or xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]

    ean_col = find_col(df.columns, ["sku", "ean"])
    qty_col = find_col(
        df.columns, ["qtyavailable", "qty available", "available", "qty"]
    )

    out = pd.DataFrame()
    out["ean"] = df[ean_col].apply(to_str_id) if ean_col else ""
    out["wh_qty"] = df[qty_col].apply(to_number) if qty_col else 0.0
    out = out[out["ean"] != ""]
    out = out.groupby("ean", as_index=False)["wh_qty"].sum()
    return out


def read_marketplace_price_stock(file, sheet_name=None):
    """
    Marketplace (Lazada/Shopee/TikTok) price & stock export. Column
    positions vary by platform export template, so we search by header
    name. Expected useful columns: SellerSKU (EAN), Quantity, Price,
    SpecialPrice / Promo Price.

    Lazada's bulk-edit template format has THREE rows below the real
    header that are not data: a 'Mandatory'/'Optional' row, a long-form
    field description row, and a short instruction row (e.g. "*Only
    positive numbers are accepted."). We skip these by requiring the EAN
    column to look like a real SKU/barcode (digits, optionally with
    separators) rather than free-text instructions.
    """
    xls = pd.ExcelFile(file)
    sheet = sheet_name or xls.sheet_names[0]

    # Some marketplace exports have a few junk rows above the real header.
    raw = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=10)
    header_row_idx = 0
    for i in range(len(raw)):
        row_vals = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if any("sellersku" in v or "seller sku" in v for v in row_vals):
            header_row_idx = i
            break

    df = pd.read_excel(xls, sheet_name=sheet, header=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]

    ean_col = find_col(df.columns, ["sellersku", "seller sku", "ean"])
    qty_col = find_col(df.columns, ["quantity", "qty", "stock"])
    price_col = find_col(df.columns, ["price"])
    special_col = find_col(
        df.columns, ["specialprice", "special price", "promo price", "saleprice"]
    )

    out = pd.DataFrame()
    out["ean"] = df[ean_col].apply(to_str_id) if ean_col else ""
    out["mp_qty"] = df[qty_col].apply(to_number) if qty_col else 0.0
    out["mp_price"] = df[price_col].apply(to_number) if price_col else 0.0
    out["mp_special_price"] = (
        df[special_col].apply(to_number) if special_col else 0.0
    )

    # Drop template instruction/placeholder rows and non-EAN identifiers.
    # Real EAN/UPC barcodes are pure digit strings, typically 8-14 digits
    # long. Some marketplace exports also contain internal composite IDs
    # like "1510456035-1770192599165-78" (productId-skuId-variantId) for
    # unmapped/orphan listings — these contain hyphens and are excluded.
    out = out[out["ean"] != ""]
    out = out[out["ean"].str.fullmatch(r"\d{8,14}")]

    out = out.groupby("ean", as_index=False).agg(
        {"mp_qty": "sum", "mp_price": "first", "mp_special_price": "first"}
    )
    return out


# --------------------------------------------------------------------------
# Core analysis
# --------------------------------------------------------------------------
def run_analysis(
    zecom_df: pd.DataFrame,
    content_df: pd.DataFrame,
    inventory_df: pd.DataFrame,
    marketplace_df: pd.DataFrame,
    stock_threshold: int,
    min_discount_pct: float,
    min_flash_stock_pct: float,
    required_remark: str,
):
    # 1) Expand each style into its EAN children via the content file
    style_eans = content_df.groupby("style")["ean"].apply(list).to_dict()
    style_sizes = content_df.groupby("style")["size"].apply(list).to_dict()

    inv_lookup = inventory_df.set_index("ean")["wh_qty"].to_dict()
    mp_lookup = marketplace_df.set_index("ean")[
        ["mp_qty", "mp_price", "mp_special_price"]
    ].to_dict("index")

    rows = []
    for _, prod in zecom_df.iterrows():
        style = prod["style"]
        eans = style_eans.get(style, [])
        sizes = style_sizes.get(style, [])

        wh_stock = sum(inv_lookup.get(e, 0) for e in eans)
        mp_records = [mp_lookup[e] for e in eans if e in mp_lookup]
        mp_stock = sum(r["mp_qty"] for r in mp_records)
        current_price = next(
            (r["mp_price"] for r in mp_records if r["mp_price"]), prod["rrp"]
        )
        flash_price = next(
            (r["mp_special_price"] for r in mp_records if r["mp_special_price"]),
            current_price,
        )

        rrp = prod["rrp"] or current_price or 0

        has_cut_size = any(s in CUT_SIZES for s in sizes)

        remark = str(prod["remark"]).strip()
        remark_ok = remark.lower() == required_remark.lower()

        discount_pct = (
            round((rrp - flash_price) / rrp * 100, 1) if rrp else 0.0
        )
        min_flash_stock = round(wh_stock * (min_flash_stock_pct / 100))

        fails = []
        if wh_stock < stock_threshold:
            fails.append(f"WH stock {int(wh_stock)} < {stock_threshold}")
        if has_cut_size:
            fails.append("Has cut sizes")
        if not remark_ok:
            fails.append(f"Remark: '{remark or '(empty)'}'")
        if rrp == 0 or discount_pct < min_discount_pct:
            fails.append(
                f"Flash price disc {discount_pct}% < {min_discount_pct}%"
            )
        if wh_stock > 0 and mp_stock < min_flash_stock:
            fails.append(
                f"Flash stock {int(mp_stock)} < {min_flash_stock_pct:.0f}% "
                f"of WH stock (need ≥{min_flash_stock})"
            )
        elif wh_stock == 0:
            fails.append("No warehouse stock")

        platforms = []
        if prod["lazada_flag"] == "YES":
            platforms.append("Lazada")
        if prod["shopee_flag"] == "YES":
            platforms.append("Shopee")
        if prod["tiktok_flag"] == "YES":
            platforms.append("TikTok")

        rows.append(
            {
                "Style": style,
                "Description": prod["description"],
                "RBU": prod["rbu"],
                "Gender": prod["gender"],
                "WH Stock": int(wh_stock),
                "Marketplace Stock": int(mp_stock),
                "Min Flash Stock": int(min_flash_stock),
                "RRP": round(rrp, 2),
                "Current Price": round(current_price, 2) if current_price else None,
                "Flash Price": round(flash_price, 2) if flash_price else None,
                "Discount %": discount_pct,
                "Platforms": ", ".join(platforms) if platforms else "—",
                "Remark": remark,
                "Qualifies": len(fails) == 0,
                "Fail Reasons": "; ".join(fails) if fails else "",
            }
        )

    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# Sidebar — rule configuration
# --------------------------------------------------------------------------
st.sidebar.title("⚡ Shocking Sale Rules")
st.sidebar.caption("All rules must pass for a product to qualify.")

stock_threshold = st.sidebar.number_input(
    "① Minimum warehouse stock (units)", min_value=0, value=50, step=10
)
required_remark = st.sidebar.text_input(
    "③ Required remark value", value="Open for all"
)
min_discount_pct = st.sidebar.number_input(
    "④ Minimum flash price discount vs RRP (%)",
    min_value=0.0,
    value=1.0,
    step=0.5,
)
min_flash_stock_pct = st.sidebar.number_input(
    "⑤ Minimum marketplace stock as % of warehouse stock",
    min_value=0.0,
    value=20.0,
    step=5.0,
)
st.sidebar.markdown("② No cut sizes (S/M/L/XL/etc.) — always checked")

st.sidebar.divider()
st.sidebar.caption(
    "Built for PUMA MY/PH/SG flash sale qualification. "
    "Upload your 4 source files in the main panel to begin."
)

# --------------------------------------------------------------------------
# Main panel — file upload
# --------------------------------------------------------------------------
st.title("⚡ Flash Sale Product Analyzer")
st.caption(
    "Upload your zeCOM tracker, Content file, Warehouse Inventory file, "
    "and Marketplace Price/Stock export. The analyzer joins all 4 and "
    "flags which products qualify for a Shocking Sale."
)

col1, col2 = st.columns(2)
with col1:
    zecom_file = st.file_uploader(
        "1️⃣ zeCOM Tracker file (.xlsx)", type=["xlsx", "xls"], key="zecom"
    )
    content_file = st.file_uploader(
        "2️⃣ Content / Master file (.xlsx)", type=["xlsx", "xls"], key="content"
    )
with col2:
    inventory_file = st.file_uploader(
        "3️⃣ Warehouse Inventory file (.xlsx)", type=["xlsx", "xls"], key="inventory"
    )
    marketplace_file = st.file_uploader(
        "4️⃣ Marketplace Price/Stock export (.xlsx)",
        type=["xlsx", "xls"],
        key="marketplace",
    )

run_btn = st.button(
    "🔍 Analyze products",
    type="primary",
    disabled=not all(
        [zecom_file, content_file, inventory_file, marketplace_file]
    ),
)

if not all([zecom_file, content_file, inventory_file, marketplace_file]):
    st.info("Upload all 4 files to enable analysis.")

# --------------------------------------------------------------------------
# Run analysis
# --------------------------------------------------------------------------
if run_btn:
    with st.spinner("Reading files and joining data..."):
        try:
            zecom_df = read_zecom_tracker(zecom_file)
            content_df = read_content_file(content_file)
            inventory_df = read_inventory_file(inventory_file)
            marketplace_df = read_marketplace_price_stock(marketplace_file)
        except Exception as e:
            st.error(f"Error reading files: {e}")
            st.stop()

        if zecom_df.empty:
            st.error(
                "Could not find a 'Style#' column in the zeCOM tracker file. "
                "Please check the file format."
            )
            st.stop()

        result_df = run_analysis(
            zecom_df,
            content_df,
            inventory_df,
            marketplace_df,
            stock_threshold,
            min_discount_pct,
            min_flash_stock_pct,
            required_remark,
        )

    st.session_state["result_df"] = result_df

# --------------------------------------------------------------------------
# Display results
# --------------------------------------------------------------------------
if "result_df" in st.session_state:
    result_df = st.session_state["result_df"]
    total = len(result_df)
    passed = int(result_df["Qualifies"].sum())
    failed = total - passed
    pass_rate = round(passed / total * 100, 1) if total else 0

    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total products", f"{total:,}")
    m2.metric("Qualify ✓", passed)
    m3.metric("Disqualified", f"{failed:,}")
    m4.metric("Pass rate", f"{pass_rate}%")

    # Fail reason breakdown
    fail_counter = Counter()
    for reasons in result_df.loc[~result_df["Qualifies"], "Fail Reasons"]:
        for r in reasons.split(";"):
            r = r.strip()
            if not r:
                continue
            if r.startswith("WH stock"):
                fail_counter["Low warehouse stock"] += 1
            elif r == "Has cut sizes":
                fail_counter["Has cut sizes"] += 1
            elif r.startswith("Remark"):
                fail_counter["Remark not matching"] += 1
            elif r.startswith("Flash price"):
                fail_counter["Flash price discount too low"] += 1
            elif r.startswith("Flash stock"):
                fail_counter["Flash stock below threshold"] += 1
            elif r == "No warehouse stock":
                fail_counter["No warehouse stock"] += 1

    if fail_counter:
        st.subheader("Fail reason breakdown")
        fail_df = pd.DataFrame(
            sorted(fail_counter.items(), key=lambda x: -x[1]),
            columns=["Reason", "Count"],
        )
        st.bar_chart(fail_df.set_index("Reason"))

    # RBU breakdown for qualifying products
    if passed:
        st.subheader("Qualifying products by RBU")
        rbu_pass = (
            result_df.loc[result_df["Qualifies"], "RBU"]
            .value_counts()
            .rename_axis("RBU")
            .reset_index(name="Count")
        )
        st.bar_chart(rbu_pass.set_index("RBU"))

    # Filters
    st.subheader("Product results")
    f1, f2, f3 = st.columns([1, 1, 2])
    with f1:
        status_filter = st.selectbox(
            "Status", ["All", "Qualifies only", "Disqualified only"]
        )
    with f2:
        rbu_options = ["All"] + sorted(result_df["RBU"].dropna().unique().tolist())
        rbu_filter = st.selectbox("RBU", rbu_options)
    with f3:
        search = st.text_input("Search style or description")

    filtered = result_df.copy()
    if status_filter == "Qualifies only":
        filtered = filtered[filtered["Qualifies"]]
    elif status_filter == "Disqualified only":
        filtered = filtered[~filtered["Qualifies"]]
    if rbu_filter != "All":
        filtered = filtered[filtered["RBU"] == rbu_filter]
    if search:
        mask = filtered["Style"].str.contains(
            search, case=False, na=False
        ) | filtered["Description"].str.contains(search, case=False, na=False)
        filtered = filtered[mask]

    st.dataframe(
        filtered,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Qualifies": st.column_config.CheckboxColumn("Qualifies"),
            "Discount %": st.column_config.NumberColumn(
                "Discount %", format="%.1f%%"
            ),
            "RRP": st.column_config.NumberColumn("RRP", format="%.2f"),
            "Current Price": st.column_config.NumberColumn(
                "Current Price", format="%.2f"
            ),
            "Flash Price": st.column_config.NumberColumn(
                "Flash Price", format="%.2f"
            ),
        },
    )

    csv_bytes = filtered.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download filtered results as CSV",
        data=csv_bytes,
        file_name="shocking_sale_results.csv",
        mime="text/csv",
    )
