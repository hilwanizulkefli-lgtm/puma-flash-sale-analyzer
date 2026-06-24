"""
PUMA Flash Sale ("Shocking Sale") Product Analyzer
---------------------------------------------------
Joins 4 source files and produces:
  1. Style-level summary — which styles qualify and why
  2. EAN-level detail   — one row per EAN with flash price calculated

Flash price logic:
  - If SRP (Special/Sale Price) exists and > 0: Flash Price = SRP × (1 - markdown%)
  - If SRP = 0 or missing:                      Flash Price = RRP × (1 - markdown%)

5 qualifying rules (all configurable in sidebar):
  ① Warehouse stock ≥ threshold
  ② No cut sizes (S/M/L/XL etc.)
  ③ Remark = "Open for all"
  ④ Flash price must be at least X% lower than the base price (SRP or RRP)
  ⑤ Marketplace stock ≥ Y% of warehouse stock
"""

from collections import Counter

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# Page config
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="⚡ Flash Sale Analyzer",
    page_icon="⚡",
    layout="wide",
)

CUT_SIZES = {
    "XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXXS",
    "2XL", "3XL", "4XL", "5XL", "6XL",
}


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def find_col(columns, candidates):
    """Return the first column whose name contains any candidate substring."""
    cols_lower = {c: str(c).lower() for c in columns}
    for cand in candidates:
        for col, low in cols_lower.items():
            if cand.lower() in low:
                return col
    return None


def to_str_id(value):
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


def round2(v):
    return round(float(v), 2) if v else 0.0


# --------------------------------------------------------------------------
# File readers
# --------------------------------------------------------------------------
def read_zecom_tracker(file, sheet_name=None):
    """
    zeCOM tracker — multi-row header with merged cells.
    Remark columns (containing 'Open for all') sit under section-label
    rows labelled 'EXCLUSION' one row above the real column header row,
    so they appear as 'Unnamed: N' — we find them via the section row.
    """
    xls = pd.ExcelFile(file)
    sheet = sheet_name or xls.sheet_names[0]
    raw = pd.read_excel(xls, sheet_name=sheet, header=None)

    header_row_idx = None
    for i in range(min(10, len(raw))):
        row_vals = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if any(v in ("style#", "style #", "style") for v in row_vals):
            header_row_idx = i
            break
    if header_row_idx is None:
        header_row_idx = 3

    section_row = raw.iloc[max(header_row_idx - 1, 0)]
    remark_col_positions = [
        i for i, v in enumerate(section_row)
        if pd.notna(v) and str(v).strip().upper() in ("EXCLUSION", "REMARK", "REMARKS")
    ]

    df = pd.read_excel(xls, sheet_name=sheet, header=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]

    style_col   = find_col(df.columns, ["style#", "style #", "style"])
    desc_col    = find_col(df.columns, ["description", "product name", "name"])
    rbu_col     = find_col(df.columns, ["rbu"])
    gender_col  = find_col(df.columns, ["gender"])
    article_col = find_col(df.columns, ["article type", "article_type"])
    rrp_col     = find_col(df.columns, ["rrp"])
    lazada_col  = find_col(df.columns, ["lazada"])
    shopee_col  = find_col(df.columns, ["shopee"])
    tiktok_col  = find_col(df.columns, ["tiktok", "tik tok"])

    named_remark_cols = [c for c in df.columns if "remark" in str(c).lower()]
    remark_cols = (
        [df.columns[i] for i in remark_col_positions if i < len(df.columns)]
        or named_remark_cols
    )

    out = pd.DataFrame()
    out["style"]       = df[style_col].apply(to_str_id) if style_col else ""
    out["description"] = df[desc_col].astype(str) if desc_col else ""
    out["rbu"]         = df[rbu_col].astype(str) if rbu_col else ""
    out["gender"]      = df[gender_col].astype(str) if gender_col else ""
    out["article_type"]= df[article_col].astype(str) if article_col else ""
    out["rrp"]         = df[rrp_col].apply(to_number) if rrp_col else 0.0
    out["lazada_flag"] = df[lazada_col].astype(str).str.upper().str.strip() if lazada_col else ""
    out["shopee_flag"] = df[shopee_col].astype(str).str.upper().str.strip() if shopee_col else ""
    out["tiktok_flag"] = df[tiktok_col].astype(str).str.upper().str.strip() if tiktok_col else ""

    if remark_cols:
        out["remark"] = df[remark_cols].apply(
            lambda row: next(
                (str(v).strip() for v in row if pd.notna(v) and str(v).strip()), ""
            ), axis=1,
        )
    else:
        out["remark"] = ""

    return out[out["style"] != ""].reset_index(drop=True)


def read_content_file(file, sheet_name=None):
    """
    Content/master file: Style → EAN + size mapping.
    Uses 'Print Size Code (UK)' for cut-size detection (not 'Size No.'
    which is a generic internal numeric code shared by all product types).
    """
    xls = pd.ExcelFile(file)
    sheet = sheet_name or xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]

    style_col = find_col(df.columns, ["color_no", "color no", "colorno", "style"])
    ean_col   = find_col(df.columns, ["ean"])
    size_col  = find_col(df.columns, [
        "print size code", "print size", "size uk", "size_uk", "uk size", "size code",
    ]) or find_col(df.columns, ["size"])

    out = pd.DataFrame()
    out["style"] = df[style_col].apply(to_str_id) if style_col else ""
    out["ean"]   = df[ean_col].apply(to_str_id) if ean_col else ""
    out["size"]  = df[size_col].astype(str).str.strip().str.upper() if size_col else ""

    return out[(out["style"] != "") & (out["ean"] != "")].reset_index(drop=True)


def read_inventory_file(file, sheet_name=None):
    """Warehouse B2C inventory — sum QtyAvailable per EAN."""
    xls = pd.ExcelFile(file)
    sheet = sheet_name or xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]

    ean_col = find_col(df.columns, ["sku", "ean"])
    qty_col = find_col(df.columns, ["qtyavailable", "qty available", "available", "qty"])

    out = pd.DataFrame()
    out["ean"]    = df[ean_col].apply(to_str_id) if ean_col else ""
    out["wh_qty"] = df[qty_col].apply(to_number) if qty_col else 0.0
    out = out[out["ean"] != ""]
    return out.groupby("ean", as_index=False)["wh_qty"].sum()


def read_marketplace_price_stock(file, sheet_name=None):
    """
    Lazada/Shopee/TikTok price+stock export.
    Skips the 3 instruction rows below the real header row in Lazada's
    bulk-edit template by filtering to pure 8-14 digit EAN barcodes.
    Returns SRP = SpecialPrice (the current promotional/sale price).
    """
    xls = pd.ExcelFile(file)
    sheet = sheet_name or xls.sheet_names[0]

    raw = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=10)
    header_row_idx = 0
    for i in range(len(raw)):
        row_vals = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if any("sellersku" in v or "seller sku" in v for v in row_vals):
            header_row_idx = i
            break

    df = pd.read_excel(xls, sheet_name=sheet, header=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]

    ean_col     = find_col(df.columns, ["sellersku", "seller sku", "ean"])
    qty_col     = find_col(df.columns, ["quantity", "qty", "stock"])
    price_col   = find_col(df.columns, ["price"])
    special_col = find_col(df.columns, ["specialprice", "special price", "promo price", "saleprice"])

    out = pd.DataFrame()
    out["ean"]     = df[ean_col].apply(to_str_id) if ean_col else ""
    out["mp_qty"]  = df[qty_col].apply(to_number) if qty_col else 0.0
    out["mp_price"]= df[price_col].apply(to_number) if price_col else 0.0
    out["srp"]     = df[special_col].apply(to_number) if special_col else 0.0

    # Keep only real EAN barcodes (8–14 digits, no hyphens or text)
    out = out[out["ean"] != ""]
    out = out[out["ean"].str.fullmatch(r"\d{8,14}")]

    return out.groupby("ean", as_index=False).agg(
        {"mp_qty": "sum", "mp_price": "first", "srp": "first"}
    )


# --------------------------------------------------------------------------
# Flash price calculation (per EAN)
# --------------------------------------------------------------------------
def calc_flash_price(srp, rrp, markdown_pct):
    """
    Flash Price rule:
      - If SRP > 0: flash = SRP × (1 - markdown_pct/100)
      - If SRP = 0: flash = RRP × (1 - markdown_pct/100)
    Returns (base_price_used, flash_price, actual_discount_pct_vs_rrp)
    """
    factor = 1 - markdown_pct / 100
    if srp and srp > 0:
        base = srp
    elif rrp and rrp > 0:
        base = rrp
    else:
        return 0.0, 0.0, 0.0

    flash = round2(base * factor)
    disc_vs_rrp = round((rrp - flash) / rrp * 100, 1) if rrp else 0.0
    return base, flash, disc_vs_rrp


# --------------------------------------------------------------------------
# Core analysis — returns (style_df, ean_df)
# --------------------------------------------------------------------------
def run_analysis(
    zecom_df, content_df, inventory_df, marketplace_df,
    stock_threshold, markdown_pct, min_flash_stock_pct, required_remark,
):
    style_eans  = content_df.groupby("style")["ean"].apply(list).to_dict()
    style_sizes = content_df.groupby("style")["size"].apply(list).to_dict()
    ean_size    = content_df.set_index("ean")["size"].to_dict()

    inv_lookup = inventory_df.set_index("ean")["wh_qty"].to_dict()
    mp_lookup  = marketplace_df.set_index("ean")[
        ["mp_qty", "mp_price", "srp"]
    ].to_dict("index")

    style_rows = []
    ean_rows   = []

    for _, prod in zecom_df.iterrows():
        style  = prod["style"]
        rrp    = prod["rrp"] or 0.0
        eans   = style_eans.get(style, [])
        sizes  = style_sizes.get(style, [])

        wh_stock  = sum(inv_lookup.get(e, 0) for e in eans)
        mp_records= [mp_lookup[e] for e in eans if e in mp_lookup]
        mp_stock  = sum(r["mp_qty"] for r in mp_records)

        has_cut_size = any(s in CUT_SIZES for s in sizes)
        remark       = str(prod["remark"]).strip()
        remark_ok    = remark.lower() == required_remark.lower()
        min_flash_st = round(wh_stock * (min_flash_stock_pct / 100))

        platforms = []
        if prod["lazada_flag"] == "YES": platforms.append("Lazada")
        if prod["shopee_flag"] == "YES": platforms.append("Shopee")
        if prod["tiktok_flag"] == "YES": platforms.append("TikTok")

        # Style-level fails (rules ①②③⑤)
        style_fails = []
        if wh_stock < stock_threshold:
            style_fails.append(f"WH stock {int(wh_stock)} < {stock_threshold}")
        elif wh_stock == 0:
            style_fails.append("No warehouse stock")
        if has_cut_size:
            style_fails.append("Has cut sizes")
        if not remark_ok:
            style_fails.append(f"Remark: '{remark or '(empty)'}'")
        if wh_stock > 0 and mp_stock < min_flash_st:
            style_fails.append(
                f"Marketplace stock {int(mp_stock)} < {min_flash_stock_pct:.0f}% "
                f"of WH (need ≥{min_flash_st})"
            )

        # Representative price for style summary (first EAN with data)
        rep = next((mp_lookup[e] for e in eans if e in mp_lookup), {})
        rep_srp  = rep.get("srp", 0.0) or 0.0
        _, rep_flash, rep_disc = calc_flash_price(rep_srp, rrp, markdown_pct)

        style_rows.append({
            "Style":            style,
            "Description":      prod["description"],
            "RBU":              prod["rbu"],
            "Gender":           prod["gender"],
            "Article Type":     prod["article_type"],
            "Platforms":        ", ".join(platforms) or "—",
            "RRP":              round2(rrp),
            "SRP":              round2(rep_srp),
            "Flash Price":      rep_flash,
            "Disc % vs RRP":    rep_disc,
            "WH Stock":         int(wh_stock),
            "Marketplace Stock":int(mp_stock),
            "Min Flash Stock":  int(min_flash_st),
            "Remark":           remark,
            "Qualifies":        len(style_fails) == 0,
            "Fail Reasons":     "; ".join(style_fails) if style_fails else "",
        })

        # EAN-level rows — only for qualifying styles
        if len(style_fails) == 0:
            for ean in eans:
                mp  = mp_lookup.get(ean, {})
                srp = mp.get("srp", 0.0) or 0.0
                qty = mp.get("mp_qty", 0.0)
                wh  = inv_lookup.get(ean, 0.0)
                base, flash, disc = calc_flash_price(srp, rrp, markdown_pct)
                size = ean_size.get(ean, "")

                flash_stock = int(wh * (min_flash_stock_pct / 100))  # always round down

                ean_rows.append({
                    "Style":             style,
                    "Description":       prod["description"],
                    "RBU":               prod["rbu"],
                    "Gender":            prod["gender"],
                    "EAN":               ean,
                    "Size":              size,
                    "Platforms":         ", ".join(platforms) or "—",
                    "RRP":               round2(rrp),
                    "SRP":               round2(srp),
                    "Base Price Used":   round2(base),
                    "Flash Price":       flash,
                    "Disc % vs RRP":     disc,
                    "WH Stock":          int(wh),
                    "Flash Stock (20%)": flash_stock,
                    "Marketplace Stock": int(qty),
                })

    return pd.DataFrame(style_rows), pd.DataFrame(ean_rows)


# --------------------------------------------------------------------------
# Sidebar — rules
# --------------------------------------------------------------------------
st.sidebar.title("⚡ Shocking Sale Rules")
st.sidebar.caption("All 5 rules must pass for a style to qualify.")

st.sidebar.markdown("**① Warehouse stock**")
stock_threshold = st.sidebar.number_input(
    "Minimum warehouse stock (units)", min_value=0, value=50, step=10,
    label_visibility="collapsed"
)

st.sidebar.markdown("**② No cut sizes** — always checked (S/M/L/XL etc.)")

st.sidebar.markdown("**③ Remark**")
required_remark = st.sidebar.text_input(
    "Required remark value", value="Open for all",
    label_visibility="collapsed"
)

st.sidebar.markdown("**④ Flash price markdown %**")
st.sidebar.caption(
    "Flash price = SRP × (1 − this %) if SRP > 0, else RRP × (1 − this %)"
)
markdown_pct = st.sidebar.number_input(
    "Price markdown % to apply", min_value=0.0, max_value=100.0,
    value=1.0, step=0.5, label_visibility="collapsed"
)

st.sidebar.markdown("**⑤ Marketplace stock**")
min_flash_stock_pct = st.sidebar.number_input(
    "Min marketplace stock as % of WH stock", min_value=0.0,
    value=20.0, step=5.0, label_visibility="collapsed"
)

st.sidebar.divider()
st.sidebar.caption("PUMA MY / PH / SG flash sale qualifier")


# --------------------------------------------------------------------------
# Main — header
# --------------------------------------------------------------------------
st.title("⚡ Flash Sale Product Analyzer")
st.caption(
    "Upload your 4 source files. The app joins them, applies your 5 rules, "
    "and outputs a full EAN-level list with calculated flash prices."
)

# --------------------------------------------------------------------------
# File upload
# --------------------------------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    zecom_file     = st.file_uploader("1️⃣ zeCOM Tracker (.xlsx)", type=["xlsx","xls"], key="zecom")
    content_file   = st.file_uploader("2️⃣ Content / Master file (.xlsx)", type=["xlsx","xls"], key="content")
with col2:
    inventory_file = st.file_uploader("3️⃣ Warehouse Inventory (.xlsx)", type=["xlsx","xls"], key="inv")
    marketplace_file = st.file_uploader("4️⃣ Marketplace Price/Stock export (.xlsx)", type=["xlsx","xls"], key="mp")

all_uploaded = all([zecom_file, content_file, inventory_file, marketplace_file])

if not all_uploaded:
    st.info("Upload all 4 files to enable analysis.")

run_btn = st.button("🔍 Analyze products", type="primary", disabled=not all_uploaded)

# --------------------------------------------------------------------------
# Run
# --------------------------------------------------------------------------
if run_btn:
    with st.spinner("Reading and joining all 4 files…"):
        try:
            zecom_df       = read_zecom_tracker(zecom_file)
            content_df     = read_content_file(content_file)
            inventory_df   = read_inventory_file(inventory_file)
            marketplace_df = read_marketplace_price_stock(marketplace_file)
        except Exception as e:
            st.error(f"Error reading files: {e}")
            st.stop()

        if zecom_df.empty:
            st.error("Could not find 'Style#' column in zeCOM tracker. Check file format.")
            st.stop()

        style_df, ean_df = run_analysis(
            zecom_df, content_df, inventory_df, marketplace_df,
            stock_threshold, markdown_pct, min_flash_stock_pct, required_remark,
        )

    st.session_state["style_df"] = style_df
    st.session_state["ean_df"]   = ean_df

# --------------------------------------------------------------------------
# Display
# --------------------------------------------------------------------------
if "style_df" in st.session_state:
    style_df = st.session_state["style_df"]
    ean_df   = st.session_state["ean_df"]

    total   = len(style_df)
    passed  = int(style_df["Qualifies"].sum())
    failed  = total - passed
    rate    = round(passed / total * 100, 1) if total else 0

    st.divider()

    # Metrics
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total styles",   f"{total:,}")
    c2.metric("Qualify ✓",      passed)
    c3.metric("Disqualified",   f"{failed:,}")
    c4.metric("Pass rate",      f"{rate}%")
    c5.metric("Qualifying EANs",f"{len(ean_df):,}")

    # Fail reason chart
    fail_counter = Counter()
    for reasons in style_df.loc[~style_df["Qualifies"], "Fail Reasons"]:
        for r in reasons.split(";"):
            r = r.strip()
            if not r: continue
            if r.startswith("WH stock") or r == "No warehouse stock":
                fail_counter["Low / no warehouse stock"] += 1
            elif r == "Has cut sizes":
                fail_counter["Has cut sizes"] += 1
            elif r.startswith("Remark"):
                fail_counter["Remark not matching"] += 1
            elif r.startswith("Marketplace stock"):
                fail_counter["Marketplace stock too low"] += 1

    if fail_counter:
        st.subheader("Fail reason breakdown")
        st.bar_chart(
            pd.DataFrame(
                sorted(fail_counter.items(), key=lambda x: -x[1]),
                columns=["Reason","Count"]
            ).set_index("Reason")
        )

    if passed:
        st.subheader("Qualifying styles by RBU")
        st.bar_chart(
            style_df[style_df["Qualifies"]]["RBU"]
            .value_counts().rename_axis("RBU").reset_index(name="Count")
            .set_index("RBU")
        )

    # ---- TABS: Style summary | EAN detail ----
    tab1, tab2 = st.tabs(["📦 Style summary", "🏷️ EAN detail (flash price list)"])

    with tab1:
        st.caption("All styles — filter by status, RBU, or search.")
        f1, f2, f3 = st.columns([1, 1, 2])
        with f1:
            status_f = st.selectbox("Status", ["All","Qualifies only","Disqualified only"], key="sf1")
        with f2:
            rbu_opts = ["All"] + sorted(style_df["RBU"].dropna().unique().tolist())
            rbu_f    = st.selectbox("RBU", rbu_opts, key="sf2")
        with f3:
            search_f = st.text_input("Search style / description", key="sf3")

        view = style_df.copy()
        if status_f == "Qualifies only":   view = view[view["Qualifies"]]
        if status_f == "Disqualified only":view = view[~view["Qualifies"]]
        if rbu_f != "All":                 view = view[view["RBU"] == rbu_f]
        if search_f:
            m = (view["Style"].str.contains(search_f, case=False, na=False) |
                 view["Description"].str.contains(search_f, case=False, na=False))
            view = view[m]

        st.dataframe(
            view, use_container_width=True, hide_index=True,
            column_config={
                "Qualifies":    st.column_config.CheckboxColumn("Qualifies"),
                "RRP":          st.column_config.NumberColumn("RRP",          format="MYR %.2f"),
                "SRP":          st.column_config.NumberColumn("SRP",          format="MYR %.2f"),
                "Flash Price":  st.column_config.NumberColumn("Flash Price",  format="MYR %.2f"),
                "Disc % vs RRP":st.column_config.NumberColumn("Disc % vs RRP",format="%.1f%%"),
            }
        )
        st.download_button(
            "⬇️ Download style summary CSV",
            data=view.to_csv(index=False).encode("utf-8"),
            file_name="flash_sale_style_summary.csv", mime="text/csv", key="dl1"
        )

    with tab2:
        st.caption(
            f"**{len(ean_df):,} EANs** from qualifying styles — each row shows the "
            "calculated flash price. Flash Price = SRP × (1 − markdown%) if SRP > 0, "
            "else RRP × (1 − markdown%)."
        )

        e1, e2, e3 = st.columns([1, 1, 2])
        with e1:
            rbu_opts2 = ["All"] + sorted(ean_df["RBU"].dropna().unique().tolist()) if not ean_df.empty else ["All"]
            rbu_f2    = st.selectbox("RBU", rbu_opts2, key="ef1")
        with e2:
            plat_opts = ["All","Lazada","Shopee","TikTok"]
            plat_f    = st.selectbox("Platform", plat_opts, key="ef2")
        with e3:
            search_f2 = st.text_input("Search style / description / EAN", key="ef3")

        ean_view = ean_df.copy() if not ean_df.empty else ean_df
        if not ean_view.empty:
            if rbu_f2 != "All":  ean_view = ean_view[ean_view["RBU"] == rbu_f2]
            if plat_f != "All":  ean_view = ean_view[ean_view["Platforms"].str.contains(plat_f, na=False)]
            if search_f2:
                m2 = (ean_view["Style"].str.contains(search_f2, case=False, na=False) |
                      ean_view["Description"].str.contains(search_f2, case=False, na=False) |
                      ean_view["EAN"].str.contains(search_f2, case=False, na=False))
                ean_view = ean_view[m2]

        if ean_view.empty:
            st.info("No qualifying EANs yet — run the analysis first.")
        else:
            st.dataframe(
                ean_view, use_container_width=True, hide_index=True,
                column_config={
                    "RRP":               st.column_config.NumberColumn("RRP",               format="MYR %.2f"),
                    "SRP":               st.column_config.NumberColumn("SRP",               format="MYR %.2f"),
                    "Base Price Used":   st.column_config.NumberColumn("Base Price Used",   format="MYR %.2f"),
                    "Flash Price":       st.column_config.NumberColumn("Flash Price",       format="MYR %.2f"),
                    "Disc % vs RRP":     st.column_config.NumberColumn("Disc % vs RRP",    format="%.1f%%"),
                    "WH Stock":          st.column_config.NumberColumn("WH Stock"),
                    "Flash Stock (20%)": st.column_config.NumberColumn("Flash Stock (20%)", help="Floor of 20% of this EAN's warehouse stock"),
                    "Marketplace Stock": st.column_config.NumberColumn("Marketplace Stock"),
                }
            )
            st.download_button(
                "⬇️ Download EAN flash price list CSV",
                data=ean_view.to_csv(index=False).encode("utf-8"),
                file_name="flash_sale_ean_list.csv", mime="text/csv", key="dl2"
            )
