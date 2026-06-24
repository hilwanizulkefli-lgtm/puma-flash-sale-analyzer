"""
PUMA Flash Sale ("Shocking Sale") Product Analyzer
---------------------------------------------------
Joins 4 source files and produces:
  1. Style-level summary — which styles qualify and why
  2. EAN-level detail   — one row per EAN with flash price & flash stock

Qualifying rules (all configurable in sidebar):
  ① Marketplace     — Lazada / Shopee / TikTok (single select)
  ② Country         — MY / PH / SG (single select, reads correct zeCOM sheet)
  ③ Price tier      — RRP or SRP as flash base price (user picks)
  ④ Price markdown  — % to apply to chosen price tier
  ⑤ No cut sizes    — style must have ≥ N sizes with main WH stock > 0
  ⑥ Remark          — must equal required value (default: "Open for all")
  ⑦ Flash stock     — % of each EAN's main WH stock (floor)
"""

from collections import Counter
import pandas as pd
import streamlit as st

st.set_page_config(page_title="⚡ Flash Sale Analyzer", page_icon="⚡", layout="wide")


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def find_col(columns, candidates):
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
        return default if pd.isna(value) else float(value)
    except (TypeError, ValueError):
        return default


def r2(v):
    return round(float(v), 2) if v else 0.0


# --------------------------------------------------------------------------
# File readers
# --------------------------------------------------------------------------
def read_zecom_tracker(file, country="MY"):
    """
    Reads zeCOM tracker for the selected country sheet (MY / PH / SG).
    Header row detected by finding 'Style#'. Remark columns found via
    the section-label row above (labelled 'EXCLUSION' in merged cells).
    Also reads RRP and MD Price for price tier selection.
    """
    xls   = pd.ExcelFile(file)
    sheet = next(
        (s for s in xls.sheet_names if country.upper() in s.upper()),
        xls.sheet_names[0]
    )

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
    md_col      = find_col(df.columns, ["md price", "md_price", "markdown price"])
    lazada_col  = find_col(df.columns, ["lazada"])
    shopee_col  = find_col(df.columns, ["shopee"])
    tiktok_col  = find_col(df.columns, ["tiktok", "tik tok"])

    named_remark_cols = [c for c in df.columns if "remark" in str(c).lower()]
    remark_cols = (
        [df.columns[i] for i in remark_col_positions if i < len(df.columns)]
        or named_remark_cols
    )

    out = pd.DataFrame()
    out["style"]        = df[style_col].apply(to_str_id) if style_col else ""
    out["description"]  = df[desc_col].astype(str)       if desc_col  else ""
    out["rbu"]          = df[rbu_col].astype(str)         if rbu_col   else ""
    out["gender"]       = df[gender_col].astype(str)      if gender_col else ""
    out["article_type"] = df[article_col].astype(str)     if article_col else ""
    out["rrp"]          = df[rrp_col].apply(to_number)    if rrp_col   else 0.0
    out["md_price"]     = df[md_col].apply(to_number)     if md_col    else 0.0
    out["lazada_flag"]  = df[lazada_col].astype(str).str.upper().str.strip() if lazada_col else ""
    out["shopee_flag"]  = df[shopee_col].astype(str).str.upper().str.strip() if shopee_col else ""
    out["tiktok_flag"]  = df[tiktok_col].astype(str).str.upper().str.strip() if tiktok_col else ""

    if remark_cols:
        out["remark"] = df[remark_cols].apply(
            lambda row: next(
                (str(v).strip() for v in row if pd.notna(v) and str(v).strip()), ""
            ), axis=1,
        )
    else:
        out["remark"] = ""

    return out[out["style"] != ""].reset_index(drop=True), sheet


def read_content_file(file):
    xls   = pd.ExcelFile(file)
    df    = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    df.columns = [str(c).strip() for c in df.columns]

    style_col = find_col(df.columns, ["color_no", "color no", "colorno", "style"])
    ean_col   = find_col(df.columns, ["ean"])
    size_col  = find_col(df.columns, [
        "print size code", "print size", "size uk", "size_uk", "uk size", "size code",
    ]) or find_col(df.columns, ["size"])

    out = pd.DataFrame()
    out["style"] = df[style_col].apply(to_str_id) if style_col else ""
    out["ean"]   = df[ean_col].apply(to_str_id)   if ean_col   else ""
    out["size"]  = df[size_col].astype(str).str.strip().str.upper() if size_col else ""

    return out[(out["style"] != "") & (out["ean"] != "")].reset_index(drop=True)


def read_inventory_file(file):
    xls   = pd.ExcelFile(file)
    df    = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    df.columns = [str(c).strip() for c in df.columns]

    ean_col = find_col(df.columns, ["sku", "ean"])
    qty_col = find_col(df.columns, ["qtyavailable", "qty available", "available", "qty"])

    out = pd.DataFrame()
    out["ean"]    = df[ean_col].apply(to_str_id) if ean_col else ""
    out["wh_qty"] = df[qty_col].apply(to_number) if qty_col else 0.0
    out = out[out["ean"] != ""]
    return out.groupby("ean", as_index=False)["wh_qty"].sum()


def read_marketplace_file(file):
    """
    Reads Lazada/Shopee/TikTok price+stock export.
    Lazada template has 3 instruction rows below the header — filtered
    out by requiring pure 8–14 digit EAN barcodes.
    """
    xls   = pd.ExcelFile(file)
    sheet = xls.sheet_names[0]

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
    out["ean"]      = df[ean_col].apply(to_str_id)   if ean_col     else ""
    out["mp_qty"]   = df[qty_col].apply(to_number)   if qty_col     else 0.0
    out["mp_price"] = df[price_col].apply(to_number) if price_col   else 0.0
    out["srp"]      = df[special_col].apply(to_number) if special_col else 0.0

    out = out[out["ean"] != ""]
    out = out[out["ean"].str.fullmatch(r"\d{8,14}")]
    return out.groupby("ean", as_index=False).agg(
        {"mp_qty": "sum", "mp_price": "first", "srp": "first"}
    )


# --------------------------------------------------------------------------
# Flash price calculation
# --------------------------------------------------------------------------
def calc_flash_price(base_price, markdown_pct, rrp):
    """
    Flash Price = base_price × (1 − markdown_pct / 100)
    Also returns discount % vs RRP for reference.
    """
    if not base_price or base_price <= 0:
        return 0.0, 0.0
    flash       = r2(base_price * (1 - markdown_pct / 100))
    disc_vs_rrp = round((rrp - flash) / rrp * 100, 1) if rrp else 0.0
    return flash, disc_vs_rrp


# --------------------------------------------------------------------------
# Core analysis
# --------------------------------------------------------------------------
def run_analysis(
    zecom_df, content_df, inventory_df, marketplace_df,
    marketplace, price_tier, markdown_pct,
    min_sizes_with_stock, required_remark, flash_stock_pct,
):
    flag_col_map = {
        "Lazada": "lazada_flag",
        "Shopee": "shopee_flag",
        "TikTok": "tiktok_flag",
    }
    flag_col = flag_col_map[marketplace]

    style_eans = content_df.groupby("style")["ean"].apply(list).to_dict()
    ean_size   = content_df.set_index("ean")["size"].to_dict()
    inv_lookup = inventory_df.set_index("ean")["wh_qty"].to_dict()
    mp_lookup  = marketplace_df.set_index("ean")[
        ["mp_qty", "mp_price", "srp"]
    ].to_dict("index")

    style_rows, ean_rows = [], []

    for _, prod in zecom_df.iterrows():
        style = prod["style"]
        rrp   = prod["rrp"]      or 0.0
        md    = prod["md_price"] or 0.0

        eans   = style_eans.get(style, [])
        ean_wh = {e: inv_lookup.get(e, 0.0) for e in eans}
        wh_total = sum(ean_wh.values())

        # ① Marketplace flag check
        on_platform = prod[flag_col] == "YES"

        # ③ Resolve base price from chosen price tier
        if price_tier == "RRP":
            base_price = rrp
        else:  # SRP — take from marketplace file, fall back to RRP
            base_price = next(
                (mp_lookup[e]["srp"] for e in eans
                 if e in mp_lookup and mp_lookup[e]["srp"]),
                0.0
            ) or rrp

        # ⑤ Cut-size rule: count sizes with main WH stock > 0
        sizes_with_stock = sum(1 for e in eans if ean_wh[e] > 0)
        has_cut_size     = (len(eans) > 0) and (sizes_with_stock < min_sizes_with_stock)

        # ⑦ Flash stock
        total_flash_stock = sum(int(ean_wh[e] * flash_stock_pct / 100) for e in eans)
        mp_records        = [mp_lookup[e] for e in eans if e in mp_lookup]
        mp_stock_total    = sum(r["mp_qty"] for r in mp_records)

        # ⑥ Remark
        remark    = str(prod["remark"]).strip()
        remark_ok = remark.lower() == required_remark.lower()

        # ④ Flash price
        flash_price, disc_vs_rrp = calc_flash_price(base_price, markdown_pct, rrp)
        price_ok = flash_price > 0

        # Collect fails
        fails = []
        if not on_platform:
            fails.append(f"Not listed on {marketplace}")
        if has_cut_size:
            fails.append(
                f"Cut sizes — {sizes_with_stock} size(s) with stock "
                f"(need ≥ {min_sizes_with_stock})"
            )
        if not remark_ok:
            fails.append(f"Remark: '{remark or '(empty)'}'")
        if not price_ok:
            fails.append(f"{price_tier} missing/zero — cannot calculate flash price")
        if wh_total > 0 and mp_stock_total < total_flash_stock:
            fails.append(
                f"Platform stock {int(mp_stock_total)} < flash stock needed "
                f"({int(total_flash_stock)})"
            )
        if wh_total == 0:
            fails.append("No main WH stock")

        style_rows.append({
            "Style":                       style,
            "Description":                 prod["description"],
            "RBU":                         prod["rbu"],
            "Gender":                      prod["gender"],
            "Article Type":                prod["article_type"],
            "Marketplace":                 marketplace,
            "RRP":                         r2(rrp),
            "MD Price":                    r2(md),
            "Price Tier Used":             price_tier,
            "Base Price":                  r2(base_price),
            "Flash Price":                 flash_price,
            "Disc % vs RRP":               disc_vs_rrp,
            "Sizes with Stock":            sizes_with_stock,
            "Main WH Stock":               int(wh_total),
            "Total Flash Stock to Submit": int(total_flash_stock),
            "Remark":                      remark,
            "Qualifies":                   len(fails) == 0,
            "Fail Reasons":                "; ".join(fails) if fails else "",
        })

        # EAN-level rows — qualifying styles only
        if len(fails) == 0:
            for ean in eans:
                mp       = mp_lookup.get(ean, {})
                srp      = mp.get("srp", 0.0) or 0.0
                wh       = ean_wh[ean]
                size     = ean_size.get(ean, "")
                ean_base = srp if (price_tier == "SRP" and srp > 0) else rrp
                f_price, f_disc = calc_flash_price(ean_base, markdown_pct, rrp)
                flash_stock = int(wh * flash_stock_pct / 100)

                ean_rows.append({
                    "Style":                 style,
                    "Description":           prod["description"],
                    "RBU":                   prod["rbu"],
                    "Gender":                prod["gender"],
                    "EAN":                   ean,
                    "Size":                  size,
                    "Marketplace":           marketplace,
                    "RRP":                   r2(rrp),
                    "MD Price":              r2(md),
                    "SRP":                   r2(srp),
                    "Price Tier Used":       price_tier,
                    "Base Price":            r2(ean_base),
                    "Flash Price":           f_price,
                    "Disc % vs RRP":         f_disc,
                    "Main WH Stock":         int(wh),
                    "Flash Stock to Submit": flash_stock,
                })

    return pd.DataFrame(style_rows), pd.DataFrame(ean_rows)


# --------------------------------------------------------------------------
# Sidebar — configuration
# --------------------------------------------------------------------------
st.sidebar.title("⚡ Flash Sale Rules")
st.sidebar.caption("Configure all rules, then upload files and click Analyze.")

st.sidebar.markdown("---")
st.sidebar.markdown("### 🌏 Eligibility")

st.sidebar.markdown("**① Marketplace**")
marketplace = st.sidebar.selectbox(
    "Marketplace", ["Lazada", "Shopee", "TikTok"],
    label_visibility="collapsed"
)

st.sidebar.markdown("**② Country**")
country = st.sidebar.selectbox(
    "Country", ["MY", "PH", "SG"],
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 💰 Price")

st.sidebar.markdown("**③ Price tier (flash base price)**")
st.sidebar.caption(
    "RRP = full tag price (from zeCOM). "
    "SRP = current sale price (from marketplace export)."
)
price_tier = st.sidebar.selectbox(
    "Price tier", ["RRP", "SRP"],
    label_visibility="collapsed"
)

st.sidebar.markdown("**④ Price markdown %**")
st.sidebar.caption(f"Flash Price = {price_tier} × (1 − this %)")
markdown_pct = st.sidebar.number_input(
    "Markdown %", min_value=0.0, max_value=100.0,
    value=1.0, step=0.5, label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📦 Stock")

st.sidebar.markdown("**⑤ Flash stock % of main WH stock**")
st.sidebar.caption("Flash Stock to Submit = floor(EAN main WH stock × this %)")
flash_stock_pct = st.sidebar.number_input(
    "Flash stock %", min_value=0.0, value=20.0, step=5.0,
    label_visibility="collapsed"
)

st.sidebar.markdown("**⑥ Min sizes with stock (cut-size rule)**")
st.sidebar.caption(
    "Fail if fewer than this many sizes have main WH stock > 0."
)
min_sizes_with_stock = st.sidebar.number_input(
    "Min sizes with stock", min_value=1, value=3, step=1,
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 🏷️ Remark")

st.sidebar.markdown("**⑦ Required remark value**")
required_remark = st.sidebar.text_input(
    "Required remark", value="Open for all",
    label_visibility="collapsed"
)

st.sidebar.divider()
st.sidebar.caption("PUMA MY / PH / SG flash sale qualifier")


# --------------------------------------------------------------------------
# Main panel
# --------------------------------------------------------------------------
st.title("⚡ Flash Sale Product Analyzer")
st.caption(
    "Configure rules in the sidebar → upload 4 files → click **Analyze** "
    "to get your qualifying EAN list with flash prices and flash stock."
)

# Active config summary banner
with st.expander("📋 Active configuration — click to expand", expanded=True):
    a1, a2, a3, a4 = st.columns(4)
    a1.metric("Marketplace",   marketplace)
    a2.metric("Country",       country)
    a3.metric("Price tier",    price_tier)
    a4.metric("Markdown",      f"{markdown_pct}%")
    b1, b2, b3, b4 = st.columns(4)
    b1.metric("Flash stock",   f"{flash_stock_pct:.0f}% of WH")
    b2.metric("Min sizes",     f"≥ {min_sizes_with_stock} with stock")
    b3.metric("Remark",        required_remark)
    b4.metric("zeCOM sheet",   f"'{country}' tab")

st.divider()

col1, col2 = st.columns(2)
with col1:
    zecom_file     = st.file_uploader(
        f"1️⃣ zeCOM Tracker (.xlsx) — '{country}' sheet will be read",
        type=["xlsx","xls"], key="zecom"
    )
    content_file   = st.file_uploader(
        "2️⃣ Content / Master file (.xlsx)",
        type=["xlsx","xls"], key="content"
    )
with col2:
    inventory_file = st.file_uploader(
        "3️⃣ Warehouse Inventory (.xlsx)",
        type=["xlsx","xls"], key="inv"
    )
    marketplace_file = st.file_uploader(
        f"4️⃣ {marketplace} Price/Stock export (.xlsx)",
        type=["xlsx","xls"], key="mp"
    )

all_uploaded = all([zecom_file, content_file, inventory_file, marketplace_file])
if not all_uploaded:
    st.info("Upload all 4 files to enable analysis.")

run_btn = st.button(
    f"🔍 Analyze — {marketplace} · {country} · {price_tier} base",
    type="primary", disabled=not all_uploaded
)


# --------------------------------------------------------------------------
# Run
# --------------------------------------------------------------------------
if run_btn:
    with st.spinner(f"Reading & joining files — {country} · {marketplace} · {price_tier}…"):
        try:
            zecom_df, sheet_used = read_zecom_tracker(zecom_file, country)
            content_df           = read_content_file(content_file)
            inventory_df         = read_inventory_file(inventory_file)
            marketplace_df       = read_marketplace_file(marketplace_file)
        except Exception as e:
            st.error(f"Error reading files: {e}")
            st.stop()

        if zecom_df.empty:
            st.error(
                f"Could not find a '{country}' sheet in the zeCOM tracker. "
                "Check that the sheet name contains 'MY', 'PH', or 'SG'."
            )
            st.stop()

        st.success(f"✅ Read zeCOM sheet: **{sheet_used}** — {len(zecom_df):,} styles loaded")

        style_df, ean_df = run_analysis(
            zecom_df, content_df, inventory_df, marketplace_df,
            marketplace, price_tier, markdown_pct,
            min_sizes_with_stock, required_remark, flash_stock_pct,
        )

    st.session_state["style_df"] = style_df
    st.session_state["ean_df"]   = ean_df
    st.session_state["cfg"]      = {
        "marketplace": marketplace, "country": country,
        "price_tier": price_tier, "markdown_pct": markdown_pct,
        "flash_stock_pct": flash_stock_pct,
    }


# --------------------------------------------------------------------------
# Display
# --------------------------------------------------------------------------
if "style_df" in st.session_state:
    style_df = st.session_state["style_df"]
    ean_df   = st.session_state["ean_df"]
    cfg      = st.session_state.get("cfg", {})

    total  = len(style_df)
    passed = int(style_df["Qualifies"].sum())
    failed = total - passed
    rate   = round(passed / total * 100, 1) if total else 0
    fsp    = cfg.get("flash_stock_pct", 20)

    st.divider()
    st.markdown(
        f"### Results — **{cfg.get('marketplace','')}** · **{cfg.get('country','')}** · "
        f"**{cfg.get('price_tier','')}** base · **{cfg.get('markdown_pct','')}%** markdown"
    )

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total styles",    f"{total:,}")
    m2.metric("Qualify ✓",       passed)
    m3.metric("Disqualified",    f"{failed:,}")
    m4.metric("Pass rate",       f"{rate}%")
    m5.metric("Qualifying EANs", f"{len(ean_df):,}")

    # Fail reason breakdown chart
    fail_counter = Counter()
    for reasons in style_df.loc[~style_df["Qualifies"], "Fail Reasons"]:
        for r in reasons.split(";"):
            r = r.strip()
            if not r: continue
            if "not listed" in r.lower():
                fail_counter[f"Not on {cfg.get('marketplace','')}"] += 1
            elif "cut sizes" in r.lower():
                fail_counter["Cut sizes (insufficient size run)"] += 1
            elif "remark" in r.lower():
                fail_counter["Remark not matching"] += 1
            elif "missing" in r.lower() or "zero" in r.lower():
                fail_counter["Price missing / zero"] += 1
            elif "platform stock" in r.lower():
                fail_counter["Platform stock < flash stock needed"] += 1
            elif "no main wh" in r.lower():
                fail_counter["No main WH stock"] += 1

    if fail_counter:
        st.subheader("Fail reason breakdown")
        st.bar_chart(
            pd.DataFrame(
                sorted(fail_counter.items(), key=lambda x: -x[1]),
                columns=["Reason", "Count"]
            ).set_index("Reason")
        )

    if passed:
        st.subheader("Qualifying styles by RBU")
        st.bar_chart(
            style_df[style_df["Qualifies"]]["RBU"]
            .value_counts().rename_axis("RBU").reset_index(name="Count")
            .set_index("RBU")
        )

    tab1, tab2 = st.tabs(["📦 Style summary", "🏷️ EAN flash price list"])

    # ── Style summary ──────────────────────────────────────────────────────
    with tab1:
        f1, f2, f3 = st.columns([1, 1, 2])
        with f1:
            status_f = st.selectbox(
                "Status", ["All", "Qualifies only", "Disqualified only"], key="sf1"
            )
        with f2:
            rbu_opts = ["All"] + sorted(style_df["RBU"].dropna().unique().tolist())
            rbu_f    = st.selectbox("RBU", rbu_opts, key="sf2")
        with f3:
            search_f = st.text_input("Search style / description", key="sf3")

        view = style_df.copy()
        if status_f == "Qualifies only":    view = view[view["Qualifies"]]
        if status_f == "Disqualified only": view = view[~view["Qualifies"]]
        if rbu_f != "All":                  view = view[view["RBU"] == rbu_f]
        if search_f:
            m = (view["Style"].str.contains(search_f, case=False, na=False) |
                 view["Description"].str.contains(search_f, case=False, na=False))
            view = view[m]

        st.dataframe(
            view, use_container_width=True, hide_index=True,
            column_config={
                "Qualifies":       st.column_config.CheckboxColumn("Qualifies"),
                "RRP":             st.column_config.NumberColumn("RRP",        format="MYR %.2f"),
                "MD Price":        st.column_config.NumberColumn("MD Price",   format="MYR %.2f"),
                "Base Price":      st.column_config.NumberColumn("Base Price", format="MYR %.2f"),
                "Flash Price":     st.column_config.NumberColumn("Flash Price",format="MYR %.2f"),
                "Disc % vs RRP":   st.column_config.NumberColumn("Disc % vs RRP", format="%.1f%%"),
                "Total Flash Stock to Submit": st.column_config.NumberColumn(
                    "Total Flash Stock to Submit",
                    help=f"Sum of floor({fsp:.0f}% × each EAN's main WH stock)"
                ),
            }
        )
        st.download_button(
            "⬇️ Download style summary CSV",
            data=view.to_csv(index=False).encode("utf-8"),
            file_name=(
                f"flash_{cfg.get('marketplace','').lower()}_"
                f"{cfg.get('country','').lower()}_styles.csv"
            ),
            mime="text/csv", key="dl1"
        )

    # ── EAN detail ─────────────────────────────────────────────────────────
    with tab2:
        st.caption(
            f"**{len(ean_df):,} EANs** from qualifying styles. "
            f"Flash Price = {cfg.get('price_tier','')} × "
            f"(1 − {cfg.get('markdown_pct','')}%).  "
            f"Flash Stock to Submit = floor({fsp:.0f}% × main WH stock per EAN)."
        )

        e1, e2, e3 = st.columns([1, 1, 2])
        with e1:
            rbu_opts2 = ["All"] + (
                sorted(ean_df["RBU"].dropna().unique().tolist())
                if not ean_df.empty else []
            )
            rbu_f2 = st.selectbox("RBU", rbu_opts2, key="ef1")
        with e2:
            gender_opts = ["All"] + (
                sorted(ean_df["Gender"].dropna().unique().tolist())
                if not ean_df.empty else []
            )
            gender_f = st.selectbox("Gender", gender_opts, key="ef2")
        with e3:
            search_f2 = st.text_input("Search style / description / EAN", key="ef3")

        ean_view = ean_df.copy() if not ean_df.empty else ean_df
        if not ean_view.empty:
            if rbu_f2 != "All":   ean_view = ean_view[ean_view["RBU"] == rbu_f2]
            if gender_f != "All": ean_view = ean_view[ean_view["Gender"] == gender_f]
            if search_f2:
                m2 = (
                    ean_view["Style"].str.contains(search_f2, case=False, na=False) |
                    ean_view["Description"].str.contains(search_f2, case=False, na=False) |
                    ean_view["EAN"].str.contains(search_f2, case=False, na=False)
                )
                ean_view = ean_view[m2]

        if ean_view.empty:
            st.info("No qualifying EANs — run the analysis first or adjust your rules.")
        else:
            st.dataframe(
                ean_view, use_container_width=True, hide_index=True,
                column_config={
                    "RRP":                   st.column_config.NumberColumn("RRP",         format="MYR %.2f"),
                    "MD Price":              st.column_config.NumberColumn("MD Price",    format="MYR %.2f"),
                    "SRP":                   st.column_config.NumberColumn("SRP",         format="MYR %.2f"),
                    "Base Price":            st.column_config.NumberColumn("Base Price",  format="MYR %.2f"),
                    "Flash Price":           st.column_config.NumberColumn("Flash Price", format="MYR %.2f"),
                    "Disc % vs RRP":         st.column_config.NumberColumn("Disc % vs RRP", format="%.1f%%"),
                    "Main WH Stock":         st.column_config.NumberColumn(
                        "Main WH Stock", help="Units in main warehouse only"
                    ),
                    "Flash Stock to Submit": st.column_config.NumberColumn(
                        "Flash Stock to Submit",
                        help=f"Floor of {fsp:.0f}% of this EAN's main WH stock"
                    ),
                }
            )
            st.download_button(
                "⬇️ Download EAN flash price list CSV",
                data=ean_view.to_csv(index=False).encode("utf-8"),
                file_name=(
                    f"flash_{cfg.get('marketplace','').lower()}_"
                    f"{cfg.get('country','').lower()}_eans.csv"
                ),
                mime="text/csv", key="dl2"
            )
