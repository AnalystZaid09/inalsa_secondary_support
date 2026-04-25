import streamlit as st
import pandas as pd
import zipfile
import io

st.set_page_config(page_title="Inalsa Secondary Support Report", layout="wide")
st.title("📦 Inalsa Secondary Support")
st.markdown("Upload your files below to generate the Credit Note summary.")

# ─────────────────────────────────────────────────────────────
# File uploaders
# ─────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    b2b_zips = st.file_uploader("B2B Report ZIP(s)", type=["zip"], key="b2b", accept_multiple_files=True)
with col2:
    b2c_zips = st.file_uploader("B2C Report ZIP(s)", type=["zip"], key="b2c", accept_multiple_files=True)

col3, col4 = st.columns(2)
with col3:
    pm_file = st.file_uploader("PM File (Excel)", type=["xlsx", "xls"], key="pm")
with col4:
    unified_csv = st.file_uploader("Unified Transaction CSV", type=["csv"], key="unified")

storage_csv = st.file_uploader("Storage Fee CSV (399153020430.csv style)", type=["csv"], key="storage")

# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────
def read_zip(uploaded_zip):
    """Read the first file inside a ZIP and return as DataFrame."""
    with zipfile.ZipFile(io.BytesIO(uploaded_zip.read()), "r") as z:
        file_name = z.namelist()[0]
        with z.open(file_name) as f:
            if file_name.endswith(".csv"):
                return pd.read_csv(f)
            elif file_name.endswith((".xlsx", ".xls")):
                return pd.read_excel(f)
            else:
                raise ValueError(f"Unsupported file format inside zip: {file_name}")

def read_zips(zip_list):
    """Read multiple ZIPs and concatenate into one DataFrame."""
    frames = [read_zip(z) for z in zip_list]
    return pd.concat(frames, ignore_index=True)

def fmt(val):
    try:
        return f"₹{val:,.2f}"
    except Exception:
        return val

def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────
# Process only when all files are uploaded
# ─────────────────────────────────────────────────────────────
all_uploaded = b2b_zips and b2c_zips and pm_file and unified_csv and storage_csv

if not all_uploaded:
    st.info("⬆️  Please upload all files to proceed.")
    st.stop()

with st.spinner("Processing files…"):

    # ── 1. Load B2B + B2C reports (multiple ZIPs each) ────────
    df_b2b = read_zips(b2b_zips)
    df_b2c = read_zips(b2c_zips)
    final_df = pd.concat([df_b2b, df_b2c], ignore_index=True)

    st.caption(
        f"Loaded {len(b2b_zips)} B2B file(s) and {len(b2c_zips)} B2C file(s) "
        f"→ {len(final_df):,} total rows before filtering."
    )

    # ── 2. Load PM file — Brand (col G) & Purchase Price (col J)
    pm_df = pd.read_excel(pm_file)

    pm_lookup = pm_df.iloc[:, [0, 6]].copy()
    pm_lookup.columns = ["ASIN", "Brand"]

    # Merge Brand
    final_df = final_df.merge(pm_lookup, left_on="Asin", right_on="ASIN", how="left")
    final_df.drop(columns=["ASIN"], inplace=True)

    # ── 3. Filter brands ──────────────────────────────────────
    final_df = final_df[final_df["Brand"].isin(["Inalsa", "Taurus"])]

    # ── 4. Remove Cancel transactions ────────────────────────
    final_df = final_df[final_df["Transaction Type"] != "Cancel"]

    # ── 5. FreeReplacement → Quantity = 0 ────────────────────
    final_df = final_df.copy()
    final_df["Transaction Type"] = final_df["Transaction Type"].astype(str).str.strip()
    final_df.loc[
        final_df["Transaction Type"].str.lower() == "freereplacement", "Quantity"
    ] = 0

    # ── 6. Refund → Quantity negative ────────────────────────
    final_df["Quantity"] = pd.to_numeric(final_df["Quantity"], errors="coerce")
    mask_refund = final_df["Transaction Type"].str.lower() == "refund"
    final_df.loc[mask_refund, "Quantity"] = -final_df.loc[mask_refund, "Quantity"].abs()

    # ── 7. Load Unified Transaction CSV ──────────────────────
    df = pd.read_csv(unified_csv, encoding="utf-8", low_memory=False, header=11)
    df = df[df["type"].isin(["Fulfilment Fee Refund", "Order", "Refund"])]

    numeric_cols = [
        "product sales", "shipping credits", "promotional rebates",
        "selling fees", "fba fees", "other transaction fees",
    ]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    pivot_df = df.groupby("order id")[numeric_cols].sum().reset_index()

    # ── 8. Select required columns from final_df ─────────────
    required_columns = [
        "Seller Gstin", "Invoice Number", "Invoice Date", "Transaction Type",
        "Order Id", "Shipment Id", "Shipment Date", "Order Date",
        "Shipment Item Id", "Quantity", "Item Description", "Asin", "Brand",
        "Invoice Amount", "Tax Exclusive Gross", "Total Tax Amount",
    ]
    available = [c for c in required_columns if c in final_df.columns]
    final_df = final_df[available]

    # Merge pivot
    final_df = final_df.merge(pivot_df, left_on="Order Id", right_on="order id", how="left")
    final_df.drop(columns=["order id"], inplace=True, errors="ignore")

    # ── 9. Capture NaN product sales BEFORE removing them ────
    for col in numeric_cols:
        if col in final_df.columns:
            final_df[col] = pd.to_numeric(final_df[col], errors="coerce")

    final_df["Amazon Fees"] = final_df[[c for c in numeric_cols if c in final_df.columns]].sum(axis=1)

    nan_product_sales_df = final_df[final_df["product sales"].isna()].copy()
    final_df = final_df[final_df["product sales"].notna()]

    # ── 10. With GST Amount Fees ──────────────────────────────
    final_df["With GST Amount Fees"] = (final_df["Amazon Fees"] / 1.18).round(2)

    # ── 11. Base PM & Purchase Cost ───────────────────
    pm_cp_lookup = pm_df.iloc[:, [0, 9]].copy()
    pm_cp_lookup.columns = ["ASIN", "Base PM"]

    final_df = final_df.merge(pm_cp_lookup, left_on="Asin", right_on="ASIN", how="left")
    final_df.drop(columns=["ASIN"], inplace=True, errors="ignore")

    final_df["Base PM"] = pd.to_numeric(final_df["Base PM"], errors="coerce").fillna(0)
    final_df["Quantity"] = pd.to_numeric(final_df["Quantity"], errors="coerce").fillna(0)
    final_df["As Per Qty Base"] = (final_df["Base PM"] * final_df["Quantity"]).round(2)

    # ── 12. Purchase Cost, Gross & Net Margin, Agreed Margin ───────
    final_df["Purchase Cost"] = (final_df["As Per Qty Base"] * 1.18).round(2)
    final_df["Gross Margin"] = (final_df["Tax Exclusive Gross"] - final_df["As Per Qty Base"]).round(2)
    final_df["Net Margin"] = (final_df["Gross Margin"] + final_df["With GST Amount Fees"]).round(2)
    final_df["Agreed Margin"] = (final_df["Tax Exclusive Gross"] * 0.04).round(2)
    final_df["Amount of CN"] = (final_df["Net Margin"] - final_df["Agreed Margin"]).round(2)

    # ── 13. Grand Total row ───────────────────────────────────
    numeric_cols_final = final_df.select_dtypes(include="number").columns
    total_row = {col: "" for col in final_df.columns}
    total_row["Seller Gstin"] = "Grand Total"
    for col in numeric_cols_final:
        total_row[col] = final_df[col].sum()
    final_with_total = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)

    grand_total    = final_with_total.iloc[-1]
    net_sales      = grand_total["Tax Exclusive Gross"]
    minimum_margin = grand_total["Agreed Margin"]
    gross_margin   = grand_total["Gross Margin"]
    amazon_fees    = grand_total["With GST Amount Fees"]

    # ── 14. Storage fees ──────────────────────────────────────
    storage_df = pd.read_csv(storage_csv)

    pm_brand_lookup = pm_df.iloc[:, [0, 6]].copy()
    pm_brand_lookup.columns = ["ASIN", "Brand"]

    storage_df = storage_df.merge(pm_brand_lookup, left_on="asin", right_on="ASIN", how="left")
    storage_df.drop(columns=["ASIN"], inplace=True, errors="ignore")
    storage_df = storage_df[storage_df["Brand"].isin(["Inalsa", "Taurus"])]

    storage_df["estimated-monthly-storage-fee"] = pd.to_numeric(
        storage_df["estimated-monthly-storage-fee"], errors="coerce"
    ).fillna(0)

    total_storage_fee = storage_df["estimated-monthly-storage-fee"].sum()
    total_storage_without_gst = -abs(round(total_storage_fee / 1.18, 2))
    storage_fees = total_storage_without_gst

    # ── 15. CN Summary ────────────────────────────────────────
    total_abc = round(sum([
        float(gross_margin),
        float(amazon_fees),
        float(storage_fees)
    ]), 2)
    credit_note_amount = round(float(minimum_margin) - float(total_abc), 2)

    cn_summary = pd.DataFrame({
        "Particulars": [
            "Net Sales",
            "Minimum Margin (4%)",
            "a. Gross Margin",
            "b. Amazon Fees (Without GST)",
            "c. Storage Fees (Without GST)",
            "Total (a+b+c)",
            "Credit Note Amount",
            "Cost of Operation / Sale",
        ],
        "Amount (₹)": [
            round(float(net_sales), 2),
            round(float(minimum_margin), 2),
            round(float(gross_margin), 2),
            round(float(amazon_fees), 2),
            round(float(storage_fees), 2),
            round(float(total_abc), 2),
            round(float(credit_note_amount), 2),
            round(float(credit_note_amount) + float(gross_margin), 2),
        ],
    })

# ─────────────────────────────────────────────────────────────
# Display Results
# ─────────────────────────────────────────────────────────────

st.success("✅ Processing complete!")

# KPI cards
st.subheader("📊 Summary")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Net Sales", fmt(float(net_sales)))
c2.metric("Minimum Margin (4%)", fmt(float(minimum_margin)))
c3.metric("Total (a+b+c)", fmt(float(total_abc)))
c4.metric("💳 Credit Note Amount", fmt(float(credit_note_amount)))

st.divider()

# CN Summary table
st.subheader("📋 Credit Note Summary")
def highlight_cn(row):
    if row["Particulars"] == "Credit Note Amount":
        return ["background-color: #d4edda; font-weight: bold"] * len(row)
    if row["Particulars"] == "Total (a+b+c)":
        return ["background-color: #fff3cd"] * len(row)
    return [""] * len(row)

st.dataframe(cn_summary.style.apply(highlight_cn, axis=1), use_container_width=True, hide_index=True)

st.divider()

# Detailed transaction data
st.subheader("🗂️ Detailed Transaction Data")
display_cols = [c for c in [
    "Seller Gstin", "Invoice Number", "Invoice Date", "Transaction Type",
    "Order Id", "Asin", "Brand", "Quantity", "Invoice Amount",
    "Tax Exclusive Gross", "Total Tax Amount", "Amazon Fees",
    "With GST Amount Fees", "Base PM", "As Per Qty Base", "Purchase Cost",
    "Gross Margin", "Net Margin", "Agreed Margin", "Amount of CN",
] if c in final_with_total.columns]

with st.expander("Show / Hide Full Table", expanded=False):
    st.dataframe(final_with_total[display_cols], use_container_width=True)

st.divider()

# NaN Product Sales Report
st.subheader("⚠️ NaN Report — Orders Not Found in Unified Transaction CSV")
st.markdown(
    f"**{len(nan_product_sales_df):,} rows** have no matching data in the Unified Transaction file "
    f"(i.e., `product sales` is NaN). These are excluded from the Credit Note calculation."
)

nan_display_cols = [c for c in [
    "Seller Gstin", "Invoice Number", "Invoice Date", "Transaction Type",
    "Order Id", "Asin", "Brand", "Quantity", "Invoice Amount",
    "Tax Exclusive Gross", "Total Tax Amount",
] if c in nan_product_sales_df.columns]

with st.expander(f"Show NaN Report ({len(nan_product_sales_df):,} rows)", expanded=False):
    st.dataframe(nan_product_sales_df[nan_display_cols], use_container_width=True)

st.divider()

# Storage fee summary
st.subheader("🏭 Storage Fee Summary")
st.markdown(f"- **Total Storage Fee (with GST):** {fmt(float(total_storage_fee))}")
st.markdown(f"- **Total Storage Fee (without GST):** {fmt(float(total_storage_without_gst))}")

storage_display_cols = [c for c in [
    "asin", "fnsku", "product-name", "fulfillment-center",
    "estimated-monthly-storage-fee", "Brand",
] if c in storage_df.columns]

with st.expander("Show Storage Data", expanded=False):
    st.dataframe(storage_df[storage_display_cols], use_container_width=True)

# Download buttons
st.divider()
st.subheader("⬇️ Download Results")

col_dl1, col_dl2, col_dl3, col_dl4 = st.columns(4)
with col_dl1:
    st.download_button(
        "📥 Download CN Summary (Excel)",
        data=to_excel_bytes(cn_summary),
        file_name="cn_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col_dl2:
    st.download_button(
        "📥 Download Detailed Report (Excel)",
        data=to_excel_bytes(final_with_total[display_cols]),
        file_name="detailed_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col_dl3:
    st.download_button(
        "📥 Download Storage Report (Excel)",
        data=to_excel_bytes(storage_df[storage_display_cols]),
        file_name="storage_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col_dl4:
    st.download_button(
        "📥 Download NaN Report (Excel)",
        data=to_excel_bytes(nan_product_sales_df[nan_display_cols]),
        file_name="nan_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )