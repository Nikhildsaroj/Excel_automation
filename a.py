import pandas as pd
import streamlit as st
import openpyxl
from io import BytesIO

# =====================
# HELPER FUNCTIONS
# =====================
def calc_shipping(w):
    try:
        return 65 if float(w) <= 1 else float(w) * 65
    except:
        return 0

def process_website_orders(df, cost_df):
    df = df[df["Order From"].str.lower().str.contains("website")].copy()
    if df.empty: 
        return df

    cost_df_unique = cost_df.drop_duplicates(subset="SKU", keep="first")
    df = df.merge(cost_df_unique[["SKU", "Landing Cost GST"]], on="SKU", how="left")

    df["Landing Cost GST"] = df["Landing Cost GST"].fillna("NA")
    df["Landing Cost GST Num"] = pd.to_numeric(df["Landing Cost GST"], errors="coerce").fillna(0)

    df["Shipping"] = df["Weight(KG)"].apply(calc_shipping)
    df["Selling Price with gst"] = df["Dis Price"] * 1.18
    df["Website charge"] = df["Selling Price with gst"] * 0.0185

    # GP formula (with Website charge)
    df["GP"] = df["Selling Price with gst"] - (
        df["Landing Cost GST Num"] + df["Shipping"] + df["Website charge"]
    )
    df["GP %"] = df.apply(lambda r: (r["GP"] / r["Selling Price with gst"] * 100) if r["Selling Price with gst"] else 0, axis=1)

    if "Sr.No" not in df.columns:
        df.insert(0, "Sr.No", range(1, len(df) + 1))

    website_cols = [
        "Sr.No", "Model Name", "SKU", "Product Type", "Brand Company", "QR Code",
        "Weight(KG)", "Order From", "Order Id", "Qty", "Dis Price", "Date",
        "Landing Cost GST", "Shipping", "Website charge",
        "Selling Price with gst", "GP", "GP %"
    ]
    return df.reindex(columns=[c for c in website_cols if c in df.columns])

def process_office_orders(df, cost_df):
    office_sources = [
        "nan", "Tender", "Direct Sales", "Chat Tawk",
        "Reseller", "Instagram", "Facebook", "India Mart", "",
        "Just Dial", "Exhibition"
    ]
    df = df[df["Order From"].isin(office_sources)].copy()
    if df.empty:
        return df

    cost_df_unique = cost_df.drop_duplicates(subset="SKU", keep="first")
    df = df.merge(cost_df_unique[["SKU", "Landing Cost GST"]], on="SKU", how="left")

    df["Landing Cost GST"] = df["Landing Cost GST"].fillna("NA")
    df["Landing Cost GST Num"] = pd.to_numeric(df["Landing Cost GST"], errors="coerce").fillna(0)

    df["Shipping"] = df["Weight(KG)"].apply(calc_shipping)
    df["Selling Price with gst"] = df["Dis Price"] * 1.18

    # GP formula (NO Website charge)
    df["GP"] = df["Selling Price with gst"] - (
        df["Landing Cost GST Num"] + df["Shipping"]
    )
    df["GP %"] = df.apply(lambda r: (r["GP"] / r["Selling Price with gst"] * 100) if r["Selling Price with gst"] else 0, axis=1)

    if "Sr.No" not in df.columns:
        df.insert(0, "Sr.No", range(1, len(df) + 1))

    office_cols = [
        "Sr.No", "Model Name", "SKU", "Product Type", "Brand Company", "QR Code",
        "Weight(KG)", "Order From", "Order Id", "Qty", "Dis Price", "Date",
        "Contact", "Email", "Shipping State", "Sales Person",
        "Landing Cost GST", "Shipping", "Selling Price with gst", "GP", "GP %"
    ]
    return df.reindex(columns=[c for c in office_cols if c in df.columns])

def build_summary_sheet(df, label):
    if df.empty:
        return pd.DataFrame()
    summary = df.groupby("Product Type").agg({
        "Selling Price with gst": "sum",
        "GP": "sum"
    }).reset_index()
    summary["GP%"] = (summary["GP"] / summary["Selling Price with gst"] * 100).round(0).astype(int).astype(str) + "%"
    totals = pd.DataFrame({
        "Product Type": ["Grand Total"],
        "Selling Price with gst": [summary["Selling Price with gst"].sum()],
        "GP": [summary["GP"].sum()],
        "GP%": [str(round(summary["GP"].sum() / summary["Selling Price with gst"].sum() * 100)) + "%"]
    })
    summary = pd.concat([summary, totals], ignore_index=True)
    summary.insert(0, label, "")
    return summary


# =====================
# STREAMLIT APP
# =====================
st.set_page_config(page_title="Sales Analysis Tool", layout="wide")
st.title("📊 Sales Analysis Tool for Orders")

col1, col2 = st.columns(2)
with col1:
    sales_file = st.file_uploader("Upload Sales Excel File", type=["xlsx"], key="sales")
with col2:
    cost_file = st.file_uploader("Upload Cost Excel File", type=["xlsx"], key="cost")

if sales_file and cost_file:
    try:
        df = pd.read_excel(sales_file)
        cost_df = pd.read_excel(cost_file)
        df["Order From"] = df["Order From"].astype(str).str.strip()

        # === FILTER OPTIONS ===
        st.header("2. Order Source Filter")
        filter_option = st.radio(
            "Select which orders to include:",
            ["Website Only", "Office Sales Sources", "Select Specific Sources"]
        )

        if filter_option == "Website Only":
            df = df[df["Order From"].str.lower().str.contains("website")].copy()
        elif filter_option == "Office Sales Sources":
            office_sources = [
                "nan", "Tender", "Direct Sales", "Chat Tawk",
                "Reseller", "Instagram", "Facebook", "India Mart", "",
                "Just Dial", "Exhibition"
            ]
            df = df[df["Order From"].isin(office_sources)].copy()
        elif filter_option == "Select Specific Sources":
            available_sources = sorted(df["Order From"].dropna().unique().tolist())
            selected_sources = st.multiselect("Choose one or more order sources:", options=available_sources)
            if selected_sources:
                df = df[df["Order From"].isin(selected_sources)].copy()
            else:
                df = pd.DataFrame(columns=df.columns)

        if df.empty:
            st.warning("No matching orders found.")
        else:
            website_df = process_website_orders(df, cost_df)
            office_df = process_office_orders(df, cost_df)

            if not website_df.empty:
                st.header("📑 Website Orders")
                st.dataframe(website_df)
                web_excel = BytesIO()
                with pd.ExcelWriter(web_excel, engine="openpyxl") as writer:
                    website_df.to_excel(writer, sheet_name="Website Orders", index=False)
                    build_summary_sheet(website_df, "WEBSITE").to_excel(writer, sheet_name="Website Summary", index=False)
                st.download_button("📥 Download Website Orders Excel", web_excel.getvalue(), file_name="website_orders.xlsx", mime="application/vnd.ms-excel")

            if not office_df.empty:
                st.header("📑 Office Sales Orders")
                st.dataframe(office_df)
                off_excel = BytesIO()
                with pd.ExcelWriter(off_excel, engine="openpyxl") as writer:
                    office_df.to_excel(writer, sheet_name="Office Orders", index=False)
                    build_summary_sheet(office_df, "OFFLINE SALES").to_excel(writer, sheet_name="Office Summary", index=False)
                st.download_button("📥 Download Office Sales Excel", off_excel.getvalue(), file_name="office_orders.xlsx", mime="application/vnd.ms-excel")

            if not website_df.empty or not office_df.empty:
                combined_excel = BytesIO()
                with pd.ExcelWriter(combined_excel, engine="openpyxl") as writer:
                    if not website_df.empty:
                        website_df.to_excel(writer, sheet_name="Website Orders", index=False)
                        build_summary_sheet(website_df, "WEBSITE").to_excel(writer, sheet_name="Website Summary", index=False)
                    if not office_df.empty:
                        office_df.to_excel(writer, sheet_name="Office Orders", index=False)
                        build_summary_sheet(office_df, "OFFLINE SALES").to_excel(writer, sheet_name="Office Summary", index=False)
                st.download_button("📥 Download Combined Excel (Website + Office)", combined_excel.getvalue(), file_name="combined_orders.xlsx", mime="application/vnd.ms-excel")

                st.success("✅ Reports generated successfully.")

    except Exception as e:
        st.error(f"Error: {str(e)}")
else:
    st.info("👆 Please upload both sales and cost files to begin analysis.")
