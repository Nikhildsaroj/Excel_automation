import pandas as pd
import streamlit as st
import openpyxl
from io import BytesIO

# Set up the Streamlit app
st.set_page_config(page_title="Sales Analysis Tool", layout="wide")
st.title("📊 Sales Analysis Tool for Orders")
st.markdown("Upload your sales and cost files to analyze order details")

# File upload section
st.header("1. Upload Your Files")
col1, col2 = st.columns(2)

with col1:
    sales_file = st.file_uploader("Upload Sales Excel File", type=["xlsx"], key="sales")
with col2:
    cost_file = st.file_uploader("Upload Cost Excel File", type=["xlsx"], key="cost")

# Instructions
with st.expander("📋 Instructions (Click to Expand)"):
    st.markdown("""
    ### How to use this tool:
    
    1. **Prepare your files**:
       - Sales file should contain these columns: 
         - 'Order From'
         - 'SKU'
         - 'Weight(KG)'
         - 'Dis Price'
       - Cost file should contain:
         - 'SKU'
         - 'Landing Cost GST'
    
    2. **Upload both files**
    
    3. **Choose filter option**:
       - Website only
       - Office sales sources
       - Select specific sources
    
    4. **Review the analysis**
    
    5. **Download the results**
    """)

if sales_file and cost_file:
    try:
        # Read uploaded files
        df = pd.read_excel(sales_file)
        cost_df = pd.read_excel(cost_file)

        # Required column checks
        required_sales_cols = ['Order From', 'SKU', 'Weight(KG)', 'Dis Price']
        required_cost_cols = ['SKU', 'Landing Cost GST']
        missing_sales = [col for col in required_sales_cols if col not in df.columns]
        missing_cost = [col for col in required_cost_cols if col not in cost_df.columns]

        if missing_sales:
            st.error(f"Sales file missing: {', '.join(missing_sales)}")
        elif missing_cost:
            st.error(f"Cost file missing: {', '.join(missing_cost)}")
        else:
            # === FILTER OPTIONS ===
            st.header("2. Order Source Filter")
            filter_option = st.radio(
                "Select which orders to include:",
                ["Website Only", "Office Sales Sources", "Select Specific Sources"]
            )

            # Define office sales sources
            office_sources = [
                "nan", "Tender", "Direct Sales", "Chat Tawk",
                "Reseller", "Instagram", "Facebook", "India Mart", "","Just Dial","Exhibition",
            ]

            # Normalize "Order From" column
            df["Order From"] = df["Order From"].astype(str).str.strip()

            if filter_option == "Website Only":
                filtered_df = df[df["Order From"].str.lower().str.contains("website")].copy()

            elif filter_option == "Office Sales Sources":
                filtered_df = df[df["Order From"].isin(office_sources)].copy()

            elif filter_option == "Select Specific Sources":
                available_sources = sorted(df["Order From"].dropna().unique().tolist())
                selected_sources = st.multiselect(
                    "Choose one or more order sources:",
                    options=available_sources
                )
                if selected_sources:
                    filtered_df = df[df["Order From"].isin(selected_sources)].copy()
                else:
                    filtered_df = pd.DataFrame(columns=df.columns)

            else:
                filtered_df = df.copy()

            if filtered_df.empty:
                st.warning("No matching orders found with the selected filter.")
            else:
                # Deduplicate cost_df
                cost_df_unique = cost_df.drop_duplicates(subset="SKU", keep="first")

                # Merge Landing Cost GST
                filtered_df = filtered_df.merge(
                    cost_df_unique[["SKU", "Landing Cost GST"]],
                    on="SKU", how="left"
                )

                # Replace missing cost
                filtered_df["Landing Cost GST"] = filtered_df["Landing Cost GST"].fillna("NA")
                filtered_df["Landing Cost GST Num"] = pd.to_numeric(
                    filtered_df["Landing Cost GST"], errors="coerce"
                ).fillna(0)

                # SHIPPING calculation
                def calc_shipping(w):
                    try:
                        return 65 if float(w) <= 1 else float(w) * 65
                    except:
                        return 0
                filtered_df["Shipping"] = filtered_df["Weight(KG)"].apply(calc_shipping)

                # SELLING PRICE WITH GST
                filtered_df["Selling Prize with gst"] = filtered_df["Dis Price"] * 1.18

                # WEBSITE CHARGE
                filtered_df["Website charge"] = filtered_df["Selling Prize with gst"] * 0.0185

                # === Sorting Option ===
                st.header("3. Sorting Options")
                sort_col = st.selectbox("Order by column:", filtered_df.columns.tolist())
                sort_order = st.radio("Sort order:", ["Ascending", "Descending"])
                filtered_df = filtered_df.sort_values(
                    by=sort_col, ascending=(sort_order == "Ascending")
                ).reset_index(drop=True)

                # === Results ===
                st.header("4. Analysis Results")
                st.metric("Total Orders", len(filtered_df))

                st.subheader("Detailed Orders")
                st.dataframe(filtered_df)

                # === Export ===
                output_excel, output_csv = BytesIO(), BytesIO()
                export_df = filtered_df.drop(columns=['Landing Cost GST Num'], errors='ignore')

                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    export_df.to_excel(writer, sheet_name='Orders Analysis', index=False)
                    summary_df = pd.DataFrame({
                        'Metric': ['Total Orders', 'Total Cost', 'Total Shipping',
                                   'Total Website Charges'],
                        'Value': [
                            len(filtered_df),
                            filtered_df['Landing Cost GST Num'].sum(),
                            filtered_df['Shipping'].sum(),
                            filtered_df['Website charge'].sum()
                        ]
                    })
                    summary_df.to_excel(writer, sheet_name='Summary', index=False)

                export_df.to_csv(output_csv, index=False)

                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        "📥 Download Excel File", output_excel.getvalue(),
                        file_name="orders_analysis.xlsx", mime="application/vnd.ms-excel"
                    )
                with col2:
                    st.download_button(
                        "📥 Download CSV File", output_csv.getvalue(),
                        file_name="orders_analysis.csv", mime="text/csv"
                    )

                st.success("✅ Analysis completed successfully.")

    except Exception as e:
        st.error(f"Error: {str(e)}")

else:
    st.info("👆 Please upload both sales and cost files to begin analysis.")

# Footer
st.markdown("---")
st.markdown("### Need help?")
st.markdown("Contact support if you encounter issues.")
