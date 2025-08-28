import pandas as pd
import streamlit as st
import openpyxl
from io import BytesIO

# Set up the Streamlit app
st.set_page_config(page_title="Sales Analysis Tool", layout="wide")
st.title("📊 Sales Analysis Tool for Website Orders")
st.markdown("Upload your sales and cost files to analyze website order profitability")

# File upload section
st.header("1. Upload Your Files")
col1, col2 = st.columns(2)

with col1:
    sales_file = st.file_uploader("Upload Sales Excel File", type=["xlsx"], key="sales")
with col2:
    cost_file = st.file_uploader("Upload Cost Excel File", type=["xlsx"], key="cost")

# Instructions section
with st.expander("📋 Instructions (Click to Expand)"):
    st.markdown("""
    ### How to use this tool:
    
    1. **Prepare your files**:
       - Sales file should contain these columns: 
         - 'Order From' (to identify website orders)
         - 'SKU' (product code)
         - 'Weight(KG)' (product weight)
         - 'Dis Price' (discounted price)
       - Cost file should contain:
         - 'SKU' (matching product codes)
         - 'Landing Cost GST' (product cost including GST)
    
    2. **Upload both files** using the upload areas above
    
    3. **Review the analysis** that appears automatically after upload
    
    4. **Download the results** using the download button at the bottom
    
    ### What the analysis includes:
    - Filtering of website orders only
    - Calculation of shipping costs (₹65 for first kg, ₹65/kg additional)
    - Calculation of selling price with GST (1.8 × discounted price)
    - Website charges (1.85% of selling price with GST)
    - Profitability metrics
    """)

if sales_file and cost_file:
    try:
        # Read the uploaded files
        df = pd.read_excel(sales_file)
        cost_df = pd.read_excel(cost_file)
        
        # Check if required columns exist
        required_sales_cols = ['Order From', 'SKU', 'Weight(KG)', 'Dis Price']
        required_cost_cols = ['SKU', 'Landing Cost GST']
        
        missing_sales = [col for col in required_sales_cols if col not in df.columns]
        missing_cost = [col for col in required_cost_cols if col not in cost_df.columns]
        
        if missing_sales:
            st.error(f"Sales file is missing these required columns: {', '.join(missing_sales)}")
        elif missing_cost:
            st.error(f"Cost file is missing these required columns: {', '.join(missing_cost)}")
        else:
            # Process the data
            with st.spinner("Analyzing your data..."):
                # Filter only Website orders
                website_df = df[df["Order From"].str.lower().str.contains("website")].copy()
                
                if website_df.empty:
                    st.warning("No website orders found in the sales data. Please check if 'Order From' column contains 'website' values.")
                else:
                    # Deduplicate cost_df to avoid multiple matches per SKU
                    cost_df_unique = cost_df.drop_duplicates(subset="SKU", keep="first")
                    
                    # Merge Landing Cost GST from cost.xlsx based on SKU
                    website_df = website_df.merge(
                        cost_df_unique[["SKU", "Landing Cost GST"]],
                        on="SKU",
                        how="left"
                    )
                    
                    # Check for missing cost data
                    missing_cost_skus = website_df[website_df["Landing Cost GST"].isna()]["SKU"].unique()
                    if len(missing_cost_skus) > 0:
                        st.warning(f"Cost data not found for these SKUs: {', '.join(map(str, missing_cost_skus[:5]))}{'...' if len(missing_cost_skus) > 5 else ''}")
                    
                    # SHIPPING calculation
                    def calc_shipping(w):
                        try:
                            return 65 if float(w) <= 1 else float(w) * 65
                        except:
                            return 0
                    
                    website_df["Shipping"] = website_df["Weight(KG)"].apply(calc_shipping)
                    
                    # SELLING PRICE WITH GST
                    website_df["Selling Prize with gst"] = website_df["Dis Price"] * 1.8
                    
                    # WEBSITE CHARGE
                    website_df["Website charge"] = website_df["Selling Prize with gst"] * 0.0185
                    
                    # Calculate profit
                    website_df["Profit"] = website_df["Dis Price"] - website_df["Landing Cost GST"] - website_df["Shipping"] - website_df["Website charge"]
                    
                    # Display results
                    st.header("2. Analysis Results")
                    
                    # Summary metrics
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Website Orders", len(website_df))
                    with col2:
                        st.metric("Total Revenue", f"₹{website_df['Dis Price'].sum():,.2f}")
                    with col3:
                        st.metric("Total Profit", f"₹{website_df['Profit'].sum():,.2f}")
                    with col4:
                        avg_profit = website_df['Profit'].mean() if len(website_df) > 0 else 0
                        st.metric("Average Profit per Order", f"₹{avg_profit:,.2f}")
                    
                    # Show data table
                    st.subheader("Detailed Order Analysis")
                    st.dataframe(website_df)
                    
                    # Profit distribution chart
                    st.subheader("Profit Distribution")
                    st.bar_chart(website_df["Profit"])
                    
                    # Prepare files for download - WITHOUT Profit column
                    output_excel = BytesIO()
                    output_csv = BytesIO()
                    
                    # Create DataFrame without Profit column for export
                    export_df = website_df.drop(columns=['Profit'], errors='ignore')
                    
                    # Excel file
                    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                        export_df.to_excel(writer, sheet_name='Website Sales Analysis', index=False)
                        
                        # Create a summary sheet
                        summary_data = {
                            'Metric': ['Total Orders', 'Total Revenue', 'Total Cost', 'Total Shipping', 
                                      'Total Website Charges', 'Total Profit', 'Average Profit per Order'],
                            'Value': [
                                len(website_df),
                                website_df['Dis Price'].sum(),
                                website_df['Landing Cost GST'].sum(),
                                website_df['Shipping'].sum(),
                                website_df['Website charge'].sum(),
                                website_df['Profit'].sum(),
                                website_df['Profit'].mean()
                            ]
                        }
                        summary_df = pd.DataFrame(summary_data)
                        summary_df.to_excel(writer, sheet_name='Summary', index=False)
                    
                    # CSV file
                    export_df.to_csv(output_csv, index=False)
                    
                    # Download buttons
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.download_button(
                            label="📥 Download Excel File",
                            data=output_excel.getvalue(),
                            file_name="website_sales_analysis.xlsx",
                            mime="application/vnd.ms-excel"
                        )
                    
                    with col2:
                        st.download_button(
                            label="📥 Download CSV File",
                            data=output_csv.getvalue(),
                            file_name="website_sales_analysis.csv",
                            mime="text/csv"
                        )
                    
                    st.success("Analysis completed successfully! Download your results using the buttons above.")
    
    except Exception as e:
        st.error(f"An error occurred: {str(e)}. Please check your files and try again.")

else:
    st.info("👆 Please upload both sales and cost files to begin analysis.")

# Footer
st.markdown("---")
st.markdown("### Need help?")
st.markdown("Check the instructions above or contact support if you encounter any issues.")