import streamlit as st
import pandas as pd

from automation import (
    analyze_sales,
    create_report,
    generate_insights
)


st.set_page_config(
    page_title="AI Business Analytics",
    page_icon="📊",
    layout="wide"
)



st.title("📊 AI Business Analytics")

st.write(
    "Automated Excel Sales Analysis & Business Insights"
)


st.divider()



uploaded_file = st.file_uploader(
    "📂 Upload Excel Sales File",
    type=["xlsx"]
)



if uploaded_file:


    file_path = "uploaded_sales.xlsx"


    with open(file_path, "wb") as f:

        f.write(
            uploaded_file.getbuffer()
        )



    df = pd.read_excel(file_path)


    df["Revenue"] = (
        df["Quantity"] *
        df["Price"]
    )



    result = analyze_sales(file_path)



    st.success(
        "Analysis completed successfully ✅"
    )



    st.subheader(
        "📌 Business KPIs"
    )



    col1, col2, col3, col4 = st.columns(4)



    with col1:

        st.metric(
            "Total Revenue",
            f"${result['total_sales']}"
        )



    with col2:

        st.metric(
            "Total Quantity",
            result["total_quantity"]
        )



    with col3:

        st.metric(
            "Best Product",
            result["best_product"]
        )



    with col4:

        st.metric(
            "Best Region",
            result["best_region"]
        )



    st.divider()



    st.subheader(
        "🤖 AI Business Insights"
    )



    insights = generate_insights(result)



    for insight in insights:

        st.info(insight)



    st.divider()



    st.subheader(
        "📈 Product Performance"
    )



    product_sales = (
        df.groupby("Product")["Revenue"]
        .sum()
        .sort_values(ascending=False)
    )


    st.bar_chart(
        product_sales
    )



    st.subheader(
        "🌎 Regional Performance"
    )


    region_sales = (
        df.groupby("Region")["Revenue"]
        .sum()
        .sort_values(ascending=False)
    )


    st.bar_chart(
        region_sales
    )



    st.divider()



    with st.expander(
        "View Raw Data"
    ):

        st.dataframe(
            df,
            use_container_width=True
        )



    output = "sales_report.xlsx"



    create_report(
        file_path,
        output
    )



    with open(output, "rb") as file:


        st.download_button(

            label="⬇️ Download Excel Report",

            data=file,

            file_name="sales_report.xlsx",

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )



else:


    st.warning(
        "Please upload an Excel file."
    )
    
    
    
    