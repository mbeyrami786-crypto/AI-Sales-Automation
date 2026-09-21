import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment


def analyze_sales(file_path):

    df = pd.read_excel(file_path)

    df["Revenue"] = df["Quantity"] * df["Price"]


    total_sales = df["Revenue"].sum()
    total_quantity = df["Quantity"].sum()


    best_product = (
        df.groupby("Product")["Revenue"]
        .sum()
        .sort_values(ascending=False)
        .head(1)
    )


    best_region = (
        df.groupby("Region")["Revenue"]
        .sum()
        .sort_values(ascending=False)
        .head(1)
    )


    result = {

        "total_sales": round(total_sales, 2),

        "total_quantity": int(total_quantity),

        "best_product": best_product.index[0],

        "best_product_sales": round(
            best_product.values[0],
            2
        ),

        "best_region": best_region.index[0],

        "best_region_sales": round(
            best_region.values[0],
            2
        )

    }


    return result



def create_report(file_path, output_path):

    df = pd.read_excel(file_path)

    df["Revenue"] = df["Quantity"] * df["Price"]


    product_report = (
        df.groupby("Product")["Revenue"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )


    region_report = (
        df.groupby("Region")["Revenue"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )


    summary = pd.DataFrame({

        "Metric": [

            "Total Revenue",
            "Total Quantity",
            "Best Product",
            "Best Region"

        ],

        "Value": [

            df["Revenue"].sum(),
            df["Quantity"].sum(),
            product_report.iloc[0]["Product"],
            region_report.iloc[0]["Region"]

        ]

    })


    with pd.ExcelWriter(
        output_path,
        engine="openpyxl"
    ) as writer:


        summary.to_excel(
            writer,
            sheet_name="Executive Summary",
            index=False
        )


        df.to_excel(
            writer,
            sheet_name="Sales Data",
            index=False
        )


        product_report.to_excel(
            writer,
            sheet_name="By Product",
            index=False
        )


        region_report.to_excel(
            writer,
            sheet_name="By Region",
            index=False
        )



    workbook = load_workbook(output_path)


    for sheet in workbook:

        sheet.freeze_panes = "A2"


        for cell in sheet[1]:

            cell.font = Font(
                bold=True
            )

            cell.alignment = Alignment(
                horizontal="center"
            )


        for column in sheet.columns:

            length = 0

            letter = column[0].column_letter


            for cell in column:

                if cell.value:

                    length = max(
                        length,
                        len(str(cell.value))
                    )


            sheet.column_dimensions[
                letter
            ].width = length + 3


    workbook.save(output_path)


    return output_path




def generate_insights(result):

    insights = []


    insights.append(
        f"🏆 Best Product: {result['best_product']}"
    )


    insights.append(
        f"🌎 Best Region: {result['best_region']}"
    )


    if result["total_sales"] > 10000:

        insights.append(
            "📈 Sales performance is strong."
        )

    else:

        insights.append(
            "📊 Sales performance can be improved."
        )


    return insights