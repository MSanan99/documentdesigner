import os
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert

os.makedirs("outputbills", exist_ok=True)

data = pd.read_excel("Electricity_Bill_Dummy_Data_v1.xlsx")

for index, row in data.iterrows():

    doc = DocxTemplate("Electricity_Bill_Template_v1.docx")
    context = {
        "BILL_PERIOD_MONTH": row["BILL_PERIOD_MONTH"],
        "BILL_PERIOD_YEAR": row["BILL_PERIOD_YEAR"],
        "CustomerName": row["CustomerName"],
      "CustomerID": str(row["CustomerID"]).upper(),
        "Address_line1": row["AddressLine1"],
        "Address_line2": row["AddressLine2"],
        "BillDate": row["BillDate"],
        "BILL_PERIOD_STARTDATE": row["BILL_PERIOD_STARTDATE"],
        "BILL_PERIOD_ENDDATE": row["BILL_PERIOD_ENDDATE"],
        "totalMonthlyRecurringCharges": row["totalMonthlyRecurringCharges"],
        "totalUsageCharges": row["totalUsageCharges"],
        "totalOtherCharges": row["totalOtherCharges"],
        "total_taxes": row["total_taxes"],
        "TotalAmount": row["TotalAmount"],
        "DueDate": row["DueDate"],
        "M": row["M"],
        "M1": row["M1"],
        "M2": row["M2"],
        "MTotAmount": row["MTotAmount"],
        "M1TotAmount": row["M1TotAmount"],
        "M2TotAmount": row["M2TotAmount"]
    }

    doc.render(context)

    output_file = os.path.join(
        "outputbills",
        f"Bill_{row['CustomerID']}.docx"
    )

    doc.save(output_file)
    pdf_file = output_file.replace(".docx", ".pdf")
    convert(output_file, pdf_file)
    print(f"Generated DOCX: {output_file}")
    print(f"Generated PDF : {pdf_file}")

print("Done")