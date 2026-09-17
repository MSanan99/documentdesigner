import os
from datetime import date, datetime, time

import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert


os.makedirs("outputbills", exist_ok=True)

data = pd.read_excel("Telecom_Bill_Dummy_Data_v1.xlsx")
call_df = pd.read_excel("Telecom_Bill_Dummy_Data_v1.xlsx", "CallData")
usage_chrgs = pd.read_excel("Telecom_Bill_Dummy_Data_v1.xlsx", "UsageChrg")
mrc_chrgs = pd.read_excel("Telecom_Bill_Dummy_Data_v1.xlsx", "MrcChrgs")
oth_chrgs = pd.read_excel("Telecom_Bill_Dummy_Data_v1.xlsx", "OthChrgs")


def format_date(val):
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, pd.Timestamp):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, date):
        return val.strftime("%d/%m/%Y")

    text = str(val).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return text


def format_time(val):
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, pd.Timestamp):
        return val.strftime("%H:%M:%S")
    if isinstance(val, datetime):
        return val.strftime("%H:%M:%S")
    if isinstance(val, time):
        return val.strftime("%H:%M:%S")
    return str(val)


def format_inr(val):
    if val is None or pd.isna(val):
        return "\u20b90.00"
    return f"\u20b9{float(val):.2f}"


generated_files = []

for index, row in data.iterrows():

    doc = DocxTemplate("Telecom_Bill_Template2_v1.docx")

    customer_id = str(row["CustomerID"]).upper()

    customer_calls = call_df[
        call_df["CustomerID"].str.upper() == customer_id
    ]

    call_details = []
    for rowno, call in enumerate(customer_calls.itertuples(), start=1):
        call_details.append({
            "RowNo": rowno,
            "CallDate": format_date(call.CallDate),
            "CallTime": format_time(call.CallTime),
            "Direction": call.Direction,
            "PhoneNumber": call.PhoneNumber,
            "CallCategory": call.CallCategory,
            "Duration": call.Duration,
            "CallCharge": f"{float(call.CallCharge):.2f}"
        })

    cust_usage_charges = usage_chrgs[
        usage_chrgs["CustomerID"].str.upper() == customer_id
    ]
    usage_charges = []
    for rowno, uc in enumerate(cust_usage_charges.itertuples(), start=1):
        usage_charges.append({
            "RowNo": rowno,
            "servtype": uc.servtype,
            "category": uc.category,
            "totalunit": uc.totalunit,
            "allowance": uc.allowance,
            "overage": uc.overage,
            "rate": f"{float(uc.rate):.2f}",
            "charge": f"{float(uc.charge):.2f}"
        })

    cust_mrc_charges = mrc_chrgs[
        mrc_chrgs["CustomerID"].str.upper() == customer_id
    ]
    mrc_charges = []
    for rowno, mrc in enumerate(cust_mrc_charges.itertuples(), start=1):
        mrc_charges.append({
            "RowNo": rowno,
            "planname": mrc.planname,
            "startdt": format_date(mrc.startdt),
            "enddt": format_date(mrc.enddt),
            "chrgamt": f"{float(mrc.chrgamt):.2f}"
        })

    cust_other_charges = oth_chrgs[
        oth_chrgs["CustomerID"].str.upper() == customer_id
    ]
    other_charges = []
    for rowno, other in enumerate(cust_other_charges.itertuples(), start=1):
        other_charges.append({
            "RowNo": rowno,
            "chrgdesc": other.chrgdesc,
            "chrgdt": format_date(other.chrgdt),
            "chrgamt": f"{float(other.chrgamt):.2f}"
        })

    context = {
        "BILL_PERIOD_MONTH": row["BILL_PERIOD_MONTH"],
        "BILL_PERIOD_YEAR": row["BILL_PERIOD_YEAR"],
        "CustomerName": row["CustomerName"],
        "CustomerID": customer_id,
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
        "Customer_specific_messages": row["Customer_specific_messages"],
        "DueDate": row["DueDate"],
        "CallDetails": call_details,
        "MrcCharges": mrc_charges,
        "OtherCharges": other_charges,
        "UsageCharges": usage_charges,
        "prevbill": format_inr(row["prevbill"]),
        "paymrcvd": format_inr(row["paymrcvd"]),
        "curchrgs": format_inr(row["curchrgs"]),
        "curtaxes": format_inr(row["curtaxes"]),
        "curbill": format_inr(row["curbill"]),
        "totdue": format_inr(row["totdue"]),
        "contactno": row["contactno"],
        "custemail": row["custemail"],
        "custmsdn": row["custmsdn"]


    }

    doc.render(context)

    output_file = os.path.join(
        "outputbills",
        f"Bill_{customer_id}.docx"
    )

    doc.save(output_file)
    generated_files.append(output_file)

    print(f"Generated DOCX: {output_file}")

for output_file in generated_files:
    pdf_file = os.path.splitext(output_file)[0] + ".pdf"
    convert(output_file, pdf_file)
    print(f"Generated PDF : {pdf_file}")
