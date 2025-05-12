import streamlit as st
import pandas as pd
from datetime import datetime
import os
from openpyxl import load_workbook

DATA_FILE = "all_branch_data.xlsx"

# Branch Code → (Branch Name, [Riders])
branch_data = {
    "0001": ("Korangi Branch", ["Ali", "Hamza", "Ikhlaq"]),
    "6661": ("Phase II Branch", ["Sherry", "Hammad"]),
    "7860": ("Lahore Branch", ["Qadeer", "Rashid", "Usman"]),
}

st.title("📦 Commission Slip Submission Form")

branch_code = st.text_input("Enter Your Branch Code")

if branch_code in branch_data:
    branch_name, riders = branch_data[branch_code]
    st.success(f"✅ Welcome, {branch_name}!")

    date = st.date_input("Select Date", datetime.today())
    rider = st.selectbox("Select Employee", riders)
    cash_slips = st.number_input("Cash Slips", min_value=0, step=1)
    online_slips = st.number_input("Online Slips", min_value=0, step=1)

    if st.button("Submit Entry"):
        total = cash_slips + online_slips
        commission = cash_slips * 25 + online_slips * 50

        new_data = {
            "Date": date.strftime("%Y-%m-%d"),
            "Branch": branch_name,
            "Rider": rider,
            "Cash Slips": cash_slips,
            "Online Slips": online_slips,
            "Total Slips": total,
            "Commission": commission
        }

        new_df = pd.DataFrame([new_data])

        if os.path.exists(DATA_FILE):
            with pd.ExcelWriter(DATA_FILE, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                # Load existing workbook
                book = writer.book
                if branch_name in book.sheetnames:
                    existing_df = pd.read_excel(DATA_FILE, sheet_name=branch_name)
                    updated_df = pd.concat([existing_df, new_df], ignore_index=True)
                else:
                    updated_df = new_df

                updated_df.to_excel(writer, sheet_name=branch_name, index=False)
        else:
            with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
                new_df.to_excel(writer, sheet_name=branch_name, index=False)

        st.success("✅ Slip submitted and saved to your branch sheet.")

else:
    if branch_code != "":
        st.error("❌ Invalid Branch Code. Please try again.")
