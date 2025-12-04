import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# ============================================================
# UI HEADER
# ============================================================

st.title("Scotmid Report Mapper")

st.write("""
          1. Export the previous month's data
          2. Drop the file in the below box, it should then give you the output file in your downloads
          3. Standard bits - paste over new data
          4. Copy and paste over values etc!!!
          5. Done.
          """)

# ============================================================
# FILE UPLOADER
# ============================================================

uploaded_file = st.file_uploader("Upload audits_basic_data_export.csv", type=["csv"])

if uploaded_file is not None:

    # ============================================================
    # COLUMN MAP
    # ============================================================

    COLUMN_MAP = {
        "Store Number": "site_code",
        "Order Number": "order_internal_id",
        "Client Name": "client_name",
        "Visit Code": "internal_id",
        "Site Code": "site_internal_id",
        "Order Deadline": "end_date",
        "Responsibility": "responsibility",
        "Premises Name": "site_name",
        "Address1": "site_address_1",
        "Address2": "site_address_2",
        "Address3": "site_address_3",
        "City": None,
        "Post Code": "site_post_code",
        "Start Date": "submitted_date",
        "End Date": "approval_date",
        "Item to Order": "item_to_order",
        "Actual Visit Date": "date_of_visit",
        "Actual Visit Time": "time_of_visit",
        "AMPM": None,
        "Pass-Fail": "primary_result",
        "Pass-Fail2": None,
        "Abort Reason": "site_code",
        "Extra Site 1": "site_code",
        "Extra Site 2": "At which type of till was the purchase made?",
        "Extra Site 3": "At which type of till was the purchase made?",
        "Extra Site 4": "At which type of till was the purchase made?",
        "Extra Site 5": None,
        "Extra Site 6": None,
        "Extra Site 7": None,
        "Extra Site 8": None,
        "Extra Site 9": None,
        "Extra Site 10": "auditor_gender",
        "What type of alcohol did you purchase?": ["What type of alcohol did you purchase?", "What type of E-cigarette product did you purchase/attempt to purchase?"],
        "Please give details of the alcohol purchased (brand and size):": ["Please give details of the alcohol that you purchased:", "Please give details of the cigarettes that you purchased:", "Please give details of the e-cig product that you purchased:"],
        "Did you make the purchase on its own or as part of a larger shop?": "Did you make the purchase on its own or as part of a larger shop?",
        "At the till / bar / counter, did the person ask you your age during the transaction?": "Did the staff member who served you ask your age?",
        "At the till / bar / counter, did the person (or their supervisor) ask you for ID during the transaction?": "Did the staff member who served you ask for ID?",
        "Was a supervisor called at any time during the transaction?": "Was a supervisor called at any point during the transaction?",
        "If a supervisor was called, please give an accurate description of the person (hair style and colour / age / build / height / any distinguishing features):": "Please accurately describe the supervisor:",
        "Was the server working entirely alone (i.e. no-one else working in the store)?": "Was the staff member who served you working entirely alone?",
        "Did the person who served you make eye contact with you during the transaction?": "Did the staff member who served you make eye contact with you during the transaction?",
        "If eye contact was made, when did the person who served you FIRST make eye contact?": "When was eye contact first made?",
        "Did the server look at you long enough to make an assessment of your age?": "Did the staff member who served you look at you long enough to assess your age?  ",
        "How many people were waiting in the queue (if there was no queue, enter 0)?": "How many people were in the queue?",
        "blank1": "At which type of till was the purchase made?",
        "blank2": None,
        "What was the gender of the server?": "What was the gender of the staff member who served you?",
        "What was the approximate age of the server?": "What was the approximate age of the staff member?",
        "Please describe the hair colour, length and style of the server's hair:": "Please accurately describe the staff member who served you:",
        "Was the server wearing a name badge?": "Was the staff member who served you wearing a name badge?",
        "What is the servers name on the name badge (if visible):": "What name was on the name badge?",
        "From the receipt, what is the server's name?": "Please enter the receipt code shown after the date and time on the receipt:",
        "Please enter the 'Operator' code:": None,
        "Please enter the 'Till' code:": None,
        "Please enter the 'Store' code:": None,
        "To help us to identify the site, please describe the surrounding area (i.e. local landmarks, names of stores etc on both sides if possible):": "Please describe the location of the store:",
        "Please enter the 'Transaction' code:": "Please enter the time from the receipt:",
        "Did you see or hear anything you think we or our client should know about?": "Please use this space to explain anything unusual about your visit or to clarify any detail of your report:",
        "blank3": "Please confirm below whether or not you were asked for ID:"
    }

    # ============================================================
    # LOAD DATA
    # ============================================================

    df = pd.read_csv(uploaded_file, dtype=str).fillna("")

    # ============================================================
    # REMOVE ONLY ABORTS
    # ============================================================

    df = df[df["primary_result"].str.strip().str.lower() != "abort"]

    # ============================================================
    # ROW FILTER
    # ============================================================

    df = df[df["item_to_order"].str.strip().isin(["Alcohol", "Cigarettes", "E-Cig"])]

    # ============================================================
    # DATE HELPERS
    # ============================================================

    def get_month_from_date(date_str):
        date_str = (date_str or "").strip()
        if not date_str:
            return None
        try:
            if "/" in date_str:
                d, m, y = date_str.split("/")
                return int(m)
            if "-" in date_str:
                y, m, d = date_str.split("-")
                return int(m)
        except:
            return None
        return None

    def increment_order_id(oid):
        oid = (oid or "").strip()
        m = re.match(r"^(.*?)(\d+)$", oid)
        if not m:
            return oid
        prefix, num = m.groups()
        return f"{prefix}{int(num) + 1}"

    def adjust_order_id(row):
        oid = row.get("order_internal_id", "")
        date_str = row.get("date_of_audit", "") or row.get("date_of_visit", "")
        month = get_month_from_date(date_str)
        if month is not None and month % 2 == 1:
            return increment_order_id(oid)
        return oid

    df["order_internal_id"] = df.apply(adjust_order_id, axis=1)

    # ============================================================
    # MAP FUNCTION
    # ============================================================

    def map_value(row, mapping):
        if mapping is None:
            return ""
        if isinstance(mapping, list):
            parts = [str(row.get(c, "")).strip() for c in mapping if str(row.get(c, "")).strip()]
            return " | ".join(parts)
        return str(row.get(mapping, "")).strip()

    # ============================================================
    # BUILD OUTPUT
    # ============================================================

    ordered_cols = list(COLUMN_MAP.keys())
    out = pd.DataFrame()

    for col in ordered_cols:
        out[col] = df.apply(lambda r: map_value(r, COLUMN_MAP[col]), axis=1)

    final_headers = [("" if col.startswith("blank") else col) for col in ordered_cols]

    # ============================================================
    # OUTPUT CSV TO STREAMLIT
    # ============================================================

    buffer = io.BytesIO()
    out.to_csv(buffer, index=False, header=final_headers, encoding="utf-8-sig")
    buffer.seek(0)

    st.success(f"Processed successfully! {len(out)} rows mapped.")

    st.download_button(
        label="Download Scotmid Report Data.csv",
        data=buffer.getvalue(),
        file_name="Scotmid Report Data.csv",
        mime="text/csv"
    )
