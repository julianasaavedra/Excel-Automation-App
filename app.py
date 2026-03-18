import streamlit as st
import pandas as pd
import re
from datetime import datetime, date

st.title("📊 Patient Excel Automation Tool")

st.write("Paste the copied patient data below:")

# Input box
text_input = st.text_area("Input Data")

if st.button("Process Data"):

    data = {}

    # Parse input
    lines = text_input.split("\n")
    for line in lines:
        if ":" in line:
            key, value = line.split(":", 1)
            data[key.strip()] = value.strip()

    # Extract fields
    patient_number = data.get("Patient #")
    patient_name = data.get("Patient Name")
    street_address = data.get("Street Address")

    city_province = data.get("City, Province", "")
    if "," in city_province:
        city, province = [x.strip() for x in city_province.split(",")]
    else:
        city, province = "", ""

    postal_code = data.get("Postal Code")
    phone_number = data.get("Patient Primary Phone #")
    cycle = data.get("Delivery Cycle", "")
    initial_delivery = data.get("Requested Initial Delivery Date")
    clinic_name = data.get("Clinic Name")
    hpr = data.get("HPR")
    type_req = data.get("Type")
    status = data.get("Status")
    delivery_information = data.get("Extra Delivery Information/Comments")

    if not initial_delivery:
        st.error("Missing Initial Delivery Date")
        st.stop()

    # Calendar
    calendar = pd.date_range(start="2026-01-01", end="2026-12-31")
    df_calendar = pd.DataFrame({"date": calendar})

    df_calendar["week_of_year"] = df_calendar["date"].dt.isocalendar().week
    df_calendar["delivery_week"] = ((df_calendar["week_of_year"] - 1) % 4) + 1

    match = df_calendar.loc[
        df_calendar["date"].dt.date == pd.to_datetime(initial_delivery).date(),
        "delivery_week"
    ]

    delivery_week = match.iloc[0] if not match.empty else None

    # Day letter
    date_obj = datetime.strptime(initial_delivery, "%m/%d/%Y")
    days = {0: "M", 1: "T", 2: "W", 3: "H", 4: "F", 5: "S", 6: "S"}
    day_letter = days[date_obj.weekday()]

    # DC Codes
    dc_codes = {
        "AB":149,"BC":150,"MB":148,"NB":145,"NL":144,"NS":145,
        "NT":149,"NU":148,"ON":147,"PE":145,"QC":145,"SK":149,"YT":149
    }

    dc = 145 if clinic_name == "The Ottawa Hospital" else dc_codes.get(province, "Unknown")

    today = date.today()

    cycle_days = int(re.search(r"\d+", cycle).group()) if re.search(r"\d+", cycle) else None

    record = {
        "Date Received from Baxter": today,
        "Requested by": hpr,
        "Patient Name": patient_name,
        "City": city,
        "Province": province,
        "Delivery Week": delivery_week,
        "Delivery Day": day_letter,
        "Cycle Days": cycle_days,
        "Branch Plant": dc,
        "Notes": delivery_information
    }

    df = pd.DataFrame([record])

    st.success("✅ Data processed successfully!")
    st.dataframe(df)

    # Download button
    file_name = "patient_output.xlsx"
    df.to_excel(file_name, index=False)

    with open(file_name, "rb") as f:
        st.download_button(
            label="📥 Download Excel",
            data=f,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )