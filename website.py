import pdfplumber
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
import openpyxl
import streamlit as st
import tempfile
import io


st.title("Changing Pdf to Excel")
pdf_path=st.file_uploader("enter Pdf file",type="pdf",accept_multiple_files=True)


for pdf in pdf_path:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(pdf.read())
        tmp_path = tmp_file.name
    base_name = os.path.splitext(os.path.basename(tmp_path))[0]





table_settings = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "intersection_tolerance": 5,
    "snap_tolerance": 3,
    "join_tolerance": 3,
    "edge_min_length": 3,
}


def extract_morning_depth(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "MORNING DEPTH":
                    return row[index_2 + 6]


def extract_mid_night_depth(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "MID-NIGHT DEPTH":
                    return row[index_2 + 6]


def extract_date_persian(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Date:":
                    return row[index_2 + 2]


def extract_date_eng(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Date:":
                    next_row = table[index_1 + 1]
                    return next_row[37]


def extract_rep(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Rep. #":
                    return row[index_2 + 3]


def extract_bit_size(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "BIT SIZE (in)":
                    return row[index_2 + 6]


def extract_nozzele_size(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "NOZZLE SIZE":
                    return row[index_2 + 6]



def extract_tfa(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "TFA":
                    return row[index_2 + 6]


def extract_bit_type(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "BIT TYPE / IADC":
                    if "/" in row[index_2 + 6]:
                        new = row[index_2 + 6]
                        return new.split("/")[0]


def extract_wob_max(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "WOB (MIN/MAX)":
                    return row[index_2 + 8]


def extract_wob_min(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "WOB (MIN/MAX)":
                    return row[index_2 + 6]


def extract_rpm(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "RPM / TQ(kft.lb)":
                    if "/" in row[index_2 + 6]:
                        new = row[index_2 + 6]
                        return new.split("/")[0]
                    else:
                        return row[index_2 + 6]


def extract_tq(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "RPM / TQ(kft.lb)":
                    if "/" in row[index_2 + 6]:
                        new = row[index_2 + 6]
                        return new.split("/")[1]


def extract_rop(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "ROP (m/hr)":
                    return row[index_2 + 6]


def extract_gpm1(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "GPM":
                    return row[index_2 + 6]



def extract_gpm2(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "GPM":
                    return row[index_2 + 8]



def extract_PUMP_Pr(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "PUMP Pr. (psi)":
                    return row[index_2 + 6]


def extract_ecd1(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "ECD (PCF)":
                    return row[index_2 + 6]


def extract_ecd2(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "ECD (PCF)":
                    return row[index_2 + 8]



def extract_W_R(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "W&R (M,H)":
                    return row[index_2 + 6]


def extract_mud_type(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Mud Type":
                    return row[index_2 + 2]


def extract_pv(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "PV (cp)":
                    return row[index_2 + 2]


def extract_yp(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "YP(lb/100ft²)":
                    return row[index_2 + 2]


def extract_R600(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "R600/R300":
                    if "/" in row[index_2 + 2]:
                        text = row[index_2 + 2]
                        return text.split("/")[0]


def extract_R300(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "R600/R300":
                    if "/" in row[index_2 + 2]:
                        text = row[index_2 + 2]
                        return text.split("/")[1]


def extract_mf_vise(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "MF Vis":

                    s = row[index_2+2]
                    num = int(''.join(filter(str.isdigit, s)))
                    return num



def extract_mw_min(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "MW (pcf)":
                    return row[index_2 + 3]



def extract_mw_max(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "MW (pcf)":
                    return row[index_2 + 6]


def extract_F_L(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "F.L(cc)":
                    return row[index_2 + 3]


def extract_cake(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Cake(1/32)":
                    return row[index_2 + 3]


def extract_gel_s(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "10s/10m Gel":
                    if "/" in row[index_2 + 2]:
                        new = row[index_2 + 2]
                        return new.split("/")[0]

def extract_gel_m(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "10s/10m Gel":
                    if "/" in row[index_2 + 2]:
                        new = row[index_2 + 2]
                        return new.split("/")[1]



def extract_ph_alka(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "PH / Alka":
                    return row[index_2 + 2]


def extract_hpht(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "HPHT F.L":
                    return row[index_2 + 3]


def extract_water(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Water":
                    return row[index_2 + 2]


def extract_solid(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Solid (%)":
                    return row[index_2 + 3]


def extract_e_s(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "E.S":
                    return row[index_2 + 1]


def extract_kci(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "KCl (wt%)":
                    return row[index_2 + 3]


def extract_oil(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Oil":
                    return row[index_2 + 1]


def extract_diesel(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Diesel":
                    return row[index_2 + 2]


def extract_bf(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "B.F. (%)":
                    return row[index_2 + 2]


def extract_sand(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Sand (%)":
                    return row[index_2 + 3]


def extract_pf(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Pf / Mf":
                    if "/" in row[index_2 + 3]:
                        new = row[index_2 + 3]
                        return new.split("/")[0]


def extract_mf(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Pf / Mf":
                    if "/" in row[index_2 + 3]:
                        new = row[index_2 + 3]
                        return new.split("/")[1]


def extract_ca(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Ca":
                    return row[index_2 + 1]


def extract_mbt(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "MBT(ppb)":
                    return row[index_2 + 3]


def extract_cl(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Cl":
                    return row[index_2 + 1]


def extract_Desander_Desilter(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Desander/Desilter":
                    return row[index_2 + 3]


def extract_loss(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Loss":
                    return row[index_2 + 2]


def extract_gain(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Gain":
                    return (row[index_2 + 2])


def extract_form(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Form." and index_2==0:
                    one=table[index_1+2]
                    main_one=one[0]
                    two=table[index_1+3]
                    main_two = two[0]
                    three=table[index_1+4]
                    main_three = three[0]
                    four=table[index_1+5]
                    main_four = four[0]
                    five=table[index_1+6]
                    main_five = five[0]
                    six=table[index_1+7]
                    main_six = six[0]
                    seven=table[index_1+8]
                    main_seven = seven[0]
                    return main_one,main_two,main_three,main_four,main_five,main_six,main_seven

def extract_top_act(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Act.":
                    one = table[index_1 + 1]
                    main_one = one[5]
                    two = table[index_1 + 2]
                    main_two = two[5]
                    three = table[index_1 + 3]
                    main_three = three[5]
                    four = table[index_1 + 4]
                    main_four = four[5]
                    five = table[index_1 + 5]
                    main_five = five[5]
                    six = table[index_1 + 6]
                    main_six = six[5]
                    seven = table[index_1 + 7]
                    main_seven = seven[5]
                    return main_one, main_two, main_three, main_four, main_five, main_six, main_seven


def extract_lithology(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "Lithology":
                    one = table[index_1 + 2]
                    main_one = one[7]
                    two = table[index_1 + 3]
                    main_two = two[7]
                    three = table[index_1 + 4]
                    main_three = three[7]
                    four = table[index_1 + 5]
                    main_four = four[7]
                    five = table[index_1 + 6]
                    main_five = five[7]
                    six = table[index_1 + 7]
                    main_six = six[7]
                    seven = table[index_1 + 8]
                    main_seven = seven[7]
                    return main_one, main_two, main_three, main_four, main_five, main_six, main_seven


def extract_summary(tables):
    for index_0, table in enumerate(tables):
        for index_1, row in enumerate(table):
            for index_2, item in enumerate(row):
                if item == "SUMMARY":
                    return row[index_2 + 2]


my_list=[]
for pdf in pdf_path:
    with pdfplumber.open(pdf) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            print(f"\n Page {page_num}")
            tables = page.extract_tables(table_settings=table_settings)

            if not tables:
                print(" No tables found on this page.")
                continue


            data = [[extract_rep(tables),
                 extract_date_persian(tables),
                 extract_date_eng(tables),
                 extract_morning_depth(tables),
                 extract_mid_night_depth(tables),
                 extract_bit_size(tables),
                 extract_nozzele_size(tables),
                 extract_tfa(tables),
                 extract_tq(tables),
                 extract_bit_type(tables),
                 extract_wob_min(tables),
                 extract_wob_max(tables),
                 extract_rpm(tables),
                 extract_rop(tables),
                 extract_gpm1(tables),
                 extract_gpm2(tables),
                 extract_PUMP_Pr(tables),
                 extract_ecd1(tables),
                 extract_ecd2(tables),
                 extract_W_R(tables),
                 extract_mud_type(tables),
                 extract_pv(tables),
                 extract_yp(tables),
                 extract_R600(tables),
                 extract_R300(tables),
                 extract_mf_vise(tables),
                 extract_mw_min(tables),
                 extract_mw_max(tables),
                 extract_F_L(tables),
                 extract_cake(tables),
                 extract_gel_s(tables),
                 extract_gel_m(tables),
                 extract_ph_alka(tables),
                 extract_hpht(tables),
                 extract_water(tables),
                 extract_solid(tables),
                 extract_e_s(tables),
                 extract_kci(tables),
                 extract_oil(tables),
                 extract_diesel(tables),
                 extract_bf(tables),
                 extract_sand(tables, ),
                 extract_pf(tables),
                 extract_mf(tables),
                 extract_ca(tables),
                 extract_mbt(tables),
                 extract_cl(tables),
                 extract_Desander_Desilter(tables),
                 extract_loss(tables),
                 extract_gain(tables),
                 extract_form(tables),
                 extract_top_act(tables),
                 extract_lithology(tables),
                 extract_summary(tables)
                ]]

            columns = [
                "Rep. #",
                "Date fa",
                "Date eng",
                "MORNING DEPTH (m)",
                "MID NIGHT DEPTH (m)",
                "Bit size",
                "Nozzele size",
                "TFA",
                "TQ (klbs)",
                "Bit Type",
                "WOB_min(klbs)",
                "WOB_max(klbs)",
                "RPM",
                "ROP (m/hr)",
                "GPM-1",
                "GPM-2",
                "PUMP Pr (psi)",
                "ECD-1",
                "ECD-2",
                "W&R (M&H)",
                "Mud Type",
                "PV",
                "YP",
                "R600",
                "R300",
                "MF VIS",
                "MW min",
                "MW max",
                "API FL (cc/30min)",
                "cake",
                "Gel 10s",
                "Gel 10m",
                "PH",
                "HPHT FL (cc/30min)",
                "Water(%)",
                "solid(%)",
                "E.S",
                "KCI",
                "oil",
                "Diesel",
                "B.F.(%)",
                "sand(%)",
                "Pf",
                "Mf",
                "Ca",
                "MBT",
                "CL",
                "Trip Loss",
                "Loss",
                "Gain",
                "Form.",
                "Top Act(MD/TVD) ",
                "Lithology",
                "SUMMARY"
            ]



            index_value = -1
            existing_df = None
            if os.path.exists("./Report.xlsx"):
                existing_df = pd.read_excel("./Report.xlsx")
                index_value = existing_df.index.max()
                new_index = pd.RangeIndex(start=0, stop=index_value + 1, name="idx")
                existing_df = existing_df.set_index(new_index)
            df = pd.DataFrame(data, columns=columns,
                              index=pd.RangeIndex(start=index_value + 1, stop=index_value + 2, name="idx"))
            # df.to_excel("./Re.xlsx")
            if existing_df is None:
                combined = df
                my_list.append(combined)
            else:

                combined = pd.concat([existing_df, df],ignore_index=True)
                my_list.append(combined)
            if my_list:
                combined = pd.concat(my_list, ignore_index=True)
            else:
                combined = pd.DataFrame()
                    #combined = pd.concat(my_list, ignore_index=True)

            #combined.to_excel("./Report.xlsx",index=False)
output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    combined.to_excel(writer, index=False, sheet_name="Sheet1")
excel_data = output.getvalue()
st.download_button(
label=" دانلود Excel",
data=excel_data,
file_name="output.xlsx",
mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

