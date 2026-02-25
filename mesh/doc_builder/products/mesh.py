import pandas as pd
from pathlib import Path
import streamlit as st
from excel_utils import (
    get_total_cols,
    get_max_row,
)

INSTALLATION_RATE = 125


def get_data(xl):

    vars = {}

    ss_type = xl.cell(row=1, column=2).value
    vars["ss_type"] = ss_type.replace(" MESH", "")

    rope_dia = xl.cell(row=1, column=1).value
    vars["rope_dia"] = round(float(rope_dia.replace("Ø", "")), 1)

    mesh_size = xl.cell(row=1, column=6).value
    vars["mesh_size"] = int(mesh_size.replace("MW", ""))

    return vars


def update_data(xl, curr_row):

    vars = get_data(xl)
    vars.update(get_total_cols(xl, curr_row, 6, 11))

    return vars


def generate_df(window_data, installation, rate):

    offer_rate_formula = f"=CEILING({rate}*1.01, 10)"
    area = round(window_data["Area (ft2)"], 0)

    data = [
        [
            "Supply of Mesh at {}mm opening".format(
                window_data["mesh_size"]
            ),
            area,
            "ft\u00b2",
            offer_rate_formula,
        ]
    ]

    if installation:
        data.append(["Installation Charges", area, "ft\u00b2", INSTALLATION_RATE])

    return data


def convert(window_wb, data, installation):

    window_xl = window_wb.worksheets[0]
    max_row = get_max_row(window_xl)
    vars = update_data(window_xl, max_row)
    vars = data | vars

    BASE_DIR = Path(__file__).resolve().parents[2]
    path = BASE_DIR/"files"/"reference_xls"/"price_xls"/"mesh_price_xl.xlsx"
    price_df = pd.read_excel(path)

    filtered_df = price_df[
        (price_df["SS Material"] == vars["ss_type"]) &
        (price_df["Rope Diameter (mm)"] == vars["rope_dia"]) &
        (price_df["Mesh Size"] == vars["mesh_size"])
    ]
    if len(filtered_df) == 0:
        st.error("Pricing for given mesh dimensions unavailable.")
        st.stop()
    else:
        price = filtered_df["Sales Price/ft2.1"].iloc[0]

    data = generate_df(vars, installation, price)

    return data
