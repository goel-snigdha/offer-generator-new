import pandas as pd
import streamlit as st


def main():

    st.subheader("")
    st.subheader("Commercials Table")
    uploaded_file = st.file_uploader("Upload your commercials file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        st.success(f"Uploaded: {uploaded_file.name}")

    st.subheader("Price Distribution")
    df = pd.DataFrame(
        {"": ["Material", "Installation"], "Percentage": [None, None], "Absolute Value": [None, None]},
    )
    st.data_editor(df, disabled=[""], hide_index=True)

    st.subheader("Payment Schedule")
    payment_df = pd.DataFrame({
        "Percentage": ["60%", "35% + Full GST", "5%", ""],
        "Description": [
            "As Advance",
            "Against Proforma Invoice before delivery of material to the site",
            "After installation completion",
            "",
        ],
        "Amount": ["Rs. 6,00,000", "Rs. 4,68,000", "Rs. 50,000", ""],
    })
    st.data_editor(payment_df, hide_index=True)
