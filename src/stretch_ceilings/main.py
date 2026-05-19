import streamlit as st


def get_user_input():
    all_data = {}

    st.subheader("")
    st.subheader("Area Spreadsheets")

    all_data["num_areas"] = st.number_input("Line Items:", min_value=1, value=1, step=1)
    all_data["areas"] = {}

    for item in range(all_data["num_areas"]):
        line_item = st.expander(f"Line Item {item + 1}", expanded=False)
        all_data["areas"][item + 1] = []

        with line_item:
            options = st.number_input(
                "Options:", min_value=1, value=1, step=1, key=f"opt_{line_item}"
            )
            line_items = all_data["areas"][item + 1]

            for opt in range(options):
                area_xl = st.file_uploader(
                    f"Area spreadsheet for Option {opt + 1}",
                    type=["xlsx"],
                    key=f"area_xl_{item}_{opt}",
                )
                line_items.append({"area_table": area_xl})

    submit = st.button("Submit")

    return submit, all_data
