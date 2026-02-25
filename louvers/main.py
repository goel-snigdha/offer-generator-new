import streamlit as st

FINISHES = ["Mill", "Powder Coated - Single Color", "Anodized", "Wood"]


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
                col1, col2 = st.columns([2, 1])
                with col1:
                    area_xl = st.file_uploader(
                        f"Area spreadsheet for Option {opt+1}",
                        type=["xlsx"],
                        key=f"area_xl_{item}_{opt}",
                    )
                with col2:
                    finish = st.selectbox(
                        "Finish", FINISHES, key=f"finish_{item}_{opt}"
                    )
                    finish_clean = "Powder Coated" if finish == "Powder Coated - Single Color" else finish
                line_items.append({"area_table": area_xl, "finish": finish_clean})

    submit = st.button("Submit")

    return submit, all_data
