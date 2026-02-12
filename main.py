import re
import io
import zipfile
import streamlit as st
import doc_builder.excel_processor as excel_processor
from doc_builder.doc_processor import main as generate_offer_doc

FINISHES = ["Mill", "Powder Coated - Single Color", "Anodized", "Wood"]


def get_user_input():
    all_data = {}
    offer_data = {}

    st.subheader("Project Details")

    pattern = r"^VT-\d{4}/[A-G]\d{3}(_R\d{1,2})?$"
    offer_num = st.text_input("Offer Number", value="VT-2526/B014")
    if not re.match(pattern, offer_num):
        st.error("Offer Number format should match that in Salesforce.")
    offer_data["OfferNumber"] = offer_num
    offer_data["ProjectName"] = st.text_input(
        "Project Name", placeholder="Residential/Commercial/Hotel/...", value="Hotel Ritz Carlton"
    )
    offer_data["ProjectCity"] = st.text_input(
        "Project City", placeholder="Project City", value="Navi Mumbai"
    )
    offer_data["installation"] = st.checkbox("Installation Included", value=True)

    st.subheader("")
    st.subheader("Recipient Information")

    offer_data["CompanyName"] = st.text_input("Company Name", placeholder="Company Name", value="ABC Hospitality Ltd")
    offer_data["FullName"] = st.text_input("Addressee Full Name", placeholder="Full Name", value="Sanjay Goel")
    offer_data["Mobile"] = st.text_input("Mobile", value="+918875621020")
    offer_data["CompanyCity"] = st.text_input(
        "Company City", placeholder="Company City", value="New Delhi"
    )

    all_data["offer_data"] = offer_data

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


def handle_conversion(data):
    output_xls, updated_data = excel_processor.convert(data)
    output_doc = generate_offer_doc(updated_data)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for file_name, xl in output_xls + [output_doc]:
            zip_file.writestr(file_name, xl.getvalue())

    st.session_state["zip_file"] = (
        zip_buffer,
        f'Files for {updated_data["offer_data"]["OfferNumber"]}.zip',
    )
    st.session_state["conversion_done"] = True


def process(submit, data):
    if submit:
        handle_conversion(data)
        st.rerun()

    if "conversion_done" in st.session_state:
        btn = st.download_button(
            "Download ZIP",
            file_name=st.session_state["zip_file"][1],
            mime="application/zip",
            data=st.session_state["zip_file"][0],
        )
        if btn:
            st.write("ZIP Downloaded Successfully")


def main():
    st.title("Vibrant Technik Offer Generator")
    submit, data = get_user_input()
    process(submit, data)


main()
