import io
import re
import zipfile
import streamlit as st
from louvers.main import get_user_input as louvers_input
from mesh.main import get_user_input as mesh_input
import excel_processor
import doc_processor


def get_product():

    st.sidebar.success("Select a product:")

    page_names_to_funcs = {
        "Aluminium Louvers": louvers_input,
        "SS316 Ropes & Meshes": mesh_input
    }

    demo_name = st.sidebar.selectbox("Choose a demo", page_names_to_funcs.keys())
    submit, area_data = page_names_to_funcs[demo_name]()

    return demo_name, submit, area_data


def get_general_input():
    all_data = {}
    offer_data = {}

    st.subheader("")
    st.subheader("Project Details")

    pattern = r"^VT-\d{4}/[A-G]\d{3}(_R\d{1,2})?$"
    offer_num = st.text_input("Offer Number")
    if not re.match(pattern, offer_num):
        st.error("Offer Number format should match that in Salesforce.")
    offer_data["OfferNumber"] = offer_num
    offer_data["ProjectName"] = st.text_input(
        "Project Name", placeholder="Residential/Commercial/Hotel/..."
    )
    offer_data["ProjectCity"] = st.text_input(
        "Project City", placeholder="Project City"
    )
    offer_data["installation"] = st.checkbox("Installation Included", value=True)

    st.subheader("")
    st.subheader("Recipient Information")

    offer_data["CompanyName"] = st.text_input("Company Name", placeholder="Company Name")
    offer_data["FullName"] = st.text_input("Addressee Full Name", placeholder="Full Name")
    offer_data["Mobile"] = st.text_input("Mobile")
    offer_data["CompanyCity"] = st.text_input(
        "Company City", placeholder="Company City"
    )

    all_data["offer_data"] = offer_data

    return all_data


def handle_conversion(product, data):
    output_xls, updated_data = excel_processor.convert(product, data)
    output_doc = doc_processor.main(product, updated_data)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for file_name, xl in output_xls + [output_doc]:
            zip_file.writestr(file_name, xl.getvalue())

    st.session_state["zip_file"] = (
        zip_buffer,
        f'Files for {updated_data["offer_data"]["OfferNumber"]}.zip',
    )
    st.session_state["conversion_done"] = True


def process(product, submit, data):
    if submit:
        handle_conversion(product, data)
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

    st.write("# Welcome to Vibrant Technik! 👋")

    data = get_general_input()
    product, submit, area_data = get_product()
    data = data | area_data
    process(product, submit, data)


main()
