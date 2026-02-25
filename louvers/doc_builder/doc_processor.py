from pathlib import Path
from mailmerge import MailMerge
from doc_utils import convert_to_doc


def merge_product_section_vars(area_data):
    BASE_DIR = Path(__file__).resolve().parent.parent
    path = BASE_DIR/"files"/"offer_templates"/"products"
    section_template = "{}/{}.docx".format(
        path, area_data["product"].lower().replace(" ", "-")
    )
    section_doc = MailMerge(section_template)

    section_doc.merge(
        Finish=area_data["finish"],
        Orientation=area_data["orientation"]
    )

    if area_data['product'] in ['Grille', 'Aerofoil', 'Rectangular']:

        section_doc.merge(
            Pitch=str(area_data["pitch"]) + " mm",
            SectionType=area_data["section_type"]
        )

        if area_data['product'] == 'Aerofoil':
            section_doc.merge(
                Fixing=area_data['installation_method'],
            )

    if area_data['product'] in ['S-Louvers', 'Fluted']:
        section_doc.merge(
            Coverage=area_data['finish_sides']
        )

    if area_data['product'] in ['Cottal']:
        section_doc.merge(
            SectionType=area_data["section_type"]
        )

    section_doc_obj = convert_to_doc(section_doc)

    return section_doc_obj


def create_commercials_section(area_data):

    BASE_DIR = Path(__file__).resolve().parents[2]
    section_template = BASE_DIR/"files"/"offer_templates"/"areas_commercials.docx"

    section_doc = MailMerge(section_template)

    section_doc.merge(
        LineItem=area_data["line_item_str"],
        Option=area_data["option_str"],
        OfferNumber=area_data["OfferNumber"]
    )

    section_doc_obj = convert_to_doc(section_doc)

    return section_doc_obj


def create_product_section(data):

    areas = data["areas"]

    section_files = []
    master_product_title = ""

    for line_item in areas:

        for opt_idx in range(len(areas[line_item])):

            option_data = areas[line_item][opt_idx] | {'OfferNumber': data['offer_data']['OfferNumber']}
            section_files.append(merge_product_section_vars(option_data))
            section_files.append(create_commercials_section(option_data))

            if not master_product_title:
                master_product_title = option_data['product']
            if option_data['product'] != master_product_title:
                master_product_title = "Louvers"

    return master_product_title, section_files
