from io import BytesIO
from mailmerge import MailMerge
from docxcompose.composer import Composer
from doc_builder.doc_utils import get_merge_fields, convert_to_doc


def create_product_section(area_data):
    path = "files/offer_templates/products"
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
            # Pitch=str(area_data["pitch"]) + " mm",
            SectionType=area_data["section_type"]
        )

    section_doc_obj = convert_to_doc(section_doc)

    return section_doc_obj


def create_commercials_section(area_data):
    section_template = "files/offer_templates/areas_commercials.docx"
    section_doc = MailMerge(section_template)

    section_doc.merge(
        LineItem=area_data["line_item_str"],
        Option=area_data["option_str"]
    )

    section_doc_obj = convert_to_doc(section_doc)

    return section_doc_obj


def combine_documents(header, section_files, footer):

    composer = Composer(header)
    for section in section_files + [footer]:
        header.add_page_break()
        composer.append(section)
    combined_doc_obj = convert_to_doc(header)

    return combined_doc_obj


def merge_data(template_path, merge_fields):

    document = MailMerge(template_path)
    document.merge(**merge_fields)
    document_obj = convert_to_doc(document)

    return document_obj


def main(data):

    header_template = "files/offer_templates/cover.docx"
    footer_template = "files/offer_templates/closing.docx"

    merge_fields = get_merge_fields(data)
    offer_header = merge_data(header_template, merge_fields)
    offer_footer = merge_data(footer_template, merge_fields)

    areas = data["areas"]

    section_files = []
    master_product_title = ""

    for line_item in areas:

        for opt_idx in range(len(areas[line_item])):

            option_data = areas[line_item][opt_idx]
            section_files.append(create_product_section(option_data))
            section_files.append(create_commercials_section(option_data))

            if not master_product_title:
                master_product_title = option_data['product']
            if option_data['product'] != master_product_title:
                master_product_title = "Louvers"

    combined_offer = combine_documents(offer_header, section_files, offer_footer)

    offer_data = data['offer_data']
    filename = "{}_Offer for Supply {}of {} for {}.docx".format(
        offer_data["OfferNumber"].replace("/", " ").replace("-", ""),
        "and Installation " if offer_data["installation"] else "",
        master_product_title,
        offer_data["FullName"],
    )

    output = BytesIO()
    combined_offer.save(output)
    output.seek(0)

    return (filename, output)
