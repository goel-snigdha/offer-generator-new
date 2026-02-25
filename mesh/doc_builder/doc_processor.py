from pathlib import Path
from mailmerge import MailMerge
from doc_utils import convert_to_doc


def merge_product_section_vars(data):
    BASE_DIR = Path(__file__).resolve().parent.parent
    path = BASE_DIR/"files"/"offer_templates"/"mesh.docx"
    section_doc = MailMerge(path)

    section_doc.merge(
        SSType=data["ss_type"],
        Opening=str(data["mesh_size"]) + "mm",
        Colour=str(data["colour"]),
        RopeDia=str(data["rope_dia"]) + "mm",
        OfferNumber=data["OfferNumber"]
    )

    section_doc_obj = convert_to_doc(section_doc)

    return section_doc_obj


def create_commercials_section(data):

    BASE_DIR = Path(__file__).resolve().parents[2]
    section_template = BASE_DIR/"files"/"offer_templates"/"areas_commercials.docx"

    section_doc = MailMerge(section_template)

    section_doc.merge(
        LineItem=data["line_item_str"],
        Option=data["option_str"],
        OfferNumber=data["OfferNumber"]
    )

    section_doc_obj = convert_to_doc(section_doc)

    return section_doc_obj


def create_product_section(data):

    areas = data["areas"]

    section_files = []
    master_product_title = ""

    for line_item in areas:

        for opt_idx in range(len(areas[line_item])):

            option_data = areas[line_item][opt_idx] | {'OfferNumber': data["offer_data"]["OfferNumber"]}
            section_files.append(merge_product_section_vars(option_data))
            section_files.append(create_commercials_section(option_data))

            master_product_title = "Mesh"

    return master_product_title, section_files
