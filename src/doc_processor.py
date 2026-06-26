from io import BytesIO
from pathlib import Path
from mailmerge import MailMerge
from docxcompose.composer import Composer
from doc_utils import get_merge_fields, convert_to_doc

import louvers.doc_builder.doc_processor as louvers_doc
import mesh.doc_builder.doc_processor as mesh_doc
import stretch_ceilings.doc_builder.doc_processor as sc_doc

PRODUCT_SECTION_FUNC = {
    "Aluminium Louvers": louvers_doc.create_product_section,
    "SS316 Ropes & Meshes": mesh_doc.create_product_section,
    "Stretch Ceilings": sc_doc.create_product_section

}
BASE_DIR = Path(__file__).resolve().parent
PRODUCT_TEMPLATES = {
    "Aluminium Louvers": BASE_DIR/"louvers"/"files"/"offer_templates",
    "SS316 Ropes & Meshes": BASE_DIR/"mesh"/"files"/"offer_templates",
    "Stretch Ceilings": BASE_DIR/"stretch_ceilings"/"files"/"offer_templates"
}


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


def main(product, data):

    header_template = PRODUCT_TEMPLATES[product]/"cover.docx"
    footer_template = PRODUCT_TEMPLATES[product]/"closing.docx"

    offer_data = data['offer_data']
    merge_fields = get_merge_fields(data)
    offer_header = merge_data(header_template, merge_fields)
    offer_footer = merge_data(footer_template, merge_fields)
    _, section_files = PRODUCT_SECTION_FUNC[product](data)

    total_pv_path = PRODUCT_TEMPLATES[product] / "total_project_value.docx"
    if total_pv_path.exists():
        section_files.append(merge_data(total_pv_path, {"OfferNumber": offer_data["OfferNumber"]}))

    combined_offer = combine_documents(offer_header, section_files, offer_footer)
    filename = "Proposal for {}, {} - {}.docx".format(
        offer_data["ProjectName"],
        offer_data["ProjectCity"],
        offer_data["OfferNumber"].replace("/", " ").replace("-", ""),
    )

    output = BytesIO()
    combined_offer.save(output)
    output.seek(0)

    return (filename, output)
