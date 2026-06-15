from pathlib import Path
from mailmerge import MailMerge
from doc_utils import convert_to_doc

_BASE_DIR          = Path(__file__).resolve().parent.parent
_TEMPLATES         = _BASE_DIR / "files" / "offer_templates" / "products"
_AREAS_TEMPLATE    = Path(__file__).resolve().parents[2] / "files" / "offer_templates" / "areas_commercials.docx"


def _is_translucent(sheet_grade):
    grade = str(sheet_grade or "").lower()
    return "plain" in grade or "translucent" in grade


def _merge_section(template_path, fields=None):
    doc = MailMerge(str(template_path))
    if fields:
        doc.merge(**fields)
    return convert_to_doc(doc)


def create_product_section(data):

    areas = data["areas"]
    section_files = []

    for line_item in areas:
        for opt_idx in range(len(areas[line_item])):
            option_data = areas[line_item][opt_idx]
            sheet_grade = option_data.get("sheet_grade", "")

            section_files.append(
                _merge_section(_AREAS_TEMPLATE, {
                    "LineItem":    str(option_data.get("line_item_str", "")),
                    "Option":      str(option_data.get("option_str", "")),
                    "OfferNumber": str(data["offer_data"]["OfferNumber"]),
                })
            )

            section_files.append(
                _merge_section(_TEMPLATES / "plain_translucent.docx")
            )

            if _is_translucent(sheet_grade):
                section_files.append(
                    _merge_section(_TEMPLATES / "lights.docx", {
                        "LightType": str(option_data.get("light_type") or ""),
                        "BoxDepth":  str(option_data.get("box_depth") or ""),
                    })
                )

    return "Stretch Ceilings", section_files
