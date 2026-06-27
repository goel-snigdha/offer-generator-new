from pathlib import Path
from mailmerge import MailMerge
from doc_utils import convert_to_doc

_BASE_DIR  = Path(__file__).resolve().parent.parent
_TEMPLATES = _BASE_DIR / "files" / "offer_templates"
_PROD_TEMPLATES = _TEMPLATES / "products"


_GRADE_TEMPLATE = {
    "mirror":    "mirror.docx",
    "lacquered": "lacquered.docx",
}
def _is_translucent(sheet_grade):
    grade = str(sheet_grade or "").lower()
    return "plain" in grade or "translucent" in grade


def _get_product_template(sheet_grade):
    grade = str(sheet_grade or "").lower()
    for key, tmpl in _GRADE_TEMPLATE.items():
        if key in grade:
            return tmpl
    return None


def _merge_section(template_path, fields=None):
    doc = MailMerge(str(template_path))
    if fields:
        doc.merge(**fields)
    return convert_to_doc(doc)


def _has_lights(option_data):
    """True only if the option's uploaded sheet has a non-empty Lights BOM.
    An empty Lights BOM means lights aren't part of this offer, so the lights
    product page should be skipped (mirrors get_lights_commercials)."""
    import openpyxl
    from stretch_ceilings.doc_builder.products._base import read_lights_bom

    wb_file = option_data.get("area_table")
    if wb_file is None:
        return False
    try:
        if hasattr(wb_file, "seek"):
            wb_file.seek(0)
        wb = openpyxl.load_workbook(wb_file, data_only=True)
        return bool(read_lights_bom(wb))
    except Exception:
        return False


def offer_has_lights(data):
    """True if any option across the offer includes lights (translucent sheet with a
    non-empty Lights BOM) — i.e. a lights page is actually part of this offer."""
    for line_item in data["areas"]:
        for option_data in data["areas"][line_item]:
            if _is_translucent(option_data.get("sheet_grade", "")) and _has_lights(option_data):
                return True
    return False


def create_product_section(data):

    areas      = data["areas"]
    offer_number = data["offer_data"]["OfferNumber"]
    section_files = []

    for line_item in areas:
        options = areas[line_item]

        # 1. Areas of work — once per line item
        sheet_grades = [o.get("sheet_grade", "") for o in options]
        product_desc = ", ".join(str(g) for g in sheet_grades if g)
        section_files.append(
            _merge_section(_TEMPLATES / "areas_of_work.docx", {
                "LineItem":         str(line_item),
                "OfferNumber":      offer_number,
                "ProductOptionsDesc": product_desc,
            })
        )

        for option_data in options:
            sheet_grade = option_data.get("sheet_grade", "")

            # 2. Product section
            tmpl = _get_product_template(sheet_grade)
            if tmpl:
                section_files.append(_merge_section(_PROD_TEMPLATES / tmpl))

            # 3. Area commercials
            section_files.append(
                _merge_section(_TEMPLATES / "area_commercials.docx", {
                    "OfferNumber":        offer_number,
                    "ProductOptionsDesc": str(sheet_grade),
                })
            )

            # 4. Lights page (translucent sheets that actually include lights)
            if _is_translucent(sheet_grade) and _has_lights(option_data):
                section_files.append(
                    _merge_section(_PROD_TEMPLATES / "lights.docx", {
                        "LightType": str(option_data.get("light_type") or ""),
                        "BoxDepth":  str(option_data.get("box_depth") or ""),
                    })
                )

    return "Stretch Ceilings", section_files
