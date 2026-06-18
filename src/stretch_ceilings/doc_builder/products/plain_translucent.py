from ._base import AREA_UNITS, read_chargeable_area, read_bom, get_profile_rate

SHEET_RATE   = 550
INSTALL_RATE = 200


def generate_df(bom, chargeable_area, installation):

    df = []

    for item in bom:
        name = item["description"]
        qty  = item["qty"]
        unit = item["unit"]

        if unit.lower().replace(" ", "") in AREA_UNITS:
            df.append([f"Supply of {name.title()}", qty, "ft²", SHEET_RATE])
        else:
            rate = get_profile_rate(name)
            df.append([name.title(), qty, unit, rate if rate is not None else 1])

    if installation:
        df.append(["Installation Charges", chargeable_area, "ft²", INSTALL_RATE])

    return df


def convert(window_wb, data, installation):

    areas_ws = window_wb.worksheets[0]
    bom_ws   = window_wb.worksheets[1]

    chargeable_area    = read_chargeable_area(areas_ws)
    data["total_area"] = chargeable_area

    bom = read_bom(bom_ws)
    return generate_df(bom, chargeable_area, installation)
