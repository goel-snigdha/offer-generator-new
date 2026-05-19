def _find_header_row(ws):
    for row in ws.iter_rows(min_row=3):
        non_empty = [c for c in row if c.value is not None]
        if len(non_empty) >= 2:
            return row[0].row
    return None


def _find_area_columns(ws, header_row_idx):
    return {
        str(cell.value).strip(): cell.column
        for cell in ws[header_row_idx]
        if cell.value and "area" in str(cell.value).lower()
    }


def _find_total_row(ws, header_row_idx):
    for row in ws.iter_rows(min_row=header_row_idx + 1):
        for cell in row:
            if cell.value and str(cell.value).strip().lower() == "total":
                return cell.row
    return None


def get_areas_data(areas_ws):
    result = {}

    result["sheet_grade"] = next(
        (cell.value for cell in areas_ws[2] if cell.value is not None), None
    )

    header_row_idx = _find_header_row(areas_ws)
    if header_row_idx is None:
        return result

    area_cols = _find_area_columns(areas_ws, header_row_idx)
    total_row_idx = _find_total_row(areas_ws, header_row_idx)

    if total_row_idx is None or not area_cols:
        return result

    if len(area_cols) == 1:
        col_idx = next(iter(area_cols.values()))
        val = areas_ws.cell(row=total_row_idx, column=col_idx).value
        result["chargeable_area"] = val
        result["actual_area"] = val
    else:
        for col_name, col_idx in area_cols.items():
            val = areas_ws.cell(row=total_row_idx, column=col_idx).value
            if "chargeable" in col_name.lower():
                result["chargeable_area"] = val
            elif "actual" in col_name.lower():
                result["actual_area"] = val

    return result


def get_profiles_data(profiles_ws):
    result = {}

    header_row_idx = _find_header_row(profiles_ws)
    if header_row_idx is None:
        return result

    # Map each header to its column index
    col_map = {
        str(cell.value).strip(): cell.column
        for cell in profiles_ws[header_row_idx]
        if cell.value is not None
    }

    # Grade column: first column in the header map
    # Length column: whichever header contains "length"
    grade_col = next(iter(col_map.values()), None)
    length_col = next(
        (col for header, col in col_map.items() if "length" in header.lower()), None
    )

    if grade_col is None or length_col is None:
        return result

    profiles = []
    for row in profiles_ws.iter_rows(min_row=header_row_idx + 1):
        grade_val = profiles_ws.cell(row=row[0].row, column=grade_col).value
        length_val = profiles_ws.cell(row=row[0].row, column=length_col).value
        if grade_val is not None:
            profiles.append({"grade": grade_val, "length": length_val})

    result["profiles"] = profiles
    return result


def get_area_data(wb):
    areas_ws = wb.worksheets[0]
    profiles_ws = wb.worksheets[1]

    data = get_areas_data(areas_ws)
    data.update(get_profiles_data(profiles_ws))
    data["product"] = "Stretch Ceilings"
    return data


def product_convert(option):
    return None
