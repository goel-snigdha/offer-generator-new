FINISH_RATE_COLS = {
    "Mill": 2,
    "Powder Coated \n Single Color": 5,
    "Anodized": 8,
    "Wood": 8,
}


def get_orientation(ws, start_row=4, col=3):

    orientations = set()
    row = start_row

    while True:
        cell_value = ws.cell(row=row, column=col).value

        if not cell_value:
            break

        val = str(cell_value)
        if val in {"Horizontal", "Vertical"}:
            orientations.add(val)

        row += 1

    if not orientations:
        return ""
    elif len(orientations) == 1:
        return orientations.pop()
    else:
        return ", ".join(sorted(orientations))
