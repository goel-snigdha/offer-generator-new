from excel_utils import get_total_cols, get_max_row

INSTALLATION_COST = 200


def get_data(xl):

    return {}


def update_data(xl, curr_row):

    vars = get_data(xl)
    vars.update(get_total_cols(xl, curr_row, 6, 6))

    return vars


def generate_df(data, finish, rate):

    df = []
    string = f'Supply of Aluminium CNC Sheets in {finish} Finish at 2mm Thickness'
    area = round(data["Area (ft2)"], 0)

    df = [
        [
            string,
            area,
            "ft\u00b2",
            rate,
        ],
        ["Installation Charges", area, "ft\u00b2", 200]
    ]

    return df


def convert(window_wb, data, installation):

    window_xl = window_wb.worksheets[0]
    max_row = get_max_row(window_xl)
    vars = update_data(window_xl, max_row)

    rate = 950
    data = generate_df(vars, data['finish'], rate)

    return data
