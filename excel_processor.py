import openpyxl
import excel_utils

import louvers.doc_builder.excel_processor as louvers_excel
import mesh.doc_builder.excel_processor as mesh_excel

GET_AREA_DATA = {
    "Aluminium Louvers": louvers_excel.get_area_data,
    "SS316 Ropes & Meshes": mesh_excel.get_area_data,
}
GET_CONVERT_FUNC = {
    "Aluminium Louvers": louvers_excel.product_convert,
    "SS316 Ropes & Meshes": mesh_excel.product_convert,
}


def convert(product_key, data):

    output_dfs = []
    output_xls = []
    found_options = False

    for area in data["areas"]:
        if len(data["areas"][area]) > 1:
            found_options = True

        options = data["areas"][area]

        line_item_str = str(area)
        option_str = ""

        for idx in range(len(options)):
            option = data["areas"][area][idx]

            wb = option["area_table"]
            window_wb = openpyxl.load_workbook(wb, data_only=True)

            option_str = ""
            if len(options) > 1:
                option_str = "Option " + excel_utils.number_to_alpha(idx + 1)

            option["line_item_str"] = line_item_str
            option["option_str"] = option_str
            option_data_from_xl = GET_AREA_DATA[product_key](window_wb.worksheets[0])
            option.update(option_data_from_xl)

            convert = GET_CONVERT_FUNC[product_key](option)

            if convert:
                output_df = convert(
                    window_wb, option, data['offer_data']["installation"]
                )
                output_dfs.append(output_df)

                line_item_title = option["product"] + " " + option["line_item_str"]
                line_item_title = line_item_title + option["option_str"].replace("Option ", "")
                filename = f"Commercials - {line_item_title}.xlsx"

                commercial_xl = excel_utils.generate_commercial_table(output_df)
                output_xls.append((filename, commercial_xl))
                option["CommercialTable"] = commercial_xl

            else:
                raise ValueError("This Excel could not be processed.")

    if not found_options:
        offer_num = data['offer_data']["OfferNumber"].replace("/", "-")
        combined_xl = (f"Commercials for {offer_num}.xlsx", excel_utils.combine_xls(output_dfs))
        output_xls.append(combined_xl)
        data["CombinedCommercials"] = combined_xl

    return output_xls, data
