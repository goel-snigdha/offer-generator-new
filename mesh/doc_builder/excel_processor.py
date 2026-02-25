from .products import mesh

PRODUCT_FUNCTIONS = {
    "MESH": [mesh.get_data, mesh.convert]
}


def get_area_data(xl):

    area_data = {}

    val = xl.cell(row=1, column=2).value
    for product in PRODUCT_FUNCTIONS.keys():
        if val and product in val:
            area_data['product'] = product

    get_data_func = PRODUCT_FUNCTIONS.get(area_data["product"], [])[0]
    new_area_data = get_data_func(xl)
    area_data.update(new_area_data)

    return area_data


def product_convert(option):
    return PRODUCT_FUNCTIONS.get(option['product'], [])[1]
