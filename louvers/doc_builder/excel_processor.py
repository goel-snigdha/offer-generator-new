from .excel_utils import (
    get_orientation
)
from .products import (
    grille,
    cottal,
    fluted,
    aerofoil,
    slouvers,
    rectangular,
    beamc,
)
PRODUCT_FUNCTIONS = {
    "Grille": [grille.get_data, grille.convert],
    "Cottal": [cottal.get_data, cottal.convert],
    "Fluted": [fluted.get_data, fluted.convert],
    "Aerofoil": [aerofoil.get_data, aerofoil.convert],
    "S-Louvers": [slouvers.get_data, slouvers.convert],
    "Rectangular": [rectangular.get_data, rectangular.convert],
    "Beam C-Channel": [beamc.get_data, beamc.convert],
}


def get_area_data(xl):

    area_data = {}

    val = xl.cell(row=1, column=3).value
    for product in PRODUCT_FUNCTIONS.keys():
        if val and val.startswith(product):
            area_data['product'] = product

    if xl.cell(row=1, column=1).value.startswith("Beam C-Channel"):
        area_data['product'] = "Beam C-Channel"

    area_data['orientation'] = get_orientation(xl)
    get_data_func = PRODUCT_FUNCTIONS.get(area_data["product"], [])[0]
    new_area_data = get_data_func(xl)
    area_data.update(new_area_data)

    return area_data


def product_convert(option):
    return PRODUCT_FUNCTIONS.get(option['product'], [])[1]
