import pint
from rctreportviewer.constants import path_to_ureg

ureg = pint.UnitRegistry(path_to_ureg, autoconvert_offset_to_baseunit=True)


def convert_unit(value, from_unit, to_unit):
    """Convert a numerical value from one unit to another and return the magnitude."""
    quantity = value * ureg[from_unit]
    return quantity.to(ureg[to_unit]).magnitude
