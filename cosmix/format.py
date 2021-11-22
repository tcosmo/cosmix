from pint import Quantity


class Format:
    end = "\033[0m"
    underline = "\033[4m"
    bold = "\033[1m"


def format_quantity(q: Quantity):
    base_number = q.to_tuple()[0]
    len_integerpart = len(str(base_number).split(".")[0])

    try:
        return ("{:." + str(len_integerpart + 2) + "g~P}").format(q)
    except ValueError as e:
        return "{:~P}".format(q)


def gsheets_quantity_format(q: Quantity, show_units=True):
    # #.##\ "uL"
    base_number = q.to_tuple()[0]
    is_fractional = len(str(base_number).split(".")) > 1

    if show_units:
        if is_fractional:
            return '#.##\ "{:~P}"'.format(q.units)
        else:
            return '#\ "{:~P}"'.format(q.units)
    if is_fractional:
        return "#.##"
    return "#"


GSHEETS_BANDING_COLORS = {
    "headerColor": {
        "red": 0.742,
        "green": 0.742,
        "blue": 0.742,
        "alpha": 1,
    },
    "firstBandColor": {
        "red": 1,
        "green": 1,
        "blue": 1,
        "alpha": 1,
    },
    "secondBandColor": {
        "red": 0.952,
        "green": 0.952,
        "blue": 0.952,
        "alpha": 1,
    },
}
