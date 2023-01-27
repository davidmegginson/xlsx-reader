""" Utility methods to simplify working with the DOM """

def get_attr(attributes, name):
    try:
        return attributes.getValue(name)
    except KeyError:
        return None

def to_num(s):
    if '.' in s:
        return to_float(s)
    else:
        return to_int(s)

def to_int(s):
    try:
        return int(s)
    except TypeError:
        return s

def to_float(s):
    try:
        return float(s)
    except TypeError:
        return s

def to_bool(s):
    if s is None:
        return False
    else:
        return s.lower() in ('t', 'true', '1', 'yes', 'y')

def parse_col_num(cell_name):
    """ Calculate the column number from a cell name """
    col_num = 0
    for c in cell_name.upper():
        n = ord(c)
        if 65 <= n <= 90:
            col_num = col_num * 26 + n - 65
    return col_num

