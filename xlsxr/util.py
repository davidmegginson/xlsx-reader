""" Utility methods """

def get_attr(attributes, name):
    """ Try looking up a DOM attribute, handling an exception """
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

def parse_cell_ref(s):
    """ Return a tuple of the row and column number, zero-based.
    D3 will return (3, 2)
    """
    row_num = None
    col_num = None
    for c in s.upper():
        n = ord(c)
        if 48 <= n <= 57: # a digit
            if row_num is None:
                row_num = n - 48
            else:
                row_num = row_num * 10 + n - 48
        elif 65 <= n <= 90: # a letter
            if col_num is None:
                col_num = n - 65
            else:
                col_num = (col_num + 1) * 26 + n - 65
    return (row_num - 1, col_num,)

def parse_cell_range(s):
    """ Parse a cell range, like D5:H11
    Returns a tuple of two tuples, zero-based, like ((3, 4,), (7, 10,),)
    """
    start, end, = s.split(':')
    start_ref = parse_cell_ref(start)
    end_ref = parse_cell_ref(end)
    return (start_ref, end_ref,)
