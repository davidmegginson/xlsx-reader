""" Utility methods to simplify working with the DOM """

def getAttr(attributes, name):
    try:
        return attributes.getValue(name)
    except KeyError:
        return None

def makeInt(s):
    try:
        return int(s)
    except ValueError:
        return s

def makeFloat(s):
    try:
        return float(s)
    except ValueError:
        return s

def makeBool(s):
    if s is None:
        return False
    else:
        return s.lower() in ('t', 'true', '1', 'yes', 'y')
