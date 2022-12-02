""" Utility methods to simplify working with the DOM """

def getAtt(attributes, name):
    try:
        return attributes.getValue(name)
    except KeyError:
        return None
