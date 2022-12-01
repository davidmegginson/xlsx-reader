""" Utility methods to simplify working with the DOM """

def getChildText(node, name, default=None):
    """ Get the text for the first named child of a DOM node (if it exists)
    
    Parameters:
        node: the parent DOM element node
        name: the name of the child node
        default: the value to return if the child node doesn't exist (defaults to None)

    Return:
        The text of the first matching child element, or the default if none is found.

    """
    
    for child in node.childNodes:
        if child.nodeType == node.ELEMENT_NODE and child.tagName == name:
            return getText(child)
    return default
    

def getText(node):
    """ Get all the text at the top level of a DOM element
    Will skip comments and child elements

    Parameters:
        node: a DOM element node

    Return:
        The concatenated text of all child text or CDATA section nodes, or None if none exist.

    """

    s = None
    for child in node.childNodes:
        if child.nodeType in (node.TEXT_NODE, node.CDATA_SECTION_NODE):
            if s is None:
                s = child.data
            else:
                s += child.data
    return s
