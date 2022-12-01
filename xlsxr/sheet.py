""" Class representing an Excel XLSX sheet in a workbook

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-21

"""

import datetime, logging, xlsxr.util, xml.dom.pulldom

logger = logging.getLogger(__name__)


class Sheet:
    """ An Excel XLSX worksheet (tab) """

    def __init__(self, workbook, name, sheet_id, state, relation_id, filename):
        """ Open a sheet inside an Excel workbook.

        @param workbook: the parent Workbook object
        @param name: the sheet name
        @param sheet_id: the sheet identifier
        @param state: the sheet state (normally 'visible')
        @param relation_id: the relation identifier for filename lookup
        @param filename: the resolved sheet filename

        """
        
        self.workbook = workbook
        self.name = name
        self.sheet_id = sheet_id
        self.state = state
        self.relation_id = relation_id
        self.filename = filename

        
    @property
    def rows(self):
        """Parse the rows on demand each time there is a request.
        Uses a streaming parsing model to minimise memory usage.

        If the parent workbook's convert_values flag is True, then convert
        number strings to numbers and date values to date objects; otherwise,
        everything will be a string.

        It is safe to repeat calls to iterate over the same sheet
        multiple times.

        Returns:
            An iterable over lists of scalar values, each representing a row

        """
        
        with self.workbook.archive.open(self.filename) as stream:

            # Used to construct each row
            row = None

            # The streaming XML parser
            doc = xml.dom.pulldom.parse(stream)

            # Walk through the events
            for event, node in doc:

                
                if event == xml.dom.pulldom.START_ELEMENT:

                    # start a new row (don't expand, in case there are many columns)
                    if node.localName == 'row':
                        row = []

                    # extract a value (expands the node)
                    elif node.localName == 'c':
                        doc.expandNode(node)
                        row.append(self.get_value(node))
                        
                elif event == xml.dom.pulldom.END_ELEMENT:

                    # finish the row and yield it to the iterator
                    if node.localName == 'row':
                        yield row
                        

    def get_value(self, node):
        """ Clean up a value according to the datatype and the workbook's convert_values flag.

        Parameters:
            datatype: the value's data type (b, d, e, inlineStr, n, s, or str)
            value: the value as a string

        Return:
            The fixed value, possibly converted to a number or date object (if requested),
            and looked up if it's a shared string.

        """

        datatype = node.getAttribute('t')
        value = xlsxr.util.getChildText(node, 'v')

        # Handle the value based on the datatype (unless it's None)

        if datatype == 'b': # boolean
            pass

        elif datatype == 'd': # date
            pass

        elif datatype == 'e': # error
            pass

        elif datatype == 'inlineStr': # TODO complex inline string
            inline_node = xlsxr.util.getChild(node, "is")
            if inline_node is not None:
                value = xlsxr.util.getChildText(inline_node, "t")

        elif datatype == 'n': # number
            if self.workbook.convert_values:
                try:
                    if '.' in value:
                        value = float(value)
                    else:
                        value = int(value)
                except ValueError:
                    logger.warning("Cannot convert %s to a number", value)

        elif datatype == 's': # shared string
            value = self.workbook.shared_strings[int(value)]

        elif datatype == 'str': # simple inline string
            pass

        # return the modified value
        return value
