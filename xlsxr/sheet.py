""" Class representing an Excel XLSX sheet in a workbook

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-21

"""

import datetime, logging, xml.sax

from xlsxr.util import getAttr, makeFloat, makeInt, makeBool

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

        # information from the workbook
        self.workbook = workbook
        self.name = name
        self.sheet_id = sheet_id
        self.state = state
        self.relation_id = relation_id
        self.filename = filename

        # information parsed from the spreadsheet XML file on demand
        self._raw_cols = None
        self._raw_rows = None
        self._raw_merges = None

    def get_col(self, index):
        """ Get info about a single column

        Parameters:
          index(int): the 0-based index of the column

        Return:
          A dict with column info, or None if there was no match

        """
        for col in self.cols:
            if col["min"] <= index + 1 <= col["max"]:
                return col
        return None

    @property
    def cols(self):
        """ Get the columns, parsing the sheet on demand """
        if self._raw_cols is None:
            self.__parse_sheet()
        return self._raw_cols

    @property
    def rows(self):
        """ Get the rows, parsing the sheet on demand """
        if self._raw_rows is None:
            self.__parse_sheet()
        return self._raw_rows

    @property
    def merges(self):
        """ Get the merges, parsing the sheet on demand """
        if self._raw_merges is None:
            self.__parse_sheet()
        return self._raw_merges

    def __parse_sheet(self):
        """ On-demand parsing of the sheet itself """

        handler = Sheet.__SheetSAXHandler(self)

        with self.workbook.archive.open(self.filename) as stream:
            xml.sax.parse(stream, handler)


    class __SheetSAXHandler(xml.sax.ContentHandler):
        """ SAX content handler for parsing a sheet XML file

        Populates the following lists in the parent sheet:
        
        - _raw_cols
        - _raw_rows
        - _raw_merges

        TODO: add XML Namespace support

        """

        def __init__(self, sheet):
            super().__init__()
            self.__sheet = sheet
            self.__workbook = sheet.workbook

            # Reset accumulators in parent sheet
            self.__sheet._raw_cols = []
            self.__sheet._raw_rows = []
            self.__sheet._raw_merges = []

            # Local accumulators for the handler
            self.__row = None
            self.__datatype = None
            self.__chunks = [] # we can reuse this list

            # Very simple parse context
            self.__in_row = False
            self.__in_c = False
            self.__in_v = False
            self.__in_is = False
            self.__in_t = False


        def startElement(self, name, attributes):

            if name == 'col':
                self.__sheet._raw_cols.append({
                    "collapsed": makeBool(getAttr(attributes, "collapsed")),
                    "hidden": makeBool(getAttr(attributes, "hidden")),
                    "min": makeInt(getAttr(attributes, "min")),
                    "max": makeInt(getAttr(attributes, "max")),
                    "style": getAttr(attributes, "style"),
                })

            if name == 'row':
                self.__in_row = True
                self.__row = []

            elif name == 'c' and self.__in_row:
                self.__in_c = True
                self.__datatype = getAttr(attributes, 't')
                self.style = getAttr(attributes, 's')

            elif name == 'v' and self.__in_c:
                self.__in_v = True

            elif name == 'is' and self.__in_c:
                self.__in_is = True

            elif name == 't' and self.__in_is:
                self.__in_t = True

            elif name == 'mergeCell':
                self.__sheet._raw_merges.append(getAttr(attributes, 'ref'))


        def endElement(self, name):

            if name == 'row':
                self.__in_row = False
                self.__sheet._raw_rows.append(self.__row)

            elif name == 'c' and self.__in_row:
                self.__in_c = False
                self.__row.append(self.__make_value())
                self.__chunks.clear()
                self.__datatype = None
                self.style = None

            elif name == 'v' and self.__in_c:
                self.__in_v = False

            elif name == 'is' and self.__in_c:
                self.__in_is = False

            elif name == 't' and self.__in_is:
                self.__in_t = False


        def characters(self, content):

            if self.__in_v or self.__in_t:
                self.__chunks.append(content)


        def __make_value(self):
            """ Figure out the scalar value to include for a cell 

            Uses the current type and style, and may look up styles and shared strings
            in the parent workbook.

            """

            # Special case: if we haven't seen any text chunks, return None
            if len(self.__chunks) == 0:
                return None

            # Merge all the text chunks (more efficient than using + each time
            value = ''.join(self.__chunks)

            if self.__datatype == 'b': # boolean
                pass

            elif self.__datatype == 'd': # date
                pass

            elif self.__datatype == 'e': # error
                pass

            elif self.__datatype == 'inlineStr':
                pass

            elif self.__datatype == 'n': # number
                if self.__workbook.convert_values:
                    try:
                        if '.' in value:
                            value = float(value)
                        else:
                            value = int(value)
                    except ValueError:
                        logger.warning("Cannot convert %s to a number", value)

            elif self.__datatype == 's': # shared string
                value = self.__workbook.shared_strings[int(value)]

            elif self.__datatype == 'str': # simple inline string
                pass

            # return the modified value
            return value
