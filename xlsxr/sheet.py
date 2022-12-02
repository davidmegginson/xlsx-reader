""" Class representing an Excel XLSX sheet in a workbook

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-21

"""

import datetime, logging, xml.sax

from xlsxr.util import getAtt

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
        self.raw_rows = None
        self.raw_merges = None

        
    @property
    def rows(self):
        """ Get the rows, parsing the sheet on demand """
        if self.raw_rows is None:
            self.__parse_sheet()
        return self.raw_rows

    @property
    def merges(self):
        """ Get the merges, parsing the sheet on demand """
        if self.raw_merges is None:
            self.__parse_sheet()
        return self.raw_merges

    def __parse_sheet(self):
        """ On-demand parsing of the sheet itself """

        self.raw_rows = []
        self.raw_merges = []
        handler = _SheetSAXHandler(self)

        with self.workbook.archive.open(self.filename) as stream:
            xml.sax.parse(stream, handler)


class _SheetSAXHandler(xml.sax.ContentHandler):

    def __init__(self, sheet):
        super().__init__()
        self.sheet = sheet
        self.workbook = sheet.workbook

        # Accumulators
        self.row = None
        self.datatype = None
        self.chunks = [] # we can reuse this list

        # Very simple parse context
        self.in_row = False
        self.in_c = False
        self.in_v = False
        self.in_is = False
        self.in_t = False

        
    def startElement(self, name, attributes):

        if name == 'row':
            self.in_row = True
            self.row = []

        elif name == 'c' and self.in_row:
            self.in_c = True
            self.datatype = getAtt(attributes, 't')
            self.style = getAtt(attributes, 's')

        elif name == 'v' and self.in_c:
            self.in_v = True

        elif name == 'is' and self.in_c:
            self.in_is = True

        elif name == 't' and self.in_is:
            self.in_t = True

        elif name == 'mergeCell':
            self.sheet.raw_merges.append(getAtt(attributes, 'ref'))


    def endElement(self, name):

        if name == 'row':
            in_row = False
            self.sheet.raw_rows.append(self.row)

        elif name == 'c' and self.in_row:
            in_c = False
            self.row.append(self.__make_value())
            self.chunks.clear()
            self.datatype = None
            self.style = None

        elif name == 'v' and self.in_c:
            self.in_v = False

        elif name == 'is' and self.in_c:
            self.in_is = False

        elif name == 't' and self.in_is:
            self.in_t = False


    def characters(self, content):

        if self.in_v or self.in_t:
            self.chunks.append(content)

            
    def __make_value(self):
        """ Figure out the scalar value to include for a cell 

        Uses the current type and style, and may look up styles and shared strings
        in the parent workbook.

        """

        # Special case: if we haven't seen any text chunks, return None
        if len(self.chunks) == 0:
            return None

        # Merge all the text chunks (more efficient than using + each time
        value = ''.join(self.chunks)

        if self.datatype == 'b': # boolean
            pass

        elif self.datatype == 'd': # date
            pass

        elif self.datatype == 'e': # error
            pass

        elif self.datatype == 'inlineStr':
            pass

        elif self.datatype == 'n': # number
            if self.workbook.convert_values:
                try:
                    if '.' in value:
                        value = float(value)
                    else:
                        value = int(value)
                except ValueError:
                    logger.warning("Cannot convert %s to a number", value)

        elif self.datatype == 's': # shared string
            value = self.workbook.shared_strings[int(value)]

        elif self.datatype == 'str': # simple inline string
            pass

        # return the modified value
        return value
