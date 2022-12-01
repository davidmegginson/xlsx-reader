""" Class representing an Excel XLSX sheet in a workbook

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-21

"""

import datetime, logging, xml.dom.pulldom

logger = logging.getLogger(__name__)


class Sheet:
    """ An Excel XLSX worksheet (tab) """

    def __init__(self, index, workbook, sheet_info):
        """ Open an Excel file.
        @param index: the sheet index (1-based)
        @param worksheet: the parent Workbook object
        """

        self.index = index
        self.workbook = workbook
        self.name = sheet_info.get('name', None)
        self.sheetId = sheet_info.get('sheetId', None)
        self.state = sheet_info.get('state', None)
        self.rel_id = sheet_info.get('rel_id', None)
        self.target = workbook.rels[self.rel_id]
        if self.target.startswith('/'):
            self.target = self.target[1:]
        else:
            self.target = 'xl/' + self.target

    @property
    def rows(self):
        with self.workbook.archive.open(self.target) as stream:
            row = None
            t = None
            v = None
            doc = xml.dom.pulldom.parse(stream)
            for event, node in doc:
                if event == xml.dom.pulldom.START_ELEMENT:
                    if node.localName == 'row':
                        row = []
                    elif node.localName == 'c':
                        t = node.getAttribute('t')
                    elif node.localName == 'v':
                        v = ''
                elif event == xml.dom.pulldom.CHARACTERS:
                    if node.data is None:
                        v = None
                    elif v is None:
                        v = node.data
                    else:
                        v += node.data
                elif event == xml.dom.pulldom.END_ELEMENT:
                    if node.localName == 'row':
                        yield row
                    elif node.localName == 'c':
                        if t == 'b': # boolean
                            pass
                        elif t == 'd': # date
                            pass
                        elif t == 'e': # error
                            pass
                        elif t == 'inlineStr': # TODO complex inline string
                            pass
                        elif t == 'n': # number
                            if self.workbook.convert_values:
                                try:
                                    if '.' in v:
                                        v = float(v)
                                    else:
                                        v = int(v)
                                except ValueError:
                                    logger.warning("Cannot convert %s to a number", v)
                        elif t == 's': # shared string
                            v = self.workbook.shared_strings[int(v)]
                        elif t == 'str': # simple inline string
                            pass
                        row.append(v)
                        
