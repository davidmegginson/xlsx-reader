""" Class representing an Excel XLSX sheet in a workbook

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-21

"""

import logging

import xml.dom.pulldom

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

    @property
    def rows(self):
        with self.workbook.archive.open('xl/' + self.target) as stream:
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
                        elif t == 's': # shared string
                            v = self.workbook.shared_strings[int(v)]
                        elif t == 'str': # simple inline string
                            pass
                        row.append(v)
                        
