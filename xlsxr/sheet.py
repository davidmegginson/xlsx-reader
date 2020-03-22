""" Class representing an Excel XLSX sheet in a workbook

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-21

"""

import xml.dom.pulldom


class Sheet:
    """ An Excel XLSX worksheet (tab) """

    def __init__(self, index, worksheet):
        """ Open an Excel file.
        @param index: the sheet index (1-based)
        @param worksheet: the parent Worksheet object
        """

        self.index = index
        self.worksheet = worksheet


    def rows(self):

        pass
