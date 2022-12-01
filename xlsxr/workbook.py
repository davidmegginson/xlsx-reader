""" Class representing an Excel XLSX workbook

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-20
"""

import logging, requests, shutil, tempfile, xlsxr.sheet, xml.dom.pulldom, zipfile

logger = logging.getLogger(__name__)


class Workbook:
    """ An Excel XLSX workbook
    """

    NAMESPACE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def __init__(self, filename=None, stream=None, url=None, convert_values=False):
        """ Open an Excel file.
        One of filename, stream, and url must be specified.

        Parameters:
            filename: path to an Excel file on the local system.
            stream: file-like object (byte stream)
            url: web address of a remote Excel file
            convert_values: if True, convert numbers and dates from strings to Python values (default is False)
        """

        self.convert_values = convert_values

        if filename is not None:
            logger.debug("Opening from file %s", filename)
            self.archive = zipfile.ZipFile(filename)
        elif stream is not None:
            logger.debug("Opening from a byte stream")
            self.archive = zipfile.ZipFile(stream)
        elif url is not None:
            logger.debug("Opening from a URL %s", url)
            tmpfile = tempfile.TemporaryFile()
            with requests.get(url, stream=True) as response:
                response.raise_for_status() # force an exception if there's a problem
                shutil.copyfileobj(response.raw, tmpfile)
            self.archive = zipfile.ZipFile(tmpfile)
        else:
            raise ValueError("Must specify filename, stream, or url argument")

        self.setup() # will throw an exception if it's not an XLSX file
            

    def setup(self):
        """ Set up the workbook 
        @raises TypeError: if the zip file is not an XLSX file
        """

        self.sheet_info = list()
        self.shared_strings = list()
        self.rels = dict()
        
        try:
            with self.archive.open("xl/workbook.xml", "r") as stream:
                self.parse_workbook(stream)
            with self.archive.open("xl/sharedStrings.xml", "r") as stream:
                self.parse_shared_strings(stream)
            with self.archive.open("xl/_rels/workbook.xml.rels", "r") as stream:
                self.parse_rels(stream)
        except KeyError:
            raise TypeError("Zip archive is not an Excel XLSX workbook")


    def get_sheet(self, index):
        if index < 1 or index > self.sheet_count:
            raise IndexError("Sheet index out of range")
        sheet_info = self.sheet_info[index - 1]
        logger.debug("Opening sheet %d", index)
        return xlsxr.sheet.Sheet(index, self, sheet_info)


    @property
    def sheet_count(self):
        return len(self.sheet_info)

    def parse_workbook(self, stream):
        """ Parse the workbook metadata """
        doc = xml.dom.pulldom.parse(stream)
        for event, node in doc:
            if event == xml.dom.pulldom.START_ELEMENT and node.localName == 'sheet':
                self.sheet_info.append({
                    'name': node.getAttribute('name'),
                    'sheetId': node.getAttribute('sheetId'),
                    'state': node.getAttribute('state'),
                    'rel_id': node.getAttribute('r:id'), # FIXME use namespace
                })
        logger.debug("Workbook has %d sheets", self.sheet_count)


    def parse_shared_strings(self, stream):
        """ Parse the workbook shared strings """
        
        in_t = False # reading actual text
        text = None # text accumulator
        doc = xml.dom.pulldom.parse(stream)

        for event, node in doc:
            if event == xml.dom.pulldom.START_ELEMENT:
                if node.localName == 'si':
                    text = None
                elif node.localName == 't':
                    in_t = True
            elif event == xml.dom.pulldom.END_ELEMENT:
                if node.localName == 'si':
                    self.shared_strings.append(text)
                    text = None
                elif node.localName == 't':
                    in_t = False
            elif event == xml.dom.pulldom.CHARACTERS:
                if text is None:
                    text = node.data
                else:
                    text.append(node.data)

        logger.debug("Workbook has %d shared strings", len(self.shared_strings))

    def parse_rels(self, stream):
        """ Parse the workbook relations """

        doc = xml.dom.pulldom.parse(stream)

        for event, node in doc:
            if event == xml.dom.pulldom.START_ELEMENT:
                if node.localName == 'Relationship':
                    self.rels[node.getAttribute('Id')] = node.getAttribute('Target')

        logger.debug("Workbook has %d relations", len(self.rels))

    def parse_styles(self, stream):
        """ Parse the workbook styles """

        doc = xml.dom.pulldom.parse(stream)
        in_cellXfs = False
        current_style = {}

        for event, node in doc:
            if event == xml.dom.pulldom.START_ELEMENT:
                if node.localName == 'cellXfs':
                    in_cellXfs = True
                elif node.localName == 'xf' and in_cellXfs:
                    pass
            elif event == xml.dom.pulldom.END_ELEMENT:
                if node.localName == 'cellXfs':
                    in_cellXfs = False



            
