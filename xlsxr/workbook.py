""" Class representing an Excel XLSX workbook

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-20
"""

import io, logging, requests, shutil, tempfile, xlsxr.style, xlsxr.sheet, xml.dom.pulldom, zipfile

logger = logging.getLogger(__name__)

SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

class Workbook:
    """ An Excel XLSX workbook
    """

    def __init__(self, filename=None, stream=None, url=None, convert_values=False, fill_merged=False):
        """ Open an Excel file.
        One of filename, stream, and url must be specified.

        Parameters:
            filename: path to an Excel file on the local system.
            stream: file-like object (byte stream)
            url: web address of a remote Excel file
            convert_values: if True, convert numbers and dates from strings to Python values (default is False)
            fill_merged: if True, fill merged areas with repeated values
        """

        self.convert_values = convert_values

        self.fill_merged = fill_merged

        if filename is not None:
            logger.debug("Opening from file %s", filename)
            self.archive = zipfile.ZipFile(filename, "r")
        elif stream is not None:
            logger.debug("Opening from a byte stream")
            self.archive = zipfile.ZipFile(stream, "r")
        elif url is not None:
            logger.debug("Opening from a URL %s", url)
            with requests.get(url, stream=True) as response:
                self.archive = zipfile.ZipFile(io.BytesIO(response.content))
        else:
            raise ValueError("Must specify filename, stream, or url argument")

        self.sheets = []
        """ List of xlxr.sheet.Sheet objects """

        self.shared_strings = []

        self.relations = dict()
        """ Dict of relations """

        self.styles = None
        """ Object of type xlsxr.style.Styles with style information """

        self.setup() # will throw an exception if it's not an XLSX file
            

    def setup(self):
        """ Set up the workbook 
        @raises TypeError: if the zip file is not an XLSX file

        """

        try:
            with self.archive.open("xl/_rels/workbook.xml.rels", "r") as stream:
                self.parse_rels(stream)
            with self.archive.open("xl/workbook.xml", "r") as stream:
                self.parse_workbook(stream)
        except KeyError:
            raise TypeError("Zip archive is not an Excel XLSX workbook")

        try:
            with self.archive.open("xl/sharedStrings.xml", "r") as stream:
                self.parse_shared_strings(stream)
        except KeyError:
            logger.info("No sharedStrings.xml in this workbook")
            
        self.styles = xlsxr.style.Styles(self, "xl/styles.xml")
        

    def parse_workbook(self, stream):
        """ Parse the workbook metadata

        Parameters:
            stream: a file-like object (from the archive)
        """
        doc = xml.dom.pulldom.parse(stream)
        
        for event, node in doc:
            if event == xml.dom.pulldom.START_ELEMENT:
                if node.namespaceURI == SPREADSHEETML_NS and node.localName == 'sheet':
                    name = node.getAttribute('name')
                    sheet_id = node.getAttribute('sheetId')
                    state = node.getAttribute('state')
                    relation_id = node.getAttributeNS(RELATIONSHIPS_NS, 'id')
                    filename = self.relations.get(relation_id)
                    if filename.startswith('/'):
                        filename = filename[1:]
                    else:
                        filename = 'xl/' + filename
                    logger.debug("Creating sheet %s", name)
                    sheet = xlsxr.sheet.Sheet(self, name, sheet_id, state, relation_id, filename)
                    self.sheets.append(sheet)
        logger.debug("Workbook has %d sheets", len(self.sheets))


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
                    self.relations[node.getAttribute('Id')] = node.getAttribute('Target')

        logger.debug("Workbook has %d relations", len(self.relations))

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



