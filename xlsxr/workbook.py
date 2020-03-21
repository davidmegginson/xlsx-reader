""" Class representing an Excel XLSX workbook
Started by David Megginson, 2020-03-20

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-20

"""

import requests, shutil, tempfile, zipfile

import xml.parsers.expat


class Workbook:
    """ An Excel XLSX workbook
    """

    def __init__(self, filename=None, stream=None, url=None):
        """ Open an Excel file.
        One of filename, stream, and url must be specified.
        @param filename: path to an Excel file on the local system.
        @param stream: file-like object (byte stream)
        @param url: web address of a remote Excel file
        """

        if filename is not None:
            self.archive = zipfile.ZipFile(filename)
        elif stream is not None:
            self.archive = zipfile.ZipFile(stream)
        elif url is not None:
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
        
        try:
            with self.archive.open("xl/workbook.xml", "r") as stream:
                self.parse_workbook(stream)
        except KeyError:
            raise TypeError("Zip archive is not an Excel XLSX workbook")

    def parse_workbook(self, stream):

        def start_element(name, atts):
            if name == 'sheet':
                self.sheet_info.append((atts['name'], atts['sheetId'],))

        parser = xml.parsers.expat.ParserCreate()
        parser.StartElementHandler = start_element
        parser.ParseFile(stream)
        
    @property
    def sheet_count(self):
        return len(self.sheet_info)

            
