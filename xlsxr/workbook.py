""" Class representing an Excel XLSX workbook
Started by David Megginson, 2020-03-20

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-20

"""

import requests, shutil, tempfile, zipfile


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
            response = requests.get(url, stream=True)
            response.raise_for_status() # force an exception if there's a problem
            tmpfile = tempfile.TemporaryFile()
            shutil.copyfileobj(response.raw, tmpfile)
            self.archive = zipfile.ZipFile(tmpfile)
        else:
            raise ValueError("Must specify filename, stream, or url argument")

        self.setup() # will throw an exception if it's not an XLSX file
            

    def setup (self):
        """ Set up the workbook 
        @raises TypeError: if the zip file is not an XLSX file
        """
        for item in self.archive.infolist():
            if item.filename == "xl/workbook.xml":
                self.workbook_item = item
                return
        raise TypeError("Zip archive is not an Excel XLSX file")
            
