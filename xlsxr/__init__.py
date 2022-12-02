""" Root package for xlsx-reader
Started by David Megginson, 2020-03-20

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-20

"""

import sys

if sys.version_info < (3,):
    raise RuntimeError("xlsx-reader requires Python 3 or higher")

from xlsxr.workbook import Workbook

SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

__version__="0.1"
"""Module version number
see https://www.python.org/dev/peps/pep-0396/
"""

