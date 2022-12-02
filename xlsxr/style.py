import xml.sax

from xlsxr.util import get_attr, to_bool

class Styles:

    def __init__(self, workbook, filename):
        self.workbook = workbook

        self.number_formats = {}
        self.cell_style_formats = []
        self.cell_formats = []
        self.cell_styles = []

        handler = Styles.__SAXHandler(self)
        with self.workbook.archive.open(filename, "r") as stream:
            xml.sax.parse(stream, handler)

    class __SAXHandler(xml.sax.handler.ContentHandler):

        def __init__(self, styles):
            super().__init__()
            self.__styles = styles

            self.__format = None
            
            self.__in_cellStyleXfs = False
            self.__in_cellXfs = False
            self.__in_xf = False

        def startElement(self, name, attributes):

            if name == 'numFmt':
                self.__styles.number_formats[get_attr(attributes, 'numFmtId')] = get_attr(attributes, 'formatCode')

            elif name == 'cellStyleXfs':
                self.__in_cellStyleXfs = True
                pass

            elif name == 'cellXfs':
                self.__in_cellXfs = True

            elif name == 'xf':
                self.__in_xf = True
                self.__format = {
                    "numFmtId": get_attr(attributes, "numFmtId"),
                    "applyProtection": to_bool(get_attr(attributes, "applyProtection")),
                }

            elif name == 'protection' and self.__in_xf:
                self.__format["protection"] = {
                    "locked": to_bool(get_attr(attributes, "locked")),
                    "hidden": to_bool(get_attr(attributes, "hidden")),
                }

            elif name == 'cellStyle':
                self.__styles.cell_styles.append({
                    "name": get_attr(attributes, "name"),
                    "xfId": get_attr(attributes, "xfId"),
                    "builtinId": get_attr(attributes, "builtinId"),
                })

        def endElement(self, name):

            if name == "cellStyleXfs":
                self.__in_cellStyleXfs = False

            elif name == "cellXfs":
                self.__in_cellXfs = False
                
            elif name == "xf":
                if self.__in_cellXfs:
                    self.__styles.cell_formats.append(self.__format)
                elif self.__in_cellStyleXfs:
                    self.__styles.cell_style_formats.append(self.__format)
                self.__format = None
