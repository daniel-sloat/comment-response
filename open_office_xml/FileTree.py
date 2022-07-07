from lxml import etree

import re
from pathlib import Path
from zipfile import ZipFile
from functools import cached_property


class FileTree:

    NAMESPACES = {
        "w": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "r1": "http://schemas.openxmlformats.org/package/2006/relationships",
        "re": "http://exslt.org/regular-expressions"
    }

    def __init__(self, filepath):
        self.filepath = Path(filepath)

    @cached_property
    def _xml_tree(self):
        with ZipFile(self.filepath, "r") as z:
            xml_tree = {}
            regex = r".+(?:\.xml|\.rels)$"
            for xml_file in (name for name in z.namelist() if re.search(regex, name)):
                xml_tree[xml_file] = etree.fromstring(z.read(xml_file))
            return xml_tree
