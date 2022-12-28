from collections import UserDict

from lxml.etree import QName


class attrib_dict(UserDict):
    """UserDict that removes URI from element tag string in key."""

    def __setitem__(self, key, value):
        qkey = QName(key).localname
        self.data[qkey] = value


def get_attrib(prop_element):
    d = attrib_dict()
    if len(prop_element):
        key = QName(prop_element).localname
        d[key] = {"_element": attrib_dict(prop_element.attrib)}
        for prop in prop_element:
            key2 = QName(prop).localname
            if len(prop):
                d[key] |= get_attrib(prop)
            else:
                d[key] |= {key2: attrib_dict(prop.attrib)}
    return d


# def get_attrib(prop_element):
#     d = attrib_dict()
#     if len(prop_element):
#         key = QName(prop_element).localname
#         d[key] = {"_element": attrib_dict(prop_element.attrib)}
#         for prop in prop_element:
#             key2 = QName(prop).localname
#             if len(prop):
#                 d[key] |= get_attrib(prop)
#             else:
#                 d[key] |= {key2: attrib_dict(prop.attrib)}
#     else:
#         key = QName(prop_element).localname
#         d[key] = {"_element": attrib_dict(prop_element.attrib)}
#     return d
