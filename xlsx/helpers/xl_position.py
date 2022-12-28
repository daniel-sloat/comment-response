import re


def xl_position(cell_ref: str) -> tuple[str | str]:
    regex = r"^(?P<letters>\w*?)(?P<numbers>\d*)$"
    col_letters, row = re.search(regex, cell_ref).groups()
    col, _pow = 0, 1
    for letter in reversed(col_letters.upper()):
        col += (ord(letter) - ord("A") + 1) * _pow
        _pow *= 26
    return row, str(col)


def xl_position_reverse(column_int):
    start_index = 1  #  it can start either at 0 or at 1
    letter = ""
    while column_int > 25 + start_index:
        letter += chr(65 + int((column_int - start_index) / 26) - 1)
        column_int = column_int - (int((column_int - start_index) / 26)) * 26
    letter += chr(65 - start_index + (int(column_int)))
    return letter


class XlPosition:
    def __init__(self, reference):
        self.reference = reference
