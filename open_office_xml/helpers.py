def excel_col_name_to_number(col_index: str) -> int:
    col_num, pow = 0, 1
    for letter in reversed(col_index.upper()):
        col_num += (ord(letter) - ord("A") + 1) * pow
        pow *= 26
    return col_num
