import re
from itertools import chain

from open_office_xml.dataclasses import RichText, Run


def _relevant_data(sheet, config_columns, remove_double_spaces=True):
    comment_response_data = []
    for row in sheet.group_by_row():
        sort_cols = []
        numbered_sort = []
        data = {}
        for col_data in row:
            
            for cat, col in config_columns["commentresponse"].items():
                if col_data.col_name == col:
                    data[cat] = _clean_text(col_data.rich, remove_double_spaces)
            for cat, col in config_columns["other"].items():
                if col_data.col_name == col:
                    data[cat] = col_data.value
            
            if col_data.col_name in config_columns["numbered_sort"]:
                if col_data.value is None:
                    col_data.value = float("inf")
                numbered_sort.append(col_data.value)
            if col_data.col_name in config_columns["sort"]:
                if col_data.value is None:
                    col_data.value = ""
                sort_cols.append(col_data.value)


        while len(sort_cols) < len(config_columns["sort"]):
            sort_cols.append("")
        # Make both sort columns equal in length. Assumes sort_cols will be
        # filled completely and correctly.
        while len(numbered_sort) < len(sort_cols):
            numbered_sort.append(float("inf"))
        data["numbered_sort"] = numbered_sort
        data["sort"] = sort_cols

        if not data.get("response"):
            data["response"] = RichText()
            
        comment_response_data.append(data)
        
    return comment_response_data


def _clean_text(rich, remove_double_spaces):
    for paragraph in rich.paragraphs:
        for run in paragraph.runs:
            if remove_double_spaces:
                run.text = re.sub(r"\s{2,}", " ", run.text)
            if run == paragraph.runs[-1]:
                run.text = run.text.rstrip()
            elif run == paragraph.runs[0]:
                run.text = run.text.lstrip()
    return rich


def _append_comment_tags(data, key="tag", props=""):
    # Appends comment tag to the end of the last paragraph of each comment.
    for d in data:
        text = f" ({d[key]})"
        for para in d["comment"].paragraphs:
            if para == d["comment"].paragraphs[-1]:
                para.runs.append(Run(props=props, text=text))
    return data


def _create_sort(data, sort):
    """Create sort key that combines the numbered sort and column sort."""
    # Use specified columns in the worksheet for custom sort.
    if sort["sort"] == "worksheet":
        for d in data:
            d["sort"] = tuple(chain.from_iterable(zip(d["numbered_sort"], d["sort"])))
            d.pop("numbered_sort")
    # Use the config file for custom sort
    else:
        order_level_1 = {v: k for k, v in enumerate(sort["custom_level_1"])}
        order_level_2 = {v: k for k, v in enumerate(sort["custom_level_2"])}

        for d in data:
            d_numsort = []
            for count, value in enumerate(d["sort"]):
                if count == 0 and value in order_level_1:
                    d_numsort.append(order_level_1[value])
                elif value in order_level_2:
                    d_numsort.append(order_level_2[value])
                else:
                    d_numsort.append(float("inf"))
            d["sort"] = tuple(chain.from_iterable(zip(d_numsort, d["sort"])))
            d.pop("numbered_sort")
    return data


def comment_data(sheet, config_file):
    data = _relevant_data(
        sheet,
        config_file["columns"],
        config_file["other"]["remove_double_spaces"],
    )
    data = _create_sort(data, config_file["sort"])
    data = _append_comment_tags(data)
    return data
