import re
from itertools import chain, groupby

from open_office_xml.dataclasses import RichText


def relevant_data(sheet, config_columns, clean=True):
    comment_response_data = []
    for row in sheet.group_by_row():
        sort_cols = []
        numbered_sort = []
        data = {}
        for col_data in row:
            for cat, col in config_columns["commentresponse"].items():
                if col_data.col_name == col:
                    if clean:
                       data[cat] = _clean_text(col_data.rich)
                    else:
                       data[cat] = col_data.rich
            for cat, col in config_columns["other"].items():
                if col_data.col_name == col:
                    data[cat] = col_data.value
            # 'is not None' is needed for this situation, as checking for truthy will also filter
            # out zeroes, which may be used for sorting purposes.
            if col_data.col_name in config_columns["numbered_sort"] and col_data.value is not None:
                numbered_sort.append(col_data.value)
            if col_data.col_name in config_columns["sort"] and col_data.value is not None:
                sort_cols.append(col_data.value)

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


def _clean_text(rich):
    for paragraph in rich.paragraphs:
        for run in paragraph.runs:
            run.text = re.sub(r"\s{2,}", " ", run.text)
            if run == paragraph.runs[-1]:
                run.text = run.text.rstrip()
            elif run == paragraph.runs[0]:
                run.text = run.text.lstrip()
    return rich


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


def _group_comments_and_responses(group):
    # Combine comments into list[list] (paragraphs still denoted by '\n')
    # Some response cells will be empty. In case there are multiple cells of
    # responses, combine into list as well (responses will be in the order
    # they are found).
    grouped_comments, grouped_responses = [], []
    comments_and_responses = {}
    comments_and_responses["comment_data"] = {}
    for g in group:
        # Selects only those rows with comments.
        if g["comment"].paragraphs:
            grouped_comments.append(g["comment"])
            # Only select responses attached to a comment
            if g["response"].paragraphs:
                grouped_responses.append(g["response"])
    # If the response is empty, insert blank response (a list of RichText).
    if not grouped_responses:
        grouped_responses = [RichText()]
    comments_and_responses["comment_data"]["comments"] = grouped_comments
    comments_and_responses["comment_data"]["response"] = grouped_responses
    return comments_and_responses


def _initial_sort_and_group(comment_response_data, key_sort):
    # Performs first (lowest-level) sorting and grouping. This is important 
    # because the comments and any responses are grouped. Sorting occurs
    # here before grouping, as the itertools groupby function requires that
    # to group all similar keys. After this initial sort and group, all that
    # remains is going through the remaining headings (the "comment data"
    # is all grouped here).
    comment_response_data = sorted(comment_response_data, key=key_sort)
    initial_grouping = []
    for key, group in groupby(comment_response_data, key=key_sort):
        grouped_data = _group_comments_and_responses(group)
        initial_grouping.append(
            {"sort": key[:-2], "heading": key[-1], "data": grouped_data}
        )
    return initial_grouping


def _sort_by_comment_count(initial_grouping):
    # To sort by comment count, preserve the "sort" grouping first, then
    # sort by (descending) comment count, then alphabetically for those
    # headings that have the same comment count. Default without sorting by
    # comment count is just alphabetically.
    sorted_grouping = sorted(
        initial_grouping,
        key=lambda x: (
            x["sort"],
            -len(x["data"]["comment_data"]["comments"]),
            x["heading"],
        ),
    )
    return sorted_grouping


def _following_groupings(grouped_data, key_sort, s=True):
    # Group remaining levels until key is exhausted.
    def _grouper(grouped_data,key_sort):
        new_grouped_data = []
        for key, group in groupby(grouped_data, key=key_sort):
            if key:
                # Remove used sort keys (only top-level sort key is used)
                data = [{k: v for k, v in g.items() if k != "sort"} for g in group]
                # Append new (reduced) sort key)
                new_grouped_data.append(
                    {"sort": key[:-2], "heading": key[-1], "data": data}
                )
            else:
                # If not sort key, grouping is complete, so extend the elements of
                # the group iterable to keep the same.
                new_grouped_data.extend([g for g in group])
        return new_grouped_data
    
    while s:
        grouped_data = _grouper(grouped_data, key_sort)
        s = any([x.get("sort") for x in grouped_data])

    return grouped_data


def group_data(comment_response_data, sort):
    key_sort = lambda x: x["sort"]
    comment_response_data = _create_sort(comment_response_data, sort)
    grouped_data = _initial_sort_and_group(comment_response_data, key_sort)
    
    if sort["type"] == "count":
        grouped_data = _sort_by_comment_count(grouped_data)

    grouped_data = _following_groupings(grouped_data, key_sort)
        
    return grouped_data
