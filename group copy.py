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
            # if col_data.row == 137:
            #    print(col_data)

            # FIX:
            # Apparently some cells are not written to spreadsheet, so
            # there can be a missing column value for 'response', because
            # the cell is empty. If can't find value, then need to insert blank value.
            for cat, col in config_columns["commentresponse"].items():
                if col_data.col_name == col:
                    data[cat] = col_data.rich
                    # if clean:
                    #    data[cat] = _clean_text(col_data.rich)
                    # else:
                    #    data[cat] = col_data.rich
            for cat, col in config_columns["other"].items():
                if col_data.col_name == col:
                    data[cat] = col_data.value
            if col_data.col_name in config_columns["numbered_sort"] and col_data.value:
                numbered_sort.append(col_data.value)
            if col_data.col_name in config_columns["sort"] and col_data.value:
                sort_cols.append(col_data.value)
                
        if not numbered_sort:
            data["numbered_sort"] = [float("inf"), float("inf"), float("inf")]
        else:
            while len(numbered_sort) < 3:
                numbered_sort.append(float("inf"))
            data["numbered_sort"] = numbered_sort
            
        while len(sort_cols) < 3:
            sort_cols.append(None)
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


def relevant_data2(sheet, config_columns):
    comment_response_data = []
    for row in sheet.group_by_row2():
        sort_cols = []
        data = {}
        for col_data in row:
            for cat, col in config_columns["commentresponse"].items():
                if col_data.col_name == col:
                    data[cat] = col_data.rich
            for cat, col in config_columns["other"].items():
                if col_data.col_name == col:
                    data[cat] = col_data.value
            if col_data.col_name in config_columns["sort"]:
                sort_cols.append(col_data.value)
        data["sort"] = sort_cols
        comment_response_data.append(data)
    return comment_response_data


def _custom_sort(data, sort):

    if sort["sort"] == "worksheet":
        for d in data:
            d["sort"] = tuple(chain.from_iterable(zip(d["numbered_sort"], d["sort"])))

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
    comments_and_responses["length"] = -len(grouped_comments)
    comments_and_responses["comment_data"]["comments"] = grouped_comments
    comments_and_responses["comment_data"]["response"] = grouped_responses
    return comments_and_responses


def _initial_sort_and_group(comment_response_data, key_sort):
    comment_response_data = sorted(comment_response_data, key=key_sort)
    initial_grouping = []
    for key, group in groupby(comment_response_data, key=key_sort):
        grouped_data = _group_comments_and_responses(group)
        initial_grouping.append(
            {"sort": key[:-2], "heading": key[-1], "data": grouped_data}
        )
    return initial_grouping


def _sort_by_comment_count(initial_grouping):
    
    sorted_grouping = sorted(
        initial_grouping,
        key=lambda x: (
            x["sort"][:-2],
            -len(x["data"]["comment_data"]["comments"]),
            x["heading"],
        ),
    )
    return sorted_grouping


def _sort_by_comment_count2(initial_grouping):
    sorted_grouping = sorted(
        initial_grouping,
        key=lambda x: (
            x["sort"][:-2],
            -len(x["data"]["comment_data"]["comments"]),
            x["heading"],
        ),
    )
    return sorted_grouping


def _following_groupings(grouped_data, key_sort):
    new_grouped_data = []
    for key, group in groupby(grouped_data, key=key_sort):
        combined = []
        for g in group:
            if not g.get("heading"):
                combined = g["data"]
            else:
                combined.append(g)
            g.pop("sort")
        #print(key)
        if key[:-2]:
            new_grouped_data.append(
                {"sort": key[:-2], "heading": key[-1], "data": combined}
            )
        else:
            new_grouped_data.append({"heading": key[-1], "data": combined})
            #continue
    return new_grouped_data


from pprint import pprint


def group_data(comment_response_data, sort):
    key_sort = lambda x: x["sort"]
    comment_response_data = _custom_sort(comment_response_data, sort)
    # pprint(comment_response_data,sort_dicts=False,width=148)
    combo_list = _initial_sort_and_group(comment_response_data, key_sort)
    #pprint(combo_list,sort_dicts=False,width=148)
    if sort["type"] == "count":
        combo_list = _sort_by_comment_count(combo_list)
        #pprint(combo_list,sort_dicts=False,width=148)
    s = max([len(x["sort"]) for x in combo_list])

    while s:
        #print(s)
        # pprint(combo_list[0],sort_dicts=False,width=148)
        combo_list = _following_groupings(combo_list, key_sort)
        s = max([len(x.get("sort",[])) for x in combo_list])
    return combo_list
