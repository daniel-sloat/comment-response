from itertools import groupby
from copy import deepcopy

from open_office_xml.dataclasses import RichText


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
    # If the response is empty, insert blank response (a list with one RichText object).
    if not grouped_responses:
        grouped_responses = [RichText()]
    comments_and_responses["comment_data"]["comments"] = grouped_comments
    comments_and_responses["comment_data"]["response"] = grouped_responses
    return comments_and_responses


def _initial_sort(comment_response_data, sort_type, key_sort):
    
    if sort_type == "nosort":
        return sorted(comment_response_data, key=lambda x: x["sort"][0::2])
    else:
        return sorted(comment_response_data, key=key_sort)
    

def _initial_group(comment_response_data, key_sort):
    return [
        {"sort": key[1::2], "data": _group_comments_and_responses(group)}
        for key, group in groupby(comment_response_data, key=key_sort)
    ]


def _sort_by_comment_count(groups):
    # To sort by comment count, preserve the "sort" grouping first, then
    # sort by (descending) comment count, then alphabetically for those
    # headings that have the same comment count. Default without sorting by
    # comment count is just alphabetically.
    def keysort(x):
        if isinstance(x["data"], dict):
            return -len(x["data"]["comment_data"]["comments"])
        else:
            return float("inf")

    new_obj = []
    for dictionary in groups:

        def recursive(dictionary):

            if isinstance(dictionary.get("data"), list):
                dictionary["data"] = _sort_by_comment_count(dictionary["data"])

            elif isinstance(dictionary.get("data"), dict):
                return recursive(dictionary["data"])

        recursive(dictionary)

    new_obj = sorted(groups, key=keysort)

    return new_obj


def depth_of_sort(grouped_data, continue_sort=True):
    # Creates dictionary of group keys ("headings") with the number of subheadings under
    # the group key.

    def get_subheading_count(gd):
        sort_data = {}
        for key, group in groupby(gd, key=lambda x: x["sort"]):
            sort_data[key] = len(list(group))
        return sort_data

    subheadings_count = {}
    gd = deepcopy(grouped_data)
    while continue_sort:
        subheadings_count |= get_subheading_count(gd)
        for group in gd:
            group["sort"] = group["sort"][:-1]
        continue_sort = max([len(group["sort"]) for group in gd])

    return subheadings_count


def _fix_empty_sort(grouped_data, sort_data):
    # If all subheadings in the next level are blank, then remove empty heading.
    # Else (if there are other headings including the empty heading), keep the
    # empty heading.
    new_group = []
    for key, group in groupby(grouped_data, key=lambda x: x["sort"]):
        group = list(group)
        if sort_data[key[:-1]] == 1 and key[-1] == "":
            for g in group:
                g["sort"] = g["sort"][:-1]
            new_group.extend(group)
        else:
            new_group.extend(group)

    return new_group


def _following_groupings(grouped_data, key_sort, continue_sort=True):
    # Group remaining levels until key is exhausted.
    def _grouper(grouped_data, key_sort):
        new_grouped_data = []
        for key, group in groupby(grouped_data, key=key_sort):
            group = list(group)
            if key:
                data = group
                for g in group:
                    if "heading" not in g:
                        data = g["data"]
                new_grouped_data.append(
                    {"sort": key[:-1], "heading": key[-1], "data": data}
                )
            else:
                # If not sort key, grouping is complete, so extend the elements of
                # the group iterable to keep the same.
                new_grouped_data.extend([g for g in group])
        return new_grouped_data

    while continue_sort:
        grouped_data = _grouper(grouped_data, key_sort)
        continue_sort = any([x.get("sort") for x in grouped_data])

    return grouped_data


def group_data(comment_response_data, sort):
    key_sort = lambda x: x["sort"]
    grouped_data = _initial_sort(comment_response_data, sort["type"], key_sort)
    grouped_data = _initial_group(grouped_data, key_sort)
    sort_data = depth_of_sort(grouped_data)
    grouped_data = _fix_empty_sort(grouped_data, sort_data)
    grouped_data = _following_groupings(grouped_data, key_sort)

    if sort["comment_sort"] == "count":
        grouped_data = _sort_by_comment_count(grouped_data)

    return grouped_data
