import logging
from itertools import groupby
from copy import deepcopy

from open_office_xml.dataclasses import RichText


def _group_comments_and_responses(group):
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


def _initial_sort(comment_response_data, sort_type):
    if sort_type == "nosort":
        # This leaves headings sorted as found in the worksheet, but obeys
        # the numbered sort for worksheet and configuration file custom sort.
        return sorted(comment_response_data, key=lambda x: x["sort"][0::2])
    else:
        # Obey numbered sort, and sort alphabetically.
        return sorted(comment_response_data, key=lambda x: x["sort"])


def _initial_group(comment_response_data):
    group = [
        {"sort": key[1::2], "data": _group_comments_and_responses(group)}
        for key, group in groupby(comment_response_data, key=lambda x: x["sort"])
    ]
    sort_data = _depth_of_sort(group)
    group = _trim_headings(group, sort_data)
    return group


def _sort_by_comment_count(groups):
    # Sorts the last subheading by comment count, recursively.
    def keysort(x):
        # Prioritize all subheadings that have comments.
        # Deeper subheadings are at the end.
        if isinstance(x["data"], dict):
            return -len(x["data"]["comment_data"]["comments"])
        else:
            return float("inf")

    for dictionary in groups:

        def data_key_fork(dictionary):
            if isinstance(dictionary.get("data"), list):
                dictionary["data"] = _sort_by_comment_count(dictionary["data"])
            elif isinstance(dictionary.get("data"), dict):
                return data_key_fork(dictionary["data"])

        data_key_fork(dictionary)

    return sorted(groups, key=keysort)


def _depth_of_sort(grouped_data, continue_sort=True):
    # Creates flat dictionary of group keys ("headings") with the number of
    # subheadings under the group key.

    def get_subheading_count(gd):
        sort_data = {}
        for key, group in groupby(gd, key=lambda x: x["sort"]):
            sort_data[key] = len(list(group))
        return sort_data

    subheadings_count = {}
    grouped_data_copy = deepcopy(grouped_data)
    while continue_sort:
        subheadings_count |= get_subheading_count(grouped_data_copy)
        for group in grouped_data_copy:
            group["sort"] = group["sort"][:-1]
        continue_sort = max(len(group["sort"]) for group in grouped_data_copy)

    return subheadings_count


def _trim_headings(grouped_data, sort_data):
    # If all subheadings in the next level are blank, then remove empty heading.
    # Else (if there are other headings including the empty heading), keep the
    # empty heading.
    trimmed_data = []
    for key, group in groupby(grouped_data, key=lambda x: x["sort"]):
        group = list(group)
        if sort_data.get(key[:-1]) == 1 and key[-1] == "":
            for g in group:
                g["sort"] = g["sort"][:-1]
            trimmed_data.extend(group)
        else:
            trimmed_data.extend(group)

    return trimmed_data


def _following_groupings(
    grouped_data,
    continue_sort=True,
):
    # Group remaining levels until key is exhausted.
    def _grouper(grouped_data):
        new_grouped_data = []
        for key, group in groupby(grouped_data, key=lambda x: x["sort"]):
            group = list(group)
            if key:
                for g in group:
                    if "heading" not in g:
                        group = g["data"]
                new_grouped_data.append(
                    {
                        "sort": key[:-1],
                        "heading": key[-1],
                        "data": group,
                    }
                )
            else:
                # If not sort key, grouping is complete, so extend the elements of
                # the group iterable to keep the same.
                new_grouped_data.extend(g for g in group)
        return new_grouped_data

    while continue_sort:
        grouped_data = _grouper(grouped_data)
        continue_sort = any(x.get("sort") for x in grouped_data)

    return grouped_data

from pprint import pformat

def group_data(comment_response_data, sort):
    logging.info("Using sort configuration:\n" + pformat(sort, sort_dicts=False))
    grouped_data = _initial_sort(comment_response_data, sort["type"])
    grouped_data = _initial_group(grouped_data)
    grouped_data = _following_groupings(grouped_data)

    if sort["comment_sort"] == "count":
        grouped_data = _sort_by_comment_count(grouped_data)

    return grouped_data
