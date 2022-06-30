from itertools import groupby
from pprint import pprint

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
        group = list(group)
        grouped_data = _group_comments_and_responses(group)
        if key[-1]:
            # print("IF")
            initial_grouping.append(
                {"sort": key[:-2], "heading": key[-1], "data": grouped_data}
            )
        else:
            data = {k: v for k, v in grouped_data.items() if k == "comment_data"}
            # print(group)
            initial_grouping.append({"sort": key[:-2], "data": data})
    return initial_grouping


from operator import itemgetter


def _sort_by_comment_count(initial_grouping):
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
    for dictionary in initial_grouping:
        
        def recursive(dictionary):
        
            if "data" in dictionary:
                
                if isinstance(dictionary["data"], list):
                    dictionary["data"] = _sort_by_comment_count(dictionary["data"])
                    
                elif isinstance(dictionary["data"], dict):
                    return recursive(dictionary["data"])
        
        recursive(dictionary)
        
    new_obj = sorted(initial_grouping, key=keysort)
                
    return new_obj


def _following_groupings(grouped_data, key_sort, continue_sort=True):
    # Group remaining levels until key is exhausted.
    def _grouper(grouped_data, key_sort):
        new_grouped_data = []
        for key, group in groupby(grouped_data, key=key_sort):
            group = list(group)
            # print(group)
            if key:
                # Remove used sort keys (only top-level sort key is used)

                # Append new (reduced) sort key)
                if key[-1]:
                    # print("YES")
                    # new_group = []
                    for g in group:
                        if "heading" not in g:
                            data = g["data"]
                        else:
                            data = group
                        # for k,v in g.items():
                        #     if k == "data":
                        #         comment_data = v
                    # print(new_group)
                    # pprint(group, sort_dicts=False, width=140)
                    new_grouped_data.append(
                        {"sort": key[:-2], "heading": key[-1], "data": data}
                    )
                else:
                    print("NO")
                    data = {k: v for k, v in group.items() if k == "data"}
                    new_grouped_data.append({"sort": key[:-2], "data": data})
            else:
                # If not sort key, grouping is complete, so extend the elements of
                # the group iterable to keep the same.
                new_grouped_data.extend([g for g in group])
        return new_grouped_data

    while continue_sort:
        grouped_data = _grouper(grouped_data, key_sort)
        continue_sort = any([x.get("sort") for x in grouped_data])
        # continue_sort = False
    # else:
    #     # When complete, remove "sort" key from top level (not necessary, but for cleanup)
    #     grouped_data = [
    #         {k: v for k, v in x.items() if k != "sort"} for x in grouped_data
    #     ]

    return grouped_data


def group_data(comment_response_data, sort):
    key_sort = lambda x: x["sort"]
    grouped_data = _initial_sort_and_group(comment_response_data, key_sort)
    # pprint(grouped_data[3:5], sort_dicts=False, width=140)

    grouped_data = _following_groupings(grouped_data, key_sort)
    # pprint(grouped_data[0:2], sort_dicts=False, width=146)
    #print(grouped_data[0:2])
    if sort["type"] == "count":
        grouped_data = _sort_by_comment_count(grouped_data)
    # pprint(grouped_data[14:16], sort_dicts=False, width=140)
    return grouped_data
