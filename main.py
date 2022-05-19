# %% [markdown]
# Takes input from xlsx spredsheet that has comments listed per row with grouping categories specified into two or three headings, and outputs a docx document that is grouped by those headings, with comments listed under the lowest-tier heading and response following. The response must be on the first row of a grouping, with other response rows under the same grouping being blank.

# %%
# //TODO Create toml configuration file
# //TODO Make sure that the first nonempty response is taken for each comment grouping
# //TODO Finalize and create another requirements.txt with only needed requirements

# %%
from pathlib import Path
import pandas as pd
import docx_tools
import win32_tools
import ooxml
from config import *

# %%
def get_comment_index_tags(
    df_worksheet: pd.DataFrame, comment_tag_column: str
) -> pd.Series:
    """Create comment index tags to be appended to the end of comments for identification.
    Once comment index tags are created and appended, can be used with AutoMark to create index.

    Args:
        df_worksheet (pd.DataFrame): Dataframe representation of worksheet.

    Returns:
        pd.Series: Comment index tags.
    """
    # Regex captures two groups: (1) filename without extension, and (2) the file extension.
    # Also covers files starting with ".", common on Unix.
    regex = r"(?P<filename>.+?)(?P<ext>\.[^.]*$|$)"
    # Prefix and suffix added in attempt to make sure only unique identifiers are marked.
    replacement = f"zyx\\1xyz"
    comment_tags = df_worksheet[comment_tag_column].str.replace(
        regex, replacement, regex=True
    )
    comment_tags.name = "CommentTags"
    return comment_tags


# %%
def append_comment_tags(comment_column_list: list, comment_tags: pd.Series) -> list:
    """Appends tags to the end of each comment.

    Args:
        comment_column_list (list): Untagged comment list.
        comment_tags (pd.Series): Tags to append.

    Returns:
        list: Tagged comment list.
    """
    for cmt, tag in zip(comment_column_list, comment_tags):
        for para in cmt:
            if para == cmt[-1]:
                tag_run = ("", " (" + tag + ")")
                para.append(tag_run)
    return comment_column_list


# %%
def working_df(
    df_worksheet: pd.DataFrame,
    comment_col: list,
    response_col: list,
) -> pd.DataFrame:
    """A "working" dataframe that includes relevant columns
    for subsequent steps of combining and grouping to get
    into format suitable for writing docx file.

    Args:
        df_worksheet (pd.DataFrame):
            DataFrame of plain sharedStrings.
        comment_col (list): For rich text:
            [Comment[Paragraph[Run[Format,RunText]]]]
        response_col (list): For rich text:
            [Response[Paragraph[Run[Format,RunText]]]]

    Returns:
        pd.DataFrame: Focused dataframe used for grouping.
    """
    df_sheet = df_worksheet[[LEVEL_1, LEVEL_2, LEVEL_3]].copy()
    df_sheet[COMMENTS_COLUMN] = comment_col
    df_sheet[RESPONSE_COLUMN] = response_col
    return df_sheet


# %%
def group_comments_and_responses(df: pd.DataFrame) -> pd.DataFrame:
    """Group comments at lowest-level of hierarchy. Counts
    the number of comments in the group.

    Args:
        df (pd.DataFrame): Relevant dataframe (grouping columns,
        comment column)

    Returns:
        pd.DataFrame: Grouped at lowest-level hierarchy with
        comment count.
    """
    df_group = df.groupby([LEVEL_1, LEVEL_2, LEVEL_3], dropna=False)
    comments_grouped = df_group[COMMENTS_COLUMN].apply(tuple)
    comment_count = df_group[COMMENTS_COLUMN].count().rename("CommentCount")
    responses = df_group[RESPONSE_COLUMN].first()
    # Comment groups with no response will not be iterable (NoneType).
    # Replace with empty run: (Response[Para(Run)])
    empty_response = ([("", "")],)
    responses = responses.apply(
        lambda x: x if isinstance(x, tuple | list) else empty_response
    )
    response_count = df_group[RESPONSE_COLUMN].count().rename("ResponseCount")
    concat = pd.concat(
        [comments_grouped, comment_count, responses, response_count], axis=1
    ).reset_index()
    return concat


# %%
def check_response_count(df: pd.DataFrame) -> None:
    """Raises message regarding number of responses. If number
    of responses != 1, show error message.

    Args:
        responses_with_count (pd.DataFrame): Grouped with
        response count.
    """
    count = df["ResponseCount"]
    if count.max() > 1:
        print(
            "ERROR: More than one response for at least one comment group detected. "
            + "Keeping only the first response (which may not be desired)."
        )
    if count.min() < 1:
        print(
            "WARNING: No response for at least one comment group detected. "
            + "Empty response inserted."
        )


# %%
def sorting(sort: int = 0) -> tuple:
    match sort:
        case "alpha":
            print("Sorting alphabetically (ascending).")
            level3_columns = [LEVEL_1, LEVEL_2, LEVEL_3]
            level3_ascending = [True, True, True]
            level2_columns = [LEVEL_1, LEVEL_2]
            level2_ascending = [True, True]
        case "count":
            print("Sorting by comment count (descending).")
            level3_columns = [LEVEL_1, LEVEL_2, "CommentCount", LEVEL_3]
            level3_ascending = [True, True, False, True]
            level2_columns = [LEVEL_1, "CommentCount", LEVEL_2]
            level2_ascending = [True, False, True]
        case _:
            print("No sorting.")
            return None
    return level3_columns, level3_ascending, level2_columns, level2_ascending


# %%
def group_by_level(df: pd.DataFrame) -> tuple:
    LEVEL_3_DATA = "Level3Data"
    LEVEL_2_DATA = "Level2Data"
    LEVEL_1_DATA = "Level1Data"
    sort = sorting(SORT)

    def level3(df: pd.DataFrame) -> pd.DataFrame:
        if sort:
            df = df.sort_values(
                by=sort[0],
                ascending=sort[1],
            ).reset_index(drop=True)
        df[LEVEL_3] = df[LEVEL_3].fillna("Blank")
        df[LEVEL_3_DATA] = tuple(
            zip(df[COMMENTS_COLUMN], df[LEVEL_3], df[RESPONSE_COLUMN])
        )
        return df

    def level2(df: pd.DataFrame) -> pd.DataFrame:
        df_group = df.groupby([LEVEL_1, LEVEL_2])
        comments_level_2 = df_group[LEVEL_3_DATA].apply(tuple)
        df = pd.DataFrame(comments_level_2).reset_index()
        if sort:
            comment_count = df_group["CommentCount"].first()
            df = pd.merge(
                comments_level_2, comment_count, left_index=True, right_index=True
            ).reset_index()
            df = df.sort_values(
                by=sort[2],
                ascending=sort[3],
            ).reset_index(drop=True)
        df[LEVEL_2_DATA] = tuple(zip(df[LEVEL_3_DATA], df[LEVEL_2]))
        return df

    def level1(df: pd.DataFrame) -> tuple:
        df_group = df.groupby([LEVEL_1])
        comments_level_1 = df_group[LEVEL_2_DATA].apply(tuple)
        df = pd.DataFrame(comments_level_1).reset_index()
        if LEVEL_1_SORT:
            df_mapping = pd.DataFrame({"sort": LEVEL_1_SORT})
            sort_mapping = df_mapping.reset_index().set_index("sort")
            df["sort"] = df[LEVEL_1].map(sort_mapping["index"])
            df = df.sort_values("sort").reset_index()
        df[LEVEL_1_DATA] = tuple(zip(df[LEVEL_2_DATA], df[LEVEL_1]))
        return tuple(df[LEVEL_1_DATA])

    return level1(level2(level3(df)))


# %%
def mark_index_entries(comment_tags: list) -> None:
    """Mark index entries by creating AutoMark document
    and opening Word and marking entries, and adding
    index.

    Args:
        comment_tags (list): Comment tags to index.
    """
    regex = f"^zyx((\d+?)-(.+?))xyz$"
    index_entry = comment_tags.replace(regex, r"\1", regex=True)
    automark_list = list(zip(comment_tags, index_entry))
    docx_tools.automarkdoc(automark_list)
    # win32_tools requires Office to be installed.
    win32_tools.mark_index_entries(add_index=True)
    return None


# %%
def main():
    # Read ooxml file and retrieve relevant data
    ooxml_file = ooxml.SpreadSheetML(FILENAME)
    sheet = ooxml_file.sheet(SHEETNAME)
    coded_sheet = sheet.to_dataframe_codes()
    # Remove empty comment rows. All rows should have a comment associated with it.
    remove_empty_comment_rows = coded_sheet[COMMENTS_COLUMN].notna()
    coded_sheet = coded_sheet[remove_empty_comment_rows]
    sharedstrings_rich = sheet.get_rich_strings()
    df_worksheet = sheet.to_dataframe()
    df_worksheet = df_worksheet[remove_empty_comment_rows]
    comment_codes = coded_sheet[COMMENTS_COLUMN]
    response_codes = coded_sheet[RESPONSE_COLUMN]
    comment_tags = get_comment_index_tags(df_worksheet, COMMENT_TAGS_COLUMN)
    # Decode comment and response columns
    comment_code_data = ooxml.RichText(sharedstrings_rich, comment_codes)
    response_code_data = ooxml.RichText(sharedstrings_rich, response_codes)
    formats = list(
        set(comment_code_data.formats_used() + response_code_data.formats_used())
    )
    comment_column_list = comment_code_data.decode()
    response_column_list = response_code_data.decode()
    comment_column_list = append_comment_tags(comment_column_list, comment_tags)
    df_working = working_df(df_worksheet, comment_column_list, response_column_list)
    comments_and_response = group_comments_and_responses(df_working)
    check_response_count(comments_and_response)
    # Group headings, comments, and responses into multi-level list
    grouped = group_by_level(comments_and_response)
    docx_tools.commentsectiondoc(grouped, formats, levels=3)
    mark_index_entries(comment_tags)
    return None


# %%
if __name__ == "__main__":
    main()

# %%
