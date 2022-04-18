# %%
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# %%
from pathlib import Path
import pandas as pd
#from IPython.display import display

# %%
# Import other python module functions
import docx_tools
import win32_tools
import ooxml

# %%
# DEFINE REQUIRED INFORMATION
# Specify filename (in current working directory) and sheet.
COMMENT_RESPONSE_XLSX_FILENAME = "Alex_45 Day_Comments_Ref.xlsx"
# COMMENT_RESPONSE_SHEET_NUM actually refers to the sheet number ("sheet#"), not name //TODO
COMMENT_RESPONSE_SHEET_NUM = "sheet1"
# Specify relevant columns.
COMMENT_COLUMN = "CommentText"
RESPONSE_COLUMN = "draft Agency Response"
BATCH_STATUS_COLUMN = "Batch status"
RESPONSE_STATUS_COLUMN = "response status"
RESPONSE_NOTES_COLUMN = "response notes & questions"
# Column used to create index comment codes.
COMMENT_TAG_COLUMN = "File Name"
LEVEL_1 = "FSOR section heading level 1"
LEVEL_2 = "FSOR section heading level 2"
LEVEL_3 = "FSOR section heading level 3"

# %%
# Used for prefixing and suffixing comment tags. String is reversed for suffix.
# Shouldn't need to be changed. Must be in [a-z] or [A-Z].
TAG_PREFIX = "zyx"
TAG_SUFFIX = TAG_PREFIX[::-1]

# %%
def get_comment_index_tags(
    df_worksheet: pd.DataFrame, 
    comment_tag_column: str
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
    replacement = TAG_PREFIX + r"\1" + TAG_SUFFIX
    comment_tags = df_worksheet[comment_tag_column].str.replace(regex,replacement,regex=True)
    comment_tags.name = "CommentTags"
    return comment_tags

# %%
def append_comment_tags(
    comment_column_list: list,
    comment_tags: pd.Series
) -> list:
    """Appends tags to the end of each comment.

    Args:
        comment_column_list (list): Untagged comment list.
        comment_tags (pd.Series): Tags to append.

    Returns:
        list: Tagged comment list.
    """
    for cmt, tag in zip(comment_column_list,comment_tags):
        for para in cmt:
            if para == cmt[-1]:
                tag_run = (""," (" + tag + ")")
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
    df_sheet = df_worksheet[[LEVEL_1,LEVEL_2,LEVEL_3,
                             BATCH_STATUS_COLUMN,RESPONSE_STATUS_COLUMN,RESPONSE_NOTES_COLUMN
                             ]].copy()
    df_sheet[COMMENT_COLUMN] = comment_col
    df_sheet[RESPONSE_COLUMN] = response_col
    df_sheet = df_sheet.sort_values(
        by=[LEVEL_1,LEVEL_2,LEVEL_3], 
        ascending=[True,True,True]
        ).reset_index(drop=True)
    return df_sheet

# %%
def group_comments(df: pd.DataFrame) -> pd.DataFrame:
    """Group comments at lowest-level of hierarchy. Counts
    the number of comments in the group.

    Args:
        df (pd.DataFrame): Relevant dataframe (grouping columns, 
        comment column)

    Returns:
        pd.DataFrame: Grouped at lowest-level hierarchy with 
        comment count.
    """
    df_group = df.groupby([LEVEL_1,LEVEL_2,LEVEL_3],dropna=False)
    comments_grouped = df_group[COMMENT_COLUMN].apply(tuple)
    comment_count = df_group[COMMENT_COLUMN].count().rename("CommentCount")
    comments_with_count = pd.merge(comments_grouped,comment_count,
                                   left_index=True,right_index=True)
    return comments_with_count

# %%
def find_responses(df: pd.DataFrame) -> pd.DataFrame:
    """Groups at lowest-level of hierarchy, and takes first
    entry as response. Counts the number of responses.

    Args:
        df (pd.DataFrame): Relevant dataframe (grouping columns, 
        response column)

    Returns:
        pd.DataFrame: Grouped at lowest-level hierarchy with 
        response count.
    """
    df_group = df.groupby([LEVEL_1,LEVEL_2,LEVEL_3],dropna=False)
    responses = df_group[RESPONSE_COLUMN].first()
    # Comment groups with no response will not be iterable (NoneType). 
    # Replace with empty run: (Response[Para(Run)])
    empty_response = ([("","")],)
    responses = responses.apply(
        lambda x: x if isinstance(x, tuple) else empty_response
        )
    response_count = df_group[RESPONSE_COLUMN].count().rename("ResponseCount")
    responses_with_count = pd.merge(responses,response_count,
                                    left_index=True,right_index=True)
    return responses_with_count

# %%
def find_response_metadata(df: pd.DataFrame) -> pd.DataFrame:
    """Groups at lowest-level of hierarchy, and takes first
    entries of metadata columns for tracking.

    Args:
        df (pd.DataFrame): Relevant dataframe

    Returns:
        pd.DataFrame: Grouped at lowest-level hierarchy
    """
    df_group = df.groupby([LEVEL_1,LEVEL_2,LEVEL_3],dropna=False)
    batch_status = df_group[BATCH_STATUS_COLUMN].first().fillna("No batch status")
    response_status = df_group[RESPONSE_STATUS_COLUMN].first().fillna("No response status")
    response_notes = df_group[RESPONSE_NOTES_COLUMN].first().fillna("No response notes")
    metadata = r"{{{ " + batch_status + " - " + response_status + " - " + response_notes + r" }}}"
    metadata.name = "Metadata"
    return metadata

# %%
def check_response_count(responses_with_count: pd.DataFrame) -> None:
    """Raises message regarding number of responses. If number
    of responses != 1, show error message.

    Args:
        responses_with_count (pd.DataFrame): Grouped with 
        response count.
    """
    count = responses_with_count["ResponseCount"]
    if count.max() > 1:
        print("ERROR: More than one response for at least one comment group detected. "
              + "Keeping only the first response (which may not be desired).")
    if count.min() < 1:
        print("WARNING: No response for at least one comment group detected. "
              + "Empty response inserted.")

# %%
def combine_and_sort_comments_and_responses(
    comments_with_count: pd.DataFrame,
    responses_with_count: pd.DataFrame,
    metadata: pd.DataFrame
) -> pd.DataFrame:
    """Merges and sorts comments and responses for
    grouping.

    Args:
        comments_with_count (pd.DataFrame): Grouped with count.
        responses_with_count (pd.DataFrame): Grouped with count.

    Returns:
        pd.DataFrame: Combined dataframe, sorted alphabetically
        except comments are sorted by descending.
    """
    # //TODO Merge this function with group_by_level. Put in level_3_group function
    section_grouping = pd.merge(comments_with_count,responses_with_count,
                                left_index=True,right_index=True
                                )
    section_grouping = pd.merge(section_grouping,metadata,
                                left_index=True,right_index=True
                                ).reset_index()
    section_grouping = section_grouping.sort_values(
        by=[LEVEL_1,LEVEL_2,"CommentCount",LEVEL_3], 
        ascending=[True,True,False,True], 
        na_position="first"
        ).reset_index(drop=True)
    return section_grouping

# %%
def group_by_level(df: pd.DataFrame) -> pd.DataFrame:
    LEVEL_3_DATA = "Level3Data"
    LEVEL_2_DATA = "Level2Data"
    LEVEL_1_DATA = "Level1Data"   
    
    def level3_group(df: pd.DataFrame) -> pd.DataFrame:
        # Groups the lowest-level heading (e.g., Heading 3)
        # Comments and responses at this level are already grouped and merged.
        # Provides data combination for further grouping steps.
        # //TODO Merge this with combine_and_sort_comments_and_responses function.
        df[LEVEL_3] = df[LEVEL_3].fillna("Blank")
        df[LEVEL_3_DATA] = tuple(zip(
            df[COMMENT_COLUMN],df[LEVEL_3],df[RESPONSE_COLUMN],df["Metadata"]
            ))
        return df
    
    def level2_group(df: pd.DataFrame) -> pd.DataFrame:
        df_group = df.groupby([LEVEL_1,LEVEL_2])
        comments_level_2 = df_group[LEVEL_3_DATA].apply(tuple)
        comment_count = df_group["CommentCount"].first()
        df_comments_level_2 = pd.merge(comments_level_2,comment_count,
                                       left_index=True,right_index=True
                                       ).reset_index()
        df_comments_level_2 = df_comments_level_2.sort_values(
            by=[LEVEL_1,"CommentCount",LEVEL_2], 
            ascending=[True,False,True],
            ).reset_index(drop=True)
        df_comments_level_2[LEVEL_2_DATA] = tuple(zip(
            df_comments_level_2[LEVEL_3_DATA],
            df_comments_level_2[LEVEL_2]
            ))
        df_comments_level_2 = df_comments_level_2.drop(
            [LEVEL_2,LEVEL_3_DATA], axis=1)
        return df_comments_level_2
    
    def level1_group(df: pd.DataFrame) -> pd.DataFrame:
        df_group = df.groupby([LEVEL_1])
        comments_level_1 = df_group[LEVEL_2_DATA].apply(tuple)
        df_comments_level_1 = pd.DataFrame(comments_level_1).reset_index()
        df_comments_level_1[LEVEL_1_DATA] = tuple(zip(
            df_comments_level_1[LEVEL_2_DATA],
            df_comments_level_1[LEVEL_1]
            ))
        df_comments_level_1 = df_comments_level_1[LEVEL_1_DATA]
        return df_comments_level_1
    
    df = level3_group(df)
    df = level2_group(df)
    df = level1_group(df)
    return df

# %%
def mark_index_entries(comment_tags: list) -> None:
    """Mark index entries by creating AutoMark document
    and opening Word and marking entries, and adding
    index.

    Args:
        comment_tags (list): Comment tags to index.
    """
    regex = f"^{TAG_PREFIX}((\d+?)-(.+?)){TAG_SUFFIX}$"
    index_entry = comment_tags.replace(regex,r"\1",regex=True)
    automark_list = list(zip(comment_tags,index_entry))
    docx_tools.automarkdoc(automark_list)
    # win32_tools requires Office to be installed.
    win32_tools.mark_index_entries(add_index=True)
    return None

# %%
def main():
    # Read ooxml file and retrieve relevant data
    filepath = Path().cwd() / COMMENT_RESPONSE_XLSX_FILENAME
    ooxml_file = ooxml.SpreadSheetML(filepath)
    sheet = ooxml_file.sheet(COMMENT_RESPONSE_SHEET_NUM)
    coded_sheet = sheet.to_dataframe_codes()
    # Remove empty comment rows. All rows should have a comment associated with it.
    remove_empty_comment_rows = coded_sheet[COMMENT_COLUMN].notna()
    coded_sheet = coded_sheet[remove_empty_comment_rows]
    sharedstrings_rich = sheet.get_rich_strings()
    df_worksheet = sheet.to_dataframe()
    df_worksheet = df_worksheet[remove_empty_comment_rows]
    # Get coded comment and response columns
    comment_codes = coded_sheet[COMMENT_COLUMN]
    response_codes = coded_sheet[RESPONSE_COLUMN]
    # Get comment index tags
    comment_tags = get_comment_index_tags(df_worksheet,COMMENT_TAG_COLUMN)
    # Decode comment and response columns
    comment_code_data = ooxml.RichText(sharedstrings_rich,comment_codes)
    response_code_data = ooxml.RichText(sharedstrings_rich,response_codes)
    formats = list(set(comment_code_data.formats_used()
                       + response_code_data.formats_used()))
    comment_column_list = comment_code_data.decode()
    response_column_list = response_code_data.decode()
    # Append comment tags to comments
    comment_column_list = append_comment_tags(comment_column_list,comment_tags)
    # Create working dataframe for next steps of grouping
    df_working = working_df(df_worksheet,comment_column_list,response_column_list)
    # Initial group of comments and responses
    comments_with_count = group_comments(df_working)
    responses_with_count = find_responses(df_working)
    metadata = find_response_metadata(df_working)
    # Error check for number of responses
    check_response_count(responses_with_count)
    # Group headings, comments, and responses into multi-level list
    section_grouping = combine_and_sort_comments_and_responses(comments_with_count,responses_with_count,metadata)
    section_grouping = group_by_level(section_grouping)
    grouped_comment_and_response_list = tuple(section_grouping)
    # Create comment response document and mark index entries
    docx_tools.commentsectiondoc(grouped_comment_and_response_list,formats,levels=3)
    mark_index_entries(comment_tags)
    return None

# %%
if __name__ == "__main__":
    main()
    pass

# %%



