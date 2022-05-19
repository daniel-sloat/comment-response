# Load config file and parameters
FILENAME = "Comment-Response.xlsx"
SHEETNAME = "sheet1"
COMMENTS_COLUMN = "Comments"
RESPONSE_COLUMN = "Response"
COMMENT_TAGS_COLUMN = "Tags"
LEVEL_1 = "Heading 1"
LEVEL_2 = "Heading 2"
LEVEL_3 = "Heading 3"

# Specify level 1, 2, and 3 sorting.
# Alphabetical (ascending): 'alpha', Comment Count (descending): 'count', No sort: anything else
SORT = "count"

# Specify custom sort of level 1. Use empty list [] for no custom sort.
# For headings not listed here, it will be at the end and sorted as specified above.
LEVEL_1_SORT = [
    "Comments Group 1",
    "Comments Group 2",
    "Comments Group 3"
]