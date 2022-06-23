# Comment-Response Script

Creates a formatted docx file using comment and response data from an xlsx spreadsheet. Groups comments for response and organizes according to headings. Tags each comment for tracking and identification.

## Examples
Examples can be found...

## Input
An xlsx spreadsheet, with:
- One comment per row
- One response per comment *group*. Additional responses will be carried over into the final document.
- Tags to append to the end of each comment when writing document.
- Headings defined (Excel has a maximum of 9 headings). The hierarchy of headings can contain up to 9 headings (if beginning at heading number 1). *(Two or three headings is more than enough!)*
  - For grouping purposes, heading names are case-sensitive.
- Custom sorts can be applied by either specifying in the configuration file, or by specifying columns to be used in sorting.
  - The configuration file allows specific sorting for the highest-level headings, and then a "general" sort for all sub-headings. This sub-heading level sort in the configuration file is good for specifying a sub-heading that should always go first - such as a "General" sub-heading, as any values that are not filled in are put at the end of the sort group. A specific sort for all sub-headings can be achieved using this sub-heading sort, but it may be easier to specify additional columns to sort by in the Excel spreadsheet.
  - For a very specific customized sort, use the same number of columns as there are for headings, and use numbers (or letters) to specify the sort. Any values that are not filled in are put at the end of the sort group.

## How to Use

Copy from other...