# Comment-Response Script

Creates a formatted docx file using comment and response data from an xlsx spreadsheet. Groups comments for response and organizes according to headings. Tags each comment for tracking and identification.

## Examples
Example input can be found in the 'tests' folder. Example output can be found in the 'output' folder.

## Input
An xlsx spreadsheet, with:
- One comment per row
- One response per comment *group*. Additional responses will be carried over into the final document.
- A column that identies the specific document the comment came from (or a special unique id that identifies that specific comment). Each of these comment ids are to appended to the end of each comment when writing document.
- Headings defined (Excel has a maximum of 9 headings). The hierarchy of headings can contain up to 9 headings (if beginning at heading number 1). *(Two or three headings is more than enough!)*
  - For grouping purposes, heading names are case-sensitive.
- Custom sorts can be applied by either specifying in the configuration file, or by specifying columns to be used in sorting.
  - The configuration file allows specific sorting for the highest-level headings, and then a "general" sort for all sub-headings. This sub-heading level sort in the configuration file is good for specifying a sub-heading that should always go first - such as a "General" sub-heading, as any values that are not filled in are put at the end of the sort group. A specific sort for all sub-headings can be achieved using this sub-heading sort, but it may be easier to specify additional columns to sort by in the Excel spreadsheet.
  - For a very specific customized sort, use the same number of columns as there are for headings, and use numbers (or letters) to specify the sort. Any values that are not filled in are put at the end of the sort group.

## Output
A single, formatted comment-response docx file grouped and sorted as specified.

## How to Use

- Use Python 3.10+
- Clone the repository or download the repository as a zip file and extract.
- Setup virtual environment (Windows-specific):
  - Using terminal (e.g., Command Prompt, PowerShell):
    - Navigate to cloned/extracted folder.
    - Enter:

          python3.10 -m venv env
          env\Scripts\activate
          pip install -r requirements.txt
          
- Customize configuration file (config.toml) in text-editer:
  - ***IMPORTANT***: Rename config.SAMPLE.toml to config.toml!

- Run the script in the terminal:  

        python comment_response.py

- The output file will be placed in the output folder, unless customized in the configuration file.

## Similar Repositories

- [daniel-sloat\comment-extract](https://github.com/daniel-sloat/comment-extract)
  - Produces an xlsx spreadsheet for grouping and response, for input into the comment-response script.