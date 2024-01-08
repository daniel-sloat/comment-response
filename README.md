# Comment-Response Script

Creates a formatted docx file using comment and response data from an xlsx spreadsheet. Groups comments for response and organizes according to headings. Tags each comment for tracking and identification.

## Examples
Example input can be found in the 'tests' folder. Run the configuration file as-is (after renaming it to config.toml) to run the script on the files in the tests folder to obtain an example output.

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

- Use Python 3.11+
- Clone the repository or download the repository as a zip file and extract.
- Setup virtual environment:

        python -m venv env
        env\Scripts\activate
        pip install -r requirements.txt

- Customize configuration file (config.toml) in text-editer:
  - ***IMPORTANT***: Rename config.SAMPLE.toml to config.toml!

- Run the script in the terminal:  

        python main.py

- The output file will be placed in the output folder, unless customized in the configuration file.

## Similar Repositories

- [daniel-sloat\comment-extract](https://github.com/daniel-sloat/comment-extract)
  - Produces an xlsx spreadsheet for grouping and response, for input into the comment-response script.

## Further Reading

### DOCX/OOXML

- https://rsmith.home.xs4all.nl/howto/reading-xlsx-files-with-python.html
- https://virantha.com/2013/08/16/reading-and-writing-microsoft-word-docx-files-with-python/
- https://blog.adimian.com/2018/09/04/fast-xlsx-parsing-with-python/
- https://www.toptal.com/xml/an-informal-introduction-to-docx
- https://github.com/booktype/python-ooxml/blob/master/ooxml/parse.py

### XPath

- https://stackoverflow.com/questions/40410269/xpath-to-select-all-nodes-between-two-text-markers-in-ooxml
- http://plasmasturm.org/log/xpath101/
- http://zvon.org/comp/r/tut-XPath_1.html
- https://stackoverflow.com/questions/8181856/xpath-between-two-elements