# Convert multilingual workbook to multiple TMX files 
<!--- [task 20.3000] -->

This script converts a multilingual workbook into as many TMX files as target languages (or language pairs) it contains.

* Input: multilingual spreadsheet (Excel file)
* Output: one TMX file per language pair

## Execution 

To know how this utility must be run:
```
python3.11 conv_xls2tmx.py --help
```

## Requirements 

The first worksheet in the workbook must be called `config` and include the following options, to be updated by the PM.


| KEY                    | VALUE                            | DESCRIPTION |
|:--------------------	|:-------------------------------	|:-------------	|
| container           	| `<container>`                   	| String: Add the the name of the actual container (case-sensitive). It needs to exist in the containers manager.                                                                                                                                                                                                                      |
| source_lang         	| eng-ZZZ                        	| String: Include if not English.                                                                                                                                                                                           |
| tmx_file_names         | `<container>`, `<target_lang>`, `TM` | List of strings: List all elements that should be included in the name of the TMX   files. Container must be the first one and they must appear in order   (separated by commas). Placeholders (e.g. `<this>`) refer to keys in   this config sheet (first column). All elements in this list will be joined   with underscore, e.g. `<container>_<xxx-XXX>_TM.tmx`. If you want   to include word "QQ" between container and language, this cell   should contain [`<container>, QQ, <target_lang>, TM`] (without   the square brackets), and that will produce   `<container>_QQ_<xxx-XXX>_TM.tmx`.                                                                                                                                                                                                   |
| segmentation           | yes                              | Boolean: Contents of cells will be segmented if possible (if the same number of   sentences, line breaks, etc. Is found on both languages)                                                                                                                                                                                                    |
| remove_html_tags       | yes                              | Boolean: Parts of the text matched by `<[^>]+>` will be removed.                                                                                                                                                   |
| remove_linebreaks      | no                               | Boolean: Linebreaks (i.e. `[^\r\n]+`) will be removed withing translation units.                                                                                                                                  |
| remove_multiple_spaces | no                               | Boolean: Double or multiple normal spaces will be replaced by one single space.                                                                                                                                   |
| remove_pattern         |                                  | Regex: Parts of the text matched by this expression will be removed.So for example   if you wanted to have parts like “[[Privacy_Policy.ENG.pdf\|προστασία της   ιδιωτικής ζωής]]” or "((*\|Meetings include online meetings))"   removed, the remove pattern should be something like `\[\[[^\]]+\]\]` or   `\(\([^)]+\)\)`, respectively. You (or the PM) don’t have to write that, you   can just provide examples and an explanation.                                                                                                                                                             |
| ofuscate_pattern       |                                  | Regex: Parts of the text matched by this expression will be ofuscated, e.g. Xxxxxx                                                                                                                                |
| ignore_cell_pattern    |                                  | Regex: Cells matched by this expression will not be included in the TMX file.                                                                                                                                     |
| header_row             | 0                                | Integer: Indicate the row containing the language codes. Starts with index 0 (row 1).                             |
| comment_column         |                                  | Letter or string: Indicate whether any column contains a comment or description that   should be included in every translation unit. Add letter name of the column   or exact text content of the cell at the `header_row`.                                                                                                                   |

<!-- Workbook template: [multilingual_tmwb_template.xlsx](multilingual_tmwb_template.xlsx) -->

Workbook template: [data/multilingual_tmwb_template.xlsx](data/multilingual_tmwb_template.xlsx)



## TODO list:
* remove markup (TBC by the PM!)
* segment the cells whenever possible
* clean up tags
* create API
* config.json (instead of config sheet)
* accept boolean (true/false) in config



## WIP

- done:
    1. parse excel, extract each language pair (cols 1.3)
    2. put each language pair in one tmx

- todo:
    1. segment paragraphs whenever possible
    2. clean up tags
