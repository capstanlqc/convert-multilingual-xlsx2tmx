# Convert multilingual workbook to multiple TMX files 
<!--- [task 20.3000] -->

This script converts a bilingual or multilingual workbook into as many TMX files as target languages (or language pairs) it contains.

* Input: multilingual spreadsheet (Excel file)
* Output: one TMX file per language pair


## Changes


| Date      | Change                            | Responsible |
|:--------- |:-------------------------------	|:-------------	|
| 20240710  | Refactored the original scritp using pandas consistently | Manuel |
| 20240711  | Added argument `--config` (or `-c`) to use `config.json` (if provided, it overrides `config` worksheet) | Manuel |

## Getting started

Clone [this repo](https://github.com/capstanlqc/convert-multilingual-xlsx2tmx) and install dependencies in a virtual environment:

```
gh repo clone capstanlqc/convert-multilingual-xlsx2tmx
cd convert-multilingual-xlsx2tmx
python3.11 -m venv venv
source venv/bin/activate
python3.11 -m pip install -r requirements.txt
```
To know how to run this utility:
```
python3.11 conv_xls2tmx.py --help
```

## Configuration 

The script expects some configuration in order to know what data it must process or how to process it. Configuration may be provided in the first worksheet of the workbook (named `config`), or as a separate `config.json` file. 

The following `config` values are expected:


| KEY                    | VALUE                            | DESCRIPTION |
|:--------------------	|:-------------------------------	|:-------------	|
| container           	| `<container>`                   	| String: Add the the name of the actual container (case-sensitive). It needs to exist in the containers manager.   |
| langtag_convention   	| cApStAn                        	| String. Convention used for language codes, it can be cApStAn or BCP47. If empty, BCP47 will be assumed as default. |
| source_lang         	| eng-ZZZ                        	| String: Mandatory code in the  `<langtag_convention>` specified. |
| tmx_file_names         | `<container>`, `<target_lang>`, `TM` | List of comma-separated strings: List all elements that should be included in the name of the TMX   files. Container must be the first one and they must appear in order   (separated by commas). Placeholders (e.g. `<this>`) refer to keys in   this config sheet (first column). All elements in this list will be joined   with underscore, e.g. `<container>_<xxx-XXX>_TM.tmx`. For example, if you want   to include word "QQ" between container and language, this cell   should contain `<container>, QQ, <target_lang>, TM`, which will produce   `<container>_QQ_<xxx-XXX>_TM.tmx`.    |
| segmentation           | yes                              | Boolean: Contents of cells will be segmented if possible (if the same number of   sentences, line breaks, etc. is found on both languages)        |
| remove_html_tags       | yes                              | Boolean: Parts of the text matched by `<[^>]+>` will be removed.        |
| remove_linebreaks      | no                               | Boolean: Linebreaks (i.e. `[^\r\n]+`) will be removed withing translation units.    |
| remove_multiple_spaces | no                               | Boolean: Double or multiple normal spaces will be replaced by one single space.    |
| remove_pattern         |                                  | Regex: Parts of the text matched by this expression will be removed. For example   if you wanted to have parts like “[[Privacy_Policy.ENG.pdf\|προστασία της   ιδιωτικής ζωής]]” or "((*\|Meetings include online meetings))"   removed, the remove pattern should be something like `\[\[[^\]]+\]\]` or   `\(\([^)]+\)\)`, respectively. You (or the PM) don’t have to write that, you   can just provide examples and an explanation.   |
| ofuscate_pattern       |                                  | Regex: Parts of the text matched by this expression will be ofuscated, e.g. Xxxxxx                                  |
| ignore_cell_pattern    |                                  | Regex: Cells matched by this expression will not be included in the TMX file.    |
| worksheet              | FOO              | String. The name of the worksheet containing the translations. This value is mandatory if there are more than one worksheet, but it can be left empty if there’s only one worksheet (other than this ‘config’). | 
| header_row             | 0                                | Integer: Indicate the row containing the language codes. Starts with index 0 (row 1). **DEPRECATED**: Value 0 (first row) is mandatory.  |
| comment_column         |                                  | Letter or string: Indicate whether any column contains a comment or description that   should be included in every translation unit. Add letter name of the column   or exact text content of the cell at the `<header_row>`.   |

<!-- Workbook template: [multilingual_tmwb_template.xlsx](multilingual_tmwb_template.xlsx) -->

Workbook template: [data/multilingual_tmwb_template.xlsx](data/multilingual_tmwb_template.xlsx)


## TODO:

> A number of options in the configuration have not been implemented, since nobody has requested them: segmentation, remove_html_tags, remove_linebreaks, remove_multiple_spaces, remove_pattern, ofuscate_pattern, ignore_cell_pattern, comment_column

* remove markup (TBC by the PM!)
* segment cells whenever possible
* clean up tags
* create API (input excel file, output zip with tmx files)
* accept boolean (true/false) in config

## References

- [Language tags basic API](https://github.com/capstanlqc/langtags_basic_api)