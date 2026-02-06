#!/usr/bin/env python3

#  This file is part of cApps.
#
#  This script is free software: you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation, either version 3 of the License, or
#  (at your option) any later version.
#
#  This script is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#
#  You should have received a copy of the GNU General Public License
#  along with cApps.  If not, see <https://www.gnu.org/licenses/>.

# ############# AUTHORSHIP INFO ###########################################

__author__ = "Manuel Souto Pico"
__copyright__ = "Copyright 2021, cApps/cApStAn"
__credits__ = ["Manuel Souto Pico"]
__license__ = "GPL"
__version__ = "0.2.0"
__maintainer__ = "Manuel Souto Pico"
__email__ = "manuel.souto@capstan.be"
__status__ = "Testing / pre-production" # "Production"

# ############# IMPORTS ###########################################

import regex as re
from typing import Dict
from pathlib import Path
import os
import sys
import json
import argparse
## import xlrd
from yattag import Doc, indent
import pandas as pd
import numpy as np
from rich import print
from conf.langtags import fetch_langtags_data
from conf.langtags import get_correspondent_tag
from conf.langtags import get_langtags_in_scheme
## import openpyxl
# from pprint import pprint as print
# import xml.dom.minidom
from mod.markup import strip_html


def main():
    # ############# PROGRAM DESCRIPTION ###########################################

    text = "This is TM Workbook Converter: it takes a spreadsheet/workbook where each \
    column contains a language version and produces as many TMX files as target \
    languages the workbook has."

    # intialize arg parser with a description
    parser = argparse.ArgumentParser(description=text)
    parser.add_argument("-V", "--version", help="show program version",
                        action="store_true")
    parser.add_argument("-i", "--input", help="specify path to mandatory input file")
    parser.add_argument("-c", "--config", help="specify path to optional config file")

    # read arguments from the command line
    args = parser.parse_args()

    # check for -V or --version
    if args.version:
        print("This is program TM Workbook Converter version 0.2")
        sys.exit()

    if not args.input:
        exit("Argument -i/--input not found.")
    elif not Path(args.input).is_file():
        exit(f"Input file '{args.input}' not found.")
    else:
        path_to_wb = Path(args.input.strip())
        print(f"Processing file: '{path_to_wb.name}'")

    if not args.config:
        exit("Argument -c/--config not found.")
    # elif "config" not in Path(args.config).name or not args.config.endswith(".json"):
    elif not args.config.endswith(".json"):
        exit(f"Config file '{args.config}' not correct.")
    else:
        config_fpath = Path(args.config.strip())
        print(f"Using configuration from '{config_fpath.name}'")


    # ############# FUNCTIONS #####################################################

    def normalize_values(config: Dict) -> Dict:
        config = {
            k: True if isinstance(v, str) and v.lower() == "yes"
               else False if isinstance(v, str) and v.lower() == "no"
               else v
            for k, v in config.items()
        }
        return config

    def remove_linebreaks(text: str) -> str:
        # return re.sub(r"([^>\n])[\r\n]+", r"\1@", text.strip())
        # return re.sub(r"\p{L}[\r\n]\p{L}", r"@", text.strip())
        return text

    def clean_seg_text(s):
        # collapse all whitespace runs (including hidden newlines)
        return re.sub(r'\s+', ' ', str(s)).strip()

    def get_config(config_fpath, wb):
        if args.config:  # config.json
            with open(config_fpath) as json_file:
                return json.load(json_file)
        elif "config" in wb.sheet_names:
            print("Read configuration from from 'config' sheet in workbook.")
            # only if config.json was not provided as arg
            return read_config_sheet(wb)
        else:
            raise ValueError(
                "Configuration not provided: either specify a 'config.json' file or include a 'config' sheet in the workbook.")

    def get_worksheet(wb, config):
        # if the extraction sheet is not specified,
        if config["worksheet"] is None:
            # and there are only two sheets, then use the one that is not config
            print(f"{wb.sheet_names=}")
            if len(wb.sheet_names) == 1:
                return wb.sheet_names[0]
            elif len(wb.sheet_names) == 2 and "config" in wb.sheet_names:
                return wb.sheet_names[1] if wb.sheet_names[0] == "config" else wb.sheet_names[0]
            # if there are more or just config, then fail
            else:
                print("ERROR: The worksheet to be extracted is not specified in config")
                sys.exit()
        return config["worksheet"]

    def get_langtags():
        # langtags = pd.read_csv(langtags_csv)
        return fetch_langtags_data('https://capps.capstan.be/langtags_json.php')

    def read_config_sheet(wb):
        print(type(wb))
        breakpoint()
        config_sheet = wb.parse("config").replace(np.nan, None)
        ## sheet = wb.sheet_by_index(sheet_idx)
        parameters = config_sheet['KEY']
        values = config_sheet['VALUE']
        return dict(zip(parameters, values))

    def get_data(wb, sheet_name, source_col, target_col):
        # sheet = wb.sheet_by_index(sheet_idx)
        df = wb.parse(sheet_name)
        df = df[df[target_col].notna()]
        return set(zip(df[source_col], df[target_col]))

    def get_headers(wb, sheet_name):
        # COMMENT: enforce first row as headers!
        # sheet_name = wb.sheet_names[sheet_idx]
        df = wb.parse(sheet_name)
        return df.columns

    def build_tmx(langpair_set, xml_source_lang, xml_target_lang, config):
        # convert to tmx
        doc, tag, text = Doc().tagtext()

        doc.asis('<?xml version="1.0" encoding="UTF-8"?>')
        with tag('tmx', version="1.4"):
            with tag('header', creationtool="cApps", creationtoolversion="2020.10",
                     segtype="paragraph", adminlang="en",
                     datatype="HTML", srclang=xml_source_lang):
                doc.attr(
                    ('o-tmf', "omt")  # o_tmf="omt",
                )
                text('')
            with tag('body'):
                for tu in langpair_set:
                    src_txt = str(tu[0]).strip()
                    tgt_txt = str(tu[1]).strip()

                    markup_cleanup = config['remove_html_tags']
                    if markup_cleanup:
                        src_txt = strip_html(src_txt)
                        tgt_txt = strip_html(tgt_txt)

                    with tag('tu'):
                        with tag('tuv'):
                            doc.attr(
                                ('xml:lang', xml_source_lang)
                            )
                            with tag('seg'):
                                text(src_txt)
                        with tag('tuv'):
                            doc.attr(
                                ('xml:lang', xml_target_lang)
                            )
                            with tag('seg'):
                                text(tgt_txt)

        # tmx_output = doc.getvalue()
        tmx_output = indent(
            doc.getvalue(),
            indentation=' ' * 2,
            newline='\r\n',
            indent_text=False,
        )

        return remove_linebreaks(tmx_output)  # .replace("o_tmf=", "o-tmf=")
        # return tmx_output  # .replace("o_tmf=", "o-tmf=")

    def get_lang_headers(columns, config):
        if config["langtag_scheme"] == "cApStAn":
            return [tag for tag in columns
                    if re.match(r'[a-z]{3}-[A-Z]{3}', tag) and tag != config['source_lang']]
        else:
            return [tag for tag in columns
                    if tag in bcp47_langtags and tag != config['source_lang']]

    def write_tmx_file(config, tmx_output):
        # build filename
        config['tmx_file_names'] = config['tmx_file_names'].replace('<', '').replace('>', '')
        fn_parts = [config[x.strip()] if x.strip() in config.keys()
                    else x.strip()
                    for x in config['tmx_file_names'].split(',')]

        # writing output
        filename = "_".join(fn_parts) + ".tmx"
        output_dir = "output"
        output_tmx_fpath = Path.cwd() / output_dir / filename

        os.makedirs(output_dir) if not os.path.exists(output_dir) else None

        with open(output_tmx_fpath, "w", encoding="utf-8") as f:
            print(f"Writing TMX output to file '{output_tmx_fpath}'")
            f.write(tmx_output)

    # all source language variables should be global!: path_to_file, wb, langtags
    def convert_wb_to_tmx_files(config_fpath, path_to_file):

        # wb = xlrd.open_workbook(path_to_file)
        # wb = openpyxl.load_workbook(path_to_file)

        # df = pd.read_excel(path_to_file)

        wb = pd.ExcelFile(path_to_file)

        try:
            config = get_config(config_fpath, wb)
            config = normalize_values(config)
            if config['source_lang'] is None:
                raise ValueError("The 'source_lang' field in the configuration is missing.")
            if config['source_column'] is None:
                raise ValueError("The 'source_column' field in the configuration is missing.")
        except ValueError as e:
            print(f"Error: {e}")
            sys.exit(1)  # exit the script gracefully

        worksheet = get_worksheet(wb, config)
        print(f"{worksheet=}")

        columns = get_headers(wb, worksheet)  # assuming config is 0
        print(f"{columns=}")

        if config['source_column'] not in columns:
            exit("""ERROR: The specified column is not found in the worksheet""")

        source_col = config['source_column']
        print(f"{source_col=}")
        target_columns = get_lang_headers(columns, config)
        print(f"{target_columns=}")

        bcp47_source_langtag = get_correspondent_tag(
            langtags, config['source_lang'], config['langtag_scheme'], "BCP47"
        ) if config['langtag_scheme'] == "cApStAn" else config['source_lang']
        print(f"{bcp47_source_langtag=}")

        # convert_colpair_to_tmx_file() for index, column in cols if column in lang_list
        for index, column in enumerate(columns):
            if column in target_columns:  # this excludes notes etc.
                print("-----------")
                print(f"{index=}: {column=}")
                # configuration of this language pair
                lang_config = dict(config, target_lang=column)  # update dict without modify original dictionary
                bcp47_target_langtag = get_correspondent_tag(
                    langtags, column, config['langtag_scheme'], "BCP47"
                ) if config['langtag_scheme'] == "cApStAn" else column
                print(f"{bcp47_target_langtag=}")

                if bcp47_target_langtag is None:
                    print(f"""ERROR: The target language of {column} is not recognized""")
                    continue

                langpair_set = get_data(wb, worksheet, source_col=source_col, target_col=column)
                tmx_output = build_tmx(langpair_set, bcp47_source_langtag, bcp47_target_langtag, config)
                write_tmx_file(lang_config, tmx_output)

    # ############# EXECUTION #####################################################
    langtags = get_langtags()
    bcp47_langtags = get_langtags_in_scheme(langtags, "BCP47")
    convert_wb_to_tmx_files(config_fpath, path_to_wb)


if __name__ == "__main__":
    main()


# todo:
# use bcp47 convention, or add convention key to config
# if config["langtag_scheme"] is None, then BCP47 should be assumed
# use the langtags api
# add option to use config.json
# use a default config in function read_config_sheet (if neither config sheet nor config json are found)
# add logging
# add other conventions to funcion get_lang_headers
# add if convention is not capstan, it must be BCP47, in that case all headers should be found in the list of BCP47 tags