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

import re
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
from conf.langtags import get_langtags_in_convention
## import openpyxl
# from pprint import pprint as print
# import xml.dom.minidom

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

if not os.path.exists(args.input):
    print(f"Input file '{args.input}' not found.")
    sys.exit()
elif args.input:
    print(f"Processing file '{os.path.basename(args.input)}'.")
    path_to_wb = args.input.rstrip('/')
else:
    print("Argument -i not found.")
    sys.exit()

if args.config and os.path.basename(args.config) == "config.json":
    print(f"Using configuration from '{args.config}'.")
    config_fpath = args.config

# ############# FUNCTIONS #####################################################

def get_config(wb):

    if args.config: # config.json
        with open(config_fpath) as json_file:
            return json.load(json_file)
    elif "config" in wb.sheet_names:
        print("Read configuration from from 'config' sheet in workbook.")
        # only if config.json was not provided as arg
        return read_config_sheet(wb)
    else:
        raise ValueError("Configuration not provided: either specify a 'config.json' file or include a 'config' sheet in the workbook.")


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
    config_sheet = wb.parse("config").replace(np.nan, None)
    ## sheet = wb.sheet_by_index(sheet_idx)
    parameters = config_sheet['KEY']
    values = config_sheet['VALUE']
    return dict(zip(parameters, values))


def get_data(wb, sheet_name, source_col, target_col):
    # sheet = wb.sheet_by_index(sheet_idx)
    df = wb.parse(sheet_name)
    source_texts = df[source_col]
    target_texts = df[target_col]
    return set(zip(source_texts, target_texts))


def get_headers(wb, sheet_name):
    # COMMENT: enforce first row as headers!
    # sheet_name = wb.sheet_names[sheet_idx]
    df = wb.parse(sheet_name)
    return df.columns


def build_tmx(langpair_set, xml_source_lang, xml_target_lang):
    # convert to tmx
    doc, tag, text = Doc().tagtext()

    doc.asis('<?xml version="1.0" encoding="UTF-8"?>')
    with tag('tmx', version="1.4"):
        with tag('header', creationtool="cApps", creationtoolversion="2020.10",
                 segtype="paragraph", adminlang="en",
                 datatype="HTML", srclang=xml_source_lang):
            doc.attr(
                ('o-tmf', "omt") # o_tmf="omt",
            )
            text('')
        with tag('body'):
            for tu in langpair_set:
                src_txt = str(tu[0]).strip()
                tgt_txt = str(tu[1]).strip()
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

    tmx_output = indent(
        doc.getvalue(),
        indentation=' '*2,
        newline='\r\n'
    )
    return tmx_output  # .replace("o_tmf=", "o-tmf=")


def get_lang_headers(columns, config):
    if config["langtag_convention"] == "cApStAn":
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
    print("Writing TMX output to file " + filename)

    output_dir = "output"
    os.makedirs(output_dir) if not os.path.exists(output_dir) else None

    with open(os.path.join(output_dir, filename), "w") as f:
        print(tmx_output, file=f)


# all source language variables should be global!: path_to_file, wb, langtags
def convert_wb_to_tmx_files(path_to_file):

    # wb = xlrd.open_workbook(path_to_file)
    # wb = openpyxl.load_workbook(path_to_file)

    # df = pd.read_excel(path_to_file)

    wb = pd.ExcelFile(path_to_file)

    try:
        config = get_config(wb)
        if config['source_lang'] is None:
            raise ValueError("The 'source_lang' field in the configuration is missing.")
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)  # exit the script gracefully

    worksheet = get_worksheet(wb, config)
    print(f"{worksheet=}")

    columns = get_headers(wb, worksheet) # assuming config is 0
    print(f"{columns=}")

    if config['source_lang'] not in columns:
        print("""ERROR: The specified source language is not found in the
            columns headers of the worksheet to be extracted""")
        return

    source_col = config['source_lang']
    print(f"{source_col=}")
    lang_list = get_lang_headers(columns, config)
    print(f"{lang_list=}")

    bcp47_source_langtag = get_correspondent_tag(
        langtags, config['source_lang'], config['langtag_convention'], "BCP47"
    ) if config['langtag_convention'] == "cApStAn" else config['source_lang']
    print(f"{bcp47_source_langtag=}")

    # convert_colpair_to_tmx_file() for index, column in cols if column in lang_list
    for index, column in enumerate(columns):
        if column in lang_list: # this excludes notes etc.
            print("-----------")
            print(f"{index=}: {column=}")
            # configuration of this language pair
            lang_config = dict(config, target_lang=column)  # update dict without modify original dictionary
            bcp47_target_langtag = get_correspondent_tag(
                langtags, column, config['langtag_convention'], "BCP47"
            ) if config['langtag_convention'] == "cApStAn" else column
            print(f"{bcp47_target_langtag=}")

            if bcp47_target_langtag is None:
                print("""ERROR: The target language {bcp47_target_lang} is not recognized""")
                continue

            langpair_set = get_data(wb, worksheet, source_col=source_col, target_col=column)
            tmx_output = build_tmx(langpair_set, bcp47_source_langtag, bcp47_target_langtag)
            write_tmx_file(lang_config, tmx_output)



# ############# EXECUTION #####################################################

if __name__ == "__main__":
    langtags = get_langtags()
    bcp47_langtags = get_langtags_in_convention(langtags, "BCP47")
    convert_wb_to_tmx_files(path_to_wb)


# todo:
# use bcp47 convention, or add convention key to config
# if config["langtag_convention"] is None, then BCP47 should be assumed
# use the langtags api
# add option to use config.json
# use a default config in function read_config_sheet (if neither config sheet nor config json are found)
# add logging
# add other conventions to funcion get_lang_headers
# add if convention is not capstan, it must be BCP47, in that case all headers should be found in the list of BCP47 tags