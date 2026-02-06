import tempfile
import subprocess
import json
from rich import print
from pathlib import Path
import argparse
import pandas as pd

from mod.conf import get_config

text = "This script processes a large number of Excel files in bulk"

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
    exit("This is program TM Workbook Converter version 0.2")

if not args.input:
    exit("Argument -i/--input not found.")
elif not Path(args.input).is_file():
    exit(f"Input file '{args.input}' not found.")
else:
    file_fpath = Path(args.input.strip())
    print(f"Processing file: '{file_fpath.name}'")

if not args.config:
    exit("Argument -c/--config not found.")
elif "config" not in Path(args.config).name or not args.config.endswith(".json"):
    exit(f"Config file '{args.config}' not correct.")
else:
    config_fpath = Path(args.config.strip())
    print(f"Config file: '{config_fpath.name}'")


df = pd.read_excel(file_fpath, engine="odf")
headers = df.columns.tolist()

try:
    expected_headers = ["sheet", "source", "tag scheme", "path", "stem"]
    assert all(h in headers for h in expected_headers)
except AssertionError as e:
    exit("Some expected headers are missing.")


config = get_config(config_fpath)

# configs = df[["stem", "sheet", "source", "tag scheme", "path"]].rename(
#     columns={"stem": "container", "tag scheme": "tag_scheme"}
# ).to_dict(orient="records")
#
batch_conf = df.to_dict(orient="records")

for file in batch_conf:
    excel_fpath = Path(file["path"])
    config["worksheet"] = file["sheet"]
    config["source_lang"] = file["source"]
    config["container"] = file["stem"]

    print(config)
    with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=True) as tmp:
        print(f"{tmp.name}")
        json.dump(config, tmp, indent=2)
        tmp.flush()


        # Call external script, pass temp file
        # Example: subprocess.run(["python", "other_script.py", tmp.name])
        # print(f"Calling script with config {tmp.name}")

        # Call the other script
        result = subprocess.run(
            ["python", "conv_xls2tmx.py", "-i", excel_fpath, "-c", tmp.name],
            capture_output=True,  # captures stdout and stderr
            text=True             # returns output as strings instead of bytes
        )

        print(f"Return code: {result.returncode}")
        print(f"Stdout:\n{result.stdout}") if result.stdout else None
        print(f"Stderr:\n {result.stderr}") if result.stderr else None
        print("----------------------------")