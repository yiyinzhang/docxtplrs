"""Command line interface: render a docx template with json data.

Usage: python -m docxtplrs [-h] [-o] [-q] template_path json_path output_filename
(mirrors python -m docxtpl)
"""

import argparse
import json
import os
import sys

from docxtplrs import DocxTemplate, TemplateError

TEMPLATE_ARG = "template_path"
JSON_ARG = "json_path"
OUTPUT_ARG = "output_filename"
OVERWRITE_ARG = "overwrite"
QUIET_ARG = "quiet"


def make_arg_parser():
    parser = argparse.ArgumentParser(
        usage="python -m docxtplrs [-h] [-o] [-q] {} {} {}".format(
            TEMPLATE_ARG, JSON_ARG, OUTPUT_ARG
        ),
        description="Make docx file from existing template docx and json data.",
    )
    parser.add_argument(TEMPLATE_ARG, type=str, help="The path to the template docx file.")
    parser.add_argument(JSON_ARG, type=str, help="The path to the json file with the data.")
    parser.add_argument(OUTPUT_ARG, type=str, help="The filename to save the generated docx.")
    parser.add_argument(
        "-o",
        "--" + OVERWRITE_ARG,
        action="store_true",
        help="If output file already exists, overwrites without asking for confirmation",
    )
    parser.add_argument(
        "-q",
        "--" + QUIET_ARG,
        action="store_true",
        help="Do not display unnecessary messages",
    )
    return parser


def check_exists_ask_overwrite(arg_value, overwrite):
    if os.path.exists(arg_value) and not overwrite:
        msg = (
            "File %s already exists, would you like to overwrite the existing file? "
            "(y/n)" % arg_value
        )
        if input(msg).lower() == "y":
            return True
        raise RuntimeError(
            "File %s already exists, please choose a different name." % arg_value
        )
    return True


def main():
    parser = make_arg_parser()
    args = vars(parser.parse_args())

    template_path = args[TEMPLATE_ARG]
    json_path = args[JSON_ARG]
    output_filename = args[OUTPUT_ARG]
    overwrite = args[OVERWRITE_ARG]
    quiet = args[QUIET_ARG]

    if not (os.path.isfile(template_path) and template_path.endswith(".docx")):
        parser.error("The template file must be an existing .docx file")
    if not (os.path.isfile(json_path) and json_path.endswith(".json")):
        parser.error("The json file must be an existing .json file")
    if not output_filename.endswith(".docx"):
        parser.error("The output file must be a .docx file")
    check_exists_ask_overwrite(output_filename, overwrite)

    with open(json_path, encoding="utf-8") as f:
        context = json.load(f)

    tpl = DocxTemplate(template_path)
    try:
        tpl.render(context)
    except TemplateError as exc:
        if not quiet:
            print("Template error:", exc, file=sys.stderr)
        raise SystemExit(1)
    tpl.save(output_filename)
    if not quiet:
        print("Saved:", output_filename)


if __name__ == "__main__":
    main()
