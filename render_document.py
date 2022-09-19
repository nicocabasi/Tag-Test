"""
render_document is a command line tool that can be used to render
a Jinja template and the accompanying json file

Usage:
    python render_document.py <json_input_file> <docx_template_file> <docx_output_file>

Args:
    <json_input_file> is the path to a json file in Airbrush format
    <docx_template_file> is the path to the docx Jinja template
    <docx_output_file> is the path to the rendered doc to create
"""
import argparse
from pathlib import Path
import json
from docxtpl import DocxTemplate

BASE_DIR = Path(__file__).parent

ARGNAME_INPUT_JSON_PATH = "json_input_file"
ARGNAME_OUTPUT_DOCX_PATH = "docx_output_file"
ARGNAME_DOCX_JINJA_TEMPLATE_PATH = "docx_template_file"


def validate_file_type(parser, file_path, ext, must_exist=True):
    """
    Validates file path string passed to an argparse parser. Checks that string has a valid
    extension and is a file that exists if must_exist = True. Raises validation errors
    if any of the conditions fail.

    Args:
        parser: An argparse.ArgumentParser instance used to raise validation errors
        file_path: The file path string to be validated
        ext: The file extenion to validate the file path against
        must_exist: Boolean for whether to check existence of file in validation. Defaults to True

    Returns:
        str | None: The string passed to it if its valid, otherwise None
    """
    path = Path(file_path)
    if must_exist and not (path.exists() and path.is_file):
        parser.error(f"The file {file_path} does not exist")
    elif path.suffix != f".{ext}":
        parser.error(f"The file {file_path} has the wrong extension. Expected {ext}")
    return file_path


def create_arg_parser():
    """
    Creates an argument parser to be used when the file is run directly as a command line tool.

    Returns:
        argparse.ArgumentParser: A parser to be used to parse arguments from sys.argv
    """
    DESCRIPTION = (
        "A command line tool to render a jinja docx template using "
        "the Origin tags and system and an input "
        "json file"
    )
    parser = argparse.ArgumentParser(description=DESCRIPTION)
    parser.add_argument(
        ARGNAME_INPUT_JSON_PATH,
        type=lambda f: validate_file_type(parser, f, "json"),
        help="The path to the json input file",
    )
    parser.add_argument(
        ARGNAME_DOCX_JINJA_TEMPLATE_PATH,
        type=lambda f: validate_file_type(parser, f, "docx", must_exist=False),
        help="The path to the jinja docx template",
    )
    parser.add_argument(
        ARGNAME_OUTPUT_DOCX_PATH,
        type=lambda f: validate_file_type(parser, f, "docx", must_exist=False),
        help="The path to create the docx doc at",
    )
    return parser


def render_document(input_file_path, docx_template, output_file_path):
    """
    Create a doc in docx format using the Origin template and a json input file.

    Args:
        input_file_path: A json input file path to render the doc with
        output_file_path: A docx output file path to write to on success
    """
    with open(input_file_path, "r") as f:
        doc_fields = json.load(f)

    doc_template = DocxTemplate(docx_template)
    doc_template.render(doc_fields)
    doc_template.save(output_file_path)


if __name__ == "__main__":
    parser = create_arg_parser()
    args = vars(parser.parse_args())

    doc_json_file = args[ARGNAME_INPUT_JSON_PATH]
    docx_template_path = args[ARGNAME_DOCX_JINJA_TEMPLATE_PATH]
    doc_output_file = args[ARGNAME_OUTPUT_DOCX_PATH]

    render_document(doc_json_file, docx_template_path, doc_output_file)
    print(f"Successfully created new document: {doc_output_file}")
