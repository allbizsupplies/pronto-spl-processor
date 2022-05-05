
import os
import sys

from spl_proc import readers
from spl_proc.writer import export_supplier_pricelist


def main():
    input_filepath = "SPL.xlsx"
    output_filepath = "supplier_pricelist.csv"

    try:
        reader_class_name = sys.argv[1]
    except IndexError:
        print_help()
        quit()

    try:
        reader_class = get_reader(reader_class_name)
    except AttributeError:
        print(f"Reader does not exist: {reader_class_name}")
        quit()

    reader = reader_class(input_filepath)
    spl_items = reader.load_spl_items_from_worksheet()
    export_supplier_pricelist(output_filepath, spl_items)


def get_reader(reader_class: str):
    return getattr(readers, reader_class)


def print_help():
    print("Usage: python -m spl_proc <reader_class>")


if __name__ == "__main__":
    main()
