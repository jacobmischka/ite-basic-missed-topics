#!/usr/bin/env python3

from argparse import ArgumentParser

from .basic_section import BasicSection
from .basic_excel import dump_section_xlsx

import sys


def extract(inpath):
    rows = []
    row_in_progress = []

    with open(inpath, "r") as infile:
        for line in infile:
            if should_skip(line):
                continue

            elif is_data_line(line):
                try:
                    if row_in_progress:
                        rows.append([" ".join(row_in_progress)])
                        row_in_progress = []
                    rows.append(extract_data_row(line))
                except AttributeError as e:
                    print(
                        "Could not append row, skipping: {}".format(e), file=sys.stderr
                    )
            else:
                # Probably a new section
                row_in_progress.append(line.strip())

    return rows


def extract_data_row(line):
    return line.strip().rsplit(maxsplit=4)


def should_skip(line):
    return (
        not line
        or len(line.strip()) == 0
        or "Page" in line
        or "#" in line
        or "Examinees" in line
        or "Your Program" in line
    )


def is_data_line(line):
    return "%" in line


def extract_sections(rows, percentages_last=False):
    sections = []
    heading = None
    subheading = None
    items = []

    for row in rows:
        if len(row) == 1:
            if items:
                try:
                    sections.append(
                        BasicSection(
                            heading,
                            subheading,
                            items,
                            percentages_last=percentages_last,
                        )
                    )
                    heading = None
                    subheading = None
                    items = []
                except Exception as e:
                    print(e, file=sys.stderr)

            if not heading:
                heading, subheading = row[0].strip().split("(")
                heading = heading.strip()
                subheading = subheading[:-1].strip()

        else:
            items.append(row)

    try:
        sections.append(
            BasicSection(heading, subheading, items, percentages_last=percentages_last)
        )
    except Exception as e:
        print(e, file=sys.stderr)

    return sections


def main():
    parser = ArgumentParser(
        description="Create spreadsheet summary of ITE missed topics report"
    )
    parser.add_argument(
        "inpath", help="Input txt file (convert from pdf using `pdftotext -raw`)"
    )
    parser.add_argument("outpath", help="Output file path")
    parser.add_argument(
        "-f",
        "--format",
        dest="format",
        default="xlsx",
        choices=["xlsx"],
        help="Output format (default xlsx)",
    )

    args = parser.parse_args()

    body = extract(args.inpath)
    sections = extract_sections(body)

    if args.format == "xlsx":
        dump_section_xlsx(sections, args.outpath)
    # elif args.format == 'csv':
    # 	dump_section_csv(sections, args.outpath)


if __name__ == "__main__":
    main()
