#!/usr/bin/env python3

from argparse import ArgumentParser

from .basic_section import BasicSection
from .basic_excel import dump_section_xlsx

import sys, re


def extract(inpath, numbers_inline):
    rows = []
    row_in_progress = []

    with open(inpath, "r") as infile:
        for line in infile:
            if should_skip(line):
                continue

            elif is_num_line(line):
                rows.append([line.strip()])

            elif is_data_line(line):
                try:
                    if row_in_progress:
                        rows.append([" ".join(row_in_progress)])
                        row_in_progress = []
                    rows.append(extract_data_row(line, numbers_inline=numbers_inline))
                except AttributeError as e:
                    print(
                        "Could not append row, skipping: {}".format(e), file=sys.stderr
                    )
            else:
                # Probably a new section
                row_in_progress.append(line.strip())

    return rows


def extract_data_row(line, numbers_inline=False):
    maxsplit = 4 if numbers_inline else 2
    return line.strip().rsplit(maxsplit=maxsplit)


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


def is_num_line(line):
    return "N =" in line


def get_nums(line):
    m = re.match(r"N = (\d+) N = (\d+)", line)
    return m.group(1, 2)


def extract_sections(rows, numbers_inline=False, percentages_last=False):
    sections = []
    heading = None
    subheading = None
    total_num = None
    program_num = None
    items = []

    for row in rows:
        if len(row) == 1:
            if is_num_line(row[0]):
                total_num, program_num = get_nums(row[0])
                continue

            if items:
                try:
                    sections.append(
                        BasicSection(
                            heading,
                            subheading,
                            items,
                            total_num=total_num,
                            program_num=program_num,
                            numbers_inline=numbers_inline,
                            percentages_last=percentages_last,
                        )
                    )
                    heading = None
                    subheading = None
                    items = []
                except Exception as e:
                    print("Failed creating section", e, file=sys.stderr)

            if not heading:
                heading, subheading = row[0].strip().split("(")
                heading = heading.strip()
                subheading = subheading[:-1].strip()

        else:
            items.append(row)

    try:
        sections.append(
            BasicSection(
                heading,
                subheading,
                items,
                total_num=total_num,
                program_num=program_num,
                numbers_inline=numbers_inline,
                percentages_last=percentages_last,
            )
        )
    except Exception as e:
        print("Failed creating section", e, file=sys.stderr)

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
    parser.add_argument(
        "--numbers-inline",
        dest="numbers_inline",
        action="store_true",
        help="Look for N count inline in each row (prior to 2021)",
    )

    args = parser.parse_args()

    body = extract(args.inpath, args.numbers_inline)
    sections = extract_sections(body)

    if args.format == "xlsx":
        dump_section_xlsx(sections, args.outpath)
    # elif args.format == 'csv':
    # 	dump_section_csv(sections, args.outpath)


if __name__ == "__main__":
    main()
