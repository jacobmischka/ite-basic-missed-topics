#!/usr/bin/env python3

from argparse import ArgumentParser

from ite_basic_missed_topics.basic import basic, basic_excel
from ite_basic_missed_topics.ite import ite, ite_csv, ite_excel


def main():
    parser = ArgumentParser(
        description="Create spreadsheet summary of ITE missed topics report"
    )
    parser.add_argument(
        "inpath", help="Input txt file (convert from pdf using `pdftotext -raw`)"
    )
    parser.add_argument("outpath", help="Output file path")
    parser.add_argument(
        "-t",
        "--type",
        dest="type",
        choices=["basic", "ite"],
        required=True,
        help="Report type",
    )
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
    parser.add_argument(
        "--percentages-last",
        dest="percentages_last",
        action="store_true",
        help="Assume percentages at end of basic line (useful with fix_higher_percentages.py)",
    )

    args = parser.parse_args()

    if args.type == "ite":
        _, _, body = ite.extract(args.inpath)
        sections = ite.extract_sections(body)

        if args.format == "xlsx":
            ite_excel.dump_section_xlsx(sections, args.outpath)
        elif args.format == "csv":
            ite_csv.dump_section_csv(sections, args.outpath)
    elif args.type == "basic":
        body = basic.extract(args.inpath, args.numbers_inline)
        sections = basic.extract_sections(
            body,
            numbers_inline=args.numbers_inline,
            percentages_last=args.percentages_last,
        )

        if args.format == "xlsx":
            basic_excel.dump_section_xlsx(sections, args.outpath)
        # elif args.format == 'csv':
        # 	basic_csv.dump_section_csv(sections, args.outpath)


if __name__ == "__main__":
    main()
