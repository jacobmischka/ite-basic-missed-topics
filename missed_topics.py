#!/usr/bin/env python3

from gooey import Gooey, GooeyParser
import pdftotext

from ite_basic_missed_topics.basic import basic, basic_excel
from ite_basic_missed_topics.ite import ite, ite_excel
from strip_header import strip_header

import os


def main():
    parser = GooeyParser(
        description="Create spreadsheet summary of ITE missed topics report"
    )
    required_group = parser.add_argument_group("Required inputs")
    required_group.add_argument(
        "inpath",
        help="Input PDF file (ITE ProgramItem report or Basic Program Summary of Examinees' Item Performance report)",
        widget="FileChooser",
    )
    required_group.add_argument(
        "outpath", help="Output file path, should end in .xlsx", widget="FileSaver"
    )
    required_group.add_argument(
        "-t",
        "--type",
        dest="type",
        choices=["basic", "ite"],
        required=True,
        help="Report type",
    )
    optional_group = parser.add_argument_group("Optional settings")
    optional_group.add_argument(
        "--numbers-inline",
        dest="numbers_inline",
        action="store_true",
        help="Look for N count inline in each row (prior to 2021)",
    )
    optional_group.add_argument(
        "--percentages-last",
        dest="percentages_last",
        action="store_true",
        help="Assume percentages at end of basic line (useful with fix_higher_percentages.py)",
    )

    args = parser.parse_args()

    with open(args.inpath, "rb") as infile:
        raw = "".join(pdftotext.PDF(infile, raw=True))
        stripped = "\n".join(strip_header(raw.splitlines()))

    if args.type == "ite":
        _, _, body = ite.extract(stripped)
        sections = ite.extract_sections(body)

        ite_excel.dump_section_xlsx(sections, args.outpath)

    elif args.type == "basic":
        body = basic.extract(stripped, args.numbers_inline)
        sections = basic.extract_sections(
            body,
            numbers_inline=args.numbers_inline,
            percentages_last=args.percentages_last,
        )

        basic_excel.dump_section_xlsx(sections, args.outpath)


if __name__ == "__main__":
    if not os.getenv("GUI_DISABLE"):
        main = Gooey(main)

    main()
