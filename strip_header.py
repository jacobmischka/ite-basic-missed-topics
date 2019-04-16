#!/usr/bin/env python3

from argparse import ArgumentParser

STRINGS_TO_SKIP = [
    "AMERICAN BOARD OF ANESTHESIOLOGY",
    "In-Training Examination",
    "156002",
    "Program :",
    "Program Summary",
    "BASIC Examination",
    "Page",
    "Listed below are the keyword phrases describing",
    "(B) after the keyword",
    "(A) after the keyword",
    "The numbers printed to the right of each",
    "PERCENT(%)",
    "This report is designed to help you identify specific",
]


def should_skip(line):
    for s in STRINGS_TO_SKIP:
        if s in line:
            return True

    return False


def strip_header(inpath, outpath):
    with open(inpath, "r") as infile, open(outpath, "w") as outfile:
        for line in infile:
            if not should_skip(line):
                outfile.write(line)


def main():
    parser = ArgumentParser()
    parser.add_argument("inpath")
    parser.add_argument("outpath")

    args = parser.parse_args()

    strip_header(args.inpath, args.outpath)


if __name__ == "__main__":
    main()
