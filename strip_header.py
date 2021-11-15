#!/usr/bin/env python3

from argparse import ArgumentParser

STRINGS_TO_SKIP = [
    "AMERICAN BOARD OF ANESTHESIOLOGY",
    "In-Training Examination",
    "Program :",
    "Program:",
    "Program Summary",
    "BASIC Examination",
    "Page",
    "Listed below are the keyword phrases describing",
    "(B) after the keyword",
    "(A) after the keyword",
    "The numbers printed to the right of each",
    "PERCENT(%)",
    "This report is designed to help you identify specific",
    "This report assists you",
    "performance of each of your residents",
    "For each resident,",
    "These scores range from a low",
    "high of 50",
    "Level of Training",
    "Scaled Score",
    "Improvement In Performance Report",
]


def should_skip(line):
    if len(line) == 0:
        return True

    for s in STRINGS_TO_SKIP:
        if s in line:
            return True

    return False


def strip_header(infile):
    for line in infile:
        if not should_skip(line):
            yield line


def main():
    parser = ArgumentParser()
    parser.add_argument("inpath")
    parser.add_argument("outpath")

    args = parser.parse_args()

    with open(args.inpath, "r") as infile, open(args.outpath, "w") as outfile:
        for line in strip_header(infile):
            outfile.write(line)


if __name__ == "__main__":
    main()
