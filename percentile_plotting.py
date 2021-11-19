#!/usr/bin/env python3

from gooey import Gooey, GooeyParser
from dataclasses import dataclass
import pdftotext
import matplotlib.pyplot as plt

from strip_header import strip_header

from random import shuffle
import os, sys


@dataclass
class Trainee:
    last_name: str
    first_name: str
    id_number: int
    scores: dict

    @property
    def full_name(self) -> str:
        return "{}, {}".format(self.last_name, self.first_name)

    @classmethod
    def from_row(cls, row):
        comma_index = row.index(",")
        last_name = row[:comma_index]
        first_name_pieces = []
        row = row[comma_index + 2 :]
        words = row.split(" ")

        id_number = None
        scores = {}

        for word in words:
            if word.isdigit():
                id_number = int(word)
            elif id_number is None:
                first_name_pieces.append(word)
            else:
                scores[word] = 0

        if id_number is None or len(scores) == 0:
            raise ValueError

        return cls(last_name, " ".join(first_name_pieces), id_number, scores)

    def parse_scores(self, row):
        scores = row.split(" ")
        for i, k in enumerate(self.scores.keys()):
            self.scores[k] = int(scores[i])


def main():
    parser = GooeyParser(
        description="Create training level bar charts from ITE Improvement in Performance report"
    )
    required_group = parser.add_argument_group("Required inputs")
    required_group.add_argument(
        "scores",
        help="Path to ITE Improvement in Performance PDF",
        widget="FileChooser",
    )
    required_group.add_argument(
        "norm_file",
        help="Path to ITE Guideline Norm Table Scaled Scores PDF",
        widget="FileChooser",
    )

    args = parser.parse_args()

    with open(args.scores, "rb") as scores_file:
        raw = "".join(pdftotext.PDF(scores_file, raw=True))
        stripped = "\n".join(strip_header(raw.splitlines()))
        years, trainees = parse_scores(stripped)

    with open(args.norm_file, "rb") as norm_table_file:
        raw = "".join(pdftotext.PDF(norm_table_file, raw=True))
        norm_table = parse_norm_table(raw)

    for i, year in enumerate(years):
        plot_year(year, trainees, i, norm_table)


def plot_year(year, trainees, score_index, norm_table):
    trainees = [t for t in trainees if len(t.scores) > score_index]
    training_levels = {}
    for trainee in trainees:
        tl = list(trainee.scores.keys())[score_index]
        if tl not in training_levels:
            training_levels[tl] = []

        training_levels[tl].append(trainee)

    print(year)

    for level, trainees in training_levels.items():
        shuffle(trainees)

        fig, ax = plt.subplots()
        points = [norm_table[level][t.scores[level]] for t in trainees]
        x = range(len(points))
        ax.bar(x, points)
        ax.hlines(avg(points), -1, len(points), colors="red", label="Average")
        ax.hlines(median(points), -1, len(points), colors="orange", label="Median")
        ax.set_title(year)
        ax.set_ylabel("Percentile rank")
        ax.set_yticks(range(0, 100, 10))
        ax.set_xlabel(level)
        ax.set_xticks(x)
        fig.legend()

        fig.tight_layout()
        plt.tight_layout()

        print("\t", level)
        for i, trainee in enumerate(trainees):
            print("\t\t{}: {}".format(i, trainee.full_name))
        print()

    plt.show()


def avg(points):
    return sum(points) / len(points)


def median(points):
    return sorted(points)[len(points) // 2]


def parse_scores(pdf_text):
    years = None

    trainees = []
    trainee_lines = []
    trainee = None

    for line in pdf_text.splitlines():
        line = line.strip()
        if len(line) == 0:
            continue

        if is_header(line):
            if years is None:
                years = get_years(line)
            continue

        if trainee is None:
            trainee_lines.append(line)

            try:
                trainee = Trainee.from_row(" ".join(trainee_lines))
            except Exception as e:
                print(e, file=sys.stderr)
                pass
        else:
            trainee.parse_scores(line)
            trainees.append(trainee)
            trainee_lines = []
            trainee = None

    return years, trainees


def is_header(line):
    return "Resident Name ID Number" in line


def get_years(header_line):
    words = header_line.split(" ")
    return words[-4:]


def parse_norm_table(norm_table_text):
    year_names = []
    years = []

    in_header = True
    last_row = False

    for line in norm_table_text.splitlines():
        line = line.strip()

        if in_header:
            if line == "New":
                in_header = False
            else:
                continue

        if line.startswith("("):
            continue

        words = line.split(" ")
        if len(words) == 0:
            continue

        if len(words) == 1:
            year_names.append(words[0])
            years.append({})
        else:
            if words[0] == "<=":
                words.pop(0)
                last_row = True

            scaled_score = int(words.pop(0))

            for i, percentile_rank in enumerate(words):
                years[i][scaled_score] = int(percentile_rank)

                if last_row:
                    for score in range(scaled_score - 1, 0, -1):
                        years[i][score] = int(percentile_rank)

            if last_row:
                break

    return {year_name: years[i] for i, year_name in enumerate(year_names)}


if __name__ == "__main__":
    if not os.getenv("GUI_DISABLE"):
        main = Gooey(main, show_stop_warning=False)

    main()
