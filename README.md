# Missed topic summaries for ITE and Basic exams

Produces a spreadsheet summarizing the aggregate results for AGCME's ITE
(ProgramItem report) or Basic (Program Summary of Examinees' Item Performance
report) exams.

```bash
$ pdftotext -raw infile.pdf outfile.txt
$ strip_acgme_header outfile.txt outfile-stripped.txt
$ acgme_missed_topics -t $TYPE outfile-stripped.txt outfile.xlsx
```

`$TYPE` is one of `basic` or `ite`.

## Installation

To install globally, build the wheel using `poetry build`, and install it using `pip`.

The following binaries are provided when installing, corresponding to the
`main` functions of their corresponding scripts:

- `strip_acgme_header`: `strip_header.py`
- `acgme_missed_topics`: `missed_topics.py`
- `ite_plot_percentiles`: `percentile_plotting.py`

## Criteria and configuration

The key at the top isn't particularly readable, but it's what fuels the
coloring and the marks in the Deficient columns.

If you expand the columns a bit, it shows that one is marked deficient if either:

1. the difference from the national mean is greater than 15% and the % missed is greater than 0%
2. the % missed is 100%

![Deficient legend](static/deficient.png)

Similarly, the numbers to the right of each color control the highlighting
cutoffs for the corresponding category. For example, the "Difference" section
is marked as green if it's between 0 and 20%, blue if 20% or greater, and light
red if between -10% and -20%.

![Highlighting legend](static/highlighting.png)

All of the numbers at the top of the spreadsheet can be changed, and the
highlighting and "Deficient Area" marks should update automatically!


## Percentile plotting


Produces training level bar charts from ITE Improvement in Performance report.
Report PDF should be converted to text and stripped as above.

```bash
$ ite_plot_percentiles < improvement-in-performance-stripped.txt 3< scaled-score-guideline-norm-table.txt
```

Takes ITE Guideline Norm Table Scaled Scores as input on file descriptor 3.
This should be converted to text using `pdftotext -raw`, and then manually
stripped of its header and everything below the STD DEV row of the Scaled Score
table.

It should look like this:

```
New
(N = 126)
CB
(N = 1142)
CA1
(N = 1807)
CA2
(N = 1786)
CA3
(N = 1674)
50 99 99 99 99 98
49 99 99 99 98 96
48 99 99 99 97 95
47 99 99 98 96 94
46 99 99 98 94 92
45 99 99 97 92 89
44 99 99 95 89 85
43 99 99 94 86 82
42 99 99 92 81 76
41 99 99 89 76 70
40 99 99 86 72 64
39 99 99 82 64 56
38 99 99 77 56 48
37 99 98 71 49 41
36 98 98 67 42 35
35 98 97 61 36 28
34 98 96 54 29 22
33 97 95 47 22 16
32 96 94 40 17 11
31 95 92 33 12 8
30 94 90 27 9 6
29 92 87 21 6 4
28 91 84 16 4 2
27 89 81 12 3 1
26 86 76 9 2 1
25 81 71 7 1 1
24 76 64 5 1 1
23 71 57 3 1 1
22 63 50 2 1 1
21 55 41 1 1 1
20 49 33 1 1 1
19 41 25 1 1 1
18 33 18 1 1 1
17 27 13 1 1 1
16 19 9 1 1 1
15 12 5 1 1 1
14 7 3 1 1 1
13 4 2 1 1 1
12 3 1 1 1 1
11 2 1 1 1 1
<= 10 1 1 1 1 1
MEAN 21 23 34 37 38
STD DEV 5 6 6 5 5
```

