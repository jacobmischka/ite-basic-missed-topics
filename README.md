# Missed topic summaries for ITE and Basic exams

Produces a spreadsheet summarizing the aggregate results for AGCME's ITE or
Basic exams.

```bash
$ pdftotext -raw infile.pdf outfile.txt
$ ./strip_header.py outfile.txt outfile-stripped.txt
$ ./missed_topics.py -t $TYPE outfile-stripped.txt outfile.xlsx
```

`$TYPE` is one of `basic` or `ite`.

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

