# Missed topic summaries for ITE and Basic exams

Produces a spreadsheet summarizing the aggregate results for AGCME's ITE or
Basic exams.

```bash
$ pdftotext -raw infile.pdf outfile.txt
$ ./strip_header.py outfile.txt outfile-stripped.txt
$ ./missed_topics.py -t $TYPE outfile-stripped.txt outfile.xlsx
```

`$TYPE` is one of `basic` or `ite`.
