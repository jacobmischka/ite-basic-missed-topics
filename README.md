# Missed topic summaries for ITE and Basic exams

``` {.bash}
$ pdftotext -raw infile.pdf outfile.txt
$ ./strip_header.py outfile.txt outfile-stripped.txt
$ ./missed_topics.py -t $TYPE outfile-stripped.txt outfile.xlsx
```
