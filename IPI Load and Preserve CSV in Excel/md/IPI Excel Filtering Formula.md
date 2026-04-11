# Excel Filtering

```excel
=LET(
  src, OFFSET(Sheet1!$($colLetter)2,,, COUNTA(Sheet1!$($colLetter):$colLetter) - 1, 1),
  res, INDEX(SORT(UNIQUE(FILTER(src, UPPER(src) > "")), 1, -1), SEQUENCE(50,,1)),
  TEXTJOIN("|", TRUE, IFERROR(res, ""))
)
```
