# Formula to filter dates by

## Current

**Note:** Formula to be used in PowerShell module.

```excel
=LET(
  src,  FILTER(Data!$A:$J, (INDEX(Data!$A:$J,, 1) > "")),
  hdrs, INDEX(src, 1,),
  dt,   FILTER(Dates!$A:$A, (INDEX(Dates!$A:$A,, 1) > 0)*(ISNUMBER(Dates!$A:$A))),
  flt,  FILTER(src, (ISNUMBER(XMATCH(INDEX(src,, 1), TEXT(dt, "yyyy-mm-dd"), 0, 1)))),
  rowA, ROWS(hdrs),
  rowB, ROWS(flt),
  seq,  SEQUENCE(rowA + rowB),
  IFERROR(
     IF(seq <= rowA, INDEX(hdrs, seq, {1,4,6,7,8}), INDEX(flt, seq - rowA, {1,4,6,7,8})),
  "")
)
```
