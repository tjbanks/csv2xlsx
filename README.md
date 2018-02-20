# csv2xlsx
Convert csv info to model readable xlsx.

```

usage: csv2xlsx.py [-h] input output

Convert CI model csv to xlsx.

positional arguments:
  input       Input csv file
  output      Output xlsx file

optional arguments:
  -h, --help  show this help message and exit

```

Input file is formatted as follows:
-------------------------------------------------------------------------------
Cell1Name Cell2Name ...

[%-connectivity-row1-and-col1 weight] [%-connectivity-row1-and-col2 weight] ...

[%-connectivity-row2-and-col1 weight] [%-connectivity-row2-and-col2 weight] ...

#OfCell1    #OfCell2

-------------------------------------------------------------------------------
!!Important!!
Rows correspond to column cell types.
