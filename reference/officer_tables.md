# Get tables from a rdocx or rpptx object

Get one or more tables from a rdocx or rpptx object. `officer_tables()`
returns a list of data frames and `officer_table()` returns a single
table as a data frame. These functions are based on example code on
extracting Word document and PowerPoint slides in the [officeverse
documentation](https://ardata-fr.github.io/officeverse/extract-content.html#word-tables).
Some additional features including the type_convert parameter and the
addition of doc_index values as the default names for the returned list
of tables are based on [this blog post by Matt
Dray](https://www.rostrum.blog/2023/06/07/rectangular-officer/).

## Usage

``` r
officer_tables(
  x,
  index = NULL,
  has_header = TRUE,
  col = NULL,
  preserve = FALSE,
  ...,
  stack = FALSE,
  type_convert = FALSE,
  nm = NULL,
  call = caller_env()
)

officer_table(
  x,
  index = NULL,
  has_header = TRUE,
  col = NULL,
  ...,
  call = caller_env()
)
```

## Arguments

- x:

  A rdocx or rpptx object or a data frame created with
  [`officer_summary()`](https://elipousson.github.io/officerExtras/reference/officer_summary.md).

- index:

  A index value matching a doc_index value for a table in the summary
  data frame, Default: `NULL`

- has_header:

  If `TRUE` (default), tables are expected to have implicit headers even
  if the Word table does not have an explicit header row. If `FALSE`,
  only explicit header rows will be used as column names.

- col:

  If col is supplied, `officer_table()` passes col and the additional
  parameters in ... to
  [`fill_with_pattern()`](https://elipousson.github.io/officerExtras/reference/fill_with_pattern.md).
  This allows the addition of preceding headings or captions as a column
  within the data.frame returned by `officer_tables()`. This is an
  experimental feature and may be modified or removed. Defaults to
  `NULL`.

- preserve:

  If `FALSE` (default), text in table cells is collapsed into a single
  line. If `TRUE`, line breaks in table cells are preserved as a "\n"
  character. This feature is adapted from
  `docxtractr::docx_extract_tbl()` published under a [MIT
  licensed](https://github.com/hrbrmstr/docxtractr/blob/master/LICENSE)
  in the 'docxtractr' package by Bob Rudis.

- ...:

  Additional parameters passed to
  [`fill_with_pattern()`](https://elipousson.github.io/officerExtras/reference/fill_with_pattern.md).

- stack:

  If `TRUE` and all tables share the same number of columns, return a
  single combined data frame instead of a list. Defaults to `FALSE`.

- type_convert:

  If `TRUE`, convert columns for the returned data frames to the
  appropriate type using
  [`utils::type.convert()`](https://rdrr.io/r/utils/type.convert.html).

- nm:

  Names to use for returned list of tables. If `NULL` (default), the
  names are set to the table_index values using the pattern
  "table_index\_\<table_index_number\>".

- call:

  The execution environment of a currently running function, e.g.
  `call = caller_env()`. The corresponding function call is retrieved
  and mentioned in error messages as the source of the error.

  You only need to supply `call` when throwing a condition from a helper
  function which wouldn't be relevant to mention in the message.

  Can also be `NULL` or a [defused function
  call](https://rlang.r-lib.org/reference/topic-defuse.html) to
  respectively not display any call or hard-code a code to display.

  For more information about error calls, see [Including function calls
  in error
  messages](https://rlang.r-lib.org/reference/topic-error-call.html).

## Value

A list of data frames or, if stack is `TRUE`, a single data frame.

## See also

`docxtractr::docx_extract_all()`

## Examples

``` r
docx_example <- read_docx_ext(
  filename = "example.docx",
  path = system.file("doc_examples", package = "officer")
)

officer_tables(docx_example)
#> $table_index_1
#>         Petals   Internode       Sepal                                Bract
#> 1  5,621498349        NULL 2,462106579                           18,2034091
#> 2  4,994616997          AA 2,429320759                          17,65204912
#> 3  4,767504884        NULL         AAA                                 NULL
#> 4   25,9242382        NULL 2,066051345                          18,37915478
#> 5  6,489375001 25,21130805 2,901582763 17,31304737, 17,07215724, 18,2902189
#> 6    5,7858682 25,52433147 2,655642742                                 NULL
#> 7  5,645575295 Merged cell 2,278691288                                 NULL
#> 8  4,828953215        NULL 2,238467716                          19,87376227
#> 9  6,783500773        NULL 2,202762147                          19,85326662
#> 10 5,395076839        NULL 2,538375992                          19,56545356
#> 11 4,683617783  29,2459239 2,601945544                          18,95335451
#> 12        Note        NULL        NULL                        New line note
#> 

pptx_example <- read_pptx_ext(
  filename = "example.pptx",
  path = system.file("doc_examples", package = "officer")
)

officer_tables(pptx_example)[[1]]
#>   Header 1  Header 2       Header 3
#> 2         A    12.23      blah blah
#> 3         B     1.23 blah blah blah
#> 4         B      9.0          Salut
#> 5         C        6          Hello
```
