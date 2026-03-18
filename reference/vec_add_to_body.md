# Add a vector of objects to a rdocx object using `add_to_body()`

`vec_add_to_body()` is a vectorized variant
[`add_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
that allows users to supply a vector of values or str inputs to add
multiple blocks of text, images, plots, or tables to a document.
Alternatively, the function also supports adding a single object at
multiple locations or using multiple different styles. All parameters
are recycled using
[`vctrs::vec_recycle_common()`](https://vctrs.r-lib.org/reference/vec_recycle.html)
so inputs must be length 1 or match the length of the longest input
vector. Optionally, the function can apply a separator between each
element by passing a value to
[`add_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
or passing docx to a function, such as
[`officer::body_add_break()`](https://davidgohel.github.io/officer/reference/body_add_break.html).

## Usage

``` r
vec_add_to_body(
  docx,
  ...,
  .sep = NULL,
  .pos = "after",
  .size = NULL,
  .call = caller_env()
)
```

## Arguments

- docx:

  A rdocx object.

- ...:

  Arguments passed on to
  [`add_to_body`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)

  `gt_object`

  :   A gt object converted to an OOXML string with
      [`gt::as_word()`](https://gt.rstudio.com/reference/as_word.html)
      then passed to
      [`add_xml_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
      as str parameter. Required for
      [`add_gt_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md).

  `keyword,id`

  :   A keyword string used to place cursor with
      [`officer::cursor_reach()`](https://davidgohel.github.io/officer/reference/cursor.html)
      or bookmark id with
      [`officer::cursor_bookmark()`](https://davidgohel.github.io/officer/reference/cursor.html).
      Defaults to `NULL`. If keyword or id are not provided, the gt
      object is inserted at the front of the document.

  `index`

  :   A integer matching a doc_index value appearing in a summary of the
      docx object created with
      [`officer::docx_summary()`](https://davidgohel.github.io/officer/reference/docx_summary.html).
      If index is for a paragraph value, the text of the pargraph is
      used as a keyword.

  `call`

  :   The execution environment of a currently running function, e.g.
      `caller_env()`. The function will be mentioned in error messages
      as the source of the error. See the `call` argument of
      [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
      information.

  `value`

  :   object to add in the document. Supported objects are vectors,
      data.frame, graphics, block of formatted paragraphs, unordered
      list of formatted paragraphs, pretty tables with package
      flextable, 'Microsoft' charts with package mschart.

  `style`

  :   paragraph style name. These names are available with function
      [styles_info](https://davidgohel.github.io/officer/reference/styles_info.html)
      and are the names of the Word styles defined in the base document
      (see argument `path` from
      [read_docx](https://davidgohel.github.io/officer/reference/read_docx.html)).

  `str`

  :   a wml string

  `pos`

  :   where to add the new element relative to the cursor, one of
      "after", "before", "on".

- .sep:

  A bare function, such as
  [officer::body_add_break](https://davidgohel.github.io/officer/reference/body_add_break.html)
  or another object passed to
  [`add_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
  as the value parameter.

- .pos:

  String passed to pos parameter if
  [`add_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
  with .sep if .sep is not a function. Defaults to "after".

- .size:

  Desired output size.

- .call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

## Examples

``` r
docx_example <- read_officer()

docx_example <- vec_add_to_body(
  docx_example,
  value = c("Sample text 1", "Sample text 2", "Sample text 3"),
  style = c("heading 1", "heading 2", "Normal")
)

docx_example <- vec_add_to_body(
  docx_example,
  value = rep("Text", 5),
  style = "Normal",
  .sep = officer::body_add_break
)

officer_summary(docx_example)
#> # A tibble: 12 × 11
#>    doc_index content_type style_name text   table_index row_id cell_id is_header
#>        <int> <chr>        <chr>      <chr>        <int>  <int>   <int> <lgl>    
#>  1         2 paragraph    heading 1  "Samp…          NA     NA      NA NA       
#>  2         3 paragraph    heading 2  "Samp…          NA     NA      NA NA       
#>  3         4 paragraph    Normal     "Samp…          NA     NA      NA NA       
#>  4         5 paragraph    Normal     "Text"          NA     NA      NA NA       
#>  5         6 paragraph    NA         "\n"            NA     NA      NA NA       
#>  6         7 paragraph    Normal     "Text"          NA     NA      NA NA       
#>  7         8 paragraph    NA         "\n"            NA     NA      NA NA       
#>  8         9 paragraph    Normal     "Text"          NA     NA      NA NA       
#>  9        10 paragraph    NA         "\n"            NA     NA      NA NA       
#> 10        11 paragraph    Normal     "Text"          NA     NA      NA NA       
#> 11        12 paragraph    NA         "\n"            NA     NA      NA NA       
#> 12        13 paragraph    Normal     "Text"          NA     NA      NA NA       
#> # ℹ 3 more variables: row_span <int>, col_span <chr>, table_stylename <chr>

if (rlang::is_installed("gt")) {
  gt_tbl <- gt::gt(gt::gtcars[1:2, 1:2])

  # list inputs such as gt tables must be passed within a list to avoid
  # issues
  docx_example <- vec_add_to_body(
    docx_example,
    gt_object = list(gt_tbl, gt_tbl),
    keyword = c("Sample text 1", "Sample text 2")
  )

  officer_summary(docx_example)
}
#> # A tibble: 24 × 11
#>    doc_index content_type style_name text   table_index row_id cell_id is_header
#>        <int> <chr>        <chr>      <chr>        <int>  <int>   <int> <lgl>    
#>  1         2 paragraph    heading 1  "Samp…          NA     NA      NA NA       
#>  2         9 paragraph    heading 2  "Samp…          NA     NA      NA NA       
#>  3        16 paragraph    Normal     "Samp…          NA     NA      NA NA       
#>  4        17 paragraph    Normal     "Text"          NA     NA      NA NA       
#>  5        18 paragraph    NA         "\n"            NA     NA      NA NA       
#>  6        19 paragraph    Normal     "Text"          NA     NA      NA NA       
#>  7        20 paragraph    NA         "\n"            NA     NA      NA NA       
#>  8        21 paragraph    Normal     "Text"          NA     NA      NA NA       
#>  9        22 paragraph    NA         "\n"            NA     NA      NA NA       
#> 10        23 paragraph    Normal     "Text"          NA     NA      NA NA       
#> # ℹ 14 more rows
#> # ℹ 3 more variables: row_span <int>, col_span <chr>, table_stylename <chr>
```
