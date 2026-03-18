# Summarize a rdocx or rpptx object

`officer_summary()` extends
[`officer::docx_summary()`](https://davidgohel.github.io/officer/reference/docx_summary.html)
and other officer summary functions by handling multiple input types
within a single function. The preserve parameter is supported by officer
version \>= 0.6.3 (currently the development version) and it is ignored
unless a minimum supported version of officer is installed.

## Usage

``` r
officer_summary(
  x,
  summary_type = "doc",
  index = NULL,
  preserve = FALSE,
  as_tibble = TRUE,
  call = caller_env()
)
```

## Arguments

- x:

  A rdocx or rpptx object passed to
  [`officer::docx_summary()`](https://davidgohel.github.io/officer/reference/docx_summary.html),
  [`officer::pptx_summary()`](https://davidgohel.github.io/officer/reference/pptx_summary.html),
  [`officer::slide_summary()`](https://davidgohel.github.io/officer/reference/slide_summary.html),
  or
  [`officer::layout_summary()`](https://davidgohel.github.io/officer/reference/layout_summary.html).
  If x is a data frame created with one of those functions, it is
  returned as is (this feature may be removed in the future).

- summary_type:

  Summary type. Options "doc", "docx", "pptx", "slide", or "layout".
  "doc" requires the data.frame include a "content_type" column but
  allows columns for either a docx or pptx summary.

- index:

  slide index

- preserve:

  If `FALSE` (default), text in table cells is collapsed into a single
  line. If `TRUE`, line breaks in table cells are preserved as a "\n"
  character. This feature is adapted from
  `docxtractr::docx_extract_tbl()` published under a [MIT
  licensed](https://github.com/hrbrmstr/docxtractr/blob/master/LICENSE)
  in the 'docxtractr' package by Bob Rudis.

- as_tibble:

  If `TRUE` (default), return a tibble data frame.

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

A tibble or data frame object.

## See also

Other summary functions:
[`check_officer_summary()`](https://elipousson.github.io/officerExtras/reference/check_officer_summary.md),
[`officer_summary_levels()`](https://elipousson.github.io/officerExtras/reference/officer_summary_levels.md)

## Examples

``` r
docx_example <- read_officer(
  system.file("doc_examples", "example.docx", package = "officer")
)

officer_summary(docx_example)
#> # A tibble: 56 × 11
#>    doc_index content_type style_name  text  table_index row_id cell_id is_header
#>        <int> <chr>        <chr>       <chr>       <int>  <int>   <int> <lgl>    
#>  1         1 paragraph    heading 1   "Tit…          NA     NA      NA NA       
#>  2         2 paragraph    NA          "Lor…          NA     NA      NA NA       
#>  3         3 paragraph    heading 1   "Tit…          NA     NA      NA NA       
#>  4         4 paragraph    List Parag… "Qui…          NA     NA      NA NA       
#>  5         5 paragraph    List Parag… "Aug…          NA     NA      NA NA       
#>  6         6 paragraph    List Parag… "Sap…          NA     NA      NA NA       
#>  7         7 paragraph    heading 2   "Sub…          NA     NA      NA NA       
#>  8         8 paragraph    List Parag… "Qui…          NA     NA      NA NA       
#>  9         9 paragraph    List Parag… "Aug…          NA     NA      NA NA       
#> 10        10 paragraph    List Parag… "Sap…          NA     NA      NA NA       
#> # ℹ 46 more rows
#> # ℹ 3 more variables: row_span <int>, col_span <chr>, table_stylename <chr>

pptx_example <- read_officer(
  "example.pptx", system.file("doc_examples", package = "officer")
)

officer_summary(pptx_example)
#> # A tibble: 54 × 9
#>    text  id    content_type slide_id row_id cell_id col_span row_span media_file
#>    <chr> <chr> <chr>           <int>  <int>   <int>    <int>    <int> <chr>     
#>  1 "Tit… 12    paragraph           1     NA      NA       NA       NA NA        
#>  2 "A t… 13    paragraph           1     NA      NA       NA       NA NA        
#>  3 "and… 13    paragraph           1     NA      NA       NA       NA NA        
#>  4 "and… 13    paragraph           1     NA      NA       NA       NA NA        
#>  5 "and… 13    paragraph           1     NA      NA       NA       NA NA        
#>  6 "Hea… 18    table cell          1      1       1        1        1 NA        
#>  7 "A"   18    table cell          1      2       1        1        1 NA        
#>  8 "B"   18    table cell          1      3       1        1        1 NA        
#>  9 "B"   18    table cell          1      4       1        1        1 NA        
#> 10 "C"   18    table cell          1      5       1        1        1 NA        
#> # ℹ 44 more rows

officer_summary(pptx_example, "slide", 1)
#> # A tibble: 5 × 11
#>   type  id    ph_label     offx  offy    cx    cy rotation fld_id fld_type text 
#>   <chr> <chr> <chr>       <dbl> <dbl> <dbl> <dbl>    <int> <chr>  <chr>    <chr>
#> 1 title 12    Titre 11       NA    NA    NA    NA       NA NA     NA       "Tit…
#> 2 body  13    Espace rés…    NA    NA    NA    NA       NA NA     NA       "A t…
#> 3 body  18    Espace rés…    NA    NA    NA    NA       NA NA     NA       "{5C…
#> 4 body  15    Espace rés…    NA    NA    NA    NA       NA NA     NA       "R l…
#> 5 body  17    Espace rés…    NA    NA    NA    NA       NA NA     NA       ""   
```
