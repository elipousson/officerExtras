# Create new columns to an officer summary data.frame based on specific levels

`officer_summary_levels()` works with
[`fill_with_pattern()`](https://elipousson.github.io/officerExtras/reference/fill_with_pattern.md)
to help convert hierarchically organized text, e.g. text with leveled
headings or other styles into a data.frame where new columns hold the
value of the preceding (or succeeding) heading text.

## Usage

``` r
officer_summary_levels(
  x,
  levels = NULL,
  levels_from = "style_name",
  exclude_levels = NULL,
  fill_col = "text",
  direction = c("down", "up", "downup", "updown"),
  ...,
  strict = FALSE,
  call = caller_env()
)
```

## Arguments

- x:

  A input data.frame (assumed to be from
  [`officer_summary()`](https://elipousson.github.io/officerExtras/reference/officer_summary.md)
  for default values).

- levels:

  Levels to use. If `NULL` (default), levels is set to unique, non-NA
  values from the levels_from column.

- levels_from:

  Column name to use for identifying levels. Defaults to "style_name".

- exclude_levels:

  Levels to exclude from process of adding new columns.

- fill_col:

  Name of column to fill , Defaults to "text".

- direction:

  Direction of fill passed to
  [`vctrs::vec_fill_missing()`](https://vctrs.r-lib.org/reference/vec_fill_missing.html),
  Default: c("down", "up", "downup", "updown")

- ...:

  Arguments passed on to
  [`officer_summary`](https://elipousson.github.io/officerExtras/reference/officer_summary.md)

  `as_tibble`

  :   If `TRUE` (default), return a tibble data frame.

  `summary_type`

  :   Summary type. Options "doc", "docx", "pptx", "slide", or "layout".
      "doc" requires the data.frame include a "content_type" column but
      allows columns for either a docx or pptx summary.

  `index`

  :   slide index

  `preserve`

  :   If `FALSE` (default), text in table cells is collapsed into a
      single line. If `TRUE`, line breaks in table cells are preserved
      as a "\n" character. This feature is adapted from
      `docxtractr::docx_extract_tbl()` published under a [MIT
      licensed](https://github.com/hrbrmstr/docxtractr/blob/master/LICENSE)
      in the 'docxtractr' package by Bob Rudis.

- strict:

  If `TRUE`, error unless all values provided to `level` are present in
  the `levels_from` column. Defaults to `FALSE` which warns if invalid
  values are present.

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

## Value

A summary data frame with additional columns based on the supplied
levels.

## See also

Other summary functions:
[`check_officer_summary()`](https://elipousson.github.io/officerExtras/reference/check_officer_summary.md),
[`officer_summary()`](https://elipousson.github.io/officerExtras/reference/officer_summary.md)
