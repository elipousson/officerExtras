# Fill a new column based on an existing column depending on a pattern

This function uses
[`vctrs::vec_fill_missing()`](https://vctrs.r-lib.org/reference/vec_fill_missing.html)
to convert hierarchically nested headings and text into a rectangular
data.frame. It is an experimental function that may be modified or
removed. At present, it is only used by
[`officer_tables()`](https://elipousson.github.io/officerExtras/reference/officer_tables.md).

## Usage

``` r
fill_with_pattern(
  x,
  pattern = "^heading",
  pattern_col = "style_name",
  fill_col = "text",
  col = "heading",
  direction = c("down", "up", "downup", "updown"),
  call = caller_env()
)
```

## Arguments

- x:

  A input data.frame (assumed to be from
  [`officer_summary()`](https://elipousson.github.io/officerExtras/reference/officer_summary.md)
  for default values).

- pattern:

  Passed to [`grepl()`](https://rdrr.io/r/base/grep.html) as the pattern
  identifying which rows of the fill_col should have values pulled into
  the new column named by col. Defaults to "^heading" which matches the
  default heading style names.

- pattern_col:

  Name of column to use for pattern matching, Defaults to "style_name".

- fill_col:

  Name of column to fill , Defaults to "text".

- col:

  Name of new column to fill with values from fill_col, Default:
  Defaults to "heading"

- direction:

  Direction of fill passed to
  [`vctrs::vec_fill_missing()`](https://vctrs.r-lib.org/reference/vec_fill_missing.html),
  Default: c("down", "up", "downup", "updown")

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

## Value

A data.frame with an additional column taking the name from col and the
values from the column named in fill_col.

## See also

[`vctrs::vec_fill_missing()`](https://vctrs.r-lib.org/reference/vec_fill_missing.html)
