# Check a summary data.frame created from a rdocx or rpptx object

Check an object and error if an object is not a data.frame with the
required column names to be a summary data.frame created from a rdocx or
rpptx object. Optionally check for the number of rows, a specific
content_type value, or if tables are included in the document the
summary was created from.

## Usage

``` r
check_officer_summary(
  x,
  n = NULL,
  content_type = NULL,
  summary_type = "doc",
  tables = FALSE,
  ...,
  arg = caller_arg(x),
  call = parent.frame()
)
```

## Arguments

- x:

  An object to check if it is a data.frame object created with
  [`officer::docx_summary()`](https://davidgohel.github.io/officer/reference/docx_summary.html)
  or another summary function.

- n:

  Required number of rows. Optional. If n is more than length 1, checks
  to make sure the number of rows is within the range of `max(n) `and
  `min(n)`. Defaults to `NULL`.

- content_type:

  Required content_type, e.g. "paragraph", "table cell", or "image".
  Optional. Defaults to `NULL`.

- summary_type:

  Summary type. Options "doc", "docx", "pptx", "slide", or "layout".
  "doc" requires the data.frame include a "content_type" column but
  allows columns for either a docx or pptx summary.

- tables:

  If `TRUE`, require that the summary include the column names indicated
  a table is present in the rdocx or rpptx summary.

- ...:

  Additional parameters passed to
  [`cli::cli_abort()`](https://cli.r-lib.org/reference/cli_abort.html)

- arg:

  Argument name to use in error messages. Defaults to `caller_arg(x)`

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

## See also

Other summary functions:
[`officer_summary()`](https://elipousson.github.io/officerExtras/reference/officer_summary.md),
[`officer_summary_levels()`](https://elipousson.github.io/officerExtras/reference/officer_summary_levels.md)
