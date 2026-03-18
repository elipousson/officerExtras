# Set cursor position in rdocx object based on keyword, id, or index

A combined function for setting cursor position with
[`officer::cursor_reach()`](https://davidgohel.github.io/officer/reference/cursor.html),
[`officer::cursor_bookmark()`](https://davidgohel.github.io/officer/reference/cursor.html),
or using a doc_index value from
[`officer::docx_summary()`](https://davidgohel.github.io/officer/reference/docx_summary.html).
Defaults to using
[`officer::cursor_end()`](https://davidgohel.github.io/officer/reference/cursor.html),
[`officer::cursor_begin()`](https://davidgohel.github.io/officer/reference/cursor.html),
[`officer::cursor_backward()`](https://davidgohel.github.io/officer/reference/cursor.html),
or
[`officer::cursor_forward()`](https://davidgohel.github.io/officer/reference/cursor.html)
if keyword, id, and index are all `NULL`.

## Usage

``` r
cursor_docx(
  docx,
  keyword = NULL,
  id = NULL,
  index = NULL,
  default = "end",
  quiet = FALSE,
  call = caller_env()
)
```

## Arguments

- docx:

  A rdocx object.

- keyword, id:

  A keyword string used to place cursor with
  [`officer::cursor_reach()`](https://davidgohel.github.io/officer/reference/cursor.html)
  or bookmark id with
  [`officer::cursor_bookmark()`](https://davidgohel.github.io/officer/reference/cursor.html).
  Defaults to `NULL`. If keyword or id are not provided, the gt object
  is inserted at the front of the document.

- index:

  A integer matching a doc_index value appearing in a summary of the
  docx object created with
  [`officer::docx_summary()`](https://davidgohel.github.io/officer/reference/docx_summary.html).
  If index is for a paragraph value, the text of the pargraph is used as
  a keyword.

- default:

  Character string with one of the following options:
  `c("end", "begin", "backward", "forward")` to set cursor position.
  Only used if keyword, id, and index are all `NULL`.

- quiet:

  If `FALSE` (default) warn when keyword is not found.

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

## See also

[`officer::cursor_begin()`](https://davidgohel.github.io/officer/reference/cursor.html),
[`officer::docx_summary()`](https://davidgohel.github.io/officer/reference/docx_summary.html)
