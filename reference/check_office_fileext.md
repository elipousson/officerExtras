# Check file path for a docx, pptx, or xlsx file extension

Check file path for a docx, pptx, or xlsx file extension

## Usage

``` r
check_office_fileext(
  x,
  arg = caller_arg(x),
  fileext = c("docx", "pptx", "xlsx"),
  allow_null = FALSE,
  call = caller_env(),
  ...
)

check_docx_fileext(x, arg = caller_arg(x), call = caller_env(), ...)

check_pptx_fileext(x, arg = caller_arg(x), call = caller_env(), ...)

check_xlsx_fileext(x, arg = caller_arg(x), call = caller_env(), ...)
```

## Arguments

- x:

  Object to check.

- arg:

  Argument name of object to check. Used to improve
  [`cli::cli_abort()`](https://cli.r-lib.org/reference/cli_abort.html)
  messages. Defaults to `caller_arg(x)`.

- fileext:

  File extensions to allow without error. Defaults to "docx", "pptx",
  "xlsx".

- allow_null:

  If `TRUE` and x is `NULL`, invisibly return `NULL`. If `FALSE`
  (default), error unless x is a character vector with file extension
  matching the supplied fileext value.

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

- ...:

  Additional parameters passed to
  [`cli::cli_abort()`](https://cli.r-lib.org/reference/cli_abort.html)
