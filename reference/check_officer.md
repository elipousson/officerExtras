# Check if x is a rdocx, rpptx, or rxlsx object

Check if x is a rdocx, rpptx, or rxlsx object

## Usage

``` r
check_officer(
  x,
  arg = caller_arg(x),
  what = c("rdocx", "rpptx", "rxlsx"),
  call = caller_env(),
  ...
)

check_docx(x, arg = caller_arg(x), call = caller_env(), ...)

check_pptx(x, arg = caller_arg(x), call = caller_env(), ...)

check_xlsx(x, arg = caller_arg(x), call = caller_env(), ...)

check_block_list(
  x,
  arg = caller_arg(x),
  allow_empty = FALSE,
  allow_null = FALSE,
  call = caller_env()
)
```

## Arguments

- x:

  Object to check.

- arg:

  Argument name of object to check. Used to improve
  [`cli::cli_abort()`](https://cli.r-lib.org/reference/cli_abort.html)
  messages. Defaults to `caller_arg(x)`.

- what:

  Class names to check

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

- ...:

  Additional parameters passed to
  [`cli::cli_abort()`](https://cli.r-lib.org/reference/cli_abort.html)

- allow_empty:

  If `TRUE`, `check_block_list()` allows an empty block list. Defaults
  to `FALSE`.

- allow_null:

  If `FALSE` (default), error if x is `NULL`.
