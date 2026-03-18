# Read a docx, pptx, potx, or xlsx file or use an existing object from officer if provided

`read_officer()` is a variant of
[`officer::read_docx()`](https://davidgohel.github.io/officer/reference/read_docx.html),
[`officer::read_pptx()`](https://davidgohel.github.io/officer/reference/read_pptx.html),
and
[`officer::read_xlsx()`](https://davidgohel.github.io/officer/reference/read_xlsx.html)
that allows users to read different Microsoft Office file types with a
single function. `read_docx_ext()`, `read_pptx_ext()`, and
`read_xlsx_ext()` are wrappers for `read_officer()` that require the
matching input file type. All versions allow both a filename and path
(the officer functions only use a path). If a rdocx, rpptx, or rxlsx
class object is provided to x, the object is checked based on the
fileext parameter and then returned as is.

## Usage

``` r
read_officer(
  filename = NULL,
  path = NULL,
  fileext = c("docx", "pptx", "xlsx"),
  x = NULL,
  arg = caller_arg(x),
  allow_null = TRUE,
  quiet = TRUE,
  call = parent.frame(),
  ...
)

read_docx_ext(
  filename = NULL,
  path = NULL,
  docx = NULL,
  allow_null = FALSE,
  quiet = TRUE
)

read_pptx_ext(
  filename = NULL,
  path = NULL,
  pptx = NULL,
  allow_null = FALSE,
  quiet = TRUE
)

read_xlsx_ext(
  filename = NULL,
  path = NULL,
  xlsx = NULL,
  allow_null = FALSE,
  quiet = TRUE
)
```

## Arguments

- filename, path:

  File name and path. Default: `NULL`. Must include a "docx", "pptx", or
  "xlsx" file path. "dotx" and "potx" files are also supported.

- fileext:

  File extensions to allow without error. Defaults to "docx", "pptx",
  "xlsx".

- x:

  A rdocx, rpptx, or rxlsx class object If x is provided, filename and
  path are ignored. Default: `NULL`

- arg:

  Argument name of object to check. Used to improve
  [`cli::cli_abort()`](https://cli.r-lib.org/reference/cli_abort.html)
  messages. Defaults to `caller_arg(x)`.

- allow_null:

  If `TRUE`, function supports the default behavior of
  [`officer::read_docx()`](https://davidgohel.github.io/officer/reference/read_docx.html),
  [`officer::read_pptx()`](https://davidgohel.github.io/officer/reference/read_pptx.html),
  or
  [`officer::read_xlsx()`](https://davidgohel.github.io/officer/reference/read_xlsx.html)
  and returns an empty document if x, filename, and path are all `NULL`.
  If `FALSE`, one of the three parameters must be supplied.

- quiet:

  If `FALSE`, warn if docx is provided when filename and/or path are
  also provided. Default: `TRUE`.

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

- ...:

  Additional parameters passed to
  [`cli::cli_abort()`](https://cli.r-lib.org/reference/cli_abort.html)

- docx, pptx, xlsx:

  A rdocx, rpptx, or rxlsx class object passed to the x parameter of
  `read_officer()` by the variant functions. Defaults to `NULL`.

## Value

A rdocx, rpptx, or rxlsx object.

## See also

[`officer::read_docx()`](https://davidgohel.github.io/officer/reference/read_docx.html)
