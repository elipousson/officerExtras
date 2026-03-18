# Combine multiple rdocx objects or docx files

`combine_docx()` is a variant of
[`officer::body_add_docx()`](https://davidgohel.github.io/officer/reference/body_add_docx.html)
that allows any number of input files and supports rdocx objects as well
as Word file paths. Optionally use a separator between files or objects.

Please note that when you create a new rdocx object with this function
(or
[`officer::body_add_docx()`](https://davidgohel.github.io/officer/reference/body_add_docx.html))
the added content will not appear in a summary data frame created with
[`officer_summary()`](https://elipousson.github.io/officerExtras/reference/officer_summary.md)
and is not accessible to other functions until the document is *opened
and edited* with Microsoft Word. This is part of how the OOXML [AltChunk
Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.altchunk?view=openxml-2.8.1)
works and can't be avoided.

## Usage

``` r
combine_docx(
  ...,
  docx = NULL,
  .list = list2(...),
  pos = "after",
  sep = NULL,
  call = caller_env()
)
```

## Arguments

- ...:

  Any number of additional rdocx objects or docx file paths.

- docx:

  A rdocx object or a file path with a docx file extension. Defaults to
  `NULL`.

- .list:

  Any number of additional rdocx objects or docx file paths passed as a
  list. Defaults to
  [`rlang::list2()`](https://rlang.r-lib.org/reference/list2.html)

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".

- sep:

  Separator to use docx files. A bare function that takes a rdocx object
  as the only parameter, such as
  [officer::body_add_break](https://davidgohel.github.io/officer/reference/body_add_break.html)
  or another object passed to
  [`add_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md)
  as the value parameter. Optional. Defaults to `NULL`.

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

## Value

A rdocx object.

## See also

[`officer::body_add_docx()`](https://davidgohel.github.io/officer/reference/body_add_docx.html)

## Examples

``` r
if (FALSE) { # \dontrun{
if (interactive()) {
  docx_path <- system.file("doc_examples", "example.docx", package = "officer")
  docx <- read_officer(docx_path)

  combine_docx(
    docx,
    docx_path
  )
}
} # }
```
