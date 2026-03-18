# Convert an rodcx object or Word file to another format with pandoc

`convert_docx()` uses
[`rmarkdown::pandoc_convert()`](https://pkgs.rstudio.com/rmarkdown/reference/pandoc_convert.html)
to convert a rodcx object or Word file to another format with pandoc. If
the "to" parameter contains a file extension, it is assumed to be an
output file name. If you want to convert a file to a Word document, use
the input parameter for the path to the Markdown, HTML, or other file.

## Usage

``` r
convert_docx(
  docx = NULL,
  to = NULL,
  input = NULL,
  output = NULL,
  path = NULL,
  options = NULL,
  extract_media = TRUE,
  overwrite = TRUE,
  quiet = TRUE,
  ...
)
```

## Arguments

- docx:

  A rdocx object or path to a docx file. Optional if input is provided.

- to:

  Format to convert to (if not specified, you must specify `output`)

- input:

  Character vector containing paths to input files (files must be UTF-8
  encoded)

- output:

  Output file (if not specified then determined based on format being
  converted to).

- path:

  File path for converted file passed to wd parameter of
  [`rmarkdown::pandoc_convert()`](https://pkgs.rstudio.com/rmarkdown/reference/pandoc_convert.html).
  If docx is a rdocx object, path defaults to current working directory
  instead of the base directory of the input file path.

- options:

  Character vector of command line options to pass to pandoc.

- extract_media:

  If `TRUE`, append `"--extract-media=."` to options. Defaults to
  `FALSE`.

- overwrite:

  If `TRUE` (default), remove file at path if it already exists. If
  `FALSE` and file exists, this function aborts.

- quiet:

  If `TRUE`, suppress alert messages and pass `FALSE` to verbose
  parameter of
  [`rmarkdown::pandoc_convert()`](https://pkgs.rstudio.com/rmarkdown/reference/pandoc_convert.html).
  Defaults `TRUE`.

- ...:

  Arguments passed on to
  [`rmarkdown::pandoc_convert`](https://pkgs.rstudio.com/rmarkdown/reference/pandoc_convert.html)

  `from`

  :   Format to convert from (if not specified then the format is
      determined based on the file extension of `input`).

  `citeproc`

  :   `TRUE` to run the pandoc-citeproc filter (for processing
      citations) as part of the conversion.

## Value

Executes a call to pandoc using
[`rmarkdown::pandoc_convert()`](https://pkgs.rstudio.com/rmarkdown/reference/pandoc_convert.html)
to create a file from an officer object or a docx file.

## Details

Supported input and output formats are described in the [pandoc user
guide](https://pandoc.org/MANUAL.html).

The system path as well as the version of pandoc shipped with RStudio
(if running under RStudio) are scanned for pandoc and the highest
version available is used.

## See also

[`rmarkdown::pandoc_convert()`](https://pkgs.rstudio.com/rmarkdown/reference/pandoc_convert.html)

## Examples

``` r
docx_example <- read_officer(
  system.file("doc_examples/example.docx", package = "officer")
)

convert_docx(
  docx_example,
  to = "markdown"
)

withr::with_tempdir({
  convert_docx(
    docx_example,
    output = "docx_example.html"
  )
})
```
