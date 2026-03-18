# Read a Google Docs document, Slides, or Sheets URL to a rdocx, rpptx, or rxlsx object

Uses the Google Docs, Slides, or Sheets URL to export a file locally,
read to an officer object. If filename is `NULL`, exported file is
removed after export.

## Usage

``` r
read_docs_url(
  url,
  format = NULL,
  filename = NULL,
  path = NULL,
  overwrite = TRUE,
  quiet = TRUE
)
```

## Arguments

- url:

  A URL for a Google Doc, Google Slides presentation, or Google Sheets.

- format:

  File format to use for export URL (typically set automatically).
  Options are `NULL` (default), "doc", "pptx", or "xslx". "pdf" and
  "csv" be supported in the future.

- filename:

  Destination file name. Optional. If filename is `NULL`, downloaded
  file is removed as part of the function execution.

- path:

  Folder path. Optional.

- overwrite:

  If `TRUE` (default), overwrite any existing file at the same location
  specified by `filename` and `path`.

- quiet:

  If `TRUE`, suppress messages when downloading file.

## See also

[`read_officer()`](https://elipousson.github.io/officerExtras/reference/read_officer.md)
