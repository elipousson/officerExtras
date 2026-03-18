# Copy media from a docx or pptx file or a rdocx or rpptx object to a target folder

Unzip a docx or pptx file to a temporary directory, check if the
directory contains a media folder, and copy media files to the directory
set by dir. If a rdocx or rpptx object is provided, files are copied
from the temporary package_dir associated with the object (accessible
via `x[["package_dir"]]`).

## Usage

``` r
officer_media(
  filename = NULL,
  path = NULL,
  x = NULL,
  target = "media",
  list = FALSE,
  overwrite = TRUE
)
```

## Arguments

- filename, path:

  File name and path for a `docx` or `pptx` file. One of the two must be
  provided. Ignored if x is provided. Defaults to `NULL`.

- x:

  A rdocx or rpptx object containing one or more media file.

- target:

  Folder name or path to copy media files. dir is created if no folder
  exists at that location.

- list:

  If `TRUE`, display a message listing files contained in the docx or
  pptx file but do not copy the files to dir. Defaults to `FALSE`.

- overwrite:

  If `TRUE` (default), overwrite any files with the same names at target
  path. If `FALSE`, abort if files with the same names already exist.

## See also

[`officer::media_extract()`](https://davidgohel.github.io/officer/reference/media_extract.html)

## Examples

``` r
officer_media(
  system.file("doc_examples/example.pptx", package = "officer"),
  list = TRUE
)
#> ℹ 1 media file found in example.pptx
#> • image1.png
```
