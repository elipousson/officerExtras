# Write a rdocx, rpptx, or rxlsx object to a file

Write a rdocx, rpptx, or rxlsx object to a file

## Usage

``` r
write_officer(x, path, overwrite = TRUE, modified_by = Sys.getenv("USER"), ...)
```

## Arguments

- x:

  A rdocx, rpptx, or rxlsx object to save. Document properties won't be
  set for rxlsx objects.

- path:

  File path.

- overwrite:

  If `TRUE` (default), remove file at path if it already exists. If
  `FALSE` and file exists, this function aborts.

- modified_by:

  If the withr package is installed, modified_by overrides the default
  value for the lastModifiedBy property assigned to the output file by
  officer. Defaults to `Sys.getenv("USER")` (the same value used by
  officer).

- ...:

  Arguments passed on to
  [`officer::set_doc_properties`](https://davidgohel.github.io/officer/reference/set_doc_properties.html)

  `title,subject,creator,description`

  :   text fields

  `created`

  :   a date object

  `hyperlink_base`

  :   a string specifying the base URL for relative hyperlinks in the
      document (only for rdocx).

  `values`

  :   a named list (names are field names), each element is a single
      character value specifying value associated with the corresponding
      field name. If `values` is provided, argument `...` will be
      ignored.

## Value

Returns the input object (invisibly) and writes the rdocx, rpptx, or
rxlsx object to a file with a name and location matching the provided
path.
