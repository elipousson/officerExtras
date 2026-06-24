# Preview a rdocx, rpptx, or rxlsx object in local default applications

`officer_open()` uses
[`officer::open_file()`](https://davidgohel.github.io/officer/reference/open_file.html)
to open a file ceated from a rdocx, rpptx, or rxlsx object.

## Usage

``` r
officer_open(
  x,
  path = NULL,
  ...,
  overwrite = FALSE,
  interactive = is_interactive()
)
```

## Arguments

- x:

  A rdocx, rpptx, or rxlsx object to save. Document properties won't be
  set for rxlsx objects.

- path:

  Optional. Set as temporary file with file extension matching type of
  input object `x`.

- ...:

  Arguments passed on to
  [`write_officer`](https://elipousson.github.io/officerExtras/reference/write_officer.md)

  `modified_by`

  :   If the withr package is installed, modified_by overrides the
      default value for the lastModifiedBy property assigned to the
      output file by officer. Defaults to `Sys.getenv("USER")` (the same
      value used by officer).

- overwrite:

  Defaults to `FALSE`.

- interactive:

  If `FALSE`, warn the user and return `x`.

## Value

Input object `x` without modification.
