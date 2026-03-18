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

- overwrite:

  Defaults to `FALSE`.

- interactive:

  If `FALSE`, warn the user and return `x`.

## Value

Input object `x` without modification.
