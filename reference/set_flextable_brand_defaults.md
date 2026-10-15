# Set flextable defaults from a brand.yml file

`set_flextable_brand_defaults()` sets
[`flextable::set_flextable_defaults()`](https://davidgohel.github.io/flextable/reference/set_flextable_defaults.html)
from a brand.yml file, so tables created with flextable match the
`Table` style in a reference document created with
[`brand_reference_docx()`](https://elipousson.github.io/officerExtras/reference/brand_reference_docx.md).
flextable formats each table directly (rather than using the reference
document's `Table` style), so the defaults are needed for tables to
match.

## Usage

``` r
set_flextable_brand_defaults(brand = NULL, theme_fun = NULL, ...)
```

## Arguments

- brand:

  Path to a brand.yml file or a directory containing one, a `brand_yml`
  object from
  [`brand.yml::read_brand_yml()`](https://rdrr.io/pkg/brand.yml/man/read_brand_yml.html),
  or a list with the brand settings (e.g. from
  [`yaml12::read_yaml()`](https://posit-dev.github.io/r-yaml12/reference/parse_yaml.html)).
  If `NULL` (default), the path is taken from the `BRAND_YML_PATH`
  environment variable or, if that is not set, found by searching the
  working directory and its parent directories for `_brand.yml` or
  `_brand.yaml`, either directly or in a `brand` or `_brand`
  subdirectory (the same search used by Quarto and
  [`brand.yml::read_brand_yml()`](https://rdrr.io/pkg/brand.yml/man/read_brand_yml.html)).

- theme_fun:

  A flextable theme function applied to each new table, before the
  `header` settings. Defaults to
  [`flextable::theme_box()`](https://davidgohel.github.io/flextable/reference/theme_box.html)
  (borders on every cell, as in the reference document `Table` style).

- ...:

  Additional arguments passed to
  [`flextable::set_flextable_defaults()`](https://davidgohel.github.io/flextable/reference/set_flextable_defaults.html).
  These take precedence over the brand settings.

## Value

The previous flextable defaults, invisibly. Pass them to
`do.call(flextable::set_flextable_defaults, old)` to restore them.

## Details

The following brand.yml settings are used:

- `typography.base`: `family`, `size`, and `line-height` (if unitless)
  for the table text.

- `color.foreground` for the text color.

- `defaults.docx.styles.Table`: `border-color`, `border-width`,
  `cell-padding`, and `header` (`family`, `size`, `weight`, `style`,
  `color`, and `background-color` for the header row).

## Examples

``` r
if (rlang::is_installed(c("flextable", "yaml12"))) {
  brand_path <- system.file("extdata", "_brand.yml", package = "officerExtras")

  old <- set_flextable_brand_defaults(brand_path)
  flextable::flextable(head(mtcars, 3))

  do.call(flextable::set_flextable_defaults, old)
}
```
