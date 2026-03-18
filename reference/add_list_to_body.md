# Add a formatted list of items to a `rdocx` object

This function assumes that the input docx object has a list style. Use
[`read_docx_ext()`](https://elipousson.github.io/officerExtras/reference/read_officer.md)
with `allow_null = TRUE` to create a new rdocx object with the "List
Bullet" and "List Number" styles. "List Paragraph" can be used but the
resulting formatted list will not have bullets or numbers.

## Usage

``` r
add_list_to_body(
  docx,
  values,
  style = "List Bullet",
  keep_na = FALSE,
  before = NULL,
  after = "",
  ...
)
```

## Arguments

- docx:

  A `rdocx` object.

- values:

  Character vector with values for list.

- style:

  Style to use for list of values, Default: 'List Bullet'

- keep_na:

  If `TRUE`, keep `NA` values in values Default: FALSE

- before:

  Value to insert before the list, Default: NULL

- after:

  Value to insert after the list, Default: ”

- ...:

  Additional parameters passed to
  [`cursor_docx()`](https://elipousson.github.io/officerExtras/reference/cursor_docx.md)

## Value

A modified `rdocx` object.

## See also

[`add_to_body()`](https://elipousson.github.io/officerExtras/reference/add_to_body.md),
[`vec_add_to_body()`](https://elipousson.github.io/officerExtras/reference/vec_add_to_body.md)
