# Reduce a list to a single officer object

`reduce_officer()` is a wrapper for
[`purrr::reduce()`](https://purrr.tidyverse.org/reference/reduce.html).

## Usage

``` r
reduce_officer(
  x = NULL,
  .f = function(x, value, ...) {
     vec_add_to_body(x, value = value, ...)
 },
  value = NULL,
  ...,
  .path = NULL
)
```

## Arguments

- x:

  File path or officer object.

- .f:

  Any function taking an officer object as the first parameter and a
  value as the second parameter, Defaults to anonymous function:
  `function(x, value, ...) { vec_add_to_body(x, value = value, ...) }`

- value:

  A vector of values that are support by the function passed to .f,
  Default: `NULL`

- ...:

  Additional parameters passed to
  [`purrr::reduce()`](https://purrr.tidyverse.org/reference/reduce.html).

- .path:

  If .path not `NULL`, it should be a file path that is passed to
  [`write_officer()`](https://elipousson.github.io/officerExtras/reference/write_officer.md),
  allowing you to modify a docx file and write it back to a file in a
  single function call.

## Value

A rdocx, rpptx, or rxlsx object.

## See also

[`purrr::reduce()`](https://purrr.tidyverse.org/reference/reduce.html),
[`vec_add_to_body()`](https://elipousson.github.io/officerExtras/reference/vec_add_to_body.md)

## Examples

``` r
if (FALSE) { # \dontrun{
if (interactive()) {
  x <- reduce_officer(value = LETTERS)

  officer_summary(x)
}
} # }
```
