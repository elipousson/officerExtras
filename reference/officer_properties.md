# Get doc properties for a rdocx or rpptx object as a list

`officer_properties()` is a variant on
[`officer::doc_properties()`](https://davidgohel.github.io/officer/reference/doc_properties.html)
that will warn instead of error if document properties can't be found

## Usage

``` r
officer_properties(x, values = list(), keep.null = FALSE, call = caller_env())
```

## Arguments

- x:

  A rdocx or rpptx object.

- values:

  A named list with new properties to replace existing document
  properties before they are returned as a named list.

- keep.null:

  Passed to
  [`utils::modifyList()`](https://rdrr.io/r/utils/modifyList.html). If
  `TRUE`, retain properties in returned list even if they have `NULL`
  values.

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

## Value

A named list of existing document properties or (if values is supplied)
modified document properties.
