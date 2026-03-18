# Is this object an officer class object?

`is_officer()` and variants are basic wrappers for
[`inherits()`](https://rdrr.io/r/base/class.html) to check for object
classes created by functions from the officer package. These functions
are intended for internal use but are exposed for use by other extension
developers.

## Usage

``` r
is_officer(x, what = c("rdocx", "rpptx", "rxlsx"))

is_rdocx(x)

is_rpptx(x)

is_block_list(x)
```

## Arguments

- x:

  A object to test

- what:

  Class or classes passed to
  [`inherits()`](https://rdrr.io/r/base/class.html)
