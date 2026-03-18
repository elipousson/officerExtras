# Make a list of blocks to add to a Word document or PowerPoint presentation

**\[experimental\]**

`make_block_list()` extends
[`officer::block_list()`](https://davidgohel.github.io/officer/reference/block_list.html)
by supporting a list of inputs and optionally combining parameters with
an existing block list (using the blocks parameter). Unlike
[`officer::block_list()`](https://davidgohel.github.io/officer/reference/block_list.html),
`make_block_list()` errors if no input parameters are provided.

`combine_blocks()` takes any number of `block_list` objects and combined
them into a single `block_list`. Both functions are not yet working as
expected.

## Usage

``` r
make_block_list(blocks = NULL, ..., allow_empty = FALSE, call = caller_env())

combine_blocks(...)
```

## Arguments

- blocks:

  A list of parameters to pass to
  [`officer::block_list()`](https://davidgohel.github.io/officer/reference/block_list.html)
  or a `block_list` object. If parameters are provided to both `...` and
  blocks is a `block_list`, the additional parameters are appended to
  the end of blocks.

- ...:

  For `make_block_list()`, these parameters are passed to
  [`officer::block_list()`](https://davidgohel.github.io/officer/reference/block_list.html)
  and must *not* include `block_list` objects. For `combine_blocks()`,
  these parameters must all be `block_list` objects.

- allow_empty:

  If `TRUE`,
  [`check_block_list()`](https://elipousson.github.io/officerExtras/reference/check_officer.md)
  allows an empty block list. Defaults to `FALSE`.

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

## See also

Other block list functions:
[`officer_add_blocks()`](https://elipousson.github.io/officerExtras/reference/officer_add_blocks.md)
