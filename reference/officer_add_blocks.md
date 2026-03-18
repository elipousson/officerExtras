# Add a list of blocks into a Word document or PowerPoint presentation

**\[experimental\]**

`officer_add_blocks()` supports adding a list of blocks to rdocx (using
`add_blocks_to_body()`) or rpptx objects (using
[`officer::ph_with()`](https://davidgohel.github.io/officer/reference/ph_with.html)).
`add_blocks_to_body()` is a variant of
[`officer::body_add_blocks()`](https://davidgohel.github.io/officer/reference/body_add_blocks.html)
that allows users to set the cursor position before adding the block
list using
[`cursor_docx()`](https://elipousson.github.io/officerExtras/reference/cursor_docx.md).

## Usage

``` r
officer_add_blocks(
  x,
  blocks,
  pos = "after",
  location = NULL,
  level_list = integer(0),
  ...,
  call = caller_env()
)

add_blocks_to_body(
  docx,
  blocks,
  pos = "after",
  keyword = NULL,
  id = NULL,
  index = NULL,
  ...
)
```

## Arguments

- x:

  A rdocx or rpptx object. Required.

- blocks:

  set of blocks to be used as footnote content returned by function
  [`block_list()`](https://davidgohel.github.io/officer/reference/block_list.html).

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".

- location:

  If `NULL`, location defaults to
  [`officer::ph_location_type()`](https://davidgohel.github.io/officer/reference/ph_location_type.html)

- level_list:

  The list of levels for hierarchy structure as integer values. If used
  the object is formated as an unordered list. If 1 and 2, item 1 level
  will be 1, item 2 level will be 2.

- ...:

  further arguments passed to or from other methods. When adding a
  `ggplot` object or `plot_instr`, these arguments will be used by the
  png function.

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

- docx:

  A rdocx object.

- keyword, id:

  A keyword string used to place cursor with
  [`officer::cursor_reach()`](https://davidgohel.github.io/officer/reference/cursor.html)
  or bookmark id with
  [`officer::cursor_bookmark()`](https://davidgohel.github.io/officer/reference/cursor.html).
  Defaults to `NULL`. If keyword or id are not provided, the gt object
  is inserted at the front of the document.

- index:

  A integer matching a doc_index value appearing in a summary of the
  docx object created with
  [`officer::docx_summary()`](https://davidgohel.github.io/officer/reference/docx_summary.html).
  If index is for a paragraph value, the text of the pargraph is used as
  a keyword.

## See also

Other block list functions:
[`make_block_list()`](https://elipousson.github.io/officerExtras/reference/make_block_list.md)
