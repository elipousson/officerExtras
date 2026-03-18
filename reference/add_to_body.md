# Add a xml string, text paragraph, or gt object at a specified position in a rdocx object

Wrappers for
[`officer::body_add_par()`](https://davidgohel.github.io/officer/reference/body_add_par.html),
[`officer::body_add_gg()`](https://davidgohel.github.io/officer/reference/body_add_gg.html),
and
[`officer::body_add_xml()`](https://davidgohel.github.io/officer/reference/body_add_xml.html)
that use the
[`cursor_docx()`](https://elipousson.github.io/officerExtras/reference/cursor_docx.md)
helper function to allow users to pass the value and keyword, id, or
index value used to place a "cursor" within the document using a single
function. If `pos = NULL`, `add_to_body()` calls
[`officer::body_add()`](https://davidgohel.github.io/officer/reference/body_add.html)
instead of
[`officer::body_add_par()`](https://davidgohel.github.io/officer/reference/body_add_par.html).
If value is a `gt_tbl` object, value is passed as the gt_object
parameter for `add_gt_to_body()`.

- `add_text_to_body()` passes value to
  [`glue::glue()`](https://glue.tidyverse.org/reference/glue.html) to
  add support for glue string interpolation.

- `add_gt_to_body()` converts gt tables to OOXML with
  [`gt::as_word()`](https://gt.rstudio.com/reference/as_word.html).

- `add_gg_to_body()` adds a caption following the plots using the labels
  from the plot object.

## Usage

``` r
add_to_body(
  docx,
  keyword = NULL,
  id = NULL,
  index = NULL,
  value = NULL,
  str = NULL,
  style = NULL,
  pos = "after",
  ...,
  gt_object = NULL,
  call = caller_env()
)

add_text_to_body(
  docx,
  value,
  style = NULL,
  pos = "after",
  .na = "NA",
  .null = NULL,
  .envir = parent.frame(),
  ...
)

add_xml_to_body(docx, str, pos = "after", ..., call = caller_env())

add_gt_to_body(
  docx,
  gt_object,
  align = "center",
  caption_location = c("top", "bottom", "embed"),
  caption_align = "left",
  split = FALSE,
  keep_with_next = TRUE,
  pos = "after",
  tablecontainer = TRUE,
  ...,
  call = caller_env()
)

add_gg_to_body(
  docx,
  value,
  caption = "title",
  caption_style = style,
  autonum = NULL,
  style = "Normal",
  pos = "after",
  ...
)

add_value_with_keys(docx, value, ..., .f = add_text_to_body)

add_str_with_keys(docx, str, ..., .f = add_xml_to_body)
```

## Arguments

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

- value:

  object to add in the document. Supported objects are vectors,
  data.frame, graphics, block of formatted paragraphs, unordered list of
  formatted paragraphs, pretty tables with package flextable,
  'Microsoft' charts with package mschart.

- str:

  a wml string

- style:

  paragraph style name. These names are available with function
  [styles_info](https://davidgohel.github.io/officer/reference/styles_info.html)
  and are the names of the Word styles defined in the base document (see
  argument `path` from
  [read_docx](https://davidgohel.github.io/officer/reference/read_docx.html)).

- pos:

  where to add the new element relative to the cursor, one of "after",
  "before", "on".

- ...:

  Additional parameters passed to
  [`officer::body_add_par()`](https://davidgohel.github.io/officer/reference/body_add_par.html),
  [`officer::body_add_gg()`](https://davidgohel.github.io/officer/reference/body_add_gg.html),
  or
  [`officer::body_add()`](https://davidgohel.github.io/officer/reference/body_add.html).

- gt_object:

  A gt object converted to an OOXML string with
  [`gt::as_word()`](https://gt.rstudio.com/reference/as_word.html) then
  passed to `add_xml_to_body()` as str parameter. Required for
  `add_gt_to_body()`.

- call:

  The execution environment of a currently running function, e.g.
  `caller_env()`. The function will be mentioned in error messages as
  the source of the error. See the `call` argument of
  [`abort()`](https://rlang.r-lib.org/reference/abort.html) for more
  information.

- .na:

  \[`character(1)`: ‘NA’\]  
  Value to replace `NA` values with. If `NULL` missing values are
  propagated, that is an `NA` result will cause `NA` output. Otherwise
  the value is replaced by the value of `.na`.

- .null:

  \[`character(1)`: ‘character()’\]  
  Value to replace NULL values with. If
  [`character()`](https://rdrr.io/r/base/character.html) whole output is
  [`character()`](https://rdrr.io/r/base/character.html). If `NULL` all
  NULL values are dropped (as in
  [`paste0()`](https://rdrr.io/r/base/paste.html)). Otherwise the value
  is replaced by the value of `.null`.

- .envir:

  \[`environment`:
  [`parent.frame()`](https://rdrr.io/r/base/sys.parent.html)\]  
  Environment to evaluate each expression in. Expressions are evaluated
  from left to right. If `.x` is an environment, the expressions are
  evaluated in that environment and `.envir` is ignored. If `NULL` is
  passed, it is equivalent to
  [`emptyenv()`](https://rdrr.io/r/base/environment.html).

- align:

  *Table alignment*

  `scalar<character>` // *default:* `"center"`

  An option for table alignment. Can either be `"center"`, `"left"`, or
  `"right"`.

- caption_location:

  *Caption location*

  `singl-kw:[top|bottom|embed]` // *default:* `"top"`

  Determines where the caption should be positioned. This can either be
  `"top"`, `"bottom"`, or `"embed"`.

- caption_align:

  *Caption alignment*

  Determines the alignment of the caption. This is either `"left"` (the
  default), `"center"`, or `"right"`. This option is only used when
  `caption_location` is not set as `"embed"`.

- split:

  *Allow splitting of a table row across pages*

  `scalar<logical>` // *default:* `FALSE`

  A logical value that indicates whether to activate the Word option
  `Allow row to break across pages`.

- keep_with_next:

  *Keeping rows together*

  `scalar<logical>` // *default:* `TRUE`

  A logical value that indicates whether a table should use Word option
  `Keep rows together`.

- tablecontainer:

  If `TRUE` (default), add tables inside of a tablecontainer tag that
  automatically adds a table number and converts the gt title into a
  table caption. This feature is based on code from the [gto
  package](https://github.com/GSK-Biostatistics/gto/) by Ellis Hughes to
  transform the gt_object to OOXML and insert the XML into the docx
  object.

- caption:

  Name of the ggplot2 label to use as a caption if plot passed to value
  has a label for this value. Defaults to "title".

- caption_style:

  Passed to style for
  [`officer::body_add_caption()`](https://davidgohel.github.io/officer/reference/body_add_caption.html).
  Defaults to same value as style.

- autonum:

  Automatic Table Numbering

  `scalar<logical>` // *default:* `TRUE`

  A logical value that indicates whether a table should use Words
  built-in auto table numbering option in the caption.
  `Automatic Table Numbering`.

- .f:

  Any function that takes a docx and value parameter and returns a rdocx
  object. A keyword parameter must also be supported if named is TRUE.
  Defaults to `add_text_to_body()`.

## Value

A rdocx object with xml, gt tables, or paragraphs of text added.

## Details

Using `add_value_with_keys()` or `add_str_with_keys()`

`add_value_with_keys()` supports value vectors of length 1 or longer. If
value is named, the names are assumed to be keywords indicating the
cursor position for adding each value in the vector. If value is not
named, a keyword parameter with the same length as value must be
provided. When `named = FALSE`, no keyword parameter is required. Add
`add_str_with_keys()` works identically but uses a str parameter and .f
defaults to `add_xml_to_body()`.

Note, as of July 2023, both `add_value_with_keys()` and
`add_str_with_keys()` are superseded by
[`vec_add_to_body()`](https://elipousson.github.io/officerExtras/reference/vec_add_to_body.md).

## Author

Ellis Hughes <ellis.h.hughes@gsk.com>
([ORCID](https://orcid.org/0000-0003-0637-4436))

## Examples

``` r
if (rlang::is_installed("gt")) {
  docx_example <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  tab_1 <-
    gt::gt(
      gt::exibble,
      rowname_col = "row",
      groupname_col = "group"
    )

  add_gt_to_body(
    docx_example,
    tab_1,
    keyword = "Title 1"
  )

  tab_str <- gt::as_word(tab_1)

  add_str_with_keys(
    docx_example,
    str = c("Title 1" = tab_str, "Title 2" = tab_str)
  )
}
#> rdocx document with 22 element(s)
#> 
#> * styles:
#>                 Normal              heading 1              heading 2 
#>            "paragraph"            "paragraph"            "paragraph" 
#> Default Paragraph Font           Normal Table                No List 
#>            "character"                "table"            "numbering" 
#>         Heading 1 Char         Heading 2 Char          Light Shading 
#>            "character"            "character"                "table" 
#>         List Paragraph                 header            Header Char 
#>            "paragraph"            "paragraph"            "character" 
#>                 footer            Footer Char 
#>            "paragraph"            "character" 
```
