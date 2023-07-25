#' Add a list of blocks into a Word document or PowerPoint presentation
#'
#' @description
#' `r lifecycle::badge('experimental')`
#'
#' [officer_add_blocks()] supports adding a list of blocks to rdocx (using
#' [add_blocks_to_body()]) or rpptx objects (using [officer::ph_with()]).
#' add_blocks_to_body() is a variant of [officer::body_add_blocks()] that allows
#' users to set the cursor position before adding the block list using
#' [cursor_docx()].
#'
#' @name officer_add_blocks
#' @param x A rdocx or rpptx object. Required.
#' @param location If `NULL`, location defaults to [officer::ph_location()]
#' @inheritParams officer::body_add_blocks
#' @inheritParams officer::ph_with
#' @inheritParams check_officer
#' @export
officer_add_blocks <- function(x,
                               blocks,
                               pos = "after",
                               location = NULL,
                               level_list = integer(0),
                               ...,
                               call = caller_env()) {
  check_officer(x, what = c("rdocx", "rpptx"), call = call)

  check_block_list(blocks, call = call)

  if (is_rdocx(x)) {
    x <- add_blocks_to_body(
      docx = x,
      blocks = blocks,
      pos = pos,
      ...
    )

    return(x)
  }

  stopifnot(
    is_rpptx(x)
  )

  location <- location %||% officer::ph_location_type()

  officer::ph_with(
    x = x,
    value = blocks,
    location = location,
    level_list = level_list,
    ...
  )
}

#' @rdname officer_add_blocks
#' @param docx A rdocx object.
#' @inheritParams cursor_docx
#' @export
add_blocks_to_body <- function(docx,
                               blocks,
                               pos = "after",
                               keyword = NULL,
                               id = NULL,
                               index = NULL,
                               ...) {
  if (!is_all_null(c(keyword, id, index))) {
    docx <- cursor_docx(docx, keyword, id, index)
  }

  officer::body_add_blocks(
    docx,
    blocks = blocks,
    pos = pos
  )
}
