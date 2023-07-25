#' Make a list of blocks to add to a Word document or PowerPoint presentation
#'
#' @description
#' `r lifecycle::badge('experimental')`
#'
#' `make_block_list()` extends [officer::block_list()] by supporting a list of
#' inputs and optionally combining parameters with an existing block list (using
#' the blocks parameter). Unlike [officer::block_list()], [make_block_list()]
#' errors if no input parameters are provided.
#'
#' `combine_blocks()` takes any number of `block_list` objects and combined them
#' into a single `block_list`. Both functions are not yet working as expected.
#'
#' @param ... For [make_block_list()], these parameters are passed to
#'   [officer::block_list()] and must *not* include `block_list` objects. For
#'   [combine_blocks()], these parameters must all be `block_list` objects.
#' @param blocks A list of parameters to pass to [officer::block_list()] or a
#'   `block_list` object. If parameters are provided to both `...` and blocks is
#'   a `block_list`, the additional parameters are appended to the end of
#'   blocks.
#' @inheritParams check_block_list
#' @family block list functions
#' @export
make_block_list <- function(blocks = NULL,
                            ...,
                            allow_empty = FALSE,
                            call = caller_env()) {
  has_blocks <- ...length() > 0

  if (is.null(blocks)) {
    if (!allow_empty && !has_blocks) {
      cli::cli_abort("{.arg blocks} or {.arg ...} must be supplied.")
    }
    has_blocks <- FALSE
    blocks <- officer::block_list(...)
  } else if (!is_block_list(blocks)) {
    stopifnot(is.list(blocks) && (allow_empty || !is_empty(blocks)))
    blocks <- exec(officer::block_list, !!!blocks)
  }

  stopifnot(
    is_block_list(blocks)
  )

  if (has_blocks) {
    blocks <- combine_blocks(blocks, officer::block_list(...))
  }

  blocks
}

#' @rdname make_block_list
#' @name combine_blocks
#' @export
combine_blocks <- function(...) {
  params <- list2(...)

  stopifnot(
    all(vapply(params, is_block_list, TRUE))
  )

  blocks <- c(...)

  class(blocks) <- class(params[[1]])

  blocks
}
