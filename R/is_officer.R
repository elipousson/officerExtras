#' Is this object an officer class object?
#'
#' [is_officer()] and variants are basic wrappers for [inherits()] to check for
#' object classes created by functions from the officer package. These functions
#' are intended for internal use but are exposed for use by other extension
#' developers.
#'
#' @rdname is_officer
#' @export
is_officer <- function(x, what = c("rdocx", "rpptx", "rxlsx")) {
  inherits(x, what)
}

#' @rdname is_officer
#' @export
is_rdocx <- function(x) {
  inherits(x, "rdocx")
}

#' @rdname is_officer
#' @export
is_rpptx <- function(x) {
  inherits(x, "rpptx")
}

#' @rdname is_officer
#' @export
is_block_list <- function(x) {
  inherits_all(x, c("block_list", "block"))
}
