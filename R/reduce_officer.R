#' Reduce a list to a single officer object
#'
#' [reduce_officer()] is a wrapper for [purrr::reduce()].
#'
#' @param x File path or officer object.
#' @param .f Any function taking an officer object as the first parameter and a
#'  value as the second parameter, Defaults to anonymous function:
#'  `function(x, value, ...) {
#'  vec_add_to_body(x, value = value, ...) }`
#' @param value A vector of values that are support by the function passed to
#'   .f, Default: `NULL`
#' @param ... Additional parameters passed to [purrr::reduce()].
#' @param .path If .path not `NULL`, it should be a file path that
#'   is passed to [write_officer()], allowing you to modify a docx file and
#'   write it back to a file in a single function call.
#' @returns A rdocx, rpptx, or rxlsx object.
#' @examples
#' \dontrun{
#' if(interactive()){
#'  x <- reduce_officer(value = LETTERS)
#'
#'  officer_summary(x)
#'  }
#' }
#' @seealso
#'  [purrr::reduce()], [vec_add_to_body()]
#' @rdname reduce_officer
#' @export
reduce_officer <- function(x = NULL,
                           .f = \(x, value, ...) {
                             vec_add_to_body(x, value = value, ...)
                           },
                           value = NULL,
                           ...,
                           .path = NULL) {
  if (!is_officer(x)) {
    x <- read_officer(x)
  }

  x <- reduce(
    value,
    .f = .f,
    ...,
    .init = x
  )

  if (is.null(.write_path)) {
    return(x)
  }

  write_officer(x, path = .path)
}
