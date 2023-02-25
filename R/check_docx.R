#' Check if x is a rdocx object or if x has a docx file extension
#'
#' @param x Object to check.
#' @param arg Argument name for x. Defaults to `caller_arg(x)`.
#' @param ... Additional parameters passed to [cli::cli_abort()]
#' @inheritParams cli::cli_abort
#' @export
#' @importFrom rlang caller_arg
#' @importFrom cli cli_abort
check_docx <- function(x, arg = caller_arg(x), ...) {
  rlang::check_required(x, arg = arg)
  if (!inherits(x, "rdocx")) {
    cli_abort("{.arg {arg}} must be a {.cls rdocx} object.", ...)
  }
}

#' @name check_docx_fileext
#' @rdname check_docx
#' @export
#' @importFrom rlang caller_arg current_env
#' @importFrom cli cli_abort
check_docx_fileext <- function(x, arg = caller_arg(x), ..., .envir = current_env()) {
  if (!has_docx_fileext(x)) {
    cli_abort(
      "{.arg {arg}} must use a {.val docx} file extension.",
      ...,
      .envir = .envir
    )
  }
}

#' Does x have a docx file extension?
#'
#' @inheritParams isstatic::has_fileext
#' @keywords internal
has_docx_fileext <- function(string = NULL, ignore.case = TRUE) {
  has_fileext(string, "docx")
}
