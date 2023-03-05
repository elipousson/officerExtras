#' @keywords internal
#' @noRd
is_officer <- function(x, what = c("rdocx", "rpptx", "rxlsx")) {
  inherits(x, what)
}

#' Check if x is a rdocx, rpptx, or rxlsx object
#'
#' @param x Object to check.
#' @param arg Argument name for x. Defaults to `caller_arg(x)`.
#' @param what Class names to check
#' @param ... Additional parameters passed to [cli::cli_abort()]
#' @export
#' @importFrom rlang caller_arg
#' @importFrom cli cli_abort
check_officer <- function(x,
                          arg = caller_arg(x),
                          what = c("rdocx", "rpptx", "rxlsx"),
                          ...) {
  rlang::check_required(x, arg = arg)
  what <- match.arg(what, several.ok = TRUE)
  if (!inherits(x, what)) {
    what <- cli_vec_cls(what)

    cli_abort("{.arg {arg}} must be a {what} object.", ...)
  }
}

#' @name check_docx
#' @rdname check_officer
#' @export
check_docx <- function(x, arg = caller_arg(x), ...) {
  check_officer(x, what = "rdocx", ...)
}

#' @name check_pptx
#' @rdname check_officer
#' @export
check_pptx <- function(x, arg = caller_arg(x), ...) {
  check_officer(x, what = "rpptx", ...)
}

#' @name check_xlsx
#' @rdname check_officer
#' @export
check_xlsx <- function(x, arg = caller_arg(x), ...) {
  check_officer(x, what = "rxlsx", ...)
}

#' Check file path for a docx, pptx, or xlsx file extension
#'
#' @inheritParams check_docx
#' @param fileext File extensions to allow without error. Defaults to "docx",
#'   "pptx", "xlsx".
#' @param arg Argument name for x. Used to improve error messages with
#'   [cli::cli_abort()].
#' @param .envir Passed to [cli::cli_abort()].
#' @export
#' @importFrom rlang caller_arg current_env
#' @importFrom cli cli_abort
check_office_fileext <- function(x,
                                 arg = caller_arg(x),
                                 ...,
                                 fileext = c("docx", "pptx", "xlsx"),
                                 .envir = current_env()) {
  fileext <- match.arg(fileext, several.ok = TRUE)

  if (!is.null(x) & !any(is_fileext_path(x, fileext))) {
    cli_abort(
      "{.arg {arg}} must use a {.val {cli_vec_last(fileext)}} file extension.",
      ...,
      .envir = .envir
    )
  }
}

#' @name check_docx_fileext
#' @rdname check_office_fileext
#' @export
check_docx_fileext <- function(x,
                               arg = caller_arg(x),
                               ...,
                               .envir = current_env()) {
  check_office_fileext(x, arg, ..., fileext = "docx", .envir = .envir)
}

#' @name check_pptx_fileext
#' @rdname check_office_fileext
#' @export
#' @importFrom rlang caller_arg current_env
check_pptx_fileext <- function(x,
                               arg = caller_arg(x),
                               ...,
                               .envir = current_env()) {
  check_office_fileext(x, arg, ..., fileext = "pptx", .envir = .envir)
}

#' @name check_xlsx_fileext
#' @rdname check_office_fileext
#' @export
#' @importFrom rlang caller_arg current_env
check_xlsx_fileext <- function(x,
                               arg = caller_arg(x),
                               ...,
                               .envir = current_env()) {
  check_office_fileext(x, arg, ..., fileext = "xlsx", .envir = .envir)
}
