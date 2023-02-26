#' Check if x is a rdocx object or if x has a docx file extension
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
                          what = c("rdocx", "rpptx"),
                          ...) {
  rlang::check_required(x, arg = arg)
  if (!inherits(x, what)) {
    what <-
      cli_vec_last(
        what,
        style = list(
          before = "<",
          after = ">",
          color = "blue"
        )
      )

    cli_abort("{.arg {arg}} must be a {what} object.", ...)
  }
}

#' @name check_docx
#' @rdname check_officer
#' @export
check_docx <- function(x, arg = caller_arg(x), ...) {
  check_officer(x, what = "rdocx")
}

#' Check file path for a docx or pptx file extension
#'
#' @inheritParams check_docx
#' @export
#' @importFrom rlang caller_arg current_env
#' @importFrom cli cli_abort
check_office_fileext <- function(x,
                                 arg = caller_arg(x),
                                 ...,
                                 fileext = c("docx", "pptx", "xlsx"),
                                 .envir = current_env()) {
  fileext <- match.arg(fileext, several.ok = TRUE)
  if (!is.null(x) & !all(is_fileext_path(x, fileext))) {
    cli_abort(
      "{.arg {arg}} must use a {.val docx} file extension.",
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