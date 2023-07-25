#' Check if x is a rdocx, rpptx, or rxlsx object
#'
#' @param x Object to check.
#' @param arg Argument name of object to check. Used to improve
#'   [cli::cli_abort()] messages. Defaults to `caller_arg(x)`.
#' @param what Class names to check
#' @param ... Additional parameters passed to [cli::cli_abort()]
#' @inheritParams rlang::args_error_context
#' @export
check_officer <- function(x,
                          arg = caller_arg(x),
                          what = c("rdocx", "rpptx", "rxlsx"),
                          call = caller_env(),
                          ...) {
  check_required(x, arg = arg, call = call)
  check_character(what)
  what <- match.arg(what, several.ok = TRUE)
  if (inherits(x, what)) {
    return(invisible(NULL))
  }

  what <- cli_vec_cls(what)
  cli_abort("{.arg {arg}} must be a {what} object.", call = call, ...)
}

#' @name check_docx
#' @rdname check_officer
#' @export
check_docx <- function(x, arg = caller_arg(x), call = caller_env(), ...) {
  check_officer(x, what = "rdocx", arg = arg, call = call, ...)
}

#' @name check_pptx
#' @rdname check_officer
#' @export
check_pptx <- function(x, arg = caller_arg(x), call = caller_env(), ...) {
  check_officer(x, what = "rpptx", arg = arg, call = call, ...)
}

#' @name check_xlsx
#' @rdname check_officer
#' @export
check_xlsx <- function(x, arg = caller_arg(x), call = caller_env(), ...) {
  check_officer(x, what = "rxlsx", arg = arg, call = call, ...)
}

#' @name check_block_list
#' @rdname check_officer
#' @param allow_empty If `TRUE`, [check_block_list()] allows an empty block
#'   list.
#' @export
check_block_list <- function(x,
                             arg = caller_arg(x),
                             allow_empty = FALSE,
                             allow_null = FALSE,
                             call = caller_env()) {
  if (is_block_list(x)) {
    if (allow_empty || !is_empty(x)) {
      return(invisible(NULL))
    }

    cli::cli_abort("{.arg {arg}} can't be empty.", call = call)
  }

  stop_input_type(
    x,
    "block list",
    allow_null = allow_null,
    call = call
  )
}

#' Check file path for a docx, pptx, or xlsx file extension
#'
#' @inheritParams check_docx
#' @param fileext File extensions to allow without error. Defaults to "docx",
#'   "pptx", "xlsx".
#' @param arg Argument name of object to check. Used to improve
#'   [cli::cli_abort()] messages. Defaults to `caller_arg(x)`.
#' @inheritParams check_officer
#' @export
#' @importFrom rlang caller_arg current_env
#' @importFrom cli cli_abort
check_office_fileext <- function(x,
                                 arg = caller_arg(x),
                                 fileext = c("docx", "pptx", "xlsx"),
                                 call = caller_env(),
                                 ...) {
  check_required(x, call = call)
  check_character(fileext, allow_null = TRUE, call = call)
  fileext <- match.arg(fileext, several.ok = TRUE)

  if (!is_null(x) && !any(is_fileext_path(x, fileext))) {
    cli_abort(
      "{.arg {arg}} must use a {.val {cli_vec_last(fileext)}} file extension.",
      ...,
      call = call
    )
  }
}

#' @name check_docx_fileext
#' @rdname check_office_fileext
#' @export
check_docx_fileext <- function(x,
                               arg = caller_arg(x),
                               call = caller_env(),
                               ...) {
  check_office_fileext(x, arg, fileext = "docx", call = call, ...)
}

#' @name check_pptx_fileext
#' @rdname check_office_fileext
#' @export
#' @importFrom rlang caller_arg current_env
check_pptx_fileext <- function(x,
                               arg = caller_arg(x),
                               call = caller_env(),
                               ...) {
  check_office_fileext(x, arg, fileext = "pptx", call = call, ...)
}

#' @name check_xlsx_fileext
#' @rdname check_office_fileext
#' @export
#' @importFrom rlang caller_arg current_env
check_xlsx_fileext <- function(x,
                               arg = caller_arg(x),
                               call = caller_env(),
                               ...) {
  check_office_fileext(x, arg, fileext = "xlsx", call = call, ...)
}
