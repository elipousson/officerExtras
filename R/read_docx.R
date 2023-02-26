#' Read a docx file or use a rdox object if provided
#'
#' @param filename,path File name and path. Default: `NULL`
#' @param docx rdocx class object, If docx is provided, filename and path are
#'   ignored. Default: `NULL`
#' @param quiet If `FALSE`, warn if docx is provided when filename and/or path
#'   are also provided. Default: `TRUE`.
#' @return A rdocx object.
#' @seealso
#'  [officer::read_docx()]
#' @rdname read_docx_ext
#' @export
#' @importFrom cli cli_alert_warning cli_alert_success symbol
#' @importFrom officer read_docx
read_docx_ext <- function(filename = NULL, path = NULL, docx = NULL, quiet = TRUE) {
  if (is.null(docx)) {
    docx <- officer::read_docx(path = set_docx_path(filename, path))
  } else {
    if (!is_all_null(c(filename, path)) & isFALSE(quiet)) {
      cli::cli_alert_warning(
        "{.arg filename} and {.arg path} are ignored if {.arg docx} is provided."
      )
    }

    check_docx(docx)
  }

  if (isFALSE(quiet)) {
    if (!is.null(filename)) {
      cli::cli_alert_success("Reading {.filename {filename}}{cli::symbol$ellipsis}")
    }
    list_doc_properties(docx, filename)
  }

  invisible(docx)
}

#' List document properties
#'
#' @keywords internal
#' @export
#' @importFrom officer doc_properties
#' @importFrom cli cli_rule symbol cli_dl
#' @importFrom rlang set_names
list_doc_properties <- function(docx, filename = NULL) {
  props <- officer::doc_properties(docx)
  if (!is.null(filename)) {
    cli::cli_rule("{cli::symbol$info} {.filename {filename}} properties:")
  } else {
    cli::cli_rule("{cli::symbol$info} document properties:")
  }

  cli::cli_dl(
    items = discard(
      rlang::set_names(props[["value"]], props[["tag"]]),
      ~ .x == ""
    )
  )
}

#' Set filepath for docx file
#'
#' @keywords internal
#' @importFrom cli cli_vec
set_docx_path <- function(filename = NULL, path = NULL) {
  if (is.null(path)) {
    path <- filename
  } else {
    path <- file.path(path, filename)
  }

  check_docx_fileext(
    path,
    arg = cli::cli_vec(
      c("filename", "path"),
      style = list("vec-last" = " or ")
    )
  )

  path
}
