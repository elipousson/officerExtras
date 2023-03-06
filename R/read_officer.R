#' Read a docx, pptx, potx, or xlsx file or use an existing object from officer
#' if provided
#'
#' [read_docx_ext()], [read_pptx_ext()], and [read_xlsx_ext()] are wrappers for
#' [read_officer()] which is a wrapper for [officer::read_docx()],
#' [officer::read_pptx()], and [officer::read_xlsx()]. This function uses both a
#' filename and path (the original officer functions only use a path) and
#' default to use a rdocx, rpptx, or rxlsx class object if provided.
#'
#' @param filename,path File name and path. Default: `NULL`. Must include a
#'   "docx", "pptx", or "xlsx" file path. "dotx" and "potx" files are also
#'   supported.
#' @param x A rdocx, rpptx, or rxlsx class object, If docx is provided, filename
#'   and path are ignored. Default: `NULL`
#' @param docx,pptx,xlsx A rdocx, rpptx, or rxlsx class object passed to the x
#'   parameter by the corresponding function. Defaults to `NULL`.
#' @param quiet If `FALSE`, warn if docx is provided when filename and/or path
#'   are also provided. Default: `TRUE`.
#' @inheritParams check_office_fileext
#' @return A rdocx, rpptx, or rxlsx object.
#' @seealso
#'  [officer::read_docx()]
#' @rdname read_officer
#' @export
#' @importFrom cli cli_alert_warning cli_alert_success symbol
#' @importFrom rlang current_call
#' @importFrom officer read_docx
read_officer <- function(filename = NULL,
                         path = NULL,
                         fileext = c("docx", "pptx", "xlsx"),
                         x = NULL,
                         arg = caller_arg(x),
                         quiet = TRUE,
                         call = parent.frame(),
                         ...) {
  if (is.null(x)) {
    path <- set_office_path(filename, path, fileext = fileext, call = call)

    x <- switch(str_extract_fileext(path),
      "docx" = officer::read_docx(path),
      "pptx" = officer::read_pptx(path),
      "xlsx" = officer::read_xlsx(path)
    )
  } else {
    if (!is_all_null(c(filename, path)) & isFALSE(quiet)) {
      cli::cli_alert_warning(
        "{.arg filename} and {.arg path} are ignored if {.arg {arg}} is provided."
      )
    }

    check_officer(x, what = paste0("r", fileext), call = call, ...)
  }

  if (isFALSE(quiet)) {
    if (!is.null(filename)) {
      cli::cli_alert_success("Reading {.filename {filename}}{cli::symbol$ellipsis}")
    }

    cli_doc_properties(x, filename)
  }

  invisible(x)
}

#' @name read_docx_ext
#' @rdname read_officer
#' @export
read_docx_ext <- function(filename = NULL,
                          path = NULL,
                          docx = NULL,
                          quiet = TRUE) {
  read_officer(
    filename = filename,
    path = path,
    fileext = "docx",
    x = docx,
    quiet = quiet
  )
}

#' @name read_pptx_ext
#' @rdname read_officer
#' @export
read_pptx_ext <- function(filename = NULL,
                          path = NULL,
                          pptx = NULL,
                          quiet = TRUE) {
  read_officer(
    filename = filename,
    path = path,
    fileext = "pptx",
    x = pptx,
    quiet = quiet
  )
}

#' @name read_xlsx_ext
#' @rdname read_officer
#' @export
read_xlsx_ext <- function(filename = NULL,
                          path = NULL,
                          xlsx = NULL,
                          quiet = TRUE) {
  read_officer(
    filename = filename,
    path = path,
    fileext = "xlsx",
    x = xlsx,
    quiet = quiet
  )
}

#' List document properties for a rdocx or rpptx object
#'
#' @keywords internal
#' @export
#' @importFrom cli cli_rule symbol cli_dl
cli_doc_properties <- function(x, filename = NULL) {
  if (!is.null(filename)) {
    cli::cli_rule("{cli::symbol$info} {.filename {filename}} properties:")
  } else {
    cli::cli_rule("{cli::symbol$info} document properties:")
  }

  cli::cli_dl(
    items = discard(
      officer_properties(x),
      ~ .x == ""
    )
  )
}

#' Get doc properties for a rdocx or rpptx object as list
#'
#' @keywords internal
#' @export
#' @importFrom officer doc_properties
#' @importFrom rlang set_names
#' @importFrom utils modifyList
officer_properties <- function(x, values = list(), keep.null = FALSE) {
  props <- officer::doc_properties(x)
  utils::modifyList(
    rlang::set_names(as.list(props[["value"]]), props[["tag"]]),
    values,
    keep.null
  )
}

#' Set filepath for docx file
#'
#' @keywords internal
#' @importFrom cli cli_vec
set_office_path <- function(filename = NULL,
                            path = NULL,
                            fileext = c("docx", "pptx", "xlsx"),
                            call = parent.frame()) {
  if (is.null(path)) {
    path <- filename
  } else if (!is.null(filename)) {
    path <- file.path(path, filename)
  }

  if (!is.character(path) & !is.null(path)) {
    cli::cli_abort(
      '{.arg {cli_vec_last(c("filename", "path"))}} must be `NULL` or
      {.cls character} objects not {.cls {class(path)}}.'
    )
  }

  fileext <- match.arg(fileext, several.ok = TRUE)

  if (("pptx" %in% fileext) & is_fileext_path(path, "potx")) {
    return(path)
  } else if (("dotx" %in% fileext) & is_fileext_path(path, "dotx")) {
    return(path)
  }

  check_office_fileext(
    path,
    arg = cli_vec_last(
      c("filename", "path")
    ),
    fileext = fileext,
    call = call
  )

  path
}
