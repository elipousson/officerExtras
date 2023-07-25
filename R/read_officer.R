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
#' @param x A rdocx, rpptx, or rxlsx class object If x is provided, filename and
#'   path are ignored. Default: `NULL`
#' @param docx,pptx,xlsx A rdocx, rpptx, or rxlsx class object passed to the x
#'   parameter of [read_officer()] by the variant functions. Defaults to `NULL`.
#' @param allow_null If `TRUE`, function supports the default behavior of
#'   [officer::read_docx()], [officer::read_pptx()], or [officer::read_xlsx()]
#'   and returns an empty document if x, filename, and path are all `NULL`. If
#'   `FALSE`, one of the three parameters must be supplied.
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
                         allow_null = TRUE,
                         quiet = TRUE,
                         call = parent.frame(),
                         ...) {
  cli_quiet(quiet)

  has_input_file <- !is_null(c(filename, path))

  if (is.null(x)) {
    if (has_input_file || !allow_null) {
      path <- set_office_path(filename, path, fileext = fileext, call = call)
      filename <- basename(path)
      fileext <- str_extract_fileext(path)
    } else {
      fileext <- match.arg(fileext)
      new_obj <- switch(fileext,
        "docx" = "empty document",
        "pptx" = "pptx document with 0 slides",
        "xlsx" = "xlsx document with 1 sheet"
      )

      cli::cli_alert_success("Creating a new {new_obj}")
    }

    x <- rlang::try_fetch(
      switch(fileext,
        "docx" = officer::read_docx(path),
        "dotx" = officer::read_docx(path),
        "pptx" = officer::read_pptx(path),
        "potx" = officer::read_pptx(path),
        "xlsx" = officer::read_xlsx(path)
      ),
      error = function(cnd) {
        cli::cli_abort("{.val {fileext}} file can't be read.", parent = cnd)
      },
      warning = function(cnd) {
        cli::cli_warn(message = cnd)
      }
    )
  } else {
    if (has_input_file) {
      cli::cli_alert_warning(
        "{.arg filename} and {.arg path} are ignored if {.arg {arg}} is provided."
      )
    }

    check_officer(x, what = paste0("r", fileext), call = call, ...)
  }

  if (!is.null(filename)) {
    cli::cli_alert_success(
      "Reading {.filename {filename}}{cli::symbol$ellipsis}"
    )
  }

  cli_doc_properties(x, filename)

  invisible(x)
}

#' @name read_docx_ext
#' @rdname read_officer
#' @export
read_docx_ext <- function(filename = NULL,
                          path = NULL,
                          docx = NULL,
                          allow_null = FALSE,
                          quiet = TRUE) {
  read_officer(
    filename = filename,
    path = path,
    fileext = "docx",
    x = docx,
    allow_null = allow_null,
    quiet = quiet
  )
}

#' @name read_pptx_ext
#' @rdname read_officer
#' @export
read_pptx_ext <- function(filename = NULL,
                          path = NULL,
                          pptx = NULL,
                          allow_null = FALSE,
                          quiet = TRUE) {
  read_officer(
    filename = filename,
    path = path,
    fileext = "pptx",
    x = pptx,
    allow_null = allow_null,
    quiet = quiet
  )
}

#' @name read_xlsx_ext
#' @rdname read_officer
#' @export
read_xlsx_ext <- function(filename = NULL,
                          path = NULL,
                          xlsx = NULL,
                          allow_null = FALSE,
                          quiet = TRUE) {
  read_officer(
    filename = filename,
    path = path,
    fileext = "xlsx",
    x = xlsx,
    allow_null = allow_null,
    quiet = quiet
  )
}

#' List document properties for a rdocx or rpptx object
#'
#' @keywords internal
#' @noRd
#' @importFrom cli cli_rule symbol cli_dl
cli_doc_properties <- function(x, filename = NULL) {
  props <- officer_properties(x)

  if (is_null(props)) {
    return(props)
  }

  msg <- "{cli::symbol$info} document properties:"

  if (!is.null(filename)) {
    msg <- "{cli::symbol$info} {.filename {filename}} properties:"
  }

  cli::cli_rule(msg)

  cli::cli_dl(
    items = discard(props, function(x) {
      x == ""
    })
  )
}

#' Get doc properties for a rdocx or rpptx object as a list
#'
#' @keywords internal
#' @noRd
#' @importFrom officer doc_properties
#' @importFrom rlang set_names
#' @importFrom utils modifyList
#' @importFrom cli cli_warn
officer_properties <- function(x, values = list(), keep.null = FALSE) {
  props <- rlang::try_fetch(
    officer::doc_properties(x),
    error = function(cnd) {
      cli::cli_warn(
        "Document properties can't be found for {.filename {x}}",
        parent = cnd
      )
      NULL
    }
  )

  if (is_null(props)) {
    return(props)
  }

  utils::modifyList(
    rlang::set_names(as.list(props[["value"]]), props[["tag"]]),
    values,
    keep.null
  )
}

#' Set filepath for docx file
#'
#' @keywords internal
#' @noRd
#' @importFrom cli cli_vec
set_office_path <- function(filename = NULL,
                            path = NULL,
                            fileext = c("docx", "pptx", "xlsx"),
                            call = parent.frame()) {
  check_string(filename, allow_null = TRUE, call = call)
  check_string(path, allow_null = TRUE, call = call)

  if (is.null(path)) {
    if (is.null(filename)) {
      args <- c("filename", "path")
      cli::cli_abort("{.arg {args}} can't both be {.code NULL}")
    }
    path <- filename
  } else if (!is.null(filename)) {
    path <- file.path(path, filename)
  }

  fileext <- match.arg(fileext, several.ok = TRUE)

  if ((("pptx" %in% fileext) && is_fileext_path(path, "potx")) ||
    (("dotx" %in% fileext) && is_fileext_path(path, "dotx"))) {
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
