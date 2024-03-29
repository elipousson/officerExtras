#' Summarize a rdocx or rpptx object
#'
#' `officer_summary()` extends `officer::docx_summary()` and other officer
#' summary functions by handling multiple input types within a single function.
#' The preserve parameter is supported by officer version >= 0.6.3 (currently
#' the development version) and it is ignored unless a minimum supported version
#' of officer is installed.
#'
#' @param x A rdocx or rpptx object passed to [officer::docx_summary()],
#'   [officer::pptx_summary()], [officer::slide_summary()], or
#'   [officer::layout_summary()]. If x is a data frame created with one of those
#'   functions, it is returned as is (this feature may be removed in the future).
#' @inheritParams check_officer_summary
#' @inheritParams officer::slide_summary
#' @inheritParams officer::docx_summary
#' @param as_tibble If `TRUE` (default), return a tibble data frame.
#' @returns A tibble or data frame object.
#' @family summary functions
#' @examples
#' docx_example <- read_officer(
#'   system.file("doc_examples", "example.docx", package = "officer")
#' )
#'
#' officer_summary(docx_example)
#'
#' pptx_example <- read_officer(
#'   "example.pptx", system.file("doc_examples", package = "officer")
#' )
#'
#' officer_summary(pptx_example)
#'
#' officer_summary(pptx_example, "slide", 1)
#'
#' @export
#' @importFrom officer docx_summary pptx_summary slide_summary layout_summary
officer_summary <- function(x,
                            summary_type = "doc",
                            index = NULL,
                            preserve = FALSE,
                            as_tibble = TRUE,
                            call = caller_env()) {
  # FIXME: This feature should likely be removed.
  if (is_officer_summary(x, summary_type, call = call)) {
    if (as_tibble) {
      return(tibble::as_tibble(x))
    }

    return(x)
  }

  what <- switch(summary_type,
    "docx" = "rdocx",
    "pptx" = "rpptx",
    "slide" = "rpptx",
    "layout" = "rpptx",
    c("rdocx", "rpptx")
  )

  check_officer(x, what = what, call = call)

  if (identical(summary_type, "doc")) {
    summary_type <- NULL
  }

  summary_type <- summary_type %||% class(x)

  if (is_installed("officer (>= 0.6.3)") &&
    (summary_type %in% c("rdocx", "docx", "rpptx", "pptx"))) {
    if (summary_type %in% c("rdocx", "docx")) {
      summary_df <- officer::docx_summary(x, preserve = preserve)
    }

    if (summary_type %in% c("rpptx", "pptx")) {
      summary_df <- officer::pptx_summary(x, preserve = preserve)
    }
  } else {
    summary_df <- switch(summary_type,
      "rdocx" = officer::docx_summary(x),
      "rpptx" = officer::pptx_summary(x),
      "docx" = officer::docx_summary(x),
      "pptx" = officer::pptx_summary(x),
      "slide" = officer::slide_summary(x, index = index),
      "layout" = officer::layout_summary(x)
    )
  }

  if (as_tibble) {
    return(tibble::as_tibble(summary_df))
  }

  summary_df
}


#' Check a summary data.frame created from a rdocx or rpptx object
#'
#' Check an object and error if an object is not a data.frame with the required
#' column names to be a summary data.frame created from a rdocx or rpptx object.
#' Optionally check for the number of rows, a specific content_type value, or if
#' tables are included in the document the summary was created from.
#'
#' @param x An object to check if it is a data.frame object created with
#'   [officer::docx_summary()] or another summary function.
#' @param n Required number of rows. Optional. If n is more than length 1,
#'   checks to make sure the number of rows is within the range of `max(n) `and
#'   `min(n)`. Defaults to `NULL`.
#' @param content_type Required content_type, e.g. "paragraph", "table cell", or
#'   "image". Optional. Defaults to `NULL`.
#' @param summary_type Summary type. Options "doc", "docx", "pptx", "slide", or
#'   "layout". "doc" requires the data.frame include a "content_type" column but
#'   allows columns for either a docx or pptx summary.
#' @param tables If `TRUE`, require that the summary include the column names
#'   indicated a table is present in the rdocx or rpptx summary.
#' @param arg Argument name to use in error messages. Defaults to
#'   `caller_arg(x)`
#' @param ... Additional parameters passed to [cli::cli_abort()]
#' @family summary functions
#' @inheritParams cli::cli_abort
#' @export
#' @importFrom cli cli_abort
check_officer_summary <- function(x,
                                  n = NULL,
                                  content_type = NULL,
                                  summary_type = "doc",
                                  tables = FALSE,
                                  ...,
                                  arg = caller_arg(x),
                                  call = parent.frame()) {
  if (!is_officer_summary(x, summary_type, tables)) {
    fn_nm <- c("officer_summary", "docx_summary", "pptx_summary")
    cli_abort(
      "{.arg {arg}} must be a data frame from
      {.fn {cli_vec_last(fn_nm)}}.",
      ...,
      call = call
    )
  }

  n_rows <- nrow(x)

  if (!is.null(n) && !((n_rows <= max(n)) && (n_rows >= min(n)))) {
    cli_abort(
      "{.arg {arg}} must have {n} row{?s}, not {n_rows}.",
      ...,
      call = call
    )
  }

  if (!is.null(content_type) && !all(x[["content_type"]] %in% content_type)) {
    cli_abort(
      "{.arg {arg}} must include only {.val {content_type}} content type{?s}.",
      ...,
      call = call
    )
  }
}

#' @keywords internal
#' @noRd
#' @importFrom rlang has_name
is_officer_summary <- function(x,
                               summary_type = "doc",
                               tables = FALSE,
                               call = caller_env()) {
  check_required(x, call = call)
  check_bool(tables, call = call)
  check_string(summary_type, allow_null = TRUE, call = call)

  if (!is.data.frame(x)) {
    return(FALSE)
  }

  summary_type <- arg_match0(
    summary_type,
    values = c("doc", "docx", "pptx", "slide", "layout"),
    error_call = call
  )

  nm <- "content_type"
  docx_nm <- c("doc_index", "content_type", "style_name", "text")
  pptx_nm <- c("text", "id", "content_type", "slide_id")

  if (tables) {
    nm <- c(nm, "row_id", "cell_id")
  }

  nm <- switch(summary_type,
    "docx" = c(nm, docx_nm),
    "pptx" = c(nm, pptx_nm),
    "slide" = c("type", "id"),
    "layout" = c("layout", "master"),
    nm
  )

  nm_check <- TRUE
  if (identical(summary_type, "doc")) {
    nm_check <- any(has_name(x, c(docx_nm, pptx_nm)))
  }

  all(has_name(x, nm)) && nm_check
}
