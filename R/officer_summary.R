#' Summarize a rdocx or rpptx object
#'
#' @param x A rdocx or rpptx object passed to [officer::docx_summary()],
#'   [officer::pptx_summary()], [officer::slide_summary()], or
#'   [officer::layout_summary()]. If x is a data.frame created with one of those
#'   functions, it is returned as is.
#' @inheritParams check_officer_summary
#' @inheritParams officer::slide_summary
#' @returns A data.frame object.
#' @export
#' @importFrom officer docx_summary pptx_summary slide_summary layout_summary
officer_summary <- function(x, summary_type = "doc", index = NULL) {
  if (is_officer_summary(x, summary_type)) {
    return(x)
  }

  what <- switch(summary_type,
    "docx" = "rdocx",
    "pptx" = "rpptx",
    "slide" = "rpptx",
    "layout" = "rpptx",
    c("rdocx", "rpptx")
  )

  check_officer(x, what = what)

  if (is.null(summary_type) | (summary_type == "doc")) {
    summary_type <- class(x)
  }

  switch(summary_type,
    "rdocx" = officer::docx_summary(x),
    "rpptx" = officer::pptx_summary(x),
    "docx" = officer::docx_summary(x),
    "pptx" = officer::pptx_summary(x),
    "slide" = officer::slide_summary(x, index),
    "layout" = officer::layout_summary(x)
  )
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
      "{.arg {arg}} must be a {.cls data.frame} from
      {.fn {cli_vec_last(fn_nm)}}.",
      ...,
      call = call
    )
  }

  if (!is.null(n) & !((nrow(x) <= max(n)) & (nrow(x) >= min(n)))) {
    cli_abort(
      "{.arg {arg}} must have {n} row{?s}.",
      ...,
      call = call
    )
  }

  if (!is.null(content_type) & !all(x[["content_type"]] %in% content_type)) {
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
                               tables = FALSE) {
  if (!is.data.frame(x)) {
    return(FALSE)
  }

  summary_type <- match.arg(summary_type, c("doc", "docx", "pptx", "slide", "layout"))

  nm <- "content_type"
  docx_nm <- c("doc_index", "content_type", "style_name", "text")
  pptx_nm <- c("text", "id", "content_type", "slide_id")

  if (isTRUE(tables)) {
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
  if (summary_type == "doc") {
    nm_check <- any(rlang::has_name(x, c(docx_nm, pptx_nm)))
  }

  all(rlang::has_name(x, nm)) & nm_check
}
