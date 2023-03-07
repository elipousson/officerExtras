#' Set cursor position in rdocx object based on keyword, id, or index
#'
#' A combined function for setting cursor position with
#' [officer::cursor_reach()], [officer::cursor_bookmark()], or using a doc_index
#' value from [officer::docx_summary()].
#'
#' @param docx A rdocx object.
#' @param keyword,id A keyword string used to place cursor with
#'   [officer::cursor_reach()] or bookmark id with [officer::cursor_bookmark()].
#'   Defaults to `NULL`. If keyword or id are not provided, the gt object is
#'   inserted at the front of the document.
#' @param index A integer matching a doc_index value appearing in a summary of
#'   the docx object created with [officer::docx_summary()]. If index is for a
#'   paragraph value, the text of the pargraph is used as a keyword.
#' @param quiet If `FALSE` (default) warn when keyword is not found.
#' @seealso
#'  [officer::cursor_begin()], [officer::docx_summary()]
#' @export
#' @importFrom rlang check_required
#' @importFrom cli cli_alert_warning cli_abort
#' @importFrom officer cursor_reach_test cursor_reach docx_bookmarks
#'   cursor_bookmark docx_summary
cursor_docx <- function(docx, keyword = NULL, id = NULL, index = NULL, quiet = FALSE) {
  check_docx(docx)

  if (!is.null(keyword)) {
    if (isFALSE(officer::cursor_reach_test(docx, keyword)) & isFALSE(quiet)) {
      cli::cli_alert_warning(
        "{.arg keyword} {.val {keyword}} can't be found in {.arg docx}."
      )

      return(docx)
    }

    return(officer::cursor_reach(docx, keyword))
  }

  if (!is.null(id)) {
    bookmarks <- officer::docx_bookmarks(docx)
    id <- match.arg(id, bookmarks)
    return(officer::cursor_bookmark(docx, id))
  }

  if (!is.null(index)) {
    return(cursor_index(docx, index))
  }

  cli_abort(
    "{.arg keyword}, {.arg id}, or {.arg index} must be supplied."
  )
}

#' @keywords internal
#' @noRd
#' @export
#' @importFrom officer docx_summary cursor_reach
cursor_index <- function(docx, index = NULL) {
  docx_df <- subset_index(officer::docx_summary(docx), index)

  check_officer_summary(
    docx_df,
    n = 1,
    content_type = "paragraph",
    summary_type = "docx"
  )

  officer::cursor_reach(docx, keyword = docx_df[["text"]])
}
