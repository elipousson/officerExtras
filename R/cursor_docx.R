
#' Set cursor position in rdocx object based on keyword, id, or index
#'
#' @param docx A rdocx object.
#' @param keyword,id A keyword string used to place cursor with
#'   [officer::cursor_reach()] or bookmark id with [officer::cursor_bookmark()].
#'   Defaults to `NULL`. If keyword or id are not provided, the gt object is
#'   inserted at the front of the document.
#' @param index A integer matching a doc_index value appearing in a summary of
#'   the docx object created with [officer::docx_summary()]. If index is for a
#'   paragraph value, the text of the pargraph is used as a keyword.
#' @seealso
#'  [officer::cursor_begin()], [officer::docx_summary()]
#' @export
#' @importFrom rlang check_required
#' @importFrom cli cli_abort
#' @importFrom officer cursor_reach cursor_bookmark docx_summary
cursor_docx <- function(docx, keyword = NULL, id = NULL, index = NULL) {
  rlang::check_required(docx)

  if (is_all_null(c(keyword, id, index))) {
    cli::cli_abort(
      "One of {.arg keyword}, {.arg id}, or {.arg index} must be provided."
    )
  }

  if (!is.null(keyword)) {
    officer::cursor_reach(docx, keyword)
  } else if (!is.null(id)) {
    officer::cursor_bookmark(docx, id)
  } else if (!is.null(index)) {
    docx_summary_df <- officer::docx_summary(docx)
    docx_summary_df <- docx_summary_df[docx_summary_df[["doc_index"]] == index, ]

    check_docx_summary(docx_summary_df)

    officer::cursor_reach(docx, keyword = docx_summary_df[["text"]])
  }
}

#' Check a docx summary data.frame
#'
#' @param x A data.frame object created with [officer::docx_summary()].
#' @param n Required number of rows.
#' @param content_type Required content_type.
#' @param arg Argument name to use in error messages. Defaults to
#'   `caller_arg(x)`
#' @param ... Additional parameters passed to [cli::cli_abort()]
#' @inheritParams cli::cli_abort
#' @keywords internal
#' @importFrom cli cli_abort
check_docx_summary <- function(x,
                               n = 1,
                               content_type = "text",
                               ...,
                               arg = caller_arg(x),
                               call = parent.frame()) {
  if (!is.null(n) & nrow(x) != n) {
    cli::cli_abort(
      "{.arg {arg}} must have {n} rows.",
      ...,
      call = call,
    )
  }

  if (!is.null(content_type) & !all(x[["content_type"]] %in% content_type)) {
    cli::cli_abort(
      "{.arg {arg}} must only include content_type {.val {content_type}}.",
      ...,
      call = call,
    )
  }
}
