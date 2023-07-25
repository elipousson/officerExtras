#' Add a vector of objects to a rdocx object using `add_to_body()`
#'
#' [vec_add_to_body()] is a vectorized variant [add_to_body()] that allows users
#' to supply a vector of values or str inputs to add multiple blocks of text,
#' images, plots, or tables to a document. Alternatively, the function also
#' supports adding a single object at multiple locations or using multiple
#' different styles.  All parameters are recycled using
#' [vctrs::vec_recycle_common()] so inputs must be length 1 or match the length
#' of the longest input vector.
#'
#' @inheritParams add_to_body
#' @inheritDotParams add_to_body
#' @inheritParams vctrs::vec_recycle_common
#' @examples
#' docx_example <- read_officer()
#'
#' docx_example <- vec_add_to_body(
#'   docx_example,
#'   value = c("Sample text 1", "Sample text 2", "Sample text 3"),
#'   style = c("heading 1", "heading 2", "Normal")
#' )
#'
#' officer_summary(docx_example)
#'
#' if (is_installed("gt")) {
#'   gt_tbl <- gt::gt(gt::gtcars[1:2, 1:2])
#'
#'   # list inputs such as gt tables must be passed within a list to avoid
#'   # issues
#'   docx_example <- vec_add_to_body(
#'     docx_example,
#'     gt_object = list(gt_tbl, gt_tbl),
#'     keyword = c("Sample text 1", "Sample text 2")
#'   )
#'
#'   officer_summary(docx_example)
#' }
#' @export
vec_add_to_body <- function(docx,
                            ...,
                            .size = NULL,
                            .call = caller_env()) {
  params <- vctrs::vec_recycle_common(..., .size = .size, .call = .call)

  for (i in seq_along(params[[1]])) {
    docx <- exec(
      .fn = add_to_body,
      docx = docx,
      !!!lapply(params, vctrs::vec_slice, i = i, error_call = call)
    )
  }

  docx
}
