#' Write a rdocx object to a docx file
#'
#' @param docx A rdocx or rpptx object to save.
#' @param path File path.
#' @param overwrite If `TRUE` (default), remove file at path if it already
#'   exists. If `FALSE` and file exists, this function aborts.
#' @inheritDotParams officer::set_doc_properties -x
#' @export
#' @returns Returns rdocx input object (invisibly) and writes the rdocx object
#'   to a file with a name and location matching path.
#' @importFrom rlang check_required check_installed
#' @importFrom cli cli_abort
#' @importFrom officer set_doc_properties
write_docx <- function(docx, path, overwrite = TRUE, ...) {
  rlang::check_required(docx)
  rlang::check_required(path)

  check_officer(docx)
  check_office_fileext(path, fileext = c("docx", "pptx"))

  params <- rlang::list2(...)

  if (isFALSE(rlang::is_empty(params))) {
    docx <-
      officer::set_doc_properties(
        docx,
        ...
      )
  }

  if (file.exists(path)) {
    if (!isTRUE(overwrite)) {
      cli::cli_abort("{.arg overwrite} must be {.code TRUE} if {.arg path} exists.")
    }
    file.remove(path)
  }

  print(docx, target = path)

  invisible(docx)
}
