#' Write a rdocx object to a docx file
#'
#' @param docx A rdocx object to save.
#' @param path File path.
#' @param overwrite If `TRUE` (default), remove file at path if it already
#'   exists. If `FALSE` and file exists, this function aborts.
#' @export
#' @returns Returns rdocx input object (invisibly) and writes the rdocx object
#'   to a file with a name and location matching path.
#' @importFrom rlang check_required check_installed
#' @importFrom cli cli_abort
write_docx <- function(docx, path, overwrite = TRUE) {
  rlang::check_required(docx)
  rlang::check_required(path)

  check_docx(docx)
  check_docx_fileext(path)

  if (file.exists(path)) {
    if (!isTRUE(overwrite)) {
      cli::cli_abort("{.arg overwrite} must be {.code TRUE} if {.arg path} exists.")
    }
    file.remove(path)
  }

  print(docx, target = path)

  invisible(docx)
}
