#' Write a rdocx object to a docx file
#'
#' @param docx A rdocx object to save.
#' @param path File path.
#' @param overwrite If `FALSE` and file exists, abort save function. If `TRUE`,
#'   save file even if file exists at path.
#' @export
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
}
