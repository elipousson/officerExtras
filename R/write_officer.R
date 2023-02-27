#' Write a rdocx object to a docx file
#'
#' @param x A rdocx, rpptx, or rxlsx object to save. Document properties won't
#'   be set for rxlsx objects.
#' @param path File path.
#' @param overwrite If `TRUE` (default), remove file at path if it already
#'   exists. If `FALSE` and file exists, this function aborts.
#' @inheritDotParams officer::set_doc_properties -x
#' @export
#' @returns Returns rdocx input object (invisibly) and writes the rdocx object
#'   to a file with a name and location matching path.
#' @importFrom rlang check_required check_installed current_call
#' @importFrom cli cli_abort
#' @importFrom officer set_doc_properties
write_officer <- function(x, path, overwrite = TRUE, ...) {
  rlang::check_required(x)
  rlang::check_required(path)

  check_officer(x, call = rlang::current_call())
  check_office_fileext(
    path,
    fileext = c("docx", "pptx", "xlsx"),
    call = rlang::current_call()
  )

  is_xlsx <- isTRUE(is_fileext_path(path, "xlsx"))

  params <- rlang::list2(...)

  if (isFALSE(rlang::is_empty(params)) & !is_xlsx) {
    x <-
      officer::set_doc_properties(
        x,
        ...
      )
  }

  if (file.exists(path)) {
    if (!isTRUE(overwrite)) {
      cli_abort("{.arg overwrite} must be {.code TRUE} if {.arg path} exists.")
    }
    file.remove(path)
  }

  print(x, target = path)

  invisible(x)
}
