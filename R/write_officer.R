#' Write a rdocx, rpptx, or rxlsx object to a file
#'
#' @param x A rdocx, rpptx, or rxlsx object to save. Document properties won't
#'   be set for rxlsx objects.
#' @param path File path.
#' @param overwrite If `TRUE` (default), remove file at path if it already
#'   exists. If `FALSE` and file exists, this function aborts.
#' @param modified_by If the withr package is installed, modified_by overrides
#'   the default value for the lastModifiedBy property assigned to the output
#'   file by officer. Defaults to `Sys.getenv("USER")` (the same value used by
#'   officer).
#' @inheritDotParams officer::set_doc_properties -x
#' @returns Returns the input object (invisibly) and writes the rdocx, rpptx, or
#'   rxlsx object to a file with a name and location matching the provided path.
#' @export
#' @importFrom rlang check_required check_installed current_call
#' @importFrom cli cli_abort
#' @importFrom officer set_doc_properties
write_officer <- function(x,
                          path,
                          overwrite = TRUE,
                          modified_by = Sys.getenv("USER"),
                          ...) {
  check_officer(x, call = current_call())

  check_office_fileext(
    path,
    fileext = c("docx", "pptx", "xlsx"),
    call = current_call()
  )

  is_xlsx <- isTRUE(is_fileext_path(path, "xlsx"))

  if (!is_empty(list2(...)) && !is_xlsx) {
    x <- officer::set_doc_properties(x, ...)
  }

  if (file.exists(path)) {
    if (!isTRUE(overwrite)) {
      cli_abort(
        "{.arg overwrite} must be {.code TRUE} to replace existing file
        {.file {basename(path)}}."
      )
    }
    file.remove(path)
  }

  if (!is_installed("withr")) {
    print(x, target = path)
  } else {
    withr::with_envvar(
      list("USER" = modified_by),
      {
        print(x, target = path)
      }
    )
  }

  invisible(x)
}
