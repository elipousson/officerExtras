#' Convert an rodcx object or Word file to another format with pandoc
#'
#' Use [rmarkdown::pandoc_convert()] to convert a rodcx object or Word file to
#' another format with pandoc.
#'
#' @param docx A rdocx object or path to a docx file. Optional if input is
#'   provided.
#' @inheritParams rmarkdown::pandoc_convert
#' @inheritParams write_officer
#' @inheritDotParams rmarkdown::pandoc_convert -verbose
#' @param quiet If `TRUE`, suppress alert messages and pass `FALSE` to verbose
#'   parameter of [rmarkdown::pandoc_convert()]. Defaults `TRUE`.
#' @inherit rmarkdown::pandoc_convert details
#' @returns Executes a call to pandoc using [rmarkdown::pandoc_convert()] to
#'   create a file from an officer object or a docx file.
#' @examples
#' \dontrun{
#' if (interactive()) {
#'   docx <- read_officer(
#'     system.file("doc_examples/example.docx", package = "officer")
#'   )
#'
#'   withr::with_tempdir({
#'     docx_convert(
#'       docx,
#'       output = "docx.md"
#'     )
#'   })
#' }
#' }
#' @seealso
#'  [rmarkdown::pandoc_convert()]
#' @rdname docx_convert
#' @export
#' @importFrom rlang check_installed is_true is_false
#' @importFrom rmarkdown pandoc_convert
docx_convert <- function(docx,
                         to = NULL,
                         input = NULL,
                         output = NULL,
                         path = NULL,
                         quiet = TRUE,
                         overwrite = TRUE,
                         ...) {
  rlang::check_installed("rmarkdown")

  if (is_officer(docx, "rdocx")) {
    if (!is.null(input) && is_true(quiet)) {
      cli::cli_alert_warning(
        "{.arg input} is ignored if {.arg x} is a {.cls rdocx} object.",
        wrap = TRUE
      )
    }

    input <- tempfile(fileext = ".docx")

    write_officer(
      docx,
      path = input,
      overwrite = overwrite
    )
  } else if (is.null(input)) {
    rlang::check_required(docx)
    check_docx_fileext(docx)
    input <- docx
  }

  if (!is.null(path) && !dir.exists(path)) {
    dir.create(path)
  }

  rmarkdown::pandoc_convert(
    input = input,
    to = to,
    output = output,
    verbose = is_false(quiet),
    wd = path,
    ...
  )
}
