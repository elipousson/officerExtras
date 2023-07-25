#' Convert an rodcx object or Word file to another format with pandoc
#'
#' `convert_docx()` uses [rmarkdown::pandoc_convert()] to convert a rodcx object
#' or Word file to another format with pandoc. If the "to" parameter contains a
#' file extension, it is assumed to be an output file name. If you want to
#' convert a file to a Word document, use the input parameter for the path to
#' the Markdown, HTML, or other file.
#'
#' @param docx A rdocx object or path to a docx file. Optional if input is
#'   provided.
#' @inheritParams rmarkdown::pandoc_convert
#' @inheritParams write_officer
#' @param path File path for converted file passed to wd parameter of
#'   [rmarkdown::pandoc_convert()]. If docx is a rdocx object, path defaults to
#'   current working directory instead of the base directory of the input file
#'   path.
#' @inheritDotParams rmarkdown::pandoc_convert -verbose -wd
#' @param quiet If `TRUE`, suppress alert messages and pass `FALSE` to verbose
#'   parameter of [rmarkdown::pandoc_convert()]. Defaults `TRUE`.
#' @param extract_media If `TRUE`, append `"--extract-media=."` to
#'   options. Defaults to `FALSE`.
#' @inherit rmarkdown::pandoc_convert details
#' @returns Executes a call to pandoc using [rmarkdown::pandoc_convert()] to
#'   create a file from an officer object or a docx file.
#' @examples
#' docx_example <- read_officer(
#'   system.file("doc_examples/example.docx", package = "officer")
#' )
#'
#' convert_docx(
#'   docx_example,
#'   to = "markdown"
#' )
#'
#' withr::with_tempdir({
#'   convert_docx(
#'     docx_example,
#'     output = "docx_example.html"
#'   )
#' })
#' @seealso
#'  [rmarkdown::pandoc_convert()]
#' @rdname convert_docx
#' @export
#' @importFrom rlang check_installed is_true is_false
convert_docx <- function(docx = NULL,
                         to = NULL,
                         input = NULL,
                         output = NULL,
                         path = NULL,
                         options = NULL,
                         extract_media = TRUE,
                         overwrite = TRUE,
                         quiet = TRUE,
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
      path = input
    )

    path <- path %||% getwd()
  } else if (is.null(input)) {
    rlang::check_required(docx)
    check_docx_fileext(docx)
    input <- docx
  }

  if (is_bool(extract_media) && is_true(extract_media)) {
    options <- c(options, "--extract-media=.")
  }

  if (!is.null(path) && !dir.exists(path)) {
    dir.create(path)
  }

  if (has_fileext(to) && is.null(output)) {
    output <- to
    to <- NULL
  }

  if (!is.null(output)) {
    output_exists <- file.exists(output)
    if (!is.null(path)) {
      output_exists <- file.exists(file.path(path, output))
    }

    if (is_true(output_exists) && is_false(overwrite)) {
      cli::cli_abort(
        c("{.arg output} {.filename {output}} already exists.",
          "*" = "Set {.code overwrite = TRUE} to replace existing file."
        )
      )
    }
  }

  rmarkdown::pandoc_convert(
    input = input,
    to = to,
    output = output,
    options = options,
    verbose = is_false(quiet),
    wd = path,
    ...
  )
}
