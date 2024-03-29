#' Combine multiple rdocx objects or docx files
#'
#' @description
#' [combine_docx()] is a variant of [officer::body_add_docx()] that allows any
#' number of input files and supports rdocx objects as well as Word file paths.
#' Optionally use a separator between files or objects.
#'
#' Please note that when you create a new rdocx object with this function (or
#' [officer::body_add_docx()]) the added content will not appear in a summary
#' data frame created with [officer_summary()] and is not accessible to other
#' functions until the document is *opened and edited* with Microsoft Word. This
#' is part of how the OOXML [AltChunk
#' Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.altchunk?view=openxml-2.8.1)
#' works and can't be avoided.
#'
#' @param ... Any number of additional rdocx objects or docx file paths.
#' @param .list Any number of additional rdocx objects or docx file paths passed
#'   as a list. Defaults to [list2()]
#' @param docx A rdocx object or a file path with a docx file extension.
#'   Defaults to `NULL`.
#' @inheritParams officer::body_add_docx
#' @param sep Separator to use docx files. A bare function that takes a rdocx
#'   object as the only parameter, such as [officer::body_add_break] or another
#'   object passed to [add_to_body()] as the value parameter. Optional. Defaults
#'   to `NULL`.
#' @inheritParams check_officer
#' @return A rdocx object.
#' @examples
#' \dontrun{
#' if (interactive()) {
#'   docx_path <- system.file("doc_examples", "example.docx", package = "officer")
#'   docx <- read_officer(docx_path)
#'
#'   combine_docx(
#'     docx,
#'     docx_path
#'   )
#' }
#' }
#' @seealso
#'  [officer::body_add_docx()]
#' @rdname combine_docx
#' @export
#' @importFrom vctrs vec_recycle
#' @importFrom officer body_add_docx
combine_docx <- function(...,
                         docx = NULL,
                         .list = list2(...),
                         pos = "after",
                         sep = NULL,
                         call = caller_env()) {
  if (is_string(docx)) {
    check_office_fileext(docx, fileext = "docx", call = call)
    docx <- read_docx_ext(docx)
  } else if (is.null(docx)) {
    docx <- read_docx_ext(allow_null = TRUE)
  }

  check_officer(docx, what = "rdocx", call = call)

  .list <- .list %||% list2(...)
  size <- length(.list)

  if (!is.null(sep)) {
    sep <- vctrs::vec_recycle(sep, size = size, x_arg = "size", call = call)
  }

  for (i in seq(size)) {
    src <- .list[[i]]

    if (is_rdocx(.list[[i]])) {
      src <- officer_temp(fileext = "docx")
      write_officer(.list[[i]], path = src)
    }

    if (!is_fileext_path(src, "docx") || !file.exists(src)) {
      cli_abort(
        "Every object in {.arg .list} must be a {.cls rdocx} object or a path to
        an existing docx file.",
        call = call
      )
    }

    docx <- officer::body_add_docx(
      x = docx,
      src = src
    )

    if (!is.null(sep) && (i < size)) {
      if (is_function(sep[[i]])) {
        docx <- eval(sep, docx)
      } else {
        docx <- add_to_body(docx, value = sep[[i]], pos = pos, call = call)
      }
    }
  }

  docx
}
