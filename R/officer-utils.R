#' Preview a rdocx, rpptx, or rxlsx object in local default applications
#'
#' `officer_open()` uses `officer::open_file()` to open a file ceated from a
#' rdocx, rpptx, or rxlsx object.
#'
#' @inheritParams write_officer
#' @param path Optional. Set as temporary file with file extension matching type
#' of input object `x`.
#' @param overwrite Defaults to `FALSE`.
#' @param interactive If `FALSE`, warn the user and return `x`.
#' @returns Input object `x` without modification.
#' @export
officer_open <- function(
  x,
  path = NULL,
  ...,
  overwrite = FALSE,
  interactive = is_interactive()
) {
  if (!interactive) {
    cli::cli_warn(
      "{.fn officer_open} can't be used in a non-interactive session."
    )

    return(x)
  }

  new_path <- path %||% tempfile(fileext = officer_fileext(x, prefix = "."))

  write_officer(
    x,
    path = new_path,
    ...,
    overwrite = overwrite
  )

  officer::open_file(new_path)

  x
}

#' Convert officer object class into equivalent file extension
#'
#' @keywords internal
#' @noRd
officer_fileext <- function(x, prefix = "") {
  paste0(prefix, str_remove(class(x), "^r"))
}

#' Subset officer object summary by content_type
#'
#' @keywords internal
#' @noRd
#' @importFrom rlang has_name
subset_type <- function(x, type) {
  if (has_name(x, "content_type")) {
    x[x[["content_type"]] %in% type, ]
  }
}

#' Subset officer object summary by style_name
#'
#' @keywords internal
#' @noRd
#' @importFrom rlang has_name
subset_style <- function(x, style) {
  if (has_name(x, "style_name")) {
    if (!is.na(style)) {
      x[!is.na(x[["style_name"]]) && x[["style_name"]] %in% style, ]
    } else {
      x[is.na(x[["style_name"]]), ]
    }
  }
}

#' Subset officer object summary by doc_index or id
#'
#' @keywords internal
#' @noRd
#' @importFrom rlang has_name
subset_index <- function(x, index, index_type = NULL) {
  if (!is.null(index_type) && has_name(x, index_type)) {
    stopifnot(is_string(index_type))
    x[x[[index_type]] %in% index, ]
  } else if (has_name(x, "doc_index")) {
    x[x[["doc_index"]] %in% index, ]
  } else if (has_name(x, "id")) {
    x[x[["id"]] %in% index, ]
  }
}

#' @keywords internal
#' @noRd
#' @importFrom rlang is_true is_false
subset_header <- function(x, header = TRUE) {
  if (is_true(header)) {
    x[x[["is_header"]], ]
  } else if (is_false(header)) {
    x[!x[["is_header"]], ]
  } else {
    x
  }
}
