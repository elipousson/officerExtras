#' @keywords internal
#' @noRd
subset_type <- function(x, type) {
  if (rlang::has_name(x, "content_type")) {
    x[x[["content_type"]] %in% type, ]
  }
}

#' @keywords internal
#' @noRd
subset_style <- function(x, style) {
  if (rlang::has_name(x, "style_name")) {
    if (!is.na(style)) {
      x[!is.na(x[["style_name"]]) & x[["style_name"]] %in% style, ]
    } else {
      x[is.na(x[["style_name"]]), ]
    }
  }
}

#' @keywords internal
#' @noRd
subset_index <- function(x, index) {
  if (rlang::has_name(x, "doc_index")) {
    x[x[["doc_index"]] %in% index, ]
  } else if (rlang::has_name(x, "id")) {
    x[x[["id"]] %in% index, ]
  }
}

#' @keywords internal
#' @noRd
subset_header <- function(x, header = TRUE) {
  if (isTRUE(header)) {
    x[isTRUE(x[["is_header"]]), ]
  } else if (isFALSE(header)) {
    x[isFALSE(x[["is_header"]]), ]
  } else {
    x
  }
}
