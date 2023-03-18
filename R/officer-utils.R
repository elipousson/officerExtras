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
subset_type <- function(x, type) {
  if (rlang::has_name(x, "content_type")) {
    x[x[["content_type"]] %in% type, ]
  }
}

#' Subset officer object summary by style_name
#'
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

#' Subset officer object summary by doc_index or id
#'
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
