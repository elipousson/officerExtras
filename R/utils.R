# @staticimports pkg:isstatic
#   is_all_null is_any_null is_fileext_path is_ggplot str_extract_fileext
#   has_fileext

# @staticimports pkg:stringstatic
#   str_detect str_remove str_c

#' @noRd
#' @importFrom glue glue
wrap_tag <- function(..., tag) {
  glue::glue(
    "<{tag}>",
    ...,
    "</{tag}>"
  )
}

#' @noRd
office_temp <- function(fileext = NULL, path = NULL, ...) {
  str_remove(
    tempfile(
      ...,
      tmpdir = path %||% "",
      fileext = paste0(".", fileext)
    ),
    paste0("^", .Platform$file.sep)
  )
}

#' @param vec_last Used as value for "vec-last" item in style object.
#' @keywords internal
#' @noRd
#' @importFrom cli cli_vec
cli_vec_last <- function(x, style = list(), vec_last = " or ") {
  cli::cli_vec(
    x,
    style = c(
      style,
      "vec-last" = vec_last
    )
  )
}

#' @keywords internal
#' @noRd
cli_vec_cls <- function(x) {
  cli_vec_last(
    x,
    style = list(
      before = "<",
      after = ">",
      color = "blue"
    )
  )
}

#' Fill a new column based on an existing column depending on a pattern
#'
#' This function uses [vctrs::vec_fill_missing()] to convert hierarchically
#' nested headings and text into a rectangular data.frame. It is an experimental
#' function that may be modified or removed. At present, it is only used by
#' [officer_tables()].
#'
#' @param x A input data.frame (assumed to be from [officer_summary()] for
#'   default values).
#' @param pattern Passed to [grepl()] as the pattern identifying which rows of the
#'   fill_col should have values pulled into the new column named by col.
#'   Defaults to "^heading" which matches the default heading style names.
#' @param pattern_col Name of column to use for pattern matching, Defaults to
#'   "style_name".
#' @param fill_col Name of column to fill , Defaults to "text".
#' @param col Name of new column to fill with values from fill_col, Default:
#'   Defaults to "heading"
#' @param direction Direction of fill passed to [vctrs::vec_fill_missing()],
#'   Default: c("down", "up", "downup", "updown")
#' @inheritParams rlang::args_error_context
#' @returns A data.frame with an additional column taking the name from col and
#'   the values from the column named in fill_col.
#' @seealso
#'  \code{\link[vctrs]{vec_fill_missing}}
#' @rdname fill_with_pattern
#' @export
fill_with_pattern <- function(x,
                              pattern = "^heading",
                              pattern_col = "style_name",
                              fill_col = "text",
                              col = "heading",
                              direction = c("down", "up", "downup", "updown"),
                              call = caller_env()) {
  check_name(col, call = call)
  check_name(fill_col, call = call)

  if (has_name(x, col)) {
    cli_abort(
      "{.arg col} {.field {col}} can't be used when
      {.arg x} has a column name {.field {col}}."
    )
  }

  pattern <- str_detect(x[[pattern_col]], pattern)
  pattern[is.na(pattern)] <- FALSE

  if (sum(pattern, na.rm = TRUE) == 0) {
    return(x)
  }
  check_installed("vctrs")
  x[[col]] <- NA_character_
  x[pattern, ][[col]] <- x[pattern, ][[fill_col]]
  # x[!pattern, ][[col]] <- rep(NA_character_, nrow(x))[!pattern]
  x[[col]] <- vctrs::vec_fill_missing(x[[col]], direction = direction)
  x
}


# ---
# repo: r-lib/rlang
# file: standalone-purrr.R
# last-updated: 2023-02-23
# license: https://unlicense.org
# ---
map <- function(.x, .f, ...) {
  .f <- rlang::as_function(.f, env = rlang::global_env())
  lapply(.x, .f, ...)
}

#' @keywords internal
map_lgl <- function(.x, .f, ...) {
  .rlang_purrr_map_mold(.x, .f, logical(1), ...)
}

#' @keywords internal
#' @importFrom rlang as_function global_env
.rlang_purrr_map_mold <- function(.x, .f, .mold, ...) {
  .f <- as_function(.f, env = global_env())
  out <- vapply(.x, .f, .mold, ..., USE.NAMES = FALSE)
  names(out) <- names(.x)
  out
}

#' @keywords internal
discard <- function(.x, .p, ...) {
  sel <- .rlang_purrr_probe(.x, .p, ...)
  .x[is.na(sel) | !sel]
}

#' @importFrom rlang is_logical as_function
#' @keywords internal
.rlang_purrr_probe <- function(.x, .p, ...) {
  if (is_logical(.p)) {
    stopifnot(length(.p) == length(.x))
    .p
  } else {
    .p <- as_function(.p, env = global_env())
    map_lgl(.x, .p, ...)
  }
}
