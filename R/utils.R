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
officer_temp <- function(...,
                         path = NULL,
                         fileext = c("docx", "pptx", "xslx")) {
  fileext <- match.arg(fileext)

  tempfile(
    ...,
    tmpdir = path %||% tempdir(),
    fileext = paste0(".", fileext)
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
#' @param pattern Passed to [grepl()] as the pattern identifying which rows of
#'   the fill_col should have values pulled into the new column named by col.
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
#'  [vctrs::vec_fill_missing()]
#' @rdname fill_with_pattern
#' @export
#' @importFrom vctrs vec_fill_missing
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
      {.arg x} has a column name {.field {col}}.",
      call = call
    )
  }

  pattern <- str_detect(x[[pattern_col]], pattern)
  pattern[is.na(pattern)] <- FALSE

  if (sum(pattern, na.rm = TRUE) == 0) {
    return(x)
  }

  x[[col]] <- NA_character_
  x[pattern, ][[col]] <- x[pattern, ][[fill_col]]
  # x[!pattern, ][[col]] <- rep(NA_character_, nrow(x))[!pattern]
  x[[col]] <- vctrs::vec_fill_missing(x[[col]], direction = direction)
  x
}
