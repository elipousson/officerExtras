#' Create new columns to an officer summary data.frame based on specific levels
#'
#' [officer_summary_levels()] works with [fill_with_pattern()] to help convert
#' hierarchically organized text, e.g. text with leveled headings or other
#' styles into a data.frame where new columns hold the value of the preceding
#' (or succeeding) heading text.
#'
#' @param levels Levels to use. If `NULL` (default), levels is set to unique,
#'   non-NA values from the levels_from column.
#' @param levels_from Column name to use for identifying levels. Defaults to
#'   "style_name".
#' @param exclude_levels Levels to exclude from process of adding new columns.
#' @inheritParams fill_with_pattern
#' @param strict If `TRUE`, error unless all values provided to `level` are
#'   present in the `levels_from` column. Defaults to `FALSE` which warns if
#'   invalid values are present.
#' @inheritDotParams officer_summary
#' @returns A summary data frame with additional columns based on the supplied
#'   levels.
#' @family summary functions
#' @export
officer_summary_levels <- function(x,
                                   levels = NULL,
                                   levels_from = "style_name",
                                   exclude_levels = NULL,
                                   fill_col = "text",
                                   direction = c("down", "up", "downup", "updown"),
                                   ...,
                                   strict = FALSE,
                                   call = caller_env()) {
  check_required(x, call = call)
  if (!is_officer_summary(x)) {
    x <- officer_summary(x, ..., call = call)
  }

  # Validate `levels_from`
  check_name(levels_from, call = call)
  levels_from <- arg_match(levels_from, values = names(x), error_call = call)

  # If levels is `NULL`, set levels to match unique values in `levels_from`
  # column
  level_values <- unique(x[[levels_from]])
  levels <- levels %||% level_values
  levels <- levels[!is.na(levels)]
  check_character(levels, call = call)

  # Optionally exclude specified levels
  if (!is_null(exclude_levels)) {
    check_character(exclude_levels, call = call)
    levels <- levels[!(tolower(levels) %in% tolower(exclude_levels))]
  }

  # Handle invalid levels values with error or warning
  if (strict) {
    levels <- arg_match(levels, values = level_values, error_call = call)
  } else {
    valid_levels <- levels %in% level_values

    if (!all(valid_levels)) {
      levels_invalid <- levels[!valid_levels]

      cli_bullets(
        c("!" = "`levels` contains invalid values that are not
        present in {levels_from} column: {levels_invalid}")
      )
    }
  }

  x[["officer_summary_level_name"]] <- factor(x[[levels_from]], levels = levels)

  for (lvl in levels[!is.na(levels)]) {
    x <- fill_with_pattern(
      x,
      pattern = lvl,
      pattern_col = "officer_summary_level_name",
      fill_col = fill_col,
      col = underscore(lvl),
      direction = direction,
      call = call
    )
  }

  x[["officer_summary_level_name"]] <- NULL
  x
}

#' @noRd
underscore <- function(x) {
  sub(" ", "_", tolower(x))
}
