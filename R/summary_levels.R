#' Create new columns to an officer summary data.frame based on specific levels
#'
#' This function works with [fill_with_pattern()] to help convert hierarchically
#' organized text, e.g. text with leveled headings or other styles into a
#' data.frame where new columns hold the value of the preceding (or succeeding)
#' heading text.
#'
#' @param levels Levels to use. If `NULL` (default), levels is set to unique,
#'   non-NA values from the levels_from column.
#' @param levels_from Column name to use for identifying levels. Defaults to
#'   "style_name".
#' @param exclude_levels Levels to exclude from process of adding new columns.
#' @inheritParams fill_with_pattern
#' @inheritDotParams officer_summary
#' @export
officer_summary_levels <- function(x,
                                   levels = NULL,
                                   levels_from = "style_name",
                                   exclude_levels = NULL,
                                   fill_col = "text",
                                   direction = c("down", "up", "downup", "updown"),
                                   ...) {
  check_required(x)
  if (!is_officer_summary(x)) {
    x <- officer_summary(x, ..., call = call)
  }

  check_name(levels_from, call = call)
  levels <- levels %||% unique(x[[levels_from]])
  levels <- levels[!is.na(levels)]

  if (!is_null(exclude_levels)) {
    check_character(exclude_levels, call = call)
    levels <- levels[!(tolower(levels) %in% tolower(exclude_levels))]
  }

  check_character(levels, call = call)

  x[["officer_summary_level_name"]] <- factor(x[[levels_from]], levels = levels)
  levels <- levels[!is.na(levels)]
  for (lvl in levels) {
    x <-
      fill_with_pattern(
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
