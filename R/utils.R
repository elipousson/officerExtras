# @staticimports pkg:isstatic
#   is_fileext_path is_all_null str_extract_fileext

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

# ---
# repo: r-lib/rlang
# file: standalone-purrr.R
# last-updated: 2023-02-23
# license: https://unlicense.org
# ---
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
