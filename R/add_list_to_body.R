#' Add a formatted list of items to a `rdocx` object
#'
#' This function assumes that the input docx object has a list style. Use
#' [read_docx_ext()] with `allow_null = TRUE` to create a new rdocx object with
#' the "List Bullet" and "List Number" styles. "List Paragraph" can be used but
#' the resulting formatted list will not have bullets or numbers.
#'
#' @param docx A `rdocx` object.
#' @param values Character vector with values for list.
#' @param style Style to use for list of values, Default: 'List Bullet'
#' @param keep_na If `TRUE`, keep `NA` values in values Default: FALSE
#' @param before Value to insert before the list, Default: NULL
#' @param after Value to insert after the list, Default: ''
#' @param ... Additional parameters passed to [cursor_docx()]
#' @returns A modified `rdocx` object.
#' @seealso
#'  [add_to_body()], [vec_add_to_body()]
#' @rdname add_list_to_body
#' @export
#' @importFrom rlang has_length
add_list_to_body <- function(docx,
                             values,
                             style = "List Bullet",
                             keep_na = FALSE,
                             before = NULL,
                             after = "",
                             ...) {

  check_docx(docx)
  check_character(values)
  if (!keep_na && !all(is.na(values))) {
    values <- values[!is.na(values)]
  }

  if (all(is.na(values)) || rlang::has_length(values, 0)) {
    return(docx)
  }

  docx <- cursor_docx(docx, ...)

  if (!is.null(before)) {
    docx <- add_to_body(docx = docx, value = before)
  }

  docx <- vec_add_to_body(
    docx = docx,
    value = values,
    style = style
  )

  if (!is.null(after)) {
    docx <- add_to_body(docx = docx, value = after)
  }
}
