
#' Add an xml string or paragraph of text at a specified position in a Word document
#'
#' Wrappers for [officer::body_add()] and [officer::body_add_xml()] that use the
#' [cursor_docx()] helper functions to allow input of filename and path and use
#' of multiple options for cursor placement. [add_text_to_body()] passed value
#' to [glue::glue()] to add support for glue string interpolation.
#' [add_gt_to_body()] converts gt tables to OOXML with [gt::as_word()].
#'
#' @inheritParams cursor_docx
#' @inheritParams read_docx_ext
#' @inheritParams cursor_docx
#' @inheritParams officer::body_add
#' @inheritParams officer::body_add_xml
#' @param ... Additional parameters passed to [officer::body_add()] or
#'   [officer::body_add_xml()].
#' @returns A rdocx object with added xml, gt tables, or paragraphs of text.
#' @export
#' @importFrom officer body_add body_add_xml
#' @importFrom cli cli_alert_warning
add_to_body <- function(docx,
                        keyword = NULL,
                        id = NULL,
                        index = NULL,
                        value = NULL,
                        str = NULL,
                        ...) {
  check_docx(docx)
  docx <- cursor_docx(docx, keyword, id, index)

  if (!is.null(value)) {
    return(officer::body_add(docx, value, ...))
  }

  if (!is.null(str)) {
    return(officer::body_add_xml(docx, str, ...))
  }

  cli::cli_alert_warning(
    "{.arg value} or {.arg str} are required to modify the {.arg docx} object."
  )

  docx
}

#' @inheritParams glue::glue
#' @name add_text_to_body
#' @rdname add_to_body
#' @export
#' @importFrom rlang check_required
#' @importFrom glue glue
add_text_to_body <- function(docx,
                             value,
                             style = NULL,
                             .na = "NA",
                             .null = NULL,
                             .envir = parent.frame(),
                             ...) {
  rlang::check_required(value)
  value <- glue::glue(value, .na = .na, .null = .null, .envir = .envir)
  add_to_body(docx, value = value, style = style, ...)
}

#' @inheritParams officer::body_add_par
#' @name add_xml_to_body
#' @rdname add_to_body
#' @export
add_xml_to_body <- function(docx,
                            str = NULL,
                            pos = "after",
                            ...) {
  add_to_body(docx, str = str, pos = pos, ...)
}

#' @param gt_object A gt object converted to an OOXML string with
#'   [gt::as_word()] then passed to [add_xml_to_body()]. Required.
#' @name add_gt_to_body
#' @rdname add_to_body
#' @export
#' @importFrom rlang check_required check_installed
add_gt_to_body <- function(docx,
                           gt_object,
                           pos = "after",
                           ...) {
  rlang::check_required(gt_object)
  rlang::check_installed("gt")
  add_xml_to_body(docx, str = gt::as_word(gt_object), pos = pos, ...)
}
