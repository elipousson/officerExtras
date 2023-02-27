
#' Add an xml string, text paragraph, or gt object at a specified position in a
#' Word document
#'
#' Wrappers for [officer::body_add_par()] and [officer::body_add_xml()] that use
#' the [cursor_docx()] helper functions to allow input of filename and path and
#' use of multiple options for cursor placement. [add_text_to_body()] passes
#' value to [glue::glue()] to add support for glue string interpolation.
#' [add_gt_to_body()] converts gt tables to OOXML with [gt::as_word()]. If `pos
#' = NULL`, [add_to_body()] calls [officer::body_add()] instead of
#' [officer::body_add_par()].
#'
#' @details Using [add_value_with_keys()] or [add_str_with_keys()]
#'
#' [add_value_with_keys()] supports value vectors of length 1 or longer. If
#' value is named, the names are assumed to be keywords indicating the cursor
#' position for adding each value in the vector. If value is not named, a
#' keyword parameter with the same length as value must be provided. When `named
#' = FALSE`, no keyword parameter is required. Add [add_str_with_keys()] works
#' identically but uses a str parameter and .f defaults to [add_xml_to_body()].
#'
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
#' @importFrom cli cli_abort cli_alert_warning
add_to_body <- function(docx,
                        keyword = NULL,
                        id = NULL,
                        index = NULL,
                        value = NULL,
                        str = NULL,
                        pos = "after",
                        ...) {
  check_docx(docx)

  if ((!is.null(str) & !is.null(value)) | is_all_null(c(str, value))) {
    cli_abort(
      "Either {.arg str} or {.arg value} must be provided."
    )
  }

  if (!is_all_null(c(keyword, id, index))) {
    docx <- cursor_docx(docx, keyword, id, index)
  }

  if (!is.null(str)) {
    return(officer::body_add_xml(docx, str, pos))
  }

  if (!is.null(pos) & !is.null(value)) {
    return(officer::body_add_par(docx, value, pos = pos, ...))
  }

  officer::body_add(docx, value, ...)
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
                             pos = "after",
                             .na = "NA",
                             .null = NULL,
                             .envir = parent.frame(),
                             ...) {
  rlang::check_required(value)
  value <- glue::glue(value, .na = .na, .null = .null, .envir = .envir)
  add_to_body(docx, value = value, style = style, pos = pos, ...)
}

#' @inheritParams officer::body_add_par
#' @name add_xml_to_body
#' @rdname add_to_body
#' @export
add_xml_to_body <- function(docx,
                            str,
                            pos = "after",
                            ...) {
  rlang::check_required(str)
  add_to_body(docx, str = str, pos = pos, ...)
}

#' @param gt_object A gt object converted to an OOXML string with
#'   [gt::as_word()] then passed to [add_xml_to_body()] as str parameter.
#'   Required for [add_gt_to_body()].
#' @inheritParams gt::as_word
#' @name add_gt_to_body
#' @rdname add_to_body
#' @export
#' @importFrom rlang check_required check_installed
add_gt_to_body <- function(docx,
                           gt_object,
                           align = "center",
                           caption_location = c("top", "bottom", "embed"),
                           caption_align = "left",
                           split = FALSE,
                           keep_with_next = TRUE,
                           pos = "after",
                           ...) {
  rlang::check_required(gt_object)
  rlang::check_installed("gt")
  add_xml_to_body(
    docx,
    str = gt::as_word(
      gt_object,
      align = align,
      caption_location = caption_location,
      caption_align = caption_align,
      split = split,
      keep_with_next = keep_with_next
    ),
    pos = pos,
    ...
    )
}


#' @name add_value_with_keys
#' @rdname add_to_body
#' @param .f Any function that takes a docx and value parameter and returns a
#'   rdocx object. A keyword parameter must also be supported if named is TRUE.
#'   Defaults to [add_text_to_body()].
#' @export
#' @importFrom rlang check_required as_function is_named
add_value_with_keys <- function(docx,
                                value,
                                ...,
                                .f = add_text_to_body) {
  arg <- "keyword"
  rlang::check_required(value)
  .f <- rlang::as_function(.f)

  value <- set_vec_value_names(value, arg = arg, ...)
  params <- rlang::list2(...)

  if (rlang::has_name(params, arg)) {
    params[[arg]] <- NULL
  }

  for (i in seq_along(value)) {
    docx <-
      rlang::exec(
        .f,
        docx,
        value = value[[i]],
        keyword = names(value)[[i]],
        !!!params
      )
  }

  docx
}

#' @name add_str_with_keys
#' @rdname add_to_body
#' @export
#' @importFrom rlang check_required as_function is_named
add_str_with_keys <- function(docx,
                              str,
                              ...,
                              .f = add_xml_to_body) {
  arg <- "keyword"
  rlang::check_required(str)
  .f <- rlang::as_function(.f)

  str <- set_vec_value_names(str, arg = arg, ...)
  params <- rlang::list2(...)

  if (rlang::has_name(params, arg)) {
    params[[arg]] <- NULL
  }

  for (i in seq_along(str)) {
    docx <-
      rlang::exec(
        .f,
        docx,
        str = str[[i]],
        keyword = names(str)[[i]],
        !!!params
      )
  }

  docx
}

#' Set names for a value vector object
#'
#' @keywords internal
#' @param nm Names for value.
#' @param arg Name of the argument in ... to use as names for value when nm is `NULL`.
#' @noRd
#' @importFrom rlang is_named list2 set_names
#' @importFrom cli cli_abort
set_vec_value_names <- function(value, nm = NULL, arg = "keyword", ...) {
  if (rlang::is_named(value)) {
    return(value)
  }

  if (is.null(nm)) {
    params <- rlang::list2(...)
    if (is.null(params[[arg]]) | (length(params[[arg]]) != length(value))) {
      cli_abort(
        "{.arg value} must {.arg {arg}} must be be the same length as {.arg value}."
      )
    }
    nm <- params[[arg]]
  }


  rlang::set_names(value, nm)
}
