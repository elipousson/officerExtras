#' Get tables from a rdocx or rpptx object
#'
#' Get one or more tables from a rdocx or rpptx object. These functions are
#' based on example code on extracting Word document and PowerPoint slides in
#' the [officeverse
#' documentation](https://ardata-fr.github.io/officeverse/extract-content.html#word-tables).
#' #' [officer_table()] is a lower-level helper function to extract a single
#' table from a document. Some additional features including the type_convert
#' parameter and the addition of doc_index values as the default names for the
#' returned list of tables are based on [this blog post by Matt
#' Dray](https://www.rostrum.blog/2023/06/07/rectangular-officer/).
#'
#' @param x A rdocx or rpptx object or a data frame created with
#'   [officer_summary()].
#' @param index A index value matching a doc_index value for a table in the
#'   summary data frame, Default: `NULL`
#' @param has_header If `TRUE` (default), tables are expected to have implicit
#'   headers even if the Word table does not have an explicit header row. If
#'   `FALSE`, only explicit header rows will be used as column names.
#' @param col If col is supplied, [officer_table()] passes col and the
#'   additional parameters in ... to [fill_with_pattern()]. This allows the
#'   addition of preceding headings or captions as a column within the
#'   data.frame returned by [officer_tables()]. This is an experimental feature
#'   and may be modified or removed. Defaults to `NULL`.
#' @param ... Additional parameters passed to [fill_with_pattern()].
#' @param stack If `TRUE` and all tables share the same number of columns,
#'   return a single combined data frame instead of a list. Defaults to `FALSE`.
#' @param type_convert If `TRUE`, convert columns for the returned data frames
#'   to the appropriate type using [utils::type.convert()].
#' @param nm Names to use for returned list of tables. If `NULL` (default), the
#'   names are set to the doc_index values using the pattern
#'   "doc_index_<doc_index_number>".
#' @inheritParams rlang::args_error_context
#' @return A list of data frames or, if stack is `TRUE`, a single data frame.
#' @examples
#' docx <- read_docx_ext(
#'   filename = "example.docx",
#'   path = system.file("doc_examples", package = "officer")
#' )
#'
#' officer_tables(docx)
#'
#' pptx <- read_pptx_ext(
#'   filename = "example.pptx",
#'   path = system.file("doc_examples", package = "officer")
#' )
#'
#' officer_tables(pptx)[[1]]
#'
#' @rdname officer_tables
#' @seealso [docxtractr::docx_extract_all()]
#' @export
#' @importFrom utils type.convert
#' @importFrom rlang set_names
officer_tables <- function(x,
                           index = NULL,
                           has_header = TRUE,
                           col = NULL,
                           preserve = FALSE,
                           ...,
                           stack = FALSE,
                           type_convert = FALSE,
                           nm = NULL,
                           call = caller_env()) {
  check_required(x)
  if (is_officer(x, c("rdocx", "rpptx"))) {
    x <- officer_summary(x, preserve = preserve, call = call)
  }

  if (!has_name(x, "content_type") || !is.data.frame(x)) {
    cli_abort(
      "{.arg x} must be a {cli_vec_cls(c('rdocx', 'rpptx'))} object or a data
      frame created with {.fn officer_summary}",
      call = call
    )
  }

  index <- index %||% officer_table_index(x)

  tables <- vector("list", length(index))

  tables <-
    map(
      index,
      function(i) {
        officer_table(
          x = x,
          index = i,
          has_header = has_header,
          col = col,
          call = call
        )
      }
    )

  if (type_convert) {
    tables <- lapply(tables, utils::type.convert, as.is = TRUE)
  }

  if (stack) {
    n_cols <- unique(vapply(tables, ncol, 1))
    if (length(n_cols) > 1) {
      cli_abort(
        "{.arg stack} must be {.code FALSE} when {.arg x}
        includes tables with a varying number of columns.",
        call = call
      )
    }

    return(do.call("rbind", tables))
  }

  rlang::set_names(tables, nm %||% glue("doc_index_{index}"))
}

#' @rdname officer_tables
#' @name officer_table
#' @export
#' @importFrom rlang has_name set_names
#' @importFrom utils head tail
officer_table <- function(x,
                          index = NULL,
                          has_header = TRUE,
                          col = NULL,
                          ...,
                          call = caller_env()) {
  check_string(col, allow_null = TRUE, call = call)

  if (!is_null(col)) {
    x <- fill_with_pattern(x, ..., col = col, call = call)
  }

  # Subset by doc_index or content_type
  if (!is.null(index)) {
    table_cells <- subset_index(x, index)
  } else {
    table_cells <- subset_type(x, "table cell")
  }

  if (!is_null(col)) {
    col_value <- table_cells[[col]]
  }

  body_cells <- table_cells
  header_cells <- data.frame()

  if (rlang::has_name(x, "is_header")) {
    header_cells <- officer_table_pivot(subset_header(table_cells, TRUE))
    body_cells <- subset_header(table_cells, FALSE)
  }

  body_cells <- officer_table_pivot(body_cells)

  n_header_rows <- nrow(header_cells)
  n_body_rows <- nrow(body_cells)

  if (!all(is_empty(col))) {
    col_value <- unique(table_cells[[col]])

    if (n_header_rows > 1) {
      cli_abort(
        "{.arg col} can't be used with tables with more than 1 header row."
      )
    }
  }

  if (n_header_rows == 0) {
    if (!is_null(col) && !is_null(col_value)) {
      check_string(col_value, call = call)

      body_col <- data.frame(
        c(col, rep(col_value, n_body_rows - 1))
      )

      body_col <- set_names(body_col, ncol(body_cells) + 1)

      body_cells <-
        cbind(
          body_cells,
          body_col
        )
    }

    if (is_false(has_header)) {
      return(body_cells)
    }

    nm <- utils::head(body_cells, 1)

    return(
      rlang::set_names(
        utils::tail(body_cells, n_body_rows - 1),
        nm
      )
    )
  }

  if (n_header_rows == 1) {
    nm <- header_cells

    if (!is_null(col)) {
      nm <- c(nm, col)
      body_cells[[col]] <- rep(col_value, n_body_rows)
    }

    return(rlang::set_names(body_cells, nm))
  }

  list(body_cells, header_cells)
}

#' @keywords internal
#' @noRd
#' @importFrom rlang has_name
officer_table_index <- function(x) {
  tables <- subset_type(x, "table cell")
  if (rlang::has_name(x, "doc_index")) {
    unique(tables[["doc_index"]])
  } else if (rlang::has_name(x, "id")) {
    unique(tables[["id"]])
  }
}

#' @keywords internal
#' @noRd
officer_table_pivot <- function(x,
                                call = caller_env()) {
  tbl <-
    data.frame(
      tapply(
        x[["text"]],
        list(
          row_id = x[["row_id"]],
          cell_id = x[["cell_id"]]
        ),
        FUN = I
      )
    )

  rownames(tbl) <- NULL

  tbl
}
