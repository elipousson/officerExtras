#' @keywords internal
#' @noRd
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
officer_table_pivot <- function(x) {
  x <- x[, c("row_id", "cell_id", "text")]

  as.data.frame(
    tapply(
      x[["text"]],
      list(
        row_id = x[["row_id"]],
        cell_id = x[["cell_id"]]
      ),
      FUN = I
    )
  )
}

#' @keywords internal
#' @noRd
#' @importFrom rlang set_names
#' @importFrom utils tail head
officer_table <- function(x, index = NULL, has_header = TRUE) {
  table_cells <- subset_type(x, "table cell")

  if (!is.null(index)) {
    table_cells <- subset_index(x, index)
  }

  if (rlang::has_name(x, "is_header")) {
    table_body <- officer_table_pivot(subset_header(table_cells, FALSE))
    table_header <- officer_table_pivot(subset_header(table_cells, TRUE))
  } else {
    table_body <- officer_table_pivot(table_cells)
    table_header <- data.frame()
  }

  if ((nrow(table_header) == 1)) {
    return(rlang::set_names(table_body, table_header))
  }

  if (nrow(table_header) == 0) {
    if (!isTRUE(has_header)) {
      return(table_body)
    }

    return(
      rlang::set_names(
        utils::tail(table_body, nrow(table_body) - 1),
        utils::head(table_body, 1)
      )
    )
  }

  list(table_body, table_header)
}

#' Get tables from a rdocx or rpptx object
#'
#' Get one or more tables from a rdocx or rpptx object. Functions based on
#' example code on extracting Word document and PowerPoint slides in the
#' [officeverse
#' documentation](https://ardata-fr.github.io/officeverse/extract-content.html#word-tables).
#'
#' @param x A rdocx or rpptx object or a data.frame created with
#'   [officer_summary()]
#' @param index A index value matching a doc_index value for a table in the
#'   summary data.frame, Default: `NULL`
#' @param has_header If `TRUE`, tables are expected to have implicit headers
#'   even if the object Default: `TRUE`
#' @return A data.frame or list of data.frame (or list objects).
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
#' @export
#' @importFrom rlang has_name %||%
officer_tables <- function(x, index = NULL, has_header = TRUE) {
  if (is_officer(x, c("rdocx", "rpptx"))) {
    x <- officer_summary(x)
  }

  if (!rlang::has_name(x, "content_type") | !is.data.frame(x)) {
    cli_abort(
      "{.arg x} must be a {cli_vec_cls(c('rdocx', 'rpptx'))} object or a
      {.cls data.frame} created with {.fn officer_summary}"
    )
  }

  tables <-
    map(
      index %||% officer_table_index(x),
      ~ officer_table(x, .x, has_header)
    )

  if (length(tables) == 1) {
    return(tables[[1]])
  }

  tables
}
