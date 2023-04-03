#' Read a Google Docs document, Slides, or Sheets URL to a rdocx, rpptx, or
#' rxlsx object
#'
#' Uses the Google Docs, Slides, or Sheets URL to export a file locally, read to
#' an officer object. If filename is `NULL`, exported file is removed after
#' export.
#'
#' @param url A URL for a Google Doc, Google Slides presentation, or Google
#'   Sheets.
#' @param format File format to use for export URL. Options are `NULL`
#'   (default), "doc", "pptx", or "xslx". "pdf" and "csv" be supported in the
#'   future.
#'   in the future.
#' @param filename Destination file name. Optional. If destfile is `NULL`,
#'   downloaded file is removed as part of the function execution.
#' @param path Folder path. Optional.
#' @param quiet If `TRUE`, suppress messages when downloading file.
#' @export
#' @importFrom glue glue
read_docs_url <- function(url,
                          format = NULL,
                          filename = NULL,
                          path = NULL,
                          quiet = TRUE) {
  export <- prep_docs_export(url, format)

  if (!is.null(destfile)) {
    check_office_fileext(destfile)
  }

  path <- str_c(
    path,
    filename %||% export[["filename"]],
    sep = .Platform$file.sep
  )

  if (!file.exists(path)) {
    download.file(export[["url"]], path, quiet = quiet)
  }

  docx <- read_officer(path, quiet = quiet)

  if (is.null(filename)) {
    file.remove(path)
  }

  docx
}

#' @noRd
prep_docs_export <- function(url, format = NULL) {
  if (is_docs_url(url)) {
    suffix <- "edit"
    format <- match.arg(format, c("doc", "pdf"))
    fileext <- "docx"
  } else if (is_slides_url(url)) {
    suffix <- "export"
    format <- match.arg(format, c("pptx", "pdf"))
    fileext <- "pptx"
  } else if (is_slides_url(url)) {
    suffix <- "edit"
    format <- match.arg(format, c("xlsx", "csv", "pdf"))
    fileext <- "xlsx"
  }

  id <- extract_docs_id(url, suffix)

  # https://www.labnol.org/internet/direct-links-for-google-drive/28356/
  url <-
    switch(fileext,
      "docx" = glue("https://docs.google.com/document/d/{id}/export?format={format}"),
      "pptx" = glue("https://docs.google.com/presentation/d/{id}/export/{format}"),
      "xlsx" = glue("https://docs.google.com/spreadsheets/d/{id}/export?format={format}")
    )

  list(
    "url" = url,
    "filename" = office_temp(fileext)
  )
}

#' @noRd
is_gdoc_url <- function(x) {
  grepl("^https://docs.google.com/", x)
}

#' @noRd
is_docs_url <- function(x, type = "document") {
  grepl(glue("^https://docs.google.com/{type}/"), x)
}

#' @noRd
is_slides_url <- function(x) {
  is_docs_url(x, "presentation")
}

#' @noRd
is_sheets_url <- function(x) {
  is_docs_url(x, "spreadsheets")
}

#' @noRd
extract_docs_id <- function(url, suffix = "edit") {
  str_extract(
    url,
    glue("(?<=/d/).+(?=/{suffix})")
  )
}
