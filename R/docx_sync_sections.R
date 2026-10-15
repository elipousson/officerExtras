#' Copy footer and page settings from the last section to every section
#'
#' @description
#' `docx_sync_sections()` copies the header and footer references, page
#' margins, and page size from the last section of a Word document to the
#' other sections, filling in what each section is missing.
#'
#' This fixes documents rendered by Quarto or Pandoc where section breaks don't
#' carry the reference document's page settings, for example when using
#' Quarto's `.landscape` divs (which add section breaks with no footer, no
#' margins, and an A4 landscape page size) or a custom OpenXML template. See
#' the "Page settings and sections" section of [brand_reference_docx()].
#'
#' For each section other than the last:
#'
#' - If the section has no header or footer references, the last section's
#'   references are added.
#' - If the section has no page margins, the last section's margins are added.
#' - If the section has no page size, the last section's page size is added.
#'   If `page_size = TRUE`, landscape sections also get the last section's page
#'   size with the width and height swapped (and portrait sections get the last
#'   section's page size).
#'
#' To use `docx_sync_sections()` in a Quarto project, add a `post-render`
#' script to `_quarto.yml`:
#'
#' ```yaml
#' project:
#'   post-render: post-render.R
#' ```
#'
#' With a `post-render.R` file:
#'
#' ```r
#' output_files <- strsplit(Sys.getenv("QUARTO_PROJECT_OUTPUT_FILES"), "\n")[[1]]
#'
#' for (file in output_files[grepl("\\.docx$", output_files)]) {
#'   officerExtras::docx_sync_sections(file)
#' }
#' ```
#'
#' @param x A rdocx object or a path to a docx file.
#' @param path Output file path. If `NULL` (default) and `x` is a file path,
#'   `x` is overwritten. If `NULL` and `x` is a rdocx object, the document is
#'   not written to a file.
#' @param page_size If `TRUE` (default), set the page size of every section to
#'   the last section's page size (with the width and height swapped for
#'   landscape sections). If `FALSE`, only add a page size to sections without
#'   one.
#' @returns A rdocx object (invisibly if a file is written).
#' @examples
#' docx <- officer::read_docx()
#' docx <- officer::body_add_par(docx, "Portrait")
#' docx <- officer::body_end_section_portrait(docx)
#' docx <- officer::body_add_par(docx, "Landscape")
#' docx <- officer::body_end_section_landscape(docx)
#'
#' docx_sync_sections(docx)
#' @export
docx_sync_sections <- function(x, path = NULL, page_size = TRUE) {
  if (is_officer(x, "rdocx")) {
    doc <- x
  } else {
    check_docx_fileext(x)

    if (!file.exists(x)) {
      cli_abort("{.arg x} file {.file {x}} does not exist.")
    }

    doc <- officer::read_docx(x)
    path <- path %||% x
  }

  check_bool(page_size)

  body <- officer::docx_body_xml(doc)
  ns <- w_ns(body)
  remove_officer_ref_marks(body)
  last <- xml2::xml_find_first(body, "//w:body/w:sectPr", ns)

  if (inherits(last, "xml_missing")) {
    cli_abort("{.arg x} has no section properties.")
  }

  sections <- xml2::xml_find_all(body, "//w:body//w:p/w:pPr/w:sectPr", ns)

  for (sect_pr in sections) {
    sync_section(sect_pr, last, page_size = page_size)
  }

  if (is.null(path)) {
    return(doc)
  }

  write_officer(doc, path = path)
  invisible(doc)
}

#' Fill in a section's header and footer references, margins, and page size
#' from `last`
#' @noRd
sync_section <- function(sect_pr, last, page_size = TRUE) {
  ns <- w_ns(sect_pr)
  find <- function(node, name) {
    xml2::xml_find_first(node, paste0("w:", name), ns)
  }
  is_missing <- function(node) inherits(node, "xml_missing")

  refs <- xml2::xml_find_all(sect_pr, "w:headerReference|w:footerReference", ns)

  if (length(refs) == 0) {
    last_refs <- xml2::xml_find_all(last, "w:headerReference|w:footerReference", ns)

    # References come first in w:sectPr; add them in reverse so they keep
    # their order
    for (ref in rev(last_refs)) {
      xml2::xml_add_child(sect_pr, ref, .where = 0)
    }
  }

  last_pg_mar <- find(last, "pgMar")

  if (is_missing(find(sect_pr, "pgMar")) && !is_missing(last_pg_mar)) {
    xml2::xml_add_child(sect_pr, last_pg_mar)
  }

  last_pg_sz <- find(last, "pgSz")
  pg_sz <- find(sect_pr, "pgSz")

  if (!is_missing(last_pg_sz) && (is_missing(pg_sz) || page_size)) {
    landscape <- !is_missing(pg_sz) &&
      identical(xml2::xml_attr(pg_sz, "orient"), "landscape")

    size <- xml2::xml_attrs(last_pg_sz)
    size <- size[names(size) %in% c("w", "h", "code")]
    width <- as.numeric(size[["w"]])
    height <- as.numeric(size[["h"]])

    if (landscape) {
      size[c("w", "h")] <- as.character(c(max(width, height), min(width, height)))
      size[["orient"]] <- "landscape"
    }

    set_ooxml_child(sect_pr, "pgSz", size)
  }

  sort_sect_pr(sect_pr)
}

#' Remove officer's marks from header and footer references that point to a
#' part
#'
#' officer marks the header and footer references it adds with
#' `w:officer="true"` and, when the document is printed, writes the content
#' stored in each marked reference to a new part. The mark is left on the
#' reference in the written file (with the content removed), so printing a
#' document read from that file fails. References that point to a part (with
#' an `r:id` other than officer's `"%s"` placeholder) don't need the mark.
#' @noRd
remove_officer_ref_marks <- function(body) {
  ns <- w_ns(body)
  refs <- xml2::xml_find_all(
    body,
    "//w:headerReference[@w:officer]|//w:footerReference[@w:officer]",
    ns
  )

  for (ref in refs) {
    if (!identical(xml2::xml_attrs(ref)[["id"]], "%s")) {
      xml2::xml_set_attr(ref, "w:officer", NULL, ns = ns)
    }
  }

  invisible(body)
}
