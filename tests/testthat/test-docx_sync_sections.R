# A document like Quarto's .landscape output: an empty section break, a
# landscape A4 section with no margins or footer, and a final letter section
# with a footer
quarto_landscape_docx <- function() {
  w_ns <- 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  doc <- officer::read_docx()
  doc <- officer::body_add_par(doc, "Portrait before")
  doc <- officer::body_add_xml(
    doc,
    sprintf("<w:p %s><w:pPr><w:sectPr></w:sectPr></w:pPr></w:p>", w_ns)
  )
  doc <- officer::body_add_par(doc, "Landscape")
  doc <- officer::body_add_xml(
    doc,
    sprintf(
      '<w:p %s><w:pPr><w:sectPr><w:pgSz w:h="11906" w:w="16838" w:orient="landscape"/></w:sectPr></w:pPr></w:p>',
      w_ns
    )
  )
  doc <- officer::body_add_par(doc, "Portrait after")
  officer::body_set_default_section(
    doc,
    officer::prop_section(
      page_size = officer::page_size(width = 8.5, height = 11),
      page_margins = officer::page_mar(top = 1, bottom = 1, left = 1.25, right = 1.25),
      footer_default = officer::block_list(officer::fpar("Footer"))
    )
  )
}

section_props <- function(doc) {
  body <- officer::docx_body_xml(doc)
  ns <- xml2::xml_ns(body)
  sections <- xml2::xml_find_all(body, "//w:body//w:p/w:pPr/w:sectPr", ns)

  lapply(sections, function(sect_pr) {
    find_attrs <- function(name) {
      xml2::xml_attrs(xml2::xml_find_first(sect_pr, paste0("w:", name), ns))
    }

    list(
      refs = length(xml2::xml_find_all(sect_pr, "w:footerReference", ns)),
      pgSz = find_attrs("pgSz"),
      pgMar = find_attrs("pgMar"),
      children = xml2::xml_name(xml2::xml_children(sect_pr))
    )
  })
}

test_that("docx_sync_sections fills in footers, margins, and page size", {
  doc <- docx_sync_sections(quarto_landscape_docx())
  sections <- section_props(doc)

  expect_length(sections, 2)

  portrait <- sections[[1]]
  landscape <- sections[[2]]

  expect_identical(portrait$refs, 1L)
  expect_identical(landscape$refs, 1L)

  expect_identical(portrait$pgSz[c("w", "h")], c(w = "12240", h = "15840"))
  expect_identical(
    landscape$pgSz[c("w", "h", "orient")],
    c(w = "15840", h = "12240", orient = "landscape")
  )

  expect_identical(portrait$pgMar[["left"]], "1800")
  expect_identical(landscape$pgMar[["left"]], "1800")

  # Elements are in schema order
  expect_identical(landscape$children, c("footerReference", "pgSz", "pgMar"))
})

test_that("docx_sync_sections keeps existing page sizes if page_size = FALSE", {
  doc <- docx_sync_sections(quarto_landscape_docx(), page_size = FALSE)
  sections <- section_props(doc)

  # Missing page size is added, existing page size is kept
  expect_identical(sections[[1]]$pgSz[c("w", "h")], c(w = "12240", h = "15840"))
  expect_identical(sections[[2]]$pgSz[c("w", "h")], c(w = "16838", h = "11906"))
})

test_that("docx_sync_sections keeps existing section settings", {
  doc <- officer::read_docx()
  doc <- officer::body_add_par(doc, "First")
  doc <- officer::body_end_block_section(
    doc,
    officer::block_section(officer::prop_section(
      page_margins = officer::page_mar(left = 2),
      header_default = officer::block_list(officer::fpar("Own header"))
    ))
  )
  doc <- officer::body_set_default_section(
    doc,
    officer::prop_section(footer_default = officer::block_list(officer::fpar("Footer")))
  )

  sections <- section_props(docx_sync_sections(doc))

  # A section with its own header keeps its references
  expect_identical(sections[[1]]$refs, 0L)
  expect_identical(sections[[1]]$pgMar[["left"]], "2880")
})

test_that("docx_sync_sections reads and writes files", {
  path <- withr::local_tempfile(fileext = ".docx")
  print(quarto_landscape_docx(), target = path)

  expect_invisible(docx_sync_sections(path))
  sections <- section_props(officer::read_docx(path))
  expect_identical(sections[[1]]$refs, 1L)

  output <- withr::local_tempfile(fileext = ".docx")
  docx_sync_sections(path, path = output)
  expect_true(file.exists(output))

  expect_error(docx_sync_sections("missing.docx"), "does not exist")
  expect_error(docx_sync_sections(path, page_size = "yes"), "page_size")
})
