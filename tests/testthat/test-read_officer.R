test_that("read_officer works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_true(
    inherits(docx, "rdocx")
  )

  expect_identical(
    read_docx_ext(docx = docx),
    docx
  )

  expect_identical(
    read_docx_ext(allow_null = TRUE)[["styles"]],
    officer::read_docx()[["styles"]]
  )

  expect_error(
    read_docx_ext()
  )

  expect_message(
    read_docx_ext(docx = docx, quiet = FALSE)
  )

  expect_message(
    read_docx_ext(filename = "test.docx", docx = docx, quiet = FALSE)
  )

  pptx <- read_pptx_ext(
    filename = "example.pptx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_true(
    inherits(pptx, "rpptx")
  )

  xlsx <- read_xlsx_ext(
    filename = "template.xlsx",
    path = system.file("template", package = "officer")
  )

  expect_true(
    inherits(xlsx, "rxlsx")
  )
})
