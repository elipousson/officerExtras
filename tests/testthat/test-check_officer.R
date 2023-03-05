test_that("check_officer works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_null(
    check_docx(docx)
  )

  expect_error(
    check_docx(0)
  )

  expect_error(
    check_pptx(docx)
  )

  expect_error(
    check_xlsx(docx)
  )
})

test_that("check_office_fileext works", {
  expect_null(
    check_docx_fileext("test.docx")
  )
  expect_null(
    check_pptx_fileext("test.pptx")
  )
  expect_null(
    check_xlsx_fileext("test.xlsx")
  )
  expect_error(
    check_docx_fileext("test")
  )
  expect_error(
    check_pptx_fileext("test.docx")
  )
})
