test_that("check_docx works", {
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
})

test_that("check_docx_fileext works", {
  expect_null(
    check_docx_fileext("test.docx")
  )
  expect_error(
    check_docx_fileext("test")
  )
})
