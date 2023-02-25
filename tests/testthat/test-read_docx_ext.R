test_that("read_docx_ext works", {
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

  expect_message(
    read_docx_ext(filename = "test.docx", docx = docx, quiet = FALSE)
  )
})
