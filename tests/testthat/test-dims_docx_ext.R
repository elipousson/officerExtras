test_that("dims_docx_ext works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_snapshot(dims_docx_ext(docx))
})
