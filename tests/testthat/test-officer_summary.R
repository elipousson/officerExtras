test_that("officer_summary works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_s3_class(officer_summary(pptx), "data.frame")

  pptx <- read_pptx_ext(
    filename = "example.pptx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_s3_class(officer_summary(pptx), "data.frame")
})
