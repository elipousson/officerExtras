test_that("officer_summary works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  comparison_docx_summary <- officer::docx_summary(docx)

  expect_identical(
    officer_summary(docx),
    comparison_docx_summary
  )

  expect_identical(
    officer_summary(officer::docx_summary(docx)),
    comparison_docx_summary
  )

  expect_identical(
    officer_summary(officer::docx_summary(docx)),
    comparison_docx_summary
  )

  pptx <- read_pptx_ext(
    filename = "example.pptx",
    path = system.file("doc_examples", package = "officer")
  )

  comparison_pptx_summary <- officer::pptx_summary(pptx)

  expect_identical(
    officer_summary(pptx),
    comparison_pptx_summary
  )

  expect_identical(
    officer_summary(pptx, "pptx"),
    comparison_pptx_summary
  )

  expect_identical(
    officer_summary(pptx, "layout"),
    officer::layout_summary(pptx)
  )

  expect_identical(
    officer_summary(pptx, "slide", index = 1),
    officer::slide_summary(pptx, 1)
  )
})
