test_that("officer_summary works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_true(
    is_officer_summary(
      officer_summary(docx)
    )
  )

  expect_true(
    is_officer_summary(
      officer_summary(officer::docx_summary(docx))
    )
  )

  pptx <- read_pptx_ext(
    filename = "example.pptx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_true(
    is_officer_summary(officer_summary(pptx))
  )

  expect_true(
    is_officer_summary(
      officer_summary(pptx, "pptx"),
      "pptx"
    )
  )

  expect_true(
    is_officer_summary(
      officer_summary(pptx, "layout"),
      "layout"
    )
  )

  expect_true(
    is_officer_summary(
      officer_summary(pptx, "slide", index = 1),
      "slide"
    )
  )
})
