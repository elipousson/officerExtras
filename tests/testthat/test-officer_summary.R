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

test_that("is_officer_summary returns FALSE for non data.frame input", {
  expect_false(is_officer_summary("not a data frame"))
  expect_false(is_officer_summary(list(content_type = "paragraph")))
})

test_that("officer_summary as_tibble = FALSE returns a plain data.frame", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  summary_df <- officer_summary(docx, as_tibble = FALSE)

  expect_false(inherits(summary_df, "tbl_df"))
  expect_s3_class(summary_df, "data.frame")
})

test_that("officer_summary errors for objects that aren't rdocx/rpptx", {
  expect_error(
    officer_summary("not an officer object")
  )
})
