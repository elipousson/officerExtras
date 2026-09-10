test_that("fill_with_pattern fills a new column based on a heading pattern", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )
  summary_df <- officer_summary(docx)

  filled <- fill_with_pattern(summary_df)

  expect_true("heading" %in% names(filled))
  expect_identical(filled[["heading"]][[1]], "Title 1")
  expect_identical(filled[["heading"]][[3]], "Title 2")
})

test_that("fill_with_pattern supports custom column names", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )
  summary_df <- officer_summary(docx)

  filled <- fill_with_pattern(
    summary_df,
    col = "section",
    fill_col = "text",
    pattern = "^heading 1$"
  )

  expect_true("section" %in% names(filled))
  expect_identical(filled[["section"]][[1]], "Title 1")
})

test_that("fill_with_pattern errors when col already exists", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )
  summary_df <- officer_summary(docx)

  expect_error(
    fill_with_pattern(summary_df, col = "text")
  )
})

test_that("fill_with_pattern returns x unchanged when pattern has no matches", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )
  summary_df <- officer_summary(docx)

  no_match <- fill_with_pattern(summary_df, pattern = "no style matches this")

  expect_identical(no_match, summary_df)
})
