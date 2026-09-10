test_that("officer_summary_levels adds columns for each heading level", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  level_summary <- officer_summary_levels(docx)

  expect_s3_class(level_summary, "data.frame")
  expect_true(all(has_name(level_summary, c("heading_1", "heading_2"))))
  expect_identical(level_summary[["heading_1"]][[1]], "Title 1")

  level_tables <- officer_tables(officer_summary(docx), col = "heading_1")
  expect_s3_class(level_tables[[1]], "data.frame")
  expect_identical(level_tables[[1]][["heading_1"]][[1]], "Sub title 2")
})

test_that("officer_summary_levels supports exclude_levels", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  level_summary <- officer_summary_levels(docx, exclude_levels = "List Paragraph")

  expect_false("list_paragraph" %in% names(level_summary))
  expect_true(all(has_name(level_summary, c("heading_1", "heading_2"))))
})

test_that("officer_summary_levels warns for invalid levels unless strict", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_message(
    officer_summary_levels(
      docx,
      levels = c("heading 1", "not a style"),
      strict = FALSE
    ),
    "invalid values"
  )

  expect_error(
    officer_summary_levels(
      docx,
      levels = c("heading 1", "not a style"),
      strict = TRUE
    )
  )
})

test_that("officer_summary_levels supports a custom levels_from column", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  level_summary <- officer_summary_levels(docx, levels_from = "content_type")

  expect_true(all(has_name(level_summary, c("paragraph", "table_cell"))))
})
