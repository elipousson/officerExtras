test_that("multiplication works", {
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
