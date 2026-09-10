test_that("vec_add_to_body works", {
  docx_example <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  docx_rows <- nrow(officer_summary(docx_example))
  new_rows <- 15
  values <- rep("new text", new_rows)

  docx_update <- vec_add_to_body(
    docx_example,
    value = values
  )

  docx_update_summary <- officer_summary(docx_update)

  expect_equal(
    nrow(docx_update_summary),
    docx_rows + new_rows
  )

  new_idx <- which(docx_update_summary[["text"]] == "new text")

  # New rows must all be present and inserted as a contiguous block
  expect_length(new_idx, new_rows)
  expect_true(all(diff(new_idx) == 1))
})

test_that("vec_add_to_body recycles style and value parameters", {
  docx_example <- read_officer()

  docx_update <- vec_add_to_body(
    docx_example,
    value = c("Sample text 1", "Sample text 2", "Sample text 3"),
    style = c("heading 1", "heading 2", "Normal")
  )

  summary_df <- officer_summary(docx_update)

  expect_identical(
    summary_df[["text"]],
    c("Sample text 1", "Sample text 2", "Sample text 3")
  )

  expect_identical(
    summary_df[["style_name"]],
    c("heading 1", "heading 2", "Normal")
  )
})

test_that("vec_add_to_body supports .sep as a function", {
  docx_example <- read_officer()

  docx_update <- vec_add_to_body(
    docx_example,
    value = rep("Text", 3),
    style = "Normal",
    .sep = officer::body_add_break
  )

  summary_df <- officer_summary(docx_update)

  expect_equal(
    sum(summary_df[["text"]] == "Text"),
    3
  )
})

test_that("vec_add_to_body supports .sep as a value passed to add_to_body", {
  docx_example <- read_officer()

  docx_update <- vec_add_to_body(
    docx_example,
    value = c("A", "B"),
    .sep = "separator"
  )

  summary_df <- officer_summary(docx_update)

  expect_identical(
    summary_df[["text"]],
    c("A", "separator", "B")
  )
})
