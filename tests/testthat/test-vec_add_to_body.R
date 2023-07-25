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

  expect_identical(
    docx_update_summary[c((docx_rows + 1):(docx_rows + new_rows)), ][["text"]],
    values
  )
})
