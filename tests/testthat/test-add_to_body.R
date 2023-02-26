test_that("add_to_body works", {
  docx_example <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )
  docx <- add_text_to_body(
    docx_example,
    value = "ABCDEFG",
    keyword = "Title 1"
  )
  expect_snapshot(docx)
})

test_that("add_gt_to_body works", {
  docx <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  skip_if_not_installed("gt")
  tab_1 <-
    gt::gt(
      gt::exibble,
      rowname_col = "row",
      groupname_col = "group"
    )
  docx <-
    add_gt_to_body(
      docx,
      tab_1,
      keyword = "Sub title 1"
    )
  expect_snapshot(docx)
})
