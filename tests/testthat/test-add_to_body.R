test_that("add_to_body works", {
  docx_example <- read_docx_ext(
    filename = "example.docx",
    path = system.file("doc_examples", package = "officer")
  )

  expect_error(add_to_body(docx_example))

  docx_text <- add_text_to_body(
    docx_example,
    value = "ABCDEFG",
    keyword = "Title 1"
  )

  expect_snapshot(docx_text)

  docx_value_key <- add_value_with_keys(
    docx_example,
    value = c("ABC", "CDE"),
    keyword = c("Title 1", "Title 2")
  )

  expect_snapshot(docx_value_key)

  docx_value_named <- add_value_with_keys(
    docx_example,
    value = c("Title 1" = "ABC", "Title 2" = "CDE")
  )

  expect_snapshot(docx_value_named)
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

  docx_gt <-
    add_gt_to_body(
      docx,
      tab_1,
      keyword = "Sub title 1"
    )

  expect_snapshot(docx_gt)

  tab_str <- gt::as_word(tab_1)

  docx_str_keys <-
    add_str_with_keys(
      docx,
      str = c("Title 1" = tab_str, "Title 2" = tab_str)
    )

  expect_snapshot(docx_str_keys)
})
